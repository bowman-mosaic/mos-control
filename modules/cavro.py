"""
Tecan Cavro XCalibur syringe-pump control module — exposes operations via
Flask API.  Wraps syringe_pump/tecan_cavro.py (+ ftdi_serial.py, motion.py).

All 4 pumps share one RS-485 serial connection (single COM port).
A bus-level lock (_bus_lock) serializes every serial transaction so
multiple pumps can operate concurrently from different threads.
"""

from modules._api import expose, push_event
import threading
import time
import sys
import os

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "syringe_pump"))
from ftdi_serial import Serial as FtdiSerial              # noqa: E402
from tecan_cavro import (                                          # noqa: E402
    TecanCavro, TecanCavroDeviceError, TecanCavroDeviceBusyError,
)
from syringe_pump_control import list_serial_ports         # noqa: E402

NUM_CAVRO = 4

# Aspiration / refill plunger speed (counts/s). Pump max (~2900) cavitates
# on the pull; 2000 was clean in bench tests. Dispense/flush still use max.
REFILL_VELOCITY_COUNTS = 2000

_serial = None
_serial_lock = threading.Lock()   # protects connect/disconnect
_bus_lock = threading.RLock()      # serializes all serial I/O (reentrant for retries)
_pumps: list = [None] * NUM_CAVRO
_proto_threads: list = [None] * NUM_CAVRO
_proto_stops = [threading.Event() for _ in range(NUM_CAVRO)]
_cont_threads: list = [None] * NUM_CAVRO
_cont_stops = [threading.Event() for _ in range(NUM_CAVRO)]
_flush_port: list = [1] * NUM_CAVRO  # role-aware final destination per pump
_stop_all_thread = None


def _patch_pump_for_concurrency(pump):
    """Patch a TecanCavro instance so all serial I/O is bus-locked and
    wait_for_ready releases the lock between polls.

    command_request_raw does request() + read() which must be atomic,
    so we lock the entire method rather than individual serial calls.
    """
    orig_cmd_raw = pump.command_request_raw

    def _locked_cmd_raw(*args, **kwargs):
        with _bus_lock:
            return orig_cmd_raw(*args, **kwargs)
    pump.command_request_raw = _locked_cmd_raw

    def _wait(poll_interval=0.5):
        start = time.time()
        serial_errors = 0
        max_serial_errors = 10
        while True:
            if pump.wait_timeout is not None and (time.time() - start) > pump.wait_timeout:
                from syringe_pump.tecan_cavro import TecanCavroReadyTimeout
                raise TecanCavroReadyTimeout(
                    f'Timeout waiting for Cavro (address: {pump.address})')
            try:
                if pump.ready(log=False):
                    return
                serial_errors = 0
            except TecanCavroDeviceError as e:
                if e.status == TecanCavroDeviceError.COMMAND_OVERFLOW:
                    time.sleep(poll_interval)
                    continue
                serial_errors += 1
                backoff = min(poll_interval * (2 ** serial_errors), 10)
                print(f"[Cavro P{pump.address}] serial error during poll "
                      f"({serial_errors}/{max_serial_errors}): {e}")
                if serial_errors >= max_serial_errors:
                    raise
                time.sleep(backoff)
                continue
            except Exception as e:
                serial_errors += 1
                backoff = min(poll_interval * (2 ** serial_errors), 10)
                print(f"[Cavro P{pump.address}] serial error during poll "
                      f"({serial_errors}/{max_serial_errors}): {e}")
                if serial_errors >= max_serial_errors:
                    raise
                time.sleep(backoff)
                continue
            time.sleep(poll_interval)
    pump.wait_for_ready = _wait


def _max_vel(pump):
    return min(pump.counts_per_stroke, int(5800 / pump.velocity_scale))


def _refill_vel(pump):
    """Plunger speed for filling the syringe from a reservoir."""
    return max(3, min(REFILL_VELOCITY_COUNTS, _max_vel(pump)))


def _check(idx):
    if idx < 0 or idx >= NUM_CAVRO or _pumps[idx] is None:
        return False
    return True


def _hms_to_seconds(hms):
    if not hms or hms == "00:00:00":
        return 0
    parts = hms.split(":")
    return int(parts[0]) * 3600 + int(parts[1]) * 60 + int(parts[2])


def _is_overflow(err):
    return (isinstance(err, TecanCavroDeviceError)
            and err.status == TecanCavroDeviceError.COMMAND_OVERFLOW)


def _is_retryable(err):
    return _is_overflow(err) or isinstance(err, TecanCavroDeviceBusyError)


def _halt_pumps(*idxs):
    """Terminate motion and drop any locally queued batch (T does not clear
    the Python-side batch string, so leftover I/V/A would get appended)."""
    for ci in idxs:
        if ci is not None and _check(ci):
            p = _pumps[ci]
            try:
                p.clear_batch()
            except Exception:
                pass
            try:
                p.halt()
            except Exception:
                pass


def _recover_after_halt(*pumps):
    """Wait until pumps will accept a new command after Terminate."""
    for p in pumps:
        try:
            p.clear_batch()
        except Exception:
            pass
    time.sleep(0.25)
    for p in pumps:
        deadline = time.time() + 8
        while time.time() < deadline:
            try:
                if p.ready(log=False):
                    break
            except Exception as e:
                if not _is_retryable(e):
                    time.sleep(0.2)
                    continue
            time.sleep(0.2)
        else:
            try:
                p.halt()
            except Exception:
                pass
            time.sleep(0.3)


def _flush_to_port(pump, port, max_vel, label):
    """Empty the syringe through ``port`` as separate short commands.

    After a halt the XCalibur buffer is often still full, so a combined
    valve+move batch (the previous approach) raises Error 15. The pump
    may also still report busy — retry that the same way as overflow.
    """
    try:
        vol = pump.volume_ml
    except Exception:
        vol = 1.0
    if vol <= 0.001:
        return
    print(f"{label}: {vol:.3f} mL → port {port}")
    last_err = None
    for attempt in range(6):
        try:
            _recover_after_halt(pump)
            pump.clear_batch()
            pump.switch_valve(port, wait=True)
            pump.move_absolute_counts(0, velocity_counts=max_vel, wait=True)
            return
        except Exception as e:
            last_err = e
            if not _is_retryable(e):
                raise
            print(f"{label}: {e}, retry {attempt + 1}")
            try:
                pump.halt()
            except Exception:
                pass
            time.sleep(0.4)
    if last_err:
        raise last_err


def _send_batch_retry(pump, retries=5):
    """Load the Python-side batch into the pump, retrying Error 15."""
    if not getattr(pump, "batch_command", ""):
        return
    last_err = None
    for attempt in range(retries):
        try:
            pump.send_batch()
            return
        except Exception as e:
            last_err = e
            if not _is_retryable(e):
                raise
            print(f"[Cavro P{pump.address}] overflow loading batch, "
                  f"retry {attempt + 1}/{retries}")
            try:
                pump.halt()
            except Exception:
                pass
            _recover_after_halt(pump)
    if last_err:
        raise last_err


def _sync_execute(*pumps):
    """Start only these pumps.

    Do not use the bus-wide broadcast R — that also executes leftover
    commands on idle pumps sharing the RS-485 line and overflows them.
    Load each buffer, then fire addressed R back-to-back.
    """
    for p in pumps:
        _send_batch_retry(p)
    for p in pumps:
        p.execute(broadcast=False, wait=False)
    TecanCavro.wait_for_all(*pumps)


def _broadcast_valves(*pump_ports):
    """Queue valve positions and start those pumps together."""
    pumps = [p for p, _ in pump_ports]
    last_err = None
    for attempt in range(4):
        try:
            _recover_after_halt(*pumps)
            for pump, port in pump_ports:
                pump.start_batch()
                pump.switch_valve(port)
            _sync_execute(*pumps)
            return
        except Exception as e:
            last_err = e
            if not _is_retryable(e):
                raise
            print(f"[Cavro] valve sync overflow, retry {attempt + 1}")
            for p in pumps:
                try:
                    p.halt()
                    p.clear_batch()
                except Exception:
                    pass
    if last_err:
        raise last_err


def _broadcast_moves(*pump_moves):
    """Queue plunger moves and start those pumps together.

    pump_moves: (pump, position_counts, velocity_counts)
    """
    pumps = [p for p, _, _ in pump_moves]
    last_err = None
    for attempt in range(4):
        try:
            _recover_after_halt(*pumps)
            for pump, counts, vel in pump_moves:
                pump.start_batch()
                pump.move_absolute_counts(counts, velocity_counts=vel, wait=False)
            _sync_execute(*pumps)
            return
        except Exception as e:
            last_err = e
            if not _is_retryable(e):
                raise
            print(f"[Cavro] move sync overflow, retry {attempt + 1}")
            for p in pumps:
                try:
                    p.halt()
                    p.clear_batch()
                except Exception:
                    pass
    if last_err:
        raise last_err


# ── Port discovery (shared with Harvard module) ─────────────────────────────

@expose
def cavro_get_ports():
    try:
        return list_serial_ports()
    except Exception as e:
        return {"error": str(e)}


# ── Bus connection ───────────────────────────────────────────────────────────

@expose
def cavro_connect_bus(port, syringe_ml=0.5, valve_positions=6,
                      addresses=None):
    """Open the shared RS-485 serial line and create pump instances.

    addresses: list of 4 ints (RS-485 addresses), default [0,1,2,3].
    """
    global _serial
    if addresses is None:
        addresses = list(range(NUM_CAVRO))

    with _serial_lock:
        if _serial is not None:
            return {"error": "Bus already connected — disconnect first"}
        try:
            TecanCavro.instances.clear()
            _serial = FtdiSerial(
                device_port=port,
                baudrate=9600,
                read_timeout=100,
                write_timeout=100,
            )
            for i in range(NUM_CAVRO):
                _pumps[i] = TecanCavro(
                    _serial,
                    address=int(addresses[i]),
                    syringe_volume_ml=float(syringe_ml),
                    total_valve_positions=int(valve_positions),
                )
                _patch_pump_for_concurrency(_pumps[i])
            return {"ok": True,
                    "msg": f"Bus connected on {port}, {NUM_CAVRO} pumps"}
        except Exception as e:
            _serial = None
            for i in range(NUM_CAVRO):
                _pumps[i] = None
            TecanCavro.instances.clear()
            return {"error": str(e)}


@expose
def cavro_disconnect_bus():
    global _serial
    # Signal all background threads to stop
    _coord_stop.set()
    _dual_stop.set()
    for i in range(NUM_CAVRO):
        _cont_stops[i].set()
        _proto_stops[i].set()
    # Halt all pumps immediately so they stop moving
    for i in range(NUM_CAVRO):
        if _pumps[i] is not None:
            try:
                _pumps[i].halt()
            except Exception:
                pass
    # Wait for background threads to finish
    if _coord_thread and _coord_thread.is_alive():
        _coord_thread.join(timeout=10)
    if _dual_thread and _dual_thread.is_alive():
        _dual_thread.join(timeout=10)
    for i in range(NUM_CAVRO):
        if _cont_threads[i] and _cont_threads[i].is_alive():
            _cont_threads[i].join(timeout=10)
        if _proto_threads[i] and _proto_threads[i].is_alive():
            _proto_threads[i].join(timeout=10)
    with _serial_lock:
        for i in range(NUM_CAVRO):
            _pumps[i] = None
        TecanCavro.instances.clear()
        if _serial is not None:
            try:
                _serial.close()
            except Exception:
                pass
            _serial = None
    return {"ok": True}


@expose
def cavro_is_connected():
    return _serial is not None


# ── Homing ───────────────────────────────────────────────────────────────────

@expose
def cavro_home(idx):
    if not _check(idx):
        return {"error": "Not connected"}
    try:
        _pumps[idx].home()
        return {"ok": True}
    except Exception as e:
        return {"error": str(e)}


@expose
def cavro_home_all():
    if _serial is None:
        return {"error": "Bus not connected"}
    try:
        TecanCavro.home_all()
        return {"ok": True}
    except Exception as e:
        return {"error": str(e)}


# ── Movement ─────────────────────────────────────────────────────────────────

@expose
def cavro_dispense(idx, volume_ml, from_port, to_port, velocity_ml=None):
    if not _check(idx):
        return {"error": "Not connected"}
    try:
        _flush_port[idx] = int(from_port)
        pump = _pumps[idx]
        max_v = int(5800 / pump.velocity_scale)
        kwargs = {
            "velocity_counts": _refill_vel(pump),
            "dispense_velocity_counts": pump._velocity_counts,
        }
        if velocity_ml is not None:
            vel_counts = min(float(velocity_ml) * pump.counts_per_ml, max_v)
            kwargs["dispense_velocity_counts"] = max(3, round(vel_counts))
        pump.dispense_ml(
            float(volume_ml),
            from_port=int(from_port),
            to_port=int(to_port),
            **kwargs,
        )
        return {"ok": True}
    except Exception as e:
        return {"error": str(e)}


@expose
def cavro_withdraw(idx, volume_ml, from_port, velocity_ml=None, outlet_port=None):
    """Aspirate (withdraw) fluid from from_port and push it out to outlet_port.

    Loops pulling strokes from from_port at the specified velocity and pushing
    each one to outlet_port at max speed.  Always ends with a push to outlet
    so the syringe is left empty.
    """
    if not _check(idx):
        return {"error": "Not connected"}
    if outlet_port is None:
        return {"error": "outlet_port required"}
    try:
        _flush_port[idx] = int(from_port)
        pump = _pumps[idx]
        vol = float(volume_ml)
        syr = pump.syringe_volume_ml
        fp = int(from_port)
        op = int(outlet_port)
        max_v = int(5800 / pump.velocity_scale)
        if velocity_ml is not None:
            vel_counts = max(3, min(round(float(velocity_ml) * pump.counts_per_ml), max_v))
        else:
            vel_counts = pump._velocity_counts

        full_counts = int(syr * pump.counts_per_ml)
        iterations = int(vol // syr)
        remaining = vol % syr

        for _ in range(iterations):
            pump.start_batch()
            pump.switch_valve(fp)
            pump.move_absolute_counts(full_counts, velocity_counts=vel_counts, wait=False)
            pump.execute(wait=True)
            pump.start_batch()
            pump.switch_valve(op)
            pump.move_absolute_counts(0, velocity_counts=max_v, wait=False)
            pump.execute(wait=True)

        if remaining > 0.0001:
            target_counts = round(remaining * pump.counts_per_ml)
            pump.start_batch()
            pump.switch_valve(fp)
            pump.move_absolute_counts(target_counts, velocity_counts=vel_counts, wait=False)
            pump.execute(wait=True)
            pump.start_batch()
            pump.switch_valve(op)
            pump.move_absolute_counts(0, velocity_counts=max_v, wait=False)
            pump.execute(wait=True)

        return {"ok": True}
    except Exception as e:
        return {"error": str(e)}


@expose
def cavro_move_absolute(idx, position_ml, velocity_ml=None):
    if not _check(idx):
        return {"error": "Not connected"}
    try:
        kwargs = {}
        if velocity_ml is not None:
            kwargs["velocity_ml"] = float(velocity_ml)
        _pumps[idx].move_absolute_ml(float(position_ml), **kwargs)
        return {"ok": True}
    except Exception as e:
        return {"error": str(e)}


@expose
def cavro_move_relative(idx, delta_ml, velocity_ml=None):
    if not _check(idx):
        return {"error": "Not connected"}
    try:
        kwargs = {}
        if velocity_ml is not None:
            kwargs["velocity_ml"] = float(velocity_ml)
        _pumps[idx].move_relative_ml(float(delta_ml), **kwargs)
        return {"ok": True}
    except Exception as e:
        return {"error": str(e)}


@expose
def cavro_switch_valve(idx, position):
    if not _check(idx):
        return {"error": "Not connected"}
    try:
        _pumps[idx].switch_valve(int(position))
        return {"ok": True}
    except Exception as e:
        return {"error": str(e)}


# ── Status / halt ────────────────────────────────────────────────────────────

@expose
def cavro_halt(idx):
    if not _check(idx):
        return {"error": "Not connected"}
    try:
        _pumps[idx].halt()
        return {"ok": True}
    except Exception as e:
        return {"error": str(e)}


def _finish_stop_all(threads, flush_targets):
    """Wait for role-specific thread cleanup, then flush any remainder.

    Coordinated and dual-push threads perform their own final flushes. Waiting
    for them before this fallback prevents a generic valve command from racing
    those role-aware flushes.
    """
    seen = set()
    for thread in threads:
        if thread is None or id(thread) in seen:
            continue
        seen.add(id(thread))
        if thread.is_alive():
            thread.join(timeout=60)

    for idx, pump, port in flush_targets:
        if pump is None:
            continue
        try:
            max_vel = min(
                pump.counts_per_stroke,
                int(5800 / pump.velocity_scale),
            )
            _flush_to_port(
                pump, port, max_vel,
                f"[Cavro STOP] final flush P{idx+1}",
            )
        except Exception as e:
            print(f"[Cavro STOP] final flush P{idx+1} failed: {e}")


@expose
def cavro_stop_all():
    """Halt all pumps and flush each active role to its safe destination.

    Inlet pumps flush to their configured ``from`` ports. Coordinated and
    dual-push outlet pumps flush to their configured ``to`` ports.
    """
    global _stop_all_thread
    _coord_stop.set()
    _dual_stop.set()
    for i in range(NUM_CAVRO):
        _cont_stops[i].set()
        _proto_stops[i].set()

    threads = (
        _coord_thread,
        _dual_thread,
        *_cont_threads,
        *_proto_threads,
    )
    flush_targets = tuple(
        (i, _pumps[i], _flush_port[i]) for i in range(NUM_CAVRO)
    )

    for i in range(NUM_CAVRO):
        if _pumps[i] is not None:
            try:
                _pumps[i].halt()
            except Exception:
                pass

    _stop_all_thread = threading.Thread(
        target=_finish_stop_all,
        args=(threads, flush_targets),
        daemon=True,
    )
    _stop_all_thread.start()
    return {"ok": True}


@expose
def cavro_get_status(idx):
    if not _check(idx):
        return {"error": "Not connected"}
    try:
        ready = _pumps[idx].ready()
        pos = _pumps[idx].volume_ml
        valve = _pumps[idx].valve_position
        return {
            "ok": True,
            "ready": ready,
            "position_ml": round(pos, 4),
            "valve_position": valve,
        }
    except Exception as e:
        return {"error": str(e)}


# ── Semi-continuous pumping ───────────────────────────────────────────────────

@expose
def cavro_continuous_start(idx, from_port, to_port, velocity_ml=None,
                           cycles=0):
    """Pump fluid semi-continuously: aspirate full syringe from from_port,
    dispense to to_port, repeat.

    cycles: number of full aspirate/dispense cycles.  0 = run until stopped.
    """
    if not _check(idx):
        return {"error": "Not connected"}
    _flush_port[idx] = int(from_port)
    if _cont_threads[idx] and _cont_threads[idx].is_alive():
        _cont_stops[idx].set()
        try:
            _pumps[idx].halt()
        except Exception:
            pass
        _cont_threads[idx].join(timeout=60)
    if _proto_threads[idx] and _proto_threads[idx].is_alive():
        return {"error": "Protocol already running on this pump"}
    _cont_stops[idx].clear()
    ready = threading.Event()
    t = threading.Thread(
        target=_continuous_thread,
        args=(idx, int(from_port), int(to_port),
              float(velocity_ml) if velocity_ml is not None else None,
              int(cycles), ready),
        daemon=True,
    )
    _cont_threads[idx] = t
    t.start()
    ready.wait(timeout=30)
    return {"ok": True}


@expose
def cavro_continuous_stop(idx):
    """Stop continuous pumping and wait for flush to complete."""
    _cont_stops[idx].set()
    if _check(idx):
        try:
            _pumps[idx].halt()
        except Exception:
            pass
    if _cont_threads[idx] and _cont_threads[idx].is_alive():
        _cont_threads[idx].join(timeout=60)
    return {"ok": True}


@expose
def cavro_continuous_is_running(idx):
    if _cont_threads[idx] and _cont_threads[idx].is_alive():
        return True
    return False


def _continuous_thread(idx, from_port, to_port, velocity_ml, cycles, ready=None):
    pump = _pumps[idx]
    syringe_ml = pump.syringe_volume_ml
    max_vel = _max_vel(pump)
    refill_vel = _refill_vel(pump)
    # Convert mL/s to counts/s, clamped to pump max
    max_v = int(5800 / pump.velocity_scale)
    if velocity_ml:
        dispense_vel = max(3, min(round(velocity_ml * pump.counts_per_ml), max_v))
    else:
        dispense_vel = None
    n = 0
    print(f"[Cavro C{idx+1}] continuous: from={from_port} to={to_port} "
          f"vel_ml={velocity_ml} dispense_vel_counts={dispense_vel} "
          f"refill_vel={refill_vel} max_vel={max_vel} syringe={syringe_ml}mL")
    try:
        push_event("onCavroContinuousUpdate", idx, "started",
                   f"Continuous: {from_port}→{to_port}" +
                   (f", {cycles} cycles" if cycles else ", until stopped"))

        # Pre-fill: batch valve switch + refill into one command
        push_event("onCavroContinuousUpdate", idx, "running",
                   "Pre-filling syringe...")
        full_counts = int(syringe_ml * pump.counts_per_ml)
        pump.start_batch()
        pump.switch_valve(from_port)
        pump.move_absolute_counts(full_counts, velocity_counts=refill_vel, wait=False)
        pump.execute(wait=True)
        print(f"[Cavro C{idx+1}] pre-fill done, plunger at {pump.volume_ml:.3f} mL")
        if ready:
            ready.set()

        while not _cont_stops[idx].is_set():
            if cycles and n >= cycles:
                break
            n += 1

            # Push: batch valve switch + dispense into one command
            print(f"[Cavro C{idx+1}] cycle {n}: switching to port {to_port}, pushing to 0")
            pump.start_batch()
            pump.switch_valve(to_port)
            try:
                vel = dispense_vel if dispense_vel else max_vel
                pump.move_absolute_counts(0, velocity_counts=vel, wait=False)
                pump.execute(wait=True)
            except Exception:
                pass
            if _cont_stops[idx].is_set():
                print(f"[Cavro C{idx+1}] halted mid push")
                break
            print(f"[Cavro C{idx+1}] cycle {n}: push done, plunger at {pump.volume_ml:.3f} mL")

            # Pull: batch valve switch + refill into one command
            print(f"[Cavro C{idx+1}] cycle {n}: switching to port {from_port}, pulling")
            pump.start_batch()
            pump.switch_valve(from_port)
            try:
                pump.move_absolute_counts(full_counts, velocity_counts=refill_vel, wait=False)
                pump.execute(wait=True)
            except Exception:
                pass
            if _cont_stops[idx].is_set():
                print(f"[Cavro C{idx+1}] halted mid pull")
                break
            print(f"[Cavro C{idx+1}] cycle {n}: pull done, plunger at {pump.volume_ml:.3f} mL")

            volume_so_far = n * syringe_ml
            push_event("onCavroContinuousUpdate", idx, "running",
                       f"Cycle {n} done — {volume_so_far:.2f} mL total")

        # Final flush: batch valve switch + push back to source at max speed
        if pump.volume_ml > 0.001:
            print(f"[Cavro C{idx+1}] final flush: {pump.volume_ml:.3f} mL back to port {from_port}")
            pump.start_batch()
            pump.switch_valve(from_port)
            pump.move_absolute_counts(0, velocity_counts=max_vel, wait=False)
            pump.execute(wait=True)
            print(f"[Cavro C{idx+1}] final flush done")

        status = "stopped" if _cont_stops[idx].is_set() else "complete"
        total = n * syringe_ml
        push_event("onCavroContinuousUpdate", idx, status,
                   f"Continuous {status}: {n} cycles, {total:.2f} mL")
    except Exception as e:
        import traceback
        print(f"[Cavro C{idx+1}] ERROR: {e}")
        traceback.print_exc()
        if ready:
            ready.set()
        try:
            pump.halt()
        except Exception:
            pass
        push_event("onCavroContinuousUpdate", idx, "error",
                   f"Continuous error: {e}")


# ── Coordinated two-pump push-pull ────────────────────────────────────────────

_coord_stop = threading.Event()
_coord_thread = None
_coord_pump_idxs = (None, None)


@expose
def cavro_coordinated_start(idx1, idx2, from1, to1, from2, to2,
                            velocity_ml=None):
    """Run two pumps in lockstep: Pump 1 pushes while Pump 2 pulls
    (through a device), then Pump 1 refills while Pump 2 dispenses.

    idx1/idx2: pump indices (0-based)
    from1/to1: inlet pump source port / device port
    from2/to2: outlet pump device port / destination port
    velocity_ml: push/pull rate in mL/s (same for both pumps in active phase)
    """
    global _coord_thread, _coord_pump_idxs
    if not _check(idx1) or not _check(idx2):
        return {"error": "Pump(s) not connected"}
    if _coord_thread and _coord_thread.is_alive():
        _coord_stop.set()
        _halt_pumps(*_coord_pump_idxs)
        _coord_thread.join(timeout=60)
    _coord_stop.clear()
    i1, i2 = int(idx1), int(idx2)
    f1, t1_ = max(1, min(6, int(from1))), max(1, min(6, int(to1)))
    f2, t2_ = max(1, min(6, int(from2))), max(1, min(6, int(to2)))
    _flush_port[i1] = f1
    _flush_port[i2] = t2_
    _coord_pump_idxs = (i1, i2)
    ready = threading.Event()
    t = threading.Thread(
        target=_coordinated_thread,
        args=(i1, i2, f1, t1_, f2, t2_,
              float(velocity_ml) if velocity_ml is not None else None,
              ready),
        daemon=True,
    )
    _coord_thread = t
    t.start()
    ready.wait(timeout=30)
    return {"ok": True}


@expose
def cavro_coordinated_stop():
    """Stop coordinated pumping and wait for flush to complete."""
    global _coord_thread
    _coord_stop.set()
    _halt_pumps(*_coord_pump_idxs)
    if _coord_thread and _coord_thread.is_alive():
        _coord_thread.join(timeout=60)
    return {"ok": True}


@expose
def cavro_coordinated_is_running():
    return _coord_thread is not None and _coord_thread.is_alive()


def _coordinated_thread(idx1, idx2, from1, to1, from2, to2, velocity_ml, ready=None):
    pump1 = _pumps[idx1]
    pump2 = _pumps[idx2]
    syringe_ml = pump1.syringe_volume_ml
    max_v = int(5800 / pump1.velocity_scale)
    max_vel = min(pump1.counts_per_stroke, max_v)
    refill_vel = _refill_vel(pump1)

    if velocity_ml:
        push_vel = max(3, min(round(velocity_ml * pump1.counts_per_ml), max_v))
    else:
        push_vel = max_vel

    n = 0
    print(f"[Cavro COORD] P{idx1+1}({from1}→{to1}) + P{idx2+1}({from2}→{to2}) "
          f"vel_ml={velocity_ml} push_vel={push_vel} refill_vel={refill_vel} "
          f"max_vel={max_vel} syringe={syringe_ml}mL")
    try:
        push_event("onCavroCoordinatedUpdate", "started",
                   f"Coordinated: P{idx1+1}({from1}→{to1}) + "
                   f"P{idx2+1}({from2}→{to2})")

        # Pre-fill: Pump 1 fills from source (batch valve+move)
        print(f"[Cavro COORD] pre-filling P{idx1+1} from port {from1}")
        pump1.start_batch()
        pump1.switch_valve(from1)
        pump1.move_absolute_counts(
            int(syringe_ml * pump1.counts_per_ml),
            velocity_counts=refill_vel, wait=False)
        pump1.execute(wait=True)
        print(f"[Cavro COORD] pre-fill done")
        if ready:
            ready.set()

        while not _coord_stop.is_set():
            n += 1

            # Phase A: Pump 1 pushes to device, Pump 2 pulls from device
            # Batch valve switch + move into single commands, then broadcast
            print(f"[Cavro COORD] cycle {n} phase A: "
                  f"P{idx1+1} push to {to1}, P{idx2+1} pull from {from2}")
            pump1.start_batch()
            pump1.switch_valve(to1)
            pump1.move_absolute_counts(0, velocity_counts=push_vel, wait=False)
            pump2.start_batch()
            pump2.switch_valve(from2)
            pump2.move_absolute_counts(
                int(syringe_ml * pump2.counts_per_ml),
                velocity_counts=push_vel, wait=False)
            TecanCavro.broadcast_execute(pump1, pump2)
            TecanCavro.wait_for_all(pump1, pump2)
            if _coord_stop.is_set():
                print(f"[Cavro COORD] halted mid phase A")
                break
            print(f"[Cavro COORD] cycle {n} phase A done")

            # Phase B: Pump 1 refills from source,
            #          Pump 2 dispenses to destination (fast)
            print(f"[Cavro COORD] cycle {n} phase B: "
                  f"P{idx1+1} refill from {from1} @ {refill_vel}, "
                  f"P{idx2+1} dispense to {to2}")
            pump1.start_batch()
            pump1.switch_valve(from1)
            pump1.move_absolute_counts(
                int(syringe_ml * pump1.counts_per_ml),
                velocity_counts=refill_vel, wait=False)
            pump2.start_batch()
            pump2.switch_valve(to2)
            pump2.move_absolute_counts(0, velocity_counts=max_vel, wait=False)
            TecanCavro.broadcast_execute(pump1, pump2)
            TecanCavro.wait_for_all(pump1, pump2)
            if _coord_stop.is_set():
                print(f"[Cavro COORD] halted mid phase B")
                break
            print(f"[Cavro COORD] cycle {n} phase B done")

            total = n * syringe_ml
            push_event("onCavroCoordinatedUpdate", "running",
                       f"Cycle {n} — {total:.2f} mL total")

        # Final flush: after halt the pump buffer is often still full.
        _recover_after_halt(pump1, pump2)
        _flush_to_port(pump1, from1, max_vel,
                       f"[Cavro COORD] final flush P{idx1+1}")
        _flush_to_port(pump2, to2, max_vel,
                       f"[Cavro COORD] final flush P{idx2+1}")
        print(f"[Cavro COORD] final flush done")

        status = "stopped" if _coord_stop.is_set() else "complete"
        total = n * syringe_ml
        push_event("onCavroCoordinatedUpdate", status,
                   f"Coordinated {status}: {n} cycles, {total:.2f} mL")
    except Exception as e:
        import traceback
        print(f"[Cavro COORD] ERROR: {e}")
        traceback.print_exc()
        if ready:
            ready.set()
        for p in (pump1, pump2):
            try:
                p.halt()
            except Exception:
                pass
        push_event("onCavroCoordinatedUpdate", "error",
                   f"Coordinated error: {e}")


# ── Dual-push (two inlets + one outlet) ──────────────────────────────────────
# Two inlet pumps push 1:1 into the device while the outlet pulls the combined
# volume. Each cycle: each inlet dispenses half a syringe, outlet fills one
# full syringe (same total volume), then inlets refill and outlet wastes.

_dual_stop = threading.Event()
_dual_thread = None
_dual_pump_idxs = (None, None, None)


@expose
def cavro_dualpush_start(idx1, idx2, idx3, from1, to1, from2, to2,
                         from3, to3, velocity_ml=None):
    """Run two inlet pumps and one outlet in lockstep.

    idx1: inlet 1, idx2: inlet 2, idx3: outlet (0-based)
    from1/to1: inlet 1 source port / device port
    from2/to2: inlet 2 source port / device port
    from3/to3: outlet device port / waste port
    velocity_ml: per-inlet push rate in mL/s (outlet pulls at 2× so volumes match)
    """
    global _dual_thread, _dual_pump_idxs
    i1, i2, i3 = int(idx1), int(idx2), int(idx3)
    if len({i1, i2, i3}) < 3:
        return {"error": "Dual push needs three different pumps"}
    if not _check(i1) or not _check(i2) or not _check(i3):
        return {"error": "Pump(s) not connected"}
    if _dual_thread and _dual_thread.is_alive():
        _dual_stop.set()
        _halt_pumps(*_dual_pump_idxs)
        _dual_thread.join(timeout=60)
    _dual_stop.clear()
    f1, t1_ = max(1, min(6, int(from1))), max(1, min(6, int(to1)))
    f2, t2_ = max(1, min(6, int(from2))), max(1, min(6, int(to2)))
    f3, t3_ = max(1, min(6, int(from3))), max(1, min(6, int(to3)))
    _flush_port[i1] = f1
    _flush_port[i2] = f2
    _flush_port[i3] = t3_
    _dual_pump_idxs = (i1, i2, i3)
    ready = threading.Event()
    t = threading.Thread(
        target=_dualpush_thread,
        args=(i1, i2, i3, f1, t1_, f2, t2_, f3, t3_,
              float(velocity_ml) if velocity_ml is not None else None,
              ready),
        daemon=True,
    )
    _dual_thread = t
    t.start()
    ready.wait(timeout=30)
    return {"ok": True}


@expose
def cavro_dualpush_stop():
    """Stop dual-push pumping and wait for flush to complete."""
    global _dual_thread
    _dual_stop.set()
    _halt_pumps(*_dual_pump_idxs)
    if _dual_thread and _dual_thread.is_alive():
        _dual_thread.join(timeout=60)
    return {"ok": True}


@expose
def cavro_dualpush_is_running():
    return _dual_thread is not None and _dual_thread.is_alive()


def _dualpush_thread(idx1, idx2, idx3, from1, to1, from2, to2,
                     from3, to3, velocity_ml, ready=None):
    p1 = _pumps[idx1]
    p2 = _pumps[idx2]
    p3 = _pumps[idx3]
    syringe_ml = p1.syringe_volume_ml
    max_v = int(5800 / p1.velocity_scale)
    max_vel = min(p1.counts_per_stroke, max_v)
    refill_vel = _refill_vel(p1)
    full = int(syringe_ml * p1.counts_per_ml)
    half = full // 2

    if velocity_ml:
        inlet_vel = max(3, min(round(velocity_ml * p1.counts_per_ml), max_v))
    else:
        inlet_vel = max_vel
    # Outlet always 2× inlet. If 2× would exceed the pump max, drop both
    # inlets together so they stay matched and outlet stays exactly 2×.
    if inlet_vel * 2 > max_v:
        inlet_vel = max(3, max_v // 2)
    outlet_vel = inlet_vel * 2

    n = 0
    print(f"[Cavro DUAL] P{idx1+1}({from1}→{to1}) + P{idx2+1}({from2}→{to2}) "
          f"+ out P{idx3+1}({from3}→{to3}) vel_ml={velocity_ml} "
          f"inlet_vel={inlet_vel} outlet_vel={outlet_vel} (2× inlet) "
          f"refill_vel={refill_vel} max_vel={max_vel} syringe={syringe_ml}mL")
    try:
        push_event("onCavroDualPushUpdate", "started",
                   f"Dual push: P{idx1+1}+P{idx2+1} @ {inlet_vel} cts/s, "
                   f"P{idx3+1} @ {outlet_vel} cts/s (2×)")

        print(f"[Cavro DUAL] pre-filling inlets P{idx1+1} and P{idx2+1} in sync")
        _recover_after_halt(p1, p2, p3)
        _broadcast_valves((p1, from1), (p2, from2))
        _broadcast_moves(
            (p1, full, refill_vel),
            (p2, full, refill_vel),
        )
        print(f"[Cavro DUAL] pre-fill done")
        if ready:
            ready.set()

        while not _dual_stop.is_set():
            n += 1

            # Phase A: both inlet valves to device together, then both
            # plungers push at inlet_vel while outlet pulls at 2×.
            print(f"[Cavro DUAL] cycle {n} phase A: "
                  f"sync valves, then P{idx1+1}+P{idx2+1} push @ {inlet_vel}, "
                  f"P{idx3+1} pull @ {outlet_vel}")
            _broadcast_valves((p1, to1), (p2, to2), (p3, from3))
            if _dual_stop.is_set():
                print(f"[Cavro DUAL] halted mid phase A (valves)")
                break
            _broadcast_moves(
                (p1, half, inlet_vel),
                (p2, half, inlet_vel),
                (p3, full, outlet_vel),
            )
            if _dual_stop.is_set():
                print(f"[Cavro DUAL] halted mid phase A")
                break
            print(f"[Cavro DUAL] cycle {n} phase A done")

            # Phase B: inlets refill in sync, outlet to waste
            print(f"[Cavro DUAL] cycle {n} phase B: "
                  f"sync valves, inlets refill @ {refill_vel}, "
                  f"P{idx3+1} waste to {to3}")
            _broadcast_valves((p1, from1), (p2, from2), (p3, to3))
            if _dual_stop.is_set():
                print(f"[Cavro DUAL] halted mid phase B (valves)")
                break
            _broadcast_moves(
                (p1, full, refill_vel),
                (p2, full, refill_vel),
                (p3, 0, max_vel),
            )
            if _dual_stop.is_set():
                print(f"[Cavro DUAL] halted mid phase B")
                break
            print(f"[Cavro DUAL] cycle {n} phase B done")

            total = n * syringe_ml
            push_event("onCavroDualPushUpdate", "running",
                       f"Cycle {n} — {total:.2f} mL mixed")

        _recover_after_halt(p1, p2, p3)
        _flush_to_port(p1, from1, max_vel,
                       f"[Cavro DUAL] final flush P{idx1+1}")
        _flush_to_port(p2, from2, max_vel,
                       f"[Cavro DUAL] final flush P{idx2+1}")
        _flush_to_port(p3, to3, max_vel,
                       f"[Cavro DUAL] final flush P{idx3+1}")
        print(f"[Cavro DUAL] final flush done")

        status = "stopped" if _dual_stop.is_set() else "complete"
        total = n * syringe_ml
        push_event("onCavroDualPushUpdate", status,
                   f"Dual push {status}: {n} cycles, {total:.2f} mL mixed")
    except Exception as e:
        import traceback
        print(f"[Cavro DUAL] ERROR: {e}")
        traceback.print_exc()
        if ready:
            ready.set()
        for p in (p1, p2, p3):
            try:
                p.halt()
            except Exception:
                pass
        push_event("onCavroDualPushUpdate", "error",
                   f"Dual push error: {e}")


# ── Protocol execution ──────────────────────────────────────────────────────

@expose
def cavro_run_protocol(idx, steps):
    """Run a multi-step Cavro pump protocol.

    steps: list of dicts, each with an 'action' key.
    Supported actions:
        Dispense  — {volume_ml, from_port, to_port, velocity_ml?}
        Move      — {position_ml, velocity_ml?}
        MoveRel   — {delta_ml, velocity_ml?}
        Valve     — {position}
        Home      — (no extra keys)
        Wait      — {time: "HH:MM:SS"}
    """
    if not _check(idx):
        return {"error": "Not connected"}
    if _proto_threads[idx] and _proto_threads[idx].is_alive():
        return {"error": "Protocol already running"}
    _proto_stops[idx].clear()
    t = threading.Thread(target=_run_protocol_thread,
                         args=(idx, steps), daemon=True)
    _proto_threads[idx] = t
    t.start()
    return {"ok": True}


@expose
def cavro_stop_protocol(idx):
    _stop_protocol(idx)
    return {"ok": True}


def _stop_protocol(idx):
    _proto_stops[idx].set()
    if _pumps[idx]:
        try:
            _pumps[idx].halt()
        except Exception:
            pass
    if _proto_threads[idx] and _proto_threads[idx].is_alive():
        _proto_threads[idx].join(timeout=5)


def _run_protocol_thread(idx, steps):
    pump = _pumps[idx]
    try:
        push_event("onCavroProtocolUpdate", idx, -1, "started",
                   f"Protocol started ({len(steps)} steps)")

        for i, step in enumerate(steps):
            if _proto_stops[idx].is_set():
                break

            action = step.get("action", "")

            if action == "Dispense":
                vol = float(step.get("volume_ml", 0))
                fp = int(step.get("from_port", 1))
                tp = int(step.get("to_port", 1))
                vel = step.get("velocity_ml")
                push_event("onCavroProtocolUpdate", idx, i, "running",
                           f"Step {i+1}: Dispense {vol} mL  {fp}→{tp}")
                kwargs = {}
                if vel is not None:
                    kwargs["velocity_ml"] = float(vel)
                pump.dispense_ml(vol, from_port=fp, to_port=tp, **kwargs)

            elif action == "Move":
                pos = float(step.get("position_ml", 0))
                vel = step.get("velocity_ml")
                push_event("onCavroProtocolUpdate", idx, i, "running",
                           f"Step {i+1}: Move to {pos} mL")
                kwargs = {}
                if vel is not None:
                    kwargs["velocity_ml"] = float(vel)
                pump.move_absolute_ml(pos, **kwargs)

            elif action == "MoveRel":
                delta = float(step.get("delta_ml", 0))
                vel = step.get("velocity_ml")
                push_event("onCavroProtocolUpdate", idx, i, "running",
                           f"Step {i+1}: Move relative {delta} mL")
                kwargs = {}
                if vel is not None:
                    kwargs["velocity_ml"] = float(vel)
                pump.move_relative_ml(delta, **kwargs)

            elif action == "Valve":
                pos = int(step.get("position", 1))
                push_event("onCavroProtocolUpdate", idx, i, "running",
                           f"Step {i+1}: Valve → {pos}")
                pump.switch_valve(pos)

            elif action == "Home":
                push_event("onCavroProtocolUpdate", idx, i, "running",
                           f"Step {i+1}: Home")
                pump.home()

            elif action == "Wait":
                pass  # handled below

            # Wait phase
            wait = _hms_to_seconds(step.get("time", "00:00:00"))
            if wait > 0:
                push_event("onCavroProtocolUpdate", idx, i, "waiting",
                           f"Step {i+1}: Waiting {wait}s")
                end = time.monotonic() + wait
                while time.monotonic() < end:
                    if _proto_stops[idx].is_set():
                        break
                    remaining = end - time.monotonic()
                    rm = int(remaining // 60)
                    rs = int(remaining % 60)
                    push_event("onCavroProtocolCountdown", idx, i,
                               f"{rm:02d}:{rs:02d}")
                    time.sleep(1)

        status = "aborted" if _proto_stops[idx].is_set() else "complete"
        push_event("onCavroProtocolUpdate", idx, -1, status,
                   f"Protocol {status}")
    except Exception as e:
        try:
            pump.halt()
        except Exception:
            pass
        push_event("onCavroProtocolUpdate", idx, -1, "error",
                   f"Protocol error: {e}")


# ── Public helpers for other modules (experiment engine, etc.) ───────────────

def get_pump(idx):
    if 0 <= idx < NUM_CAVRO:
        return _pumps[idx]
    return None


def disconnect():
    """Called by control_server shutdown."""
    cavro_disconnect_bus()
