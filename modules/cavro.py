"""
Tecan Cavro XCalibur syringe-pump control module — exposes operations via
Flask API.  Wraps Syringe_pump/tecan_cavro.py (+ ftdi_serial.py, motion.py).

All 4 pumps share one RS-485 serial connection (single COM port).
A bus-level lock (_bus_lock) serializes every serial transaction so
multiple pumps can operate concurrently from different threads.
"""

from modules._api import expose, push_event
import threading
import time
import sys
import os

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "Syringe_pump"))
from ftdi_serial import Serial as FtdiSerial              # noqa: E402
from tecan_cavro import TecanCavro                         # noqa: E402
from syringe_pump_control import list_serial_ports         # noqa: E402

NUM_CAVRO = 4

_serial = None
_serial_lock = threading.Lock()   # protects connect/disconnect
_bus_lock = threading.RLock()      # serializes all serial I/O (reentrant for retries)
_pumps: list = [None] * NUM_CAVRO
_proto_threads: list = [None] * NUM_CAVRO
_proto_stops = [threading.Event() for _ in range(NUM_CAVRO)]
_cont_threads: list = [None] * NUM_CAVRO
_cont_stops = [threading.Event() for _ in range(NUM_CAVRO)]


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

    def _wait(poll_interval=0.02):
        start = time.time()
        while not pump.ready(log=False):
            if pump.wait_timeout is not None and (time.time() - start) > pump.wait_timeout:
                from tecan_cavro import TecanCavroReadyTimeout
                raise TecanCavroReadyTimeout(
                    f'Timeout waiting for Cavro (address: {pump.address})')
            time.sleep(poll_interval)
    pump.wait_for_ready = _wait


def _check(idx):
    if idx < 0 or idx >= NUM_CAVRO or _pumps[idx] is None:
        return False
    return True


def _hms_to_seconds(hms):
    if not hms or hms == "00:00:00":
        return 0
    parts = hms.split(":")
    return int(parts[0]) * 3600 + int(parts[1]) * 60 + int(parts[2])


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
        pump = _pumps[idx]
        max_v = int(5800 / pump.velocity_scale)
        kwargs = {
            "velocity_counts": min(pump.counts_per_stroke, int(5800 / pump.velocity_scale)),
            "dispense_velocity_counts": pump._velocity_counts,
        }
        if velocity_ml is not None:
            vel_counts = min(float(velocity_ml) * pump.counts_per_ml, max_v)
            kwargs["dispense_velocity_counts"] = max(1, round(vel_counts))
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


@expose
def cavro_stop_all():
    """Emergency stop: halt all pumps and kill all background threads."""
    _coord_stop.set()
    for i in range(NUM_CAVRO):
        _cont_stops[i].set()
        _proto_stops[i].set()
    for i in range(NUM_CAVRO):
        if _pumps[i] is not None:
            try:
                _pumps[i].halt()
            except Exception:
                pass
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
    max_vel = min(pump.counts_per_stroke, int(5800 / pump.velocity_scale))
    # Convert mL/s to counts/s, clamped to pump max
    max_v = int(5800 / pump.velocity_scale)
    if velocity_ml:
        dispense_vel = max(1, min(round(velocity_ml * pump.counts_per_ml), max_v))
    else:
        dispense_vel = None
    n = 0
    print(f"[Cavro C{idx+1}] continuous: from={from_port} to={to_port} "
          f"vel_ml={velocity_ml} dispense_vel_counts={dispense_vel} "
          f"max_vel={max_vel} syringe={syringe_ml}mL")
    try:
        push_event("onCavroContinuousUpdate", idx, "started",
                   f"Continuous: {from_port}→{to_port}" +
                   (f", {cycles} cycles" if cycles else ", until stopped"))

        # Pre-fill: batch valve switch + fast pull into one command
        push_event("onCavroContinuousUpdate", idx, "running",
                   "Pre-filling syringe...")
        full_counts = int(syringe_ml * pump.counts_per_ml)
        pump.start_batch()
        pump.switch_valve(from_port)
        pump.move_absolute_counts(full_counts, velocity_counts=max_vel, wait=False)
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

            # Pull: batch valve switch + fast refill into one command
            print(f"[Cavro C{idx+1}] cycle {n}: switching to port {from_port}, pulling")
            pump.start_batch()
            pump.switch_valve(from_port)
            try:
                pump.move_absolute_counts(full_counts, velocity_counts=max_vel, wait=False)
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
        for ci in _coord_pump_idxs:
            if ci is not None and _check(ci):
                try:
                    _pumps[ci].halt()
                except Exception:
                    pass
        _coord_thread.join(timeout=60)
    _coord_stop.clear()
    i1, i2 = int(idx1), int(idx2)
    f1, t1_ = max(1, min(6, int(from1))), max(1, min(6, int(to1)))
    f2, t2_ = max(1, min(6, int(from2))), max(1, min(6, int(to2)))
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
    for ci in _coord_pump_idxs:
        if ci is not None and _check(ci):
            try:
                _pumps[ci].halt()
            except Exception:
                pass
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

    if velocity_ml:
        push_vel = max(1, min(round(velocity_ml * pump1.counts_per_ml), max_v))
    else:
        push_vel = max_vel

    n = 0
    print(f"[Cavro COORD] P{idx1+1}({from1}→{to1}) + P{idx2+1}({from2}→{to2}) "
          f"vel_ml={velocity_ml} push_vel={push_vel} max_vel={max_vel} "
          f"syringe={syringe_ml}mL")
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
            velocity_counts=max_vel, wait=False)
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

            # Phase B: Pump 1 refills from source (fast),
            #          Pump 2 dispenses to destination (fast)
            print(f"[Cavro COORD] cycle {n} phase B: "
                  f"P{idx1+1} refill from {from1}, P{idx2+1} dispense to {to2}")
            pump1.start_batch()
            pump1.switch_valve(from1)
            pump1.move_absolute_counts(
                int(syringe_ml * pump1.counts_per_ml),
                velocity_counts=max_vel, wait=False)
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

        # Final flush: inlet pump back to source, outlet pump out to destination
        flush1 = pump1.volume_ml > 0.001
        flush2 = pump2.volume_ml > 0.001
        if flush1:
            print(f"[Cavro COORD] final flush P{idx1+1}: "
                  f"{pump1.volume_ml:.3f} mL back to port {from1}")
            pump1.start_batch()
            pump1.switch_valve(from1)
            pump1.move_absolute_counts(0, velocity_counts=max_vel, wait=False)
        if flush2:
            print(f"[Cavro COORD] final flush P{idx2+1}: "
                  f"{pump2.volume_ml:.3f} mL out to port {to2}")
            pump2.start_batch()
            pump2.switch_valve(to2)
            pump2.move_absolute_counts(0, velocity_counts=max_vel, wait=False)
        flush_pumps = [p for p, f in [(pump1, flush1), (pump2, flush2)] if f]
        if flush_pumps:
            TecanCavro.broadcast_execute(*flush_pumps)
            TecanCavro.wait_for_all(*flush_pumps)
        print(f"[Cavro COORD] final flush done")

        status = "stopped" if _coord_stop.is_set() else "complete"
        total = n * syringe_ml
        push_event("onCavroCoordinatedUpdate", status,
                   f"Coordinated {status}: {n} cycles, {total:.2f} mL")
    except Exception as e:
        if ready:
            ready.set()
        for p in (pump1, pump2):
            try:
                p.halt()
            except Exception:
                pass
        push_event("onCavroCoordinatedUpdate", "error",
                   f"Coordinated error: {e}")


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
