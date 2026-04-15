"""
Pump control module — exposes Harvard Apparatus syringe pump operations via
Flask API.  Wraps syringe_pump/syringe_pump_control.py without modifying it.
"""

from modules._api import expose, push_event
import threading
import time
import sys
import os

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "syringe_pump"))
from syringe_pump_control import HarvardPump, list_serial_ports  # noqa: E402

NUM_PUMPS = 4

_pumps = [None] * NUM_PUMPS
_proto_threads = [None] * NUM_PUMPS
_proto_stops = [threading.Event() for _ in range(NUM_PUMPS)]


def get_pump(idx):
    """Return a pump instance by index (used by the experiment engine)."""
    if 0 <= idx < NUM_PUMPS:
        return _pumps[idx]
    return None


def _hms_to_seconds(hms):
    if not hms or hms == "00:00:00":
        return 0
    parts = hms.split(":")
    return int(parts[0]) * 3600 + int(parts[1]) * 60 + int(parts[2])


# ── Port discovery ──────────────────────────────────────────────────────────

@expose
def pump_get_ports():
    try:
        return list_serial_ports()
    except Exception as e:
        return {"error": str(e)}


# ── Connection ──────────────────────────────────────────────────────────────

@expose
def pump_connect(idx, port, address):
    if idx < 0 or idx >= NUM_PUMPS:
        return {"error": "Invalid pump index"}
    try:
        _pumps[idx] = HarvardPump(port=port, address=int(address))
        return {"ok": True, "msg": f"Connected to {port} addr {address}"}
    except Exception as e:
        _pumps[idx] = None
        return {"error": str(e)}


@expose
def pump_disconnect(idx):
    if idx < 0 or idx >= NUM_PUMPS:
        return {"error": "Invalid pump index"}
    _stop_protocol(idx)
    if _pumps[idx]:
        try:
            _pumps[idx].close()
        except Exception:
            pass
        _pumps[idx] = None
    return {"ok": True}


@expose
def pump_is_connected(idx):
    return _pumps[idx] is not None


# ── Settings ────────────────────────────────────────────────────────────────

@expose
def pump_set_diameter(idx, diameter_mm):
    if not _pumps[idx]:
        return {"error": "Not connected"}
    try:
        _pumps[idx].set_diameter(float(diameter_mm))
        return {"ok": True}
    except Exception as e:
        return {"error": str(e)}


@expose
def pump_set_direction(idx, direction):
    if not _pumps[idx]:
        return {"error": "Not connected"}
    try:
        _pumps[idx].set_direction(direction)
        return {"ok": True}
    except Exception as e:
        return {"error": str(e)}


# ── Manual controls ─────────────────────────────────────────────────────────

@expose
def pump_set_rate(idx, rate, units):
    if not _pumps[idx]:
        return {"error": "Not connected"}
    try:
        _pumps[idx].set_rate(float(rate), units)
        return {"ok": True}
    except Exception as e:
        return {"error": str(e)}


@expose
def pump_set_volume(idx, volume):
    if not _pumps[idx]:
        return {"error": "Not connected"}
    try:
        _pumps[idx].set_volume(float(volume))
        return {"ok": True}
    except Exception as e:
        return {"error": str(e)}


@expose
def pump_run(idx):
    if not _pumps[idx]:
        return {"error": "Not connected"}
    try:
        _pumps[idx].run()
        return {"ok": True}
    except Exception as e:
        return {"error": str(e)}


@expose
def pump_stop(idx):
    if not _pumps[idx]:
        return {"error": "Not connected"}
    try:
        _pumps[idx].stop()
        return {"ok": True}
    except Exception as e:
        return {"error": str(e)}


@expose
def pump_clear_volume(idx):
    if not _pumps[idx]:
        return {"error": "Not connected"}
    try:
        _pumps[idx].clear_volume()
        return {"ok": True}
    except Exception as e:
        return {"error": str(e)}


@expose
def pump_clear_target(idx):
    if not _pumps[idx]:
        return {"error": "Not connected"}
    try:
        _pumps[idx].clear_target()
        return {"ok": True}
    except Exception as e:
        return {"error": str(e)}


@expose
def pump_get_status(idx):
    if not _pumps[idx]:
        return {"error": "Not connected"}
    try:
        s = _pumps[idx].get_status()
        return {"ok": True, "status": s}
    except Exception as e:
        return {"error": str(e)}


# ── Protocol execution ──────────────────────────────────────────────────────

@expose
def pump_run_protocol(idx, steps):
    """Run a multi-step pump protocol.
    steps: list of {action, rate, units, time} dicts.
    """
    if not _pumps[idx]:
        return {"error": "Not connected"}
    if _proto_threads[idx] and _proto_threads[idx].is_alive():
        return {"error": "Protocol already running"}
    _proto_stops[idx].clear()
    t = threading.Thread(target=_run_protocol_thread, args=(idx, steps), daemon=True)
    _proto_threads[idx] = t
    t.start()
    return {"ok": True}


@expose
def pump_stop_protocol(idx):
    _stop_protocol(idx)
    return {"ok": True}


def _stop_protocol(idx):
    _proto_stops[idx].set()
    if _pumps[idx]:
        try:
            _pumps[idx].stop()
        except Exception:
            pass
    if _proto_threads[idx] and _proto_threads[idx].is_alive():
        _proto_threads[idx].join(timeout=3)


def _run_protocol_thread(idx, steps):
    try:
        push_event("onPumpProtocolUpdate", idx, -1, "started",
                   f"Protocol started ({len(steps)} steps)")
        for i, step in enumerate(steps):
            if _proto_stops[idx].is_set():
                break
            action = step.get("action", "")
            rate = step.get("rate", "")
            units = step.get("units", "")
            time_s = step.get("time", "00:00:00")

            if action == "Run":
                push_event("onPumpProtocolUpdate", idx, i, "running",
                           f"Step {i+1}: Run {rate} {units}")
                _pumps[idx].set_rate(float(rate), units)
                _pumps[idx].run()
            elif action == "Stop":
                push_event("onPumpProtocolUpdate", idx, i, "running",
                           f"Step {i+1}: Stop")
                _pumps[idx].stop()

            wait = _hms_to_seconds(time_s)
            if wait > 0:
                end = time.monotonic() + wait
                while time.monotonic() < end:
                    if _proto_stops[idx].is_set():
                        break
                    remaining = end - time.monotonic()
                    rm = int(remaining // 60)
                    rs = int(remaining % 60)
                    push_event("onPumpProtocolCountdown", idx, i, f"{rm:02d}:{rs:02d}")
                    time.sleep(1)

        status = "aborted" if _proto_stops[idx].is_set() else "complete"
        push_event("onPumpProtocolUpdate", idx, -1, status, f"Protocol {status}")
    except Exception as e:
        try:
            _pumps[idx].stop()
        except Exception:
            pass
        push_event("onPumpProtocolUpdate", idx, -1, "error", f"Protocol error: {e}")
