#!/usr/bin/env python3
"""
Standalone web server for controlling 4x Harvard Apparatus Model 22 syringe pumps.
No microscope or camera dependencies.

Usage:
    python pump_server.py                # http://localhost:5000
    python pump_server.py --port 9000    # custom port
"""

import os
import sys
import json
import time
import threading
import argparse

from flask import Flask, request, jsonify, send_file

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from syringe_pump_control import HarvardPump, list_serial_ports

app = Flask(__name__, static_folder=None)

import logging
logging.getLogger("werkzeug").setLevel(logging.ERROR)

NUM_PUMPS = 4
_pumps = [None] * NUM_PUMPS
_proto_threads = [None] * NUM_PUMPS
_proto_stops = [threading.Event() for _ in range(NUM_PUMPS)]
_lock = threading.Lock()

# Server-sent events for protocol progress
_events = []
_event_id = 0
_event_lock = threading.Lock()


def _push_event(name, *args):
    global _event_id
    with _event_lock:
        _event_id += 1
        _events.append({"id": _event_id, "name": name, "args": list(args)})
        if len(_events) > 200:
            _events.pop(0)


# ── Static files ─────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return send_file(os.path.join(os.path.dirname(__file__), "pump_ui.html"))


@app.route("/api/events")
def events():
    since = int(request.args.get("since", 0))
    with _event_lock:
        return jsonify([e for e in _events if e["id"] > since])


# ── Port discovery ───────────────────────────────────────────────────────────

@app.route("/api/ports", methods=["POST"])
def get_ports():
    try:
        return jsonify(list_serial_ports())
    except Exception as e:
        return jsonify({"error": str(e)})


# ── Pump connection ──────────────────────────────────────────────────────────

@app.route("/api/connect", methods=["POST"])
def pump_connect():
    d = request.get_json(silent=True) or {}
    idx = int(d.get("pump", 0))
    port = d.get("port", "")
    addr = int(d.get("address", idx))
    if not 0 <= idx < NUM_PUMPS:
        return jsonify({"error": f"Invalid pump index {idx}"})
    if not port:
        return jsonify({"error": "No port specified"})
    try:
        with _lock:
            if _pumps[idx]:
                try:
                    _pumps[idx].close()
                except Exception:
                    pass
            _pumps[idx] = HarvardPump(port=port, address=addr)
        return jsonify({"ok": True})
    except Exception as e:
        return jsonify({"error": str(e)})


@app.route("/api/disconnect", methods=["POST"])
def pump_disconnect():
    d = request.get_json(silent=True) or {}
    idx = int(d.get("pump", 0))
    _proto_stops[idx].set()
    with _lock:
        if _pumps[idx]:
            try:
                _pumps[idx].close()
            except Exception:
                pass
            _pumps[idx] = None
    return jsonify({"ok": True})


# ── Pump commands ────────────────────────────────────────────────────────────

def _cmd(idx, fn_name, *args):
    if not 0 <= idx < NUM_PUMPS:
        return {"error": f"Invalid pump {idx}"}
    p = _pumps[idx]
    if not p:
        return {"error": f"Pump {idx+1} not connected"}
    try:
        fn = getattr(p, fn_name)
        result = fn(*args)
        return {"ok": True, "response": result}
    except Exception as e:
        return {"error": str(e)}


@app.route("/api/set_diameter", methods=["POST"])
def set_diameter():
    d = request.get_json(silent=True) or {}
    return jsonify(_cmd(int(d["pump"]), "set_diameter", float(d["value"])))


@app.route("/api/set_direction", methods=["POST"])
def set_direction():
    d = request.get_json(silent=True) or {}
    return jsonify(_cmd(int(d["pump"]), "set_direction", d["value"]))


@app.route("/api/set_rate", methods=["POST"])
def set_rate():
    d = request.get_json(silent=True) or {}
    return jsonify(_cmd(int(d["pump"]), "set_rate", float(d["rate"]), d["units"]))


@app.route("/api/set_volume", methods=["POST"])
def set_volume():
    d = request.get_json(silent=True) or {}
    return jsonify(_cmd(int(d["pump"]), "set_volume", float(d["value"])))


@app.route("/api/run", methods=["POST"])
def pump_run():
    d = request.get_json(silent=True) or {}
    return jsonify(_cmd(int(d["pump"]), "run"))


@app.route("/api/stop", methods=["POST"])
def pump_stop():
    d = request.get_json(silent=True) or {}
    return jsonify(_cmd(int(d["pump"]), "stop"))


@app.route("/api/clear_volume", methods=["POST"])
def clear_volume():
    d = request.get_json(silent=True) or {}
    return jsonify(_cmd(int(d["pump"]), "clear_volume"))


@app.route("/api/clear_target", methods=["POST"])
def clear_target():
    d = request.get_json(silent=True) or {}
    return jsonify(_cmd(int(d["pump"]), "clear_target"))


@app.route("/api/status", methods=["POST"])
def pump_status():
    d = request.get_json(silent=True) or {}
    return jsonify(_cmd(int(d["pump"]), "get_status"))


# ── Protocol execution ───────────────────────────────────────────────────────

@app.route("/api/proto_run", methods=["POST"])
def proto_run():
    d = request.get_json(silent=True) or {}
    idx = int(d["pump"])
    steps = d.get("steps", [])
    if not 0 <= idx < NUM_PUMPS:
        return jsonify({"error": "Invalid pump"})
    if not _pumps[idx]:
        return jsonify({"error": f"Pump {idx+1} not connected"})
    if not steps:
        return jsonify({"error": "Protocol is empty"})
    if _proto_threads[idx] and _proto_threads[idx].is_alive():
        return jsonify({"error": "Protocol already running"})
    _proto_stops[idx].clear()
    _proto_threads[idx] = threading.Thread(
        target=_proto_worker, args=(idx, steps), daemon=True)
    _proto_threads[idx].start()
    return jsonify({"ok": True})


@app.route("/api/proto_stop", methods=["POST"])
def proto_stop():
    d = request.get_json(silent=True) or {}
    idx = int(d["pump"])
    _proto_stops[idx].set()
    if _pumps[idx]:
        try:
            _pumps[idx].stop()
        except Exception:
            pass
    return jsonify({"ok": True})


def _hms_to_sec(hms):
    if not hms or hms == "00:00:00":
        return 0
    parts = hms.split(":")
    return int(parts[0]) * 3600 + int(parts[1]) * 60 + int(parts[2])


def _proto_worker(idx, steps):
    label = f"Pump {idx+1}"
    _push_event("proto_update", idx, 0, "running", f"{label}: protocol started ({len(steps)} steps)")
    try:
        for i, step in enumerate(steps):
            if _proto_stops[idx].is_set():
                break
            action = step.get("action", "Stop")
            rate = step.get("rate", "")
            units = step.get("units", "ML/MIN")
            time_s = step.get("time", "00:00:00")
            p = _pumps[idx]
            if not p:
                break

            if action == "Run" and rate:
                p.set_rate(float(rate), units)
                p.run()
                _push_event("proto_update", idx, i, "running", f"{label}: step {i+1} — Run {rate} {units}")
            else:
                p.stop()
                _push_event("proto_update", idx, i, "running", f"{label}: step {i+1} — Stop")

            wait = _hms_to_sec(time_s)
            if wait > 0:
                end = time.monotonic() + wait
                while time.monotonic() < end:
                    if _proto_stops[idx].is_set():
                        break
                    remaining = int(end - time.monotonic())
                    m, s = divmod(max(0, remaining), 60)
                    h, m = divmod(m, 60)
                    _push_event("proto_countdown", idx, i, f"{h:02d}:{m:02d}:{s:02d}")
                    time.sleep(1)

        status = "aborted" if _proto_stops[idx].is_set() else "complete"
        _push_event("proto_update", idx, -1, status, f"{label}: protocol {status}")
    except Exception as e:
        _push_event("proto_update", idx, -1, "error", f"{label}: error — {e}")
        if _pumps[idx]:
            try:
                _pumps[idx].stop()
            except Exception:
                pass


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Syringe Pump Control Server")
    parser.add_argument("--port", type=int, default=5000)
    parser.add_argument("--host", default="0.0.0.0")
    args = parser.parse_args()
    print(f"Syringe Pump Control")
    print(f"  http://localhost:{args.port}")
    app.run(host=args.host, port=args.port, threaded=True)
