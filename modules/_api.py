"""
Shared Flask application and helpers for all MOS-11 modules.

Replaces Eel with plain Flask — no gevent, no monkey-patching, real OS threads.
Each @expose'd function becomes a POST /api/<name> endpoint.
Server→client push uses a lightweight event store polled by the frontend.
"""

from flask import Flask, request, jsonify, send_from_directory
import threading
import time
import os

WEB_DIR = os.path.join(os.path.dirname(__file__), "..", "web")

app = Flask(__name__, static_folder=None)

import logging
logging.basicConfig(level=logging.INFO, format="%(asctime)s %(name)s %(levelname)s: %(message)s")
logging.getLogger("werkzeug").setLevel(logging.WARNING)
logging.getLogger("comtypes").setLevel(logging.WARNING)
logging.getLogger("nikon_ti.com").setLevel(logging.DEBUG)


# ── Static file serving (replaces eel.init) ─────────────────────────────────

@app.route("/")
def _index():
    return send_from_directory(WEB_DIR, "index.html")


@app.route("/<path:filename>")
def _static(filename):
    return send_from_directory(WEB_DIR, filename)


# ── @expose decorator ───────────────────────────────────────────────────────

def expose(fn):
    """Register a function as POST /api/<fn_name>.

    The frontend calls:  await api('fn_name', arg1, arg2, ...)
    which POSTs JSON {"args": [arg1, arg2, ...]} and gets the return value.
    """
    name = fn.__name__

    def wrapper():
        args = []
        try:
            body = request.get_json(silent=True)
            if body and "args" in body:
                args = body["args"]
        except Exception:
            pass
        try:
            result = fn(*args)
            if result is None:
                return jsonify({"ok": True})
            if isinstance(result, (dict, list, bool, int, float, str)):
                return jsonify(result)
            return jsonify({"value": result})
        except Exception as e:
            return jsonify({"error": str(e)})

    wrapper.__name__ = f"api_{name}"
    app.add_url_rule(f"/api/{name}", endpoint=name, view_func=wrapper,
                     methods=["POST"])
    return fn


# ── Event push (server → client) ────────────────────────────────────────────

_events = []
_events_lock = threading.Lock()
_event_counter = 0
_MAX_EVENTS = 200


def push_event(name, *args):
    """Store an event for the frontend to pick up via /api/events polling."""
    global _event_counter
    with _events_lock:
        _event_counter += 1
        _events.append({
            "id": _event_counter,
            "name": name,
            "args": list(args),
        })
        if len(_events) > _MAX_EVENTS:
            _events[:] = _events[-_MAX_EVENTS:]


@app.route("/api/events")
def _get_events():
    since = int(request.args.get("since", 0))
    with _events_lock:
        new = [e for e in _events if e["id"] > since]
    return jsonify(new)
