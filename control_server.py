#!/usr/bin/env python3
"""
MOS-11 Modular Control System
──────────────────────────────
Flask-based web server with WebSocket live-view streaming.
Uses real OS threads (no gevent) so blocking hardware drivers never stall HTTP.

Usage:
    python control_server.py                     # localhost:8080
    python control_server.py --port 9000         # custom port
    python control_server.py --host 0.0.0.0      # accessible over LAN
"""

import os
import sys
import signal
import atexit
import argparse

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from modules._api import app
import modules.pumps                   # noqa: F401
import modules.experiment              # noqa: F401
import modules.nikon_ti                # noqa: F401
import modules.coolsnap                # noqa: F401
import modules.intensilight            # noqa: F401

from flask import Response, jsonify
from flask_sock import Sock
import modules.coolsnap as _cs

sock = Sock(app)


# ── WebSocket live stream ───────────────────────────────────────────────────
# Single persistent connection, binary JPEG frames pushed as fast as the
# camera produces them.  No HTTP overhead per frame, no base64.

@sock.route("/cam/live")
def cam_live_ws(ws):
    """Stream binary JPEG frames over WebSocket."""
    prev_id = None
    while True:
        if not _cs.live_is_active():
            _cs._frame_event.wait(timeout=1.0)
            if not _cs.live_is_active():
                break

        _cs._frame_event.wait(timeout=0.5)
        _cs._frame_event.clear()

        jpeg = _cs.get_live_jpeg()
        if jpeg is None or jpeg is prev_id:
            continue
        prev_id = jpeg
        try:
            ws.send(jpeg)
        except Exception:
            break


# ── HTTP fallback endpoints (snap preview, FPS readout) ─────────────────────

@app.route("/cam/frame")
def cam_frame():
    """Return the latest JPEG frame (used by snap, not live view)."""
    jpeg = _cs.get_live_jpeg()
    if jpeg is None:
        return Response(status=204)
    return Response(jpeg, mimetype="image/jpeg",
                    headers={"Cache-Control": "no-store",
                             "Content-Length": str(len(jpeg))})


@app.route("/cam/fps")
def cam_fps():
    return jsonify({"fps": _cs.get_live_fps(),
                    "active": _cs.live_is_active()})


# ── Graceful shutdown ────────────────────────────────────────────────────────

import modules.nikon_ti as _ti
import modules.coolsnap as _cs_mod
import modules.intensilight as _il
import modules.pumps as _pumps_mod

_shutdown_done = False

def _shutdown():
    global _shutdown_done
    if _shutdown_done:
        return
    _shutdown_done = True
    print("\nShutting down — disconnecting hardware...")

    for name, fn in [
        ("CoolSNAP",     _cs_mod.disconnect),
        ("Nikon Ti",     _ti.disconnect),
        ("IntensiLight", _il.disconnect),
    ]:
        try:
            fn()
            print(f"  {name} disconnected")
        except Exception as e:
            print(f"  {name} disconnect failed: {e}")

    for i in range(_pumps_mod.NUM_PUMPS):
        if _pumps_mod._pumps[i] is not None:
            try:
                _pumps_mod._pumps[i].close()
                _pumps_mod._pumps[i] = None
                print(f"  Pump {i} disconnected")
            except Exception as e:
                print(f"  Pump {i} disconnect failed: {e}")

    print("Shutdown complete.")

atexit.register(_shutdown)
signal.signal(signal.SIGINT, lambda *_: sys.exit(0))
signal.signal(signal.SIGTERM, lambda *_: sys.exit(0))


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="MOS-11 Control System")
    parser.add_argument("--host", default="0.0.0.0",
                        help="Listen address (default 0.0.0.0)")
    parser.add_argument("--port", type=int, default=8081,
                        help="Listen port (default 8081)")
    args = parser.parse_args()

    print("MOS-11 Control System")
    print(f"  Local:   http://localhost:{args.port}")
    print(f"  Network: http://<this-machine-ip>:{args.port}")
    app.run(host=args.host, port=args.port, threaded=True)
