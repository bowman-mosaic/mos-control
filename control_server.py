#!/usr/bin/env python3
"""
MOS-11 Modular Control System
──────────────────────────────
Unified web interface for synchronised syringe-pump and camera control.

Usage:
    python control_server.py                     # localhost:8080, opens no browser
    python control_server.py --port 9000         # custom port
    python control_server.py --host 0.0.0.0      # accessible over LAN / VPN
"""

import eel
import os
import sys
import argparse

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import modules.pumps       # noqa: F401, E402 – registers @eel.expose functions
import modules.camera      # noqa: F401, E402
import modules.experiment  # noqa: F401, E402

WEB_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "web")
eel.init(WEB_DIR)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="MOS-11 Control System")
    parser.add_argument("--host", default="0.0.0.0",
                        help="Listen address (default 0.0.0.0 for LAN access)")
    parser.add_argument("--port", type=int, default=808 ,
                        help="Listen port (default 8080)")
    args = parser.parse_args()

    print(f"MOS-11 Control System")
    print(f"  Local:   http://localhost:{args.port}")
    print(f"  Network: http://<this-machine-ip>:{args.port}")
    eel.start("index.html", host=args.host, port=args.port,
              mode=None, block=True)
