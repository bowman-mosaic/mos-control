"""
Experiment utilities — save / load experiment configurations and coordinated
stop-all.  Protocol execution is handled independently by modules.pumps
(pump_run_protocol) and modules.coolsnap (camera capture); the frontend
starts both at the same time to synchronise them.
"""

from modules._api import expose
import json
import os
from datetime import datetime

_EXPERIMENTS_DIR = os.path.join(os.path.dirname(__file__), "..", "experiments")


@expose
def experiment_save(name, payload):
    """Save an experiment configuration (pump + camera protocols) to JSON."""
    os.makedirs(_EXPERIMENTS_DIR, exist_ok=True)
    safe = "".join(c if c.isalnum() or c in "-_ " else "_" for c in name)
    path = os.path.join(_EXPERIMENTS_DIR, f"{safe}.json")
    with open(path, "w") as f:
        json.dump({"name": name, "payload": payload,
                   "saved": datetime.now().isoformat()}, f, indent=2)
    return {"ok": True, "path": path}


@expose
def experiment_load(name):
    safe = "".join(c if c.isalnum() or c in "-_ " else "_" for c in name)
    path = os.path.join(_EXPERIMENTS_DIR, f"{safe}.json")
    if not os.path.exists(path):
        return {"error": f"Experiment '{name}' not found"}
    with open(path) as f:
        data = json.load(f)
    return {"ok": True, "data": data}


@expose
def experiment_list_saved():
    if not os.path.isdir(_EXPERIMENTS_DIR):
        return []
    return sorted(f[:-5] for f in os.listdir(_EXPERIMENTS_DIR)
                  if f.endswith(".json"))


@expose
def experiment_delete_saved(name):
    safe = "".join(c if c.isalnum() or c in "-_ " else "_" for c in name)
    path = os.path.join(_EXPERIMENTS_DIR, f"{safe}.json")
    if os.path.exists(path):
        os.remove(path)
        return {"ok": True}
    return {"error": "Not found"}


@expose
def experiment_stop_all():
    """Stop every running pump protocol and camera capture."""
    from modules import pumps as _p
    from modules import coolsnap as _cs
    for i in range(4):
        try:
            _p.pump_stop_protocol(i)
        except Exception:
            pass
    try:
        _cs.live_stop()
        _cs.capture_stop()
    except Exception:
        pass
    return {"ok": True}
