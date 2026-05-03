"""
Persistent configuration store.

Modular configuration system:
- Each module (imaging, timeline, cavro, obsbot) has its own folder under
  ``config/modules/<module>/`` with per-name JSON files.
- A master config under ``config/masters/<name>.json`` references module
  configs by name.
"""

import os
import json
import logging
from datetime import datetime

from modules._api import expose

log = logging.getLogger(__name__)

CONFIG_DIR = os.path.join(
    os.path.normpath(os.path.join(os.path.dirname(__file__), "..")),
    "config",
)
os.makedirs(CONFIG_DIR, exist_ok=True)

MODULES_DIR = os.path.join(CONFIG_DIR, "modules")
MASTERS_DIR = os.path.join(CONFIG_DIR, "masters")
os.makedirs(MODULES_DIR, exist_ok=True)
os.makedirs(MASTERS_DIR, exist_ok=True)

MODULES = ("imaging", "timeline", "cavro", "obsbot")
for _m in MODULES:
    os.makedirs(os.path.join(MODULES_DIR, _m), exist_ok=True)


def _safe_name(name: str) -> str:
    return "".join(c if c.isalnum() or c in "-_ " else "_" for c in name)


def _check_module(module: str):
    if module not in MODULES:
        return {"error": f"Unknown module '{module}' (valid: {', '.join(MODULES)})"}
    return None


def load(key: str):
    """Read a top-level config key from disk (used by obsbot for waypoints/crops)."""
    p = os.path.join(CONFIG_DIR, f"{key}.json")
    if not os.path.exists(p):
        return None
    with open(p, "r", encoding="utf-8") as f:
        return json.load(f)


def save(key: str, data):
    """Write a top-level config key to disk (used by obsbot for waypoints/crops)."""
    with open(os.path.join(CONFIG_DIR, f"{key}.json"), "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)


# ── Module configs ──────────────────────────────────────────────────────────

@expose
def config_module_list(module):
    """List saved configs for a given module."""
    err = _check_module(module)
    if err:
        return err
    folder = os.path.join(MODULES_DIR, module)
    if not os.path.isdir(folder):
        return {"ok": True, "configs": []}
    configs = []
    for fn in sorted(os.listdir(folder)):
        if not fn.endswith(".json"):
            continue
        path = os.path.join(folder, fn)
        try:
            with open(path, "r", encoding="utf-8") as fh:
                data = json.load(fh)
            configs.append({
                "name": data.get("name", fn[:-5]),
                "saved": data.get("saved", ""),
            })
        except Exception:
            configs.append({"name": fn[:-5], "saved": ""})
    return {"ok": True, "configs": configs}


@expose
def config_module_save(module, name, payload):
    """Save a module config under the given name."""
    err = _check_module(module)
    if err:
        return err
    if not name or not str(name).strip():
        return {"error": "Configuration name is required"}
    name = str(name).strip()
    safe = _safe_name(name)
    folder = os.path.join(MODULES_DIR, module)
    os.makedirs(folder, exist_ok=True)
    path = os.path.join(folder, f"{safe}.json")
    try:
        data = {
            "name": name,
            "module": module,
            "saved": datetime.now().isoformat(),
            "payload": payload,
        }
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)
        return {"ok": True, "path": path}
    except Exception as e:
        log.error("config_module_save(%s, %s) failed: %s", module, name, e)
        return {"error": str(e)}


@expose
def config_module_load(module, name):
    """Load a single module config by name."""
    err = _check_module(module)
    if err:
        return err
    safe = _safe_name(str(name))
    path = os.path.join(MODULES_DIR, module, f"{safe}.json")
    if not os.path.exists(path):
        return {"error": f"{module} config '{name}' not found"}
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        return {"ok": True, "data": data}
    except Exception as e:
        log.error("config_module_load(%s, %s) failed: %s", module, name, e)
        return {"error": str(e)}


@expose
def config_module_delete(module, name):
    """Delete a module config by name."""
    err = _check_module(module)
    if err:
        return err
    safe = _safe_name(str(name))
    path = os.path.join(MODULES_DIR, module, f"{safe}.json")
    if os.path.exists(path):
        try:
            os.remove(path)
            return {"ok": True}
        except Exception as e:
            return {"error": str(e)}
    return {"error": "Not found"}


# ── Master configs ──────────────────────────────────────────────────────────

@expose
def config_master_list():
    """List all saved master configurations."""
    if not os.path.isdir(MASTERS_DIR):
        return {"ok": True, "masters": []}
    masters = []
    for fn in sorted(os.listdir(MASTERS_DIR)):
        if not fn.endswith(".json"):
            continue
        path = os.path.join(MASTERS_DIR, fn)
        try:
            with open(path, "r", encoding="utf-8") as fh:
                data = json.load(fh)
            masters.append({
                "name": data.get("name", fn[:-5]),
                "saved": data.get("saved", ""),
                "modules": data.get("modules", {}) or {},
            })
        except Exception:
            masters.append({"name": fn[:-5], "saved": "", "modules": {}})
    return {"ok": True, "masters": masters}


@expose
def config_master_save(name, module_refs):
    """Save a master config. module_refs: {module_name: config_name|None}."""
    if not name or not str(name).strip():
        return {"error": "Master name is required"}
    name = str(name).strip()
    safe = _safe_name(name)
    os.makedirs(MASTERS_DIR, exist_ok=True)
    path = os.path.join(MASTERS_DIR, f"{safe}.json")
    refs = {}
    if isinstance(module_refs, dict):
        for m in MODULES:
            v = module_refs.get(m)
            if v is None or v == "":
                refs[m] = None
            else:
                refs[m] = str(v).strip() or None
    try:
        data = {
            "name": name,
            "saved": datetime.now().isoformat(),
            "modules": refs,
        }
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)
        return {"ok": True, "path": path}
    except Exception as e:
        log.error("config_master_save(%s) failed: %s", name, e)
        return {"error": str(e)}


@expose
def config_master_load(name):
    """Load a master config + hydrate each referenced module's payload.

    Returns:
        {"ok": True, "master": {...}, "modules": {module: {name, payload} | None}}
        Missing/broken refs are returned as null module entries with an error string.
    """
    safe = _safe_name(str(name))
    path = os.path.join(MASTERS_DIR, f"{safe}.json")
    if not os.path.exists(path):
        return {"error": f"Master '{name}' not found"}
    try:
        with open(path, "r", encoding="utf-8") as f:
            master = json.load(f)
    except Exception as e:
        return {"error": str(e)}

    hydrated = {}
    refs = master.get("modules", {}) or {}
    for m in MODULES:
        ref_name = refs.get(m)
        if not ref_name:
            hydrated[m] = None
            continue
        ref_safe = _safe_name(ref_name)
        mpath = os.path.join(MODULES_DIR, m, f"{ref_safe}.json")
        if not os.path.exists(mpath):
            hydrated[m] = {"name": ref_name, "error": "missing"}
            continue
        try:
            with open(mpath, "r", encoding="utf-8") as fh:
                mdata = json.load(fh)
            hydrated[m] = {
                "name": mdata.get("name", ref_name),
                "payload": mdata.get("payload"),
            }
        except Exception as e:
            hydrated[m] = {"name": ref_name, "error": str(e)}
    return {"ok": True, "master": master, "modules": hydrated}


@expose
def config_master_delete(name):
    """Delete a master config (module configs are untouched)."""
    safe = _safe_name(str(name))
    path = os.path.join(MASTERS_DIR, f"{safe}.json")
    if os.path.exists(path):
        try:
            os.remove(path)
            return {"ok": True}
        except Exception as e:
            return {"error": str(e)}
    return {"error": "Not found"}
