"""
Persistent configuration store.

Saves all user configuration (presets, timeline, OBSBOT waypoints/crops) as
JSON files in the ``config/`` directory next to the server.  The frontend
syncs here on every change and loads on startup, replacing localStorage as
the source of truth.
"""

import os
import json
import logging

from modules._api import expose

log = logging.getLogger(__name__)

CONFIG_DIR = os.path.join(
    os.path.normpath(os.path.join(os.path.dirname(__file__), "..")),
    "config",
)
os.makedirs(CONFIG_DIR, exist_ok=True)

_ALLOWED_KEYS = {
    "microscope_presets",
    "acquisition_presets",
    "timeline_events",
    "timeline_lanes",
    "timeline_settings",
    "obsbot_waypoints",
    "obsbot_crops",
}


def _path(key: str) -> str:
    return os.path.join(CONFIG_DIR, f"{key}.json")


def load(key: str):
    """Read a config key from disk. Returns None if missing."""
    p = _path(key)
    if not os.path.exists(p):
        return None
    with open(p, "r", encoding="utf-8") as f:
        return json.load(f)


def save(key: str, data):
    """Write a config key to disk."""
    with open(_path(key), "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)


@expose
def config_save(key, data):
    """Save a configuration key."""
    if key not in _ALLOWED_KEYS:
        return {"error": f"Unknown config key: {key}"}
    try:
        save(key, data)
        return {"ok": True}
    except Exception as e:
        log.error("config_save(%s) failed: %s", key, e)
        return {"error": str(e)}


@expose
def config_load(key):
    """Load a configuration key."""
    if key not in _ALLOWED_KEYS:
        return {"error": f"Unknown config key: {key}"}
    try:
        data = load(key)
        return {"ok": True, "data": data}
    except Exception as e:
        log.error("config_load(%s) failed: %s", key, e)
        return {"error": str(e)}


@expose
def config_load_all():
    """Load all configuration keys at once (for startup)."""
    result = {}
    for key in _ALLOWED_KEYS:
        try:
            result[key] = load(key)
        except Exception as e:
            log.error("config_load_all key %s failed: %s", key, e)
            result[key] = None
    return {"ok": True, "data": result}
