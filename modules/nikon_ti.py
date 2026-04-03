"""
Nikon Eclipse Ti microscope control module — drives the microscope body and all
attached peripherals through the Nikon Ti SDK COM interface.

Architecture (discovered via probe):
  - Create ONE main scope object via Nikon.TiScope.NikonTi
  - All devices are sub-properties: scope.DiaShutter, scope.DiaLamp, etc.
  - Numeric values are IMipParameter objects — read/write via .RawValue
  - Drives expose MoveAbsolute(val) and MoveRelative(val) methods
  - Shutters expose Open() and Close() methods
  - Drive units are in nm (nanometers)

Hardware mapping (inverted microscope — "dia"/transmitted comes from top):
  DiaShutter          Sutter Lambda SC SmartShutter (transmitted light, top)
  EpiShutter          Not mounted on this system (IsMounted=0)
  DiaLamp             Halogen lamp (bottom) — intensity + on/off
  Nosepiece           Objective turret (positions 1–6)
  FilterBlockCassette1  Fluorescence filter cassette
  LightPathDrive      Eyepiece / port selector (manual control only)
  ZDrive              Focus (nm)
  XDrive / YDrive     Stage (nm)
  PFS                 Perfect Focus System
"""

from modules._api import expose
import threading
import queue
import ctypes

_HAS_COMTYPES = True
try:
    import importlib
    if importlib.util.find_spec("comtypes") is None:
        _HAS_COMTYPES = False
except Exception:
    _HAS_COMTYPES = False

# ── Module state ─────────────────────────────────────────────────────────────

_scope = None
_com_thread = None
_cmd_queue = queue.Queue()
_dia_lamp_intensity = 2


class TiError(Exception):
    """Raised when a Nikon Ti operation fails."""


# ── Dedicated COM worker (MTA) ──────────────────────────────────────────────
# Nikon Ti COM objects cannot be shared across threads.  All creation and
# access must happen on the SAME thread.  We use a dedicated worker thread
# that initializes COM (MTA), creates the scope, then processes commands.

def _com_worker(ready_event):
    global _scope
    import logging
    log = logging.getLogger("nikon_ti.com")
    log.info("COM worker starting on thread %s", threading.current_thread().ident)
    ctypes.windll.ole32.CoInitializeEx(None, 2)  # COINIT_MULTITHREADED
    try:
        import comtypes.client
        _scope = comtypes.client.CreateObject("Nikon.TiScope.NikonTi")
        log.info("Scope created OK: %r", _scope)
    except Exception as e:
        log.error("Scope creation failed: %s", e)
        ready_event._error = e
        ready_event.set()
        return
    ready_event._error = None
    ready_event.set()

    try:
        while True:
            item = _cmd_queue.get()
            if item is None:
                log.info("COM worker received shutdown signal")
                break
            fn, result_evt, holder = item
            try:
                holder["value"] = fn()
            except Exception as e:
                log.warning("COM call error: %s", e)
                holder["error"] = e
            result_evt.set()
    except BaseException as e:
        log.error("COM worker thread crashed: %s: %s", type(e).__name__, e, exc_info=True)
    finally:
        log.info("COM worker thread exiting")
        _scope = None
        ctypes.windll.ole32.CoUninitialize()


def _com_call(fn):
    """Dispatch fn() to the COM worker thread and block until done."""
    if _com_thread is None or not _com_thread.is_alive():
        try:
            connect()
        except Exception:
            raise TiError("Microscope not connected (auto-reconnect failed)")
    holder = {}
    evt = threading.Event()
    _cmd_queue.put((fn, evt, holder))
    if not evt.wait(timeout=15):
        raise TiError("COM call timed out")
    if "error" in holder:
        raise holder["error"]
    return holder.get("value")


# ── Internal helpers ─────────────────────────────────────────────────────────

def _dev(name):
    if _scope is None:
        raise TiError("Microscope not connected")
    try:
        return getattr(_scope, name)
    except AttributeError:
        raise TiError(f"Device '{name}' not found on scope object")


def _read_param(param):
    try:
        return param.RawValue
    except Exception:
        return None


def _write_param(param, value):
    param.RawValue = value


def _ensure_controlled(dev):
    """Enable software control on a device if it has IsControlled."""
    try:
        if dev.IsControlled.RawValue == 0:
            dev.IsControlled.RawValue = 1
    except Exception:
        pass


# ── Connection ───────────────────────────────────────────────────────────────

def connect():
    global _com_thread
    if not _HAS_COMTYPES:
        raise TiError("comtypes is not installed")
    if _com_thread and _com_thread.is_alive():
        return  # already connected — nothing to do
    ready = threading.Event()
    ready._error = None
    _com_thread = threading.Thread(target=_com_worker, args=(ready,), daemon=True)
    _com_thread.start()
    if not ready.wait(timeout=10):
        raise TiError("COM worker failed to start in time")
    if ready._error:
        _com_thread = None
        raise TiError(f"COM init failed: {ready._error}")


def disconnect():
    global _com_thread
    if _com_thread and _com_thread.is_alive():
        _cmd_queue.put(None)
        _com_thread.join(timeout=3)
    _com_thread = None


def is_connected():
    return _scope is not None and _com_thread is not None and _com_thread.is_alive()


def get_system_type():
    return _com_call(lambda: _read_param(_dev("SystemType")))


# ── Lambda SC shutter (DiaShutter — transmitted light, top) ─────────────────

def shutter_open():
    _com_call(lambda: _dev("DiaShutter").Open())

def shutter_close():
    _com_call(lambda: _dev("DiaShutter").Close())

def shutter_get_state():
    return _com_call(lambda: _read_param(_dev("DiaShutter").Value))


# ── Dia lamp (halogen intensity) ────────────────────────────────────────────

def dia_lamp_on():
    def _do():
        dev = _dev("DiaLamp")
        _ensure_controlled(dev)
        dev.On()
    _com_call(_do)

def dia_lamp_off():
    def _do():
        dev = _dev("DiaLamp")
        _ensure_controlled(dev)
        dev.Off()
    _com_call(_do)

def dia_lamp_set_intensity(value):
    v = int(value)
    def _do():
        dev = _dev("DiaLamp")
        _ensure_controlled(dev)
        _write_param(dev.Value, v)
    _com_call(_do)

def dia_lamp_get_intensity():
    return _com_call(lambda: _read_param(_dev("DiaLamp").Value))

def dia_lamp_get_state():
    def _do():
        dev = _dev("DiaLamp")
        return {"on": _read_param(dev.IsOn),
                "intensity": _read_param(dev.Value),
                "lower": _read_param(dev.LowerLimit),
                "upper": _read_param(dev.UpperLimit),
                "controlled": _read_param(dev.IsControlled)}
    return _com_call(_do)


# ── Nosepiece (objectives) ──────────────────────────────────────────────────

def nosepiece_get_position():
    return _com_call(lambda: _read_param(_dev("Nosepiece").Position))

def nosepiece_set_position(pos):
    p = int(pos)
    _com_call(lambda: _write_param(_dev("Nosepiece").Position, p))


# ── Filter block cassette 1 ─────────────────────────────────────────────────

def filter_get_position():
    return _com_call(lambda: _read_param(_dev("FilterBlockCassette1").Position))

def filter_set_position(pos):
    p = int(pos)
    _com_call(lambda: _write_param(_dev("FilterBlockCassette1").Position, p))


# ── Batch preset apply ───────────────────────────────────────────────────────

def apply_preset(objective=None, filter_pos=None, lamp_intensity=None,
                 lamp_on=None, shutter_open=None):
    """Apply multiple microscope settings in a single COM dispatch.

    Reads back positions after setting to ensure hardware has settled.
    """
    global _dia_lamp_intensity
    def _do():
        global _dia_lamp_intensity
        if objective is not None:
            dev = _dev("Nosepiece")
            _ensure_controlled(dev)
            _write_param(dev.Position, int(objective))
        if filter_pos is not None:
            dev = _dev("FilterBlockCassette1")
            _ensure_controlled(dev)
            _write_param(dev.Position, int(filter_pos))
        if lamp_intensity is not None:
            dev = _dev("DiaLamp")
            _ensure_controlled(dev)
            _write_param(dev.Value, int(lamp_intensity))
            _dia_lamp_intensity = int(lamp_intensity)
        if lamp_on is not None:
            dev = _dev("DiaLamp")
            _ensure_controlled(dev)
            if lamp_on:
                dev.On()
            else:
                dev.Off()
        if shutter_open is not None:
            dev = _dev("DiaShutter")
            _ensure_controlled(dev)
            if shutter_open:
                dev.Open()
            else:
                dev.Close()
        if objective is not None:
            _read_param(_dev("Nosepiece").Position)
        if filter_pos is not None:
            _read_param(_dev("FilterBlockCassette1").Position)
    _com_call(_do)


# ── Z drive (focus, units: nm) ──────────────────────────────────────────────

def z_get_position():
    return _com_call(lambda: _read_param(_dev("ZDrive").Position))

def z_move_absolute(nm):
    n = int(nm)
    _com_call(lambda: _dev("ZDrive").MoveAbsolute(n))

def z_move_relative(delta_nm):
    d = int(delta_nm)
    _com_call(lambda: _dev("ZDrive").MoveRelative(d))


# ── XY drive (stage, units: nm) ─────────────────────────────────────────────

def xy_get_position():
    def _do():
        return {"x": _read_param(_dev("XDrive").Position),
                "y": _read_param(_dev("YDrive").Position)}
    return _com_call(_do)

def x_move_absolute(nm):
    n = int(nm)
    _com_call(lambda: _dev("XDrive").MoveAbsolute(n))

def y_move_absolute(nm):
    n = int(nm)
    _com_call(lambda: _dev("YDrive").MoveAbsolute(n))

def x_move_relative(delta_nm):
    d = int(delta_nm)
    _com_call(lambda: _dev("XDrive").MoveRelative(d))

def y_move_relative(delta_nm):
    d = int(delta_nm)
    _com_call(lambda: _dev("YDrive").MoveRelative(d))


# ── PFS (Perfect Focus System) ──────────────────────────────────────────────

def pfs_get_status():
    def _do():
        dev = _dev("PFS")
        return {"value": _read_param(dev.Value), "position": _read_param(dev.Position),
                "status": _read_param(dev.Status), "is_mounted": _read_param(dev.IsMounted)}
    return _com_call(_do)

def pfs_enable():
    _com_call(lambda: _dev("PFS").Enable())

def pfs_disable():
    _com_call(lambda: _dev("PFS").Disable())

def pfs_search():
    _com_call(lambda: _dev("PFS").SearchPosition())


# ── Full status ─────────────────────────────────────────────────────────────

def get_full_status():
    if _scope is None:
        return {"connected": False}

    status = {"connected": True}

    def _try(fn, key):
        try:
            status[key] = fn()
        except Exception as e:
            status[key] = f"error: {e}"

    _try(get_system_type, "system_type")
    _try(shutter_get_state, "shutter")
    _try(dia_lamp_get_state, "dia_lamp")
    _try(nosepiece_get_position, "objective")
    _try(filter_get_position, "filter_block")
    _try(z_get_position, "z_position")
    _try(xy_get_position, "xy_position")
    _try(pfs_get_status, "pfs")
    return status


def probe_device(name):
    """List all COM attributes on a sub-device for debugging."""
    if not is_connected():
        return {"error": "Not connected"}
    def _do():
        dev = _dev(name)
        attrs = []
        for attr in dir(dev):
            if attr.startswith("_"):
                continue
            try:
                val = getattr(dev, attr)
                val_str = repr(val)
                if len(val_str) > 120:
                    val_str = val_str[:120] + "..."
                attrs.append({"name": attr, "value": val_str, "type": type(val).__name__})
            except Exception as ex:
                attrs.append({"name": attr, "error": str(ex)[:100]})
        return {"ok": True, "device": name, "attributes": attrs}
    try:
        return _com_call(_do)
    except Exception as e:
        return {"error": str(e)}


# ── Eel-exposed functions ───────────────────────────────────────────────────

def _wrap(fn, *args, **kwargs):
    """Call fn and return {"ok": True, ...} or {"error": "..."}."""
    try:
        result = fn(*args, **kwargs)
        if isinstance(result, dict):
            return {"ok": True, **result}
        return {"ok": True, "value": result}
    except Exception as e:
        return {"error": str(e)}


@expose
def ti_connect():
    try:
        connect()
        return {"ok": True, "msg": "Connected to Nikon Ti"}
    except Exception as e:
        return {"error": str(e)}


@expose
def ti_disconnect():
    try:
        disconnect()
        return {"ok": True}
    except Exception as e:
        return {"error": str(e)}


@expose
def ti_is_connected():
    return is_connected()


@expose
def ti_shutter_open():
    return _wrap(shutter_open)

@expose
def ti_shutter_close():
    return _wrap(shutter_close)

@expose
def ti_shutter_state():
    return _wrap(shutter_get_state)

@expose
def ti_dia_lamp_on():
    return _wrap(dia_lamp_on)

@expose
def ti_dia_lamp_off():
    return _wrap(dia_lamp_off)

@expose
def ti_dia_lamp_set_intensity(value):
    return _wrap(dia_lamp_set_intensity, value)

@expose
def ti_dia_lamp_state():
    return _wrap(dia_lamp_get_state)

@expose
def ti_nosepiece_get():
    return _wrap(nosepiece_get_position)

@expose
def ti_nosepiece_set(position):
    return _wrap(nosepiece_set_position, position)

@expose
def ti_filter_get():
    return _wrap(filter_get_position)

@expose
def ti_filter_set(position):
    return _wrap(filter_set_position, position)

@expose
def ti_apply_preset(objective=None, filter_pos=None, lamp_intensity=None,
                    lamp_on=None, shutter_open=None):
    return _wrap(apply_preset, objective, filter_pos, lamp_intensity, lamp_on,
                 shutter_open)

@expose
def ti_z_get():
    return _wrap(z_get_position)

@expose
def ti_z_move_abs(nm):
    return _wrap(z_move_absolute, nm)

@expose
def ti_z_move_rel(delta_nm):
    return _wrap(z_move_relative, delta_nm)

@expose
def ti_xy_get():
    return _wrap(xy_get_position)

@expose
def ti_x_move_rel(delta_nm):
    return _wrap(x_move_relative, delta_nm)

@expose
def ti_y_move_rel(delta_nm):
    return _wrap(y_move_relative, delta_nm)

@expose
def ti_pfs_enable():
    return _wrap(pfs_enable)

@expose
def ti_pfs_disable():
    return _wrap(pfs_disable)

@expose
def ti_pfs_status():
    return _wrap(pfs_get_status)

@expose
def ti_status():
    try:
        return get_full_status()
    except Exception as e:
        return {"error": str(e)}

@expose
def ti_probe(device_name):
    return probe_device(device_name)
