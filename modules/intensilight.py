"""
Nikon Intensilight C-HGFIE epi-fluorescence illuminator — serial control.

Protocol reverse-engineered from the Micro-Manager Nikon device adapter
(mmCoreAndDevices/DeviceAdapters/Nikon/Nikon.cpp, BSD license).

Serial format:
  TX:  command + CR
  RX:  response + CR LF
  Commands start with 'c' (write) or 'r' (read).
  Responses start with 'a' (success) or 'n' (error).

Commands:
  cSXC1       Open shutter
  cSXC2       Close shutter
  rSXR        Query shutter state  →  aSXR1 (open) / aSXR2 (closed)
  cNDM<1-6>   Set ND filter position
  rNAR        Query ND filter      →  aNAR<1-6>
  rVEN        Query firmware ver.  →  aVEN<version>

ND positions:  1→ND1  2→ND2  3→ND4  4→ND8  5→ND16  6→ND32
"""

from modules._api import expose, push_event
import serial
import threading
import logging

log = logging.getLogger("intensilight")

_lock = threading.Lock()
_ser = None
_port = None

ND_LABELS = {1: "ND1", 2: "ND2", 3: "ND4", 4: "ND8", 5: "ND16", 6: "ND32"}
ND_FROM_LABEL = {v: k for k, v in ND_LABELS.items()}

SERIAL_TIMEOUT = 3.0


class ILError(Exception):
    pass


def _send(cmd):
    """Send a command and return the response string (blocking)."""
    if _ser is None or not _ser.is_open:
        raise ILError("Intensilight not connected")
    with _lock:
        _ser.reset_input_buffer()
        _ser.write((cmd + "\r").encode("ascii"))
        _ser.flush()
        raw = _ser.read_until(b"\r\n")
    resp = raw.decode("ascii", errors="replace").strip()
    if not resp:
        raise ILError(f"No response to '{cmd}'")
    if resp[0] == "n":
        raise ILError(f"Intensilight error: {resp}")
    return resp


# ── Connection ───────────────────────────────────────────────────────────────

def connect(port="COM6"):
    global _ser, _port
    disconnect()
    _ser = serial.Serial(
        port=port,
        baudrate=9600,
        bytesize=serial.EIGHTBITS,
        parity=serial.PARITY_NONE,
        stopbits=serial.STOPBITS_ONE,
        timeout=SERIAL_TIMEOUT,
    )
    _port = port
    ver = get_version()
    log.info("Intensilight connected on %s  firmware %s", port, ver)
    return ver


def disconnect():
    global _ser, _port
    if _ser and _ser.is_open:
        try:
            _ser.close()
        except Exception:
            pass
    _ser = None
    _port = None


def is_connected():
    return _ser is not None and _ser.is_open


def get_version():
    resp = _send("rVEN")
    if len(resp) >= 4 and resp[:4] == "aVEN":
        return resp[4:].strip()
    raise ILError(f"Unexpected version response: {resp}")


# ── Shutter ──────────────────────────────────────────────────────────────────

def shutter_open():
    resp = _send("cSXC1")
    if "SXC" not in resp:
        raise ILError(f"Unexpected shutter response: {resp}")


def shutter_close():
    resp = _send("cSXC2")
    if "SXC" not in resp:
        raise ILError(f"Unexpected shutter response: {resp}")


def shutter_get_state():
    """Return True if open, False if closed."""
    resp = _send("rSXR")
    if len(resp) >= 5 and resp[1:4] == "SXR":
        return resp[4] == "1"
    raise ILError(f"Unexpected shutter state response: {resp}")


# ── ND Filter ────────────────────────────────────────────────────────────────

def nd_set(position):
    """Set ND filter position (1-6).  1=ND1(brightest) … 6=ND32(dimmest)."""
    p = int(position)
    if p < 1 or p > 6:
        raise ILError(f"ND position must be 1-6 (got {p})")
    resp = _send(f"cNDM{p}")
    if "NDM" not in resp:
        raise ILError(f"Unexpected ND set response: {resp}")


def nd_get():
    """Return current ND position (1-6)."""
    resp = _send("rNAR")
    if len(resp) >= 5 and resp[1:4] == "NAR":
        return int(resp[4:].strip())
    raise ILError(f"Unexpected ND get response: {resp}")


def get_state():
    """Return full state dict."""
    state = {"connected": True}
    try:
        state["shutter_open"] = shutter_get_state()
    except Exception as e:
        state["shutter_open"] = None
        state["shutter_error"] = str(e)
    try:
        nd = nd_get()
        state["nd_position"] = nd
        state["nd_label"] = ND_LABELS.get(nd, f"?{nd}")
    except Exception as e:
        state["nd_position"] = None
        state["nd_error"] = str(e)
    return state


# ── Exposed API ──────────────────────────────────────────────────────────────

@expose
def il_connect(port="COM6"):
    try:
        ver = connect(port)
        return {"ok": True, "version": ver, "port": port}
    except Exception as e:
        return {"error": str(e)}


@expose
def il_disconnect():
    try:
        disconnect()
        return {"ok": True}
    except Exception as e:
        return {"error": str(e)}


@expose
def il_is_connected():
    return is_connected()


@expose
def il_shutter_open():
    try:
        shutter_open()
        return {"ok": True}
    except Exception as e:
        return {"error": str(e)}


@expose
def il_shutter_close():
    try:
        shutter_close()
        return {"ok": True}
    except Exception as e:
        return {"error": str(e)}


@expose
def il_shutter_state():
    try:
        return {"ok": True, "open": shutter_get_state()}
    except Exception as e:
        return {"error": str(e)}


@expose
def il_nd_set(position):
    try:
        nd_set(position)
        return {"ok": True, "position": int(position),
                "label": ND_LABELS.get(int(position), "?")}
    except Exception as e:
        return {"error": str(e)}


@expose
def il_nd_get():
    try:
        nd = nd_get()
        return {"ok": True, "position": nd,
                "label": ND_LABELS.get(nd, "?")}
    except Exception as e:
        return {"error": str(e)}


@expose
def il_state():
    if not is_connected():
        return {"connected": False}
    try:
        return get_state()
    except Exception as e:
        return {"error": str(e)}
