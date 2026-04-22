"""
OBSBOT camera control -- VISCA over IP (PTZ) + RTSP frame capture.

VISCA: UDP socket to camera IP:52381 for pan/tilt/zoom/focus commands.
RTSP:  OpenCV VideoCapture for grabbing individual frames on demand.

Waypoints are persisted to Obs_bot/waypoints.json and shared with the
standalone CLI (Obs_bot/obs_test.py).
"""

from modules._api import expose, push_event
import socket
import threading
import time
import json
import os
import io
import re
import logging
import base64
import shutil
import subprocess
from datetime import datetime
import numpy as np
import cv2
from PIL import Image
from paddleocr import PaddleOCR

log = logging.getLogger("obsbot")

def _find_ffmpeg():
    """Locate ffmpeg binary — check PATH first, fall back to known winget location."""
    p = shutil.which("ffmpeg")
    if p:
        return p
    _WINGET = os.path.expandvars(
        r"%LOCALAPPDATA%\Microsoft\WinGet\Packages"
    )
    if os.path.isdir(_WINGET):
        for root, dirs, files in os.walk(_WINGET):
            if "ffmpeg.exe" in files:
                return os.path.join(root, "ffmpeg.exe")
    return None

_FFMPEG = _find_ffmpeg()
if _FFMPEG:
    log.info("FFmpeg found: %s", _FFMPEG)
else:
    log.warning("FFmpeg not found — OBSBOT live stream will not work")

VISCA_PORT = 52381
SEQ_MAX = 2**32 - 1

_BASE_DIR = os.path.normpath(os.path.join(os.path.dirname(__file__), ".."))
import modules.config as _cfg
WAYPOINTS_FILE = os.path.join(_BASE_DIR, "Obs_bot", "waypoints.json")  # legacy
CAPTURES_DIR = os.path.join(_BASE_DIR, "captures", "obsbot")


# ── Low-level VISCA-over-IP transport ────────────────────────────────────────

class VISCAError(Exception):
    pass


class OBSBotVISCA:
    """Thin VISCA-over-IP driver tuned for OBSBOT cameras."""

    def __init__(self, ip: str, port: int = VISCA_PORT, timeout: float = 1.0):
        self.ip = ip
        self.port = port
        self._sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        self._sock.bind(("", 0))
        self._sock.settimeout(timeout)
        self.seq = 0
        self._reset_seq()
        try:
            self._cmd("00 01")
        except VISCAError:
            pass

    def _next_seq(self):
        self.seq = (self.seq + 1) & SEQ_MAX

    def _reset_seq(self):
        msg = bytearray.fromhex("02 00 00 01 00 00 00 01 01")
        self._sock.sendto(msg, (self.ip, self.port))
        try:
            self._sock.recv(32)
        except socket.timeout:
            pass
        self.seq = 1

    def _cmd(self, hex_body: str, query: bool = False, retries: int = 5):
        preamble = b"\x81" + (b"\x09" if query else b"\x01")
        payload = preamble + bytearray.fromhex(hex_body) + b"\xff"
        plen = len(payload).to_bytes(2, "big")
        ptype = b"\x01\x00"

        last_exc = None
        for _ in range(retries):
            self._next_seq()
            seq_bytes = self.seq.to_bytes(4, "big")
            self._sock.sendto(ptype + plen + seq_bytes + payload,
                              (self.ip, self.port))
            try:
                resp = self._recv()
            except VISCAError as e:
                last_exc = e
                continue
            if resp is not None:
                return resp[1:-1]
            if not query:
                return None
        if last_exc:
            raise last_exc
        raise VISCAError("No response after retries")

    def _recv(self):
        while True:
            try:
                data = self._sock.recv(64)
            except socket.timeout:
                return None
            resp_seq = int.from_bytes(data[4:8], "big")
            if resp_seq < self.seq:
                continue
            body = data[8:]
            if len(body) > 2:
                status = body[1] >> 4
                if status not in (4, 5):
                    raise VISCAError(f"Error response: {body.hex()}")
                return body
            return None

    def close(self):
        self._sock.close()

    @staticmethod
    def _decode_pos(raw: bytes, signed: bool = True) -> int:
        nibbles = "".join(f"{b & 0x0F:x}" for b in raw)
        val = int(nibbles, 16)
        if signed and val >= 0x8000:
            val -= 0x10000
        return val

    @staticmethod
    def _encode_pos(val: int) -> str:
        raw = val.to_bytes(2, "big", signed=True)
        return " ".join(f"0{c}" for c in raw.hex())

    # ── pan / tilt ───────────────────────────────────────────────────────

    def pantilt(self, pan_speed: int, tilt_speed: int,
                pan_pos=None, tilt_pos=None, relative: bool = False):
        ps = max(0, min(abs(pan_speed), 24))
        ts = max(0, min(abs(tilt_speed), 24))

        if pan_pos is not None and tilt_pos is not None:
            mode = "03" if relative else "02"
            self._cmd(f"06 {mode} {ps:02x} {ts:02x} "
                      f"{self._encode_pos(pan_pos)} {self._encode_pos(tilt_pos)}")
        else:
            def _dir(s):
                if s < 0:  return "01"
                if s > 0:  return "02"
                return "03"
            self._cmd(f"06 01 {ps:02x} {ts:02x} "
                      f"{_dir(pan_speed)} {_dir(tilt_speed)}")

    def pantilt_stop(self):
        self._cmd("06 01 00 00 03 03")

    def pantilt_home(self):
        self._cmd("06 04")

    def get_pantilt_position(self):
        try:
            resp = self._cmd("06 12", query=True)
            if resp is None or len(resp) < 9:
                return None, None
            return self._decode_pos(resp[1:5]), self._decode_pos(resp[5:9])
        except Exception:
            return None, None

    # ── zoom ─────────────────────────────────────────────────────────────

    def zoom(self, speed: int):
        s = max(-7, min(speed, 7))
        if s == 0:   d = "0"
        elif s > 0:  d = "2"
        else:        d = "3"
        self._cmd(f"04 07 {d}{abs(s):x}")

    def zoom_stop(self):
        self.zoom(0)

    def zoom_to(self, fraction: float):
        p = max(0, min(int(fraction * 16384), 16384))
        h = f"{p:04x}"
        self._cmd("04 47 " + " ".join(f"0{c}" for c in h))

    def get_zoom_position(self):
        try:
            resp = self._cmd("04 47", query=True)
            if resp is None or len(resp) < 4:
                return None
            return self._decode_pos(resp[1:5], signed=False)
        except Exception:
            return None

    # ── focus ────────────────────────────────────────────────────────────

    def set_focus_mode(self, mode: str):
        modes = {"auto": "38 02", "manual": "38 03", "toggle": "38 10",
                 "one_push": "18 01", "infinity": "18 02"}
        if mode not in modes:
            raise VISCAError(f"Unknown focus mode: {mode}")
        self._cmd("04 " + modes[mode])

    def get_focus_mode(self):
        try:
            resp = self._cmd("04 38", query=True)
            if resp is None:
                return "unknown"
            return {2: "auto", 3: "manual"}.get(resp[-1], "unknown")
        except Exception:
            return "unknown"

    # ── power ─────────────────────────────────────────────────────────────

    def power_on(self):
        self._cmd("04 00 02")

    def power_off(self):
        self._cmd("04 00 03")

    def get_power(self) -> bool:
        try:
            resp = self._cmd("04 00", query=True)
            if resp is None or len(resp) < 2:
                return True  # assume on if we got a response at all
            return resp[-1] == 2
        except Exception:
            return False

    # ── presets ───────────────────────────────────────────────────────────

    def save_preset(self, num: int):
        self._cmd(f"04 3F 01 {num:02x}")

    def recall_preset(self, num: int):
        self._cmd(f"04 3F 02 {num:02x}")

    # ── nudge ────────────────────────────────────────────────────────────

    NUDGE_SCALE = 400
    NUDGE_MIN = 10

    def nudge(self, pan_dir: int, tilt_dir: int, level: float = 1.0):
        step = max(self.NUDGE_MIN, int(self.NUDGE_SCALE * level))
        pan, tilt = self.get_pantilt_position()
        if pan is None or tilt is None:
            raise VISCAError("Cannot read current position")
        target_pan = pan + pan_dir * step
        target_tilt = tilt + tilt_dir * step
        speed = max(6, min(int(6 + level * 4), 18))
        self.pantilt(speed, speed, pan_pos=target_pan, tilt_pos=target_tilt)

    # ── absolute move ────────────────────────────────────────────────────

    def goto(self, pan: int, tilt: int, zoom: int,
             pan_speed: int = 12, tilt_speed: int = 12):
        self.pantilt(pan_speed, tilt_speed, pan_pos=pan, tilt_pos=tilt)
        self.zoom_to(zoom / 16384.0)

    # ── status ───────────────────────────────────────────────────────────

    def status(self):
        pan, tilt = self.get_pantilt_position()
        zoom = self.get_zoom_position()
        focus = self.get_focus_mode()
        return {"pan": pan, "tilt": tilt, "zoom": zoom, "focus_mode": focus}


# ── Module state ─────────────────────────────────────────────────────────────

_cam = None          # type: OBSBotVISCA | None
_rtsp_url = None     # type: str | None
_waypoints = []      # type: list[dict]
_seq_thread = None   # type: threading.Thread | None
_seq_stop = threading.Event()

# ── Live stream (FFmpeg subprocess) ──────────────────────────────────────────

_live_thread = None
_live_stop = threading.Event()
_live_jpeg = None
_frame_event = threading.Event()
_LIVE_FPS = 10
_LIVE_SCALE_W = 960

_SOI = b'\xff\xd8'
_EOI = b'\xff\xd9'


def _live_loop():
    """Background thread: run FFmpeg to decode RTSP -> MJPEG, read JPEGs from
    stdout pipe.  FFmpeg handles H.264 decoding natively — far more reliable
    than OpenCV's built-in decoder."""
    global _live_jpeg

    if not _FFMPEG:
        log.error("FFmpeg not found")
        return

    cmd = [
        _FFMPEG, "-hide_banner", "-loglevel", "error",
        "-rtsp_transport", "tcp",
        "-i", _rtsp_url,
        "-vf", f"scale={_LIVE_SCALE_W}:-2",
        "-r", str(_LIVE_FPS),
        "-q:v", "4",
        "-f", "mjpeg",
        "-"
    ]
    log.info("FFmpeg live: %s", " ".join(cmd))
    proc = None
    try:
        proc = subprocess.Popen(
            cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0)
        )
        buf = b""
        while not _live_stop.is_set():
            chunk = proc.stdout.read(16384)
            if not chunk:
                rc = proc.poll()
                if rc is not None:
                    stderr_out = proc.stderr.read().decode(errors="replace")
                    log.error("FFmpeg exited %d: %s", rc, stderr_out[:500])
                break
            buf += chunk
            while True:
                soi = buf.find(_SOI)
                if soi < 0:
                    buf = b""
                    break
                eoi = buf.find(_EOI, soi + 2)
                if eoi < 0:
                    buf = buf[soi:]
                    break
                _live_jpeg = buf[soi:eoi + 2]
                _frame_event.set()
                buf = buf[eoi + 2:]
    except Exception as e:
        log.error("FFmpeg live error: %s", e)
    finally:
        if proc and proc.poll() is None:
            proc.terminate()
            try:
                proc.wait(timeout=3)
            except subprocess.TimeoutExpired:
                proc.kill()
        _live_jpeg = None
        _frame_event.set()
        log.info("FFmpeg live stopped")


def live_is_active():
    return _live_thread is not None and _live_thread.is_alive()


def get_live_jpeg():
    return _live_jpeg


def snap_jpeg() -> bytes:
    """Grab one clean frame via FFmpeg (single-shot)."""
    if not _FFMPEG:
        raise RuntimeError("FFmpeg not found")
    if not _rtsp_url:
        raise RuntimeError("No RTSP URL configured")

    cmd = [
        _FFMPEG, "-hide_banner", "-loglevel", "error",
        "-rtsp_transport", "tcp",
        "-i", _rtsp_url,
        "-frames:v", "1",
        "-q:v", "2",
        "-f", "mjpeg",
        "-"
    ]
    result = subprocess.run(
        cmd, capture_output=True, timeout=10,
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0)
    )
    if result.returncode != 0:
        raise RuntimeError(f"FFmpeg snap failed: {result.stderr.decode(errors='replace')[:300]}")
    if not result.stdout:
        raise RuntimeError("FFmpeg returned empty output")
    return result.stdout


# ── Waypoint persistence ────────────────────────────────────────────────────

def _load_waypoints():
    global _waypoints
    data = _cfg.load("obsbot_waypoints")
    if data is not None:
        _waypoints = data
    elif os.path.exists(WAYPOINTS_FILE):
        with open(WAYPOINTS_FILE, "r") as f:
            _waypoints = json.load(f)
        _cfg.save("obsbot_waypoints", _waypoints)
    else:
        _waypoints = []


def _save_waypoints():
    _cfg.save("obsbot_waypoints", _waypoints)


# ── RTSP frame capture ──────────────────────────────────────────────────────

def _grab_frame(rtsp_url: str):
    """Grab a single clean frame via FFmpeg subprocess."""
    if not _FFMPEG:
        raise RuntimeError("FFmpeg not found")

    _ffprobe = _FFMPEG.replace("ffmpeg", "ffprobe")
    probe_cmd = [
        _ffprobe, "-v", "error", "-rtsp_transport", "tcp",
        "-select_streams", "v:0",
        "-show_entries", "stream=width,height",
        "-of", "csv=p=0", rtsp_url
    ]
    probe = subprocess.run(probe_cmd, capture_output=True, timeout=10,
                           creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    w, h = [int(x) for x in probe.stdout.decode().strip().split(",")]

    cmd = [
        _FFMPEG, "-hide_banner", "-loglevel", "error",
        "-rtsp_transport", "tcp",
        "-i", rtsp_url,
        "-frames:v", "1",
        "-f", "image2pipe",
        "-pix_fmt", "bgr24",
        "-vcodec", "rawvideo",
        "-"
    ]
    result = subprocess.run(
        cmd, capture_output=True, timeout=10,
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0)
    )
    if result.returncode != 0:
        raise RuntimeError(f"FFmpeg grab failed: {result.stderr.decode(errors='replace')[:300]}")
    if not result.stdout:
        raise RuntimeError("FFmpeg returned no frame data")

    raw = result.stdout
    frame = np.frombuffer(raw, dtype=np.uint8).reshape((h, w, 3))
    return frame


def _save_frame(frame, name: str, meta: dict):
    """Save frame as JPEG + JSON sidecar to captures/obsbot/."""
    os.makedirs(CAPTURES_DIR, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe = "".join(c if c.isalnum() or c in "-_ " else "_" for c in name)
    basename = f"{ts}_{safe}"
    jpg_path = os.path.join(CAPTURES_DIR, f"{basename}.jpg")
    meta_path = os.path.join(CAPTURES_DIR, f"{basename}.meta.json")
    cv2.imwrite(jpg_path, frame)
    with open(meta_path, "w") as f:
        json.dump(meta, f, indent=2)
    return jpg_path


# ── Sequence runner (background thread) ─────────────────────────────────────

def _sequence_worker(dwell_s: float):
    """Iterate all waypoints: goto -> settle -> capture."""
    global _seq_thread
    filenames = []
    total = len(_waypoints)
    try:
        for i, wp in enumerate(_waypoints):
            if _seq_stop.is_set():
                break
            name = wp["name"]
            push_event("onObsbotProgress", name, i, total, "moving")
            log.info("Sequence: moving to #%d '%s'", i, name)
            _cam.goto(wp["pan"], wp["tilt"], wp["zoom"])

            # settle
            remaining = dwell_s
            while remaining > 0 and not _seq_stop.is_set():
                time.sleep(min(remaining, 0.5))
                remaining -= 0.5

            if _seq_stop.is_set():
                break

            push_event("onObsbotProgress", name, i, total, "capturing")
            log.info("Sequence: capturing at #%d '%s'", i, name)
            try:
                frame = _grab_frame(_rtsp_url)
                meta = {"waypoint": name, "index": i,
                        "pan": wp["pan"], "tilt": wp["tilt"],
                        "zoom": wp["zoom"],
                        "timestamp": datetime.now().isoformat()}
                path = _save_frame(frame, name, meta)
                filenames.append(path)
            except Exception as e:
                log.error("Capture failed at waypoint '%s': %s", name, e)
                push_event("onObsbotProgress", name, i, total,
                           f"capture_error: {e}")

        push_event("onObsbotSequenceDone", filenames)
        log.info("Sequence complete: %d captures", len(filenames))
    except Exception as e:
        log.error("Sequence failed: %s", e)
        push_event("onObsbotSequenceDone", [], str(e))
    finally:
        _seq_thread = None


# ── Exposed API ──────────────────────────────────────────────────────────────

@expose
def obsbot_connect(ip="192.168.1.31", rtsp_url="rtsp://192.168.1.31/stream2"):
    """Connect VISCA + store RTSP URL."""
    global _cam, _rtsp_url
    try:
        disconnect()
        _cam = OBSBotVISCA(ip)
        _rtsp_url = rtsp_url if rtsp_url else None
        _load_waypoints()
        s = _cam.status()
        log.info("OBSBOT connected at %s  (RTSP: %s)", ip,
                 _rtsp_url or "not set")
        return {"ok": True, "ip": ip, "rtsp_url": _rtsp_url, **s}
    except Exception as e:
        return {"error": str(e)}


@expose
def obsbot_disconnect():
    try:
        disconnect()
        return {"ok": True}
    except Exception as e:
        return {"error": str(e)}


@expose
def obsbot_is_connected():
    return _cam is not None


@expose
def obsbot_status():
    if _cam is None:
        return {"connected": False}
    try:
        s = _cam.status()
        s["connected"] = True
        s["rtsp_url"] = _rtsp_url
        s["waypoint_count"] = len(_waypoints)
        s["sequence_running"] = _seq_thread is not None
        return s
    except Exception as e:
        return {"error": str(e)}


@expose
def obsbot_power_on():
    if _cam is None:
        return {"error": "Not connected"}
    try:
        _cam.power_on()
        return {"ok": True}
    except Exception as e:
        return {"error": str(e)}


@expose
def obsbot_power_off():
    if _cam is None:
        return {"error": "Not connected"}
    try:
        _cam.power_off()
        return {"ok": True}
    except Exception as e:
        return {"error": str(e)}


@expose
def obsbot_home():
    if _cam is None:
        return {"error": "Not connected"}
    try:
        _cam.pantilt_home()
        return {"ok": True}
    except Exception as e:
        return {"error": str(e)}


@expose
def obsbot_live_start():
    """Start FFmpeg-based RTSP live stream."""
    global _live_thread
    if not _rtsp_url:
        return {"error": "No RTSP URL configured -- pass rtsp_url to obsbot_connect"}
    if live_is_active():
        return {"ok": True, "already_running": True}
    _live_stop.clear()
    _live_jpeg = None
    _live_thread = threading.Thread(target=_live_loop, daemon=True,
                                    name="obsbot-live")
    _live_thread.start()
    return {"ok": True}


@expose
def obsbot_live_stop():
    """Stop the live stream."""
    global _live_thread
    if not live_is_active():
        return {"ok": True, "was_running": False}
    _live_stop.set()
    _live_thread.join(timeout=5)
    _live_thread = None
    return {"ok": True, "was_running": True}


@expose
def obsbot_snap_preview():
    """Grab a single clean frame from RTSP, return as base64 JPEG."""
    if _cam is None:
        return {"error": "Not connected"}
    if not _rtsp_url:
        return {"error": "No RTSP URL configured"}
    try:
        jpeg = snap_jpeg()
        b64 = base64.b64encode(jpeg).decode("ascii")
        return {"ok": True, "image": b64}
    except Exception as e:
        return {"error": str(e)}


@expose
def obsbot_goto(index):
    """Move to a waypoint by index."""
    if _cam is None:
        return {"error": "Not connected"}
    idx = int(index)
    if idx < 0 or idx >= len(_waypoints):
        return {"error": f"Index {idx} out of range (0-{len(_waypoints)-1})"}
    wp = _waypoints[idx]
    try:
        _cam.goto(wp["pan"], wp["tilt"], wp["zoom"])
        return {"ok": True, "waypoint": wp}
    except Exception as e:
        return {"error": str(e)}


@expose
def obsbot_mark(name=""):
    """Read current position and save as a waypoint."""
    if _cam is None:
        return {"error": "Not connected"}
    pan, tilt = _cam.get_pantilt_position()
    zoom = _cam.get_zoom_position()
    if pan is None or zoom is None:
        return {"error": "Could not read current position"}
    wp_name = name if name else f"wp{len(_waypoints)}"
    wp = {"name": wp_name, "pan": pan, "tilt": tilt, "zoom": zoom}
    _waypoints.append(wp)
    _save_waypoints()
    log.info("Marked waypoint #%d '%s'  pan=%d tilt=%d zoom=%d",
             len(_waypoints) - 1, wp_name, pan, tilt, zoom)
    return {"ok": True, "index": len(_waypoints) - 1, "waypoint": wp}


@expose
def obsbot_list_waypoints():
    return {"ok": True, "waypoints": _waypoints}


@expose
def obsbot_delete_waypoint(index):
    idx = int(index)
    if idx < 0 or idx >= len(_waypoints):
        return {"error": f"Index {idx} out of range (0-{len(_waypoints)-1})"}
    removed = _waypoints.pop(idx)
    _save_waypoints()
    return {"ok": True, "removed": removed}


@expose
def obsbot_capture(filename=""):
    """Grab a single frame from RTSP and save to captures/obsbot/."""
    if _cam is None:
        return {"error": "Not connected"}
    if not _rtsp_url:
        return {"error": "No RTSP URL configured -- pass rtsp_url to obsbot_connect"}
    try:
        frame = _grab_frame(_rtsp_url)
        name = filename if filename else "manual"
        pan, tilt = _cam.get_pantilt_position()
        zoom = _cam.get_zoom_position()
        meta = {"name": name, "pan": pan, "tilt": tilt, "zoom": zoom,
                "timestamp": datetime.now().isoformat()}
        path = _save_frame(frame, name, meta)
        return {"ok": True, "path": path}
    except Exception as e:
        return {"error": str(e)}


@expose
def obsbot_run_sequence(dwell_s=3.0):
    """Iterate all waypoints: goto -> settle -> capture at each.
    Runs in background thread; push events signal progress."""
    global _seq_thread
    if _cam is None:
        return {"error": "Not connected"}
    if not _rtsp_url:
        return {"error": "No RTSP URL configured -- pass rtsp_url to obsbot_connect"}
    if not _waypoints:
        return {"error": "No waypoints saved"}
    if _seq_thread is not None:
        return {"error": "Sequence already running"}
    dwell = float(dwell_s)
    _seq_stop.clear()
    _seq_thread = threading.Thread(target=_sequence_worker, args=(dwell,),
                                   daemon=True, name="obsbot-seq")
    _seq_thread.start()
    log.info("Sequence started: %d waypoints, %.1fs dwell",
             len(_waypoints), dwell)
    return {"ok": True, "waypoints": len(_waypoints), "dwell_s": dwell}


@expose
def obsbot_stop_sequence():
    """Stop a running sequence."""
    if _seq_thread is None:
        return {"ok": True, "was_running": False}
    _seq_stop.set()
    _seq_thread.join(timeout=10)
    return {"ok": True, "was_running": True}


# ── Incubator monitor ────────────────────────────────────────────────────────

MONITOR_DIR = os.path.join(_BASE_DIR, "captures", "incubator")
MONITOR_CSV = os.path.join(MONITOR_DIR, "readings.csv")

_mon_thread = None
_mon_stop = threading.Event()
_mon_latest = None   # dict with parsed readings

_CROP_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                          "..", "Obs_bot", "monitor_crops.json")  # legacy
_crops = {"temp": None, "co2": None}  # each is [x1,y1,x2,y2] or None
_paddle_ocr = None  # lazy-init PaddleOCR instance


def _load_crops():
    global _crops
    data = _cfg.load("obsbot_crops")
    if data is not None:
        _crops = data
        return
    try:
        if os.path.exists(_CROP_FILE):
            with open(_CROP_FILE, "r") as f:
                _crops = json.load(f)
            _cfg.save("obsbot_crops", _crops)
    except Exception:
        pass

def _save_crops():
    _cfg.save("obsbot_crops", _crops)

_load_crops()


def _get_paddle_ocr():
    global _paddle_ocr
    if _paddle_ocr is None:
        _paddle_ocr = PaddleOCR(
            use_angle_cls=False, lang='en', use_gpu=False, show_log=False)
    return _paddle_ocr


def _ocr_led_region(jpeg_bytes: bytes, crop: list) -> str:
    """Crop a region from a JPEG, preprocess for 7-segment LED, OCR via PaddleOCR."""
    img = Image.open(io.BytesIO(jpeg_bytes))
    w, h = img.size

    if crop:
        x1 = int(crop[0] * w)
        y1 = int(crop[1] * h)
        x2 = int(crop[2] * w)
        y2 = int(crop[3] * h)
        img = img.crop((x1, y1, x2, y2))

    arr = np.array(img).astype(float)
    r, g, b = arr[:, :, 0], arr[:, :, 1], arr[:, :, 2]

    mask = ((r > 140) & (r > g * 1.4) & (r > b * 1.4)).astype(np.uint8) * 255

    contours, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    rects = [(cv2.boundingRect(c), cv2.contourArea(c))
             for c in contours if cv2.contourArea(c) >= 5]
    if not rects:
        return ''

    total_area = sum(a for _, a in rects)
    cx = sum((bx + bw / 2) * a for (bx, _, bw, _), a in rects) / total_area
    cy = sum((by + bh / 2) * a for (_, by, _, bh), a in rects) / total_area
    max_dist = max(img.size) * 0.4
    cluster = [((bx, by, bw, bh), a) for (bx, by, bw, bh), a in rects
               if ((bx + bw / 2 - cx) ** 2 + (by + bh / 2 - cy) ** 2) ** 0.5 < max_dist]
    if not cluster:
        return ''

    ax1 = min(bx for (bx, _, _, _), _ in cluster)
    ay1 = min(by for (_, by, _, _), _ in cluster)
    ax2 = max(bx + bw for (bx, _, bw, _), _ in cluster)
    ay2 = max(by + bh for (_, by, _, bh), _ in cluster)
    p = 5
    mh, mw = mask.shape[:2]
    ax1, ay1 = max(0, ax1 - p), max(0, ay1 - p)
    ax2, ay2 = min(mw, ax2 + p), min(mh, ay2 + p)

    digit_mask = mask[ay1:ay2, ax1:ax2]

    kern = cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (3, 3))
    digit_mask = cv2.morphologyEx(digit_mask, cv2.MORPH_CLOSE, kern)

    # Enlarge decimal points so OCR can see them
    dm_h, dm_w = digit_mask.shape[:2]
    dot_contours, _ = cv2.findContours(
        digit_mask.copy(), cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    dot_rects = [(cv2.boundingRect(c), cv2.contourArea(c))
                 for c in dot_contours if cv2.contourArea(c) >= 2]
    if dot_rects:
        max_blob_h = max(bh for (_, _, _, bh), _ in dot_rects)
        digit_min_area = max_blob_h * max_blob_h * 0.1
        for (bx, by, bw, bh), a in dot_rects:
            too_big = a > digit_min_area
            too_tall = bh > max_blob_h * 0.25
            too_wide = bw > max_blob_h * 0.25
            not_at_baseline = by + bh < dm_h * 0.6
            if too_big or too_tall or too_wide or not_at_baseline:
                continue
            target_size = max(2, int(max_blob_h * 0.07))
            center = (bx + bw // 2, by + bh // 2)
            cv2.circle(digit_mask, center, target_size, 255, -1)

    scaled = cv2.resize(digit_mask, None, fx=3, fy=3,
                        interpolation=cv2.INTER_NEAREST)
    rgb = cv2.cvtColor(scaled, cv2.COLOR_GRAY2BGR)
    bp = 30
    bordered = np.zeros((rgb.shape[0] + 2 * bp, rgb.shape[1] + 2 * bp, 3),
                        dtype=np.uint8)
    bordered[bp:-bp, bp:-bp] = rgb

    paddle = _get_paddle_ocr()
    result = paddle.ocr(bordered, det=False, cls=False)
    text = ''
    if result:
        for page in result:
            if page:
                for item in page:
                    text = item[0]

    cleaned = re.sub(r'[^0-9.]', '', text)

    # Fallback: insert decimal from contour layout if OCR missed it
    if '.' not in cleaned and cleaned:
        dm_h, dm_w = digit_mask.shape[:2]
        dc, _ = cv2.findContours(digit_mask, cv2.RETR_EXTERNAL,
                                 cv2.CHAIN_APPROX_SIMPLE)
        dc_rects = [(cv2.boundingRect(c), cv2.contourArea(c))
                    for c in dc if cv2.contourArea(c) >= 3]
        max_h = max(bh for (_, _, _, bh), _ in dc_rects) if dc_rects else 1

        dots, segs = [], []
        for (bx, by, bw, bh), a in dc_rects:
            if bw < max_h * 0.3 and bh < max_h * 0.3 and by + bh > dm_h * 0.6:
                dots.append((bx, by, bw, bh))
            elif a > max(15, max_h * 0.5):
                segs.append((bx, by, bw, bh))

        groups = []
        for s in sorted(segs, key=lambda r: r[0]):
            placed = False
            for g in groups:
                gx2 = max(sx + sw for sx, _, sw, _ in g)
                gx1 = min(sx for sx, _, _, _ in g)
                if s[0] < gx2 + 10 and s[0] + s[2] > gx1 - 10:
                    g.append(s)
                    placed = True
                    break
            if not placed:
                groups.append([s])
        groups.sort(key=lambda g: min(sx for sx, _, _, _ in g))

        for dx, dy, dw, dh in dots:
            dot_cx = dx + dw // 2
            insert_after = 0
            for i, g in enumerate(groups):
                gx2 = max(sx + sw for sx, _, sw, _ in g)
                if dot_cx > gx2:
                    insert_after = i + 1
            if 0 < insert_after <= len(cleaned):
                cleaned = cleaned[:insert_after] + '.' + cleaned[insert_after:]
                break

    return cleaned


def _read_display(jpeg_bytes: bytes) -> dict:
    """OCR both display regions and return parsed temperature and CO2."""
    temp_val = None
    co2_val = None
    raw_parts = []

    for label, key in [("temp", "temperature"), ("co2", "co2")]:
        crop = _crops.get(label)
        try:
            text = _ocr_led_region(jpeg_bytes, crop)
            raw_parts.append(f"{label}={text!r}")
            digits = re.sub(r"[^0-9]", "", text)
            if digits:
                if key == "temperature" and len(digits) >= 3:
                    # Temp is always ##.# (3 digits, decimal after 2nd)
                    val = float(digits[:2] + "." + digits[2])
                    temp_val = val
                elif key == "co2" and len(digits) >= 2:
                    # CO2 is always #.# (2 digits, decimal after 1st)
                    val = float(digits[0] + "." + digits[1])
                    co2_val = val
                else:
                    val = float(digits)
                    if key == "temperature":
                        temp_val = val
                    else:
                        co2_val = val
        except Exception as e:
            raw_parts.append(f"{label}=ERR:{e}")
            log.error("OCR %s failed: %s", label, e)

    return {
        "temperature": temp_val,
        "co2": co2_val,
        "raw": "  |  ".join(raw_parts),
    }


def _do_reading(save_dir: str = None) -> dict:
    """Take 4 snapshots, OCR each, use the mode for temp and CO2, log to CSV."""
    from collections import Counter

    temps = []
    co2s = []
    raw_parts = []
    last_jpeg = None

    for i in range(4):
        try:
            jpeg = snap_jpeg()
            last_jpeg = jpeg
            parsed = _read_display(jpeg)
            if parsed["temperature"] is not None:
                temps.append(parsed["temperature"])
            if parsed["co2"] is not None:
                co2s.append(parsed["co2"])
            raw_parts.append(parsed["raw"])
        except Exception as e:
            log.error("Reading attempt %d failed: %s", i + 1, e)
            raw_parts.append(f"attempt{i+1}=ERR:{e}")
        if i < 3:
            time.sleep(0.5)

    temp_val = Counter(temps).most_common(1)[0][0] if temps else None
    co2_val = Counter(co2s).most_common(1)[0][0] if co2s else None

    ts = datetime.now()
    ts_str = ts.strftime("%Y-%m-%d %H:%M:%S")
    temp_s = "" if temp_val is None else str(temp_val)
    co2_s = "" if co2_val is None else str(co2_val)
    raw_all = " | ".join(raw_parts)
    raw_esc = raw_all.replace('"', '""')

    out_dir = save_dir or MONITOR_DIR
    os.makedirs(out_dir, exist_ok=True)

    csv_path = os.path.join(out_dir, "incubator_readings.csv")
    write_header = not os.path.exists(csv_path)
    with open(csv_path, "a", newline="", encoding="utf-8") as csv_f:
        if write_header:
            csv_f.write("timestamp,temperature,co2,raw_ocr\n")
        csv_f.write(f'{ts_str},{temp_s},{co2_s},"{raw_esc}"\n')

    b64 = base64.b64encode(last_jpeg).decode("ascii") if last_jpeg else ""
    return {
        "timestamp": ts_str,
        "temperature": temp_val,
        "co2": co2_val,
        "raw_text": raw_all,
        "image": b64,
    }


def _monitor_loop(interval_s: float, waypoint_name: str):
    """Background thread: periodically snap frame, OCR, log to CSV."""
    global _mon_latest

    try:
        if waypoint_name:
            for i, wp in enumerate(_waypoints):
                if wp["name"] == waypoint_name:
                    _cam.pantilt(10, 10, pan_pos=wp["pan"], tilt_pos=wp["tilt"])
                    if "zoom" in wp:
                        _cam.zoom(wp["zoom"])
                    time.sleep(3)
                    break

        while not _mon_stop.is_set():
            try:
                reading = _do_reading()
            except Exception as e:
                log.error("Monitor reading failed: %s", e)
                _mon_stop.wait(timeout=60)
                continue
            _mon_latest = reading

            push_event("onIncubatorReading", reading["timestamp"],
                        reading.get("temperature"), reading.get("co2"),
                        reading["raw_text"])
            log.info("Incubator: %s  temp=%s  co2=%s",
                     reading["timestamp"], reading["temperature"],
                     reading["co2"])

            _mon_stop.wait(timeout=interval_s)
    except Exception as e:
        log.error("Monitor loop error: %s", e)
    finally:
        _mon_latest = None
        log.info("Incubator monitor stopped")


def monitor_is_active():
    return _mon_thread is not None and _mon_thread.is_alive()


@expose
def obsbot_monitor_start(interval_min=60, waypoint=""):
    """Start incubator monitoring."""
    global _mon_thread
    if _cam is None:
        return {"error": "Not connected"}
    if monitor_is_active():
        return {"ok": True, "already_running": True}
    if not _crops.get("temp") and not _crops.get("co2"):
        return {"error": "Set crop regions for Temp and CO2 first (use Set Crop buttons)"}
    interval_s = float(interval_min) * 60
    _mon_stop.clear()
    _mon_thread = threading.Thread(
        target=_monitor_loop, args=(interval_s, waypoint),
        daemon=True, name="obsbot-monitor"
    )
    _mon_thread.start()
    log.info("Incubator monitor started: every %s min, waypoint=%r",
             interval_min, waypoint)
    return {"ok": True, "interval_min": float(interval_min),
            "csv": MONITOR_CSV}


@expose
def obsbot_monitor_stop():
    """Stop incubator monitoring."""
    global _mon_thread
    if not monitor_is_active():
        return {"ok": True, "was_running": False}
    _mon_stop.set()
    _mon_thread.join(timeout=10)
    _mon_thread = None
    return {"ok": True, "was_running": True}


@expose
def obsbot_monitor_latest():
    """Return the latest incubator reading."""
    if _mon_latest is None:
        return {"ok": True, "reading": None}
    return {"ok": True, "reading": _mon_latest}


@expose
def obsbot_monitor_snap_now():
    """Take an immediate incubator reading (manual trigger)."""
    if _cam is None:
        return {"error": "Not connected"}
    try:
        reading = _do_reading()
    except Exception as e:
        return {"error": f"Reading failed: {e}"}
    return {"ok": True, **reading}


@expose
def obsbot_incubator_reading(save_dir=""):
    """Take a reading and save to a specific directory (for timeline use).

    If save_dir is empty, falls back to the default incubator monitor folder.
    The waypoint from the crop config is used for camera positioning.
    """
    if _cam is None:
        return {"error": "OBSBOT not connected"}
    try:
        reading = _do_reading(save_dir=save_dir or None)
    except Exception as e:
        return {"error": f"Reading failed: {e}"}
    push_event("onIncubatorReading", reading["timestamp"],
               reading.get("temperature"), reading.get("co2"),
               reading["raw_text"])
    return {"ok": True, **reading}


@expose
def obsbot_monitor_set_crop(which, x1, y1, x2, y2):
    """Set crop region for 'temp' or 'co2'. Coordinates are 0..1 fractions."""
    if which not in ("temp", "co2"):
        return {"error": "which must be 'temp' or 'co2'"}
    _crops[which] = [float(x1), float(y1), float(x2), float(y2)]
    _save_crops()
    log.info("Crop set: %s = %s", which, _crops[which])
    return {"ok": True, "crops": _crops}


@expose
def obsbot_monitor_get_crops():
    """Return current crop regions."""
    return {"ok": True, "crops": _crops}


# ── Internal helpers ─────────────────────────────────────────────────────────

def disconnect():
    global _cam, _rtsp_url, _live_thread, _mon_thread
    if monitor_is_active():
        _mon_stop.set()
        _mon_thread.join(timeout=10)
        _mon_thread = None
    if live_is_active():
        _live_stop.set()
        _live_thread.join(timeout=5)
        _live_thread = None
    if _cam is not None:
        try:
            _cam.close()
        except Exception:
            pass
        _cam = None
    _rtsp_url = None


# Load waypoints on import so they're available immediately
_load_waypoints()
