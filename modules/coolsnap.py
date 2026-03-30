"""
CoolSNAP EZ camera module — direct control via pvcam64.dll (ctypes).

Uses a custom ctypes wrapper (pvcam_raw.py) that calls the PVCAM C library
directly, bypassing PyVCAM entirely. This avoids DLL-version mismatches
between the Python package and the installed PVCAM driver.

Provides:
  - Live view (WebSocket binary JPEG stream)
  - Snap (single frame capture + save)
  - Video recording (N frames to .npy)
  - Time-lapse (frames at an interval, saved as .npy stack)
  - Exposure time control (ms)
  - Binning (1x1, 2x2, 4x4, 8x8)
"""

from modules._api import expose, push_event
from modules import pvcam_raw as pvc
import threading
import numpy as np
import base64
import io
import os
import time
from datetime import datetime

_HAS_CV2 = True
try:
    import cv2
except ImportError:
    _HAS_CV2 = False

_HAS_PIL = True
try:
    from PIL import Image as _PILImage
except ImportError:
    _HAS_PIL = False

# ── Module state ─────────────────────────────────────────────────────────────

_lock = threading.Lock()
_hcam = None              # camera handle (int16) from pvcam_raw.cam_open()
_pvc_initialized = False
_sensor_w = 0             # full sensor width
_sensor_h = 0             # full sensor height
_bit_depth = 0
_cam_name_str = ""

_exposure_ms = 20
_binning = (1, 1)

_live_thread = None
_live_stop = threading.Event()
_live_fps = 0.0

_capture_thread = None
_capture_stop = threading.Event()

_save_dir = os.path.join(os.path.dirname(__file__), "..", "captures")

_frame_lock = threading.Lock()
_latest_jpeg = None          # raw JPEG bytes for WebSocket stream
_frame_event = threading.Event()

# Circular buffer for continuous acquisition — kept alive while live view runs
_circ_buf = None
_circ_buf_size = 0

# Display range / auto-brightness
# Modes: "auto" = continuous EMA-smoothed percentile stretch (default)
#         "locked" = frozen min/max from last auto-adjust
_disp_mode = "auto"
_disp_vmin = 0.0
_disp_vmax = 65535.0
_disp_gamma = 1.0         # gamma correction: >1 darkens midtones, <1 brightens
_DISP_EMA_ALPHA = 0.25    # smoothing factor: lower = more stable, higher = more responsive
_DISP_LO_PCT = 1.0        # low percentile for auto-stretch
_DISP_HI_PCT = 99.0       # high percentile for auto-stretch

_gamma_lut = None          # cached uint8 LUT for current gamma value
_gamma_lut_val = None      # gamma value the LUT was built for


class CamError(Exception):
    """Raised when a camera operation fails."""


# ── Helpers ──────────────────────────────────────────────────────────────────

_LIVE_PREVIEW_DIM = 800

def _compute_percentiles(frame, lo_pct=None, hi_pct=None):
    """Fast percentile computation on a subsampled ravel of the frame."""
    if lo_pct is None:
        lo_pct = _DISP_LO_PCT
    if hi_pct is None:
        hi_pct = _DISP_HI_PCT
    flat = frame.ravel()
    sample = flat[:: max(1, flat.size // 10000)]
    return float(np.percentile(sample, lo_pct)), float(np.percentile(sample, hi_pct))


def _get_gamma_lut(gamma):
    """Return a uint8[256] lookup table for the given gamma value."""
    global _gamma_lut, _gamma_lut_val
    if _gamma_lut_val == gamma and _gamma_lut is not None:
        return _gamma_lut
    lut = np.arange(256, dtype=np.float32) / 255.0
    lut = np.power(lut, gamma)
    _gamma_lut = (lut * 255).astype(np.uint8)
    _gamma_lut_val = gamma
    return _gamma_lut


def _normalize_u8(frame, max_dim=_LIVE_PREVIEW_DIM):
    """Downsample + contrast-stretch + gamma-correct a uint16 sensor frame to uint8.

    In 'auto' mode, uses EMA-smoothed percentiles to reduce flicker.
    In 'locked' mode, uses the frozen _disp_vmin/_disp_vmax values.
    Gamma correction (>1 darkens midtones) is applied via a fast uint8 LUT.
    """
    global _disp_vmin, _disp_vmax

    h, w = frame.shape
    scale = max(1, max(h, w) // max_dim) if max_dim else 1
    small = frame[::scale, ::scale] if scale > 1 else frame

    if _disp_mode == "auto":
        lo, hi = _compute_percentiles(small)
        if hi <= lo:
            hi = lo + 1
        alpha = _DISP_EMA_ALPHA
        if _disp_vmin == 0.0 and _disp_vmax == 65535.0:
            _disp_vmin, _disp_vmax = lo, hi
        else:
            _disp_vmin += alpha * (lo - _disp_vmin)
            _disp_vmax += alpha * (hi - _disp_vmax)

    vmin, vmax = _disp_vmin, _disp_vmax
    if vmax <= vmin:
        vmax = vmin + 1

    u8 = np.clip((small.astype(np.float32) - vmin) / (vmax - vmin) * 255,
                  0, 255).astype(np.uint8)

    gamma = _disp_gamma
    if gamma != 1.0:
        u8 = _get_gamma_lut(gamma)[u8]

    return u8


def _frame_to_jpeg_bytes(frame, quality=80, max_dim=_LIVE_PREVIEW_DIM):
    """Convert a 2-D uint16 numpy array to raw JPEG bytes.

    Uses cv2.imencode (libjpeg-turbo, ~3ms) with Pillow fallback (~12ms).
    """
    normed = _normalize_u8(frame, max_dim)

    if _HAS_CV2:
        ok, buf = cv2.imencode(".jpg", normed,
                               [cv2.IMWRITE_JPEG_QUALITY, quality])
        if ok:
            return buf.tobytes()

    if _HAS_PIL:
        img = _PILImage.fromarray(normed, mode="L")
        buf = io.BytesIO()
        img.save(buf, format="JPEG", quality=quality)
        return buf.getvalue()

    raise CamError("No JPEG encoder available (install opencv-python or Pillow)")


def _frame_to_base64(frame, quality=85, max_dim=800):
    """Convert a 2-D uint16 numpy array to base64 JPEG for snap responses."""
    return base64.b64encode(_frame_to_jpeg_bytes(frame, quality, max_dim)).decode("ascii")


def _ensure_save_dir():
    os.makedirs(_save_dir, exist_ok=True)


def _timestamp():
    return datetime.now().strftime("%Y%m%d_%H%M%S")


# ── Connection ───────────────────────────────────────────────────────────────

def connect():
    global _hcam, _pvc_initialized, _sensor_w, _sensor_h, _bit_depth, _cam_name_str
    with _lock:
        if _hcam is not None:
            return  # already open — nothing to do
        if not _pvc_initialized:
            pvc.init()
            _pvc_initialized = True
        n = pvc.cam_count()
        if n < 1:
            raise CamError("No PVCAM cameras found")
        name = pvc.cam_name(0)
        _hcam = pvc.cam_open(name)
        _sensor_w, _sensor_h = pvc.sensor_size(_hcam)
        _bit_depth = pvc.bit_depth(_hcam)
        try:
            _cam_name_str = pvc.chip_name(_hcam)
        except Exception:
            _cam_name_str = name


def disconnect():
    global _hcam, _pvc_initialized
    live_stop()
    capture_stop()
    with _lock:
        if _hcam is not None:
            try:
                pvc.cam_close(_hcam)
            except Exception:
                pass
            _hcam = None
        if _pvc_initialized:
            try:
                pvc.uninit()
            except Exception:
                pass
            _pvc_initialized = False


def is_connected():
    return _hcam is not None


def get_camera_info():
    if _hcam is None:
        raise CamError("Camera not connected")
    with _lock:
        return {
            "name": _cam_name_str,
            "sensor_size": [_sensor_w, _sensor_h],
            "bit_depth": _bit_depth,
        }


# ── Settings ─────────────────────────────────────────────────────────────────

def set_exposure(ms):
    global _exposure_ms
    _exposure_ms = max(1, int(ms))


def get_exposure():
    return _exposure_ms


def set_binning(b):
    global _binning
    b = int(b)
    if b not in (1, 2, 4, 8):
        raise CamError(f"Binning must be 1, 2, 4, or 8 (got {b})")
    _binning = (b, b)


def get_binning():
    return _binning[0]


# ── Snap (single frame) ─────────────────────────────────────────────────────

def snap():
    """Capture a single frame via continuous-mode poll.

    We use setup_cont + poll rather than setup_seq because the CoolSNAP EZ
    (FireWire) is more reliable in continuous mode with short polling.
    """
    if _hcam is None:
        raise CamError("Camera not connected")
    with _lock:
        import ctypes
        frame_bytes = pvc.setup_cont(_hcam, _exposure_ms, _binning[0])
        n_frames = 2
        buf = (pvc.uns16 * (frame_bytes * n_frames // 2))()
        pvc.start_cont(_hcam, buf, frame_bytes * n_frames)
        try:
            frame = pvc.poll_frame_numpy(
                _hcam, _sensor_w, _sensor_h, _binning[0], timeout_s=10,
            )
            return frame
        except TimeoutError:
            raise CamError("Snap timed out (10s) — no frame received")
        finally:
            pvc.abort(_hcam)


def snap_and_save():
    """Snap a single frame, save to disk, return (frame, path)."""
    frame = snap()
    _ensure_save_dir()
    path = os.path.join(_save_dir, f"snap_{_timestamp()}.npy")
    np.save(path, frame)
    return frame, path


# ── Live view ────────────────────────────────────────────────────────────────

def live_start():
    global _live_thread
    if _hcam is None:
        raise CamError("Camera not connected")
    if _live_thread and _live_thread.is_alive():
        return
    _live_stop.clear()
    _live_thread = threading.Thread(target=_live_loop, daemon=True)
    _live_thread.start()


def live_stop():
    _live_stop.set()
    if _live_thread and _live_thread.is_alive():
        _live_thread.join(timeout=3)


def live_is_active():
    return _live_thread is not None and _live_thread.is_alive()


def _live_loop():
    """Background thread: continuously poll frames from the camera.

    Uses pvcam_raw's continuous acquisition mode with a 2-frame circular buffer.
    Each new frame is JPEG-encoded and stored for the WebSocket streamer.
    """
    global _latest_jpeg, _live_fps, _circ_buf, _circ_buf_size
    try:
        with _lock:
            frame_bytes = pvc.setup_cont(_hcam, _exposure_ms, _binning[0])
            n_circ = 2
            _circ_buf_size = frame_bytes * n_circ
            _circ_buf = (pvc.uns16 * (_circ_buf_size // 2))()
            pvc.start_cont(_hcam, _circ_buf, _circ_buf_size)

        t0 = time.monotonic()
        frames_count = 0

        while not _live_stop.is_set():
            try:
                status, _, _ = pvc.check_cont_status(_hcam)
                if status >= pvc.FRAME_AVAILABLE:
                    ptr = pvc.get_latest_frame(_hcam)
                    frame = pvc.frame_to_numpy(
                        ptr, _sensor_w, _sensor_h, _binning[0],
                    )
                    frames_count += 1
                    elapsed = time.monotonic() - t0
                    if elapsed > 0:
                        _live_fps = frames_count / elapsed
                    if elapsed > 2.0:
                        t0 = time.monotonic()
                        frames_count = 0

                    jpeg = _frame_to_jpeg_bytes(
                        frame, quality=75, max_dim=_LIVE_PREVIEW_DIM,
                    )
                    with _frame_lock:
                        _latest_jpeg = jpeg
                    _frame_event.set()
                else:
                    time.sleep(0.002)
            except Exception:
                if _live_stop.is_set():
                    break
                time.sleep(0.01)
    finally:
        try:
            with _lock:
                pvc.abort(_hcam)
        except Exception:
            pass
        _circ_buf = None
        with _frame_lock:
            _latest_jpeg = None
        _frame_event.set()


def get_live_jpeg():
    """Return the latest JPEG bytes (for MJPEG HTTP endpoint)."""
    with _frame_lock:
        return _latest_jpeg


def get_live_fps():
    return round(_live_fps, 1)


# ── Video recording ──────────────────────────────────────────────────────────

def record_video(num_frames=100):
    """Record num_frames into a 3-D numpy array (frames, H, W)."""
    if _hcam is None:
        raise CamError("Camera not connected")

    _capture_stop.clear()
    frames = []
    with _lock:
        frame_bytes = pvc.setup_cont(_hcam, _exposure_ms, _binning[0])
        n_circ = 2
        buf_size = frame_bytes * n_circ
        buf = (pvc.uns16 * (buf_size // 2))()
        pvc.start_cont(_hcam, buf, buf_size)
    try:
        collected = 0
        while collected < num_frames:
            if _capture_stop.is_set():
                break
            try:
                status, _, _ = pvc.check_cont_status(_hcam)
                if status >= pvc.FRAME_AVAILABLE:
                    ptr = pvc.get_latest_frame(_hcam)
                    frame = pvc.frame_to_numpy(
                        ptr, _sensor_w, _sensor_h, _binning[0],
                    )
                    frames.append(frame)
                    collected += 1
                else:
                    time.sleep(0.001)
            except Exception:
                time.sleep(0.001)
    finally:
        with _lock:
            pvc.abort(_hcam)

    if not frames:
        raise CamError("No frames captured")
    return np.stack(frames, axis=0)


def record_video_and_save(num_frames=100):
    """Record video and save to .npy. Returns filepath."""
    video = record_video(num_frames)
    _ensure_save_dir()
    path = os.path.join(_save_dir, f"video_{_timestamp()}.npy")
    np.save(path, video)
    return path


# ── Time-lapse ───────────────────────────────────────────────────────────────

def timelapse(num_frames=10, interval_sec=5.0):
    """Capture num_frames images at interval_sec apart, return 3-D stack."""
    if _cam is None:
        raise CamError("Camera not connected")

    _capture_stop.clear()
    frames = []
    for i in range(num_frames):
        if _capture_stop.is_set():
            break
        frame = snap()
        frames.append(frame.copy())
        push_event("onTimelapseProgress", i + 1, num_frames)
        if i < num_frames - 1:
            deadline = time.monotonic() + interval_sec
            while time.monotonic() < deadline:
                if _capture_stop.is_set():
                    break
                time.sleep(0.1)

    if not frames:
        raise CamError("No frames captured")
    return np.stack(frames, axis=0)


def timelapse_and_save(num_frames=10, interval_sec=5.0):
    """Run time-lapse and save. Returns filepath."""
    stack = timelapse(num_frames, interval_sec)
    _ensure_save_dir()
    path = os.path.join(_save_dir, f"timelapse_{_timestamp()}.npy")
    np.save(path, stack)
    return path


# ── Background capture thread ────────────────────────────────────────────────

def capture_stop():
    _capture_stop.set()
    if _capture_thread and _capture_thread.is_alive():
        _capture_thread.join(timeout=10)


def _run_capture(mode, **kwargs):
    """Run a capture task in the background."""
    global _capture_thread
    if _capture_thread and _capture_thread.is_alive():
        raise CamError("A capture is already in progress")
    _capture_stop.clear()
    _capture_thread = threading.Thread(
        target=_capture_worker, args=(mode,), kwargs=kwargs, daemon=True)
    _capture_thread.start()


def _capture_worker(mode, **kwargs):
    try:
        if mode == "video":
            push_event("onCamStatus", "recording",
                        f"Recording {kwargs.get('num_frames', 100)} frames...")
            path = record_video_and_save(**kwargs)
        elif mode == "timelapse":
            n = kwargs.get('num_frames', 10)
            iv = kwargs.get('interval_sec', 5)
            push_event("onCamStatus", "recording",
                        f"Time-lapse: {n} frames, {iv}s interval...")
            path = timelapse_and_save(**kwargs)
        elif mode == "snap":
            push_event("onCamStatus", "snapping", "Snapping image...")
            _, path = snap_and_save()
        else:
            return

        fname = os.path.basename(path)
        push_event("onCamStatus", "idle", f"Saved: {fname}")
        push_event("onCamCaptureComplete", fname)
    except Exception as e:
        push_event("onCamStatus", "error", str(e))


# ── Eel-exposed wrappers ────────────────────────────────────────────────────

def _wrap(fn, *args, **kwargs):
    try:
        result = fn(*args, **kwargs)
        if isinstance(result, dict):
            return {"ok": True, **result}
        return {"ok": True, "value": result}
    except Exception as e:
        return {"error": str(e)}


@expose
def cam_connect():
    try:
        connect()
        info = get_camera_info()
        return {"ok": True, "name": info["name"],
                "sensor": info["sensor_size"], "bit_depth": info["bit_depth"]}
    except Exception as e:
        return {"error": str(e)}


@expose
def cam_disconnect():
    try:
        disconnect()
        return {"ok": True}
    except Exception as e:
        return {"error": str(e)}


@expose
def cam_is_connected():
    return is_connected()


@expose
def cam_info():
    return _wrap(get_camera_info)


@expose
def cam_set_exposure(ms):
    set_exposure(ms)
    return {"ok": True, "exposure_ms": _exposure_ms}


@expose
def cam_get_exposure():
    return _exposure_ms


@expose
def cam_set_binning(b):
    return _wrap(set_binning, b)


@expose
def cam_get_binning():
    return get_binning()


@expose
def cam_snap():
    """Snap and return base64 image for preview."""
    try:
        frame = snap()
        b64 = _frame_to_base64(frame)
        return {"ok": True, "image": b64,
                "width": int(frame.shape[1]), "height": int(frame.shape[0])}
    except Exception as e:
        return {"error": str(e)}


@expose
def cam_snap_save():
    """Snap, save to disk, and return base64 + filename."""
    try:
        frame, path = snap_and_save()
        b64 = _frame_to_base64(frame)
        return {"ok": True, "image": b64, "file": os.path.basename(path),
                "width": int(frame.shape[1]), "height": int(frame.shape[0])}
    except Exception as e:
        return {"error": str(e)}


@expose
def cam_live_start():
    return _wrap(live_start)


@expose
def cam_live_stop():
    live_stop()
    return {"ok": True}


@expose
def cam_live_active():
    return live_is_active()


@expose
def cam_record_video(num_frames=100):
    try:
        _run_capture("video", num_frames=int(num_frames))
        return {"ok": True, "msg": f"Recording {num_frames} frames..."}
    except Exception as e:
        return {"error": str(e)}


@expose
def cam_timelapse(num_frames=10, interval_sec=5):
    try:
        _run_capture("timelapse", num_frames=int(num_frames),
                     interval_sec=float(interval_sec))
        return {"ok": True, "msg": f"Time-lapse: {num_frames} frames, {interval_sec}s apart"}
    except Exception as e:
        return {"error": str(e)}


@expose
def cam_capture_stop():
    capture_stop()
    return {"ok": True}


@expose
def cam_set_save_dir(path):
    global _save_dir
    _save_dir = path
    os.makedirs(_save_dir, exist_ok=True)
    return {"ok": True, "path": _save_dir}


@expose
def cam_get_save_dir():
    return _save_dir


@expose
def cam_display_mode(mode=None):
    """Get or set display mode ('auto' or 'locked')."""
    global _disp_mode
    if mode is not None:
        if mode not in ("auto", "locked"):
            return {"error": "mode must be 'auto' or 'locked'"}
        _disp_mode = mode
        if mode == "auto":
            global _disp_vmin, _disp_vmax
            _disp_vmin, _disp_vmax = 0.0, 65535.0
    return {"mode": _disp_mode, "vmin": round(_disp_vmin), "vmax": round(_disp_vmax),
            "gamma": _disp_gamma}


@expose
def cam_auto_adjust():
    """One-shot: compute display range from current frame and lock it."""
    global _disp_mode, _disp_vmin, _disp_vmax
    jpeg = get_live_jpeg()
    if jpeg is None:
        if _hcam is None:
            return {"error": "Camera not connected"}
        try:
            frame = snap()
        except Exception as e:
            return {"error": str(e)}
        lo, hi = _compute_percentiles(frame)
    else:
        return {"mode": _disp_mode, "vmin": round(_disp_vmin), "vmax": round(_disp_vmax)}
    if hi <= lo:
        hi = lo + 1
    _disp_vmin, _disp_vmax = lo, hi
    _disp_mode = "locked"
    return {"mode": "locked", "vmin": round(lo), "vmax": round(hi), "gamma": _disp_gamma}


@expose
def cam_set_display_range(vmin, vmax):
    """Manually set display min/max and lock."""
    global _disp_mode, _disp_vmin, _disp_vmax
    _disp_vmin = float(vmin)
    _disp_vmax = float(vmax)
    _disp_mode = "locked"
    return {"mode": "locked", "vmin": round(_disp_vmin), "vmax": round(_disp_vmax),
            "gamma": _disp_gamma}


@expose
def cam_get_display_range():
    """Return current display range, mode, and gamma."""
    return {"mode": _disp_mode, "vmin": round(_disp_vmin), "vmax": round(_disp_vmax),
            "gamma": _disp_gamma}


@expose
def cam_set_gamma(gamma):
    """Set gamma correction (>1 darkens midtones, <1 brightens). Range: 0.2–5.0."""
    global _disp_gamma
    gamma = float(gamma)
    gamma = max(0.2, min(5.0, gamma))
    _disp_gamma = round(gamma, 2)
    return {"gamma": _disp_gamma}


@expose
def cam_get_gamma():
    return {"gamma": _disp_gamma}
