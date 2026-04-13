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
import json
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
_live_exposure_changed = threading.Event()

_capture_thread = None
_capture_stop = threading.Event()

_save_dir = os.path.join(os.path.dirname(__file__), "..", "captures")
_base_save_dir = _save_dir

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

# Pseudo-color: None = grayscale, or (R, G, B) tuple for tinting
_pseudo_color = None
_PSEUDO_COLOR_MAP = {
    "none":    None,
    "blue":    (0, 0, 255),
    "green":   (0, 255, 0),
    "red":     (255, 0, 0),
    "cyan":    (0, 255, 255),
    "magenta": (255, 0, 255),
    "yellow":  (255, 255, 0),
}
_pseudo_color_name = "none"

_FILTER_POS_COLOR = {
    1: "none",      # DIA (brightfield)
    2: "blue",      # DAPI
    3: "green",     # GFP
    4: "red",       # TxRed
}


class CamError(Exception):
    """Raised when a camera operation fails."""


# ── Helpers ──────────────────────────────────────────────────────────────────

_LIVE_PREVIEW_DIM = 800

def _auto_range(frame):
    """Simple data-range display: map actual [min, max] to [0, 255].

    Like ImageJ on first open — no histogram tricks, just the real range.
    Uses subsampled data for speed.
    """
    flat = frame.ravel()
    sub = flat[::max(1, flat.size // 50000)]
    lo, hi = float(sub.min()), float(sub.max())
    if hi <= lo:
        hi = lo + 1
    return lo, hi


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
        lo, hi = _auto_range(small)
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

    pc = _pseudo_color
    if pc is not None:
        r, g, b = pc
        rgb = np.zeros((*u8.shape, 3), dtype=np.uint8)
        if r: rgb[:, :, 0] = (u8.astype(np.uint16) * r // 255).astype(np.uint8)
        if g: rgb[:, :, 1] = (u8.astype(np.uint16) * g // 255).astype(np.uint8)
        if b: rgb[:, :, 2] = (u8.astype(np.uint16) * b // 255).astype(np.uint8)
        return rgb

    return u8


def _frame_to_jpeg_bytes(frame, quality=80, max_dim=_LIVE_PREVIEW_DIM):
    """Convert a 2-D uint16 numpy array to raw JPEG bytes.

    Returns grayscale or RGB JPEG depending on pseudo-color setting.
    Uses cv2.imencode (libjpeg-turbo, ~3ms) with Pillow fallback (~12ms).
    """
    normed = _normalize_u8(frame, max_dim)

    if _HAS_CV2:
        if normed.ndim == 3:
            normed = cv2.cvtColor(normed, cv2.COLOR_RGB2BGR)
        ok, buf = cv2.imencode(".jpg", normed,
                               [cv2.IMWRITE_JPEG_QUALITY, quality])
        if ok:
            return buf.tobytes()

    if _HAS_PIL:
        pil_mode = "RGB" if normed.ndim == 3 else "L"
        img = _PILImage.fromarray(normed, mode=pil_mode)
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


def _safe(fn):
    """Call fn(), return result or None on any error/timeout."""
    try:
        return fn()
    except Exception:
        return None


def _gather_system_state():
    """Snapshot all hardware state. Returns None for anything disconnected."""
    state = {}

    # Camera
    state["camera"] = {
        "connected": _hcam is not None,
        "name": _cam_name_str if _hcam else None,
        "sensor": [_sensor_w, _sensor_h] if _hcam else None,
        "bit_depth": _bit_depth if _hcam else None,
        "exposure_ms": _exposure_ms,
        "binning": _binning[0],
        "pseudo_color": _pseudo_color_name,
        "display_mode": _disp_mode,
        "display_range": [round(_disp_vmin), round(_disp_vmax)],
        "gamma": _disp_gamma,
    }

    # Microscope
    try:
        import modules.nikon_ti as ti
        if ti.is_connected():
            state["microscope"] = {
                "connected": True,
                "objective": _safe(ti.nosepiece_get_position),
                "filter": _safe(ti.filter_get_position),
                "shutter": _safe(ti.shutter_get_state),
                "lamp": _safe(ti.dia_lamp_get_state),
                "z_nm": _safe(ti.z_get_position),
                "xy_nm": _safe(ti.xy_get_position),
                "pfs": _safe(ti.pfs_get_status),
            }
        else:
            state["microscope"] = {"connected": False}
    except Exception:
        state["microscope"] = None

    # Intensilight
    try:
        import modules.intensilight as il
        if il.is_connected():
            state["intensilight"] = _safe(il.get_state)
        else:
            state["intensilight"] = {"connected": False}
    except Exception:
        state["intensilight"] = None

    # Cavro pumps (read in-memory attrs only — no serial I/O)
    try:
        import modules.cavro as cavro
        pumps = []
        for i in range(cavro.NUM_CAVRO):
            p = cavro.get_pump(i)
            if p is not None:
                pumps.append({
                    "index": i,
                    "address": getattr(p, "address", None),
                    "syringe_ml": getattr(p, "syringe_volume_ml", None),
                    "continuous_running": (
                        cavro._cont_threads[i] is not None
                        and cavro._cont_threads[i].is_alive()),
                    "coordinated_running": (
                        cavro._coord_thread is not None
                        and cavro._coord_thread.is_alive()
                        and i in cavro._coord_pump_idxs),
                })
            else:
                pumps.append(None)
        state["cavro"] = {
            "connected": cavro._serial is not None,
            "pumps": pumps,
        }
    except Exception:
        state["cavro"] = None

    return state


def _save_meta(npy_path, colors=None, channels=None):
    """Write a sidecar .meta.json with full system state snapshot."""
    meta = {
        "timestamp": datetime.now().isoformat(),
        "color": _pseudo_color_name,
    }
    if colors:
        meta["colors"] = colors
    if channels:
        meta["channels"] = channels
    meta["system"] = _gather_system_state()
    try:
        with open(npy_path + ".meta.json", "w") as f:
            json.dump(meta, f, indent=2, default=str)
    except Exception:
        pass


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
    _live_exposure_changed.set()


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
    _save_meta(path)
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


_latest_raw_frame = None

def _live_loop():
    """Background thread: continuously poll frames from the camera.

    Uses pvcam_raw's continuous acquisition mode with a 2-frame circular buffer.
    Each new frame is JPEG-encoded and stored for the WebSocket streamer.
    """
    global _latest_jpeg, _live_fps, _circ_buf, _circ_buf_size, _latest_raw_frame
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
            if _live_exposure_changed.is_set():
                _live_exposure_changed.clear()
                try:
                    with _lock:
                        pvc.abort(_hcam)
                        frame_bytes = pvc.setup_cont(_hcam, _exposure_ms, _binning[0])
                        n_circ = 2
                        _circ_buf_size = frame_bytes * n_circ
                        _circ_buf = (pvc.uns16 * (_circ_buf_size // 2))()
                        pvc.start_cont(_hcam, _circ_buf, _circ_buf_size)
                except Exception:
                    if _live_stop.is_set():
                        break
                    time.sleep(0.01)
                    continue

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

                    _latest_raw_frame = frame.copy()
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

def record_video(num_frames=100, duration_sec=None, fps=None):
    """Record video into a 3-D numpy array (frames, H, W).

    If duration_sec is provided, records for that many seconds.
    fps controls frame pacing; None = capture as fast as possible.
    """
    if _hcam is None:
        raise CamError("Camera not connected")

    _capture_stop.clear()
    frames = []
    frame_interval = 1.0 / fps if fps else None
    with _lock:
        frame_bytes = pvc.setup_cont(_hcam, _exposure_ms, _binning[0])
        n_circ = 2
        buf_size = frame_bytes * n_circ
        buf = (pvc.uns16 * (buf_size // 2))()
        pvc.start_cont(_hcam, buf, buf_size)
    try:
        collected = 0
        t0 = time.monotonic()
        next_frame_time = t0
        while True:
            if _capture_stop.is_set():
                break
            if duration_sec is not None:
                if time.monotonic() - t0 >= duration_sec:
                    break
            else:
                if collected >= num_frames:
                    break
            try:
                status, _, _ = pvc.check_cont_status(_hcam)
                if status >= pvc.FRAME_AVAILABLE:
                    now = time.monotonic()
                    if frame_interval and now < next_frame_time:
                        time.sleep(max(0, next_frame_time - now))
                    ptr = pvc.get_latest_frame(_hcam)
                    frame = pvc.frame_to_numpy(
                        ptr, _sensor_w, _sensor_h, _binning[0],
                    )
                    frames.append(frame)
                    collected += 1
                    if frame_interval:
                        next_frame_time = t0 + collected * frame_interval
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


def record_video_and_save(num_frames=100, duration_sec=None, fps=None):
    """Record video and save to .npy. Returns filepath."""
    global _pseudo_color, _pseudo_color_name
    video = record_video(num_frames, duration_sec=duration_sec, fps=fps)
    _ensure_save_dir()
    path = os.path.join(_save_dir, f"video_{_timestamp()}.npy")
    np.save(path, video)
    prev_color, prev_name = _pseudo_color, _pseudo_color_name
    _pseudo_color, _pseudo_color_name = None, "none"
    _save_meta(path)
    _pseudo_color, _pseudo_color_name = prev_color, prev_name
    return path


# ── Time-lapse ───────────────────────────────────────────────────────────────

def timelapse(num_frames=10, interval_sec=5.0):
    """Capture num_frames images at interval_sec apart, return 3-D stack.

    The interval is measured from the START of each frame, so if snap takes
    300ms and interval is 1s, the wait after snap is only 700ms.
    """
    if _hcam is None:
        raise CamError("Camera not connected")

    _capture_stop.clear()
    frames = []
    for i in range(num_frames):
        if _capture_stop.is_set():
            break
        frame_t0 = time.monotonic()
        frame = snap()
        frames.append(frame.copy())
        push_event("onTimelapseProgress", i + 1, num_frames)
        if i < num_frames - 1:
            remaining = interval_sec - (time.monotonic() - frame_t0)
            if remaining > 0:
                deadline = time.monotonic() + remaining
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
    _save_meta(path)
    return path


# ── Image stack accumulation ─────────────────────────────────────────────────

_stack_frames = []
_stack_lock = threading.Lock()


_stack_was_live = False

def stack_begin():
    """Clear the frame buffer and stop live preview to avoid conflicts."""
    global _stack_frames, _stack_was_live
    _stack_was_live = live_is_active()
    if _stack_was_live:
        live_stop()
    with _stack_lock:
        _stack_frames = []


def stack_snap():
    """Snap one frame and append to the stack buffer. Returns frame index."""
    frame = snap()
    with _stack_lock:
        _stack_frames.append(frame.copy())
        return len(_stack_frames)


def stack_finish():
    """Save all accumulated frames as a single .npy and push event."""
    global _stack_frames
    with _stack_lock:
        if not _stack_frames:
            raise CamError("No frames in stack")
        stacked = np.stack(_stack_frames, axis=0)
        _stack_frames = []
    _ensure_save_dir()
    path = os.path.join(_save_dir, f"stack_{_timestamp()}.npy")
    np.save(path, stacked)
    _save_meta(path)
    fname = os.path.basename(path)
    push_event("onCamStatus", "idle", f"Saved: {fname}")
    push_event("onCamCaptureComplete", fname)
    if _stack_was_live:
        live_start()
    return path


def stack_capture(channels, keep_shutter_open=False):
    """Execute an entire stack acquisition server-side for speed.

    channels: list of dicts, each with:
      objective, filter, lamp_intensity, lamp_on, exposure_ms, binning,
      il_shutter_open, il_nd, shutter_open
    """
    import modules.nikon_ti as ti
    import modules.intensilight as il

    if _capture_thread and _capture_thread.is_alive():
        _capture_thread.join(timeout=10)

    was_live = live_is_active()
    if was_live:
        live_stop()

    frames = []
    colors = []
    use_il = il.is_connected()
    try:
        use_scope = ti.is_connected()

        for i, ch in enumerate(channels):
            push_event("onCamStatus", "busy",
                        f"Stack ch {i+1}/{len(channels)}")

            if use_scope:
                ti.apply_preset(
                    objective=ch.get("objective"),
                    filter_pos=ch.get("filter"),
                    lamp_intensity=ch.get("lamp_intensity"),
                    lamp_on=ch.get("lamp_on"),
                    shutter_open=ch.get("shutter_open"),
                )

            if use_il:
                il_nd = ch.get("il_nd")
                if il_nd is not None:
                    il.nd_set(int(il_nd))
                il_sh = ch.get("il_shutter_open")
                if il_sh is not None:
                    if il_sh:
                        il.shutter_open()
                    else:
                        il.shutter_close()

            if use_scope or use_il:
                time.sleep(0.5)

            exp = ch.get("exposure_ms")
            if exp is not None:
                set_exposure(int(exp))
            binn = ch.get("binning")
            if binn is not None:
                set_binning(int(binn))

            frame = snap()
            frames.append(frame.copy())
            ch_color = ch.get("color", "auto")
            if ch_color and ch_color != "auto":
                colors.append(ch_color)
            else:
                filt = ch.get("filter")
                colors.append(_FILTER_POS_COLOR.get(filt, "none") if filt else "none")

    finally:
        try:
            if use_scope:
                ti.shutter_close()
                ti.dia_lamp_off()
            if use_il:
                il.shutter_close()
        except Exception:
            pass
        if was_live:
            live_start()

    if not frames:
        raise CamError("No frames captured")

    stacked = np.stack(frames, axis=0)
    _ensure_save_dir()
    path = os.path.join(_save_dir, f"stack_{_timestamp()}.npy")
    np.save(path, stacked)
    _save_meta(path, colors=colors, channels=channels)
    fname = os.path.basename(path)
    push_event("onCamStatus", "idle", f"Saved: {fname}")
    push_event("onCamCaptureComplete", fname)
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
        _capture_thread.join(timeout=5)
        if _capture_thread.is_alive():
            raise CamError("A capture is already in progress")
    _capture_stop.clear()
    _capture_thread = threading.Thread(
        target=_capture_worker, args=(mode,), kwargs=kwargs, daemon=True)
    _capture_thread.start()


def _capture_worker(mode, **kwargs):
    was_live = live_is_active()
    if was_live:
        live_stop()
    fname = None
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
    except Exception as e:
        push_event("onCamStatus", "error", str(e))
    finally:
        if was_live:
            live_start()
        push_event("onCamCaptureComplete", fname)


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
def cam_record_video(num_frames=100, duration_sec=None, fps=None):
    try:
        kw = {}
        if duration_sec is not None:
            kw["duration_sec"] = float(duration_sec)
        else:
            kw["num_frames"] = int(num_frames)
        if fps is not None:
            kw["fps"] = float(fps)
        _run_capture("video", **kw)
        return {"ok": True, "msg": f"Recording..."}
    except Exception as e:
        return {"error": str(e)}


@expose
def cam_record_timelapse(num_frames=10, interval_sec=5.0):
    try:
        _run_capture("timelapse", num_frames=int(num_frames),
                     interval_sec=float(interval_sec))
        return {"ok": True, "msg": f"Timelapse: {num_frames} frames, {interval_sec}s interval..."}
    except Exception as e:
        return {"error": str(e)}


@expose
def cam_stack_begin():
    try:
        stack_begin()
        return {"ok": True}
    except Exception as e:
        return {"error": str(e)}


@expose
def cam_stack_snap():
    try:
        idx = stack_snap()
        return {"ok": True, "frame": idx}
    except Exception as e:
        return {"error": str(e)}


@expose
def cam_stack_finish():
    try:
        path = stack_finish()
        return {"ok": True, "file": os.path.basename(path)}
    except Exception as e:
        return {"error": str(e)}


@expose
def cam_stack_capture(channels, keep_shutter_open=False):
    """Execute full stack server-side. channels is a list of channel dicts."""
    try:
        path = stack_capture(channels, keep_shutter_open)
        return {"ok": True, "file": os.path.basename(path)}
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
    global _save_dir, _base_save_dir
    _save_dir = path
    _base_save_dir = path
    os.makedirs(_save_dir, exist_ok=True)
    return {"ok": True, "path": _save_dir}


@expose
def cam_get_save_dir():
    return _save_dir


@expose
def cam_experiment_start(name):
    """Create an experiment folder with timelapse/video/stack subfolders."""
    global _save_dir
    exp_dir = os.path.join(_base_save_dir, name)
    for sub in ("timelapse", "video", "stack"):
        os.makedirs(os.path.join(exp_dir, sub), exist_ok=True)
    _save_dir = exp_dir
    return {"ok": True, "path": exp_dir}


@expose
def cam_experiment_set_subdir(subdir):
    """Switch save target to a subfolder within the current experiment."""
    global _save_dir
    exp_dir = os.path.dirname(_save_dir) if os.path.basename(_save_dir) in ("timelapse", "video", "stack") else _save_dir
    _save_dir = os.path.join(exp_dir, subdir)
    os.makedirs(_save_dir, exist_ok=True)
    return {"ok": True, "path": _save_dir}


@expose
def cam_experiment_end():
    """Reset save directory back to the base captures folder."""
    global _save_dir
    _save_dir = _base_save_dir
    return {"ok": True, "path": _save_dir}


@expose
def cam_list_captures(subdir=None):
    """List .npy files in a directory, newest first.

    subdir: optional relative path under _base_save_dir (e.g. "exp1/stack").
    If None, lists files in _base_save_dir itself.
    """
    base = os.path.normpath(_base_save_dir)
    if subdir:
        target = os.path.normpath(os.path.join(base, subdir))
    else:
        target = base
    if not os.path.isdir(target):
        return {"files": []}
    files = [f for f in os.listdir(target) if f.endswith(".npy")]
    files.sort(key=lambda f: os.path.getmtime(os.path.join(target, f)), reverse=True)
    result = []
    for f in files:
        path = os.path.join(target, f)
        try:
            arr = np.load(path, mmap_mode='r')
            shape = [int(d) for d in arr.shape]
        except Exception:
            shape = []
        result.append({"name": f, "shape": shape})
    return {"files": result}


@expose
def cam_list_experiments():
    """List experiment folders in _base_save_dir (those containing subfolders)."""
    base = os.path.normpath(_base_save_dir)
    if not os.path.isdir(base):
        return {"experiments": []}
    experiments = []
    for name in sorted(os.listdir(base)):
        d = os.path.join(base, name)
        if os.path.isdir(d):
            subs = [s for s in ("video", "stack", "timelapse")
                    if os.path.isdir(os.path.join(d, s))]
            if subs:
                counts = {}
                for s in subs:
                    sd = os.path.join(d, s)
                    counts[s] = len([f for f in os.listdir(sd) if f.endswith(".npy")])
                experiments.append({"name": name, "types": subs, "counts": counts})
    return {"experiments": experiments}


def _apply_bcg(frame, brightness=0, contrast=100, gamma=1.0, max_dim=1200):
    """Apply brightness/contrast/gamma adjustments to a uint16 frame → uint8.

    At default settings (b=0, c=100, g=1.0) maps the full data range to 0-255.
    contrast > 100 tightens the window (brighter), < 100 widens it (darker).
    brightness shifts the window center.
    gamma: 0.1-5.0 (1.0 = linear). Applied via LUT.
    """
    h, w = frame.shape
    scale = max(1, max(h, w) // max_dim) if max_dim else 1
    small = frame[::scale, ::scale] if scale > 1 else frame

    lo, hi = _auto_range(small)
    rng = hi - lo
    center = (lo + hi) / 2.0

    center += brightness / 100.0 * rng

    if contrast != 100:
        scale_f = max(100.0 / max(contrast, 1), 0.01)
        rng *= scale_f

    vmin = center - rng / 2.0
    vmax = center + rng / 2.0
    if vmax <= vmin:
        vmax = vmin + 1

    u8 = np.clip((small.astype(np.float32) - vmin) / (vmax - vmin) * 255,
                  0, 255).astype(np.uint8)

    if gamma != 1.0:
        lut = _get_gamma_lut(gamma)
        u8 = lut[u8]

    pc = _pseudo_color
    if pc is not None:
        r, g, b = pc
        rgb = np.zeros((*u8.shape, 3), dtype=np.uint8)
        if r: rgb[:, :, 0] = (u8.astype(np.uint16) * r // 255).astype(np.uint8)
        if g: rgb[:, :, 1] = (u8.astype(np.uint16) * g // 255).astype(np.uint8)
        if b: rgb[:, :, 2] = (u8.astype(np.uint16) * b // 255).astype(np.uint8)
        return rgb

    return u8


@expose
def cam_npy_auto_adjust(filename, frame_idx=0, subdir=None):
    """Analyze a frame and return suggested brightness/contrast/gamma values."""
    base = os.path.normpath(_base_save_dir) if subdir else os.path.normpath(_save_dir)
    if subdir:
        base = os.path.join(base, subdir)
    path = os.path.join(base, filename)
    if not os.path.isfile(path):
        return {"error": f"File not found: {filename}"}
    try:
        arr = np.load(path)
    except Exception as e:
        return {"error": f"Cannot load: {e}"}

    frame_idx = int(frame_idx)
    if arr.ndim == 3:
        frame_idx = max(0, min(frame_idx, arr.shape[0] - 1))
        frame = arr[frame_idx]
    elif arr.ndim == 2:
        frame = arr
    else:
        return {"error": f"Unsupported shape: {arr.shape}"}

    h, w = frame.shape
    step = max(1, max(h, w) // 800)
    small = frame[::step, ::step]

    flat = small.ravel()
    data_min, data_max = float(flat.min()), float(flat.max())
    full_rng = data_max - data_min
    if full_rng <= 0:
        return {"ok": True, "brightness": 0, "contrast": 100, "gamma": 100}

    hist, _ = np.histogram(flat, bins=256, range=(data_min, data_max))
    n_px = int(hist.sum())
    limit = n_px // 10
    threshold = max(1, n_px // 5000)
    bin_w = full_rng / 256.0

    hmin = 0
    for i in range(256):
        c = int(hist[i])
        if c > limit:
            c = 0
        if c > threshold:
            hmin = i
            break

    hmax = 255
    for i in range(255, -1, -1):
        c = int(hist[i])
        if c > limit:
            c = 0
        if c > threshold:
            hmax = i
            break

    if hmin >= hmax:
        return {"ok": True, "brightness": 0, "contrast": 100, "gamma": 100}

    auto_lo = data_min + hmin * bin_w
    auto_hi = data_min + (hmax + 1) * bin_w
    auto_rng = auto_hi - auto_lo
    auto_center = (auto_lo + auto_hi) / 2.0
    data_center = (data_min + data_max) / 2.0

    contrast = int(round(100 * full_rng / auto_rng)) if auto_rng > 0 else 100
    contrast = max(50, min(300, contrast))

    brightness = int(round((data_center - auto_center) / full_rng * 100))
    brightness = max(-100, min(100, brightness))

    return {"ok": True, "brightness": brightness, "contrast": contrast, "gamma": 100}


@expose
def cam_npy_histogram(filename, frame_idx=0, bins=256):
    """Return histogram data for a frame in an .npy file."""
    path = os.path.join(os.path.normpath(_save_dir), filename)
    if not os.path.isfile(path):
        return {"error": f"File not found: {filename}"}
    try:
        arr = np.load(path)
    except Exception as e:
        return {"error": f"Cannot load: {e}"}

    frame_idx = int(frame_idx)
    bins = int(bins)
    if arr.ndim == 3:
        frame_idx = max(0, min(frame_idx, arr.shape[0] - 1))
        frame = arr[frame_idx]
    elif arr.ndim == 2:
        frame = arr
    else:
        return {"error": f"Unsupported shape: {arr.shape}"}

    flat = frame.ravel()
    sub = flat[::max(1, flat.size // 100000)]
    data_min, data_max = int(sub.min()), int(sub.max())
    if data_max <= data_min:
        data_max = data_min + 1
    hist, _ = np.histogram(sub, bins=bins, range=(data_min, data_max))
    return {"hist": hist.tolist(), "min": data_min, "max": data_max}


@expose
def cam_npy_stack(filename, mode="max", brightness=0, contrast=100, gamma=1.0, subdir=None):
    """Return a composite of all frames in an .npy file as a base64 JPEG.

    mode: 'max' (max-intensity projection), 'mean' (average), 'sum' (sum clipped),
          'color' (multi-color merge using per-channel pseudo-colors from meta).
    subdir: optional relative path under _base_save_dir.
    """
    base = os.path.normpath(_base_save_dir) if subdir else os.path.normpath(_save_dir)
    if subdir:
        base = os.path.join(base, subdir)
    path = os.path.join(base, filename)
    if not os.path.isfile(path):
        return {"error": f"File not found: {filename}"}
    try:
        arr = np.load(path)
    except Exception as e:
        return {"error": f"Cannot load: {e}"}

    if arr.ndim != 3 or arr.shape[0] < 2:
        return {"error": "Need a multi-frame file to stack"}

    meta_color = "none"
    meta_colors = None
    meta_path = path + ".meta.json"
    if os.path.isfile(meta_path):
        try:
            with open(meta_path) as f:
                meta = json.load(f)
            meta_color = meta.get("color", "none")
            meta_colors = meta.get("colors")
        except Exception:
            pass

    brightness, contrast, gamma = float(brightness), float(contrast), float(gamma)

    if mode == "color":
        h, w = arr.shape[1], arr.shape[2]
        canvas = np.zeros((h, w, 3), dtype=np.float64)
        for i in range(arr.shape[0]):
            frame = arr[i]
            norm = frame.astype(np.float64) / max(1, frame.max())
            c_name = (meta_colors[i] if meta_colors and i < len(meta_colors)
                      else "none")
            rgb = _PSEUDO_COLOR_MAP.get(c_name)
            if rgb:
                for ch in range(3):
                    canvas[:, :, ch] += norm * (rgb[ch] / 255.0)
            else:
                canvas[:, :, 0] += norm
                canvas[:, :, 1] += norm
                canvas[:, :, 2] += norm
        canvas = np.clip(canvas * 255, 0, 255).astype(np.uint8)
        if _HAS_CV2:
            bgr = cv2.cvtColor(canvas, cv2.COLOR_RGB2BGR)
            ok, buf = cv2.imencode(".jpg", bgr,
                                   [cv2.IMWRITE_JPEG_QUALITY, 90])
            b64 = base64.b64encode(buf.tobytes()).decode("ascii") if ok else ""
        elif _HAS_PIL:
            img = _PILImage.fromarray(canvas, mode="RGB")
            buf = io.BytesIO()
            img.save(buf, format="JPEG", quality=90)
            b64 = base64.b64encode(buf.getvalue()).decode("ascii")
        else:
            b64 = ""
        return {
            "image": b64,
            "width": w, "height": h,
            "n_frames": int(arr.shape[0]),
            "mode": "color",
        }

    if mode == "mean":
        composite = np.mean(arr, axis=0).astype(arr.dtype)
    elif mode == "sum":
        composite = np.clip(np.sum(arr.astype(np.float64), axis=0),
                            0, np.iinfo(arr.dtype).max).astype(arr.dtype)
    else:
        composite = np.max(arr, axis=0)

    global _pseudo_color, _pseudo_color_name
    prev_color, prev_name = _pseudo_color, _pseudo_color_name
    _pseudo_color = _PSEUDO_COLOR_MAP.get(meta_color)
    _pseudo_color_name = meta_color
    try:
        u8 = _apply_bcg(composite, brightness, contrast, gamma, max_dim=1200)
        if _HAS_CV2:
            if u8.ndim == 3:
                u8 = cv2.cvtColor(u8, cv2.COLOR_RGB2BGR)
            ok, buf = cv2.imencode(".jpg", u8,
                                   [cv2.IMWRITE_JPEG_QUALITY, 90])
            b64 = base64.b64encode(buf.tobytes()).decode("ascii") if ok else ""
        elif _HAS_PIL:
            pil_mode = "RGB" if u8.ndim == 3 else "L"
            img = _PILImage.fromarray(u8, mode=pil_mode)
            buf = io.BytesIO()
            img.save(buf, format="JPEG", quality=90)
            b64 = base64.b64encode(buf.getvalue()).decode("ascii")
        else:
            b64 = _frame_to_base64(composite, quality=90, max_dim=1200)
    finally:
        _pseudo_color, _pseudo_color_name = prev_color, prev_name

    return {
        "image": b64,
        "width": int(composite.shape[1]),
        "height": int(composite.shape[0]),
        "n_frames": int(arr.shape[0]),
        "mode": mode,
    }


@expose
def cam_npy_preview(filename, frame_idx=0, brightness=0, contrast=100, gamma=1.0, subdir=None):
    """Return a base64 JPEG preview of an .npy file.

    For 3D arrays (video/timelapse), frame_idx selects which frame.
    Reads the .meta.json sidecar to apply the saved pseudo-color.
    brightness/contrast/gamma allow viewer-side image adjustments.
    subdir: optional relative path under _base_save_dir.
    """
    global _pseudo_color, _pseudo_color_name
    base = os.path.normpath(_base_save_dir) if subdir else os.path.normpath(_save_dir)
    if subdir:
        base = os.path.join(base, subdir)
    path = os.path.join(base, filename)
    if not os.path.isfile(path):
        return {"error": f"File not found: {filename}"}
    try:
        arr = np.load(path)
    except Exception as e:
        return {"error": f"Cannot load: {e}"}

    brightness = float(brightness)
    contrast = float(contrast)
    gamma = float(gamma)

    meta_color = "none"
    meta_colors = None
    meta_path = path + ".meta.json"
    if os.path.isfile(meta_path):
        try:
            with open(meta_path) as f:
                meta = json.load(f)
            meta_color = meta.get("color", "none")
            meta_colors = meta.get("colors")
        except Exception:
            pass

    frame_idx = int(frame_idx)
    if arr.ndim == 3:
        n_frames = int(arr.shape[0])
        frame_idx = max(0, min(frame_idx, n_frames - 1))
        frame = arr[frame_idx]
    elif arr.ndim == 2:
        frame = arr
        n_frames = 1
        frame_idx = 0
    else:
        return {"error": f"Unsupported array shape: {arr.shape}"}

    if meta_colors and frame_idx < len(meta_colors):
        frame_color = meta_colors[frame_idx]
    else:
        frame_color = meta_color

    prev_color, prev_name = _pseudo_color, _pseudo_color_name
    _pseudo_color = _PSEUDO_COLOR_MAP.get(frame_color)
    _pseudo_color_name = frame_color
    try:
        u8 = _apply_bcg(frame, brightness, contrast, gamma, max_dim=1200)
        quality = 90
        if _HAS_CV2:
            if u8.ndim == 3:
                u8 = cv2.cvtColor(u8, cv2.COLOR_RGB2BGR)
            ok, buf = cv2.imencode(".jpg", u8,
                                   [cv2.IMWRITE_JPEG_QUALITY, quality])
            if ok:
                b64 = base64.b64encode(buf.tobytes()).decode("ascii")
            else:
                b64 = _frame_to_base64(frame, quality=quality, max_dim=1200)
        elif _HAS_PIL:
            pil_mode = "RGB" if u8.ndim == 3 else "L"
            img = _PILImage.fromarray(u8, mode=pil_mode)
            buf = io.BytesIO()
            img.save(buf, format="JPEG", quality=quality)
            b64 = base64.b64encode(buf.getvalue()).decode("ascii")
        else:
            b64 = _frame_to_base64(frame, quality=quality, max_dim=1200)
    finally:
        _pseudo_color, _pseudo_color_name = prev_color, prev_name
    result = {
        "image": b64,
        "width": int(frame.shape[1]),
        "height": int(frame.shape[0]),
        "n_frames": n_frames,
        "frame_idx": frame_idx,
        "shape": [int(d) for d in arr.shape],
        "color": frame_color,
    }
    if meta_colors:
        result["colors"] = meta_colors
    return result


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
    frame = _latest_raw_frame
    if frame is None:
        if _hcam is None:
            return {"error": "Camera not connected"}
        try:
            frame = snap()
        except Exception as e:
            return {"error": str(e)}
    lo, hi = _auto_range(frame)
    if hi <= lo:
        hi = lo + 1
    _disp_vmin, _disp_vmax = lo, hi
    _disp_mode = "locked"
    return {"mode": "locked", "vmin": round(lo), "vmax": round(hi), "gamma": _disp_gamma}


@expose
def cam_get_histogram(bins=256):
    """Return histogram of the latest live frame for display widget."""
    frame = _latest_raw_frame
    if frame is None:
        return {"error": "No frame available"}
    bins = int(bins)
    flat = frame.ravel()
    sub = flat[::max(1, flat.size // 100000)]
    data_min, data_max = int(sub.min()), int(sub.max())
    if data_max <= data_min:
        data_max = data_min + 1
    hist, edges = np.histogram(sub, bins=bins, range=(data_min, data_max))
    return {
        "hist": hist.tolist(),
        "min": data_min,
        "max": data_max,
        "vmin": round(_disp_vmin),
        "vmax": round(_disp_vmax),
        "mode": _disp_mode,
    }


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


@expose
def cam_set_pseudo_color(color):
    """Set pseudo-color for display. Valid: none, blue, green, red, cyan, magenta, yellow."""
    global _pseudo_color, _pseudo_color_name
    color = str(color).lower().strip()
    if color not in _PSEUDO_COLOR_MAP:
        return {"error": f"Unknown color '{color}'. Valid: {', '.join(_PSEUDO_COLOR_MAP)}"}
    _pseudo_color = _PSEUDO_COLOR_MAP[color]
    _pseudo_color_name = color
    return {"color": _pseudo_color_name}


@expose
def cam_get_pseudo_color():
    return {"color": _pseudo_color_name}
