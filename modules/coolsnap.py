"""
CoolSNAP EZ camera module — direct control via pvcam64.dll (ctypes).

Uses a custom ctypes wrapper (pvcam_raw.py) that calls the PVCAM C library
directly, bypassing PyVCAM entirely. This avoids DLL-version mismatches
between the Python package and the installed PVCAM driver.

Provides:
  - Live view (WebSocket binary JPEG stream)
  - Snap (single frame capture + save)
  - Video recording (N frames to .tif)
  - Time-lapse (frames at an interval, saved as .tif stack)
  - Exposure time control (ms)
  - Binning (1x1, 2x2, 4x4, 8x8)
"""

from modules._api import expose, push_event
from modules import pvcam_raw as pvc
from modules.image_levels import estimate_display_levels, smooth_display_levels
import threading
import numpy as np
import base64
import io
import json
import os
import shutil
import subprocess
import time
from collections import OrderedDict
from datetime import datetime
import tifffile

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
_mosaic_cancel = threading.Event()
_MOSAIC_SETTLE_S = 0.2
_live_fps = 0.0
_live_exposure_changed = threading.Event()

_capture_thread = None
_capture_stop = threading.Event()

_save_dir = os.path.join(os.path.dirname(__file__), "..", "captures")
_base_save_dir = _save_dir

_experiment_json_path = None
_experiment_meta = None
_experiment_lock = threading.Lock()

_frame_lock = threading.Lock()
_latest_jpeg = None          # raw JPEG bytes for WebSocket stream
_frame_event = threading.Event()
_raw_frame_condition = threading.Condition()
_latest_raw_frame = None
_latest_raw_frame_seq = 0

# One unsaved full-resolution snap retained while the preview confirmation
# dialog is open. The browser receives only a JPEG preview; saving writes this
# original uint16 frame as TIFF.
_pending_snap_lock = threading.Lock()
_pending_snap_frame = None
_pending_snap_scale = None

# Viewer TIFF sources. Entries are invalidated when either the file size or
# mtime changes. Pixel data is opened lazily so capture listings only read TIFF
# headers, and the selected frame is copied before releasing the lock so an
# evicted memmap cannot be closed while another request is using it.
_TIFF_CACHE_MAX_ENTRIES = 4
_tiff_cache = OrderedDict()
_tiff_cache_lock = threading.RLock()

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
_disp_brightness = 0      # -100 to 100, applied as offset after stretch
_disp_contrast = 0        # -100 to 100, applied as gain after stretch
_disp_levels_initialized = False

_gamma_lut = None          # cached uint8 LUT for current gamma value
_gamma_lut_val = None      # gamma value the LUT was built for
_gamma_lut_lock = threading.Lock()

# Pseudo-color: None = grayscale, or (R, G, B) tuple for tinting
_pseudo_color = None
_USE_GLOBAL_PSEUDO_COLOR = object()
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

# CoolSNAP EZ nominal 6.45 µm sensor pixels.  Values below are the sample-plane
# µm/pixel at 1×1 binning for the microscope's current 1× camera path.  A user
# calibration made with a stage micrometer replaces any nominal value.
_SCALE_CONFIG_KEY = "coolsnap_scale"
_OBJECTIVE_MAGNIFICATIONS = {
    1: 2.0,
    2: 4.0,
    3: 4.0,
    4: 10.0,
    5: 20.0,
    6: 40.0,
}
_CAMERA_PIXEL_PITCH_UM = 6.45


class CamError(Exception):
    """Raised when a camera operation fails."""


# ── Helpers ──────────────────────────────────────────────────────────────────

_LIVE_PREVIEW_DIM = 800


class _TiffCacheEntry:
    def __init__(self, signature, shape, axes):
        self.signature = signature
        self.shape = tuple(shape)
        self.axes = (axes or "").upper()
        self.array = None


class _UnsupportedTiffShape(ValueError):
    def __init__(self, shape):
        self.shape = tuple(shape)
        super().__init__(f"Unsupported TIFF shape: {self.shape}")


def _close_tiff_array(array):
    """Close a cached numpy memmap, if the array owns one."""
    seen = set()
    current = array
    while current is not None and id(current) not in seen:
        seen.add(id(current))
        mmap = getattr(current, "_mmap", None)
        if mmap is not None:
            try:
                mmap.close()
            except Exception:
                pass
            return
        current = getattr(current, "base", None)


def _discard_tiff_entry(entry):
    if entry.array is not None:
        _close_tiff_array(entry.array)
        entry.array = None


def _clear_tiff_cache():
    """Clear cached TIFF metadata and close any mapped files."""
    with _tiff_cache_lock:
        entries = list(_tiff_cache.values())
        _tiff_cache.clear()
        for entry in entries:
            _discard_tiff_entry(entry)


def _tiff_signature(path):
    stat = os.stat(path)
    return stat.st_mtime_ns, stat.st_size


def _read_tiff_metadata(path):
    """Read the primary TIFF series shape/axes without decoding pixels."""
    with tifffile.TiffFile(path) as tif:
        series = tif.series[0]
        return tuple(int(d) for d in series.shape), getattr(series, "axes", "")


def _get_tiff_entry(path, load_array=False):
    """Return an mtime/size-validated cached TIFF entry."""
    cache_path = os.path.normcase(os.path.abspath(path))
    signature = _tiff_signature(cache_path)
    with _tiff_cache_lock:
        entry = _tiff_cache.get(cache_path)
        if entry is not None and entry.signature != signature:
            _tiff_cache.pop(cache_path)
            _discard_tiff_entry(entry)
            entry = None

        if entry is None:
            shape, axes = _read_tiff_metadata(cache_path)
            entry = _TiffCacheEntry(signature, shape, axes)
            _tiff_cache[cache_path] = entry
        else:
            _tiff_cache.move_to_end(cache_path)

        if load_array and entry.array is None:
            try:
                entry.array = tifffile.memmap(cache_path, mode="r")
            except Exception:
                # Compressed, tiled, or otherwise non-mappable TIFF layouts
                # still benefit from avoiding a decode on every frame request.
                entry.array = tifffile.imread(cache_path)
            entry.shape = tuple(int(d) for d in entry.array.shape)

        while len(_tiff_cache) > _TIFF_CACHE_MAX_ENTRIES:
            _, evicted = _tiff_cache.popitem(last=False)
            _discard_tiff_entry(evicted)
        return entry


def _get_tiff_frame(path, frame_idx=0):
    """Return (detached frame, selected index, frame count, full shape)."""
    with _tiff_cache_lock:
        entry = _get_tiff_entry(path, load_array=True)
        array = entry.array
        shape = entry.shape

        # Recorded camera files are either one grayscale frame or a grayscale
        # frame stack. Preserve the existing API contract that every 3-D TIFF
        # is indexed on its first axis; color TIFF support would also require
        # coordinated changes to auto-adjust and export.
        if len(shape) == 2:
            selected_idx = 0
            n_frames = 1
            frame = array
        elif len(shape) == 3:
            n_frames = int(shape[0])
            selected_idx = max(0, min(int(frame_idx), n_frames - 1))
            frame = array[selected_idx]
        else:
            raise _UnsupportedTiffShape(shape)

        # Detach one frame from the source before it can be evicted/closed.
        return np.array(frame, copy=True), selected_idx, n_frames, shape


def _auto_range(frame):
    """Estimate robust display levels without changing the raw frame."""
    return estimate_display_levels(frame, max_samples=50_000)


def _reset_auto_level_state():
    """Make continuous auto-level tracking initialize from the next frame."""
    global _disp_levels_initialized
    _disp_levels_initialized = False


def _get_gamma_lut(gamma):
    """Return a uint8[256] lookup table for the given gamma value."""
    global _gamma_lut, _gamma_lut_val
    with _gamma_lut_lock:
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
    global _disp_vmin, _disp_vmax, _disp_levels_initialized

    h, w = frame.shape
    scale = max(1, max(h, w) // max_dim) if max_dim else 1
    small = frame[::scale, ::scale] if scale > 1 else frame

    if _disp_mode == "auto":
        lo, hi = _auto_range(small)
        if not _disp_levels_initialized:
            _disp_vmin, _disp_vmax = lo, hi
            _disp_levels_initialized = True
        else:
            _disp_vmin, _disp_vmax = smooth_display_levels(
                (_disp_vmin, _disp_vmax), (lo, hi)
            )

    vmin, vmax = _disp_vmin, _disp_vmax
    if vmax <= vmin:
        vmax = vmin + 1

    u8 = np.clip((small.astype(np.float32) - vmin) / (vmax - vmin) * 255,
                  0, 255).astype(np.uint8)

    gamma = _disp_gamma
    if gamma != 1.0:
        u8 = _get_gamma_lut(gamma)[u8]

    b, c = _disp_brightness, _disp_contrast
    if b != 0 or c != 0:
        gain = 1.0 + c / 100.0
        offset = b * 2.55
        u8 = np.clip(u8.astype(np.float32) * gain + offset, 0, 255).astype(np.uint8)

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


def _next_name(prefix, ext=".tif", directory=None):
    """Return the next sequential filename like 'video_1.tif', 'video_2.tif', etc."""
    d = directory or _save_dir
    existing = set()
    if os.path.isdir(d):
        for f in os.listdir(d):
            if f.startswith(prefix + "_") and f.endswith(ext):
                stem = f[len(prefix) + 1 : -len(ext)]
                if stem.isdigit():
                    existing.add(int(stem))
    n = 1
    while n in existing:
        n += 1
    return f"{prefix}_{n}{ext}"


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


def _save_experiment_json():
    """Write the in-memory experiment metadata to experiment.json."""
    if _experiment_json_path is None or _experiment_meta is None:
        return
    try:
        with open(_experiment_json_path, "w") as f:
            json.dump(_experiment_meta, f, indent=2, default=str)
    except Exception:
        pass


def _objective_scale_1x(calibration):
    """Return ``(µm/px, mm/px)`` owned by one objective at 1×1 binning."""
    if not isinstance(calibration, dict):
        return None
    if "mm_per_pixel_1x" in calibration:
        try:
            mm = float(calibration["mm_per_pixel_1x"])
        except (TypeError, ValueError):
            mm = None
        if mm is not None and np.isfinite(mm) and mm > 0:
            return mm * 1000.0, mm
    try:
        um = float(calibration.get("um_per_pixel_1x"))
    except (TypeError, ValueError):
        return None
    if np.isfinite(um) and um > 0:
        return um, um / 1000.0
    return None


def _sync_objective_scale(objective, saved=None):
    """Keep an objective's µm/px and mm/px on the same 1×1 value.

    A saved mm/px wins. Older configs that only stored µm/px are converted.
    Nominal defaults are left alone when the saved record has neither.
    """
    saved = saved if isinstance(saved, dict) else {}
    if "mm_per_pixel_1x" in saved:
        scale = _objective_scale_1x({"mm_per_pixel_1x": saved["mm_per_pixel_1x"]})
    elif "um_per_pixel_1x" in saved:
        scale = _objective_scale_1x({"um_per_pixel_1x": saved["um_per_pixel_1x"]})
    else:
        scale = _objective_scale_1x(objective)
    if scale is None:
        return objective
    um, mm = scale
    objective["um_per_pixel_1x"] = um
    objective["mm_per_pixel_1x"] = mm
    return objective


def _default_scale_config():
    objectives = {}
    for pos, mag in _OBJECTIVE_MAGNIFICATIONS.items():
        um = _CAMERA_PIXEL_PITCH_UM / mag
        objectives[str(pos)] = {
            "magnification": mag,
            "um_per_pixel_1x": um,
            "mm_per_pixel_1x": um / 1000.0,
            "calibrated": False,
        }
    return {
        "camera": "CoolSNAP EZ",
        "pixel_pitch_um": _CAMERA_PIXEL_PITCH_UM,
        "objectives": objectives,
    }


def _get_scale_config():
    """Return persisted per-objective calibration merged with safe defaults."""
    defaults = _default_scale_config()
    try:
        from modules import config as app_config
        saved = app_config.load(_SCALE_CONFIG_KEY)
    except Exception:
        saved = None
    if not isinstance(saved, dict):
        return defaults
    saved_objectives = saved.get("objectives")
    if isinstance(saved_objectives, dict):
        for pos, value in saved_objectives.items():
            if pos in defaults["objectives"] and isinstance(value, dict):
                defaults["objectives"][pos].update(value)
                _sync_objective_scale(defaults["objectives"][pos], value)
    return defaults


def mosaic_step_nm(sensor_pixels, um_per_pixel_1x, overlap):
    """Stage step between adjacent mosaic tiles, in nanometers.

    The sample-plane field of view is ``sensor_pixels * um_per_pixel_1x``.
    Binning cancels out of that product. ``overlap`` is the fraction of the
    field shared with the next tile, so 0 steps by a full field and 0.1
    keeps a 10% overlap.
    """
    pixels = float(sensor_pixels)
    um_per_pixel = float(um_per_pixel_1x)
    shared = float(overlap)
    if not (pixels > 0) or not (um_per_pixel > 0):
        raise ValueError("Sensor size and µm/pixel must be positive")
    if not (0 <= shared < 1):
        raise ValueError("Overlap must be from 0 up to, but not including, 1")
    fov_um = pixels * um_per_pixel
    return int(round(fov_um * (1.0 - shared) * 1000.0))


def mosaic_tile_offsets(cols=3, rows=3):
    """Serpentine ``(col, row)`` offsets centered on the current field.

    ``+col`` is stage +X (jog right) and ``+row`` is stage +Y (jog up).
    A 3×3 starts at the upper-left tile and snakes across each row.
    """
    cols = int(cols)
    rows = int(rows)
    if cols < 1 or rows < 1:
        raise ValueError("Mosaic size must be at least 1×1")
    col0 = -(cols // 2)
    row0 = rows // 2
    tiles = []
    for r in range(rows):
        row = row0 - r
        col_order = range(cols) if r % 2 == 0 else range(cols - 1, -1, -1)
        for c in col_order:
            tiles.append((col0 + c, row))
    return tiles


def _capture_scale_metadata(objective=None, binning=None, z_nm=None):
    """Resolve immutable scale metadata for one acquired frame."""
    if objective is None or z_nm is None:
        try:
            import modules.nikon_ti as ti
            if ti.is_connected():
                if objective is None:
                    objective = _safe(ti.nosepiece_get_position)
                if z_nm is None:
                    z_nm = _safe(ti.z_get_position)
        except Exception:
            pass

    b = int(binning if binning is not None else _binning[0])
    result = {"binning": b}
    if z_nm is not None:
        result["z_nm"] = int(z_nm)
    try:
        pos = int(objective)
    except (TypeError, ValueError):
        return result

    calibration = _get_scale_config()["objectives"].get(str(pos))
    result["objective_position"] = pos
    scale_1x = _objective_scale_1x(calibration)
    if scale_1x is None:
        return result
    base_um, base_mm = scale_1x
    um_per_pixel = base_um * b
    result.update({
        "objective_magnification": float(calibration["magnification"]),
        "um_per_pixel": um_per_pixel,
        "um_per_pixel_1x": base_um,
        "mm_per_pixel": base_mm * b,
        "mm_per_pixel_1x": base_mm,
        "calibration_source": (
            "calibrated" if calibration.get("calibrated") else "nominal"
        ),
    })
    return result


def _make_capture_entry(color=None, colors=None, channels=None, scale=None):
    entry = {
        "timestamp": datetime.now().isoformat(),
        "color": color if color is not None else _pseudo_color_name,
    }
    if colors:
        entry["colors"] = colors
    if channels:
        entry["channels"] = channels
    if scale:
        entry["scale"] = scale
    return entry


def _record_capture(tif_path, color=None, colors=None, channels=None,
                    scale=None, entry=None):
    """Record a capture in experiment.json when an experiment is active."""
    if _experiment_meta is None:
        return
    exp_dir = os.path.dirname(_experiment_json_path)
    rel_key = os.path.relpath(tif_path, exp_dir).replace("\\", "/")
    if entry is None:
        entry = _make_capture_entry(color, colors, channels, scale)
    with _experiment_lock:
        _experiment_meta["captures"][rel_key] = entry
        _save_experiment_json()


def _write_capture_tiff(path, data, color=None, colors=None, channels=None,
                        scale=None):
    """Write pixels plus portable MOS acquisition metadata to one TIFF."""
    entry = _make_capture_entry(color, colors, channels, scale)
    tifffile.imwrite(
        path, data, photometric="minisblack",
        metadata={"mos_control": entry},
    )
    _record_capture(path, entry=entry)


def _read_tiff_capture_entry(tif_path):
    """Read embedded MOS metadata without decoding image pixels."""
    try:
        with tifffile.TiffFile(tif_path) as tif:
            description = tif.pages[0].description
        data = json.loads(description) if description else {}
        entry = data.get("mos_control") if isinstance(data, dict) else None
        return entry if isinstance(entry, dict) else {}
    except Exception:
        return {}


def _experiment_root_for_path(path):
    """Return the experiment folder that contains experiment.json for `path`."""
    d = os.path.dirname(path) if os.path.isfile(path) else path
    basename = os.path.basename(d)
    if basename in ("video", "stack", "timelapse"):
        return os.path.dirname(d)
    return d


def _load_experiment_captures(exp_root):
    exp_json = os.path.join(exp_root, "experiment.json")
    if not os.path.isfile(exp_json):
        return {}
    try:
        with open(exp_json) as f:
            data = json.load(f)
    except Exception:
        return {}
    return data.get("captures", {}) or {}


def _get_capture_entry(tif_path):
    """Load embedded metadata, supplemented by the experiment capture entry."""
    entry = _read_tiff_capture_entry(tif_path)
    exp_root = _experiment_root_for_path(tif_path)
    captures = _load_experiment_captures(exp_root)
    rel_key = os.path.relpath(tif_path, exp_root).replace("\\", "/")
    experiment_entry = captures.get(rel_key)
    if isinstance(experiment_entry, dict):
        entry.update(experiment_entry)
    return entry


def _get_capture_meta(tif_path):
    """Look up (color, colors) for a capture from its experiment.json.

    Walks up from the file to find the experiment root containing
    experiment.json, then looks up the relative path in the captures dict.
    Returns ("none", None) if not found.
    """
    entry = _get_capture_entry(tif_path)
    if not entry:
        return "none", None
    return entry.get("color", "none"), entry.get("colors")


def _stack_channel_identity(ch):
    """Stable (position, marker, color, filter) key for a stack channel."""
    if not isinstance(ch, dict):
        return ("", "", "none", None)
    return (
        str(ch.get("position") or ""),
        str(ch.get("marker_id") or ""),
        str(ch.get("color") or "none"),
        ch.get("filter"),
    )


def _match_stack_channel_index(channels, identity, fallback=0):
    """Find the channel in `channels` that matches a saved identity.

    Position+filter is preferred so stack-over-time follows the same
    field of view even if channel order changes between files.
    """
    if not channels or identity is None:
        return fallback
    pos, marker_id, color, filt = identity
    for i, ch in enumerate(channels):
        if not isinstance(ch, dict):
            continue
        same_pos = ((pos and ch.get("position") == pos) or
                    (marker_id and ch.get("marker_id") == marker_id))
        if not same_pos:
            continue
        ch_color = str(ch.get("color") or "none")
        if color and ch_color == color:
            return i
        if filt is not None and ch.get("filter") == filt:
            return i
        if not color and filt is None:
            return i
    if pos or marker_id:
        for i, ch in enumerate(channels):
            if not isinstance(ch, dict):
                continue
            if ((pos and ch.get("position") == pos) or
                    (marker_id and ch.get("marker_id") == marker_id)):
                return i
    return fallback


# ── Connection ───────────────────────────────────────────────────────────────

def connect():
    global _hcam, _pvc_initialized, _sensor_w, _sensor_h, _bit_depth, _cam_name_str
    with _lock:
        if _hcam is not None:
            return  # already open — nothing to do
        # Always do a full uninit/init cycle to reset the PVCAM SDK and
        # FireWire interface — this is what Micro-Manager does and it clears
        # stale camera states left over from crashes or forced shutdowns.
        if _pvc_initialized:
            try:
                pvc.uninit()
            except Exception:
                pass
            _pvc_initialized = False
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
    global _hcam, _pvc_initialized, _pending_snap_frame, _pending_snap_scale
    live_stop()
    capture_stop()
    with _pending_snap_lock:
        _pending_snap_frame = None
        _pending_snap_scale = None
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
    if _disp_mode == "auto":
        _reset_auto_level_state()
    _live_exposure_changed.set()


def get_exposure():
    return _exposure_ms


def set_binning(b):
    global _binning
    b = int(b)
    if b not in (1, 2, 4, 8):
        raise CamError(f"Binning must be 1, 2, 4, or 8 (got {b})")
    _binning = (b, b)
    if _disp_mode == "auto":
        _reset_auto_level_state()


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
    scale = _capture_scale_metadata()
    frame = snap()
    _ensure_save_dir()
    path = os.path.join(_save_dir, _next_name("snap"))
    _write_capture_tiff(path, frame, scale=scale)
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
    global _latest_raw_frame, _latest_raw_frame_seq
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

                    with _raw_frame_condition:
                        _latest_raw_frame = frame.copy()
                        _latest_raw_frame_seq += 1
                        _raw_frame_condition.notify_all()
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


def _wait_for_live_frame(after_seq, timeout_s=2.0):
    """Return a raw frame newer than ``after_seq`` from live acquisition."""
    deadline = time.monotonic() + float(timeout_s)
    with _raw_frame_condition:
        while _latest_raw_frame is None or _latest_raw_frame_seq <= after_seq:
            remaining = deadline - time.monotonic()
            if remaining <= 0:
                raise CamError("Timed out waiting for a fresh live frame")
            _raw_frame_condition.wait(timeout=remaining)
        return _latest_raw_frame.copy(), _latest_raw_frame_seq


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


def record_video_and_save(num_frames=100, duration_sec=None, fps=None, color=None):
    """Record video and save to .tif. Returns filepath."""
    scale = _capture_scale_metadata()
    video = record_video(num_frames, duration_sec=duration_sec, fps=fps)
    _ensure_save_dir()
    path = os.path.join(_save_dir, _next_name("video"))
    _write_capture_tiff(path, video, color=color, scale=scale)
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


def timelapse_and_save(num_frames=10, interval_sec=5.0, color=None):
    """Run time-lapse and save. Returns filepath."""
    scale = _capture_scale_metadata()
    stack = timelapse(num_frames, interval_sec)
    _ensure_save_dir()
    path = os.path.join(_save_dir, _next_name("timelapse"))
    _write_capture_tiff(path, stack, color=color, scale=scale)
    return path


# ── Image stack accumulation ─────────────────────────────────────────────────

_stack_frames = []
_stack_scale_metadata = []
_stack_lock = threading.Lock()


_stack_was_live = False

def stack_begin():
    """Clear the frame buffer and stop live preview to avoid conflicts."""
    global _stack_frames, _stack_scale_metadata, _stack_was_live
    _stack_was_live = live_is_active()
    if _stack_was_live:
        live_stop()
    with _stack_lock:
        _stack_frames = []
        _stack_scale_metadata = []


def stack_snap():
    """Snap one frame and append to the stack buffer. Returns frame index."""
    scale = _capture_scale_metadata()
    frame = snap()
    with _stack_lock:
        _stack_frames.append(frame.copy())
        _stack_scale_metadata.append(scale)
        return len(_stack_frames)


def stack_finish():
    """Save all accumulated frames as a single .tif and push event."""
    global _stack_frames, _stack_scale_metadata
    with _stack_lock:
        if not _stack_frames:
            raise CamError("No frames in stack")
        stacked = np.stack(_stack_frames, axis=0)
        channels = [{"scale": scale} for scale in _stack_scale_metadata]
        _stack_frames = []
        _stack_scale_metadata = []
    _ensure_save_dir()
    path = os.path.join(_save_dir, _next_name("stack"))
    _write_capture_tiff(path, stacked, channels=channels)
    fname = os.path.basename(path)
    push_event("onCamStatus", "idle", f"Saved: {fname}")
    push_event("onCamCaptureComplete", fname)
    if _stack_was_live:
        live_start()
    return path


def autofocus(z_range_nm=20000, steps=11, exposure_ms=None, metric="lapvar"):
    """Image-based autofocus: scan Z, score each frame, move to best.

    Scores a fresh live frame, or snaps when live view is inactive, at each of
    `steps` evenly-spaced Z positions across ``current_z ± z_range_nm/2``.
    Moves the stage to the best-scoring Z and returns
    ``(best_z_nm, best_score, scores, z_positions)``.
    """
    import modules.nikon_ti as ti

    if _hcam is None:
        raise CamError("Camera not connected")
    if not ti.is_connected():
        raise CamError("Microscope not connected")
    steps = max(3, int(steps))
    z_range_nm = max(1000, int(z_range_nm))

    if _capture_thread and _capture_thread.is_alive():
        _capture_thread.join(timeout=10)

    was_live = live_is_active()

    # Save & restore exposure
    orig_exposure = _exposure_ms
    if exposure_ms is not None:
        try:
            set_exposure(int(exposure_ms))
        except Exception:
            pass

    try:
        center_z = ti.z_get_position()
        half = z_range_nm // 2
        z_positions = [int(center_z - half + i * (z_range_nm / (steps - 1)))
                       for i in range(steps)]
        scores = []
        best_idx = 0
        best_score = -1.0
        push_event("onAutofocusProgress", 0, steps, "Starting…")

        for i, z in enumerate(z_positions):
            ti.z_move_absolute(z)
            time.sleep(0.15)  # settle
            if was_live:
                # Wait for a frame acquired after the stage has settled. This
                # keeps the live acquisition and browser WebSocket connected.
                with _raw_frame_condition:
                    previous_seq = _latest_raw_frame_seq
                frame, _ = _wait_for_live_frame(previous_seq)
            else:
                frame = snap()
            score = _af_score(frame, metric)
            scores.append(score)
            if score > best_score:
                best_score = score
                best_idx = i
            push_event("onAutofocusProgress", i + 1, steps,
                       f"Z={z/1000:.1f}µm  score={score:.1f}")

        best_z = z_positions[best_idx]
        ti.z_move_absolute(best_z)
        push_event("onAutofocusComplete", best_z, best_score)
        return best_z, best_score, scores, z_positions

    finally:
        if exposure_ms is not None:
            try:
                set_exposure(int(orig_exposure))
            except Exception:
                pass


def _af_score(frame, metric="lapvar"):
    """Sharpness score (higher = sharper)."""
    if metric == "lapvar" and _HAS_CV2:
        return float(cv2.Laplacian(frame.astype(np.float32), cv2.CV_32F).var())
    # Fallback / explicit gradient: use numpy-based gradient variance
    f = frame.astype(np.float32)
    gx = np.diff(f, axis=1)
    gy = np.diff(f, axis=0)
    return float(gx.var() + gy.var())


def _stack_channel_move(ti, ch, last_xy, last_z):
    """Move XY/Z for a stack channel when coordinates are present.

    Returns (moved, last_xy, last_z). Skips axes that already match the
    previous channel so consecutive frames at the same marker stay put.
    """
    moved = False
    x_nm, y_nm = ch.get("x_nm"), ch.get("y_nm")
    z_nm = ch.get("z_nm")
    if x_nm is not None and y_nm is not None:
        xy = (int(x_nm), int(y_nm))
        if xy != last_xy:
            ti.x_move_absolute(xy[0])
            ti.y_move_absolute(xy[1])
            last_xy = xy
            moved = True
    if z_nm is not None:
        z = int(z_nm)
        if z != last_z:
            ti.z_move_absolute(z)
            last_z = z
            moved = True
    return moved, last_xy, last_z


def stack_capture(channels, keep_shutter_open=False):
    """Execute an entire stack acquisition server-side for speed.

    channels: list of dicts, each with:
      objective, filter, lamp_intensity, lamp_on, exposure_ms, binning,
      il_shutter_open, il_nd, shutter_open
      optional x_nm, y_nm, z_nm — stage position for this channel
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
    channel_meta = []
    use_il = il.is_connected()
    try:
        use_scope = ti.is_connected()
        last_xy = None
        last_z = None

        for i, ch in enumerate(channels):
            push_event("onCamStatus", "busy",
                        f"Stack ch {i+1}/{len(channels)}")

            moved = False
            if use_scope:
                moved, last_xy, last_z = _stack_channel_move(
                    ti, ch, last_xy, last_z)
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

            if use_scope or use_il or moved:
                time.sleep(0.8 if moved else 0.5)

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
            channel_meta.append({
                "color": colors[-1],
                "filter": ch.get("filter"),
                "objective": ch.get("objective"),
                "binning": int(ch.get("binning") or _binning[0]),
                "position": ch.get("position") or None,
                "marker_id": ch.get("marker_id") or None,
                "x_nm": ch.get("x_nm"),
                "y_nm": ch.get("y_nm"),
                "z_nm": ch.get("z_nm"),
                "scale": _capture_scale_metadata(
                    objective=ch.get("objective"),
                    binning=ch.get("binning"),
                    z_nm=ch.get("z_nm"),
                ),
            })

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
    path = os.path.join(_save_dir, _next_name("stack"))
    _write_capture_tiff(path, stacked, colors=colors, channels=channel_meta)
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
    color = kwargs.pop('_color', _pseudo_color_name)
    was_live = live_is_active()
    if was_live:
        live_stop()
    fname = None
    try:
        if mode == "video":
            push_event("onCamStatus", "recording",
                        f"Recording {kwargs.get('num_frames', 100)} frames...")
            path = record_video_and_save(color=color, **kwargs)
        elif mode == "timelapse":
            n = kwargs.get('num_frames', 10)
            iv = kwargs.get('interval_sec', 5)
            push_event("onCamStatus", "recording",
                        f"Time-lapse: {n} frames, {iv}s interval...")
            path = timelapse_and_save(color=color, **kwargs)
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
def cam_get_scale_calibration():
    """Return per-objective sample-plane calibration at 1×1 binning."""
    return {"ok": True, **_get_scale_config()}


@expose
def cam_set_scale_calibration(objective_position, um_per_pixel_1x):
    """Persist a stage-micrometer calibration for one turret position."""
    try:
        pos = int(objective_position)
        value = float(um_per_pixel_1x)
    except (TypeError, ValueError):
        return {"error": "Objective position and µm/pixel must be numeric"}
    if str(pos) not in _get_scale_config()["objectives"]:
        return {"error": f"Unknown objective position: {pos}"}
    if not np.isfinite(value) or value <= 0:
        return {"error": "µm/pixel must be greater than zero"}

    config = _get_scale_config()
    objective = config["objectives"][str(pos)]
    objective["um_per_pixel_1x"] = value
    objective["mm_per_pixel_1x"] = value / 1000.0
    objective["calibrated"] = True
    try:
        from modules import config as app_config
        app_config.save(_SCALE_CONFIG_KEY, config)
    except Exception as e:
        return {"error": f"Could not save calibration: {e}"}
    return {"ok": True, **config}


@expose
def cam_snap():
    """Snap and retain the full frame while returning a JPEG preview."""
    global _pending_snap_frame, _pending_snap_scale
    try:
        scale = _capture_scale_metadata()
        frame = snap()
        with _pending_snap_lock:
            _pending_snap_frame = frame.copy()
            _pending_snap_scale = scale
        b64 = _frame_to_base64(frame)
        return {"ok": True, "image": b64,
                "width": int(frame.shape[1]), "height": int(frame.shape[0])}
    except Exception as e:
        return {"error": str(e)}


@expose
def cam_snap_save_pending():
    """Save the full-resolution frame currently shown in the snap dialog."""
    global _pending_snap_frame, _pending_snap_scale
    with _pending_snap_lock:
        if _pending_snap_frame is None:
            return {"error": "No unsaved snap is available"}
        frame = _pending_snap_frame
        scale = _pending_snap_scale
        try:
            _ensure_save_dir()
            path = os.path.join(_save_dir, _next_name("snap"))
            _write_capture_tiff(path, frame, scale=scale)
        except Exception as e:
            return {"error": str(e)}
        _pending_snap_frame = None
        _pending_snap_scale = None
    return {"ok": True, "file": os.path.basename(path)}


@expose
def cam_snap_discard():
    """Discard the full-resolution frame retained for preview."""
    global _pending_snap_frame, _pending_snap_scale
    with _pending_snap_lock:
        _pending_snap_frame = None
        _pending_snap_scale = None
    return {"ok": True}


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
        kw = {"_color": _pseudo_color_name}
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
                     interval_sec=float(interval_sec),
                     _color=_pseudo_color_name)
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
def cam_autofocus(z_range_nm=20000, steps=11, exposure_ms=None, metric="lapvar"):
    """Image-based autofocus.  Z range in nm, steps count.  Returns best Z."""
    try:
        best_z, best_score, scores, z_positions = autofocus(
            z_range_nm=int(z_range_nm),
            steps=int(steps),
            exposure_ms=(int(exposure_ms) if exposure_ms is not None else None),
            metric=str(metric),
        )
        return {"ok": True, "best_z_nm": int(best_z), "best_score": float(best_score),
                "scores": [float(s) for s in scores],
                "z_positions": [int(z) for z in z_positions]}
    except Exception as e:
        return {"error": str(e)}


def mosaic_cancel():
    """Ask a running mosaic to stop after the current tile."""
    _mosaic_cancel.set()


def capture_mosaic(cols=3, rows=3, overlap=0.1):
    """Save a centered grid of fields and return the stage to the start.

    Each tile is one camera field, stepped by the calibrated sample-plane
    scale. ``overlap`` is the fraction of that field shared with the next
    tile. The current XY position is the center tile.
    """
    import modules.nikon_ti as ti

    if _hcam is None:
        raise CamError("Camera not connected")
    if not ti.is_connected():
        raise CamError("Microscope not connected")
    cols, rows = int(cols), int(rows)
    if cols != 3 or rows != 3:
        raise CamError("Only a 3×3 mosaic is supported")

    scale = _capture_scale_metadata()
    um_per_pixel = scale.get("um_per_pixel_1x")
    if not um_per_pixel:
        raise CamError("Current objective has no mm/px scale")
    step_x = mosaic_step_nm(_sensor_w, um_per_pixel, overlap)
    step_y = mosaic_step_nm(_sensor_h, um_per_pixel, overlap)
    if step_x <= 0 or step_y <= 0:
        raise CamError("Mosaic step is too small for this objective")

    origin = ti.xy_get_position()
    origin_x, origin_y = int(origin["x"]), int(origin["y"])
    tiles = mosaic_tile_offsets(cols, rows)
    was_live = live_is_active()
    if was_live:
        live_stop()

    _mosaic_cancel.clear()
    saved = []
    cancelled = False
    try:
        for index, (col, row) in enumerate(tiles, start=1):
            if _mosaic_cancel.is_set():
                cancelled = True
                break
            ti.x_move_absolute(origin_x + col * step_x)
            ti.y_move_absolute(origin_y + row * step_y)
            time.sleep(_MOSAIC_SETTLE_S)
            frame_scale = dict(_capture_scale_metadata())
            frame_scale["mosaic"] = {
                "cols": cols,
                "rows": rows,
                "col": col,
                "row": row,
                "index": index,
                "step_x_nm": step_x,
                "step_y_nm": step_y,
                "origin_x_nm": origin_x,
                "origin_y_nm": origin_y,
            }
            frame = snap()
            _ensure_save_dir()
            path = os.path.join(_save_dir, _next_name("mosaic"))
            _write_capture_tiff(path, frame, scale=frame_scale)
            name = os.path.basename(path)
            saved.append(name)
            push_event("onMosaicProgress", index, len(tiles), name)
            push_event("onCamCaptureComplete", name)
        return {
            "files": saved,
            "count": len(saved),
            "cancelled": cancelled,
            "step_x_nm": step_x,
            "step_y_nm": step_y,
            "origin_x_nm": origin_x,
            "origin_y_nm": origin_y,
        }
    finally:
        try:
            ti.x_move_absolute(origin_x)
            ti.y_move_absolute(origin_y)
        except Exception:
            pass
        if was_live:
            try:
                live_start()
            except Exception:
                pass


@expose
def cam_mosaic_cancel():
    """Stop the mosaic after the tile currently being saved."""
    mosaic_cancel()
    return {"ok": True}


@expose
def cam_capture_mosaic(cols=3, rows=3, overlap=0.1):
    """Capture a 3×3 mosaic centered on the current stage position."""
    try:
        return {"ok": True, **capture_mosaic(cols, rows, overlap)}
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


def _sanitize_filename(name):
    """Replace characters that are illegal in Windows paths."""
    import re
    name = re.sub(r'[<>:"/\\|?*]', '_', name)
    name = name.strip('. ')
    return name or "experiment"


@expose
def cam_experiment_start(name):
    """Create an experiment folder with timelapse/video/stack subfolders."""
    global _save_dir, _experiment_json_path, _experiment_meta
    name = _sanitize_filename(name)
    exp_dir = os.path.join(_base_save_dir, name)
    for sub in ("timelapse", "video", "stack"):
        os.makedirs(os.path.join(exp_dir, sub), exist_ok=True)
    _save_dir = exp_dir
    _experiment_json_path = os.path.join(exp_dir, "experiment.json")
    _experiment_meta = {
        "name": name,
        "saved": datetime.now().isoformat(),
        "payload": {},
        "captures": {},
    }
    _save_experiment_json()
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
    global _save_dir, _experiment_json_path, _experiment_meta
    _save_dir = _base_save_dir
    _experiment_json_path = None
    _experiment_meta = None
    return {"ok": True, "path": _save_dir}


@expose
def cam_experiment_set_payload(payload):
    """Store presets + timeline config into the experiment JSON."""
    if _experiment_meta is None:
        return {"error": "No experiment in progress"}
    with _experiment_lock:
        _experiment_meta["payload"] = payload
        _experiment_meta["saved"] = datetime.now().isoformat()
        _save_experiment_json()
    return {"ok": True}


@expose
def cam_list_captures(subdir=None):
    """List image files (.tif and .npy) in a directory, newest first.

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
    files = [f for f in os.listdir(target) if f.endswith(".tif")]
    files.sort(key=lambda f: os.path.getmtime(os.path.join(target, f)), reverse=True)
    exp_root = _experiment_root_for_path(target)
    captures = _load_experiment_captures(exp_root)
    result = []
    for f in files:
        path = os.path.join(target, f)
        try:
            entry = _get_tiff_entry(path)
            shape = [int(d) for d in entry.shape]
        except Exception:
            shape = []
        rel_key = os.path.relpath(path, exp_root).replace("\\", "/")
        cap = _read_tiff_capture_entry(path)
        exp_cap = captures.get(rel_key)
        if isinstance(exp_cap, dict):
            cap.update(exp_cap)
        item = {"name": f, "shape": shape}
        if cap.get("colors"):
            item["colors"] = cap["colors"]
        if cap.get("channels"):
            item["channels"] = cap["channels"]
        if cap.get("scale"):
            item["scale"] = cap["scale"]
        result.append(item)
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
                    counts[s] = len([f for f in os.listdir(sd) if f.endswith(".tif")])
                experiments.append({"name": name, "types": subs, "counts": counts})
    return {"experiments": experiments}


def _apply_bcg(frame, brightness=0, contrast=100, gamma=1.0, max_dim=1200,
               display_min=None, display_max=None,
               pseudo_color=_USE_GLOBAL_PSEUDO_COLOR):
    """Apply brightness/contrast/gamma adjustments to a uint16 frame → uint8.

    At default settings (b=0, c=100, g=1.0) maps the full data range to 0-255.
    contrast > 100 tightens the window (brighter), < 100 widens it (darker).
    brightness shifts the window center.
    gamma: 0.1-5.0 (1.0 = linear). Applied via LUT.
    display_min/display_max: if both provided, lock the display range instead
    of auto-ranging per frame (prevents brightness flicker during playback).
    """
    h, w = frame.shape
    scale = max(1, max(h, w) // max_dim) if max_dim else 1
    small = frame[::scale, ::scale] if scale > 1 else frame

    if display_min is not None and display_max is not None:
        lo, hi = float(display_min), float(display_max)
    else:
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

    pc = (_pseudo_color if pseudo_color is _USE_GLOBAL_PSEUDO_COLOR
          else pseudo_color)
    if pc is not None and u8.ndim == 2:
        r, g, b = pc
        rgb = np.zeros((*u8.shape, 3), dtype=np.uint8)
        if r: rgb[:, :, 0] = (u8.astype(np.uint16) * r // 255).astype(np.uint8)
        if g: rgb[:, :, 1] = (u8.astype(np.uint16) * g // 255).astype(np.uint8)
        if b: rgb[:, :, 2] = (u8.astype(np.uint16) * b // 255).astype(np.uint8)
        return rgb, lo, hi

    return u8, lo, hi


@expose
def cam_npy_auto_adjust(filename, frame_idx=0, subdir=None):
    """Analyze a frame and return an explicit robust display window."""
    base = os.path.normpath(_base_save_dir) if subdir else os.path.normpath(_save_dir)
    if subdir:
        base = os.path.join(base, subdir)
    path = os.path.join(base, filename)
    if not os.path.isfile(path):
        return {"error": f"File not found: {filename}"}
    try:
        arr = tifffile.imread(path)
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
    auto_lo, auto_hi = _auto_range(small)
    return {
        "ok": True,
        "display_min": auto_lo,
        "display_max": auto_hi,
        "brightness": 0,
        "contrast": 100,
        "gamma": 100,
    }


@expose
def cam_npy_histogram(filename, frame_idx=0, bins=256, subdir=None):
    """Return histogram data for a frame in an image file."""
    base = os.path.normpath(_base_save_dir) if subdir else os.path.normpath(_save_dir)
    if subdir:
        base = os.path.join(base, subdir)
    path = os.path.join(base, filename)
    if not os.path.isfile(path):
        return {"error": f"File not found: {filename}"}
    frame_idx = int(frame_idx)
    try:
        frame, _, _, _ = _get_tiff_frame(path, frame_idx)
    except _UnsupportedTiffShape as e:
        return {"error": f"Unsupported shape: {e.shape}"}
    except Exception as e:
        return {"error": f"Cannot load: {e}"}

    bins = int(bins)
    flat = frame.ravel()
    sub = flat[::max(1, flat.size // 100000)]
    data_min, data_max = int(sub.min()), int(sub.max())
    if data_max <= data_min:
        data_max = data_min + 1
    hist, _ = np.histogram(sub, bins=bins, range=(data_min, data_max))
    return {"hist": hist.tolist(), "min": data_min, "max": data_max}


@expose
def cam_npy_stack(filename, mode="max", brightness=0, contrast=100, gamma=1.0,
                  subdir=None, display_min=None, display_max=None):
    """Return a composite of all frames in an image file as a base64 JPEG.

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
        arr = tifffile.imread(path)
    except Exception as e:
        return {"error": f"Cannot load: {e}"}

    if arr.ndim != 3 or arr.shape[0] < 2:
        return {"error": "Need a multi-frame file to stack"}

    meta_color, meta_colors = _get_capture_meta(path)

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
        u8, _, _ = _apply_bcg(
            composite, brightness, contrast, gamma, max_dim=1200,
            display_min=display_min, display_max=display_max,
        )
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
def cam_npy_merge(filename, weights=None, gamma=1.0, subdir=None, ranges=None):
    """Merge stack channels into an RGB composite with per-channel intensity.

    weights: list of floats (0.0–1.0), one per channel. Controls intensity.
             If None, all channels at 1.0.
    gamma: global gamma correction.
    ranges: list of [lo, hi] per channel to lock normalization. If None,
            robustly auto-ranges each channel.
    """
    base = os.path.normpath(_base_save_dir) if subdir else os.path.normpath(_save_dir)
    if subdir:
        base = os.path.join(base, subdir)
    path = os.path.join(base, filename)
    if not os.path.isfile(path):
        return {"error": f"File not found: {filename}"}
    try:
        arr = tifffile.imread(path)
    except Exception as e:
        return {"error": f"Cannot load: {e}"}
    if arr.ndim != 3 or arr.shape[0] < 2:
        return {"error": "Need a multi-channel stack"}

    _, meta_colors = _get_capture_meta(path)

    gamma = float(gamma)
    n_ch, h, w = arr.shape
    if weights is None:
        weights = [1.0] * n_ch
    else:
        weights = [float(x) for x in weights]
        while len(weights) < n_ch:
            weights.append(1.0)

    computed_ranges = []
    canvas = np.zeros((h, w, 3), dtype=np.float64)
    for i in range(n_ch):
        frame = arr[i].astype(np.float64)
        if ranges and i < len(ranges) and ranges[i] is not None:
            lo, hi = float(ranges[i][0]), float(ranges[i][1])
        else:
            lo, hi = _auto_range(frame)
        if hi <= lo:
            hi = lo + 1
        computed_ranges.append([lo, hi])
        if weights[i] <= 0:
            continue
        norm = np.clip((frame - lo) / (hi - lo), 0, 1)
        c_name = (meta_colors[i] if meta_colors and i < len(meta_colors)
                  else "none")
        rgb = _PSEUDO_COLOR_MAP.get(c_name)
        if rgb:
            for ch in range(3):
                canvas[:, :, ch] += norm * (rgb[ch] / 255.0) * weights[i]
        else:
            for ch in range(3):
                canvas[:, :, ch] += norm * weights[i]

    if gamma != 1.0:
        canvas = np.power(np.clip(canvas, 0, 1), gamma)
    else:
        canvas = np.clip(canvas, 0, 1)

    u8 = (canvas * 255).astype(np.uint8)

    scale = max(1, max(h, w) // 1200)
    if scale > 1:
        u8 = u8[::scale, ::scale]

    if _HAS_CV2:
        bgr = cv2.cvtColor(u8, cv2.COLOR_RGB2BGR)
        ok, buf = cv2.imencode(".jpg", bgr, [cv2.IMWRITE_JPEG_QUALITY, 90])
        b64 = base64.b64encode(buf.tobytes()).decode("ascii") if ok else ""
    elif _HAS_PIL:
        img = _PILImage.fromarray(u8, mode="RGB")
        buf = io.BytesIO()
        img.save(buf, format="JPEG", quality=90)
        b64 = base64.b64encode(buf.getvalue()).decode("ascii")
    else:
        b64 = ""

    colors_out = []
    for i in range(n_ch):
        colors_out.append(meta_colors[i] if meta_colors and i < len(meta_colors)
                          else "none")

    return {
        "image": b64,
        "width": int(u8.shape[1]), "height": int(u8.shape[0]),
        "n_frames": n_ch,
        "colors": colors_out,
        "ranges": computed_ranges,
    }


@expose
def cam_npy_preview(filename, frame_idx=0, brightness=0, contrast=100, gamma=1.0,
                    subdir=None, display_min=None, display_max=None):
    """Return a base64 JPEG preview of an image file.

    For 3D arrays (video/timelapse), frame_idx selects which frame.
    Reads pseudo-color from experiment.json if available.
    brightness/contrast/gamma allow viewer-side image adjustments.
    display_min/display_max: lock display range to prevent brightness flicker.
    subdir: optional relative path under _base_save_dir.
    """
    base = os.path.normpath(_base_save_dir) if subdir else os.path.normpath(_save_dir)
    if subdir:
        base = os.path.join(base, subdir)
    path = os.path.join(base, filename)
    if not os.path.isfile(path):
        return {"error": f"File not found: {filename}"}
    frame_idx = int(frame_idx)
    try:
        frame, frame_idx, n_frames, shape = _get_tiff_frame(path, frame_idx)
    except _UnsupportedTiffShape as e:
        return {"error": f"Unsupported array shape: {e.shape}"}
    except Exception as e:
        return {"error": f"Cannot load: {e}"}

    brightness = float(brightness)
    contrast = float(contrast)
    gamma = float(gamma)

    meta_entry = _get_capture_entry(path)
    meta_color = meta_entry.get("color", "none") if meta_entry else "none"
    meta_colors = meta_entry.get("colors") if meta_entry else None
    meta_channels = meta_entry.get("channels") if meta_entry else None

    if meta_colors and frame_idx < len(meta_colors):
        frame_color = meta_colors[frame_idx]
    else:
        frame_color = meta_color
    frame_scale = meta_entry.get("scale") if meta_entry else None
    if meta_channels and frame_idx < len(meta_channels):
        channel = meta_channels[frame_idx]
        if isinstance(channel, dict) and isinstance(channel.get("scale"), dict):
            frame_scale = channel["scale"]

    dmin = float(display_min) if display_min is not None else None
    dmax = float(display_max) if display_max is not None else None
    u8, used_lo, used_hi = _apply_bcg(
        frame, brightness, contrast, gamma, max_dim=1200,
        display_min=dmin, display_max=dmax,
        pseudo_color=_PSEUDO_COLOR_MAP.get(frame_color),
    )
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
    result = {
        "image": b64,
        "width": int(frame.shape[1]),
        "height": int(frame.shape[0]),
        "n_frames": n_frames,
        "frame_idx": frame_idx,
        "shape": [int(d) for d in shape],
        "color": frame_color,
        "display_min": used_lo,
        "display_max": used_hi,
    }
    if meta_colors:
        result["colors"] = meta_colors
    if meta_channels:
        result["channels"] = meta_channels
    if frame_scale:
        result["scale"] = frame_scale
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
            _reset_auto_level_state()
    return {"mode": _disp_mode, "vmin": round(_disp_vmin), "vmax": round(_disp_vmax),
            "gamma": _disp_gamma}


@expose
def cam_auto_adjust():
    """One-shot: compute display range from current frame and lock it."""
    global _disp_mode, _disp_vmin, _disp_vmax, _disp_levels_initialized
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
    _disp_levels_initialized = True
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
    global _disp_mode, _disp_vmin, _disp_vmax, _disp_levels_initialized
    _disp_vmin = float(vmin)
    _disp_vmax = float(vmax)
    _disp_levels_initialized = True
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
def cam_set_brightness(val):
    global _disp_brightness
    _disp_brightness = max(-100, min(100, int(val)))
    return {"brightness": _disp_brightness}


@expose
def cam_set_contrast(val):
    global _disp_contrast
    _disp_contrast = max(-100, min(100, int(val)))
    return {"contrast": _disp_contrast}


@expose
def cam_reset_display():
    """Reset brightness, contrast, and gamma to defaults."""
    global _disp_brightness, _disp_contrast, _disp_gamma, _disp_mode, _disp_vmin, _disp_vmax
    _disp_brightness = 0
    _disp_contrast = 0
    _disp_gamma = 1.0
    _disp_mode = "auto"
    _disp_vmin = 0.0
    _disp_vmax = 65535.0
    _reset_auto_level_state()
    return {"brightness": 0, "contrast": 0, "gamma": 1.0}


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


# ── Export (image + video) ───────────────────────────────────────────────────

_EXPORTS_DIR = os.path.join(os.path.dirname(__file__), "..", "exports")


def _ensure_exports_dir():
    os.makedirs(_EXPORTS_DIR, exist_ok=True)


def _find_ffmpeg():
    """Locate FFmpeg, including its common WinGet installation directory."""
    path = shutil.which("ffmpeg")
    if path:
        return path
    if os.name == "nt":
        winget = os.path.expandvars(
            r"%LOCALAPPDATA%\Microsoft\WinGet\Packages"
        )
        if os.path.isdir(winget):
            for root, _dirs, files in os.walk(winget):
                if "ffmpeg.exe" in files:
                    return os.path.join(root, "ffmpeg.exe")
    return None


_FFMPEG = _find_ffmpeg()


def _encode_h264_mp4(out_path, fps, frame_w, frame_h, frames):
    """Encode RGB frames as macOS-compatible H.264/yuv420p MP4."""
    if not _FFMPEG:
        raise RuntimeError(
            "FFmpeg is required for MP4 export but was not found"
        )

    fps = max(1, int(fps))
    frame_w, frame_h = int(frame_w), int(frame_h)
    cmd = [
        _FFMPEG,
        "-y",
        "-hide_banner",
        "-loglevel", "error",
        "-f", "rawvideo",
        "-pix_fmt", "rgb24",
        "-s:v", f"{frame_w}x{frame_h}",
        "-r", str(fps),
        "-i", "pipe:0",
        "-an",
        "-c:v", "libx264",
        "-preset", "medium",
        "-crf", "18",
        "-vf", "pad=ceil(iw/2)*2:ceil(ih/2)*2",
        "-pix_fmt", "yuv420p",
        "-color_range", "tv",
        "-colorspace", "bt709",
        "-color_primaries", "bt709",
        "-color_trc", "bt709",
        "-movflags", "+faststart",
        out_path,
    ]
    proc = subprocess.Popen(
        cmd,
        stdin=subprocess.PIPE,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.PIPE,
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
    )
    written = 0
    write_error = None
    try:
        for frame in frames:
            rgb = np.asarray(frame, dtype=np.uint8)
            if rgb.shape != (frame_h, frame_w, 3):
                raise ValueError(
                    f"MP4 frame has shape {rgb.shape}; expected "
                    f"{(frame_h, frame_w, 3)}"
                )
            proc.stdin.write(np.ascontiguousarray(rgb).tobytes())
            written += 1
    except Exception as exc:
        write_error = exc
    finally:
        if proc.stdin:
            try:
                proc.stdin.close()
            except Exception:
                pass

    stderr = proc.stderr.read().decode(errors="replace") if proc.stderr else ""
    return_code = proc.wait()
    if write_error is not None or return_code != 0 or written == 0:
        try:
            os.remove(out_path)
        except OSError:
            pass
        detail = str(write_error) if write_error is not None else stderr.strip()
        raise RuntimeError(detail or "FFmpeg did not encode any frames")
    return written


def _render_frame_rgb(frame, color_name, brightness=0, contrast=100, gamma=1.0,
                      display_min=None, display_max=None):
    """Render a uint16 frame to an 8-bit RGB numpy array (H, W, 3).

    Returns (rgb_array, lo, hi).
    """
    global _pseudo_color, _pseudo_color_name
    prev_color, prev_name = _pseudo_color, _pseudo_color_name
    _pseudo_color = _PSEUDO_COLOR_MAP.get(color_name)
    _pseudo_color_name = color_name
    try:
        img, lo, hi = _apply_bcg(frame, brightness, contrast, gamma,
                                 max_dim=0, display_min=display_min,
                                 display_max=display_max)
    finally:
        _pseudo_color, _pseudo_color_name = prev_color, prev_name
    if img.ndim == 2:
        img = np.stack([img, img, img], axis=-1)
    return img, lo, hi


def _render_merge_rgb(arr, meta_colors, weights, gamma, ranges=None):
    """Render a multi-channel stack (C, H, W) to merged RGB (H, W, 3).

    Returns (rgb_array, computed_ranges).
    """
    n_ch, h, w = arr.shape
    if weights is None:
        weights = [1.0] * n_ch
    computed = []
    canvas = np.zeros((h, w, 3), dtype=np.float64)
    for i in range(n_ch):
        frame = arr[i].astype(np.float64)
        if ranges and i < len(ranges) and ranges[i] is not None:
            lo, hi = float(ranges[i][0]), float(ranges[i][1])
        else:
            lo, hi = _auto_range(frame)
        if hi <= lo:
            hi = lo + 1
        computed.append([lo, hi])
        if weights[i] <= 0:
            continue
        norm = np.clip((frame - lo) / (hi - lo), 0, 1)
        c_name = (meta_colors[i] if meta_colors and i < len(meta_colors)
                  else "none")
        rgb = _PSEUDO_COLOR_MAP.get(c_name)
        if rgb:
            for ch in range(3):
                canvas[:, :, ch] += norm * (rgb[ch] / 255.0) * weights[i]
        else:
            for ch in range(3):
                canvas[:, :, ch] += norm * weights[i]
    if gamma != 1.0:
        canvas = np.power(np.clip(canvas, 0, 1), gamma)
    else:
        canvas = np.clip(canvas, 0, 1)
    return (canvas * 255).astype(np.uint8), computed


@expose
def cam_export_image(filename, frame_idx=0, brightness=0, contrast=100,
                     gamma=1.0, subdir=None, display_min=None,
                     display_max=None):
    """Export the currently viewed frame as a full-resolution PNG."""
    base = os.path.normpath(_base_save_dir) if subdir else os.path.normpath(_save_dir)
    if subdir:
        base = os.path.join(base, subdir)
    path = os.path.join(base, filename)
    if not os.path.isfile(path):
        return {"error": f"File not found: {filename}"}
    try:
        arr = tifffile.imread(path)
    except Exception as e:
        return {"error": f"Cannot load: {e}"}

    frame_idx = int(frame_idx)
    meta_color, meta_colors = _get_capture_meta(path)

    if arr.ndim == 2:
        frame = arr
    elif arr.ndim == 3:
        frame = arr[min(frame_idx, arr.shape[0] - 1)]
    else:
        return {"error": f"Unsupported shape: {arr.shape}"}

    if meta_colors and frame_idx < len(meta_colors):
        color = meta_colors[frame_idx]
    else:
        color = meta_color

    rgb, _, _ = _render_frame_rgb(
        frame, color, brightness, contrast, gamma,
        display_min=display_min, display_max=display_max,
    )

    _ensure_exports_dir()
    stem = os.path.splitext(filename)[0]
    if arr.ndim == 3:
        out_name = f"{stem}_f{frame_idx}.png"
    else:
        out_name = f"{stem}.png"
    out_path = os.path.join(_EXPORTS_DIR, out_name)

    if _HAS_CV2:
        cv2.imwrite(out_path, cv2.cvtColor(rgb, cv2.COLOR_RGB2BGR))
    elif _HAS_PIL:
        _PILImage.fromarray(rgb, mode="RGB").save(out_path)
    else:
        return {"error": "No image library available (cv2 or PIL required)"}

    return {"ok": True, "filename": out_name}


@expose
def cam_export_merge_image(filename, weights=None, gamma=1.0, subdir=None):
    """Export the current merged stack view as a full-resolution PNG."""
    base = os.path.normpath(_base_save_dir) if subdir else os.path.normpath(_save_dir)
    if subdir:
        base = os.path.join(base, subdir)
    path = os.path.join(base, filename)
    if not os.path.isfile(path):
        return {"error": f"File not found: {filename}"}
    try:
        arr = tifffile.imread(path)
    except Exception as e:
        return {"error": f"Cannot load: {e}"}
    if arr.ndim != 3 or arr.shape[0] < 2:
        return {"error": "Need a multi-channel stack"}

    _, meta_colors = _get_capture_meta(path)
    gamma = float(gamma)
    if weights:
        weights = [float(x) for x in weights]

    rgb, _ = _render_merge_rgb(arr, meta_colors, weights, gamma)

    _ensure_exports_dir()
    stem = os.path.splitext(filename)[0]
    out_name = f"{stem}_merge.png"
    out_path = os.path.join(_EXPORTS_DIR, out_name)

    if _HAS_CV2:
        cv2.imwrite(out_path, cv2.cvtColor(rgb, cv2.COLOR_RGB2BGR))
    elif _HAS_PIL:
        _PILImage.fromarray(rgb, mode="RGB").save(out_path)
    else:
        return {"error": "No image library available"}

    return {"ok": True, "filename": out_name}


@expose
def cam_export_mp4(filename, fps=10, brightness=0, contrast=100, gamma=1.0,
                   subdir=None, display_min=None, display_max=None):
    """Export a video or timelapse .tif as an MP4 with display adjustments."""
    base = os.path.normpath(_base_save_dir) if subdir else os.path.normpath(_save_dir)
    if subdir:
        base = os.path.join(base, subdir)
    path = os.path.join(base, filename)
    if not os.path.isfile(path):
        return {"error": f"File not found: {filename}"}
    try:
        arr = tifffile.imread(path)
    except Exception as e:
        return {"error": f"Cannot load: {e}"}
    if arr.ndim != 3 or arr.shape[0] < 2:
        return {"error": "Need a multi-frame file (video/timelapse)"}

    meta_color, _ = _get_capture_meta(path)
    n_frames, h, w = arr.shape
    fps = int(fps)

    first_rgb, lo, hi = _render_frame_rgb(
        arr[0], meta_color, brightness, contrast, gamma,
        display_min=display_min, display_max=display_max,
    )
    frame_h, frame_w = first_rgb.shape[:2]

    _ensure_exports_dir()
    stem = os.path.splitext(filename)[0]
    out_name = f"{stem}.mp4"
    out_path = os.path.join(_EXPORTS_DIR, out_name)

    try:
        def rendered_frames():
            yield first_rgb
            for i in range(1, n_frames):
                rgb, _, _ = _render_frame_rgb(
                    arr[i], meta_color, brightness, contrast, gamma,
                    display_min=lo, display_max=hi,
                )
                yield rgb

        written = _encode_h264_mp4(
            out_path, fps, frame_w, frame_h, rendered_frames()
        )
    except Exception as e:
        return {"error": f"MP4 export failed: {e}"}

    return {
        "ok": True,
        "filename": out_name,
        "frames": written,
        "codec": "H.264",
        "pixel_format": "yuv420p",
    }


@expose
def cam_export_stack_tl_mp4(filenames, channel=0, fps=2, brightness=0,
                            contrast=100, gamma=1.0, subdir=None,
                            display_min=None, display_max=None):
    """Export a stack timelapse (single channel across files) as MP4."""
    if not filenames or len(filenames) < 2:
        return {"error": "Need at least 2 stack files"}

    base = os.path.normpath(_base_save_dir) if subdir else os.path.normpath(_save_dir)
    if subdir:
        base = os.path.join(base, subdir)
    fps = int(fps)
    channel = int(channel)

    first_path = os.path.join(base, filenames[0])
    if not os.path.isfile(first_path):
        return {"error": f"File not found: {filenames[0]}"}
    first_arr = tifffile.imread(first_path)
    if first_arr.ndim != 3:
        return {"error": "Files must be multi-channel stacks"}
    first_entry = _get_capture_entry(first_path)
    meta_colors = first_entry.get("colors")
    first_channels = first_entry.get("channels") or []
    identity = None
    if channel < len(first_channels):
        identity = _stack_channel_identity(first_channels[channel])
    color = (meta_colors[channel] if meta_colors and channel < len(meta_colors)
             else "none")
    if identity and identity[2]:
        color = identity[2]
    frame = first_arr[min(channel, first_arr.shape[0] - 1)]
    first_rgb, lo, hi = _render_frame_rgb(
        frame, color, brightness, contrast, gamma,
        display_min=display_min, display_max=display_max,
    )
    frame_h, frame_w = first_rgb.shape[:2]

    _ensure_exports_dir()
    out_name = _next_name(f"stack_tl_ch{channel}", ext=".mp4", directory=_EXPORTS_DIR)
    out_path = os.path.join(_EXPORTS_DIR, out_name)

    try:
        def rendered_frames():
            yield first_rgb
            for fn in filenames[1:]:
                p = os.path.join(base, fn)
                if not os.path.isfile(p):
                    continue
                arr = tifffile.imread(p)
                if arr.ndim != 3:
                    continue
                entry = _get_capture_entry(p)
                mc = entry.get("colors")
                chs = entry.get("channels") or []
                ch_idx = _match_stack_channel_index(chs, identity, channel)
                c = (mc[ch_idx] if mc and ch_idx < len(mc) else "none")
                f = arr[min(ch_idx, arr.shape[0] - 1)]
                rgb, _, _ = _render_frame_rgb(
                    f, c, brightness, contrast, gamma,
                    display_min=lo, display_max=hi,
                )
                yield rgb

        written = _encode_h264_mp4(
            out_path, fps, frame_w, frame_h, rendered_frames()
        )
    except Exception as e:
        return {"error": f"MP4 export failed: {e}"}

    return {
        "ok": True,
        "filename": out_name,
        "frames": written,
        "codec": "H.264",
        "pixel_format": "yuv420p",
    }


@expose
def cam_export_merge_mp4(filenames, weights=None, merge_gamma=1.0, fps=2,
                         subdir=None):
    """Export a merged stack timelapse as MP4."""
    if not filenames or len(filenames) < 2:
        return {"error": "Need at least 2 stack files"}

    base = os.path.normpath(_base_save_dir) if subdir else os.path.normpath(_save_dir)
    if subdir:
        base = os.path.join(base, subdir)
    fps = int(fps)
    merge_gamma = float(merge_gamma)
    if weights:
        weights = [float(x) for x in weights]

    first_path = os.path.join(base, filenames[0])
    if not os.path.isfile(first_path):
        return {"error": f"File not found: {filenames[0]}"}
    first_arr = tifffile.imread(first_path)
    if first_arr.ndim != 3:
        return {"error": "Files must be multi-channel stacks"}
    _, meta_colors = _get_capture_meta(first_path)

    first_rgb, locked_ranges = _render_merge_rgb(first_arr, meta_colors,
                                                 weights, merge_gamma)
    frame_h, frame_w = first_rgb.shape[:2]

    _ensure_exports_dir()
    out_name = _next_name("stack_tl_merge", ext=".mp4", directory=_EXPORTS_DIR)
    out_path = os.path.join(_EXPORTS_DIR, out_name)

    try:
        def rendered_frames():
            yield first_rgb
            for fn in filenames[1:]:
                p = os.path.join(base, fn)
                if not os.path.isfile(p):
                    continue
                arr = tifffile.imread(p)
                if arr.ndim != 3:
                    continue
                _, mc = _get_capture_meta(p)
                rgb, _ = _render_merge_rgb(
                    arr, mc or meta_colors, weights, merge_gamma,
                    ranges=locked_ranges,
                )
                yield rgb

        written = _encode_h264_mp4(
            out_path, fps, frame_w, frame_h, rendered_frames()
        )
    except Exception as e:
        return {"error": f"MP4 export failed: {e}"}

    return {
        "ok": True,
        "filename": out_name,
        "frames": written,
        "codec": "H.264",
        "pixel_format": "yuv420p",
    }
