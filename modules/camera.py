"""
Camera control module — wraps Pycromanager / video_capture.py and exposes
camera operations to the Eel frontend.

Micro-Manager must be running with the server enabled (port 4827) for
camera functions to work.  If pycromanager is not installed, the module
loads but all functions return an appropriate error.
"""

import eel
import threading
import numpy as np
import base64
import io
import os
import sys
import time
from datetime import datetime

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

_core = None
_capture_thread = None
_capture_stop = threading.Event()
_save_dir = os.path.join(os.path.dirname(__file__), "..", "captured_videos")
_live_timer = None
_live_active = False

_HAS_PYCROMANAGER = True
try:
    from pycromanager import Core as _PycroCore
except ImportError:
    _HAS_PYCROMANAGER = False

_HAS_PIL = True
try:
    from PIL import Image as _PILImage
except ImportError:
    _HAS_PIL = False


def get_core():
    """Return the pycromanager Core (used by experiment engine)."""
    return _core


def is_connected():
    return _core is not None


def _image_to_base64(img):
    """Convert a 2-D numpy array to a base64-encoded PNG string."""
    vmin, vmax = np.nanpercentile(img, [1, 99])
    clipped = np.clip(img.astype(float), vmin, vmax)
    if vmax > vmin:
        normed = ((clipped - vmin) / (vmax - vmin) * 255).astype(np.uint8)
    else:
        normed = np.zeros_like(img, dtype=np.uint8)

    if _HAS_PIL:
        pil_img = _PILImage.fromarray(normed, mode="L")
        buf = io.BytesIO()
        pil_img.save(buf, format="PNG")
        return base64.b64encode(buf.getvalue()).decode("ascii")

    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt

    fig, ax = plt.subplots(figsize=(4, 3), dpi=100)
    ax.imshow(normed, cmap="gray", vmin=0, vmax=255)
    ax.axis("off")
    buf = io.BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight", pad_inches=0)
    plt.close(fig)
    return base64.b64encode(buf.getvalue()).decode("ascii")


def _snap_image(core):
    """Snap a single image using the pycromanager Core."""
    try:
        if core.is_sequence_running():
            core.stop_sequence_acquisition()
    except Exception:
        pass
    core.snap_image()
    tagged = core.get_tagged_image()
    pix = tagged.pix
    if pix.ndim == 1 and "Height" in tagged.tags and "Width" in tagged.tags:
        return np.reshape(pix, (tagged.tags["Height"], tagged.tags["Width"]))
    return pix


def _record_video(core, duration_sec, fps):
    """Record a video and return a 3-D array (frames, H, W)."""
    n_frames = int(duration_sec * fps)
    interval = 1.0 / fps
    frames = []
    for i in range(n_frames):
        frames.append(_snap_image(core))
        if i < n_frames - 1:
            time.sleep(interval)
    return np.stack(frames, axis=0)


# ── Connection ──────────────────────────────────────────────────────────────

@eel.expose
def camera_connect():
    global _core
    if not _HAS_PYCROMANAGER:
        return {"error": "pycromanager is not installed"}
    try:
        _core = _PycroCore()
        return {"ok": True, "msg": "Connected to Micro-Manager"}
    except Exception as e:
        _core = None
        return {"error": str(e)}


@eel.expose
def camera_disconnect():
    global _core, _live_active
    _live_active = False
    _core = None
    return {"ok": True}


@eel.expose
def camera_is_connected():
    return _core is not None


# ── Preview ─────────────────────────────────────────────────────────────────

@eel.expose
def camera_snap():
    """Snap a single image and return it as base64 PNG for the preview pane."""
    if not _core:
        return {"error": "Camera not connected"}
    try:
        img = _snap_image(_core)
        b64 = _image_to_base64(img)
        return {"ok": True, "image": b64,
                "width": int(img.shape[1]), "height": int(img.shape[0])}
    except Exception as e:
        return {"error": str(e)}


@eel.expose
def camera_set_live(active):
    global _live_active
    _live_active = bool(active)
    return {"ok": True}


@eel.expose
def camera_is_live():
    return _live_active


# ── Capture ─────────────────────────────────────────────────────────────────

@eel.expose
def camera_set_save_dir(path):
    global _save_dir
    _save_dir = path
    os.makedirs(_save_dir, exist_ok=True)
    return {"ok": True, "path": _save_dir}


@eel.expose
def camera_get_save_dir():
    return _save_dir


@eel.expose
def camera_capture_video(duration_sec=10, fps=10):
    """Start a background video capture."""
    global _capture_thread
    if not _core:
        return {"error": "Camera not connected"}
    if _capture_thread and _capture_thread.is_alive():
        return {"error": "Capture already in progress"}
    _capture_stop.clear()
    t = threading.Thread(target=_capture_thread_fn,
                         args=(duration_sec, fps), daemon=True)
    _capture_thread = t
    t.start()
    return {"ok": True, "msg": f"Capturing {duration_sec}s @ {fps} fps"}


def capture_video_blocking(duration_sec=10, fps=10):
    """Synchronous video capture. Returns file path."""
    if not _core:
        raise RuntimeError("Camera not connected")
    video = _record_video(_core, duration_sec, fps)
    os.makedirs(_save_dir, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    path = os.path.join(_save_dir, f"video_{ts}.npy")
    np.save(path, video)
    return path


def snap_image_blocking():
    """Synchronous single-image capture. Returns file path."""
    if not _core:
        raise RuntimeError("Camera not connected")
    img = _snap_image(_core)
    os.makedirs(_save_dir, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S_%f")[:-3]
    path = os.path.join(_save_dir, f"image_{ts}.npy")
    np.save(path, img)
    return path



def _capture_thread_fn(duration_sec, fps):
    try:
        eel.onCameraStatus("capturing",
                           f"Recording {duration_sec}s video ...")()
        path = capture_video_blocking(duration_sec, fps)
        eel.onCameraStatus("idle", f"Saved: {os.path.basename(path)}")()
        eel.onCameraCaptureComplete(os.path.basename(path))()
    except Exception as e:
        eel.onCameraStatus("error", str(e))()


# ── Camera Protocol ─────────────────────────────────────────────────────────

_proto_thread = None
_proto_stop = threading.Event()


def _hms_to_seconds(hms):
    if not hms or hms == "00:00:00":
        return 0
    parts = hms.split(":")
    return int(parts[0]) * 3600 + int(parts[1]) * 60 + int(parts[2])


@eel.expose
def camera_run_protocol(steps):
    """Run a camera protocol — a sequence of Capture, Image, and Pause steps.

    steps: list of dicts, each with:
        action      : "Capture" | "Image" | "Pause"
        duration_sec: float   (Capture only)
        fps         : int     (Capture only)
        time        : "HH:MM:SS" (Pause only — how long to wait)
    """
    global _proto_thread
    if not _core:
        return {"error": "Camera not connected"}
    if _proto_thread and _proto_thread.is_alive():
        return {"error": "Camera protocol already running"}
    _proto_stop.clear()
    t = threading.Thread(target=_run_camera_protocol, args=(steps,), daemon=True)
    _proto_thread = t
    t.start()
    return {"ok": True}


@eel.expose
def camera_stop_protocol():
    _proto_stop.set()
    if _proto_thread and _proto_thread.is_alive():
        _proto_thread.join(timeout=3)
    return {"ok": True}


def _run_camera_protocol(steps):
    try:
        n = len(steps)
        eel.onCameraProtocolUpdate(-1, "started",
                                   f"Camera protocol started ({n} steps)")()

        for i, step in enumerate(steps):
            if _proto_stop.is_set():
                break

            action = step.get("action", "")

            if action == "Capture":
                dur = float(step.get("duration_sec", 10))
                fps = int(step.get("fps", 10))
                eel.onCameraProtocolUpdate(
                    i, "running",
                    f"Step {i+1}: Capturing {dur}s @{fps}fps")()
                path = capture_video_blocking(dur, fps)
                eel.onCameraCaptureComplete(os.path.basename(path))()
                eel.onCameraProtocolUpdate(
                    i, "done",
                    f"Step {i+1}: Saved {os.path.basename(path)}")()

            elif action == "Image":
                eel.onCameraProtocolUpdate(
                    i, "running", f"Step {i+1}: Snapping image")()
                path = snap_image_blocking()
                eel.onCameraCaptureComplete(os.path.basename(path))()
                eel.onCameraProtocolUpdate(
                    i, "done",
                    f"Step {i+1}: Saved {os.path.basename(path)}")()

            elif action == "Pause":
                wait = _hms_to_seconds(step.get("time", "00:00:00"))
                if wait > 0:
                    eel.onCameraProtocolUpdate(
                        i, "running", f"Step {i+1}: Pause")()
                    end = time.monotonic() + wait
                    while time.monotonic() < end:
                        if _proto_stop.is_set():
                            break
                        remaining = end - time.monotonic()
                        rm = int(remaining // 60)
                        rs = int(remaining % 60)
                        eel.onCameraProtocolCountdown(
                            i, f"{rm:02d}:{rs:02d}")()
                        time.sleep(1)

        status = "aborted" if _proto_stop.is_set() else "complete"
        eel.onCameraProtocolUpdate(-1, status,
                                   f"Camera protocol {status}")()
    except Exception as e:
        eel.onCameraProtocolUpdate(-1, "error",
                                   f"Camera protocol error: {e}")()
