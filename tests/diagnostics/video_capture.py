"""
Pycromanager video capture: get_image, get_video, capture_videos.
Start Micro-Manager first, then: Tools > Options > check "Run server on port 4827".
"""

from pycromanager import Core
from datetime import datetime
import numpy as np
import os
import time


def _stop_live_if_running(core):
    """Stop Live or sequence acquisition so snap_image() can run. No-op if nothing running."""
    try:
        if core.is_sequence_running():
            core.stop_sequence_acquisition()
    except Exception:
        pass


def get_image(core):
    """Snap and return image as 2D array."""
    _stop_live_if_running(core)
    core.snap_image()
    tagged = core.get_tagged_image()
    pix = tagged.pix
    if pix.ndim == 1 and "Height" in tagged.tags and "Width" in tagged.tags:
        return np.reshape(pix, (tagged.tags["Height"], tagged.tags["Width"]))
    return pix


def get_video(core, duration_sec=10, fps=10):
    """Capture a video as a 3D array (n_frames, height, width)."""
    n_frames = int(duration_sec * fps)
    interval_sec = 1.0 / fps
    frames = []
    for i in range(n_frames):
        frames.append(get_image(core))
        if i < n_frames - 1:
            time.sleep(interval_sec)
    return np.stack(frames, axis=0)

def capture_videos(core, interval_sec, total_time_sec, duration_sec, fps=10, save_dir=None):
    """
    Capture videos at a fixed interval for a total time.

    Args:
        core: Pycromanager Core instance.
        interval_sec: Time between the start of each video (seconds).
        total_time_sec: How long to run (wall clock, seconds).
        duration_sec: Length of each video (seconds).
        fps: Frames per second for each video.
        save_dir: If set, save each video to this directory as it's captured
                  (e.g. video_001_143022.npy). Returns list of paths instead of arrays.

    Returns:
        If save_dir is None: list of 3D arrays, each shape (n_frames, height, width).
        If save_dir is set: list of saved file paths.
    """
    if save_dir:
        os.makedirs(save_dir, exist_ok=True)

    videos = [] if save_dir is None else None
    paths = [] if save_dir else None
    t_start = time.time()
    n = 0

    while time.time() - t_start < total_time_sec:
        video = get_video(core, duration_sec=duration_sec, fps=fps)
        if save_dir:
            ts = datetime.now().strftime("%H%M%S")
            path = os.path.join(save_dir, f"video_{n + 1:03d}_{ts}.npy")
            np.save(path, video)
            paths.append(path)
            print(f"Saved {path}")
        else:
            videos.append(video)
        n += 1
        # Wait until next scheduled start (interval_sec from start of run)
        elapsed = time.time() - t_start
        next_start = n * interval_sec
        wait = next_start - elapsed
        if wait > 0:
            time.sleep(wait)

    return paths if save_dir else videos


if __name__ == "__main__":
    # Example: 10 s videos every 5 min for 1 hour, saved to disk as they come in
    SAVE_DIR = r"C:\Users\penelope\Desktop\MOS-11\captured_videos"
    core = Core()
    paths = capture_videos(
        core,
        interval_sec=300,
        total_time_sec=3600,
        duration_sec=10,
        fps=10,
        save_dir=SAVE_DIR,
    )
    print(f"Done. Saved {len(paths)} videos to {SAVE_DIR}")
