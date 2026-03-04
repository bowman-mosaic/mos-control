"""
View a video saved as .npy (shape: n_frames, height, width).
Usage: python view_npy_video.py <path/to/video.npy>
"""

import sys
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.widgets import Button
from matplotlib.animation import FuncAnimation


def watch(filepath, fps=10):
    """Open an .npy video and show it with a Play/Pause button."""
    data = np.load(filepath)
    if data.ndim != 3:
        print(f"Expected 3D (frames, height, width), got {data.shape}")
        return
    n_frames = data.shape[0]

    fig, ax = plt.subplots()
    plt.subplots_adjust(bottom=0.1)
    vmin, vmax = np.nanpercentile(data, [1, 99])
    im = ax.imshow(data[0], cmap="gray", vmin=vmin, vmax=vmax)
    ax.axis("off")

    interval_ms = int(1000 / fps)

    def update_frame(i):
        im.set_data(data[i])
        return (im,)

    anim = FuncAnimation(
        fig, update_frame, frames=n_frames, interval=interval_ms, repeat=True, blit=True
    )
    anim.pause()
    playing = [False]  # use list so we can mutate in closure

    def toggle_play(event):
        if playing[0]:
            anim.event_source.stop()
            btn.label.set_text("Play")
            playing[0] = False
        else:
            anim.event_source.start()
            btn.label.set_text("Pause")
            playing[0] = True

    btn_ax = plt.axes([0.45, 0.02, 0.1, 0.04])
    btn = Button(btn_ax, "Play")
    btn.on_clicked(toggle_play)
    plt.show()


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python view_npy_video.py <path/to/video.npy>")
        sys.exit(1)
    watch(sys.argv[1])
