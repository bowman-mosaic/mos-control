"""3×3 mosaic step size and capture sequencing, without a microscope."""

import os
import sys

import numpy as np

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import modules.coolsnap as coolsnap
import modules.nikon_ti as ti


def test_step_is_one_field_minus_the_overlap():
    # 1392 sensor pixels * 0.645 µm/px = 897.84 µm. 10% overlap leaves 808.056 µm.
    assert coolsnap.mosaic_step_nm(1392, 0.645, 0.1) == 808056
    assert coolsnap.mosaic_step_nm(1000, 1.0, 0) == 1_000_000


def test_tiles_snake_across_a_centered_3x3():
    assert coolsnap.mosaic_tile_offsets() == [
        (-1, 1), (0, 1), (1, 1),
        (1, 0), (0, 0), (-1, 0),
        (-1, -1), (0, -1), (1, -1),
    ]


def test_capture_saves_nine_tiles_and_returns_to_the_start(monkeypatch, tmp_path):
    moves = []
    monkeypatch.setattr(coolsnap, "_hcam", object())
    monkeypatch.setattr(coolsnap, "_sensor_w", 1000)
    monkeypatch.setattr(coolsnap, "_sensor_h", 800)
    monkeypatch.setattr(coolsnap, "_save_dir", str(tmp_path))
    monkeypatch.setattr(coolsnap, "_experiment_meta", None)
    monkeypatch.setattr(coolsnap, "_MOSAIC_SETTLE_S", 0)
    monkeypatch.setattr(coolsnap, "live_is_active", lambda: False)
    monkeypatch.setattr(coolsnap, "snap", lambda: np.zeros((4, 5), dtype=np.uint16))
    monkeypatch.setattr(
        coolsnap,
        "_capture_scale_metadata",
        lambda *args, **kwargs: {"um_per_pixel_1x": 1.0, "binning": 1},
    )
    monkeypatch.setattr(ti, "is_connected", lambda: True)
    monkeypatch.setattr(ti, "xy_get_position", lambda: {"x": 5_000_000, "y": 7_000_000})
    monkeypatch.setattr(ti, "x_move_absolute", lambda nm: moves.append(("x", int(nm))))
    monkeypatch.setattr(ti, "y_move_absolute", lambda nm: moves.append(("y", int(nm))))

    result = coolsnap.capture_mosaic(overlap=0)

    assert result["count"] == 9
    assert result["cancelled"] is False
    assert result["step_x_nm"] == 1_000_000
    assert result["step_y_nm"] == 800_000
    assert [path.name for path in sorted(tmp_path.glob("mosaic_*.tif"), key=lambda p: p.name)] == [
        f"mosaic_{n}.tif" for n in range(1, 10)
    ]
    # Upper-left tile, then the stage is sent home after the last tile.
    assert moves[0] == ("x", 4_000_000)
    assert moves[1] == ("y", 7_800_000)
    assert moves[-2] == ("x", 5_000_000)
    assert moves[-1] == ("y", 7_000_000)


def test_cancel_stops_between_tiles_and_still_returns_home(monkeypatch, tmp_path):
    moves = []
    monkeypatch.setattr(coolsnap, "_hcam", object())
    monkeypatch.setattr(coolsnap, "_sensor_w", 100)
    monkeypatch.setattr(coolsnap, "_sensor_h", 100)
    monkeypatch.setattr(coolsnap, "_save_dir", str(tmp_path))
    monkeypatch.setattr(coolsnap, "_experiment_meta", None)
    monkeypatch.setattr(coolsnap, "_MOSAIC_SETTLE_S", 0)
    monkeypatch.setattr(coolsnap, "live_is_active", lambda: False)
    monkeypatch.setattr(
        coolsnap,
        "_capture_scale_metadata",
        lambda *args, **kwargs: {"um_per_pixel_1x": 1.0},
    )

    def snap_and_cancel():
        coolsnap.mosaic_cancel()
        return np.zeros((2, 2), dtype=np.uint16)

    monkeypatch.setattr(coolsnap, "snap", snap_and_cancel)
    monkeypatch.setattr(ti, "is_connected", lambda: True)
    monkeypatch.setattr(ti, "xy_get_position", lambda: {"x": 0, "y": 0})
    monkeypatch.setattr(ti, "x_move_absolute", lambda nm: moves.append(("x", int(nm))))
    monkeypatch.setattr(ti, "y_move_absolute", lambda nm: moves.append(("y", int(nm))))

    result = coolsnap.capture_mosaic(overlap=0)

    assert result["cancelled"] is True
    assert result["count"] == 1
    assert moves[-2:] == [("x", 0), ("y", 0)]
