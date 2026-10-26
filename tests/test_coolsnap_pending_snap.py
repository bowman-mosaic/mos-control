import os

import numpy as np
import tifffile

import modules.coolsnap as coolsnap


def test_preview_snap_can_be_saved_as_full_resolution_tiff(monkeypatch, tmp_path):
    frame = np.arange(48, dtype=np.uint16).reshape(6, 8)
    monkeypatch.setattr(coolsnap, "snap", lambda: frame.copy())
    monkeypatch.setattr(coolsnap, "_frame_to_base64", lambda _frame: "preview")
    monkeypatch.setattr(
        coolsnap,
        "_capture_scale_metadata",
        lambda: {"objective_position": 4, "binning": 1},
    )
    monkeypatch.setattr(coolsnap, "_save_dir", str(tmp_path))
    monkeypatch.setattr(coolsnap, "_base_save_dir", str(tmp_path))
    coolsnap.cam_snap_discard()

    preview = coolsnap.cam_snap()
    assert preview == {
        "ok": True,
        "image": "preview",
        "width": 8,
        "height": 6,
    }

    saved = coolsnap.cam_snap_save_pending()
    assert saved["ok"] is True
    assert saved["file"].endswith(".tif")
    saved_path = os.path.join(tmp_path, saved["file"])
    np.testing.assert_array_equal(tifffile.imread(saved_path), frame)
    assert tifffile.imread(saved_path).dtype == np.uint16
    assert "error" in coolsnap.cam_snap_save_pending()


def test_discard_removes_pending_full_resolution_snap(monkeypatch):
    frame = np.zeros((4, 5), dtype=np.uint16)
    monkeypatch.setattr(coolsnap, "snap", lambda: frame)
    monkeypatch.setattr(coolsnap, "_frame_to_base64", lambda _frame: "preview")
    monkeypatch.setattr(coolsnap, "_capture_scale_metadata", lambda: None)

    assert coolsnap.cam_snap()["ok"] is True
    assert coolsnap.cam_snap_discard()["ok"] is True
    assert "error" in coolsnap.cam_snap_save_pending()
