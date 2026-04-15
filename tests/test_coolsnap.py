"""
Tests for modules.coolsnap — CoolSNAP EZ camera via PyVCAM.

All tests mock the PyVCAM layer so they can run without hardware.
"""

import sys
import os
import types
import time
import threading
import numpy as np
import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

# ── Stub eel ─────────────────────────────────────────────────────────────────

_fake_eel = types.ModuleType("eel")
_fake_eel.expose = lambda fn: fn
_fake_eel.onLiveFrame = lambda *a: (lambda: None)
_fake_eel.onCamStatus = lambda *a: (lambda: None)
_fake_eel.onCamCaptureComplete = lambda *a: (lambda: None)
_fake_eel.onTimelapseProgress = lambda *a: (lambda: None)
sys.modules.setdefault("eel", _fake_eel)

# ── Mock PyVCAM ──────────────────────────────────────────────────────────────

SENSOR_W, SENSOR_H = 1392, 1040


class MockCamera:
    def __init__(self):
        self.name = "CoolSNAP_EZ_Mock"
        self.sensor_size = (SENSOR_W, SENSOR_H)
        self.bit_depth = 12
        self.bit_depth_host = 16
        self._open = False
        self._live = False
        self._frame_count = 0

    def open(self):
        self._open = True

    def close(self):
        self._open = False

    def get_frame(self, exp_time=20):
        return np.random.randint(0, 4095, (SENSOR_H, SENSOR_W), dtype=np.uint16)

    def start_live(self, exp_time=20):
        self._live = True
        self._frame_count = 0

    def poll_frame(self):
        if not self._live:
            raise RuntimeError("Not in live mode")
        self._frame_count += 1
        frame = np.random.randint(0, 4095, (SENSOR_H, SENSOR_W), dtype=np.uint16)
        return ({'pixel_data': frame}, 30.0, self._frame_count)

    def finish(self):
        self._live = False

    @staticmethod
    def detect_camera():
        yield MockCamera()


_mock_cam_instance = None

_fake_pvc = types.ModuleType("pyvcam.pvc")
_fake_pvc.init_pvcam = lambda: None
_fake_pvc.uninit_pvcam = lambda: None

_fake_pyvcam = types.ModuleType("pyvcam")
_fake_pyvcam.pvc = _fake_pvc

_fake_pyvcam_camera = types.ModuleType("pyvcam.camera")
_fake_pyvcam_camera.Camera = MockCamera

sys.modules["pyvcam"] = _fake_pyvcam
sys.modules["pyvcam.pvc"] = _fake_pvc
sys.modules["pyvcam.camera"] = _fake_pyvcam_camera

from modules import coolsnap  # noqa: E402


# ── Fixtures ─────────────────────────────────────────────────────────────────

@pytest.fixture(autouse=True)
def _reset(tmp_path):
    coolsnap._cam = None
    coolsnap._pvc_initialized = False
    coolsnap._exposure_ms = 20
    coolsnap._binning = (1, 1)
    coolsnap._save_dir = str(tmp_path)
    coolsnap._live_stop.set()
    coolsnap._capture_stop.set()
    yield
    coolsnap._cam = None
    coolsnap._pvc_initialized = False


@pytest.fixture()
def cam():
    coolsnap.connect()
    return coolsnap._cam


# ── Connection ──────────────────────────────────────────────────────────────

class TestConnection:
    def test_connect(self, cam):
        assert coolsnap.is_connected()

    def test_disconnect(self, cam):
        coolsnap.disconnect()
        assert not coolsnap.is_connected()

    def test_camera_info(self, cam):
        info = coolsnap.get_camera_info()
        assert info["name"] == "CoolSNAP_EZ_Mock"
        assert info["sensor_size"] == [SENSOR_W, SENSOR_H]
        assert info["bit_depth"] == 12

    def test_info_raises_when_disconnected(self):
        with pytest.raises(coolsnap.CamError):
            coolsnap.get_camera_info()


# ── Settings ────────────────────────────────────────────────────────────────

class TestSettings:
    def test_exposure(self):
        coolsnap.set_exposure(50)
        assert coolsnap.get_exposure() == 50

    def test_exposure_min_clamp(self):
        coolsnap.set_exposure(-10)
        assert coolsnap.get_exposure() == 1

    def test_binning_valid(self):
        for b in [1, 2, 4, 8]:
            coolsnap.set_binning(b)
            assert coolsnap.get_binning() == b

    def test_binning_invalid(self):
        with pytest.raises(coolsnap.CamError):
            coolsnap.set_binning(3)


# ── Snap ────────────────────────────────────────────────────────────────────

class TestSnap:
    def test_snap_returns_array(self, cam):
        frame = coolsnap.snap()
        assert isinstance(frame, np.ndarray)
        assert frame.shape == (SENSOR_H, SENSOR_W)

    def test_snap_raises_when_disconnected(self):
        with pytest.raises(coolsnap.CamError):
            coolsnap.snap()

    def test_snap_and_save(self, cam, tmp_path):
        coolsnap._save_dir = str(tmp_path)
        frame, path = coolsnap.snap_and_save()
        assert os.path.exists(path)
        assert path.endswith(".npy")
        loaded = np.load(path)
        assert loaded.shape == frame.shape


# ── Video ───────────────────────────────────────────────────────────────────

class TestVideo:
    def test_record_video(self, cam):
        video = coolsnap.record_video(num_frames=5)
        assert video.shape[0] == 5
        assert video.shape[1] == SENSOR_H
        assert video.shape[2] == SENSOR_W

    def test_record_and_save(self, cam, tmp_path):
        coolsnap._save_dir = str(tmp_path)
        path = coolsnap.record_video_and_save(num_frames=3)
        assert os.path.exists(path)
        data = np.load(path)
        assert data.shape[0] == 3


# ── Time-lapse ──────────────────────────────────────────────────────────────

class TestTimelapse:
    def test_timelapse(self, cam):
        stack = coolsnap.timelapse(num_frames=3, interval_sec=0.01)
        assert stack.shape[0] == 3

    def test_timelapse_save(self, cam, tmp_path):
        coolsnap._save_dir = str(tmp_path)
        path = coolsnap.timelapse_and_save(num_frames=2, interval_sec=0.01)
        assert os.path.exists(path)


# ── Frame conversion ────────────────────────────────────────────────────────

class TestFrameConversion:
    def test_to_base64(self):
        frame = np.random.randint(0, 4095, (100, 100), dtype=np.uint16)
        b64 = coolsnap._frame_to_base64(frame)
        assert isinstance(b64, str)
        assert len(b64) > 100


# ── Eel wrappers ────────────────────────────────────────────────────────────

class TestEelWrappers:
    def test_cam_connect(self):
        r = coolsnap.cam_connect()
        assert r["ok"]
        assert "CoolSNAP_EZ_Mock" in r["name"]

    def test_cam_snap(self, cam):
        r = coolsnap.cam_snap()
        assert r["ok"]
        assert "image" in r
        assert r["width"] == SENSOR_W

    def test_cam_exposure(self):
        r = coolsnap.cam_set_exposure(100)
        assert r["ok"]
        assert coolsnap.cam_get_exposure() == 100

    def test_cam_binning(self):
        r = coolsnap.cam_set_binning(2)
        assert r["ok"]
        assert coolsnap.cam_get_binning() == 2

    def test_error_when_disconnected(self):
        r = coolsnap.cam_snap()
        assert "error" in r
