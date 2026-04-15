"""
Tests for modules.nikon_ti — Nikon Eclipse Ti COM interface.

All tests mock the COM layer. The real interface uses:
  - scope = CreateObject("Nikon.TiScope.NikonTi")
  - scope.EpiShutter, scope.DiaLamp, etc. as sub-device properties
  - IMipParameter objects with .RawValue for reading/writing values

Run with:  python -m pytest tests/test_nikon_ti.py -v
"""

import sys
import os
import types
import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

# ── Stub eel and comtypes ────────────────────────────────────────────────────

_fake_eel = types.ModuleType("eel")
_fake_eel.expose = lambda fn: fn
sys.modules["eel"] = _fake_eel

_fake_comtypes = types.ModuleType("comtypes")
_fake_comtypes_client = types.ModuleType("comtypes.client")
_fake_comtypes.client = _fake_comtypes_client
_fake_comtypes.CoInitialize = lambda: None
_fake_comtypes.CoUninitialize = lambda: None
_fake_comtypes_client.CreateObject = None
sys.modules["comtypes"] = _fake_comtypes
sys.modules["comtypes.client"] = _fake_comtypes_client

from modules import nikon_ti  # noqa: E402


# ── Mock COM objects ────────────────────────────────────────────────────────

class MockParam:
    """Simulates an IMipParameter with RawValue."""
    def __init__(self, value=0):
        self.RawValue = value


class MockShutter:
    def __init__(self):
        self.Value = MockParam(0)
        self.IsMounted = MockParam(1)

    def Open(self):
        self.Value.RawValue = 1

    def Close(self):
        self.Value.RawValue = 0


class MockDiaLamp:
    def __init__(self):
        self.Value = MockParam(0)
        self.Position = MockParam(20)
        self.IsMounted = MockParam(1)
        self.LowerLimit = MockParam(0)
        self.UpperLimit = MockParam(100)


class MockPositionDevice:
    def __init__(self, name="Device", unit=""):
        self.Name = name
        self.Unit = unit
        self.Position = MockParam(1)
        self.Value = MockParam(1)
        self.IsMounted = MockParam(1)
        self.LowerLimit = MockParam(1)
        self.UpperLimit = MockParam(6)


class MockDrive:
    def __init__(self, name="Drive", unit="nm"):
        self.Name = name
        self.Unit = unit
        self.Position = MockParam(0)
        self.Value = MockParam(0)
        self.IsMounted = MockParam(1)
        self.LowerLimit = MockParam(-10000000)
        self.UpperLimit = MockParam(10000000)
        self.Speed = MockParam(1)

    def MoveAbsolute(self, val):
        self.Position.RawValue = val

    def MoveRelative(self, delta):
        self.Position.RawValue += delta


class MockPFS:
    def __init__(self):
        self.Value = MockParam(0)
        self.Position = MockParam(0)
        self.Status = MockParam(0)
        self.IsMounted = MockParam(1)
        self._enabled = False

    def Enable(self):
        self._enabled = True
        self.Value.RawValue = 1

    def Disable(self):
        self._enabled = False
        self.Value.RawValue = 0

    def SearchPosition(self):
        pass


class MockScope:
    """Simulates the INikonTi root COM object with sub-device properties."""
    def __init__(self):
        self.EpiShutter = MockShutter()
        self.DiaShutter = MockShutter()
        self.DiaLamp = MockDiaLamp()
        self.Nosepiece = MockPositionDevice("Nosepiece")
        self.FilterBlockCassette1 = MockPositionDevice("FilterBlock")
        self.LightPathDrive = MockPositionDevice("LightPath")
        self.ZDrive = MockDrive("Z Drive", "nm")
        self.XDrive = MockDrive("X Drive", "nm")
        self.YDrive = MockDrive("Y Drive", "nm")
        self.PFS = MockPFS()
        self.SystemType = MockParam(1)


# ── Fixtures ────────────────────────────────────────────────────────────────

@pytest.fixture(autouse=True)
def _reset():
    nikon_ti._scope = None
    nikon_ti._com_initialized = False
    nikon_ti._HAS_COMTYPES = True
    yield
    nikon_ti._scope = None
    nikon_ti._com_initialized = False


@pytest.fixture()
def scope(monkeypatch):
    """Inject a MockScope and connect."""
    ms = MockScope()
    monkeypatch.setattr(nikon_ti.comtypes.client, "CreateObject", lambda progid: ms)
    nikon_ti.connect()
    return ms


# ── Connection ──────────────────────────────────────────────────────────────

class TestConnection:
    def test_connect(self, scope):
        assert nikon_ti.is_connected()

    def test_disconnect(self, scope):
        nikon_ti.disconnect()
        assert not nikon_ti.is_connected()

    def test_disconnect_safe_when_not_connected(self):
        nikon_ti.disconnect()

    def test_connect_fails_without_comtypes(self):
        nikon_ti._HAS_COMTYPES = False
        with pytest.raises(nikon_ti.TiError, match="comtypes"):
            nikon_ti.connect()


# ── Shutter (Lambda SC via DiaShutter) ──────────────────────────────────────

class TestShutter:
    def test_open(self, scope):
        nikon_ti.shutter_open()
        assert scope.DiaShutter.Value.RawValue == 1

    def test_close(self, scope):
        nikon_ti.shutter_open()
        nikon_ti.shutter_close()
        assert scope.DiaShutter.Value.RawValue == 0

    def test_get_state(self, scope):
        nikon_ti.shutter_open()
        assert nikon_ti.shutter_get_state() == 1

    def test_raises_when_disconnected(self):
        with pytest.raises(nikon_ti.TiError):
            nikon_ti.shutter_open()


# ── Dia lamp ────────────────────────────────────────────────────────────────

class TestDiaLamp:
    def test_on_off(self, scope):
        nikon_ti.dia_lamp_on()
        assert scope.DiaLamp.Value.RawValue == 1
        nikon_ti.dia_lamp_off()
        assert scope.DiaLamp.Value.RawValue == 0

    def test_set_intensity(self, scope):
        nikon_ti.dia_lamp_set_intensity(75)
        assert scope.DiaLamp.Position.RawValue == 75

    def test_get_state(self, scope):
        nikon_ti.dia_lamp_on()
        nikon_ti.dia_lamp_set_intensity(50)
        state = nikon_ti.dia_lamp_get_state()
        assert state["on"] == 1
        assert state["intensity"] == 50


# ── Nosepiece ───────────────────────────────────────────────────────────────

class TestNosepiece:
    def test_set_get(self, scope):
        nikon_ti.nosepiece_set_position(4)
        assert nikon_ti.nosepiece_get_position() == 4

    def test_cycle(self, scope):
        for p in [1, 2, 3, 4, 5, 6]:
            nikon_ti.nosepiece_set_position(p)
            assert nikon_ti.nosepiece_get_position() == p


# ── Filter block ────────────────────────────────────────────────────────────

class TestFilterBlock:
    def test_set_get(self, scope):
        nikon_ti.filter_set_position(3)
        assert nikon_ti.filter_get_position() == 3


# ── Light path ──────────────────────────────────────────────────────────────

class TestLightPath:
    def test_set_get(self, scope):
        nikon_ti.light_path_set_position(2)
        assert nikon_ti.light_path_get_position() == 2


# ── Z drive ─────────────────────────────────────────────────────────────────

class TestZDrive:
    def test_move_absolute(self, scope):
        nikon_ti.z_move_absolute(5000000)
        assert nikon_ti.z_get_position() == 5000000

    def test_move_relative(self, scope):
        nikon_ti.z_move_absolute(1000000)
        nikon_ti.z_move_relative(50000)
        assert nikon_ti.z_get_position() == 1050000

    def test_move_relative_negative(self, scope):
        nikon_ti.z_move_absolute(1000000)
        nikon_ti.z_move_relative(-200000)
        assert nikon_ti.z_get_position() == 800000


# ── XY drive ────────────────────────────────────────────────────────────────

class TestXYDrive:
    def test_get_position(self, scope):
        scope.XDrive.Position.RawValue = 100000
        scope.YDrive.Position.RawValue = 200000
        pos = nikon_ti.xy_get_position()
        assert pos["x"] == 100000
        assert pos["y"] == 200000

    def test_move_relative(self, scope):
        nikon_ti.x_move_relative(5000)
        nikon_ti.y_move_relative(-3000)
        pos = nikon_ti.xy_get_position()
        assert pos["x"] == 5000
        assert pos["y"] == -3000


# ── PFS ─────────────────────────────────────────────────────────────────────

class TestPFS:
    def test_enable_disable(self, scope):
        nikon_ti.pfs_enable()
        assert scope.PFS._enabled is True
        nikon_ti.pfs_disable()
        assert scope.PFS._enabled is False

    def test_get_status(self, scope):
        status = nikon_ti.pfs_get_status()
        assert "value" in status
        assert "position" in status
        assert "status" in status


# ── Full status ─────────────────────────────────────────────────────────────

class TestFullStatus:
    def test_disconnected(self):
        assert nikon_ti.get_full_status()["connected"] is False

    def test_connected(self, scope):
        s = nikon_ti.get_full_status()
        assert s["connected"] is True
        assert "shutter" in s
        assert "z_position" in s


# ── Eel wrappers ────────────────────────────────────────────────────────────

class TestEelWrappers:
    def test_connect(self, monkeypatch):
        ms = MockScope()
        monkeypatch.setattr(nikon_ti.comtypes.client, "CreateObject", lambda p: ms)
        assert nikon_ti.ti_connect()["ok"] is True

    def test_shutter_open(self, scope):
        assert nikon_ti.ti_shutter_open()["ok"] is True

    def test_dia_lamp_intensity(self, scope):
        assert nikon_ti.ti_dia_lamp_set_intensity(40)["ok"] is True

    def test_z_move(self, scope):
        nikon_ti.z_move_absolute(1000000)
        r = nikon_ti.ti_z_move_rel(50000)
        assert r["ok"] is True

    def test_error_when_disconnected(self):
        for fn in [nikon_ti.ti_shutter_open, nikon_ti.ti_dia_lamp_off,
                    nikon_ti.ti_pfs_enable]:
            assert "error" in fn()
