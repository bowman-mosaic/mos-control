import pytest

import modules.cavro as cavro


class FakePump:
    syringe_volume_ml = 0.5
    counts_per_ml = 6000
    counts_per_stroke = 3000
    velocity_scale = 2.0

    def __init__(self):
        self.volume_ml = 0.0
        self.commands = []
        self._pending_target = None

    def start_batch(self):
        self.commands.append(("batch",))

    def switch_valve(self, port, **_kwargs):
        self.commands.append(("valve", port))

    def move_absolute_counts(self, target, velocity_counts, wait=False, **_kwargs):
        self.commands.append(("move", target, velocity_counts, wait))
        self._pending_target = target

    def execute(self, wait=True):
        self.commands.append(("execute", wait))
        if self._pending_target is not None:
            self.volume_ml = self._pending_target / self.counts_per_ml
            self._pending_target = None

    def halt(self):
        self.commands.append(("halt",))


@pytest.fixture
def continuous_harness(monkeypatch):
    pump = FakePump()
    flushes = []
    events = []

    old_pump = cavro._pumps[0]
    cavro._pumps[0] = pump
    cavro._cont_stops[0].clear()
    monkeypatch.setattr(cavro, "_max_vel", lambda _pump: 2900)
    monkeypatch.setattr(cavro, "_refill_vel", lambda _pump: 1800)
    monkeypatch.setattr(
        cavro,
        "_flush_to_port",
        lambda _pump, port, velocity, label: flushes.append(
            (port, velocity, label)
        ),
    )
    monkeypatch.setattr(
        cavro,
        "push_event",
        lambda *args: events.append(args),
    )
    try:
        yield pump, flushes, events
    finally:
        cavro._pumps[0] = old_pump
        cavro._cont_stops[0].clear()


def _move_commands(pump):
    return [command for command in pump.commands if command[0] == "move"]


def test_push_mode_applies_flow_rate_to_push_and_returns_to_source(
    continuous_harness,
):
    pump, flushes, _events = continuous_harness

    cavro._continuous_thread(
        0, from_port=4, to_port=3, velocity_ml=0.1, cycles=1, mode="push"
    )

    moves = _move_commands(pump)
    assert moves[0][1:3] == (3000, 1800)  # fixed-rate refill from FROM
    assert moves[1][1:3] == (0, 600)  # selected-rate push to TO
    assert moves[2][1:3] == (3000, 1800)
    assert flushes[0][:2] == (4, 2900)  # remainder returns to FROM


def test_pull_mode_applies_flow_rate_to_pull_and_finishes_to_destination(
    continuous_harness,
):
    pump, flushes, _events = continuous_harness

    cavro._continuous_thread(
        0, from_port=4, to_port=3, velocity_ml=0.1, cycles=1, mode="pull"
    )

    moves = _move_commands(pump)
    assert moves[0][1:3] == (3000, 600)  # selected-rate pull from FROM
    assert moves[1][1:3] == (0, 1800)  # fixed-rate reset push to TO
    assert moves[2][1:3] == (3000, 600)
    assert flushes[0][:2] == (3, 1800)  # remainder completes to TO


def test_push_mode_without_mode_argument_keeps_original_defaults(
    continuous_harness,
):
    pump, flushes, _events = continuous_harness

    cavro._continuous_thread(
        0, from_port=1, to_port=2, velocity_ml=None, cycles=1
    )

    moves = _move_commands(pump)
    assert moves[0][1:3] == (3000, 1800)
    assert moves[1][1:3] == (0, 2900)
    assert flushes[0][:2] == (1, 2900)


def test_continuous_start_rejects_invalid_mode_and_same_port(monkeypatch):
    monkeypatch.setattr(cavro, "_check", lambda _idx: True)

    assert "error" in cavro.cavro_continuous_start(0, 1, 2, mode="sideways")
    assert "error" in cavro.cavro_continuous_start(0, 2, 2, mode="push")


def test_continuous_start_sets_mode_aware_flush_port(monkeypatch):
    class DummyThread:
        def __init__(self, *args, **kwargs):
            pass

        def start(self):
            pass

        def is_alive(self):
            return False

        def join(self, timeout=None):
            pass

    class InstantEvent:
        def __init__(self):
            pass

        def wait(self, timeout=None):
            return True

        def set(self):
            pass

        def clear(self):
            pass

        def is_set(self):
            return False

    monkeypatch.setattr(cavro, "_check", lambda _idx: True)
    monkeypatch.setattr(cavro.threading, "Thread", DummyThread)
    monkeypatch.setattr(cavro.threading, "Event", InstantEvent)

    cavro._flush_port[0] = 1
    assert cavro.cavro_continuous_start(0, 4, 3, 0.1, 0, "push") == {"ok": True}
    assert cavro._flush_port[0] == 4

    assert cavro.cavro_continuous_start(0, 4, 3, 0.1, 0, "pull") == {"ok": True}
    assert cavro._flush_port[0] == 3

    assert cavro.cavro_continuous_start(0, 1, 2, 0.1, 0) == {"ok": True}
    assert cavro._flush_port[0] == 1
