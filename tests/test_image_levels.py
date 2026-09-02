import numpy as np
import pytest

from modules.image_levels import (
    estimate_display_levels,
    sample_pixels,
    smooth_display_levels,
)


def test_sample_pixels_is_bounded_and_two_dimensional():
    frame = np.arange(1_000_000, dtype=np.uint16).reshape(1000, 1000)
    sample = sample_pixels(frame, max_samples=10_000)
    assert sample.ndim == 1
    assert sample.size <= 10_000


def test_estimate_levels_ignores_isolated_hot_and_dead_pixels():
    frame = np.linspace(100, 1000, 400 * 400, dtype=np.uint16).reshape(400, 400)
    frame[0, 0] = 0
    frame[-1, -1] = 65535

    low, high = estimate_display_levels(frame)

    assert 90 <= low <= 120
    assert 980 <= high <= 1010


def test_estimate_levels_handles_constant_image():
    frame = np.full((32, 32), 1234, dtype=np.uint16)
    low, high = estimate_display_levels(frame)
    assert low == 1234
    assert high == 1235


def test_estimate_levels_ignores_nonfinite_float_values():
    frame = np.arange(100, dtype=np.float32).reshape(10, 10)
    frame[0, 0] = np.nan
    frame[0, 1] = np.inf
    low, high = estimate_display_levels(frame)
    assert np.isfinite(low)
    assert np.isfinite(high)
    assert high > low


def test_estimate_levels_rejects_invalid_input():
    with pytest.raises(ValueError):
        estimate_display_levels(np.zeros((2, 2, 2), dtype=np.uint16))
    with pytest.raises(ValueError):
        estimate_display_levels(np.zeros((2, 2)), 99, 1)


def test_smoothing_expands_faster_than_it_contracts():
    expanded = smooth_display_levels((100, 1000), (0, 2000))
    contracted = smooth_display_levels((100, 1000), (200, 900))

    assert expanded == pytest.approx((50, 1500))
    assert contracted == pytest.approx((115, 985))


def test_smoothing_ignores_small_changes():
    levels = smooth_display_levels((100, 1100), (105, 1095))
    assert levels == (100, 1100)
