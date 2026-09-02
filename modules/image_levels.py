"""Display-only level estimation for scientific grayscale images.

The functions in this module never modify the source image. They estimate a
robust black/white display window and map that window to an 8-bit preview.
"""

from __future__ import annotations

import math

import numpy as np


# Match ImageJ's commonly used 0.35% total saturation, split across both tails.
DEFAULT_LOW_PERCENTILE = 0.175
DEFAULT_HIGH_PERCENTILE = 99.825
DEFAULT_MAX_SAMPLES = 100_000


def sample_pixels(frame, max_samples=DEFAULT_MAX_SAMPLES):
    """Return a deterministic, approximately uniform sample of a 2-D image."""
    image = np.asarray(frame)
    if image.ndim != 2:
        raise ValueError(f"Expected a 2-D image, got shape {image.shape}")
    if image.size == 0:
        raise ValueError("Cannot estimate levels for an empty image")

    max_samples = max(1, int(max_samples))
    if image.size <= max_samples:
        sample = image.ravel()
    else:
        step = max(1, int(math.ceil(math.sqrt(image.size / max_samples))))
        sample = image[::step, ::step].ravel()

    if np.issubdtype(sample.dtype, np.floating):
        sample = sample[np.isfinite(sample)]
        if sample.size == 0:
            raise ValueError("Image contains no finite pixel values")
    return sample


def estimate_display_levels(
    frame,
    low_percentile=DEFAULT_LOW_PERCENTILE,
    high_percentile=DEFAULT_HIGH_PERCENTILE,
    max_samples=DEFAULT_MAX_SAMPLES,
):
    """Estimate robust black and white display levels.

    A small fraction of each histogram tail is ignored so isolated hot/dead
    pixels do not flatten the useful contrast. Constant images receive a
    one-value-wide window to keep downstream normalization well-defined.
    """
    low_percentile = float(low_percentile)
    high_percentile = float(high_percentile)
    if not 0.0 <= low_percentile < high_percentile <= 100.0:
        raise ValueError("Percentiles must satisfy 0 <= low < high <= 100")

    sample = sample_pixels(frame, max_samples)
    data_min = float(np.min(sample))
    data_max = float(np.max(sample))
    low, high = np.percentile(
        sample, (low_percentile, high_percentile)
    ).astype(float)

    # Very narrow or quantized images can collapse both percentiles. Fall back
    # to their sampled data range before using the constant-frame guard.
    if high <= low and data_max > data_min:
        low, high = data_min, data_max
    if high <= low:
        high = low + 1.0

    return float(low), float(high)


def smooth_display_levels(
    current,
    target,
    slow_alpha=0.15,
    fast_alpha=0.5,
    deadband_fraction=0.01,
):
    """Stabilize continuously tracked levels without hiding sudden clipping.

    The window expands quickly when new dark or bright values appear and
    contracts slowly as they disappear. Small changes inside the deadband are
    ignored to prevent visible frame-to-frame breathing.
    """
    current_low, current_high = map(float, current)
    target_low, target_high = map(float, target)
    if current_high <= current_low:
        return target_low, target_high

    span = max(current_high - current_low, 1.0)
    deadband = max(0.0, float(deadband_fraction)) * span
    slow_alpha = min(1.0, max(0.0, float(slow_alpha)))
    fast_alpha = min(1.0, max(0.0, float(fast_alpha)))

    def update(value, wanted, expanding):
        if abs(wanted - value) <= deadband:
            return value
        alpha = fast_alpha if expanding else slow_alpha
        return value + alpha * (wanted - value)

    low = update(current_low, target_low, target_low < current_low)
    high = update(current_high, target_high, target_high > current_high)
    if high <= low:
        return target_low, target_high
    return low, high
