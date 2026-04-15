"""Unit conversion between mL and pump encoder counts."""

from typing import Optional


def ml_to_counts(ml: Optional[float], counts_per_ml: float) -> Optional[float]:
    if ml is None:
        return None
    return ml * counts_per_ml


def counts_to_ml(counts: Optional[float], counts_per_ml: float) -> Optional[float]:
    if counts is None:
        return None
    return counts / counts_per_ml
