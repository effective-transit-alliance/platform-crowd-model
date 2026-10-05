"""Tests that `_core`, the Rust `simulate` loop, matches `simulate_python` bit for bit."""

import dataclasses
import random

from platform_crowd_model import model
from platform_crowd_model._core import py_sum_for_test


def test_py_sum_adds_like_sum() -> None:
    """`_core`'s `py_sum` gives exactly what Python's compensated `sum()` does."""
    rng = random.Random(0)
    for _ in range(2000):
        values = [
            rng.choice([0.0, -0.0, 1e-15, 1e16, rng.uniform(-1, 1) * 10 ** rng.randint(-12, 12)])
            for _ in range(rng.randrange(30))
        ]
        assert py_sum_for_test(values) == sum(values)


def test_core_matches_python() -> None:
    """
    Every third scenario's summary is identical from both backends,
    with its trains stopped at the platform's east end.
    They're compared exactly, not within a tolerance,
    since even rounding differences can change which stopping position is best.
    """
    for params in model.scenarios()[::3]:
        python = model.simulate(params, False, False, backend="python")
        rust = model.simulate(params, False, False, backend="rust")
        assert rust == python, params.name


def test_core_matches_python_choosing_nearest_vce() -> None:
    """Platform 3's summary is identical from both backends with `vce_choice` `nearest`, too."""
    for params in model.scenarios():
        if params.name != "3":
            continue
        params = dataclasses.replace(
            params, assumptions=dataclasses.replace(params.assumptions, vce_choice="nearest")
        )
        python = model.simulate(params, False, False, backend="python")
        rust = model.simulate(params, False, False, backend="rust")
        assert rust == python, params.name


def test_core_time_series_matches_python() -> None:
    """
    A few scenarios' time series are identical from both backends:
    platform 3's with trains 2 minutes apart, and platform 9's, on one track, all at once.
    """
    chosen = [
        params
        for params in model.scenarios()
        if (params.name, params.headway.total_seconds()) in {("3", 120), ("9", 0)}
    ]
    assert len(chosen) == 2
    for params in chosen:
        python = model.simulate(params, print_time_series=False, backend="python")
        rust = model.simulate(params, print_time_series=False, backend="rust")
        assert rust == python, params.name
