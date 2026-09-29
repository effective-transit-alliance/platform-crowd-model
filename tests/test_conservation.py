"""Tests that the model neither creates nor loses passengers."""

import itertools

import pytest

from platform_crowd_model import model

TOLERANCE = 1
"""
How far off a total can be (pax).
The simulation stops once fewer than 1 arriving passenger is left on the platform,
so up to that many may not have gone up yet.
"""


@pytest.mark.parametrize(
    "params",
    model.scenarios(),
    ids=lambda params: f"{params.name}-{round(params.headway.total_seconds())}s",
)
def test_passengers_are_conserved(params: model.Params) -> None:
    """
    Every arriving passenger alights and goes up,
    every departing passenger comes down and boards,
    and the platform's count each second is what's come onto it minus what's left it.

    This uses the trains' default stopping position, not `best_stopping_position`'s,
    since conservation shouldn't depend on where they stop, and searching is slow.
    """
    time_series, summary = model.simulate(params, print_time_series=False)
    instants = time_series.instants
    arriving = params.trains * params.arriving_pax_per_train
    departing = params.trains * params.assumptions.departing_pax_per_train

    assert all(dwell is not None for dwell in summary.dwells)
    assert sum(instant.off_rate for instant in instants) == pytest.approx(arriving, abs=TOLERANCE)
    assert sum(instant.up_rate for instant in instants) == pytest.approx(arriving, abs=TOLERANCE)
    assert sum(instant.down_rate for instant in instants) == pytest.approx(departing, abs=TOLERANCE)
    assert sum(instant.on_rate for instant in instants) == pytest.approx(departing, abs=TOLERANCE)
    assert instants[-1].total_pax_on_platform < TOLERANCE

    on_platform = itertools.accumulate(
        instant.off_rate + instant.down_rate - instant.up_rate - instant.on_rate
        for instant in instants
    )
    for instant, expected in zip(instants, on_platform, strict=True):
        assert instant.total_pax_on_platform == pytest.approx(expected, abs=1e-6), instant.time
