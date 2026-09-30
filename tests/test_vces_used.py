"""Tests that every VCE in every scenario is used."""

import pytest

from platform_crowd_model import model


@pytest.mark.parametrize(
    "params",
    [params for params in model.scenarios() if params.headway == model.CLOSE_HEADWAY],
    ids=lambda params: params.name,
)
def test_every_vce_carries_passengers_up(params: model.Params) -> None:
    """
    With trains 2 minutes apart, stopped flush with the platform's east end,
    every VCE but a down escalator carries some arriving passengers up,
    so none is misplaced where no passenger ever reaches it, e.g. off its platform.
    The trains aren't moved to their best stopping position, which would take much longer.
    A VCE entirely west of the trains, e.g. a stair far out on a long platform,
    can go unused if every passenger has a nearer one, so it isn't checked.
    """
    time_series, _summary = model.simulate(params, print_time_series=False)
    roles = model.vce_roles(params.vces)
    train_west_end = params.platform_east_end - params.train_length
    unused = [
        vce.name
        for i, (vce, role) in enumerate(zip(params.vces, roles, strict=True))
        if role != "down"
        and vce.east_end > train_west_end
        and sum(second[i][1] for second in time_series.vces) == 0
    ]
    assert not unused, f"{params.name}'s {unused} carry nobody up"


def test_scenarios_have_every_vce() -> None:
    """Every VCE in the data is in some scenario."""
    in_scenarios = {vce for params in model.scenarios() for vce in params.vces}
    in_data = {
        vce
        for platform in (model.PLATFORM_A, *range(1, 12))
        for vce in (*model.platform_vces(platform), *model.transformation_vces(platform))
    }
    assert in_data <= in_scenarios
