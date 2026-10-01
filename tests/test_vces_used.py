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
    With trains 2 minutes apart, every VCE but a down escalator carries some arriving passengers up,
    so none is misplaced where no passenger ever reaches it, e.g. off its platform,
    or where none would choose it, e.g. past a nearer one, even when the platform is crowded.
    Trains first stop flush with the platform's east end, which is quick to simulate,
    and only if some VCE carries nobody there, at their best stopping position, as the model does.
    """

    def unused(params: model.Params) -> list[str]:
        time_series, _summary = model.simulate(params, print_time_series=False)
        roles = model.vce_roles(params.vces)
        return [
            vce.name
            for i, (vce, role) in enumerate(zip(params.vces, roles, strict=True))
            if role != "down" and sum(second[i][1] for second in time_series.vces) == 0
        ]

    if unused(params):
        params = model.best_stopping_position(params)
        assert not unused(params), f"{params.name}'s {unused(params)} carry nobody up"


def test_scenarios_have_every_vce() -> None:
    """Every VCE in the data is in some scenario."""
    in_scenarios = {vce for params in model.scenarios() for vce in params.vces}
    in_data = {
        vce
        for platform in (model.PLATFORM_A, *range(1, 12))
        for vce in (*model.platform_vces(platform), *model.transformation_vces(platform))
    }
    assert in_data <= in_scenarios
