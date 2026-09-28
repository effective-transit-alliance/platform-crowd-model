#!/usr/bin/env -S uv run

"""
This is a recursive peak-hour platform clearance calculator.
model from https://onlinepubs.trb.org/Onlinepubs/hrr/1971/355/355-001.pdf
"""

import csv
import dataclasses
import functools
import itertools
import typing
from collections import defaultdict
from collections.abc import Generator
from concurrent.futures import ProcessPoolExecutor
from dataclasses import dataclass
from pathlib import Path
from typing import TYPE_CHECKING, Annotated, Any, Literal, Self, cast

import openpyxl
import typer
from openpyxl.cell import Cell
from openpyxl.chart import Reference, ScatterChart
from openpyxl.chart.series_factory import SeriesFactory
from openpyxl.worksheet.worksheet import Worksheet
from typer import Option

if TYPE_CHECKING:
    from _typeshed import DataclassInstance

SECONDS_PER_MINUTE = 60

SQUARE_METERS_PER_SQUARE_FOOT = 0.09290304

CLOSE_HEADWAY = 120
"""Time between two trains' arrivals in the closely spaced scenarios (s), from the ETA report."""

MAX_SIMULATION_LENGTH = 7200
"""
The simulation runs until the last train departs and the platform clears,
but stops with an error if that takes longer than this (s), which means something's wrong.
"""

NORMAL_HEADWAY = 300
"""Time between two trains' arrivals in the normal scenarios (s), from the ETA report."""


@dataclass(frozen=True)
class Assumptions:
    """
    Everything the model assumes, shared by every scenario,
    as opposed to the facts about each scenario in `Params`.
    The defaults are the model's current assumptions;
    override any of them to see how sensitive the results are to it,
    e.g. `Assumptions(stair_capacity=15)`.
    """

    usable_platform_area_multiplier: Annotated[
        float, Field(name="Usable Platform Area Multiplier", units="fraction")
    ] = 0.75
    """
    Fraction of the platform's area usable by passengers,
    leaving the rest for columns, stairwells, and other obstructions.
    From the ETA report.
    """

    max_train_cars: Annotated[int, Field(name="Max Train Cars", units="car")] = 12
    """
    Cars in the longest trains, on platforms long enough for them.
    Shorter platforms get trains as long as `Params.platform_max_cars`.
    """

    seats_per_car: Annotated[int, Field(name="Seats per Car", units="pax")] = 135
    """
    Seats in each car, all of which are full on arrival, and all of whose passengers alight.
    A seated NJ Transit car,
    from the Moynihan Station Development Project environmental assessment,
    chapter 4.4, Station Circulation Analysis, Tables 4.4-10 and 4.4-19,
    which have 1,620 passengers on a 12-car train:
    https://web.archive.org/web/20241011135133/https://cdn.esd.ny.gov/subsidiaries_projects/msdc/Data/NEPA/04_4%20StationPedCirculation.pdf#page=22
    https://web.archive.org/web/20241011135133/https://cdn.esd.ny.gov/subsidiaries_projects/msdc/Data/NEPA/04_4%20StationPedCirculation.pdf#page=47
    """

    departing_pax_per_train: Annotated[
        int, Field(name="Departing Passengers per Train", units="pax")
    ] = 400
    """Passengers boarding each train, from the ETA report."""

    departing_pax_lead_time: Annotated[
        int, Field(name="Departing Passengers Lead Time", units="s")
    ] = 120
    """
    How long before its scheduled arrival a train's departing passengers
    start coming down to the platform, all at once, like when its track is announced.
    Until then, they all wait in the concourse, off the platform and the stairs,
    with none arriving later.
    Not from any source.
    """

    doors_per_car: Annotated[int, Field(name="Doors per Car", units="door")] = 4
    """
    Doors (single-door equivalents) on each car on the platform side.
    An NJ Transit MultiLevel, the worst case.
    An LIRR car has more and better doors.
    """

    car_length: Annotated[float, Field(name="Car Length", units="ft")] = 85
    """
    Length of each car, over which its doors are spread evenly.
    An NJ Transit MultiLevel.
    Only used on platforms with each VCE's position.
    """

    walking_speed: Annotated[float, Field(name="Walking Speed", units="ft/s")] = (
        250 / SECONDS_PER_MINUTE
    )
    """
    Speed arriving passengers walk from the doors to the VCEs.
    The TCQSM's design walking speed, 250 ft/min
    (p. 10-20: https://onlinepubs.trb.org/onlinepubs/tcrp/tcrp_rpt_165ch-10.pdf#page=24),
    though people walk slower in crowds, with less than 25 ft^2/pax (Exhibit 10-10, p. 10-21).
    Only used on platforms with each VCE's position.
    """

    escalator_reversal_threshold: Annotated[
        float, Field(name="Escalator Reversal Threshold", units="fraction")
    ] = 0.1
    """
    Extra escalators, i.e. beyond one going up and one going down,
    go up while the platform is alighting,
    and reverse to go down once fewer than this fraction of `Params.arriving_pax_per_train`
    are still on the platform or aboard trains that have arrived,
    i.e. the platform is nearly fully alighted.
    Not from any source.
    """

    vce_choice: Annotated[
        Literal["nearest", "quickest"], Field(name="VCE Choice", units="rule")
    ] = "quickest"
    """
    Which VCE each arriving passenger walks to:
    the `nearest`, or the `quickest` to get up,
    i.e. with the least walking time plus waiting time for everyone queued or walking there.
    Only used on platforms with each VCE's position.
    """

    car_full_fraction: Annotated[float, Field(name="Car Full Fraction", units="fraction")] = 0.9
    """
    Departing passengers go to the nearest car
    unless it has at least this fraction of `seats_per_car` boarded, waiting, or walking to it,
    in which case they go to the nearest car that doesn't.
    Not from any source.
    Only used on platforms with each VCE's position.
    """

    door_flow_rate: Annotated[float, Field(name="Door Flow Rate", units="pax/s/door")] = 1.0
    """
    Alighting and boarding rate per single-door equivalent.
    Assumes no delay for the doors to open, and the same rate both ways.
    From the ETA report.
    """

    stair_capacity: Annotated[float, Field(name="Stair Capacity", units="pax/min/ft")] = 17
    """
    Stair capacity, the LOS E/F boundary.
    Applied to all VCEs, even escalators.
    Fruin, p. 14: https://onlinepubs.trb.org/Onlinepubs/hrr/1971/355/355-001.pdf#page=14
    """

    stair_queue_space: Annotated[float, Field(name="Stair Queue Space", units="ft^2/pax")] = 5
    """
    Space per passenger queued at the stairs.
    Only used to report when the arriving passengers fit in the stair queues (the taper time).
    TCQSM p. 10-51: https://onlinepubs.trb.org/onlinepubs/tcrp/tcrp_rpt_165ch-10.pdf#page=55
    """

    stair_queue_length: Annotated[float, Field(name="Stair Queue Length", units="ft")] = 20
    """
    Length of the queue in front of each stair,
    used to report when the arriving passengers start to taper off.
    From the ETA report.
    """

    bidirectional_stair_flow_limit: Annotated[
        float, Field(name="Bidirectional Stair Flow Limit", units="pax/min/ft")
    ] = 10
    """
    Nobody comes down while the upward flow exceeds this.
    The ETA report says there's no bidirectional flow on stairs worse than LOS C,
    i.e. above the LOS C/D boundary, 10 pax/min/ft.
    Otherwise, both directions share `stair_capacity`.
    """

    emergency_stair_flow: Annotated[
        float, Field(name="Emergency Stair Flow", units="pax/min/ft")
    ] = 19
    """
    Stair flow used only for the emergency egress time.
    Fruin's maximum ascending stair flow is 18.9
    (p. 9: https://onlinepubs.trb.org/Onlinepubs/hrr/1971/355/355-001.pdf#page=9),
    more than both `stair_capacity` and NFPA 130's 16.9
    (TCQSM p. 10-79: https://onlinepubs.trb.org/onlinepubs/tcrp/tcrp_rpt_165ch-10.pdf#page=83),
    so the emergency egress time is a lower bound.
    """

    platform_los_min_space: tuple[tuple[str, float], ...] = (
        ("A", 13),
        ("B", 10),
        ("C", 7),
        ("D", 3),
        ("E", 2),
    )
    """
    Fruin's LOS for queuing and waiting areas:
    each grade needs more than this space per passenger (ft^2/pax), or else F.
    Most passengers on a platform are standing and waiting, not walking,
    so the TCQSM grades platforms with these, not Fruin's LOS for walkways.
    TCQSM, Exhibit 10-32, p. 10-55: https://onlinepubs.trb.org/onlinepubs/tcrp/tcrp_rpt_165ch-10.pdf#page=59
    """

    stair_los_max_flow: tuple[tuple[str, float], ...] = (
        ("A", 5),
        ("B", 7),
        ("C", 10),
        ("D", 13),
    )
    """
    Fruin's LOS for stairs:
    each grade allows at most this flow (pax/min per ft of width),
    then E up to `stair_capacity`, or else F.
    Fruin, pp. 12-14: https://onlinepubs.trb.org/Onlinepubs/hrr/1971/355/355-001.pdf#page=12
    """


# basic flow: train egress > platform crowd > VCE egress rate > back to
# platform crowd


def stair_flow(rate: float, w: float) -> float:
    """
    :param rate: stair flow per foot of width (pax/min/ft)
    :param w: total stair width (ft)
    :return: total stair flow (pax/s)
    """
    return rate * w / SECONDS_PER_MINUTE


def alight_rate(k: float, t: float, t0: float | None, u: float) -> float:
    """
    :param k: number of people waiting to get off train
    :param t: time pass counter (s)
    :param t0: train arrival time, or `None` if it hasn't been scheduled yet
    :param u: maximum alighting rate across all doors (pax/s)
    :return: egress rate from train to platform across all doors (pax/s)
    """
    if t0 is not None and t > t0:
        return min(k, u)
    else:
        return 0


def platform_clearance(karr: float, w: float, assumptions: Assumptions) -> float:
    """
    Arriving passengers queue at each stair,
    which discharges them at `Assumptions.stair_capacity` as long as anyone is queued.

    Fruin's stair equation relates flow to the space per passenger *on the stair*,
    which a queued stair holds near its critical density,
    so it doesn't apply to the space per passenger on the platform.

    :param karr: number of arriving passengers queued at this VCE (pax)
    :param w: this VCE's width (ft)
    :return: this VCE's upward flow (pax/s)
    """
    return min(karr, stair_flow(assumptions.stair_capacity, w))


def platform_ingress(
    kdep: float,
    widths: list[float],
    up_rates: list[float],
    directions: list[Direction],
    assumptions: Assumptions,
) -> list[float]:
    """
    Departing passengers queue upstairs and come down each VCE with whatever capacity
    its upward flow leaves, unless that exceeds `Assumptions.bidirectional_stair_flow_limit`.

    :param kdep: number of departing passengers upstairs (pax)
    :param widths: this train's share of each VCE's width (ft)
    :param up_rates: upward flow on this train's share of each VCE (pax/s)
    :param directions: which way each VCE runs now
    :return: platform ingress rate down each VCE (pax/s)
    """
    available = [
        0
        if direction == "up" or r_up > stair_flow(assumptions.bidirectional_stair_flow_limit, width)
        else stair_flow(assumptions.stair_capacity, width) - r_up
        for width, r_up, direction in zip(widths, up_rates, directions, strict=True)
    ]
    total = sum(available)
    if total <= 0:
        return [0 for _ in available]
    # If fewer are upstairs than can come down, they spread across the VCEs by capacity.
    return [rate * min(1, kdep / total) for rate in available]


def boarder_fraction(train_boarders: float, all_boarders: list[float]) -> float:
    """
    :param train_boarders: one train's departing passengers upstairs (pax)
    :param all_boarders: every train's departing passengers upstairs (pax)
    :return: that train's share of them
    """
    total = sum(all_boarders)
    if total > 0:
        return train_boarders / total
    else:
        return 1


def board_rate(
    r_max: float,
    r_off: float,
    sim_t: float,
    arr_t: float | None,
    boarders: float,
) -> float:
    """
    Nobody boards until everyone has alighted, per the ETA report,
    i.e. the second after `r_off` is last nonzero.

    :param r_max: maximum boarding rate across all doors (pax/s)
    :param r_off: train alight rate (pax/s)
    :param sim_t: time (s)
    :param arr_t: train arrival time (s), or `None` if it hasn't been scheduled yet
    :param boarders: number of passengers waiting on platform to board
    :return: train ingress rate across all doors (pax/s)
    """
    if arr_t is not None and arr_t < sim_t and r_off == 0:
        return min(r_max, boarders)
    else:
        return 0


def space_per_pax(k: float, a: float) -> float:
    """
    :param k: people on platform (pax)
    :param a: usable platform area (ft^2)
    :return: space per passenger (ft^2/pax)
    """
    if k > 0:
        return a / k
    else:
        return a


def platform_crowd_los(inst_crowding: float, assumptions: Assumptions) -> str:
    """
    :param inst_crowding: space per passenger on the platform (ft^2/pax)
    :return: its LOS, per `Assumptions.platform_los_min_space`
    """
    for grade, min_space in assumptions.platform_los_min_space:
        if inst_crowding > min_space:
            return grade
    return "F"


def egress_crowd_los(w: float, plat_egress_rate: float, assumptions: Assumptions) -> str:
    """
    :param w: VCE width (ft)
    :param plat_egress_rate: upward flow on it (pax/s)
    :return: its LOS, per `Assumptions.stair_los_max_flow` and `stair_capacity`
    """
    for grade, max_flow in (*assumptions.stair_los_max_flow, ("E", assumptions.stair_capacity)):
        if plat_egress_rate <= stair_flow(max_flow, w):
            return grade
    return "F"


def worst_egress_los(vces: tuple[Vce, ...], up_rates: list[float], assumptions: Assumptions) -> str:
    """The worst of each VCE's `egress_crowd_los`."""
    return max(
        egress_crowd_los(vce.width, r_up, assumptions)
        for vce, r_up in zip(vces, up_rates, strict=True)
    )


@dataclass
class Field:
    """A type annotation for a field."""

    name: str
    """The human-readable name of the field."""

    units: str
    """The field's units."""

    @property
    def description(self) -> str:
        return f"{self.name} ({self.units})"

    @classmethod
    def try_from_annotated(cls, annotated_type: Any) -> Self | None:
        if typing.get_origin(annotated_type) is not Annotated:
            return None
        (self,) = cast(type[Annotated[Any, Any]], annotated_type).__metadata__
        assert type(self) is cls, f"{type(self)} supposed to be f{cls}"
        return self


def annotated_field_names(
    cls: type[DataclassInstance],
) -> Generator[tuple[str, Field]]:
    """
    Iterate over the fields and `@property` methods of a `@dataclass` type
    and return the attr name and `Field` for each `Annotated[T, Field]` type.
    """

    # Resolve the annotations with `get_type_hints`, not `field.type`,
    # since some, like `Assumptions`', refer to `Field` before it's defined.
    type_hints = typing.get_type_hints(cls, include_extras=True)
    for field in dataclasses.fields(cls):
        field_meta = Field.try_from_annotated(type_hints[field.name])
        if field_meta is None:
            continue
        yield field.name, field_meta

    for attr_name in dir(cls):
        attr = getattr(cls, attr_name)
        if not isinstance(attr, property):
            continue
        prop = attr
        signature = typing.get_type_hints(prop.fget, include_extras=True)
        return_type = signature["return"]
        field_meta = Field.try_from_annotated(return_type)
        if field_meta is None:
            continue
        yield attr_name, field_meta


def annotated_field_values(
    obj: DataclassInstance,
) -> Generator[tuple[str, Any, Field]]:
    """
    Iterate over the fields and `@property` methods of a `@dataclass` object
    and return the attr name, value, and `Field` for each `Annotated[T, Field]` type.
    """

    for attr, field in annotated_field_names(obj.__class__):
        yield attr, getattr(obj, attr), field


@dataclass(frozen=True)
class Vce:
    """A VCE (vertical circulation element), i.e. a stair or escalator going upstairs."""

    name: str
    """Its name, e.g. `P3-S4`."""

    width: float
    """Its width (in feet)."""

    type: str = "stair"
    """`stair` or `escalator`."""

    west_end: float | None = None
    """Where it starts along the platform (ft east of the Master Plan's plans' west edge)."""

    east_end: float | None = None
    """Where it ends along the platform (ft east of the Master Plan's plans' west edge)."""


VCE_DATA = Path(__file__).parent / "data" / "estimated_vce_widths.csv"


@functools.cache
def platform_vces(platform: int) -> tuple[Vce, ...]:
    """
    Every VCE on `platform`, from `data/estimated_vce_widths.csv`,
    with the Master Plan's width where it has one, or else the estimated width.
    """
    with VCE_DATA.open() as f:
        return tuple(
            Vce(
                name=row["vce"],
                width=float(row["master_plan_width_in"] or row["sheet_width_in"]) / 12,
                type=row["type"],
                west_end=float(row["west_end_ft"]),
                east_end=float(row["east_end_ft"]),
            )
            for row in csv.DictReader(f)
            if int(row["platform"]) == platform
        )


def door_positions(params: Params) -> list[float]:
    """
    Where each of a train's doors is along the platform (ft),
    spread evenly along a train stopped at `Params.train_east_end`.
    """
    assert params.platform_east_end is not None
    assumptions = params.assumptions
    train_east_end = params.train_east_end or params.platform_east_end
    train_west_end = train_east_end - params.train_length
    door_spacing = assumptions.car_length / assumptions.doors_per_car
    return [train_west_end + (door + 0.5) * door_spacing for door in range(params.doors_per_train)]


def distance_to(vce: Vce, position: float) -> float:
    """Distance (ft) along the platform from `position` to the nearest end of `vce`."""
    assert vce.west_end is not None and vce.east_end is not None
    return max(0, vce.west_end - position, position - vce.east_end)


@dataclass(frozen=True)
class Door:
    """Where some of the arriving passengers come from."""

    share: float
    """Fraction of the arriving passengers."""

    walking_times: list[int]
    """Time (s) they take to walk to each VCE."""


def doors_to_vces(params: Params) -> list[Door]:
    """
    Where arriving passengers come from, and how far they are from each VCE.
    If the VCEs have no positions, the platform is one door
    with no walking time to any VCE.
    """
    if params.platform_east_end is None:
        return [Door(share=1, walking_times=[0 for _ in params.vces])]
    doors = door_positions(params)
    return [
        Door(
            share=1 / len(doors),
            walking_times=[
                round(distance_to(vce, door) / params.assumptions.walking_speed)
                for vce in params.vces
            ],
        )
        for door in doors
    ]


type Direction = Literal["both", "up", "down"]
"""Which way a VCE runs: stairs go `both` ways, and escalators go `up` or `down`."""

type Role = Literal["stair", "up", "down", "reversible"]
"""What a VCE does: a stair, an escalator always going `up` or `down`, or one that reverses."""


def vce_roles(vces: tuple[Vce, ...]) -> list[Role]:
    """
    What each VCE does.
    A platform's only escalator goes up.
    With more, the easternmost, toward 7th Avenue, goes up,
    the westernmost, toward 8th Avenue, goes down,
    and the rest are `reversible`, going up while the platform is alighting.
    This matches the AM peak, when most passengers are heading toward 7th Avenue.
    """
    escalators = sorted(
        (i for i, vce in enumerate(vces) if vce.type == "escalator"),
        key=lambda i: vces[i].west_end or 0,
    )
    roles: list[Role] = ["stair" for _ in vces]
    for i in escalators:
        roles[i] = "reversible"
    if escalators:
        roles[escalators[-1]] = "up"
    if len(escalators) > 1:
        roles[escalators[0]] = "down"
    return roles


@dataclass(frozen=True)
class Car:
    """One of a train's cars, where departing passengers wait to board."""

    doors: int
    """Doors (single-door equivalents) on the platform side."""

    walking_times: list[int]
    """Time (s) to walk from each VCE to its nearest door."""


def train_cars(params: Params) -> list[Car]:
    """
    Each train's cars.
    If the VCEs have no positions, the train is one car
    with no walking time from any VCE.
    """
    assumptions = params.assumptions
    if params.platform_east_end is None:
        return [Car(doors=params.doors_per_train, walking_times=[0 for _ in params.vces])]
    doors = door_positions(params)
    doors_per_car = assumptions.doors_per_car
    cars: list[Car] = []
    for car in range(params.cars):
        car_doors = doors[car * doors_per_car : (car + 1) * doors_per_car]
        cars.append(
            Car(
                doors=len(car_doors),
                walking_times=[
                    round(
                        min(distance_to(vce, door) for door in car_doors)
                        / assumptions.walking_speed
                    )
                    for vce in params.vces
                ],
            )
        )
    return cars


def cars_by_distance(cars: list[Car], vce: int) -> list[int]:
    """`cars`' indices, from the nearest to `vce` to the farthest, by walking time."""
    return sorted(range(len(cars)), key=lambda c: cars[c].walking_times[vce])


def choose_car(params: Params, nearest: list[int], car_loads: list[float]) -> int:
    """
    The car departing passengers coming down a VCE walk to:
    the nearest one that isn't close to full, per `Assumptions.car_full_fraction`,
    or else the nearest one.

    :param nearest: the cars, from the nearest to the VCE to the farthest, per `cars_by_distance`
    :param car_loads: passengers boarded on, waiting for, or walking to each car (pax)
    """
    if len(nearest) == 1:
        return nearest[0]
    full = params.assumptions.car_full_fraction * params.assumptions.seats_per_car
    return next((c for c in nearest if car_loads[c] < full), nearest[0])


def choose_vce(
    params: Params,
    door: Door,
    vce_queues: list[float],
    walking_to: list[float],
    directions: list[Direction],
    capacities: list[float],
) -> int:
    """
    The VCE the passengers from `door` walk to, per `Assumptions.vce_choice`,
    of those going up.

    :param vce_queues: arriving passengers queued at each VCE (pax)
    :param walking_to: arriving passengers walking to each VCE (pax)
    :param directions: which way each VCE runs now
    :param capacities: each VCE's upward capacity (pax/s)
    """
    going_up = [i for i, direction in enumerate(directions) if direction != "down"]

    def time_to_go_up(i: int) -> float:
        return door.walking_times[i] + (vce_queues[i] + walking_to[i]) / capacities[i]

    if params.assumptions.vce_choice == "nearest":
        return min(going_up, key=lambda i: door.walking_times[i])
    return min(going_up, key=time_to_go_up)


def pooled_vces(total_width: float) -> tuple[Vce, ...]:
    """
    A single VCE standing in for all of a platform's VCEs, as one pooled queue.
    Per the ETA report, `total_width` excludes one VCE per platform,
    e.g. an escalator running the other way.
    """
    return (Vce(name="All VCEs", width=total_width),)


@dataclass
class Params:
    platform: Annotated[int, Field(name="Platform", units="#")]
    """Which platform it is, e.g. 3."""

    platform_width: Annotated[int, Field(name="Platform Width", units="ft")]
    """Platform width (in feet)."""

    platform_length: Annotated[int, Field(name="Platform Length", units="ft")]
    """Platform length (in feet)."""

    platform_max_cars: Annotated[int, Field(name="Platform Max Cars", units="car")]
    """
    Cars in the longest train that fits on the platform's tracks, from
    https://www.railfanguides.us/ny/penntonewrochelle/PennStationLayout1.jpg
    """

    headway: Annotated[int, Field(name="Headway", units="s")]
    """
    Time (in seconds) between trains' scheduled arrivals.
    The first arrives at 0 s, on one track, and the second on the other.
    """

    vces: tuple[Vce, ...]
    """The VCEs (vertical circulation elements) going upstairs."""

    trains: Annotated[int, Field(name="Trains", units="train")] = 4
    """Trains arriving, alternating between the platform's two tracks."""

    platform_east_end: float | None = None
    """
    Where the platform ends to the east (ft east of the Master Plan's plans' west edge),
    if `vces` have positions.
    From `MASTER_PLAN_EAST_END_FT` in `scripts/estimate_vce_widths.py`.
    """

    train_east_end: Annotated[float | None, Field(name="Train East End", units="ft")] = None
    """
    Where the trains' east ends stop (ft east of the Master Plan's plans' west edge),
    if `vces` have positions,
    or else `platform_east_end`.
    `best_stopping_position` finds the best one.
    """

    assumptions: Assumptions = dataclasses.field(default_factory=Assumptions)
    """What the model assumes, the same for every scenario unless overridden."""

    modifier: str | None = None
    """
    What sets this scenario apart from the platform's others,
    e.g. `recon` for its VCEs after Penn Reconstruction.
    """

    @property
    def name(self) -> str:
        """The platform's name in the results table, e.g. `3 (recon)`."""
        return f"{self.platform} ({self.modifier})" if self.modifier else str(self.platform)

    @property
    def filename_prefix(self) -> str:
        """Prefix of the filename to save the spreadsheet in, e.g. `platform3_recon`."""
        return f"platform{self.platform}" + (f"_{self.modifier}" if self.modifier else "")

    @property
    def cars(self) -> Annotated[int, Field(name="Cars per Train", units="car")]:
        """Cars in each train: as many as fit on the platform, up to `max_train_cars`."""
        return min(self.platform_max_cars, self.assumptions.max_train_cars)

    @property
    def arriving_pax_per_train(
        self,
    ) -> Annotated[int, Field(name="Arriving Passengers per Train", units="pax")]:
        """Passengers arriving on each train, all of whom alight."""
        return self.cars * self.assumptions.seats_per_car

    @property
    def doors_per_train(self) -> Annotated[int, Field(name="Doors per Train", units="door")]:
        """Doors (single-door equivalents) on each train on the platform side."""
        return self.cars * self.assumptions.doors_per_car

    @property
    def train_length(self) -> Annotated[float, Field(name="Train Length", units="ft")]:
        """Length of each train."""
        return self.cars * self.assumptions.car_length

    @property
    def total_vce_width(self) -> Annotated[float, Field(name="Total VCE Width", units="ft")]:
        """Total width (in feet) of all of the VCEs."""
        return sum(vce.width for vce in self.vces)

    @property
    def los_f_egress_rate(
        self,
    ) -> Annotated[float, Field(name="LOS F Egress Rate", units="pax/s")]:
        """LOS (level of service) F egress rate (in pax/s)."""
        return stair_flow(self.assumptions.emergency_stair_flow, self.total_vce_width)

    @property
    def emergency_egress_time(
        self,
    ) -> Annotated[float, Field(name="Emergency Egress Time", units="s")]:
        """
        Time (in seconds) for everyone on both trains to go upstairs
        at `Assumptions.emergency_stair_flow`.
        """
        return 2 * self.arriving_pax_per_train / self.los_f_egress_rate


def writable_cell(sheet: Worksheet, row: int, column: int) -> Cell:
    """
    Like `sheet.cell`, but not a `MergedCell`, whose `value` is read-only.
    We never merge cells, so this always holds.
    """
    cell = sheet.cell(row=row, column=column)
    assert isinstance(cell, Cell)
    return cell


def active_worksheet(wb: openpyxl.Workbook) -> Worksheet:
    """
    `wb.active`, as a plain `Worksheet`.

    The stubs type `wb.active` as a fake subclass of both `Chartsheet` and `Worksheet`.
    Narrowing it with `type(active) is Worksheet` makes type checkers treat everything after as
    unreachable, and narrowing with `isinstance` resolves methods like `add_chart` to
    `Chartsheet`'s. Returning it as `Worksheet` avoids both.
    """
    active = wb.active
    assert isinstance(active, Worksheet)
    return active


@dataclass
class Instant:
    """
    An instant in the simulation.
    """

    time: Annotated[int, Field(name="Time", units="s")]
    """Time (in seconds)."""

    off_rate: Annotated[float, Field(name="Alighting Rate", units="pax/s")]
    """Alighting rate from every train (in pax/s)."""

    on_rate: Annotated[float, Field(name="Boarding Rate", units="pax/s")]
    """Boarding rate onto every train (in pax/s)."""

    down_rate: Annotated[float, Field(name="Downstairs Rate", units="pax/s")]
    """Downstairs rate (in pax/s)."""

    up_rate: Annotated[float, Field(name="Upstairs Rate", units="pax/s")]
    """Upstairs rate (in pax/s)."""

    departing_pax_on_platform: Annotated[
        float, Field(name="Departing Passengers on Platform", units="pax")
    ]
    """Departing passengers on the platform for every train."""

    arriving_pax_waiting_on_platform: Annotated[
        float, Field(name="Arriving Passengers on Platform", units="pax")
    ]
    """Number of arriving passengers on the platform."""

    total_pax_on_platform: Annotated[float, Field(name="Total Passengers on Platform", units="pax")]
    """Total number of passengers on platform."""

    platform_crowding: Annotated[float, Field(name="Platform Space per Passenger", units="ft^2")]
    """Platform space per passenger (in square feet)."""

    net_pax_flow_rate: Annotated[float, Field(name="Net Platform Flow Rate", units="pax/s")]
    """Net platform flow rate."""

    platform_crowd_los: Annotated[str, Field(name="Platform Crowding LOS", units="LOS")]
    """Platform crowding LOS (level of service)"""

    egress_los: Annotated[str, Field(name="Egress LOS", units="LOS")]
    """Egress LOS (level of service)."""


@dataclass
class Summary:
    """Headline results of one model run, for the results table in the README."""

    max_up_rate: float
    """Highest upstairs rate (pax/s)."""

    secs_at_capacity: int
    """Seconds the upstairs rate is at the LOS E capacity (17 pax/min/ft) of the VCEs going up."""

    taper_time: int | None
    """
    Last second the arriving passengers on the platform exceed what fits in the stair queues,
    i.e. when they start to taper off, or `None` if they never do.
    """

    clear_time: int | None
    """First second after the last arrival when all arriving passengers have left the platform."""

    arrival_times: list[int | None]
    """When each train arrives (s), or `None` if it doesn't within the simulation."""

    dwells: list[int | None]
    """
    Each train's dwell (s):
    from its arrival until all of its arriving passengers have alighted
    and all of its departing passengers have boarded,
    or `None` if that doesn't happen within the simulation.
    """

    boarded_time: int | None
    """First second when all departing passengers have boarded, or `None` if they never do."""

    max_pax_on_platform: float
    """Most passengers on the platform at once."""

    min_space_per_pax: float
    """Least platform space per passenger (sq ft)."""


def calc_workbook(
    params: Params, write_workbook: bool = True, print_time_series: bool = True
) -> tuple[openpyxl.Workbook, Summary]:
    """
    Simulate `params`, returning its spreadsheet and its results table's summary.
    Without `write_workbook`, the spreadsheet is left without its time series and charts,
    and without `print_time_series`, nothing is printed,
    e.g. when only the summary is needed.
    """
    assumptions = params.assumptions
    eff_area = (
        params.platform_width * params.platform_length * assumptions.usable_platform_area_multiplier
    )
    door_rate = params.doors_per_train * assumptions.door_flow_rate

    # Initialize counters
    arriving_pax_waiting_on_plat: float = 0
    vce_queues = [0.0 for _ in params.vces]
    """Arriving passengers queued at each VCE."""
    doors = doors_to_vces(params)
    vce_capacities = [stair_flow(assumptions.stair_capacity, vce.width) for vce in params.vces]
    """Each VCE's upward capacity (pax/s), for choosing VCEs."""
    roles = vce_roles(params.vces)
    directions: list[Direction] = [
        "both" if role == "stair" else "down" if role == "down" else "up" for role in roles
    ]
    """Which way each VCE runs now."""
    walking: defaultdict[int, list[float]] = defaultdict(lambda: [0.0 for _ in params.vces])
    """Arriving passengers walking to each VCE, by the time they reach its queue."""
    walking_totals = [0.0 for _ in params.vces]
    """Arriving passengers walking to each VCE, whenever they reach it, i.e. `walking` summed."""
    cars = train_cars(params)
    nearest_cars = [cars_by_distance(cars, i) for i in range(len(params.vces))]
    """Each VCE's cars, from the nearest to the farthest."""
    trains = range(params.trains)
    boarders_walking: defaultdict[int, list[list[float]]] = defaultdict(
        lambda: [[0.0 for _ in cars] for _ in trains]
    )
    """Each train's departing passengers walking to each car, by the time they reach it."""
    boarders_walking_totals = [0.0 for _ in trains]
    """Each train's departing passengers walking to any of its cars, `boarders_walking` summed."""
    car_waiting = [[0.0 for _ in cars] for _ in trains]
    """Each train's departing passengers waiting at each car."""
    car_loads = [list(waiting) for waiting in car_waiting]
    """Each train's passengers boarded on, waiting for, or walking to each car."""
    arrival_times: list[int | None] = [
        train * params.headway if train < 2 else None for train in trains
    ]
    """
    When each train arrives: the first two as scheduled,
    and each later one once it's scheduled and the train before it on its track has departed.
    """
    remaining_arrivals = [float(params.arriving_pax_per_train) for _ in trains]
    new_pax = [0.0 for _ in trains]
    release_times = [
        train * params.headway - assumptions.departing_pax_lead_time for train in trains
    ]
    """When each train's departing passengers start coming down to the platform."""
    start_time = min(0, *release_times)
    """
    When the simulation starts (s): before the first train arrives at 0 s,
    once its departing passengers start coming down.
    """
    boarders_upstairs = [0.0 for _ in trains]
    boarders_on_plat = [0.0 for _ in trains]
    total_pax_on_platform: float = 0
    wb = openpyxl.Workbook()
    vce_sheet = wb.create_sheet("VCEs")
    """Each VCE's queue and upward flow each second."""
    writable_cell(vce_sheet, row=1, column=1).value = "Time (s)"
    for i, vce in enumerate(params.vces):
        writable_cell(vce_sheet, row=1, column=2 + 2 * i).value = f"{vce.name} Queue (pax)"
        writable_cell(vce_sheet, row=1, column=3 + 2 * i).value = f"{vce.name} Up Rate (pax/s)"
    vce_empty_times: list[int | None] = [None for _ in params.vces]
    """When each VCE's queue last empties, i.e. its last passenger goes up."""
    trains_sheet = wb.create_sheet("Trains")
    """Each train's passengers, alighting, boarding, and departing passengers each second."""
    TRAIN_COLUMNS = [
        "Passengers (pax)",
        "Alighting Rate (pax/s)",
        "Boarding Rate (pax/s)",
        "Departing Passengers on Platform (pax)",
    ]
    if write_workbook:
        writable_cell(trains_sheet, row=1, column=1).value = "Time (s)"
        for train in trains:
            for j, column_name in enumerate(TRAIN_COLUMNS):
                writable_cell(
                    trains_sheet, row=1, column=2 + len(TRAIN_COLUMNS) * train + j
                ).value = f"Train {train + 1} {column_name}"

    sheet = active_worksheet(wb)

    writable_cell(sheet, column=1, row=1).value = "Parameter"
    writable_cell(sheet, column=2, row=1).value = "Value"

    param_values = [*annotated_field_values(params), *annotated_field_values(assumptions)]
    for i, (_attr, value, field) in enumerate(param_values):
        writable_cell(sheet, column=1, row=i + 2).value = field.description
        writable_cell(sheet, column=2, row=i + 2).value = value

    FIRST_DATA_ROW = 2

    # The parameters take up columns 1 (A) and 2 (B), so the time series starts after them.
    FIRST_DATA_COLUMN = 3

    max_pax_in_stair_queues = (
        params.total_vce_width * assumptions.stair_queue_length / assumptions.stair_queue_space
    )
    summary = Summary(
        max_up_rate=0,
        secs_at_capacity=0,
        taper_time=None,
        clear_time=None,
        arrival_times=arrival_times,
        dwells=[None for _ in trains],
        boarded_time=None,
        max_pax_on_platform=total_pax_on_platform,
        min_space_per_pax=space_per_pax(total_pax_on_platform, eff_area),
    )

    if print_time_series:
        print("Elapsed_Time", *(f"Train_{train + 1}_Pax" for train in trains))

    def get_column_for(attr_name: str) -> int:
        for i, (attr, _field) in enumerate(annotated_field_names(Instant)):
            if attr == attr_name:
                return FIRST_DATA_COLUMN + i
        raise AttributeError(Instant, attr_name)

    def walking_to_train(train: int) -> float:
        """Departing passengers walking to `train`'s cars."""
        return boarders_walking_totals[train]

    def subtract(total: float, pax: float) -> float:
        """`total - pax`, or 0 if that's only left over from rounding."""
        remaining = total - pax
        return 0.0 if remaining < 1e-9 else remaining

    def row_for(time: int) -> int:
        """The spreadsheets' row for `time`."""
        return time - start_time + FIRST_DATA_ROW

    time_after = start_time
    for time_after in itertools.count(start_time):
        if time_after - start_time >= MAX_SIMULATION_LENGTH:
            raise RuntimeError(
                f"{params.filename_prefix} hasn't finished after {MAX_SIMULATION_LENGTH} s"
            )
        for train in trains:
            if release_times[train] == time_after:
                boarders_upstairs[train] = float(assumptions.departing_pax_per_train)
        if write_workbook:
            writable_cell(vce_sheet, row=row_for(time_after), column=1).value = time_after
            writable_cell(trains_sheet, row=row_for(time_after), column=1).value = time_after
        off_rates: list[float] = []
        for train in trains:
            off_rate = alight_rate(
                remaining_arrivals[train],
                time_after,
                arrival_times[train],
                door_rate,
            )
            remaining_arrivals[train] -= off_rate
            if remaining_arrivals[train] < 0:
                remaining_arrivals[train] = 0
            off_rates.append(off_rate)
        total_pax_on_platform += sum(off_rates)
        arriving_pax_waiting_on_plat += sum(off_rates)
        walking_to = list(walking_totals)
        still_alighting = arriving_pax_waiting_on_plat
        for train in trains:
            arrival_time = arrival_times[train]
            if arrival_time is not None and time_after >= arrival_time:
                still_alighting += remaining_arrivals[train]
        nearly_alighted = (
            still_alighting
            < assumptions.escalator_reversal_threshold * params.arriving_pax_per_train
        )
        for i, role in enumerate(roles):
            if role == "reversible" and (not nearly_alighted or vce_queues[i] + walking_to[i] > 0):
                # Once reversed, nobody walks to it, so it stays down until more alight.
                directions[i] = "up"
            elif role == "reversible":
                directions[i] = "down"
        # Nobody walks to a VCE in a second nobody alights.
        for door in doors if sum(off_rates) > 0 else ():
            alighting = sum(off_rates) * door.share
            if params.platform_east_end is None:
                # Spread across the VCEs in proportion to their widths.
                for i, vce in enumerate(params.vces):
                    walking[time_after][i] += alighting * vce.width / params.total_vce_width
                    walking_totals[i] += alighting * vce.width / params.total_vce_width
                continue
            i = choose_vce(params, door, vce_queues, walking_to, directions, vce_capacities)
            walking[time_after + door.walking_times[i]][i] += alighting
            walking_totals[i] += alighting
            walking_to[i] += alighting
        for i, reaching in enumerate(walking.pop(time_after, [])):
            vce_queues[i] += reaching
            walking_totals[i] = subtract(walking_totals[i], reaching)
        vce_up_rates = [
            platform_clearance(queue, vce.width, assumptions)
            for vce, queue in zip(params.vces, vce_queues, strict=True)
        ]
        for i, r_up in enumerate(vce_up_rates):
            vce_queues[i] = max(0, vce_queues[i] - r_up)
        plat_egress_rate = sum(vce_up_rates)
        arriving_pax_waiting_on_plat -= plat_egress_rate
        if arriving_pax_waiting_on_plat < 0:
            arriving_pax_waiting_on_plat = 0
        total_pax_on_platform -= plat_egress_rate
        # Each train's boarders get a share of each VCE,
        # and so a share of the upward flow on it.
        boarder_fracs = [
            boarder_fraction(boarders_upstairs[train], boarders_upstairs) for train in trains
        ]
        vce_ingress_rates = [
            # Nobody comes down for a train with nobody upstairs.
            platform_ingress(
                boarders_upstairs[train],
                [vce.width * boarder_fracs[train] for vce in params.vces],
                [r_up * boarder_fracs[train] for r_up in vce_up_rates],
                directions,
                assumptions,
            )
            if boarders_upstairs[train] > 0
            else [0.0 for _ in params.vces]
            for train in trains
        ]
        plat_ingress_rates = [sum(rates) for rates in vce_ingress_rates]
        # Departing passengers walk from their VCE to a car.
        for train in trains:
            for i, rate in enumerate(vce_ingress_rates[train]):
                if rate <= 0:
                    continue
                car = choose_car(params, nearest_cars[i], car_loads[train])
                boarders_walking[time_after + cars[car].walking_times[i]][train][car] += rate
                boarders_walking_totals[train] += rate
                car_loads[train][car] += rate
        for train, reaching in enumerate(
            boarders_walking.pop(time_after, [[0.0 for _ in cars] for _ in trains])
        ):
            for car, pax in enumerate(reaching):
                car_waiting[train][car] += pax
                boarders_walking_totals[train] = subtract(boarders_walking_totals[train], pax)
        for train in trains:
            boarders_on_plat[train] = sum(car_waiting[train])
            total_pax_on_platform += plat_ingress_rates[train]
        # Each car boards through its own doors.
        car_on_rates = [
            [
                board_rate(
                    car.doors * assumptions.door_flow_rate,
                    off_rates[train],
                    time_after,
                    arrival_times[train],
                    waiting,
                )
                for car, waiting in zip(cars, car_waiting[train], strict=True)
            ]
            for train in trains
        ]
        for train, on_rates in enumerate(car_on_rates):
            for car, on_rate in enumerate(on_rates):
                car_waiting[train][car] -= on_rate
        on_rates = [sum(rates) for rates in car_on_rates]

        for train in trains:
            boarders_on_plat[train] -= on_rates[train]
            total_pax_on_platform -= on_rates[train]
            boarders_upstairs[train] -= plat_ingress_rates[train]
            new_pax[train] += on_rates[train]

        inst_crowding = space_per_pax(total_pax_on_platform, eff_area)
        if total_pax_on_platform < 0:
            total_pax_on_platform = 0
        for train in trains:
            if boarders_on_plat[train] < 0:
                boarders_on_plat[train] = 0
        if arriving_pax_waiting_on_plat < 0:
            arriving_pax_waiting_on_plat = 0
        if print_time_series:
            print(
                time_after,
                *(remaining_arrivals[train] + new_pax[train] for train in trains),
                arriving_pax_waiting_on_plat,
                plat_egress_rate,
            )
        summary.max_up_rate = max(summary.max_up_rate, plat_egress_rate)
        # Capacity of the VCEs going up now, not counting escalators going down.
        capacity = sum(
            stair_flow(assumptions.stair_capacity, vce.width)
            for vce, direction in zip(params.vces, directions, strict=True)
            if direction != "down"
        )
        if plat_egress_rate >= capacity - 1e-9:
            summary.secs_at_capacity += 1
        if arriving_pax_waiting_on_plat > max_pax_in_stair_queues:
            summary.taper_time = time_after
        if (
            summary.clear_time is None
            and all(
                arrival_time is not None and time_after > arrival_time
                for arrival_time in arrival_times
            )
            and arriving_pax_waiting_on_plat < 1
        ):
            summary.clear_time = time_after
        if (
            summary.boarded_time is None
            and time_after >= max(release_times)
            and sum(boarders_upstairs) + sum(boarders_on_plat) + sum(boarders_walking_totals) < 1
        ):
            summary.boarded_time = time_after
        for train in trains:
            arrival_time = arrival_times[train]
            if (
                summary.dwells[train] is None
                and arrival_time is not None
                and time_after > arrival_time
                and remaining_arrivals[train] < 1
                and boarders_upstairs[train] + boarders_on_plat[train] + walking_to_train(train) < 1
            ):
                summary.dwells[train] = time_after - arrival_time
                # The next train on its track arrives once it's scheduled and this one departs.
                if train + 2 < params.trains:
                    arrival_times[train + 2] = max((train + 2) * params.headway, time_after)
        summary.max_pax_on_platform = max(summary.max_pax_on_platform, total_pax_on_platform)
        summary.min_space_per_pax = min(summary.min_space_per_pax, inst_crowding)

        for i, (queue, r_up) in enumerate(zip(vce_queues, vce_up_rates, strict=True)):
            if queue < 1e-9 and r_up > 0:
                vce_empty_times[i] = time_after

        if write_workbook:
            net_pax_flow_rate: float = 0
            for rate in plat_ingress_rates:
                net_pax_flow_rate += rate
            for rate in off_rates:
                net_pax_flow_rate += rate
            net_pax_flow_rate -= plat_egress_rate
            for rate in on_rates:
                net_pax_flow_rate -= rate
            instant = Instant(
                time=time_after,
                arriving_pax_waiting_on_platform=arriving_pax_waiting_on_plat,
                off_rate=sum(off_rates),
                on_rate=sum(on_rates),
                down_rate=sum(plat_ingress_rates),
                departing_pax_on_platform=sum(boarders_on_plat)
                + sum(walking_to_train(train) for train in trains),
                total_pax_on_platform=total_pax_on_platform,
                platform_crowding=inst_crowding,
                up_rate=plat_egress_rate,
                net_pax_flow_rate=net_pax_flow_rate,
                platform_crowd_los=platform_crowd_los(inst_crowding, assumptions),
                egress_los=worst_egress_los(params.vces, vce_up_rates, assumptions),
            )

            for train in trains:
                for j, value in enumerate(
                    (
                        remaining_arrivals[train] + new_pax[train],
                        off_rates[train],
                        on_rates[train],
                        boarders_on_plat[train] + walking_to_train(train),
                    )
                ):
                    writable_cell(
                        trains_sheet,
                        row=row_for(instant.time),
                        column=2 + len(TRAIN_COLUMNS) * train + j,
                    ).value = value

            for i, (queue, r_up) in enumerate(zip(vce_queues, vce_up_rates, strict=True)):
                writable_cell(vce_sheet, row=row_for(instant.time), column=2 + 2 * i).value = queue
                writable_cell(vce_sheet, row=row_for(instant.time), column=3 + 2 * i).value = r_up

            for i, (_attr, value, field) in enumerate(annotated_field_values(instant)):
                column = FIRST_DATA_COLUMN + i
                writable_cell(sheet, row=1, column=column).value = field.description
                writable_cell(sheet, row=row_for(instant.time), column=column).value = value

        # Stop once the last train has departed and the platform has cleared.
        if (
            summary.clear_time is not None
            and summary.boarded_time is not None
            and all(dwell is not None for dwell in summary.dwells)
        ):
            break
    if not write_workbook:
        return wb, summary
    simulation_length = time_after - start_time + 1

    def make_chart(title: str, min_col: int, x_title: str, y_title: str) -> ScatterChart:
        chart = ScatterChart()
        chart.title = title
        chart.style = 13
        chart.x_axis.title = x_title
        chart.y_axis.title = y_title
        chart.x_axis.scaling.min = start_time
        chart.x_axis.scaling.max = time_after
        chart.legend = None

        max_row = simulation_length + FIRST_DATA_ROW - 1
        xvalues = Reference(
            sheet, min_col=get_column_for("time"), min_row=FIRST_DATA_ROW, max_row=max_row
        )
        values = Reference(sheet, min_col=min_col, min_row=FIRST_DATA_ROW - 1, max_row=max_row)
        # Y values start one row above X values so that first cell is series name.
        series = SeriesFactory(values, xvalues, title_from_data=True)
        chart.series.append(series)
        return chart

    def make_chart_with_chopped_y(
        title: str, min_col: int, x_title: str, y_title: str
    ) -> ScatterChart:
        chart = ScatterChart()
        chart.title = title
        chart.style = 13
        chart.x_axis.title = x_title
        chart.y_axis.title = y_title
        chart.x_axis.scaling.min = start_time
        chart.x_axis.scaling.max = time_after
        chart.y_axis.scaling.min = 0
        chart.y_axis.scaling.max = 50
        chart.legend = None

        max_row = simulation_length + FIRST_DATA_ROW - 1
        xvalues = Reference(
            sheet, min_col=get_column_for("time"), min_row=FIRST_DATA_ROW, max_row=max_row
        )
        values = Reference(sheet, min_col=min_col, min_row=FIRST_DATA_ROW - 1, max_row=max_row)
        # Y values start one row above X values so that first cell is series name.
        series = SeriesFactory(values, xvalues, title_from_data=True)
        chart.series.append(series)
        return chart

    def make_chart_2(
        title: str,
        col1: int,
        col2: int,
        x_title: str,
        y_title: str,
        *more_cols: int,
        data_sheet: Worksheet = sheet,
        time_col: int | None = None,
    ) -> ScatterChart:
        chart = ScatterChart()
        chart.title = title
        chart.style = 13
        chart.x_axis.title = x_title
        chart.y_axis.title = y_title
        chart.x_axis.scaling.min = start_time
        chart.x_axis.scaling.max = time_after
        assert chart.legend is not None
        chart.legend.position = "b"

        max_row = simulation_length + FIRST_DATA_ROW - 1
        xvalues = Reference(
            data_sheet,
            min_col=get_column_for("time") if time_col is None else time_col,
            min_row=FIRST_DATA_ROW,
            max_row=max_row,
        )
        for col in (col1, col2, *more_cols):
            values = Reference(data_sheet, min_col=col, min_row=FIRST_DATA_ROW - 1, max_row=max_row)
            # Y values start one row above X values so that first cell is series name.
            chart.series.append(SeriesFactory(values, xvalues, title_from_data=True))
        return chart

    sheet.add_chart(
        make_chart_2(
            "Up and Down Rates",
            get_column_for("up_rate"),
            get_column_for("down_rate"),
            "Time (s)",
            "Rate (pax/s)",
        ),
        "V4",
    )
    sheet.add_chart(
        make_chart_2(
            "Passengers Aboard Trains",
            2,
            2 + len(TRAIN_COLUMNS),
            "Time (s)",
            "Passengers",
            *(2 + len(TRAIN_COLUMNS) * train for train in trains[2:]),
            data_sheet=trains_sheet,
            time_col=1,
        ),
        "V19",
    )
    sheet.add_chart(
        make_chart_2(
            "Passengers on Platform",
            get_column_for("arriving_pax_waiting_on_platform"),
            get_column_for("total_pax_on_platform"),
            "Time (s)",
            "Passengers",
        ),
        "V34",
    )
    sheet.add_chart(
        make_chart_with_chopped_y(
            "Space per Passenger",
            get_column_for("platform_crowding"),
            "Time (s)",
            "Space per passenger (sq ft)",
        ),
        "V49",
    )
    sheet.add_chart(
        make_chart(
            "Net Platform Flow Rate",
            get_column_for("net_pax_flow_rate"),
            "Time (s)",
            "Net Flow Rate (pax/s)",
        ),
        "V64",
    )
    if print_time_series:
        print(
            f"LOS F egress rate is {params.los_f_egress_rate} pax/s. "
            f"Emergency egress time is {params.emergency_egress_time} seconds."
        )
    vce_list = wb.create_sheet("VCE List")
    vce_list.append(
        [
            "VCE",
            "Type",
            "Role",
            "Width (ft)",
            "West End (ft)",
            "East End (ft)",
            "Queue Empty Time (s)",
        ]
    )
    for vce, role, empty_time in zip(params.vces, roles, vce_empty_times, strict=True):
        vce_list.append(
            [vce.name, vce.type, role, vce.width, vce.west_end, vce.east_end, empty_time]
        )

    return wb, summary


RESULTS_COLUMNS = [
    "Platform",
    "Headway",
    "VCE width",
    "Arrivals",
    "Dwell",
    "Taper time",
    "Clear time",
    "Boarded time",
    "Time at capacity",
    "Max up rate (pax/s)",
    "Max pax on platform",
    "Max density (pax/m²)",
]
RESULTS_HEADER = "| " + " | ".join(RESULTS_COLUMNS) + " |\n" + "|---" * len(RESULTS_COLUMNS) + "|"

README = Path(__file__).parent / "README.md"
RESULTS_START = "<!-- results-table:start -->"
RESULTS_END = "<!-- results-table:end -->"
"""The README's results table is between these markers, so `--update-readme` can replace it."""


def update_readme_results(table: str) -> None:
    """Replace the results table in the README with `table`."""
    readme = README.read_text()
    start = readme.index(RESULTS_START) + len(RESULTS_START)
    end = readme.index(RESULTS_END)
    README.write_text(f"{readme[:start]}\n{table}\n{readme[end:]}")


STOPPING_POSITION_STEP = 5
"""Distance (ft) between the stopping positions `best_stopping_position` tries."""

COARSE_STOPPING_POSITION_STEP = 85
"""
Distance (ft) between the stopping positions `best_stopping_position` tries first,
a car's length, before trying every `STOPPING_POSITION_STEP` around the best of them.
"""


def best_stopping_position(params: Params) -> Params:
    """
    `params` with the trains stopped where the arriving passengers clear the platform soonest,
    or if tied, where the departing passengers finish boarding soonest,
    of every position `STOPPING_POSITION_STEP` apart with the train on the platform.
    """
    if params.platform_east_end is None:
        return params
    platform_west_end = params.platform_east_end - params.platform_length
    candidates: list[Params] = []
    train_east_end = params.platform_east_end
    while train_east_end - params.train_length >= platform_west_end:
        candidates.append(dataclasses.replace(params, train_east_end=train_east_end))
        train_east_end -= STOPPING_POSITION_STEP

    def score(candidate: Params) -> tuple[float, float]:
        # Only the summary matters.
        _wb, summary = calc_workbook(candidate, write_workbook=False, print_time_series=False)
        never = float("inf")
        return (
            never if summary.clear_time is None else summary.clear_time,
            never if summary.boarded_time is None else summary.boarded_time,
        )

    scores: dict[float, tuple[float, float]] = {}

    def cached_score(candidate: Params) -> tuple[float, float]:
        assert candidate.train_east_end is not None
        if candidate.train_east_end not in scores:
            scores[candidate.train_east_end] = score(candidate)
        return scores[candidate.train_east_end]

    # Try every car length first, then every `STOPPING_POSITION_STEP` within a car length of it.
    coarse_step = COARSE_STOPPING_POSITION_STEP // STOPPING_POSITION_STEP
    coarse = min(candidates[::coarse_step], key=cached_score)
    i = candidates.index(coarse)
    return min(candidates[max(0, i - coarse_step) : i + coarse_step + 1], key=cached_score)


def run_model(params: Params, spreadsheets: bool) -> str:
    """
    Run the model, return its row of the results table,
    and with `spreadsheets`, print its time series and save its spreadsheet.
    """
    params = best_stopping_position(params)
    wb, summary = calc_workbook(
        params=params, write_workbook=spreadsheets, print_time_series=spreadsheets
    )

    headway = params.headway
    if spreadsheets:
        wb.save(
            f"{params.filename_prefix}"
            f"_{params.arriving_pax_per_train}"
            f"_{params.arriving_pax_per_train}"
            f"_{headway}s.xlsx"
        )
    wb.close()

    def fmt_time(t: int | None) -> str:
        """`t` seconds as `m:ss`."""
        if t is None:
            return "never"
        minutes, seconds = divmod(t, SECONDS_PER_MINUTE)
        return f"{minutes}:{seconds:02}"

    return (
        f"| {params.name} | {fmt_time(headway)} | {params.total_vce_width:.4g} ft"
        f" | {', '.join(fmt_time(arrival) for arrival in summary.arrival_times)}"
        f" | {', '.join(fmt_time(dwell) for dwell in summary.dwells)}"
        f" | {fmt_time(summary.taper_time)} | {fmt_time(summary.clear_time)}"
        f" | {fmt_time(summary.boarded_time)}"
        f" | {fmt_time(summary.secs_at_capacity)} | {summary.max_up_rate:.2f}"
        f" | {summary.max_pax_on_platform:.0f}"
        f" | {1 / (summary.min_space_per_pax * SQUARE_METERS_PER_SQUARE_FOOT):.2f}"
        f" ({platform_crowd_los(summary.min_space_per_pax, params.assumptions)}) |"
    )


def main(
    update_readme: Annotated[
        bool, Option(help="Replace the results table in the README with this run's.")
    ] = False,
    spreadsheets: Annotated[
        bool,
        Option(help="Also print each scenario's time series and save its spreadsheet."),
    ] = False,
) -> None:
    """Run every scenario and print a table of their results."""

    # params are labeled  with p<platform number><time in seconds>
    # recon indicates that a platform was modelled accounting for penn reconstruction plans
    params_p3120 = Params(
        platform=3,
        platform_width=18,
        platform_length=900,
        platform_max_cars=10,
        headway=CLOSE_HEADWAY,
        vces=platform_vces(3),
        platform_east_end=719,
    )
    params_p3300 = Params(
        platform=3,
        platform_width=18,
        platform_length=900,
        platform_max_cars=10,
        headway=NORMAL_HEADWAY,
        vces=platform_vces(3),
        platform_east_end=719,
    )
    params_p3recon120 = Params(
        platform=3,
        modifier="recon",
        platform_width=18,
        platform_length=900,
        platform_max_cars=10,
        headway=CLOSE_HEADWAY,
        vces=pooled_vces(44.75),
    )
    params_p3recon300 = Params(
        platform=3,
        modifier="recon",
        platform_width=18,
        platform_length=900,
        platform_max_cars=10,
        headway=NORMAL_HEADWAY,
        vces=pooled_vces(44.75),
    )
    params_p60 = Params(
        platform=6,
        platform_width=15,
        platform_length=1100,
        platform_max_cars=17,
        headway=0,
        vces=platform_vces(6),
        platform_east_end=849,
    )
    params_p10120 = Params(
        platform=10,
        platform_width=42,
        platform_length=1100,
        platform_max_cars=12,
        headway=CLOSE_HEADWAY,
        vces=pooled_vces(70.58),
    )
    params_p11120 = Params(
        platform=11,
        platform_width=18,
        platform_length=1100,
        platform_max_cars=12,
        headway=CLOSE_HEADWAY,
        vces=pooled_vces(43.58),
    )
    params_p30 = dataclasses.replace(params_p3120, headway=0)
    params_p3recon0 = dataclasses.replace(params_p3recon120, headway=0)
    params_p6120 = dataclasses.replace(params_p60, headway=CLOSE_HEADWAY)
    params_p6300 = dataclasses.replace(params_p60, headway=NORMAL_HEADWAY)
    params_p100 = dataclasses.replace(params_p10120, headway=0)
    params_p10300 = dataclasses.replace(params_p10120, headway=NORMAL_HEADWAY)
    params_p110 = dataclasses.replace(params_p11120, headway=0)
    params_p11300 = dataclasses.replace(params_p11120, headway=NORMAL_HEADWAY)
    with ProcessPoolExecutor() as executor:
        rows = list(
            executor.map(
                functools.partial(run_model, spreadsheets=spreadsheets),
                # Group the results by headway first, then by platform.
                [
                    params_p30,
                    params_p3recon0,
                    params_p60,
                    params_p100,
                    params_p110,
                    params_p3120,
                    params_p3recon120,
                    params_p6120,
                    params_p10120,
                    params_p11120,
                    params_p3300,
                    params_p3recon300,
                    params_p6300,
                    params_p10300,
                    params_p11300,
                ],
            )
        )
    table = "\n".join([RESULTS_HEADER, *rows])
    print()
    print(table)
    if update_readme:
        update_readme_results(table)


if __name__ == "__main__":
    typer.run(main)
