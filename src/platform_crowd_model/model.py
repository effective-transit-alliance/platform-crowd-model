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
from datetime import timedelta
from enum import StrEnum
from functools import cache
from math import ceil
from pathlib import Path
from typing import TYPE_CHECKING, Annotated, Any, Literal, Self, cast

from platform_crowd_model.paths import DATA_DIR

if TYPE_CHECKING:
    from _typeshed import DataclassInstance

SECONDS_PER_MINUTE = timedelta(minutes=1).total_seconds()
"""
Computed once, since converting a `timedelta` each time in `stair_flow`,
which runs several times each simulated second,
makes the simulation about 40% slower.
"""

TIME_STEP = timedelta(seconds=1)
"""
How much time each step of the simulation covers.
Every rate is per second, so this must stay 1 s.
"""

METERS_PER_FOOT = 0.3048

SQUARE_METERS_PER_SQUARE_FOOT = METERS_PER_FOOT**2

NFPA_130_EXIT_FLOW = 0.0555 * 1000 * METERS_PER_FOOT
"""
Exit capacity of stairs and stopped escalators for evacuating a platform (pax/min/ft),
0.0555 pax/min per mm of width (1.41 pax/min per inch), per NFPA 130 (2026 edition) 5.3.5.3.
"""

NFPA_130_PLATFORM_EVACUATION_TIME = timedelta(minutes=4)
"""
Time within which NFPA 130 (2026 edition) 5.3.3.1 requires a platform's occupant load,
including those on trains, to be able to evacuate it.
"""

NFPA_130_POINT_OF_SAFETY_TIME = timedelta(minutes=6)
"""
Time within which NFPA 130 (2026 edition) 5.3.3.2 requires evacuating
from the most remote point on a platform to a point of safety.
"""

NFPA_130_MAX_TRAVEL_DISTANCE = 100 / METERS_PER_FOOT
"""
Farthest (ft) NFPA 130 (2026 edition) 5.3.3.5 lets anyone on a platform be
from where an exit leaves it, 100 m (328'1", though NFPA 130 gives 325 ft).
"""

NFPA_130_PLATFORM_WALKING_SPEED = 37.8 / METERS_PER_FOOT / SECONDS_PER_MINUTE
"""
Speed (ft/s) people evacuate along a platform at, 37.8 m/min (124 fpm), per NFPA 130 5.3.4.4.
"""

NFPA_130_STAIR_VERTICAL_SPEED = 14.63 / METERS_PER_FOOT / SECONDS_PER_MINUTE
"""
Vertical speed (ft/s) people evacuate up stairs and stopped escalators at,
14.63 m/min (48 fpm), per NFPA 130 5.3.5.3.
"""

PLATFORM_TO_CONCOURSE_RISE = 16 + 9.25 / 12
"""
Height (ft) of the concourse above the platforms, 16'9 1/4",
from the existing platforms to the existing concourse ("Level A")
on the PCIP Phase 2 plan's north-south cross section (sheet A-213, November 2020, PDF page 54).
"""

CLOSE_HEADWAY = timedelta(minutes=2)
"""Time between two trains' arrivals in the closely spaced scenarios, from the ETA report."""

MAX_SIMULATION_LENGTH = timedelta(hours=2)
"""
The simulation runs until the last train departs and the platform clears,
but stops with an error if that takes longer than this, which means something's wrong.
"""

NORMAL_HEADWAY = timedelta(minutes=5)
"""Time between two trains' arrivals in the normal scenarios, from the ETA report."""


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

    max_train_overhang: Annotated[float, Field(name="Max Train Overhang", units="ft")] = 15
    """
    How far a train can extend past its platform's end,
    e.g. 12-car LIRR trains on platform 11, 1,007' long, per the Moynihan Station EA.
    Not from any source.
    """

    car_length: Annotated[float, Field(name="Car Length", units="ft")] = 85
    """
    Length of each car, over which its doors are spread evenly.
    A NJT MultiLevel.
    """

    seats_per_car: Annotated[int, Field(name="Seats per Car", units="pax")] = 135
    """
    Seats in each car, all of which are full on arrival, and all of whose passengers alight.
    A seated NJT car,
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
        timedelta, Field(name="Departing Passengers Lead Time", units="s")
    ] = timedelta(minutes=2)
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
    A NJT MultiLevel's, the worst case.
    A LIRR car has more and better doors.
    """

    walking_speed: Annotated[float, Field(name="Walking Speed", units="ft/s")] = (
        250 / SECONDS_PER_MINUTE
    )
    """
    Speed arriving passengers walk from the doors to the VCEs.
    The TCQSM's design walking speed, 250 ft/min
    (p. 10-20: https://onlinepubs.trb.org/onlinepubs/tcrp/tcrp_rpt_165ch-10.pdf#page=24),
    though people walk slower in crowds, with less than 25 ft^2/pax (Exhibit 10-10, p. 10-21).
    """

    vce_choice: Annotated[
        Literal["nearest", "quickest"], Field(name="VCE Choice", units="rule")
    ] = "quickest"
    """
    Which VCE each arriving passenger walks to:
    the `nearest`, or the `quickest` to get up,
    i.e. with the least walking time plus waiting time for everyone queued or walking there.
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


def fmt_ft_in(ft: float) -> str:
    """
    A length in feet as feet and inches, to the nearest inch, e.g. `6'2"` for 6.17 ft,
    leaving out whichever is 0, e.g. `18'` or `6"`.
    """
    feet, inches = divmod(round(ft * 12), 12)
    if inches == 0 and feet != 0:
        return f"{feet}'"
    if feet == 0:
        return f'{inches}"'
    return f"{feet}'{inches}\""


def stair_flow(flow_per_width: float, width: float) -> float:
    """
    :param flow_per_width: stair flow per foot of width (pax/min/ft)
    :param width: total stair width (ft)
    :return: total stair flow (pax/s)
    """
    return flow_per_width * width / SECONDS_PER_MINUTE


def alight_rate(
    pax_aboard: float, time: timedelta, arrival_time: timedelta | None, max_rate: float
) -> float:
    """
    :param pax_aboard: number of people waiting to get off train
    :param time: time pass counter
    :param arrival_time: train arrival time, or `None` if it hasn't been scheduled yet
    :param max_rate: maximum alighting rate across all doors (pax/s)
    :return: egress rate from train to platform across all doors (pax/s)
    """
    if arrival_time is not None and time > arrival_time:
        return min(pax_aboard, max_rate)
    else:
        return 0


def vce_capacity(vce: Vce, assumptions: Assumptions) -> float:
    """
    `vce`'s capacity (pax/s) in one direction:
    `Assumptions.stair_capacity`, even for an escalator.
    """
    return stair_flow(assumptions.stair_capacity, vce.width)


def platform_clearance(arriving_pax_on_platform: float, capacity: float) -> float:
    """
    Arriving passengers queue at each VCE,
    which discharges them at its capacity as long as anyone is queued.

    Fruin's stair equation relates flow to the space per passenger *on the stair*,
    which a queued stair holds near its critical density,
    so it doesn't apply to the space per passenger on the platform.
    The few seconds of walking from the doors to the stairs are ignored.

    :param arriving_pax_on_platform: number of arriving passengers queued at this VCE (pax)
    :param capacity: this VCE's capacity, per `vce_capacity` (pax/s)
    :return: this VCE's upward flow (pax/s)
    """
    return min(arriving_pax_on_platform, capacity)


def down_capacities(
    capacities: list[float], up_rates: list[float], bidirectional_limits: list[float]
) -> list[float]:
    """
    Each VCE's capacity (pax/s) left for departing passengers to come down:
    whatever its upward flow leaves,
    unless that exceeds its `Assumptions.bidirectional_stair_flow_limit`.

    :param capacities: each VCE's capacity, per `vce_capacity` (pax/s)
    :param up_rates: each VCE's upward flow (pax/s)
    :param bidirectional_limits: each VCE's `Assumptions.bidirectional_stair_flow_limit` (pax/s)
    """
    return [
        0 if up_rate > limit else capacity - up_rate
        for capacity, up_rate, limit in zip(capacities, up_rates, bidirectional_limits, strict=True)
    ]


def platform_ingress(departing_pax_upstairs: float, down_capacity: float) -> float:
    """
    Departing passengers queue upstairs and come down with whatever capacity is left for them.

    :param departing_pax_upstairs: number of departing passengers upstairs (pax)
    :param down_capacity: this train's share of the VCEs' `down_capacities` (pax/s)
    :return: platform ingress rate (pax/s)
    """
    return min(departing_pax_upstairs, down_capacity)


ROUNDING_TOLERANCE = 1e-9
"""
Passengers (pax) left over from rounding, e.g. 1e-15 after the last of them come downstairs,
which count as nobody.
"""


def subtract(total: float, pax: float) -> float:
    """`total - pax`, or 0 if that's only left over from rounding, per `ROUNDING_TOLERANCE`."""
    remaining = total - pax
    return 0.0 if remaining < ROUNDING_TOLERANCE else remaining


def boarder_fraction(train_boarders: float, all_boarders: list[float]) -> float:
    """
    :param train_boarders: one train's departing passengers upstairs (pax)
    :param all_boarders: every train's departing passengers upstairs (pax)
    :return: that train's share of them
    """
    total = sum(all_boarders)
    # A total of only rounding errors would give a share of nearly 1 / 0.
    if total > ROUNDING_TOLERANCE:
        return train_boarders / total
    else:
        return 1


def board_rate(
    max_rate: float,
    off_rate: float,
    time: timedelta,
    arrival_time: timedelta | None,
    boarders: float,
) -> float:
    """
    Nobody boards until everyone has alighted, per the ETA report,
    i.e. the second after `off_rate` is last nonzero.

    :param max_rate: maximum boarding rate across all doors (pax/s)
    :param off_rate: train alight rate (pax/s)
    :param time: time
    :param arrival_time: train arrival time, or `None` if it hasn't been scheduled yet
    :param boarders: number of passengers waiting on platform to board
    :return: train ingress rate across all doors (pax/s)
    """
    if arrival_time is not None and arrival_time < time and off_rate == 0:
        return min(max_rate, boarders)
    else:
        return 0


def calc_space_per_pax(pax_on_platform: float, area: float) -> float:
    """
    :param pax_on_platform: people on platform (pax)
    :param area: usable platform area (ft^2)
    :return: space per passenger (ft^2/pax)
    """
    if pax_on_platform > 0:
        return area / pax_on_platform
    else:
        return area


def platform_crowd_los(space_per_pax: float, assumptions: Assumptions) -> str:
    """
    :param space_per_pax: space per passenger on the platform (ft^2/pax)
    :return: its LOS, per `Assumptions.platform_los_min_space`
    """
    for grade, min_space in assumptions.platform_los_min_space:
        if space_per_pax > min_space:
            return grade
    return "F"


def egress_crowd_los(capacity: float, up_rate: float, assumptions: Assumptions) -> str:
    """
    A VCE's LOS, per `Assumptions.stair_los_max_flow` and `stair_capacity`,
    as fractions of its capacity.

    :param capacity: the VCE's capacity, per `vce_capacity` (pax/s)
    :param up_rate: upward flow on it (pax/s)
    """
    for grade, max_flow in (*assumptions.stair_los_max_flow, ("E", assumptions.stair_capacity)):
        if up_rate <= capacity * max_flow / assumptions.stair_capacity:
            return grade
    return "F"


def worst_egress_los(vces: list[Vce], up_rates: list[float], assumptions: Assumptions) -> str:
    """The worst of each VCE's `egress_crowd_los`."""
    return max(
        egress_crowd_los(vce_capacity(vce, assumptions), up_rate, assumptions)
        for vce, up_rate in zip(vces, up_rates, strict=True)
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
        value = getattr(obj, attr)
        # CSVs and charts can't hold `timedelta`s, so give them in seconds, per `Field.units`.
        if isinstance(value, timedelta):
            value = round(value.total_seconds())
        # Lengths are easier to picture in feet and inches than in decimal feet.
        elif field.units == "ft":
            value = fmt_ft_in(value)
        yield attr, value, field


class VceType(StrEnum):
    """What kind of VCE it is, as written in `data/`."""

    STAIR = "stair"
    ESCALATOR = "escalator"


@dataclass(frozen=True)
class Vce:
    """A VCE (vertical circulation element), i.e. a stair or escalator going upstairs."""

    name: str
    """Its name, e.g. `P3-S4`."""

    width: float
    """Its width (in feet)."""

    west_end: float
    """Where it starts along the platform (ft east of the Master Plan's plans' west edge)."""

    east_end: float
    """Where it ends along the platform (ft east of the Master Plan's plans' west edge)."""

    type: VceType = VceType.STAIR
    """Whether it's a stair or an escalator."""


VCE_DATA = DATA_DIR / "vces.csv"
"""Every VCE on platforms 1 to 11, via `platform-crowd-model data vces`."""


PLATFORM_LENGTHS = DATA_DIR / "platform_lengths_moynihan_ea.csv"
"""Each platform's length, from the Moynihan Station EA's Table 4.4-10."""

PLATFORM_MAX_CARS = DATA_DIR / "platform_max_cars_track_map.csv"
"""
Cars in the longest train that fits on each platform's tracks,
from a track map of unknown origin found at Railfan Guides of the U.S.:
https://www.railfanguides.us/ny/penntonewrochelle/PennStationLayout1.jpg
"""

PLATFORM_EAST_ENDS = DATA_DIR / "platform_east_ends.csv"
"""
Where each platform ends to the east (ft east of the Master Plan's plans' west edge),
via `platform-crowd-model data vces`.
"""

PLATFORM_WEST_ENDS = DATA_DIR / "platform_west_ends_pcip_phase_1.csv"
"""
Where each platform ends to the west (ft east of the Master Plan's plans' west edge),
via `platform-crowd-model data platform-west-ends-pcip-phase-1`.
"""

TRANSFORMATION_VCES = DATA_DIR / "vces_transformation_fra_sos.csv"
"""Penn Transformation's new VCEs, via `platform-crowd-model data vces-transformation-fra-sos`."""

TRANSFORMATION_PLATFORMS = DATA_DIR / "platforms_transformation_fra_sos.csv"
"""
Penn Transformation's platform extensions and decluttering,
via `platform-crowd-model data vces-transformation-fra-sos`.
"""

PLATFORM_A_DATA = DATA_DIR / "platform_a_pcip_phase_1.csv"
"""
PCIP Phase 1's Platform A, a new platform south of Platform 1,
via `platform-crowd-model data platform-a-pcip-phase-1`.
"""

PLATFORM_A_VCES = DATA_DIR / "vces_platform_a_pcip_phase_1.csv"
"""Platform A's VCEs, via `platform-crowd-model data platform-a-pcip-phase-1`."""

PLATFORM_A = 0
"""Platform A's number in the model, since the existing platforms are numbered 1 to 11."""

PLATFORMS_OSM = DATA_DIR / "platforms_osm.csv"
"""
Each platform's outline's area and average width, from OpenStreetMap,
via `platform-crowd-model data platforms-osm`.
"""


@cache
def platform_a() -> dict[str, str]:
    """Platform A's row of `PLATFORM_A_DATA`."""
    with PLATFORM_A_DATA.open() as f:
        (row,) = csv.DictReader(f)
    return row


@cache
def platform_vces(platform: int) -> tuple[Vce, ...]:
    """
    Every VCE on `platform`, from `VCE_DATA`,
    with the Master Plan's width where it has one, or else the estimated width,
    or on Platform A, from `PLATFORM_A_VCES`.
    """
    if platform == PLATFORM_A:
        with PLATFORM_A_VCES.open() as f:
            return tuple(
                Vce(
                    name=f"PA-{row['label']}" + ("-E" if row["type"] == VceType.ESCALATOR else ""),
                    width=float(row["width_in"]) / 12,
                    type=VceType(row["type"]),
                    west_end=float(row["west_end_ft"]),
                    east_end=float(row["east_end_ft"]),
                )
                for row in csv.DictReader(f)
            )
    with VCE_DATA.open() as f:
        return tuple(
            Vce(
                name=row["vce_name"],
                width=float(row["master_plan_width_in"] or row["estimated_width_in"]) / 12,
                type=VceType(row["type"]),
                west_end=float(row["west_end_ft"]),
                east_end=float(row["east_end_ft"]),
            )
            for row in csv.DictReader(f)
            if int(row["platform"]) == platform
        )


@cache
def platform_lengths() -> dict[int, int]:
    """Each platform's length (ft), from `PLATFORM_LENGTHS`, and Platform A's."""
    with PLATFORM_LENGTHS.open() as f:
        lengths = {int(row["platform"]): int(row["length_ft"]) for row in csv.DictReader(f)}
    a = platform_a()
    return {PLATFORM_A: int(a["east_end_ft"]) - int(a["west_end_ft"]), **lengths}


@cache
def platform_tracks() -> dict[int, int]:
    """How many tracks each platform serves, from `PLATFORM_LENGTHS`, and Platform A."""
    with PLATFORM_LENGTHS.open() as f:
        return {
            PLATFORM_A: int(platform_a()["tracks"]),
            **{
                int(row["platform"]): len(row["track_numbers"].split("/"))
                for row in csv.DictReader(f)
            },
        }


@cache
def platform_max_cars() -> dict[int, int]:
    """
    Cars in the longest train that fits on each platform's tracks, from `PLATFORM_MAX_CARS`,
    and Platform A's.
    """
    with PLATFORM_MAX_CARS.open() as f:
        return {
            PLATFORM_A: int(platform_a()["max_cars"]),
            **{int(row["platform"]): int(row["max_cars"]) for row in csv.DictReader(f)},
        }


@cache
def platform_east_ends() -> dict[int, float]:
    """Where each platform ends to the east (ft), from `PLATFORM_EAST_ENDS`, and Platform A."""
    with PLATFORM_EAST_ENDS.open() as f:
        return {
            PLATFORM_A: float(platform_a()["east_end_ft"]),
            **{int(row["platform"]): float(row["east_end_ft"]) for row in csv.DictReader(f)},
        }


@cache
def platform_areas() -> dict[int, int]:
    """Each platform's area (sq ft), from `PLATFORMS_OSM`, and Platform A's."""
    with PLATFORMS_OSM.open() as f:
        return {
            PLATFORM_A: int(platform_a()["area_sq_ft"]),
            **{
                int(row["platform"]): int(row["area_sq_ft"])
                for row in csv.DictReader(f)
                if row["platform"] and row["level"] == "-3"
            },
        }


@cache
def platform_west_ends() -> dict[int, float]:
    """Where each platform ends to the west (ft), from `PLATFORM_WEST_ENDS`."""
    with PLATFORM_WEST_ENDS.open() as f:
        return {int(row["platform"]): float(row["west_end_ft"]) for row in csv.DictReader(f)}


@cache
def platform_widths() -> dict[int, float]:
    """Each platform's average width (ft), from `PLATFORMS_OSM`."""
    with PLATFORMS_OSM.open() as f:
        return {
            int(row["platform"]): float(row["width_ft"])
            for row in csv.DictReader(f)
            if row["platform"] and row["level"] == "-3"
        }


@cache
def transformation_platforms() -> dict[int, dict[str, str]]:
    """Each platform's row of `TRANSFORMATION_PLATFORMS`."""
    with TRANSFORMATION_PLATFORMS.open() as f:
        return {int(row["platform"]): row for row in csv.DictReader(f)}


@cache
def transformation_vces(platform: int) -> tuple[Vce, ...]:
    """Penn Transformation's new VCEs on `platform`, from `TRANSFORMATION_VCES`, e.g. `P3-T1`."""
    with TRANSFORMATION_VCES.open() as f:
        rows = [row for row in csv.DictReader(f) if int(row["platform"]) == platform]
    return tuple(
        Vce(
            name=f"P{platform}-T{n}",
            width=float(row["width_in"]) / 12,
            west_end=float(row["west_end_ft"]),
            east_end=float(row["east_end_ft"]),
            type=VceType(row["type"]),
        )
        for n, row in enumerate(rows, start=1)
    )


def door_positions(params: Params) -> list[float]:
    """
    Where each of a train's doors is along the platform (ft),
    spread evenly along a train stopped with its east end at the platform's.
    """
    assumptions = params.assumptions
    train_west_end = params.platform_east_end - params.train_length
    door_spacing = assumptions.car_length / assumptions.doors_per_car
    return [train_west_end + (door + 0.5) * door_spacing for door in range(params.doors_per_train)]


def distance_to(vce: Vce, position: float) -> float:
    """Distance (ft) along the platform from `position` to the nearest end of `vce`."""
    return max(0, vce.west_end - position, position - vce.east_end)


@dataclass(frozen=True)
class Door:
    """Where some of the arriving passengers come from."""

    share: float
    """Fraction of the arriving passengers."""

    walking_times: list[int]
    """
    Time they take to walk to each VCE, in `TIME_STEP`s,
    so the simulation can look up who arrives each step without `timedelta` arithmetic.
    """


def doors_to_vces(params: Params, vces: list[Vce]) -> list[Door]:
    """
    Where arriving passengers come from, and how far they are from each of `vces`.
    """
    doors = door_positions(params)
    return [
        Door(
            share=1 / len(doors),
            walking_times=[
                round(distance_to(vce, door) / params.assumptions.walking_speed) for vce in vces
            ],
        )
        for door in doors
    ]


def vce_wait(queue: float, walking_to: float, capacity: float) -> float:
    """
    How long (s) someone reaching a VCE now would wait to go up,
    behind everyone queued at or walking to it.

    :param queue: arriving passengers queued at the VCE (pax)
    :param walking_to: arriving passengers walking to the VCE (pax)
    :param capacity: the VCE's upward capacity (pax/s)
    """
    return (queue + walking_to) / capacity


def choose_vce(params: Params, door: Door, waits: list[float]) -> int:
    """
    The VCE the passengers from `door` walk to, per `Assumptions.vce_choice`.

    :param waits: each VCE's `vce_wait` (s)
    """
    # The first of any tied, i.e. the westernmost.
    walks = door.walking_times
    if params.assumptions.vce_choice == "nearest":
        return walks.index(min(walks))
    times = [walk + wait for walk, wait in zip(walks, waits, strict=True)]
    return times.index(min(times))


@dataclass
class Params:
    platform: Annotated[int, Field(name="Platform", units="#")]
    """
    Which platform it is, 1 to 11, or `PLATFORM_A`, 0, whose dimensions are in `data/`.
    """

    headway: Annotated[timedelta, Field(name="Headway", units="s")]
    """
    Time between trains' scheduled arrivals.
    The first arrives at 0 s, on one track, and the second on the other.
    """

    vces: tuple[Vce, ...]
    """The VCEs (vertical circulation elements) going upstairs, from west to east."""

    trains: Annotated[int, Field(name="Trains", units="train")] = 4
    """Trains arriving, alternating between the platform's tracks."""

    assumptions: Assumptions = dataclasses.field(default_factory=Assumptions)
    """What the model assumes, the same for every scenario unless overridden."""

    transformation: bool = False
    """
    Whether the platform is as Penn Transformation would leave it,
    extended, with longer trains, and decluttered, per `TRANSFORMATION_PLATFORMS`.
    Its new VCEs, `transformation_vces`, are in `vces`.
    """

    modifier: str | None = None
    """
    What sets this scenario apart from the platform's others,
    e.g. `transformation` for Penn Transformation.
    """

    def __post_init__(self) -> None:
        """Sort `vces` from west to east."""
        self.vces = tuple(sorted(self.vces, key=lambda vce: vce.west_end))

    @property
    def name(self) -> str:
        """
        The platform's name in the results table, e.g. `3T` for `transformation`,
        its number followed by its modifier's initial, to keep the table narrow.
        """
        number = "A" if self.platform == PLATFORM_A else str(self.platform)
        return number + (self.modifier[0].upper() if self.modifier else "")

    @property
    def filename_prefix(self) -> str:
        """
        Prefix of the filenames to save its time series and charts in,
        e.g. `platform3_transformation`.
        """
        number = "A" if self.platform == PLATFORM_A else str(self.platform)
        return f"platform{number}" + (f"_{self.modifier}" if self.modifier else "")

    @property
    def transformation_platform(self) -> dict[str, str]:
        """The platform's row of `TRANSFORMATION_PLATFORMS`, if `transformation`, or else none."""
        return transformation_platforms()[self.platform] if self.transformation else {}

    @property
    def platform_extension(self) -> Annotated[float, Field(name="Platform Extension", units="ft")]:
        """How far Penn Transformation extends the platform to the west (ft), if at all."""
        west_end = self.transformation_platform.get("west_end_ft")
        return platform_west_ends()[self.platform] - float(west_end) if west_end else 0

    @property
    def platform_east_end(self) -> float:
        """
        Where the platform ends to the east (ft east of the Master Plan's plans' west edge),
        from `PLATFORM_EAST_ENDS`.
        """
        return platform_east_ends()[self.platform]

    @property
    def platform_length(self) -> Annotated[float, Field(name="Platform Length", units="ft")]:
        """Platform length (in feet), from the Moynihan Station EA, plus `platform_extension`."""
        return platform_lengths()[self.platform] + self.platform_extension

    @property
    def platform_max_cars(self) -> Annotated[int, Field(name="Platform Max Cars", units="car")]:
        """
        Cars in the longest train that fits on its tracks, from `PLATFORM_MAX_CARS`,
        or with Penn Transformation's extension, from `TRANSFORMATION_PLATFORMS`.
        """
        max_cars = self.transformation_platform.get("max_cars")
        return int(max_cars) if max_cars else platform_max_cars()[self.platform]

    @property
    def tracks(self) -> Annotated[int, Field(name="Tracks", units="track")]:
        """
        Tracks the platform serves, from `PLATFORM_LENGTHS`: 2 for an island platform,
        or 1 for platform 9, which only serves track 17.
        """
        return platform_tracks()[self.platform]

    @property
    def cars(self) -> Annotated[int, Field(name="Cars per Train", units="car")]:
        """
        Cars in each train: as many as fit on the platform's tracks and along the platform,
        with up to `max_train_overhang`, up to `max_train_cars`.
        """
        return min(
            self.platform_max_cars,
            int(
                (self.platform_length + self.assumptions.max_train_overhang)
                // self.assumptions.car_length
            ),
            self.assumptions.max_train_cars,
        )

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
    def platform_area(self) -> Annotated[float, Field(name="Platform Area", units="ft^2")]:
        """
        Platform area (in square feet), from OpenStreetMap's outline of it,
        which accounts for platforms tapering.
        With Penn Transformation, decluttering adds the FRA's percentage more circulation area,
        and `platform_extension` adds its length at the platform's average width.
        """
        area = platform_areas()[self.platform]
        row = self.transformation_platform
        if not row:
            return area
        decluttered = int(row["added_circulation_area_sq_ft"]) / int(row["circulation_area_sq_ft"])
        return area * (1 + decluttered) + self.platform_extension * platform_widths()[self.platform]

    @property
    def simulated_vces(self) -> list[Vce]:
        """
        The VCEs the simulation uses, all treated as stairs,
        but the widest escalator, which is left out,
        like the ETA report's one VCE per platform, e.g. an escalator running the other way,
        or if the widest tie, the westernmost of `nfpa_130_out_of_service_choices`.
        """
        left_out = next(iter(self.nfpa_130_out_of_service_choices), None)
        return [vce for vce in self.vces if vce is not left_out]

    @property
    def total_vce_width(self) -> Annotated[float, Field(name="Total VCE Width", units="ft")]:
        """Total width (in feet) of `simulated_vces`."""
        return sum(vce.width for vce in self.simulated_vces)

    @property
    def nfpa_130_out_of_service_choices(self) -> list[Vce]:
        """
        Each escalator NFPA 130 could take out of service,
        the one "having the most adverse effect upon egress capacity" (5.3.5.4),
        i.e. the widest, since every VCE has the same capacity per width,
        or each of the widest if they tie, or none if there are no escalators.
        """
        escalators = [vce for vce in self.vces if vce.type == VceType.ESCALATOR]
        widest = max((vce.width for vce in escalators), default=None)
        return [vce for vce in escalators if vce.width == widest]

    def nfpa_130_exit_capacity_without(self, out_of_service: Vce | None) -> float:
        """
        How fast the VCEs can evacuate the platform under NFPA 130 (in pax/s),
        at `NFPA_130_EXIT_FLOW`, with `out_of_service` out of service,
        and escalators providing at most half of the capacity, per NFPA 130 5.3.5.6.
        """
        exits = [vce for vce in self.vces if vce is not out_of_service]
        stairs = sum(vce.width for vce in exits if vce.type != VceType.ESCALATOR)
        escalators = sum(vce.width for vce in exits if vce.type == VceType.ESCALATOR)
        return stair_flow(NFPA_130_EXIT_FLOW, stairs + min(escalators, stairs))

    @property
    def nfpa_130_exit_capacity(
        self,
    ) -> Annotated[float, Field(name="NFPA 130 Exit Capacity", units="pax/s")]:
        """
        `nfpa_130_exit_capacity_without` the widest escalator,
        the same for each of `nfpa_130_out_of_service_choices`.
        """
        return self.nfpa_130_exit_capacity_without(
            next(iter(self.nfpa_130_out_of_service_choices), None)
        )

    def nfpa_130_longest_walk_without(self, out_of_service: Vce | None) -> float:
        """
        Farthest anyone on the platform is from their nearest VCE (ft),
        with `out_of_service` out of service, if any,
        from either end of the platform, or from halfway between two VCEs.
        """
        exits = [vce for vce in self.vces if vce is not out_of_service]
        # From the west end to the first exit, and from the last exit to the east end.
        longest = max(
            exits[0].west_end - (self.platform_east_end - self.platform_length),
            self.platform_east_end - max(vce.east_end for vce in exits),
        )
        # Halfway between each exit and the next.
        east = exits[0].east_end
        for vce in exits[1:]:
            longest = max(longest, (vce.west_end - east) / 2)
            east = max(east, vce.east_end)
        return max(longest, 0)

    def nfpa_130_evacuation_time(self, occupants: float) -> timedelta:
        """
        How long `occupants` take to flow out through its exits at `nfpa_130_exit_capacity`,
        NFPA 130's platform evacuation time (5.3.3.1).
        """
        return timedelta(seconds=ceil(occupants / self.nfpa_130_exit_capacity))

    def nfpa_130_time_to_concourse_without(
        self, occupants: float, out_of_service: Vce | None
    ) -> timedelta:
        """
        How long the farthest of `occupants` takes to reach the concourse,
        taken as NFPA 130's point of safety (5.3.3.2), per its Annex C,
        with `out_of_service` out of service, if any:
        their walk along the platform to their nearest exit at `NFPA_130_PLATFORM_WALKING_SPEED`,
        plus their wait there, the rest of the platform's flow time after that walk,
        plus their climb up `PLATFORM_TO_CONCOURSE_RISE` at `NFPA_130_STAIR_VERTICAL_SPEED`.
        """
        walk = self.nfpa_130_longest_walk_without(out_of_service) / NFPA_130_PLATFORM_WALKING_SPEED
        flow = occupants / self.nfpa_130_exit_capacity_without(out_of_service)
        climb = PLATFORM_TO_CONCOURSE_RISE / NFPA_130_STAIR_VERTICAL_SPEED
        return timedelta(seconds=ceil(max(walk, flow) + climb))

    def nfpa_130_time_to_concourse(self, occupants: float) -> timedelta:
        """
        `nfpa_130_time_to_concourse_without` the widest escalator (5.3.5.4).
        NFPA 130 doesn't say which to choose if the widest tie,
        so this is the shortest with each of `nfpa_130_out_of_service_choices`.
        NFPA 130's travel distance limit (5.3.3.5) doesn't take an escalator out of service,
        so its walk is `nfpa_130_longest_walk_without(None)` instead.
        """
        return min(
            self.nfpa_130_time_to_concourse_without(occupants, out_of_service)
            for out_of_service in self.nfpa_130_out_of_service_choices or [None]
        )


@dataclass
class Instant:
    """
    An instant in the simulation.
    """

    time: Annotated[timedelta, Field(name="Time", units="s")]
    """Time since the first train's scheduled arrival."""

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

    time_at_capacity: timedelta
    """How long the upstairs rate is at the VCEs' LOS E capacity (17 pax/min/ft)."""

    taper_time: timedelta | None
    """
    Last second the arriving passengers on the platform exceed what fits in the stair queues,
    i.e. when they start to taper off, or `None` if they never do.
    """

    clear_time: timedelta | None
    """First second after the last arrival when all arriving passengers have left the platform."""

    arrival_times: list[timedelta | None]
    """When each train arrives, or `None` if it doesn't within the simulation."""

    dwells: list[timedelta | None]
    """
    Each train's dwell:
    from its arrival until all of its arriving passengers have alighted
    and all of its departing passengers have boarded,
    or `None` if that doesn't happen within the simulation.
    """

    boarded_time: timedelta | None
    """First second when all departing passengers have boarded, or `None` if they never do."""

    max_pax_on_platform: float
    """Most passengers on the platform at once."""

    max_occupants: float
    """
    Most passengers on the platform or aboard its trains at once,
    counting trains from their arrival until they depart.
    """

    min_space_per_pax: float
    """Least platform space per passenger (sq ft)."""

    vce_gone_up: list[float]
    """Arriving passengers who've gone up each VCE, for `unused_vces`."""


TRAIN_COLUMNS = [
    "Passengers (pax)",
    "Alighting Rate (pax/s)",
    "Boarding Rate (pax/s)",
    "Departing Passengers on Platform (pax)",
]
"""What `TimeSeries.trains` has for each train each second."""


@dataclass
class TimeSeries:
    """Everything that happens each second of one model run, for its CSVs and charts."""

    instants: list[Instant] = dataclasses.field(default_factory=list[Instant])
    """The whole platform each second."""

    trains: list[list[list[float]]] = dataclasses.field(default_factory=list[list[list[float]]])
    """Each second, each train's `TRAIN_COLUMNS`."""


def simulate(
    params: Params, record_time_series: bool = True, print_time_series: bool = True
) -> tuple[TimeSeries, Summary]:
    """
    Simulate `params`, returning its time series and its results table's summary.
    Without `record_time_series`, the time series is left empty,
    and without `print_time_series`, nothing is printed,
    e.g. when only the summary is needed.
    """
    assumptions = params.assumptions
    usable_area = params.platform_area * assumptions.usable_platform_area_multiplier
    door_rate = params.doors_per_train * assumptions.door_flow_rate

    # Initialize counters
    arriving_pax_on_platform: float = 0
    vces = params.simulated_vces
    vce_queues = [0.0 for _ in vces]
    """Arriving passengers queued at each VCE."""
    vce_capacities = [vce_capacity(vce, assumptions) for vce in vces]
    """Each VCE's capacity in one direction (pax/s)."""
    doors = doors_to_vces(params, vces)
    walking: defaultdict[int, list[float]] = defaultdict(lambda: [0.0 for _ in vces])
    """Arriving passengers walking to each VCE, by the step they reach its queue."""
    walking_totals = [0.0 for _ in vces]
    """Arriving passengers walking to each VCE, whenever they reach it, i.e. `walking` summed."""
    bidirectional_limits = [
        stair_flow(assumptions.bidirectional_stair_flow_limit, vce.width) for vce in vces
    ]
    trains = range(params.trains)
    arrival_times: list[timedelta | None] = [
        train * params.headway if train < params.tracks else None for train in trains
    ]
    """
    When each train arrives: the first on each track as scheduled,
    and each later one once it's scheduled and the train before it on its track has departed.
    """
    remaining_arrivals = [float(params.arriving_pax_per_train) for _ in trains]
    new_pax = [0.0 for _ in trains]
    release_times = [
        train * params.headway - assumptions.departing_pax_lead_time for train in trains
    ]
    """When each train's departing passengers start coming down to the platform."""
    start_time = min(timedelta(0), *release_times)
    """
    When the simulation starts: before the first train arrives at 0 s,
    once its departing passengers start coming down.
    """
    boarders_upstairs = [0.0 for _ in trains]
    boarders_on_platform = [0.0 for _ in trains]
    total_pax_on_platform: float = 0
    time_series = TimeSeries()
    gone_up: float = 0
    """Arriving passengers who've gone up, for `check_conservation`."""

    total_vce_width = params.total_vce_width
    max_pax_in_stair_queues = (
        total_vce_width * assumptions.stair_queue_length / assumptions.stair_queue_space
    )
    capacity = sum(vce_capacities)
    no_flow = [0.0 for _ in vces]
    """Each VCE's upward flow while nobody is queued."""
    summary = Summary(
        max_up_rate=0,
        time_at_capacity=timedelta(0),
        taper_time=None,
        clear_time=None,
        arrival_times=arrival_times,
        dwells=[None for _ in trains],
        boarded_time=None,
        max_pax_on_platform=total_pax_on_platform,
        max_occupants=total_pax_on_platform,
        min_space_per_pax=calc_space_per_pax(total_pax_on_platform, usable_area),
        vce_gone_up=[0.0 for _ in vces],
    )

    if print_time_series:
        print("Elapsed_Time", *(f"Train_{train + 1}_Pax" for train in trains))

    time_after = start_time
    for step in itertools.count():
        time_after = start_time + step * TIME_STEP
        if time_after - start_time >= MAX_SIMULATION_LENGTH:
            raise RuntimeError(
                f"{params.filename_prefix} hasn't finished after {MAX_SIMULATION_LENGTH}"
            )
        for train in trains:
            if release_times[train] == time_after:
                boarders_upstairs[train] = float(assumptions.departing_pax_per_train)
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
        arriving_pax_on_platform += sum(off_rates)
        if arriving_pax_on_platform > 0:
            # Each door's alighting passengers walk to a VCE,
            # though nobody does in a second nobody alights.
            alighting = sum(off_rates)
            if alighting > 0:
                waits = [
                    vce_wait(queue, walking_to, capacity)
                    for queue, walking_to, capacity in zip(
                        vce_queues, walking_totals, vce_capacities, strict=True
                    )
                ]
                for door in doors:
                    i = choose_vce(params, door, waits)
                    walking[step + door.walking_times[i]][i] += alighting * door.share
                    walking_totals[i] += alighting * door.share
                    waits[i] = vce_wait(vce_queues[i], walking_totals[i], vce_capacities[i])
            for i, reaching in enumerate(walking.pop(step, [])):
                vce_queues[i] += reaching
                walking_totals[i] = subtract(walking_totals[i], reaching)
            vce_up_rates = []
            for i, queue in enumerate(vce_queues):
                vce_up_rate = platform_clearance(queue, vce_capacities[i])
                vce_queues[i] = queue - vce_up_rate
                vce_up_rates.append(vce_up_rate)
                summary.vce_gone_up[i] += vce_up_rate
        else:
            # While nobody is queued, skip each VCE.
            vce_up_rates = no_flow
        up_rate = sum(vce_up_rates)
        gone_up += up_rate
        arriving_pax_on_platform -= up_rate
        if arriving_pax_on_platform < 0:
            arriving_pax_on_platform = 0
        total_pax_on_platform -= up_rate
        # Each train's boarders get a share of each VCE,
        # and so a share of the upward flow on it and of the capacity it leaves.
        # Their shares of a VCE's capacity, upward flow, and bidirectional limit scale together,
        # so whether it has any capacity left doesn't depend on the train.
        down_capacity = (
            capacity
            if vce_up_rates is no_flow
            else sum(down_capacities(vce_capacities, vce_up_rates, bidirectional_limits))
        )
        down_rates = [
            platform_ingress(
                boarders_upstairs[train],
                down_capacity * boarder_fraction(boarders_upstairs[train], boarders_upstairs),
            )
            for train in trains
        ]
        for train in trains:
            boarders_on_platform[train] += down_rates[train]
            total_pax_on_platform += down_rates[train]
        on_rates = [
            board_rate(
                door_rate,
                off_rates[train],
                time_after,
                arrival_times[train],
                boarders_on_platform[train],
            )
            for train in trains
        ]

        for train in trains:
            boarders_on_platform[train] -= on_rates[train]
            total_pax_on_platform -= on_rates[train]
            boarders_upstairs[train] -= down_rates[train]
            new_pax[train] += on_rates[train]

        space_per_pax = calc_space_per_pax(total_pax_on_platform, usable_area)
        if total_pax_on_platform < 0:
            total_pax_on_platform = 0
        for train in trains:
            if boarders_on_platform[train] < 0:
                boarders_on_platform[train] = 0
        if arriving_pax_on_platform < 0:
            arriving_pax_on_platform = 0
        if print_time_series:
            print(
                round(time_after.total_seconds()),
                *(remaining_arrivals[train] + new_pax[train] for train in trains),
                arriving_pax_on_platform,
                up_rate,
            )
        summary.max_up_rate = max(summary.max_up_rate, up_rate)
        if up_rate >= capacity - 1e-9:
            summary.time_at_capacity += TIME_STEP
        if arriving_pax_on_platform > max_pax_in_stair_queues:
            summary.taper_time = time_after
        if (
            summary.clear_time is None
            and all(
                arrival_time is not None and time_after > arrival_time
                for arrival_time in arrival_times
            )
            and arriving_pax_on_platform < 1
        ):
            summary.clear_time = time_after
        if (
            summary.boarded_time is None
            and time_after >= max(release_times)
            and sum(boarders_upstairs) + sum(boarders_on_platform) < 1
        ):
            summary.boarded_time = time_after
        for train in trains:
            arrival_time = arrival_times[train]
            if (
                summary.dwells[train] is None
                and arrival_time is not None
                and time_after > arrival_time
                and remaining_arrivals[train] < 1
                and boarders_upstairs[train] + boarders_on_platform[train] < 1
            ):
                summary.dwells[train] = time_after - arrival_time
                # The next train on its track arrives once it's scheduled and this one departs.
                if train + params.tracks < params.trains:
                    arrival_times[train + params.tracks] = max(
                        (train + params.tracks) * params.headway, time_after
                    )
        summary.max_pax_on_platform = max(summary.max_pax_on_platform, total_pax_on_platform)
        aboard = sum(
            remaining_arrivals[train] + new_pax[train]
            for train in trains
            if (arrival_time := arrival_times[train]) is not None
            and time_after >= arrival_time
            and summary.dwells[train] is None
        )
        summary.max_occupants = max(summary.max_occupants, total_pax_on_platform + aboard)
        summary.min_space_per_pax = min(summary.min_space_per_pax, space_per_pax)

        if record_time_series:
            net_pax_flow_rate: float = 0
            for rate in down_rates:
                net_pax_flow_rate += rate
            for rate in off_rates:
                net_pax_flow_rate += rate
            net_pax_flow_rate -= up_rate
            for rate in on_rates:
                net_pax_flow_rate -= rate
            instant = Instant(
                time=time_after,
                arriving_pax_waiting_on_platform=arriving_pax_on_platform,
                off_rate=sum(off_rates),
                on_rate=sum(on_rates),
                down_rate=sum(down_rates),
                departing_pax_on_platform=sum(boarders_on_platform),
                total_pax_on_platform=total_pax_on_platform,
                platform_crowding=space_per_pax,
                up_rate=up_rate,
                net_pax_flow_rate=net_pax_flow_rate,
                platform_crowd_los=platform_crowd_los(space_per_pax, assumptions),
                egress_los=worst_egress_los(vces, vce_up_rates, assumptions),
            )

            time_series.instants.append(instant)
            time_series.trains.append(
                [
                    [
                        remaining_arrivals[train] + new_pax[train],
                        off_rates[train],
                        on_rates[train],
                        boarders_on_platform[train],
                    ]
                    for train in trains
                ]
            )

        # Stop once the last train has departed and the platform has cleared.
        if (
            summary.clear_time is not None
            and summary.boarded_time is not None
            and all(dwell is not None for dwell in summary.dwells)
        ):
            break
    check_conservation(
        params,
        arriving=[
            ("went up", gone_up),
            ("are still aboard", sum(remaining_arrivals)),
            ("are still on the platform", arriving_pax_on_platform),
        ],
        departing=[
            ("boarded", sum(new_pax)),
            ("are still upstairs", sum(boarders_upstairs)),
            ("are still on the platform", sum(boarders_on_platform)),
        ],
    )
    return time_series, summary


CONSERVATION_TOLERANCE = 1e-6
"""How far off `check_conservation`'s totals can be (pax), for floating-point rounding."""


def check_conservation(
    params: Params, arriving: list[tuple[str, float]], departing: list[tuple[str, float]]
) -> None:
    """
    Check that a finished simulation neither created nor lost passengers:
    everywhere the arriving passengers are, e.g. gone up, still aboard, or on the platform,
    adds up to everyone who arrived,
    and likewise for the departing passengers.
    It only adds up where passengers are once per run, so it's cheap enough to check every run,
    including the `README.md` snapshot test's.

    :param arriving: how many of the arriving passengers are in each place, by its description
    :param departing: how many of the departing passengers are in each place, by its description
    """
    assumptions = params.assumptions
    for who, places, expected in (
        ("arriving", arriving, params.trains * params.arriving_pax_per_train),
        ("departing", departing, params.trains * assumptions.departing_pax_per_train),
    ):
        total = sum(pax for _place, pax in places)
        if abs(total - expected) > CONSERVATION_TOLERANCE:
            breakdown = ", ".join(f"{pax:.6g} {place}" for place, pax in places)
            raise RuntimeError(
                f"{params.filename_prefix} with a {params.headway} headway"
                f" has {total:.6g} {who} passengers, not {expected}: {breakdown}"
            )


RESULTS_COLUMNS = [
    "Platform",
    "Headway",
    "VCE width",
    "NFPA 130 evacuation",
    "NFPA 130 to concourse",
    "Arrivals",
    "Dwell",
    "Taper time",
    "Clear time",
    "Boarded time",
    "Time at capacity",
    "Max pax on platform",
    "Max density (pax/m²)",
    "Max up rate (pax/s)",
    "NFPA 130 travel distance",
    "Unused VCEs",
]
RESULTS_HEADER = "| " + " | ".join(RESULTS_COLUMNS) + " |\n" + "|---" * len(RESULTS_COLUMNS) + "|"

README = Path(__file__).parents[2] / "README.md"
RESULTS_START = "<!-- results-table:start -->"
RESULTS_END = "<!-- results-table:end -->"
"""The README's results table is between these markers, so `--update-readme` can replace it."""


def update_readme_results(table: str) -> None:
    """Replace the results table in the README with `table`."""
    readme = README.read_text()
    start = readme.index(RESULTS_START) + len(RESULTS_START)
    end = readme.index(RESULTS_END)
    README.write_text(f"{readme[:start]}\n{table}\n{readme[end:]}")


OUTPUT_DIR = Path("output")
"""Where `--charts` saves each scenario's CSVs and charts."""

SERIES_COLORS = ["#2a78d6", "#eb6834", "#1baf7a", "#eda100"]
"""Each chart's series colors, in order, from a colorblind-safe categorical palette."""


def write_csv(path: Path, header: list[str], rows: list[list[Any]]) -> None:
    """Write `header` and then `rows` to the CSV at `path`."""
    with path.open("w", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(header)
        writer.writerows(rows)


def save_time_series(params: Params, time_series: TimeSeries, stem: Path) -> None:
    """
    Save `params`' parameters and time series to CSVs,
    and its charts to an SVG, all named starting with `stem`.
    """
    param_values = [*annotated_field_values(params), *annotated_field_values(params.assumptions)]
    write_csv(
        stem.with_name(f"{stem.name}_params.csv"),
        ["Parameter", "Value"],
        [[field.description, value] for _attr, value, field in param_values],
    )
    write_csv(
        stem.with_suffix(".csv"),
        [field.description for _attr, field in annotated_field_names(Instant)],
        [
            [value for _attr, value, _field in annotated_field_values(instant)]
            for instant in time_series.instants
        ],
    )
    trains = range(params.trains)
    write_csv(
        stem.with_name(f"{stem.name}_trains.csv"),
        [
            "Time (s)",
            *(f"Train {train + 1} {column}" for train in trains for column in TRAIN_COLUMNS),
        ],
        [
            [round(instant.time.total_seconds()), *itertools.chain.from_iterable(train_values)]
            for instant, train_values in zip(time_series.instants, time_series.trains, strict=True)
        ],
    )

    # Only `--charts` needs `matplotlib`, so don't slow down every other run importing it.
    from matplotlib.figure import Figure

    times = [instant.time.total_seconds() for instant in time_series.instants]

    def column(attr: str) -> tuple[str, list[float]]:
        """The name and values of `Instant`'s `attr` each second."""
        field = dict(annotated_field_names(Instant))[attr]
        return field.name, [getattr(instant, attr) for instant in time_series.instants]

    charts: list[tuple[str, str, list[tuple[str, list[float]]]]] = [
        ("Up and Down Rates", "Rate (pax/s)", [column("up_rate"), column("down_rate")]),
        (
            "Passengers Aboard Trains",
            "Passengers",
            [
                (
                    f"Train {train + 1}",
                    [train_values[train][0] for train_values in time_series.trains],
                )
                for train in trains
            ],
        ),
        (
            "Passengers on Platform",
            "Passengers",
            [column("arriving_pax_waiting_on_platform"), column("total_pax_on_platform")],
        ),
        ("Space per Passenger", "Space per passenger (sq ft)", [column("platform_crowding")]),
        ("Net Platform Flow Rate", "Net Flow Rate (pax/s)", [column("net_pax_flow_rate")]),
    ]
    fig = Figure(figsize=(12, 3 * len(charts)), layout="constrained")
    fig.suptitle(f"Platform {params.name}, {round(params.headway.total_seconds())} s headway")
    axes = fig.subplots(len(charts), 1, sharex=True, squeeze=False)[:, 0]
    for ax, (title, y_label, series) in zip(axes, charts, strict=True):
        for color, (label, values) in zip(SERIES_COLORS, series, strict=False):
            ax.plot(times, values, label=label, color=color, linewidth=2)
        ax.set_title(title, loc="left")
        ax.set_ylabel(y_label)
        ax.set_xlim(times[0], times[-1])
        ax.grid(color="#e0e0dd", linewidth=0.5)
        ax.spines[["top", "right"]].set_visible(False)
        if len(series) > 1:
            ax.legend(loc="upper left", bbox_to_anchor=(1, 1), frameon=False)
    # Past 50 sq ft per passenger, the platform is nearly empty, so show just the crowded part.
    axes[3].set_ylim(0, 50)
    axes[-1].set_xlabel("Time (s)")
    fig.savefig(stem.with_suffix(".svg"))


def checkmark(ok: bool) -> str:
    """✓ if `ok`, or else ✗, for marking checks in the results table."""
    return "✓" if ok else "✗"


def unused_vces(params: Params, summary: Summary) -> list[str]:
    """
    The names of `params`' VCEs that no arriving passenger went up,
    e.g. because a nearer one is always quicker.
    """
    return [
        vce.name
        for vce, gone_up in zip(params.simulated_vces, summary.vce_gone_up, strict=True)
        if gone_up == 0
    ]


def run_model(params: Params, charts: bool) -> str:
    """
    Run the model, return its row of the results table,
    and with `charts`, print its time series and save its CSVs and charts in `OUTPUT_DIR`.
    """
    time_series, summary = simulate(
        params=params, record_time_series=charts, print_time_series=charts
    )

    headway = params.headway
    if charts:
        save_time_series(
            params,
            time_series,
            OUTPUT_DIR
            / (
                f"{params.filename_prefix}"
                f"_{params.arriving_pax_per_train}"
                f"_{params.arriving_pax_per_train}"
                f"_{round(headway.total_seconds())}s"
            ),
        )

    evacuation_time = params.nfpa_130_evacuation_time(summary.max_occupants)
    evacuation_ok = checkmark(evacuation_time <= NFPA_130_PLATFORM_EVACUATION_TIME)
    travel_distance = params.nfpa_130_longest_walk_without(None)
    travel_distance_ok = checkmark(travel_distance <= NFPA_130_MAX_TRAVEL_DISTANCE)
    to_concourse = params.nfpa_130_time_to_concourse(summary.max_occupants)
    to_concourse_ok = checkmark(to_concourse <= NFPA_130_POINT_OF_SAFETY_TIME)

    def fmt_time(t: timedelta | None) -> str:
        """`t` as `m:ss`."""
        if t is None:
            return "never"
        minutes, rest = divmod(t, timedelta(minutes=1))
        return f"{minutes}:{rest.seconds:02}"

    return (
        f"| {params.name} | {fmt_time(headway)} | {fmt_ft_in(params.total_vce_width)}"
        f" | {fmt_time(evacuation_time)} {evacuation_ok}"
        f" | {fmt_time(to_concourse)} {to_concourse_ok}"
        f" | {', '.join(fmt_time(arrival) for arrival in summary.arrival_times)}"
        f" | {', '.join(fmt_time(dwell) for dwell in summary.dwells)}"
        f" | {fmt_time(summary.taper_time)} | {fmt_time(summary.clear_time)}"
        f" | {fmt_time(summary.boarded_time)}"
        f" | {fmt_time(summary.time_at_capacity)}"
        f" | {summary.max_pax_on_platform:.0f}"
        f" | {1 / (summary.min_space_per_pax * SQUARE_METERS_PER_SQUARE_FOOT):.2f}"
        f" ({platform_crowd_los(summary.min_space_per_pax, params.assumptions)})"
        f" | {summary.max_up_rate:.2f}"
        f" | {fmt_ft_in(travel_distance)} {travel_distance_ok}"
        f" | {', '.join(unused_vces(params, summary))} |"
    )


def scenarios() -> list[Params]:
    """Every scenario, grouped by headway first, then by platform, as in the results table."""

    def platform_params(platform: int) -> Params:
        """`platform` with its VCEs, before choosing a headway."""
        return Params(platform=platform, headway=timedelta(0), vces=platform_vces(platform))

    def transformation_params(platform: int) -> Params:
        """`platform` as Penn Transformation would leave it, with its new VCEs."""
        params = platform_params(platform)
        return dataclasses.replace(
            params,
            transformation=True,
            modifier="transformation",
            vces=(*params.vces, *transformation_vces(platform)),
        )

    platforms = [
        # PCIP Phase 1's Platform A, south of Platform 1, which isn't part of Penn Transformation.
        platform_params(PLATFORM_A),
        *(
            params
            for platform in range(1, 12)
            for params in (platform_params(platform), transformation_params(platform))
        ),
    ]
    return [
        dataclasses.replace(params, headway=headway)
        for headway in (timedelta(0), CLOSE_HEADWAY, NORMAL_HEADWAY)
        for params in platforms
    ]


def main(update_readme: bool = False, charts: bool = False) -> None:
    """Run every scenario and print a table of their results."""
    if charts:
        OUTPUT_DIR.mkdir(exist_ok=True)
    with ProcessPoolExecutor() as executor:
        rows = list(
            executor.map(
                functools.partial(run_model, charts=charts),
                scenarios(),
            )
        )
    table = "\n".join([RESULTS_HEADER, *rows])
    print()
    print(table)
    if update_readme:
        update_readme_results(table)
