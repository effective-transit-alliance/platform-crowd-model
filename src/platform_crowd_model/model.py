"""
This is a recursive peak-hour platform clearance calculator.
model from https://onlinepubs.trb.org/Onlinepubs/hrr/1971/355/355-001.pdf
"""

import csv
import dataclasses
import functools
import itertools
import math
import typing
from collections.abc import Generator
from concurrent.futures import ProcessPoolExecutor
from dataclasses import dataclass
from datetime import timedelta
from functools import cache
from pathlib import Path
from typing import TYPE_CHECKING, Annotated, Any, Self, cast

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

SQUARE_METERS_PER_SQUARE_FOOT = 0.09290304

NFPA_130_EXIT_FLOW = 1.41 * 12
"""
Exit capacity of stairs and stopped escalators for evacuating a platform (pax/min/ft),
1.41 pax/min per inch of width in NFPA 130's 2010 edition.
TCQSM p. 10-51: https://onlinepubs.trb.org/onlinepubs/tcrp/tcrp_rpt_165ch-10.pdf#page=55
"""

NFPA_130_PLATFORM_EVACUATION_TIME = timedelta(minutes=4)
"""
Time within which NFPA 130 requires a platform's occupants,
including those on trains, to be able to evacuate it.
TCQSM p. 10-3: https://onlinepubs.trb.org/onlinepubs/tcrp/tcrp_rpt_165ch-10.pdf#page=7
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
    """Length of each car, a NJT MultiLevel's."""

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


def platform_clearance(
    arriving_pax_on_platform: float, vce_width: float, assumptions: Assumptions
) -> float:
    """
    Arriving passengers queue at the stairs,
    which discharge them at `Assumptions.stair_capacity` as long as anyone is queued.

    Fruin's stair equation relates flow to the space per passenger *on the stair*,
    which a queued stair holds near its critical density,
    so it doesn't apply to the space per passenger on the platform.
    The few seconds of walking from the doors to the stairs are ignored.

    :param arriving_pax_on_platform: number of arriving passengers on the platform (pax)
    :param vce_width: total width of vertical circulation elements (ft)
    :return: platform egress rate on stairs (pax/s)
    """
    return min(arriving_pax_on_platform, stair_flow(assumptions.stair_capacity, vce_width))


def platform_ingress(
    departing_pax_upstairs: float, vce_width: float, up_rate: float, assumptions: Assumptions
) -> float:
    """
    Departing passengers queue upstairs and come down with whatever stair capacity
    the upward flow leaves, unless it exceeds `Assumptions.bidirectional_stair_flow_limit`.

    :param departing_pax_upstairs: number of departing passengers upstairs (pax)
    :param vce_width: this train's share of the total width of vertical circulation elements (ft)
    :param up_rate: upward stair flow on this train's share of the stairs (pax/s)
    :return: platform ingress rate on stairs (pax/s)
    """
    if up_rate > stair_flow(assumptions.bidirectional_stair_flow_limit, vce_width):
        return 0
    return min(departing_pax_upstairs, stair_flow(assumptions.stair_capacity, vce_width) - up_rate)


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


def space_per_pax(pax_on_platform: float, area: float) -> float:
    """
    :param pax_on_platform: people on platform (pax)
    :param area: usable platform area (ft^2)
    :return: space per passenger (ft^2/pax)
    """
    if pax_on_platform > 0:
        return area / pax_on_platform
    else:
        return area


def platform_crowd_los(space: float, assumptions: Assumptions) -> str:
    """
    :param space: space per passenger on the platform (ft^2/pax)
    :return: its LOS, per `Assumptions.platform_los_min_space`
    """
    for grade, min_space in assumptions.platform_los_min_space:
        if space > min_space:
            return grade
    return "F"


def egress_crowd_los(vce_width: float, up_rate: float, assumptions: Assumptions) -> str:
    """
    :param vce_width: total width of vertical circulation elements (ft)
    :param up_rate: upward stair flow (pax/s)
    :return: its LOS, per `Assumptions.stair_los_max_flow` and `stair_capacity`
    """
    for grade, max_flow in (*assumptions.stair_los_max_flow, ("E", assumptions.stair_capacity)):
        if up_rate <= stair_flow(max_flow, vce_width):
            return grade
    return "F"


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


PLATFORM_LENGTHS = DATA_DIR / "platform_lengths_moynihan_ea.csv"
"""Each platform's length, from the Moynihan Station EA's Table 4.4-10."""

PLATFORM_MAX_CARS = DATA_DIR / "platform_max_cars_track_map.csv"
"""
Cars in the longest train that fits on each platform's tracks,
from a track map of unknown origin found at Railfan Guides of the U.S.:
https://www.railfanguides.us/ny/penntonewrochelle/PennStationLayout1.jpg
"""

PLATFORMS_OSM = DATA_DIR / "platforms_osm.csv"
"""
Each platform's outline's area, from OpenStreetMap,
via `platform-crowd-model data platforms-osm`.
"""


@cache
def platform_lengths() -> dict[int, int]:
    """Each platform's length (ft), from `PLATFORM_LENGTHS`."""
    with PLATFORM_LENGTHS.open() as f:
        return {int(row["platform"]): int(row["length_ft"]) for row in csv.DictReader(f)}


@cache
def platform_max_cars() -> dict[int, int]:
    """Cars in the longest train that fits on each platform's tracks, from `PLATFORM_MAX_CARS`."""
    with PLATFORM_MAX_CARS.open() as f:
        return {int(row["platform"]): int(row["max_cars"]) for row in csv.DictReader(f)}


@cache
def platform_areas() -> dict[int, int]:
    """Each platform's area (sq ft), from `PLATFORMS_OSM`."""
    with PLATFORMS_OSM.open() as f:
        return {
            int(row["platform"]): int(row["area_sq_ft"])
            for row in csv.DictReader(f)
            if row["platform"] and row["level"] == "-3"
        }


@dataclass
class Params:
    platform: Annotated[int, Field(name="Platform", units="#")]
    """Which platform it is, e.g. 3."""

    headway: Annotated[timedelta, Field(name="Headway", units="s")]
    """
    Time between trains' scheduled arrivals.
    The first arrives at 0 s, on one track, and the second on the other.
    """

    total_vce_width: Annotated[float, Field(name="Total VCE Width", units="ft")]
    """
    Total width (in feet) of all of the VCEs (vertical circulation elements) going upstairs.
    Per the ETA report, this excludes one VCE per platform,
    e.g. an escalator running the other way.
    """

    vce_widths: list[float]
    """Widths (in feet) of each VCE (vertical circulation element)."""

    trains: Annotated[int, Field(name="Trains", units="train")] = 4
    """Trains arriving, alternating between the platform's two tracks."""

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
        """Prefix of the filenames to save the time series and charts in, e.g. `platform3_recon`."""
        return f"platform{self.platform}" + (f"_{self.modifier}" if self.modifier else "")

    @property
    def platform_length(self) -> Annotated[int, Field(name="Platform Length", units="ft")]:
        """Platform length (in feet), from the Moynihan Station EA."""
        return platform_lengths()[self.platform]

    @property
    def platform_max_cars(self) -> Annotated[int, Field(name="Platform Max Cars", units="car")]:
        """Cars in the longest train that fits on its tracks, from `PLATFORM_MAX_CARS`."""
        return platform_max_cars()[self.platform]

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
    def platform_area(self) -> Annotated[int, Field(name="Platform Area", units="ft^2")]:
        """
        Platform area (in square feet), from OpenStreetMap's outline of it,
        which accounts for platforms tapering.
        """
        return platform_areas()[self.platform]

    @property
    def nfpa_130_exit_capacity(
        self,
    ) -> Annotated[float, Field(name="NFPA 130 Exit Capacity", units="pax/s")]:
        """
        How fast the VCEs can evacuate the platform under NFPA 130 (in pax/s),
        at `NFPA_130_EXIT_FLOW`.
        NFPA 130 also takes the widest escalator out of service,
        and lets escalators provide at most half of the capacity
        (TCQSM p. 10-52: https://onlinepubs.trb.org/onlinepubs/tcrp/tcrp_rpt_165ch-10.pdf#page=56),
        but we don't know which VCEs are escalators yet,
        so this counts all of `total_vce_width`, and so is optimistic.
        """
        return stair_flow(NFPA_130_EXIT_FLOW, self.total_vce_width)


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

    vce_widths = params.vce_widths

    if print_time_series:
        print("vce_widths = ", vce_widths)

    # Initialize counters
    arriving_pax_on_platform: float = 0
    trains = range(params.trains)
    arrival_times: list[timedelta | None] = [
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

    max_pax_in_stair_queues = (
        params.total_vce_width * assumptions.stair_queue_length / assumptions.stair_queue_space
    )
    capacity = stair_flow(assumptions.stair_capacity, params.total_vce_width)
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
        min_space_per_pax=space_per_pax(total_pax_on_platform, usable_area),
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
        up_rate = platform_clearance(arriving_pax_on_platform, params.total_vce_width, assumptions)
        gone_up += up_rate
        arriving_pax_on_platform -= up_rate
        if arriving_pax_on_platform < 0:
            arriving_pax_on_platform = 0
        total_pax_on_platform -= up_rate
        # Each train's boarders get a share of the stairs,
        # and so a share of the upward flow on them.
        boarder_fractions = [
            boarder_fraction(boarders_upstairs[train], boarders_upstairs) for train in trains
        ]
        down_rates = [
            platform_ingress(
                boarders_upstairs[train],
                params.total_vce_width * boarder_fractions[train],
                up_rate * boarder_fractions[train],
                assumptions,
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

        space = space_per_pax(total_pax_on_platform, usable_area)
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
                if train + 2 < params.trains:
                    arrival_times[train + 2] = max((train + 2) * params.headway, time_after)
        summary.max_pax_on_platform = max(summary.max_pax_on_platform, total_pax_on_platform)
        aboard = sum(
            remaining_arrivals[train] + new_pax[train]
            for train in trains
            if (arrival_time := arrival_times[train]) is not None
            and time_after >= arrival_time
            and summary.dwells[train] is None
        )
        summary.max_occupants = max(summary.max_occupants, total_pax_on_platform + aboard)
        summary.min_space_per_pax = min(summary.min_space_per_pax, space)

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
                platform_crowding=space,
                up_rate=up_rate,
                net_pax_flow_rate=net_pax_flow_rate,
                platform_crowd_los=platform_crowd_los(space, assumptions),
                egress_los=egress_crowd_los(params.total_vce_width, up_rate, assumptions),
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
    "Arrivals",
    "Dwell",
    "Taper time",
    "Clear time",
    "Boarded time",
    "Time at capacity",
    "Max up rate (pax/s)",
    "Max pax on platform",
    "Max density (pax/m²)",
    "NFPA 130 evacuation",
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

    evacuation_time = timedelta(
        seconds=math.ceil(summary.max_occupants / params.nfpa_130_exit_capacity)
    )
    evacuation_ok = "✓" if evacuation_time <= NFPA_130_PLATFORM_EVACUATION_TIME else "✗"

    def fmt_time(t: timedelta | None) -> str:
        """`t` as `m:ss`."""
        if t is None:
            return "never"
        minutes, rest = divmod(t, timedelta(minutes=1))
        return f"{minutes}:{rest.seconds:02}"

    return (
        f"| {params.name} | {fmt_time(headway)} | {fmt_ft_in(params.total_vce_width)}"
        f" | {', '.join(fmt_time(arrival) for arrival in summary.arrival_times)}"
        f" | {', '.join(fmt_time(dwell) for dwell in summary.dwells)}"
        f" | {fmt_time(summary.taper_time)} | {fmt_time(summary.clear_time)}"
        f" | {fmt_time(summary.boarded_time)}"
        f" | {fmt_time(summary.time_at_capacity)} | {summary.max_up_rate:.2f}"
        f" | {summary.max_pax_on_platform:.0f}"
        f" | {1 / (summary.min_space_per_pax * SQUARE_METERS_PER_SQUARE_FOOT):.2f}"
        f" ({platform_crowd_los(summary.min_space_per_pax, params.assumptions)})"
        f" | {fmt_time(evacuation_time)} {evacuation_ok} |"
    )


def main(update_readme: bool = False, charts: bool = False) -> None:
    """Run every scenario and print a table of their results."""

    # params are labeled  with p<platform number><time in seconds>
    # recon indicates that a platform was modelled accounting for penn reconstruction plans
    params_p3120 = Params(
        platform=3,
        headway=CLOSE_HEADWAY,
        total_vce_width=42.5,
        vce_widths=[w / 12 for w in (60, 60, 40, 54, 40, 54, 54, 54, 54, 54, 54)],
    )
    params_p3300 = Params(
        platform=3,
        headway=NORMAL_HEADWAY,
        total_vce_width=42.5,
        vce_widths=[w / 12 for w in (60, 60, 40, 54, 40, 54, 54, 54, 54, 54, 54)],
    )
    params_p3recon120 = Params(
        platform=3,
        modifier="recon",
        headway=CLOSE_HEADWAY,
        total_vce_width=44.75,
        vce_widths=[w / 12 for w in (60, 60, 40, 54, 40, 54, 54, 54, 54, 54, 54)],
    )
    params_p3recon300 = Params(
        platform=3,
        modifier="recon",
        headway=NORMAL_HEADWAY,
        total_vce_width=44.75,
        vce_widths=[w / 12 for w in (60, 60, 40, 54, 40, 54, 54, 54, 54, 54, 54)],
    )
    params_p60 = Params(
        platform=6,
        headway=timedelta(0),
        total_vce_width=48.168,
        vce_widths=[w / 12 for w in (60, 60, 40, 54, 40, 54, 54, 54, 54, 54, 54)],
    )
    params_p10120 = Params(
        platform=10,
        headway=CLOSE_HEADWAY,
        total_vce_width=70.58,
        vce_widths=[w / 12 for w in (60, 60, 40, 54, 40, 54, 54, 54, 54, 54, 54)],
    )
    params_p11120 = Params(
        platform=11,
        headway=CLOSE_HEADWAY,
        total_vce_width=43.58,
        vce_widths=[w / 12 for w in (60, 60, 40, 54, 40, 54, 54, 54, 54, 54, 54)],
    )
    params_p30 = dataclasses.replace(params_p3120, headway=timedelta(0))
    params_p3recon0 = dataclasses.replace(params_p3recon120, headway=timedelta(0))
    params_p6120 = dataclasses.replace(params_p60, headway=CLOSE_HEADWAY)
    params_p6300 = dataclasses.replace(params_p60, headway=NORMAL_HEADWAY)
    params_p100 = dataclasses.replace(params_p10120, headway=timedelta(0))
    params_p10300 = dataclasses.replace(params_p10120, headway=NORMAL_HEADWAY)
    params_p110 = dataclasses.replace(params_p11120, headway=timedelta(0))
    params_p11300 = dataclasses.replace(params_p11120, headway=NORMAL_HEADWAY)
    if charts:
        OUTPUT_DIR.mkdir(exist_ok=True)
    with ProcessPoolExecutor() as executor:
        rows = list(
            executor.map(
                functools.partial(run_model, charts=charts),
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
