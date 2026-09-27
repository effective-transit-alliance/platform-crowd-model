#!/usr/bin/env -S uv run

"""
This is a recursive peak-hour platform clearance calculator.
model from https://onlinepubs.trb.org/Onlinepubs/hrr/1971/355/355-001.pdf
"""

import dataclasses
import typing
from collections.abc import Generator
from dataclasses import dataclass
from typing import TYPE_CHECKING, Annotated, Any, Self, cast

import openpyxl
from openpyxl.cell import Cell
from openpyxl.chart import Reference, ScatterChart
from openpyxl.chart.series_factory import SeriesFactory
from openpyxl.worksheet.worksheet import Worksheet

if TYPE_CHECKING:
    from _typeshed import DataclassInstance

# Sources cited below:
# - Fruin, "Designing for Pedestrians: A Level-of-Service Concept" (1971):
#   https://onlinepubs.trb.org/Onlinepubs/hrr/1971/355/355-001.pdf
#   Page numbers are of the PDF.
# - TCQSM, 3rd edition, chapter 10:
#   https://onlinepubs.trb.org/onlinepubs/tcrp/tcrp_rpt_165ch-10.pdf
#   Page numbers are the manual's, and the PDF pages are in the links.

SECONDS_PER_MINUTE = 60


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

    for field in dataclasses.fields(cls):
        field_meta = Field.try_from_annotated(field.type)
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
class Assumptions:
    """
    Everything the model assumes, shared by every scenario,
    as opposed to the facts about each scenario in `Params`.
    The defaults are the model's current assumptions;
    override any of them to see how sensitive the results are to it,
    e.g. `Assumptions(stair_capacity=15)`.
    """

    door_flow_rate: Annotated[float, Field(name="Door Flow Rate", units="pax/s/door")] = 1.0
    """
    Alighting and boarding rate per single-door equivalent (pax/s/door).
    Assumes no delay for the doors to open, and the same rate both ways.
    From the ETA report.
    """

    stair_capacity: Annotated[float, Field(name="Stair Capacity", units="pax/min/ft")] = 17
    """
    Stair capacity, the LOS E/F boundary (pax/min per ft of width).
    Applied to all VCEs, even escalators, which are faster (pessimistic),
    and regardless of stair rise, which slows people down on long climbs (optimistic).
    Fruin, p. 14: https://onlinepubs.trb.org/Onlinepubs/hrr/1971/355/355-001.pdf#page=14
    """

    bidirectional_stair_flow_limit: Annotated[
        float, Field(name="Bidirectional Stair Flow Limit", units="pax/min/ft")
    ] = 10
    """
    Upward stair flow (pax/min per ft of width) above which nobody can come down,
    the LOS C/D boundary.
    Below it, both directions share `stair_capacity`.
    From the ETA report: "There is no bidirectional flow on stairwells if LOS is worse than C".
    """

    emergency_stair_flow: Annotated[
        float, Field(name="Emergency Stair Flow", units="pax/min/ft")
    ] = 19
    """
    Stair flow (pax/min per ft of width) used only for the printed emergency egress time.
    Fruin's maximum ascending stair flow is 18.9 (p. 9: https://onlinepubs.trb.org/Onlinepubs/hrr/1971/355/355-001.pdf#page=9),
    more than both `stair_capacity` and NFPA 130's 16.9
    (TCQSM p. 10-79: https://onlinepubs.trb.org/onlinepubs/tcrp/tcrp_rpt_165ch-10.pdf#page=83),
    so the emergency egress time is a lower bound.
    """

    stair_queue_space: Annotated[float, Field(name="Stair Queue Space", units="ft^2/pax")] = 5
    """
    Space per passenger queued at the stairs (ft^2/pax).
    Only used to report when the arrived passengers fit in the stair queues (the taper time).
    TCQSM p. 10-51: https://onlinepubs.trb.org/onlinepubs/tcrp/tcrp_rpt_165ch-10.pdf#page=55
    """

    stair_queue_length: Annotated[float, Field(name="Stair Queue Length", units="ft")] = 20
    """
    Length of the queue in front of each stair (ft),
    used to report when the arrived passengers start to taper off.
    From the ETA report.
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
    each grade needs more than this space per passenger (ft^2/pax), else F.
    The TCQSM only applies these to passengers waiting to board, not to everyone on the platform,
    so grading everyone with them is optimistic.
    TCQSM Exhibit 10-32, p. 10-55: https://onlinepubs.trb.org/onlinepubs/tcrp/tcrp_rpt_165ch-10.pdf#page=59
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
    then E up to `stair_capacity`, else F.
    Fruin, pp. 12-14: https://onlinepubs.trb.org/Onlinepubs/hrr/1971/355/355-001.pdf#page=12
    """

    simulation_length: Annotated[int, Field(name="Simulation Length", units="s")] = 600
    """Time to simulate (s), counted from the first train's arrival."""

    usable_platform_area_multiplier: Annotated[
        float, Field(name="Usable Platform Area Multiplier", units="fraction")
    ] = 0.75
    """
    Fraction of the platform's area usable by passengers,
    leaving the rest for columns, stairwells, and other obstructions.
    From the ETA report.
    Likely optimistic: the TCQSM's 18 in. edge buffers alone take about 17% of an 18 ft platform.
    """

    arriving_pax_per_train: Annotated[
        int, Field(name="Arriving Passengers per Train", units="pax")
    ] = 1620
    """
    Passengers arriving on each train (pax), all of whom alight.
    A crush-loaded 10-car or seated 12-car NJ Transit MultiLevel, from the ETA report.
    """

    doors_per_train: Annotated[int, Field(name="Doors per Train", units="door")] = 40
    """
    Doors (single-door equivalents) on each train on the platform side.
    A 10-car NJ Transit MultiLevel with 4 per car, the worst case.
    A 12-car LIRR train has more and better doors.
    """

    departing_pax_per_train: Annotated[
        int, Field(name="Departing Passengers per Train", units="pax")
    ] = 400
    """Passengers boarding each train (pax), from the ETA report."""

    departing_pax_on_platform_per_train: Annotated[
        int, Field(name="Departing Passengers on Platform per Train", units="pax")
    ] = 200
    """
    Of `departing_pax_per_train`, those already on the platform at the start (pax).
    The rest start upstairs, all at once, with none arriving during the simulation (optimistic).
    From the ETA report.
    """


CLOSE_HEADWAY = 120
"""Time between two trains' arrivals in the closely spaced scenarios (s), from the ETA report."""

NORMAL_HEADWAY = 300
"""Time between two trains' arrivals in the normal scenarios (s), from the ETA report."""


def stair_flow(rate: float, w: float) -> float:
    """
    :param rate: stair flow per foot of width (pax/min/ft)
    :param w: stair width (ft)
    :return: stair flow across the whole width (pax/s)
    """
    return rate * w / SECONDS_PER_MINUTE


# basic flow: train egress > platform crowd > VCE egress rate > back to
# platform crowd


# keep high VCE egress rate if queues at stairs are long
def alight_rate(k: float, t: float, t0: float, u: float) -> float:
    """
    :param k: number of people waiting to get off train
    :param t: time pass counter (s)
    :param t0: train arrival time
    :param u: doors * `Assumptions.door_flow_rate` (pax/s)
    :return: egress rate from train to platform across all doors (pax/s)
    """
    if t > t0:
        return min(k, u)
    else:
        return 0


def platform_clearance(karr: float, w: float, a: Assumptions) -> float:
    """
    Arrived passengers queue at the stairs, which discharge them at `a.stair_capacity`
    as long as anyone is queued.
    See the TCQSM's stair queuing procedure: https://onlinepubs.trb.org/onlinepubs/tcrp/tcrp_rpt_165ch-10.pdf#page=55

    Fruin's stair equation relates flow to the space per passenger *on the stair*,
    which a queued stair holds near its critical density,
    so it doesn't apply to the space per passenger on the platform.
    The few seconds of walking from the doors to the stairs are ignored.

    :param karr: number of arrived passengers on the platform (pax)
    :param w: total width of vertical circulation elements (ft)
    :return: platform egress rate on stairs (pax/s)
    """
    return min(karr, stair_flow(a.stair_capacity, w))


def platform_ingress(kdep: float, w: float, r_up: float, a: Assumptions) -> float:
    """
    Departing passengers queue upstairs and come down with whatever stair capacity
    the upward flow leaves.

    :param kdep: number of departing passengers upstairs (pax)
    :param w: width of vertical circulation elements available to these passengers (ft)
    :param: r_up: upstairs flow on those same vertical circulation elements (pax/s)
    :return: platform ingress rate on stairs (pax/s)
    """
    if r_up > stair_flow(a.bidirectional_stair_flow_limit, w):
        return 0
    return min(kdep, stair_flow(a.stair_capacity, w) - r_up)


def boarder_fraction(trainA_boarders: float, trainB_boarders: float) -> float:
    if trainA_boarders + trainB_boarders > 0:
        return trainA_boarders / (trainB_boarders + trainA_boarders)
    else:
        return 1


def board_rate(
    r_max: float,
    r_off: float,
    sim_t: float,
    arr_t: float,
    dep_t: float,
    boarders: float,
) -> float:
    """
    :param r_max: maximum train board rate, doors * `Assumptions.door_flow_rate` (pax/s)
    :param r_off: train alight rate, pass alight_rate_fn;
    nobody boards until everyone has alighted, i.e. the second after this is last nonzero
    :param sim_t: time in seconds, pass counter
    :param arr_t: train arrival time
    :param: dep_t: train departure time
    :param: boarders: number of passengers waiting on platform to board,
    pass departing_pax_on_plat
    :return: train ingress rate across all doors (pax/s)
    """
    if arr_t < sim_t < dep_t and r_off == 0:
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


def platform_crowd_los(inst_crowding: float, a: Assumptions) -> str:
    """
    Fruin's LOS for queuing and waiting areas, like a platform, not for walkways.
    See the TCQSM's Exhibit 10-32: https://onlinepubs.trb.org/onlinepubs/tcrp/tcrp_rpt_165ch-10.pdf#page=59

    :param inst_crowding: space per passenger (ft^2/pax)
    """
    for grade, min_space in a.platform_los_min_space:
        if inst_crowding > min_space:
            return grade
    return "F"


def egress_crowd_los(w: float, plat_egress_rate: float, a: Assumptions) -> str:
    """
    Fruin's LOS for stairs: https://onlinepubs.trb.org/Onlinepubs/hrr/1971/355/355-001.pdf#page=12

    :param w: total width of vertical circulation elements (ft)
    :param plat_egress_rate: upward stair flow (pax/s)
    """
    for grade, max_flow in (*a.stair_los_max_flow, ("E", a.stair_capacity)):
        if plat_egress_rate <= stair_flow(max_flow, w):
            return grade
    return "F"


@dataclass
class Params:
    filename_prefix: str
    """Prefix of filename to save the spreadsheet in."""

    platform_width: Annotated[int, Field(name="Platform Width", units="ft")]
    """Platform width (in feet)."""

    platform_length: Annotated[int, Field(name="Platform Length", units="ft")]
    """Platform length (in feet)."""

    train1_arrival_time: Annotated[int, Field(name="Train 1 Arrival Time", units="s")]
    """Time (in seconds) when train 1 arrives."""

    train2_arrival_time: Annotated[int, Field(name="Train 2 Arrival Time", units="s")]
    """Time (in seconds) when train 2 arrives."""

    total_vce_width: Annotated[float, Field(name="Total VCE Width", units="ft")]
    """
    Total width (in feet) of all of the VCEs (vertical circulation elements) going upstairs.
    Per the ETA report, this excludes one VCE per platform,
    e.g. an escalator running the other way (pessimistic).
    """

    assumptions: Assumptions = dataclasses.field(default_factory=Assumptions)
    """What the model assumes, the same for every scenario unless overridden."""

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
        """Emergency egress time (in seconds)."""
        return 2 * self.assumptions.arriving_pax_per_train / self.los_f_egress_rate


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

    train1_pax: Annotated[float, Field(name="Train 1 Passengers", units="pax")]
    """Train 1 number of passengers."""

    train2_pax: Annotated[float, Field(name="Train 2 Passengers", units="pax")]
    """Train 2 number of passengers."""

    train1_off_rate: Annotated[float, Field(name="Train 1 Alighting Rate", units="pax/s")]
    """Train 1 alighting rate (in pax/s)."""

    train2_off_rate: Annotated[float, Field(name="Train 2 Alighting Rate", units="pax/s")]
    """Train 2 alighting rate (in pax/s)."""

    train1_on_rate: Annotated[float, Field(name="Train 1 Boarding Rate", units="pax/s")]
    """Train 1 boarding rate (in pax/s)."""

    train2_on_rate: Annotated[float, Field(name="Train 2 Boarding Rate", units="pax/s")]
    """Train 2 boarding rate (in pax/s)."""

    down_rate: Annotated[float, Field(name="Downstairs Rate", units="pax/s")]
    """Downstairs rate (in pax/s)."""

    up_rate: Annotated[float, Field(name="Upstairs Rate", units="pax/s")]
    """Upstairs rate (in pax/s)."""

    train1_departing_pax_on_platform: Annotated[
        float, Field(name="Train 1 Departing Passengers on Platform", units="pax")
    ]
    """Train 1 departing passengers on platform."""

    train2_departing_pax_on_platform: Annotated[
        float, Field(name="Train 2 Departing Passengers on Platform", units="pax")
    ]
    """Train 2 departing passengers on platform."""

    arrived_pax_waiting_on_platform: Annotated[
        float, Field(name="Arrived Passengers on Platform", units="pax")
    ]
    """Number of passengers who arrived on platform."""

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
    """Seconds the upstairs rate is at the VCEs' `Assumptions.stair_capacity`."""

    taper_time: int | None
    """
    Last second the arrived passengers on the platform exceed what fits in the stair queues,
    i.e. when they start to taper off, or `None` if they never do.
    """

    clear_time: int | None
    """First second after the last arrival when all arrived passengers have left the platform."""

    boarded_time: int | None
    """First second when all departing passengers have boarded, or `None` if they never do."""

    max_pax_on_platform: float
    """Most passengers on the platform at once."""

    min_space_per_pax: float
    """Least platform space per passenger (sqft)."""


def calc_workbook(params: Params) -> tuple[openpyxl.Workbook, Summary]:
    a = params.assumptions
    eff_area = params.platform_width * params.platform_length * a.usable_platform_area_multiplier

    # Initialize counters
    arrived_pax_waiting_on_plat: float = 0
    train1_remaining_arrivals = float(a.arriving_pax_per_train)
    train2_remaining_arrivals = float(a.arriving_pax_per_train)
    train1_new_pax: float = 0
    train2_new_pax: float = 0
    train1_boarders_upstairs = float(
        a.departing_pax_per_train - a.departing_pax_on_platform_per_train
    )
    train2_boarders_upstairs = train1_boarders_upstairs
    train1_boarders_on_plat = float(a.departing_pax_on_platform_per_train)
    train2_boarders_on_plat = train1_boarders_on_plat
    total_pax_on_platform = train1_boarders_on_plat + train2_boarders_on_plat
    wb = openpyxl.Workbook()

    sheet = active_worksheet(wb)

    writable_cell(sheet, column=1, row=1).value = "Parameter"
    writable_cell(sheet, column=2, row=1).value = "Value"

    param_values = [*annotated_field_values(params), *annotated_field_values(params.assumptions)]
    for i, (_attr, value, field) in enumerate(param_values):
        writable_cell(sheet, column=1, row=i + 2).value = field.description
        writable_cell(sheet, column=2, row=i + 2).value = value

    FIRST_DATA_ROW = 2

    # The parameters take up columns 1 (A) and 2 (B), so the time series starts after them.
    FIRST_DATA_COLUMN = 3

    qmax = params.total_vce_width * a.stair_queue_length / a.stair_queue_space
    capacity = stair_flow(a.stair_capacity, params.total_vce_width)
    last_arrival_time = max(params.train1_arrival_time, params.train2_arrival_time)
    summary = Summary(
        max_up_rate=0,
        secs_at_capacity=0,
        taper_time=None,
        clear_time=None,
        boarded_time=None,
        max_pax_on_platform=total_pax_on_platform,
        min_space_per_pax=space_per_pax(total_pax_on_platform, eff_area),
    )

    print("Elapsed_Time", "Train_1_Pax", "Train_2_Pax")

    def get_column_for(attr_name: str) -> int:
        for i, (attr, _field) in enumerate(annotated_field_names(Instant)):
            if attr == attr_name:
                return FIRST_DATA_COLUMN + i
        raise AttributeError(Instant, attr_name)

    for time_after in range(0, a.simulation_length):
        train1_off_rate = alight_rate(
            train1_remaining_arrivals,
            time_after,
            params.train1_arrival_time,
            a.doors_per_train * a.door_flow_rate,
        )
        train1_remaining_arrivals -= train1_off_rate
        if train1_remaining_arrivals < 0:
            train1_remaining_arrivals = 0
        train2_off_rate = alight_rate(
            train2_remaining_arrivals,
            time_after,
            params.train2_arrival_time,
            a.doors_per_train * a.door_flow_rate,
        )
        train2_remaining_arrivals -= train2_off_rate
        if train2_remaining_arrivals < 0:
            train2_remaining_arrivals = 0
        total_pax_on_platform += train1_off_rate + train2_off_rate
        arrived_pax_waiting_on_plat += train1_off_rate + train2_off_rate
        plat_egress_rate = platform_clearance(
            arrived_pax_waiting_on_plat, params.total_vce_width, a
        )
        arrived_pax_waiting_on_plat -= plat_egress_rate
        if arrived_pax_waiting_on_plat < 0:
            arrived_pax_waiting_on_plat = 0
        total_pax_on_platform -= plat_egress_rate
        # Each train's boarders get a share of the stairs,
        # and so a share of the upward flow on them.
        train1_boarder_frac = boarder_fraction(train1_boarders_upstairs, train2_boarders_upstairs)
        train2_boarder_frac = boarder_fraction(train2_boarders_upstairs, train1_boarders_upstairs)
        plat_ingress_rate_1 = platform_ingress(
            train1_boarders_upstairs,
            params.total_vce_width * train1_boarder_frac,
            plat_egress_rate * train1_boarder_frac,
            a,
        )

        plat_ingress_rate_2 = platform_ingress(
            train2_boarders_upstairs,
            params.total_vce_width * train2_boarder_frac,
            plat_egress_rate * train2_boarder_frac,
            a,
        )
        train1_boarders_on_plat += plat_ingress_rate_1
        train2_boarders_on_plat += plat_ingress_rate_2
        total_pax_on_platform += plat_ingress_rate_1
        total_pax_on_platform += plat_ingress_rate_2
        train1_on_rate = board_rate(
            a.doors_per_train * a.door_flow_rate,
            train1_off_rate,
            time_after,
            params.train1_arrival_time,
            a.simulation_length,
            train1_boarders_on_plat,
        )
        train2_on_rate = board_rate(
            a.doors_per_train * a.door_flow_rate,
            train2_off_rate,
            time_after,
            params.train2_arrival_time,
            a.simulation_length,
            train2_boarders_on_plat,
        )

        train1_boarders_on_plat -= train1_on_rate

        train2_boarders_on_plat -= train2_on_rate

        total_pax_on_platform -= train1_on_rate

        total_pax_on_platform -= train2_on_rate

        train1_boarders_upstairs -= plat_ingress_rate_1

        train2_boarders_upstairs -= plat_ingress_rate_2

        train1_new_pax += train1_on_rate

        train2_new_pax += train2_on_rate

        inst_crowding = space_per_pax(total_pax_on_platform, eff_area)
        if total_pax_on_platform < 0:
            total_pax_on_platform = 0
        if train1_boarders_on_plat < 0:
            train1_boarders_on_plat = 0
        if train2_boarders_on_plat < 0:
            train2_boarders_on_plat = 0
        if arrived_pax_waiting_on_plat < 0:
            arrived_pax_waiting_on_plat = 0
        print(
            time_after,
            train1_remaining_arrivals + train1_new_pax,
            train2_remaining_arrivals + train2_new_pax,
            arrived_pax_waiting_on_plat,
            plat_egress_rate,
        )
        """
        print(
            "At time " + str(time_after) + " s,",
            str(train1_remaining_arrivals) + " wait to alight train 1;",
            str(train2_remaining_arrivals) + " wait to alight train 2;",
            str(train1_on_rate - train1_off_rate) + " pax/s train 1 net rate;",
            str(train2_on_rate - train2_off_rate) + " pax/s train 2 net rate;",
        )
        print(
            str(int(arrived_pax_waiting_on_plat))
            + " deboarded pax on platform;",
            str(int(train1_boarders_on_plat + int(train2_boarders_on_plat)))
            + " boarding pax on platform;",
            str(int(total_pax_on_platform)) + " total pax on platform;",
            str(int(inst_crowding)) + " sqft per pax;",
        )
        print(
            str(plat_egress_rate) + " pax/s up;",
            str((plat_ingress_rate_1 + plat_ingress_rate_2)) + "pax/s down;",
            str((train1_boarders_upstairs + train2_boarders_upstairs))
            + " pax are upstairs"
        )
        """
        summary.max_up_rate = max(summary.max_up_rate, plat_egress_rate)
        if plat_egress_rate >= capacity - 1e-9:
            summary.secs_at_capacity += 1
        if arrived_pax_waiting_on_plat > qmax:
            summary.taper_time = time_after
        if (
            summary.clear_time is None
            and time_after > last_arrival_time
            and arrived_pax_waiting_on_plat < 1
        ):
            summary.clear_time = time_after
        if (
            summary.boarded_time is None
            and train1_boarders_upstairs
            + train2_boarders_upstairs
            + train1_boarders_on_plat
            + train2_boarders_on_plat
            < 1
        ):
            summary.boarded_time = time_after
        summary.max_pax_on_platform = max(summary.max_pax_on_platform, total_pax_on_platform)
        summary.min_space_per_pax = min(summary.min_space_per_pax, inst_crowding)

        instant = Instant(
            time=time_after,
            train1_pax=train1_remaining_arrivals + train1_new_pax,
            train2_pax=train2_remaining_arrivals + train2_new_pax,
            arrived_pax_waiting_on_platform=arrived_pax_waiting_on_plat,
            train1_off_rate=train1_off_rate,
            train2_off_rate=train2_off_rate,
            train1_on_rate=train1_on_rate,
            train2_on_rate=train2_on_rate,
            down_rate=plat_ingress_rate_1 + plat_ingress_rate_2,
            train1_departing_pax_on_platform=train1_boarders_on_plat,
            train2_departing_pax_on_platform=train2_boarders_on_plat,
            total_pax_on_platform=total_pax_on_platform,
            platform_crowding=inst_crowding,
            up_rate=plat_egress_rate,
            net_pax_flow_rate=(
                plat_ingress_rate_1
                + plat_ingress_rate_2
                + train1_off_rate
                + train2_off_rate
                - plat_egress_rate
                - train1_on_rate
                - train2_on_rate
            ),
            platform_crowd_los=platform_crowd_los(inst_crowding, a),
            egress_los=egress_crowd_los(params.total_vce_width, plat_egress_rate, a),
        )

        for i, (_attr, value, field) in enumerate(annotated_field_values(instant)):
            column = FIRST_DATA_COLUMN + i
            writable_cell(sheet, row=1, column=column).value = field.description
            writable_cell(sheet, row=instant.time + 2, column=column).value = value

    def make_chart(title: str, min_col: int, x_title: str, y_title: str) -> ScatterChart:
        chart = ScatterChart()
        chart.title = title
        chart.style = 13
        chart.x_axis.title = x_title
        chart.y_axis.title = y_title
        chart.x_axis.scaling.min = 0
        chart.x_axis.scaling.max = a.simulation_length
        chart.legend = None

        max_row = a.simulation_length + FIRST_DATA_ROW - 1
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
        chart.x_axis.scaling.min = 0
        chart.x_axis.scaling.max = a.simulation_length
        chart.y_axis.scaling.min = 0
        chart.y_axis.scaling.max = 50
        chart.legend = None

        max_row = a.simulation_length + FIRST_DATA_ROW - 1
        xvalues = Reference(
            sheet, min_col=get_column_for("time"), min_row=FIRST_DATA_ROW, max_row=max_row
        )
        values = Reference(sheet, min_col=min_col, min_row=FIRST_DATA_ROW - 1, max_row=max_row)
        # Y values start one row above X values so that first cell is series name.
        series = SeriesFactory(values, xvalues, title_from_data=True)
        chart.series.append(series)
        return chart

    def make_chart_2(title: str, col1: int, col2: int, x_title: str, y_title: str) -> ScatterChart:
        chart = ScatterChart()
        chart.title = title
        chart.style = 13
        chart.x_axis.title = x_title
        chart.y_axis.title = y_title
        chart.x_axis.scaling.min = 0
        chart.x_axis.scaling.max = a.simulation_length
        assert chart.legend is not None
        chart.legend.position = "b"

        max_row = a.simulation_length + FIRST_DATA_ROW - 1
        xvalues = Reference(
            sheet, min_col=get_column_for("time"), min_row=FIRST_DATA_ROW, max_row=max_row
        )
        values1 = Reference(sheet, min_col=col1, min_row=FIRST_DATA_ROW - 1, max_row=max_row)
        values2 = Reference(sheet, min_col=col2, min_row=FIRST_DATA_ROW - 1, max_row=max_row)
        # Y values start one row above X values so that first cell is series name.
        series1 = SeriesFactory(values1, xvalues, title_from_data=True)
        chart.series.append(series1)
        series2 = SeriesFactory(values2, xvalues, title_from_data=True)
        chart.series.append(series2)
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
            get_column_for("train1_pax"),
            get_column_for("train2_pax"),
            "Time (s)",
            "Passengers",
        ),
        "V19",
    )
    sheet.add_chart(
        make_chart_2(
            "Passengers on Platform",
            get_column_for("arrived_pax_waiting_on_platform"),
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
            "Space per passenger (sqft)",
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
    print(
        f"LOS F egress rate is {params.los_f_egress_rate} pax/s. "
        f"Emergency egress time is {params.emergency_egress_time} seconds."
    )
    return wb, summary


RESULTS_COLUMNS = [
    "Platform",
    "Headway",
    "VCE width",
    "Max up rate (pax/s)",
    "Time at capacity",
    "Taper time",
    "Clear time",
    "Boarded time",
    "Max pax on platform",
    "Min space/pax (sqft)",
]
RESULTS_HEADER = "| " + " | ".join(RESULTS_COLUMNS) + " |\n" + "|---" * len(RESULTS_COLUMNS) + "|"


def run_model(params: Params) -> str:
    """Run the model, save its spreadsheet, and return its row of the results table."""
    wb, summary = calc_workbook(params=params)

    headway = params.train2_arrival_time - params.train1_arrival_time
    wb.save(
        f"{params.filename_prefix}"
        f"_{params.assumptions.arriving_pax_per_train}"
        f"_{params.assumptions.arriving_pax_per_train}"
        f"_{headway}s.xlsx"
    )
    wb.close()

    def fmt_time(t: int | None) -> str:
        return "never" if t is None else f"{t} s"

    return (
        f"| {params.filename_prefix} | {headway} s | {params.total_vce_width} ft"
        f" | {summary.max_up_rate:.2f} | {summary.secs_at_capacity} s"
        f" | {fmt_time(summary.taper_time)} | {fmt_time(summary.clear_time)}"
        f" | {fmt_time(summary.boarded_time)}"
        f" | {summary.max_pax_on_platform:.0f} | {summary.min_space_per_pax:.1f}"
        f" ({platform_crowd_los(summary.min_space_per_pax, params.assumptions)}) |"
    )


def main() -> None:
    # params are labeled  with p<platform number><time in seconds>
    # recon indicates that a platform was modelled accounting for penn reconstruction plans
    params_p3120 = Params(
        filename_prefix="platform3",
        platform_width=18,
        platform_length=900,
        train1_arrival_time=0,
        train2_arrival_time=CLOSE_HEADWAY,
        total_vce_width=42.5,
    )
    params_p3300 = Params(
        filename_prefix="platform3",
        platform_width=18,
        platform_length=900,
        train1_arrival_time=0,
        train2_arrival_time=NORMAL_HEADWAY,
        total_vce_width=42.5,
    )
    params_p3recon120 = Params(
        filename_prefix="platform3_recon",
        platform_width=18,
        platform_length=900,
        train1_arrival_time=0,
        train2_arrival_time=CLOSE_HEADWAY,
        total_vce_width=44.75,
    )
    params_p3recon300 = Params(
        filename_prefix="platform3_recon",
        platform_width=18,
        platform_length=900,
        train1_arrival_time=0,
        train2_arrival_time=NORMAL_HEADWAY,
        total_vce_width=44.75,
    )
    params_p60 = Params(
        filename_prefix="platform6",
        platform_width=15,
        platform_length=1100,
        train1_arrival_time=0,
        train2_arrival_time=0,
        total_vce_width=48.168,
    )
    params_p10120 = Params(
        filename_prefix="platform10",
        platform_width=42,
        platform_length=1100,
        train1_arrival_time=0,
        train2_arrival_time=CLOSE_HEADWAY,
        total_vce_width=70.58,
    )
    params_p11120 = Params(
        filename_prefix="platform11",
        platform_width=18,
        platform_length=1100,
        train1_arrival_time=0,
        train2_arrival_time=CLOSE_HEADWAY,
        total_vce_width=43.58,
    )
    rows = [
        run_model(params)
        for params in [
            params_p3120,
            params_p3300,
            params_p3recon120,
            params_p3recon300,
            params_p60,
            params_p10120,
            params_p11120,
        ]
    ]
    print()
    print(RESULTS_HEADER)
    for row in rows:
        print(row)


if __name__ == "__main__":
    main()
