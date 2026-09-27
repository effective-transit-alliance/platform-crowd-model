#!/usr/bin/env -S uv run

"""
This is a recursive peak-hour platform clearance calculator.
model from https://onlinepubs.trb.org/Onlinepubs/hrr/1971/355/355-001.pdf
"""

import dataclasses
import functools
import typing
from collections.abc import Generator
from concurrent.futures import ProcessPoolExecutor
from dataclasses import dataclass
from pathlib import Path
from typing import TYPE_CHECKING, Annotated, Any, Self, cast

import numpy as np
import openpyxl
import typer
from numpy.typing import NDArray
from openpyxl.cell import Cell
from openpyxl.chart import Reference, ScatterChart
from openpyxl.chart.series_factory import SeriesFactory
from openpyxl.worksheet.worksheet import Worksheet
from typer import Option

if TYPE_CHECKING:
    from _typeshed import DataclassInstance

SECONDS_PER_MINUTE = 60

SQUARE_METERS_PER_SQUARE_FOOT = 0.09290304

FRUIN_ASCENDING_STAIR_COEFFICIENTS = (111, 162)
"""
`(a, b)` in Fruin's equation for ascending stair flow, `P = (aM - b)/M^2`,
where `P` is the flow (pax/min per ft of stair width)
and `M` is the space per passenger on the stair (ft^2/pax).
Fruin, p. 9: https://onlinepubs.trb.org/Onlinepubs/hrr/1971/355/355-001.pdf#page=9
"""

CLOSE_HEADWAY = 120
"""Time between two trains' arrivals in the closely spaced scenarios (s), from the ETA report."""

NORMAL_HEADWAY = 300
"""Time between two trains' arrivals in the normal scenarios (s), from the ETA report."""


@dataclass(frozen=True)
class Assumptions:
    """
    Everything the model assumes, shared by every scenario,
    as opposed to the facts about each scenario in `Params`.
    The defaults are the model's current assumptions,
    the same as in the `penn-station-can-handle-the-load` tag that the ETA report used,
    including some that are bugs, documented in the README's "Known Bugs";
    override any of them to see how sensitive the results are to it,
    e.g. `Assumptions(stair_capacity=15)`.
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
    """

    arriving_pax_per_train: Annotated[
        int, Field(name="Arriving Passengers per Train", units="pax")
    ] = 1620
    """
    Passengers arriving on each train, all of whom alight.
    A seated 12-car NJ Transit train at 135 seats per car,
    from the Moynihan Station Development Project environmental assessment,
    chapter 4.4, Station Circulation Analysis, Tables 4.4-10 and 4.4-19:
    https://web.archive.org/web/20241011135133/https://cdn.esd.ny.gov/subsidiaries_projects/msdc/Data/NEPA/04_4%20StationPedCirculation.pdf#page=22
    https://web.archive.org/web/20241011135133/https://cdn.esd.ny.gov/subsidiaries_projects/msdc/Data/NEPA/04_4%20StationPedCirculation.pdf#page=47
    """

    departing_pax_per_train: Annotated[
        int, Field(name="Departing Passengers per Train", units="pax")
    ] = 400
    """Passengers boarding each train, from the ETA report."""

    departing_pax_on_platform_per_train: Annotated[
        int, Field(name="Departing Passengers on Platform per Train", units="pax")
    ] = 200
    """
    Of `departing_pax_per_train`, those already on the platform at the start.
    The rest start upstairs, all at once, with none arriving during the simulation.
    From the ETA report.
    """

    doors_per_train: Annotated[int, Field(name="Doors per Train", units="door")] = 40
    """
    Doors (single-door equivalents) on each train on the platform side.
    A 10-car NJ Transit MultiLevel with 4 per car, the worst case.
    A 12-car LIRR train has more and better doors.
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

    stair_queue_min_flow: Annotated[
        float, Field(name="Stair Queue Minimum Flow", units="pax/min/ft")
    ] = 10
    """
    Minimum upward stair flow while more arriving passengers are on the platform
    than fit in the stair queues, the LOS C/D boundary.
    From the ETA report.
    """

    stair_queue_space: Annotated[float, Field(name="Stair Queue Space", units="ft^2/pax")] = 5
    """
    Space per passenger queued at the stairs.
    TCQSM p. 10-51: https://onlinepubs.trb.org/onlinepubs/tcrp/tcrp_rpt_165ch-10.pdf#page=55
    """

    stair_queue_length: Annotated[float, Field(name="Stair Queue Length", units="ft")] = 20
    """
    Length of the queue in front of each stair.
    While more arriving passengers are on the platform than fit in these queues,
    the upward flow is at least `stair_queue_min_flow`;
    once they fit, they start to taper off.
    From the ETA report.
    """

    bidirectional_stair_flow_limit: Annotated[
        float, Field(name="Bidirectional Stair Flow Limit", units="pax/min/ft")
    ] = 12
    """
    Total stair flow in both directions, which the upward flow leaves for the downward flow.
    A bug: the ETA report says there's no bidirectional flow on stairs worse than LOS C,
    i.e. above 10 pax/min/ft.
    See `README.md#bidirectional-flow-stops-at-12-paxminft-not-10`.
    """

    concourse_area: Annotated[float, Field(name="Concourse Area", units="ft^2")] = 5000
    """
    Area of the concourse upstairs holding the departing passengers who haven't come down yet.
    A bug: their flow down comes from Fruin's equation applied to this area,
    so it slows as the concourse empties.
    See `README.md#downward-flow-slows-as-a-fixed-concourse-empties`.
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
        ("A", 35),
        ("B", 25),
        ("C", 15),
        ("D", 10),
        ("E", 5),
    )
    """
    Fruin's LOS for walkways:
    each grade needs more than this space per passenger (ft^2/pax), or else F.
    Fruin, p. 7: https://onlinepubs.trb.org/Onlinepubs/hrr/1971/355/355-001.pdf#page=7
    A bug: most passengers on a platform are standing and waiting, not walking,
    so the TCQSM grades platforms with Fruin's LOS for queuing and waiting areas.
    See `README.md#platform-crowding-is-graded-as-a-walkway`.
    """

    stair_los_max_flow: tuple[tuple[str, float], ...] = (
        ("A", 5),
        ("B", 7),
        ("C", 9.5),
        ("D", 13),
    )
    """
    Fruin's LOS for stairs:
    each grade allows at most this flow (pax/min per ft of width),
    then E up to `stair_capacity`, or else F.
    Fruin, pp. 12-14: https://onlinepubs.trb.org/Onlinepubs/hrr/1971/355/355-001.pdf#page=12
    A bug: Fruin puts the C/D boundary at 10, not 9.5.
    See `README.md#the-stair-los-cd-boundary-is-95-paxminft-not-10`.
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


def fruin_ascending_stair_flow(m: float) -> float:
    """
    Fruin's equation for ascending stair flow.

    :param m: space per passenger (ft^2/pax)
    :return: stair flow per foot of width (pax/min/ft)
    """
    a, b = FRUIN_ASCENDING_STAIR_COEFFICIENTS
    return (a * m - b) / m**2


def alight_rate(k: float, t: float, t0: float, u: float) -> float:
    """
    :param k: number of people waiting to get off train
    :param t: time pass counter (s)
    :param t0: train arrival time
    :param u: maximum alighting rate across all doors (pax/s)
    :return: egress rate from train to platform across all doors (pax/s)
    """
    if t > t0:
        return min(k, u)
    else:
        return 0


def platform_clearance(
    karr: float, area: float, w: float, max_pax_in_stair_queues: float, assumptions: Assumptions
) -> float:
    """
    Two bugs: Fruin's equation is used as if it gives pax/s across all of the stairs,
    not pax/min per ft of width, and with the platform's space per passenger,
    not the stair's.
    See the README's "Known Bugs".

    :param karr: number of arriving passengers on the platform heading upstairs
    :param area: usable platform area (ft^2)
    :param w: total width of vertical circulation elements (ft)
    :param max_pax_in_stair_queues: number of people that fit in the stair queues
    :return: platform egress rate on stairs (pax/s)
    """
    flow = min(
        stair_flow(assumptions.stair_capacity, w),
        fruin_ascending_stair_flow(area / max(1, karr)),
    )
    if karr <= max_pax_in_stair_queues:
        return min(karr, flow)
    else:
        return max(stair_flow(assumptions.stair_queue_min_flow, w), flow)


def platform_ingress(kdep: float, w: float, r_up: float, assumptions: Assumptions) -> float:
    """
    Several bugs: the whole upward flow is subtracted from each train's share of the stairs,
    Fruin's ascending equation is used for descending passengers
    with the same units bug as `platform_clearance`,
    and it's applied to a fixed concourse area.
    See the README's "Known Bugs".

    :param kdep: number of departing passengers upstairs
    :param w: this train's share of the total width of vertical circulation elements (ft)
    :param r_up: upward stair flow (pax/s)
    :return: platform ingress rate on stairs (pax/s)
    """
    if kdep > 0:
        return min(
            kdep,
            min(
                max(0, stair_flow(assumptions.bidirectional_stair_flow_limit, w) - r_up),
                max(
                    0, fruin_ascending_stair_flow(assumptions.concourse_area / max(1, kdep)) - r_up
                ),
            ),
        )
    else:
        return 0


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
    A bug: boarding uses whatever door capacity alighting leaves in the same second,
    though the ETA report says nobody boards until everyone has alighted.
    See the README's "Known Bugs".

    :param r_max: maximum boarding rate across all doors (pax/s)
    :param r_off: train alight rate (pax/s)
    :param sim_t: time (s)
    :param arr_t: train arrival time (s)
    :param dep_t: train departure time (s)
    :param boarders: number of passengers waiting on platform to board
    :return: train ingress rate across all doors (pax/s)
    """
    if arr_t < sim_t < dep_t:
        return min(r_max - r_off, boarders)
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
    :param w: total width of vertical circulation elements (ft)
    :param plat_egress_rate: upward stair flow (pax/s)
    :return: its LOS, per `Assumptions.stair_los_max_flow` and `stair_capacity`
    """
    for grade, max_flow in (*assumptions.stair_los_max_flow, ("E", assumptions.stair_capacity)):
        if plat_egress_rate <= stair_flow(max_flow, w):
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
        yield attr, getattr(obj, attr), field


@dataclass
class Params:
    platform: Annotated[int, Field(name="Platform", units="#")]
    """Which platform it is, e.g. 3."""

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
    e.g. an escalator running the other way.
    """

    vce_widths: NDArray[np.floating]
    """Widths (in feet) of each VCE (vertical circulation element)."""

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
    """Seconds the upstairs rate is at the VCEs' LOS E capacity (17 pax/min/ft)."""

    taper_time: int | None
    """
    Last second the arriving passengers on the platform exceed what fits in the stair queues,
    i.e. when they start to taper off, or `None` if they never do.
    """

    clear_time: int | None
    """First second after the last arrival when all arriving passengers have left the platform."""

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
    door_rate = assumptions.doors_per_train * assumptions.door_flow_rate

    www = params.vce_widths[0, :]

    if print_time_series:
        print("www = ", www)

    # Initialize counters
    arriving_pax_waiting_on_plat: float = 0
    train1_remaining_arrivals = float(assumptions.arriving_pax_per_train)
    train2_remaining_arrivals = float(assumptions.arriving_pax_per_train)
    train1_new_pax: float = 0
    train2_new_pax: float = 0
    boarders_upstairs = float(
        assumptions.departing_pax_per_train - assumptions.departing_pax_on_platform_per_train
    )
    train1_boarders_upstairs = boarders_upstairs
    train2_boarders_upstairs = boarders_upstairs
    train1_boarders_on_plat = float(assumptions.departing_pax_on_platform_per_train)
    train2_boarders_on_plat = float(assumptions.departing_pax_on_platform_per_train)
    total_pax_on_platform = train1_boarders_on_plat + train2_boarders_on_plat
    wb = openpyxl.Workbook()

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
    capacity = stair_flow(assumptions.stair_capacity, params.total_vce_width)
    last_arrival_time = max(params.train1_arrival_time, params.train2_arrival_time)
    summary = Summary(
        max_up_rate=0,
        secs_at_capacity=0,
        taper_time=None,
        clear_time=None,
        dwells=[None, None],
        boarded_time=None,
        max_pax_on_platform=total_pax_on_platform,
        min_space_per_pax=space_per_pax(total_pax_on_platform, eff_area),
    )

    if print_time_series:
        print("Elapsed_Time", "Train_1_Pax", "Train_2_Pax")

    def get_column_for(attr_name: str) -> int:
        for i, (attr, _field) in enumerate(annotated_field_names(Instant)):
            if attr == attr_name:
                return FIRST_DATA_COLUMN + i
        raise AttributeError(Instant, attr_name)

    for time_after in range(0, assumptions.simulation_length):
        train1_off_rate = alight_rate(
            train1_remaining_arrivals,
            time_after,
            params.train1_arrival_time,
            door_rate,
        )
        train1_remaining_arrivals -= train1_off_rate
        if train1_remaining_arrivals < 0:
            train1_remaining_arrivals = 0
        train2_off_rate = alight_rate(
            train2_remaining_arrivals,
            time_after,
            params.train2_arrival_time,
            door_rate,
        )
        train2_remaining_arrivals -= train2_off_rate
        if train2_remaining_arrivals < 0:
            train2_remaining_arrivals = 0
        total_pax_on_platform += train1_off_rate + train2_off_rate
        arriving_pax_waiting_on_plat += train1_off_rate + train2_off_rate
        plat_egress_rate = platform_clearance(
            arriving_pax_waiting_on_plat,
            eff_area,
            params.total_vce_width,
            max_pax_in_stair_queues,
            assumptions,
        )
        arriving_pax_waiting_on_plat -= plat_egress_rate
        if arriving_pax_waiting_on_plat < 0:
            arriving_pax_waiting_on_plat = 0
        total_pax_on_platform -= plat_egress_rate
        plat_ingress_rate_1 = platform_ingress(
            train1_boarders_upstairs,
            params.total_vce_width
            * boarder_fraction(train1_boarders_upstairs, train2_boarders_upstairs),
            plat_egress_rate,
            assumptions,
        )

        plat_ingress_rate_2 = platform_ingress(
            train2_boarders_upstairs,
            params.total_vce_width
            * boarder_fraction(train2_boarders_upstairs, train1_boarders_upstairs),
            plat_egress_rate,
            assumptions,
        )
        train1_boarders_on_plat += plat_ingress_rate_1
        train2_boarders_on_plat += plat_ingress_rate_2
        total_pax_on_platform += plat_ingress_rate_1
        total_pax_on_platform += plat_ingress_rate_2
        train1_on_rate = board_rate(
            door_rate,
            train1_off_rate,
            time_after,
            params.train1_arrival_time,
            assumptions.simulation_length,
            train1_boarders_on_plat,
        )
        train2_on_rate = board_rate(
            door_rate,
            train2_off_rate,
            time_after,
            params.train2_arrival_time,
            assumptions.simulation_length,
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
        if arriving_pax_waiting_on_plat < 0:
            arriving_pax_waiting_on_plat = 0
        if print_time_series:
            print(
                time_after,
                train1_remaining_arrivals + train1_new_pax,
                train2_remaining_arrivals + train2_new_pax,
                arriving_pax_waiting_on_plat,
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
            str(int(arriving_pax_waiting_on_plat))
            + " deboarded pax on platform;",
            str(int(train1_boarders_on_plat + int(train2_boarders_on_plat)))
            + " boarding pax on platform;",
            str(int(total_pax_on_platform)) + " total pax on platform;",
            str(int(inst_crowding)) + " sq ft per pax;",
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
        if arriving_pax_waiting_on_plat > max_pax_in_stair_queues:
            summary.taper_time = time_after
        if (
            summary.clear_time is None
            and time_after > last_arrival_time
            and arriving_pax_waiting_on_plat < 1
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
        for train, (arrival_time, remaining_arrivals, upstairs, on_plat) in enumerate(
            (
                (
                    params.train1_arrival_time,
                    train1_remaining_arrivals,
                    train1_boarders_upstairs,
                    train1_boarders_on_plat,
                ),
                (
                    params.train2_arrival_time,
                    train2_remaining_arrivals,
                    train2_boarders_upstairs,
                    train2_boarders_on_plat,
                ),
            )
        ):
            if (
                summary.dwells[train] is None
                and time_after > arrival_time
                and remaining_arrivals < 1
                and upstairs + on_plat < 1
            ):
                summary.dwells[train] = time_after - arrival_time
        summary.max_pax_on_platform = max(summary.max_pax_on_platform, total_pax_on_platform)
        summary.min_space_per_pax = min(summary.min_space_per_pax, inst_crowding)

        if write_workbook:
            instant = Instant(
                time=time_after,
                train1_pax=train1_remaining_arrivals + train1_new_pax,
                train2_pax=train2_remaining_arrivals + train2_new_pax,
                arriving_pax_waiting_on_platform=arriving_pax_waiting_on_plat,
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
                platform_crowd_los=platform_crowd_los(inst_crowding, assumptions),
                egress_los=egress_crowd_los(params.total_vce_width, plat_egress_rate, assumptions),
            )

            for i, (_attr, value, field) in enumerate(annotated_field_values(instant)):
                column = FIRST_DATA_COLUMN + i
                writable_cell(sheet, row=1, column=column).value = field.description
                writable_cell(sheet, row=instant.time + 2, column=column).value = value

    if not write_workbook:
        return wb, summary

    def make_chart(title: str, min_col: int, x_title: str, y_title: str) -> ScatterChart:
        chart = ScatterChart()
        chart.title = title
        chart.style = 13
        chart.x_axis.title = x_title
        chart.y_axis.title = y_title
        chart.x_axis.scaling.min = 0
        chart.x_axis.scaling.max = assumptions.simulation_length
        chart.legend = None

        max_row = assumptions.simulation_length + FIRST_DATA_ROW - 1
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
        chart.x_axis.scaling.max = assumptions.simulation_length
        chart.y_axis.scaling.min = 0
        chart.y_axis.scaling.max = 50
        chart.legend = None

        max_row = assumptions.simulation_length + FIRST_DATA_ROW - 1
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
        chart.x_axis.scaling.max = assumptions.simulation_length
        assert chart.legend is not None
        chart.legend.position = "b"

        max_row = assumptions.simulation_length + FIRST_DATA_ROW - 1
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
    return wb, summary


RESULTS_COLUMNS = [
    "Platform",
    "Headway",
    "VCE width",
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


def run_model(params: Params, spreadsheets: bool) -> str:
    """
    Run the model, return its row of the results table,
    and with `spreadsheets`, print its time series and save its spreadsheet.
    """
    wb, summary = calc_workbook(
        params=params, write_workbook=spreadsheets, print_time_series=spreadsheets
    )

    headway = params.train2_arrival_time - params.train1_arrival_time
    if spreadsheets:
        wb.save(
            f"{params.filename_prefix}"
            f"_{params.assumptions.arriving_pax_per_train}"
            f"_{params.assumptions.arriving_pax_per_train}"
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
        f"| {params.name} | {fmt_time(headway)} | {params.total_vce_width} ft"
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
        train1_arrival_time=0,
        train2_arrival_time=CLOSE_HEADWAY,
        total_vce_width=42.5,
        vce_widths=(
            1
            / 12
            * np.transpose(
                np.array(
                    [
                        [60, 1],
                        [60, 1],
                        [40, 1],
                        [54, 1],
                        [40, 1],
                        [54, 1],
                        [54, 1],
                        [54, 1],
                        [54, 1],
                        [54, 1],
                        [54, 1],
                    ]
                )
            )
        ),
    )
    params_p3300 = Params(
        platform=3,
        platform_width=18,
        platform_length=900,
        train1_arrival_time=0,
        train2_arrival_time=NORMAL_HEADWAY,
        total_vce_width=42.5,
        vce_widths=(
            1
            / 12
            * np.transpose(
                np.array(
                    [
                        [60, 1],
                        [60, 1],
                        [40, 1],
                        [54, 1],
                        [40, 1],
                        [54, 1],
                        [54, 1],
                        [54, 1],
                        [54, 1],
                        [54, 1],
                        [54, 1],
                    ]
                )
            )
        ),
    )
    params_p3recon120 = Params(
        platform=3,
        modifier="recon",
        platform_width=18,
        platform_length=900,
        train1_arrival_time=0,
        train2_arrival_time=CLOSE_HEADWAY,
        total_vce_width=44.75,
        vce_widths=(
            1
            / 12
            * np.transpose(
                np.array(
                    [
                        [60, 1],
                        [60, 1],
                        [40, 1],
                        [54, 1],
                        [40, 1],
                        [54, 1],
                        [54, 1],
                        [54, 1],
                        [54, 1],
                        [54, 1],
                        [54, 1],
                    ]
                )
            )
        ),
    )
    params_p3recon300 = Params(
        platform=3,
        modifier="recon",
        platform_width=18,
        platform_length=900,
        train1_arrival_time=0,
        train2_arrival_time=NORMAL_HEADWAY,
        total_vce_width=44.75,
        vce_widths=(
            1
            / 12
            * np.transpose(
                np.array(
                    [
                        [60, 1],
                        [60, 1],
                        [40, 1],
                        [54, 1],
                        [40, 1],
                        [54, 1],
                        [54, 1],
                        [54, 1],
                        [54, 1],
                        [54, 1],
                        [54, 1],
                    ]
                )
            )
        ),
    )
    params_p60 = Params(
        platform=6,
        platform_width=15,
        platform_length=1100,
        train1_arrival_time=0,
        train2_arrival_time=0,
        total_vce_width=48.168,
        vce_widths=(
            1
            / 12
            * np.transpose(
                np.array(
                    [
                        [60, 1],
                        [60, 1],
                        [40, 1],
                        [54, 1],
                        [40, 1],
                        [54, 1],
                        [54, 1],
                        [54, 1],
                        [54, 1],
                        [54, 1],
                        [54, 1],
                    ]
                )
            )
        ),
    )
    params_p10120 = Params(
        platform=10,
        platform_width=42,
        platform_length=1100,
        train1_arrival_time=0,
        train2_arrival_time=CLOSE_HEADWAY,
        total_vce_width=70.58,
        vce_widths=(
            1
            / 12
            * np.transpose(
                np.array(
                    [
                        [60, 1],
                        [60, 1],
                        [40, 1],
                        [54, 1],
                        [40, 1],
                        [54, 1],
                        [54, 1],
                        [54, 1],
                        [54, 1],
                        [54, 1],
                        [54, 1],
                    ]
                )
            )
        ),
    )
    params_p11120 = Params(
        platform=11,
        platform_width=18,
        platform_length=1100,
        train1_arrival_time=0,
        train2_arrival_time=CLOSE_HEADWAY,
        total_vce_width=43.58,
        vce_widths=(
            1
            / 12
            * np.transpose(
                np.array(
                    [
                        [60, 1],
                        [60, 1],
                        [40, 1],
                        [54, 1],
                        [40, 1],
                        [54, 1],
                        [54, 1],
                        [54, 1],
                        [54, 1],
                        [54, 1],
                        [54, 1],
                    ]
                )
            )
        ),
    )
    params_p30 = dataclasses.replace(params_p3120, train2_arrival_time=0)
    params_p3recon0 = dataclasses.replace(params_p3recon120, train2_arrival_time=0)
    params_p6120 = dataclasses.replace(params_p60, train2_arrival_time=CLOSE_HEADWAY)
    params_p6300 = dataclasses.replace(params_p60, train2_arrival_time=NORMAL_HEADWAY)
    params_p100 = dataclasses.replace(params_p10120, train2_arrival_time=0)
    params_p10300 = dataclasses.replace(params_p10120, train2_arrival_time=NORMAL_HEADWAY)
    params_p110 = dataclasses.replace(params_p11120, train2_arrival_time=0)
    params_p11300 = dataclasses.replace(params_p11120, train2_arrival_time=NORMAL_HEADWAY)
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
