"""Type stubs for the Rust extension built from `crates/core`."""

from typing import Protocol

from platform_crowd_model.model import CoreInput

class CoreResult(Protocol):
    finished: bool
    max_up_rate: float
    time_at_capacity: int
    taper_time: int | None
    clear_time: int | None
    arrival_times: list[int | None]
    dwells: list[int | None]
    boarded_time: int | None
    max_pax_on_platform: float
    max_occupants: float
    min_space_per_pax: float
    vce_gone_up: list[float]
    vce_empty_times: list[int | None]
    gone_up: float
    still_aboard: float
    arriving_pax_on_platform: float
    boarded: float
    still_upstairs: float
    still_walking_to_cars: float
    still_waiting_at_cars: float
    times: list[int]
    arriving_pax_waiting_on_platform: list[float]
    off_rate: list[float]
    on_rate: list[float]
    down_rate: list[float]
    departing_pax_on_platform: list[float]
    total_pax_on_platform: list[float]
    platform_crowding: list[float]
    up_rate: list[float]
    net_pax_flow_rate: list[float]
    vces: list[list[list[float]]]
    train_values: list[list[list[float]]]

def simulate_core(input: CoreInput, record_time_series: bool) -> CoreResult: ...
def py_sum_for_test(values: list[float]) -> float: ...
