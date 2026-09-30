"""
Extract the shapes of the platforms and of what's on them, as `shapes` describes,
from NJT's PCIP Phase 1 plan of Alternative 12, with Platform A,
the sheet `platform_a_pcip_phase_1` measures,
drawn like the existing plan `shapes_pcip_phase_1` reads, and registered to the frame the same way.

Platform A's VCEs are drawn as new stairs and escalators, and found like the existing ones.
Each of AP7 to AP12's stair and escalator side by side is drawn as one run of treads,
so it's one footprint, of type `stair`.
They're named by their labels on the sheet, e.g. `AP7`,
from `data/vces_platform_a_pcip_phase_1.csv`.

Writes `data/shapes_platform_a_pcip_phase_1.geojson`
and `data/shapes_platform_a_pcip_phase_1_lonlat.geojson`.
"""

import csv

from platform_crowd_model.paths import DATA_DIR
from platform_crowd_model.platform_a_pcip_phase_1 import PAGE, PLATFORM_A_FILL, SOURCE, VCES_CSV
from platform_crowd_model.shapes import vce_names, write
from platform_crowd_model.shapes_pcip_phase_1 import read_plan

OUT_GEOJSON = DATA_DIR / "shapes_platform_a_pcip_phase_1.geojson"
LONLAT_GEOJSON = DATA_DIR / "shapes_platform_a_pcip_phase_1_lonlat.geojson"

PLATFORM_A = "A"


def names(platform: str, vce_type: str) -> list[tuple[float, str]]:
    """
    Platform A's VCEs' labels, of any type, since a stair and escalator side by side is one,
    or the other platforms' VCEs, from `data/vces.csv`.
    """
    if platform != PLATFORM_A:
        return vce_names(platform, vce_type)
    with VCES_CSV.open() as f:
        return sorted(
            {
                ((int(row["west_end_ft"]) + int(row["east_end_ft"])) / 2, row["label"])
                for row in csv.DictReader(f)
            }
        )


def main() -> None:
    plan = read_plan(PAGE, SOURCE, {PLATFORM_A: PLATFORM_A_FILL}, names)
    write(plan, OUT_GEOJSON, LONLAT_GEOJSON)
