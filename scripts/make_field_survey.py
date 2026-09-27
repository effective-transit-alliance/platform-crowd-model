#!/usr/bin/env -S uv run --script
# /// script
# requires-python = ">=3.14"
# dependencies = []
# ///

"""
Make a blank field survey sheet for measuring every platform's VCEs in person,
prefilled with the VCEs expected from NJ Transit's January 2022 station directory
(`data/njt_directory_vces.csv`) and the Master Plan (`data/master_plan_existing_vces.csv`).

The two sources don't reconcile and can't be aligned reliably,
so each platform lists both sources' VCEs, each sorted east to west,
and surveyors mark which ones they find and add rows for any that neither lists.
Platforms are ordered by priority:
platform 3 first, as the ETA report's focus, then platform 11, which has no width data,
then the rest.

Writes `data/field_survey.csv`; the columns after `master_plan_mid_ft` are for surveyors.
"""

import csv
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
DIRECTORY_CSV = REPO / "data" / "njt_directory_vces.csv"
MASTER_PLAN_CSV = REPO / "data" / "master_plan_existing_vces.csv"
OUT_CSV = REPO / "data" / "field_survey.csv"

PLATFORM_ORDER = [3, 11, 1, 2, 4, 5, 6, 7, 8, 9, 10]

TRACKS = {
    1: "1/2",
    2: "3/4",
    3: "5/6",
    4: "7/8",
    5: "9/10",
    6: "11/12",
    7: "13/14",
    8: "15/16",
    9: "17",
    10: "18/19",
    11: "20/21",
}

SURVEY_COLUMNS = [
    "found",
    "type",
    "clear_width_in",
    "escalator_step_width_in",
    "escalator_direction_am_peak",
    "escalator_direction_pm_peak",
    "nearest_column_number",
    "distance_from_east_end_ft",
    "leads_to",
    "obstructions",
    "photos",
    "notes",
]
"""
- `found`: yes, no, or a duplicate of another row's `id`.
- `clear_width_in`: between the handrails, at the platform end.
- `escalator_direction_*`: up, down, or stopped.
- `leads_to`: the concourse, e.g. NJ Transit, Amtrak, LIRR, Exit, West End, or Moynihan.
- `obstructions`: columns, benches, bins, or narrow landings near the bottom.
"""


def main() -> None:
    with DIRECTORY_CSV.open() as f:
        directory = list(csv.DictReader(f))
    with MASTER_PLAN_CSV.open() as f:
        master_plan = list(csv.DictReader(f))
    rows = []
    for platform in PLATFORM_ORDER:
        # The directory's map has east on the right, so east to west is descending x.
        on_map = sorted(
            (v for v in directory if int(v["platform"]) == platform),
            key=lambda v: (v["level"], -int(v["map_x"])),
        )
        # The Master Plan's positions are feet west of the platforms' east end.
        in_plan = sorted(
            (v for v in master_plan if int(v["platform"]) == platform),
            key=lambda v: int(v["mid_ft"]),
        )
        for n, v in enumerate(on_map, 1):
            rows.append(
                {
                    "id": f"P{platform}-D{n}",
                    "platform": platform,
                    "tracks": TRACKS[platform],
                    "source": "2022 directory",
                    "expected_type": v["type"],
                    "expected_width_in": "",
                    "directory_level": v["level"],
                    "directory_map_x": v["map_x"],
                    "master_plan_mid_ft": "",
                }
            )
        for n, v in enumerate(in_plan, 1):
            rows.append(
                {
                    "id": f"P{platform}-M{n}",
                    "platform": platform,
                    "tracks": TRACKS[platform],
                    "source": "Master Plan",
                    "expected_type": v["type"],
                    # Alternatives sometimes disagree, e.g. `44/48`.
                    "expected_width_in": v["width_in"],
                    "directory_level": "",
                    "directory_map_x": "",
                    "master_plan_mid_ft": v["mid_ft"],
                }
            )
    with OUT_CSV.open("w", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=[*rows[0], *SURVEY_COLUMNS], lineterminator="\n")
        writer.writeheader()
        writer.writerows(rows)
    print(f"wrote {len(rows)} rows")


if __name__ == "__main__":
    main()
