"""
Make a blank field survey sheet for measuring every platform's VCEs in person,
prefilled with the VCEs expected from `data/estimated_vce_widths.csv`,
from the PCIP Phase 2 existing plan on platforms 1 to 8
and positioned from the directory on platforms 9 to 11,
NJ Transit's January 2022 station directory (`data/njt_directory_vces.csv`),
and the Master Plan (`data/master_plan_existing_vces.csv`).

The sources don't reconcile, and the directory can't be aligned with the others,
so each platform lists every source's VCEs, each sorted west to east,
and surveyors mark which ones they find and add rows for any that neither lists.
Platforms are ordered by priority:
platform 3 first, as the ETA report's focus, then platform 11, which has no width data,
then the rest.

`position_ft` is in the Master Plan's frame: feet east of its plans' west edge,
which cuts across the platforms under the West End Concourse.

Writes `data/field_survey.csv`; the columns after `position_ft` are for surveyors.
"""

import csv

from platform_crowd_model.paths import DATA_DIR

DIRECTORY_CSV = DATA_DIR / "njt_directory_vces.csv"
MASTER_PLAN_CSV = DATA_DIR / "master_plan_existing_vces.csv"
SHEET_CSV = DATA_DIR / "estimated_vce_widths.csv"
OUT_CSV = DATA_DIR / "field_survey.csv"

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
    with SHEET_CSV.open() as f:
        sheet = list(csv.DictReader(f))
    rows = []
    for platform in PLATFORM_ORDER:
        for v in (v for v in sheet if int(v["platform"]) == platform):
            rows.append(
                {
                    "id": v["vce"],
                    "platform": platform,
                    "tracks": TRACKS[platform],
                    "source": "PCIP Phase 2 existing plan"
                    if v["source"].startswith("PCIP")
                    else "2022 directory, positioned",
                    "expected_type": v["type"],
                    "expected_width_in": v["master_plan_width_in"] or v["sheet_width_in"],
                    "expected_width_status": v["width_status"],
                    "directory_level": "",
                    "directory_map_x": "",
                    "position_ft": round((int(v["west_end_ft"]) + int(v["east_end_ft"])) / 2),
                }
            )
        # The directory's map has west on the left, so west to east is ascending x.
        on_map = sorted(
            (v for v in directory if int(v["platform"]) == platform),
            key=lambda v: (v["level"], int(v["map_x"])),
        )
        # The Master Plan's positions are feet east of its plans' west edge.
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
                    "expected_width_status": "",
                    "directory_level": v["level"],
                    "directory_map_x": v["map_x"],
                    "position_ft": "",
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
                    "expected_width_status": "master plan",
                    "directory_level": "",
                    "directory_map_x": "",
                    "position_ft": v["mid_ft"],
                }
            )
    with OUT_CSV.open("w", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=[*rows[0], *SURVEY_COLUMNS], lineterminator="\n")
        writer.writeheader()
        writer.writerows(rows)
    print(f"wrote {len(rows)} rows")


if __name__ == "__main__":
    main()
