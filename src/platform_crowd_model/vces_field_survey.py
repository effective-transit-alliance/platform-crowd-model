"""
Make a blank field survey sheet for measuring every platform's VCEs in person,
prefilled with the VCEs expected from `data/vces.csv`,
from the PCIP Phase 2 existing plan on platforms 1 to 8,
positioned from the directory on platforms 9 to 11,
and from the Moynihan Station EA's plan at the platforms' west ends,
NJT's January 2022 station directory (`data/vces_njt_directory.csv`),
and the Master Plan (`data/vces_existing_master_plan.csv`).

The sources don't reconcile, and the directory can't be aligned with the others,
so each platform lists every source's VCEs, each sorted west to east,
and surveyors mark which ones they find and add rows for any that neither lists.
Platforms are ordered by priority:
platform 3 first, as the ETA report's focus, then platform 11, which has no width data,
then the rest.

`midpoint_ft` is in the Master Plan's frame: feet east of its plans' west edge,
which cuts across the platforms under the West End Concourse.

Writes `data/vces_field_survey.csv`; the columns after `midpoint_ft` are for surveyors.
Regenerating it keeps what surveyors have entered:
each row that's still generated keeps its survey columns,
matched by what identifies it rather than its `vce_name`, which can change as VCEs are added,
and every other row with anything entered, like one a surveyor added, is kept at the end.
"""

import csv
from typing import Any

from platform_crowd_model.paths import DATA_DIR
from platform_crowd_model.vces import (
    DIRECTORY_VCE_SOURCE,
    MOYNIHAN_EA_VCE_SOURCE,
    PCIP_PHASE_2_VCE_SOURCE,
)

DIRECTORY_CSV = DATA_DIR / "vces_njt_directory.csv"
MASTER_PLAN_CSV = DATA_DIR / "vces_existing_master_plan.csv"
SHEET_CSV = DATA_DIR / "vces.csv"
OUT_CSV = DATA_DIR / "vces_field_survey.csv"

SOURCE_LABELS = {
    PCIP_PHASE_2_VCE_SOURCE: "PCIP Phase 2 existing plan",
    DIRECTORY_VCE_SOURCE: "2022 directory, positioned",
    MOYNIHAN_EA_VCE_SOURCE: "Moynihan Station EA plan",
}
"""How the sheet labels each of `data/vces.csv`'s sources."""

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
- `found`: yes, no, or a duplicate of another row's `vce_name`.
- `clear_width_in`: between the handrails, at the platform end.
- `escalator_direction_*`: up, down, or stopped.
- `leads_to`: the concourse, e.g. NJT, Amtrak, LIRR, Exit, West End, or Moynihan.
- `obstructions`: columns, benches, bins, or narrow landings near the bottom.
"""


IDENTITY_COLUMNS = [
    "platform",
    "source",
    "expected_type",
    "directory_level",
    "directory_map_x",
    "midpoint_ft",
]
"""The columns that identify a generated row, since `vce_name` can change."""


def identity(row: dict[str, Any]) -> tuple[str, ...]:
    return tuple(str(row[column]) for column in IDENTITY_COLUMNS)


def surveyed(row: dict[str, Any]) -> bool:
    """Whether a surveyor has entered anything in `row`."""
    return any(row.get(column) for column in SURVEY_COLUMNS)


def keep_survey(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """
    `rows` with the survey columns already entered in `OUT_CSV` for the same VCEs,
    followed by every other row entered in `OUT_CSV` that isn't generated anymore.
    """
    if not OUT_CSV.exists():
        return rows
    with OUT_CSV.open() as f:
        old = [row for row in csv.DictReader(f) if surveyed(row)]
    by_identity = {identity(row): row for row in old}
    kept = set()
    for row in rows:
        entered = by_identity.get(identity(row))
        if entered:
            row.update({column: entered[column] for column in SURVEY_COLUMNS})
            kept.add(identity(row))
    unmatched = [row for row in old if identity(row) not in kept]
    for row in unmatched:
        print(f"keeping {row['vce_name'] or 'an added row'} on platform {row['platform']},")
        print("  which isn't generated anymore but has survey entries")
    return rows + unmatched


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
                    "vce_name": v["vce_name"],
                    "platform": platform,
                    "track_numbers": TRACKS[platform],
                    "source": SOURCE_LABELS[v["source"]],
                    "expected_type": v["type"],
                    "expected_width_in": v["master_plan_width_in"] or v["estimated_width_in"],
                    "expected_width_source": v["width_source"],
                    "directory_level": "",
                    "directory_map_x": "",
                    "midpoint_ft": round((int(v["west_end_ft"]) + int(v["east_end_ft"])) / 2),
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
            key=lambda v: int(v["midpoint_ft"]),
        )
        for n, v in enumerate(on_map, 1):
            rows.append(
                {
                    "vce_name": f"P{platform}-D{n}",
                    "platform": platform,
                    "track_numbers": TRACKS[platform],
                    "source": "2022 directory",
                    "expected_type": v["type"],
                    "expected_width_in": "",
                    "expected_width_source": "",
                    "directory_level": v["level"],
                    "directory_map_x": v["map_x"],
                    "midpoint_ft": "",
                }
            )
        for n, v in enumerate(in_plan, 1):
            rows.append(
                {
                    "vce_name": f"P{platform}-M{n}",
                    "platform": platform,
                    "track_numbers": TRACKS[platform],
                    "source": "Master Plan",
                    "expected_type": v["type"],
                    # Alternatives sometimes disagree, e.g. `44/48`.
                    "expected_width_in": v["width_in"],
                    "expected_width_source": "master_plan",
                    "directory_level": "",
                    "directory_map_x": "",
                    "midpoint_ft": v["midpoint_ft"],
                }
            )
    rows = keep_survey(rows)
    with OUT_CSV.open("w", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=[*rows[0], *SURVEY_COLUMNS], lineterminator="\n")
        writer.writeheader()
        writer.writerows(rows)
    print(f"wrote {len(rows)} rows")


if __name__ == "__main__":
    main()
