"""
Extract the approximate positions of each platform's VCEs
from the platform-level plans in the NY Penn Station Master Plan's Alternatives Report
(August 2020 draft), and match them to the widths in `data/vce_widths_master_plan.csv`.

The plans are vector drawings, with each VCE drawn as a rectangle filled with one of four colors:
new or existing, stair or escalator.
Each platform's width table lists its VCEs in the same west-to-east order as the plan,
so the `n`th rectangle on a platform, from west to east, is its `n`th VCE in the table.
The plans have west on the left:
the platforms' east ends on the plans match those on an existing-conditions plan in NJ Transit's
PCIP Phase 2 drawings (November 2020) to within about 4 ft,
and the gray wedges on the right are the 7th Ave subway.
The draft's plans and tables don't always agree, though.
Where they list the same sequence of stairs and escalators, they're matched,
keeping both their statuses (new or existing), which sometimes differ.
Other platforms are skipped and reported.
The plan is only used for positions; widths come from the tables,
since the plan is too small to measure widths from.

Writes `data/vce_positions_master_plan.csv`,
and `data/vces_existing_master_plan.csv`,
which combines the existing VCEs across all of the alternatives,
since each alternative keeps a different subset of them,
and `data/platform_east_ends_master_plan.csv`,
each platform's east end from its outline, averaged across the alternatives.
"""

import csv
import sys
import urllib.request
from dataclasses import dataclass

import pymupdf

from platform_crowd_model.paths import CACHE_DIR, DATA_DIR

WIDTHS_CSV = DATA_DIR / "vce_widths_master_plan.csv"
POSITIONS_CSV = DATA_DIR / "vce_positions_master_plan.csv"
EXISTING_CSV = DATA_DIR / "vces_existing_master_plan.csv"
EAST_ENDS_CSV = DATA_DIR / "platform_east_ends_master_plan.csv"
PDF_CACHE = CACHE_DIR / "PSMP-Alternatives-Report.pdf"
PDF_URL = (
    "https://liamblank.com/wp-content/uploads/2026/07/"
    "R-Master-Plan-MTA-NJT-AMT-20_0814-PSMP-Alternatives-Report.pdf"
)

PLAN_PAGES = {
    # Alternative: (1-indexed PDF page of its platform-level plan, `source` of its width table).
    1: (23, "draft-2020-08-p16"),
    2: (35, "draft-2020-08-p28"),
    3: (47, "draft-2020-08-p40"),
    4: (59, "draft-2020-08-p52"),
}

FILLS = {
    ("stair", "new"): (0.271, 0.565, 0.255),
    ("stair", "existing"): (0.671, 0.831, 0.549),
    ("escalator", "existing"): (0.675, 0.804, 0.925),
    ("escalator", "new"): (0.0, 0.475, 0.71),
}

PDF_UNITS_PER_FOOT = (259.57 - 175.01) / 100
"""From the plan's scale bar, whose 0' and 100' ticks are at these x coordinates."""

WEST_EDGE_X = 166
"""
x coordinate of the plan's west edge, under the West End Concourse,
where every platform's outline starts; positions are measured east from here.
"""

WEST_EDGE_TOLERANCE = 2
"""
Outlines starting within this many PDF units of `WEST_EDGE_X` are platforms',
since alternatives 3 and 4's plans are drawn about 1 unit farther west.
"""

SAME_VCE_TOLERANCE_FT = 15
"""VCEs of the same type within this many feet in different alternatives are the same VCE."""

PLATFORM_FILL = (0.984, 0.965, 0.867)
"""The plans' fill color for platforms."""

LEGEND_MIN_X = 1050
"""The legend's color swatches are to the right of this x coordinate."""

ROW_GAP = 8
"""Minimum vertical gap (PDF units) between VCEs on adjacent platforms."""

SAME_X_TOLERANCE = 3
"""
VCEs whose centers are within this many PDF units along the platform are side by side,
e.g. an escalator next to a stair, which the tables list escalator first.
"""


@dataclass
class Rect:
    type: str
    status: str
    x0: float
    y0: float
    x1: float
    y1: float

    @property
    def x(self) -> float:
        return (self.x0 + self.x1) / 2

    @property
    def y(self) -> float:
        return (self.y0 + self.y1) / 2


def pdf() -> pymupdf.Document:
    if not PDF_CACHE.exists():
        CACHE_DIR.mkdir(exist_ok=True)
        request = urllib.request.Request(
            PDF_URL, headers={"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) Firefox/130.0"}
        )
        with urllib.request.urlopen(request) as response:
            PDF_CACHE.write_bytes(response.read())
    return pymupdf.open(PDF_CACHE)


def vce_rects(page: pymupdf.Page) -> list[Rect]:
    rects = []
    for drawing in page.get_drawings():
        fill = drawing.get("fill")
        if fill is None:
            continue
        for (type_, status), color in FILLS.items():
            if all(abs(a - b) < 0.01 for a, b in zip(fill, color, strict=True)):
                r = drawing["rect"]
                if r.x0 < LEGEND_MIN_X:
                    rects.append(Rect(type_, status, r.x0, r.y0, r.x1, r.y1))
    return rects


def platform_rows(rects: list[Rect]) -> dict[int, list[Rect]]:
    """Group the VCEs into platforms, numbered 11 at the top (north) to 1 at the bottom (south)."""
    rects = sorted(rects, key=lambda r: r.y)
    rows: list[list[Rect]] = [[rects[0]]]
    for r in rects[1:]:
        if r.y - rows[-1][-1].y > ROW_GAP:
            rows.append([])
        rows[-1].append(r)
    if len(rows) != 11:
        sys.exit(f"found {len(rows)} rows of VCEs, not 11")
    return {11 - i: row for i, row in enumerate(rows)}


def platform_east_ends(page: pymupdf.Page, rows: dict[int, list[Rect]]) -> dict[int, float]:
    """
    The x coordinate of each platform's east end, from the outline around its row of VCEs.
    Only outlines starting at `WEST_EDGE_X` are platforms';
    platforms 1 and 2 are drawn differently, so they're skipped.
    """
    out = {}
    for drawing in page.get_drawings():
        fill = drawing.get("fill")
        r = drawing["rect"]
        if (
            fill is None
            or any(abs(a - b) > 0.01 for a, b in zip(fill, PLATFORM_FILL, strict=True))
            or abs(r.x0 - WEST_EDGE_X) > WEST_EDGE_TOLERANCE
        ):
            continue
        for platform, row in rows.items():
            if r.y0 <= sum(v.y for v in row) / len(row) <= r.y1:
                out[platform] = r.x1
    return out


def west_to_east(row: list[Rect]) -> list[Rect]:
    """Sort VCEs west to east, with escalators before stairs beside them."""
    row = sorted(row, key=lambda r: r.x)
    groups: list[list[Rect]] = []
    for r in row:
        if groups and r.x - groups[-1][0].x <= SAME_X_TOLERANCE:
            groups[-1].append(r)
        else:
            groups.append([r])
    return [r for g in groups for r in sorted(g, key=lambda r: r.type != "escalator")]


def main() -> None:
    with WIDTHS_CSV.open() as f:
        widths = list(csv.DictReader(f))
    doc = pdf()
    out = []
    mismatches = 0
    east_ends: dict[int, list[float]] = {}
    for alternative, (page_number, source) in PLAN_PAGES.items():
        page = doc[page_number - 1]
        rows = platform_rows(vce_rects(page))
        for platform, x in platform_east_ends(page, rows).items():
            east_ends.setdefault(platform, []).append(x)
        for platform, row in rows.items():
            table = sorted(
                (w for w in widths if w["source"] == source and int(w["platform"]) == platform),
                key=lambda w: int(w["vce_number"]),
            )
            drawn = west_to_east(row)
            drawn_types = [r.type for r in drawn]
            table_types = [w["type"] for w in table]
            if drawn_types != table_types:
                mismatches += 1
                print(
                    f"alternative {alternative} platform {platform}: "
                    "plan and table have different stairs and escalators\n"
                    f"  plan:  {drawn_types}\n  table: {table_types}",
                    file=sys.stderr,
                )
                continue
            for w, r in zip(table, drawn, strict=True):
                out.append(
                    {
                        "alternative": alternative,
                        "platform": platform,
                        "vce_number": w["vce_number"],
                        "width_in": w["width_in"],
                        "type": w["type"],
                        "table_status": w["status"],
                        "plan_status": r.status,
                        "west_end_ft": round((r.x0 - WEST_EDGE_X) / PDF_UNITS_PER_FOOT),
                        "east_end_ft": round((r.x1 - WEST_EDGE_X) / PDF_UNITS_PER_FOOT),
                    }
                )
    with POSITIONS_CSV.open("w", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=list(out[0]), lineterminator="\n")
        writer.writeheader()
        writer.writerows(out)
    print(f"wrote {len(out)} VCEs; {mismatches} platforms didn't match their tables")
    write_existing(out)
    write_east_ends(east_ends)


def write_east_ends(east_ends: dict[int, list[float]]) -> None:
    """Average each platform's east end across the alternatives."""
    with EAST_ENDS_CSV.open("w", newline="") as f:
        writer = csv.writer(f, lineterminator="\n")
        writer.writerow(["platform", "east_end_ft", "alternatives"])
        for platform, xs in sorted(east_ends.items()):
            x = sum(xs) / len(xs)
            writer.writerow([platform, round((x - WEST_EDGE_X) / PDF_UNITS_PER_FOOT), len(xs)])
    print(f"wrote {len(east_ends)} platforms' east ends")


def write_existing(vces: list[dict[str, int | str]]) -> None:
    """
    Combine the VCEs marked existing, by either the table or the plan,
    across the alternatives, by type and position.
    """
    existing = [v for v in vces if "existing" in (v["table_status"], v["plan_status"])]
    combined: list[dict[str, int | str]] = []
    for v in sorted(existing, key=lambda v: (-int(v["platform"]), int(v["west_end_ft"]))):
        mid = (int(v["west_end_ft"]) + int(v["east_end_ft"])) / 2
        for c in combined:
            if (
                c["platform"] == v["platform"]
                and c["type"] == v["type"]
                and abs(float(c["midpoint_ft"]) - mid) <= SAME_VCE_TOLERANCE_FT
            ):
                widths = {int(w) for w in str(c["width_in"]).split("/")} | {int(v["width_in"])}
                c["width_in"] = "/".join(map(str, sorted(widths)))
                c["alternatives"] = f"{c['alternatives']} {v['alternative']}"
                break
        else:
            combined.append(
                {
                    "platform": v["platform"],
                    "type": v["type"],
                    "width_in": v["width_in"],
                    "midpoint_ft": round(mid),
                    "alternatives": str(v["alternative"]),
                }
            )
    with EXISTING_CSV.open("w", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=list(combined[0]), lineterminator="\n")
        writer.writeheader()
        writer.writerows(combined)
    print(f"wrote {len(combined)} existing VCEs")


if __name__ == "__main__":
    main()
