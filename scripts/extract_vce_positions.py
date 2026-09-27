#!/usr/bin/env -S uv run --script
# /// script
# requires-python = ">=3.14"
# dependencies = ["pymupdf"]
# ///

"""
Extract the approximate positions of each platform's VCEs
from the platform-level plans in the NY Penn Station Master Plan's Alternatives Report
(August 2020 draft), and match them to the widths in `data/master_plan_vce_widths.csv`.

The plans are vector drawings, with each VCE drawn as a rectangle filled with one of four colors:
new or existing, stair or escalator.
Each platform's width table lists its VCEs in the same east-to-west order as the plan,
so the `n`th rectangle on a platform, from east to west, is its `n`th VCE in the table.
The draft's plans and tables don't always agree, though.
Where they list the same sequence of stairs and escalators, they're matched,
keeping both their statuses (new or existing), which sometimes differ.
Other platforms are skipped and reported.
The plan is only used for positions; widths come from the tables,
since the plan is too small to measure widths from.

Writes `data/master_plan_vce_positions.csv`,
and `data/master_plan_existing_vces.csv`,
which combines the existing VCEs across all of the alternatives,
since each alternative keeps a different subset of them.
"""

import csv
import sys
import urllib.request
from dataclasses import dataclass
from pathlib import Path

import pymupdf

REPO = Path(__file__).resolve().parent.parent
WIDTHS_CSV = REPO / "data" / "master_plan_vce_widths.csv"
POSITIONS_CSV = REPO / "data" / "master_plan_vce_positions.csv"
EXISTING_CSV = REPO / "data" / "master_plan_existing_vces.csv"
PDF_CACHE = REPO / ".cache" / "PSMP-Alternatives-Report.pdf"
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

EAST_EDGE_X = 166
"""
x coordinate of the east end of the platforms in the plan,
where every platform's outline starts; positions are measured west from here.
"""

SAME_VCE_TOLERANCE_FT = 15
"""VCEs of the same type within this many feet in different alternatives are the same VCE."""

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
        PDF_CACHE.parent.mkdir(exist_ok=True)
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


def east_to_west(row: list[Rect]) -> list[Rect]:
    """Sort VCEs east to west, with escalators before stairs beside them."""
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
    for alternative, (page_number, source) in PLAN_PAGES.items():
        rows = platform_rows(vce_rects(doc[page_number - 1]))
        for platform, row in rows.items():
            table = sorted(
                (w for w in widths if w["source"] == source and int(w["platform"]) == platform),
                key=lambda w: int(w["vce"]),
            )
            drawn = east_to_west(row)
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
                        "vce": w["vce"],
                        "width_in": w["width_in"],
                        "type": w["type"],
                        "table_status": w["status"],
                        "plan_status": r.status,
                        "east_end_ft": round((r.x0 - EAST_EDGE_X) / PDF_UNITS_PER_FOOT),
                        "west_end_ft": round((r.x1 - EAST_EDGE_X) / PDF_UNITS_PER_FOOT),
                    }
                )
    with POSITIONS_CSV.open("w", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=list(out[0]), lineterminator="\n")
        writer.writeheader()
        writer.writerows(out)
    print(f"wrote {len(out)} VCEs; {mismatches} platforms didn't match their tables")
    write_existing(out)


def write_existing(vces: list[dict[str, int | str]]) -> None:
    """
    Combine the VCEs marked existing, by either the table or the plan,
    across the alternatives, by type and position.
    """
    existing = [v for v in vces if "existing" in (v["table_status"], v["plan_status"])]
    combined: list[dict[str, int | str]] = []
    for v in sorted(existing, key=lambda v: (-int(v["platform"]), int(v["east_end_ft"]))):
        mid = (int(v["east_end_ft"]) + int(v["west_end_ft"])) / 2
        for c in combined:
            if (
                c["platform"] == v["platform"]
                and c["type"] == v["type"]
                and abs(float(c["mid_ft"]) - mid) <= SAME_VCE_TOLERANCE_FT
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
                    "mid_ft": round(mid),
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
