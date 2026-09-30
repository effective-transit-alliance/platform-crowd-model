"""
Measure Platform A, a new platform south of Platform 1 under West 31st Street,
and its VCEs, on NJT's PCIP Phase 1 plan of Alternative 12 (Appendix A, sheet A-021, July 2019).

Platform A runs from under the 7th Avenue Subway west under the 8th Avenue Subway
to an extension of the West End Concourse, long enough for a 12-car train.
Per the PCIP Phase 1 final report, its VCEs are
pairs of a stair and an escalator up to a new Concourse A above it (AP7 to AP12),
and, where its west end curves too much for escalators,
stairs up to the West End Concourse's extension (WP5 and WP6):
8 stairs and 6 escalators, with elevators between them.
The report doesn't give their widths,
so stairs are taken to be the 5'-0" egress stairs it sizes the other alternatives' with,
and escalators to have 40 in. steps.

The plan is a vector drawing, drawn rotated, so its shapes' coordinates are rotated first.
It's converted to the Master Plan's frame (feet east of its plans' west edge)
by fitting a line to where the existing Platforms 3 to 11 end to the east on it
and in `data/platform_east_ends.csv`.

Writes `data/platform_a_pcip_phase_1.csv` and `data/vces_platform_a_pcip_phase_1.csv`.
"""

import csv
from collections.abc import Callable

import numpy as np
import pymupdf

from platform_crowd_model.paths import DATA_DIR
from platform_crowd_model.platform_west_ends_pcip_phase_1 import pdf

PAGE = 31
"""1-indexed PDF page of sheet A-021, "Alt 12 Proposed Plan Overall"."""
SOURCE = f"PCIP Phase 1 Appendix A, sheet A-021, July 2019, PDF page {PAGE}"

EAST_ENDS_CSV = DATA_DIR / "platform_east_ends.csv"
PLATFORM_CSV = DATA_DIR / "platform_a_pcip_phase_1.csv"
VCES_CSV = DATA_DIR / "vces_platform_a_pcip_phase_1.csv"

EXISTING_PLATFORM_FILL = (0.729, 0.729, 0.729)
PLATFORM_A_FILL = (0.4, 0.514, 0.533)

PLATFORM_ROWS = {
    11: (184, 214),
    10: (248, 305),
    9: (317, 347),
    8: (380, 411),
    7: (444, 474),
    6: (508, 533),
    5: (567, 598),
    4: (631, 661),
    3: (694, 724),
}
"""Each existing platform's rows (PDF units, rotated) on the plan."""

VCES = {
    "WP5": (612, 647, ("stair",)),
    "WP6": (687, 710, ("stair",)),
    "AP7": (885, 930, ("stair", "escalator")),
    "AP8": (1019, 1076, ("stair", "escalator")),
    "AP9": (1263, 1320, ("stair", "escalator")),
    "AP10": (1542, 1599, ("stair", "escalator")),
    "AP11": (1734, 1791, ("stair", "escalator")),
    "AP12": (1860, 1905, ("stair", "escalator")),
}
"""Each labeled VCE location's columns (PDF units, rotated) on the plan, and what's there."""

STAIR_WIDTH_IN = 60
ESCALATOR_WIDTH_IN = 40
TRACKS = 2
"""Platform A is an island platform, sized for 2 fully loaded 12-car trains."""
MAX_CARS = 12


def filled(page: pymupdf.Page, fill: tuple[float, float, float]) -> list[pymupdf.Rect]:
    """The rotated bounding boxes of the shapes filled with `fill`."""
    return [
        drawing["rect"] * page.rotation_matrix
        for drawing in page.get_drawings()
        if drawing.get("fill")
        and all(abs(c - f) < 0.002 for c, f in zip(drawing["fill"], fill, strict=True))
    ]


def outline_area(page: pymupdf.Page, fill: tuple[float, float, float]) -> float:
    """The area (PDF units squared) inside the largest shape filled with `fill`."""
    best = 0.0
    for drawing in page.get_drawings():
        if not (
            drawing.get("fill")
            and all(abs(c - f) < 0.002 for c, f in zip(drawing["fill"], fill, strict=True))
        ):
            continue
        points = [item[1] for item in drawing["items"] if item[0] == "l"]
        points += [drawing["items"][-1][2]] if drawing["items"][-1][0] == "l" else []
        area = (
            abs(
                sum(
                    a.x * b.y - b.x * a.y
                    for a, b in zip(points, points[1:] + points[:1], strict=True)
                )
            )
            / 2
        )
        best = max(best, area)
    return best


def calibration(page: pymupdf.Page) -> tuple[Callable[[float], float], float]:
    """A function from the plan's rotated columns to feet, and its feet per PDF unit."""
    rects = filled(page, EXISTING_PLATFORM_FILL)
    with EAST_ENDS_CSV.open() as f:
        east_ends = {int(row["platform"]): float(row["east_end_ft"]) for row in csv.DictReader(f)}
    xs: list[float] = []
    fts: list[float] = []
    for platform, (top, bottom) in PLATFORM_ROWS.items():
        ends = [r.x1 for r in rects if top <= (r.y0 + r.y1) / 2 <= bottom]
        xs.append(max(ends))
        fts.append(east_ends[platform])
    ft_per_pt, offset = np.polyfit(xs, fts, 1)
    residuals = [f - (ft_per_pt * x + offset) for x, f in zip(xs, fts, strict=True)]
    print(
        f"{ft_per_pt:.4f} ft per PDF unit, east ends off by up to {max(map(abs, residuals)):.1f} ft"
    )
    return (lambda x: float(ft_per_pt * x + offset)), float(ft_per_pt)


def main() -> None:
    page = pdf()[PAGE - 1]
    ft, ft_per_pt = calibration(page)
    (platform_a,) = filled(page, PLATFORM_A_FILL)
    area = outline_area(page, PLATFORM_A_FILL) * ft_per_pt**2
    west, east = ft(platform_a.x0), ft(platform_a.x1)
    print(f"Platform A: {west:.0f} to {east:.0f} ft, {east - west:.0f} ft long, {area:.0f} sq ft")
    with PLATFORM_CSV.open("w", newline="") as f:
        writer = csv.writer(f, lineterminator="\n")
        writer.writerow(
            ["platform", "west_end_ft", "east_end_ft", "area_sq_ft", "tracks", "max_cars", "source"]
        )
        writer.writerow(["A", round(west), round(east), round(area), TRACKS, MAX_CARS, SOURCE])
    rows = [
        {
            "platform": "A",
            "label": label,
            "type": vce_type,
            "west_end_ft": round(ft(x0)),
            "east_end_ft": round(ft(x1)),
            "width_in": STAIR_WIDTH_IN if vce_type == "stair" else ESCALATOR_WIDTH_IN,
            "source": SOURCE,
        }
        for label, (x0, x1, types) in VCES.items()
        for vce_type in types
    ]
    types = [row["type"] for row in rows]
    print(f"{types.count('stair')} stairs and {types.count('escalator')} escalators")
    with VCES_CSV.open("w", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=list(rows[0]), lineterminator="\n")
        writer.writeheader()
        writer.writerows(rows)
