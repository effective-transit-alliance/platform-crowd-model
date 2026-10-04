"""
Extract the shapes of platforms 1 to 11 and the diagonal platform and of what's on them,
as `shapes` describes, from NJT's PCIP Phase 1 existing plan
(Appendix A, sheet A-001, "Existing Plan Overall", July 2019, PDF page 24),
the same drawing as the plan of Alternative 12 that `platform_a_pcip_phase_1` measures,
without Platform A.
Platforms 1 and 2 share one outline, since the sheet draws them as one,
and the sheet cuts platforms 5 to 8 off at their west ends.
It's from before Moynihan Train Hall opened, so it doesn't have its escalators.

The sheet is a vector drawing, drawn rotated, with no text,
so its platforms are told apart by which row of the sheet they're in.
It's marked not to scale, so it's registered to the frame like Platform A's plan:
x by fitting the platforms' east ends to `data/platform_east_ends.csv`,
and y, separately, by fitting platforms 3 to 8's centerlines to the PCIP Phase 2 plan's,
as `shapes_pcip_phase_2.read_plan` reads them; the two scales agree to within 1%.
`data/vces.csv` reads the plan the same way, so its VCEs match the shapes.
"""

from collections.abc import Callable

import numpy as np
import pymupdf
from shapely import Point, Polygon, box, unary_union

from platform_crowd_model import shapes_pcip_phase_2
from platform_crowd_model.platform_a_pcip_phase_1 import (
    EXISTING_PLATFORM_FILL,
    calibration,
)
from platform_crowd_model.platform_west_ends_pcip_phase_1 import pdf
from platform_crowd_model.shapes import XY, Plan, item_points, read_drawing

PAGE = 24
"""1-indexed PDF page of sheet A-001, "Existing Plan Overall"."""
SOURCE = f"PCIP Phase 1 Appendix A, sheet A-001, July 2019, PDF page {PAGE}"

PLATFORM_ROWS = {
    "11": (184, 214),
    "10": (248, 305),
    "9": (317, 347),
    "8": (380, 411),
    "7": (444, 474),
    "6": (507, 534),
    "5": (567, 598),
    "4": (630, 661),
    "3": (694, 725),
    "1/2": (758, 854),
}
"""Each platform's rows (PDF units, rotated) on the sheet; platforms 1 and 2 share one outline."""

SHARED_ROWS = {"2": (758, 806), "1": (806, 854)}
"""Platforms 1 and 2's halves of their shared outline's rows, platform 1 being the south one."""

DIAGONAL_PLATFORM = "diagonal"
DIAGONAL_PLATFORM_MAX_X = 420
"""The diagonal platform is the platform-filled shape west of this (PDF units, rotated)."""

Y_FIT_PLATFORMS = [str(p) for p in range(3, 9)]
MAX_SCALE_DIFFERENCE = 0.01
MAX_Y_RESIDUAL_FT = 3

WALL_GRAYS = (0.3, 0.5)
"""The range of grays the sheet draws building elements in, e.g. walls."""

CONCOURSE_FILL = (0.965, 0.965, 0.879)
"""The sheet's fill color for existing concourses."""

LABEL_FILL = (1.0, 1.0, 1.0)
MIN_LABEL_OUTLINE_WIDTH = 0.5
"""The platforms' labels are white boxes with a thick outline."""


def read_plan(
    page_number: int,
    source: str,
    fills: dict[str, tuple[float, ...]],
    names: Callable[[str, str], list[tuple[float, str]]] | None = None,
) -> Plan:
    """
    The plan on `page_number`, a sheet drawn like the existing plan,
    with platforms besides the existing ones filled with `fills`, e.g. Platform A.
    y is fitted to the PCIP Phase 2 plan's platforms (`shapes_pcip_phase_2.read_plan`).
    """
    page = pdf()[page_number - 1]
    rotate = page.rotation_matrix
    x_ft, ft_per_unit = calibration(page)

    # Each platform's outline, in rotated PDF units, from the shapes filled as platform.
    pieces: dict[str, list[Polygon]] = {}
    labels = []
    for d in page.get_drawings():
        fill = d.get("fill")
        if fill is None:
            continue
        points = [p * rotate for item in d["items"] for p in item_points(item)]
        if len(points) < 3:
            continue
        polygon = Polygon([(p.x, p.y) for p in points]).buffer(0)
        if all(abs(c - f) < 0.002 for c, f in zip(fill, EXISTING_PLATFORM_FILL, strict=True)):
            x0, y0, x1, y1 = polygon.bounds
            if x1 - x0 < 50:
                continue
            middle = (y0 + y1) / 2
            if x1 < DIAGONAL_PLATFORM_MAX_X and y1 - y0 > 50:
                name = DIAGONAL_PLATFORM
            else:
                name = next(
                    (p for p, (top, bottom) in PLATFORM_ROWS.items() if top <= middle <= bottom),
                    "",
                )
            if name:
                pieces.setdefault(name, []).append(polygon)
        elif name := next(
            (
                p
                for p, color in fills.items()
                if all(abs(c - f) < 0.002 for c, f in zip(fill, color, strict=True))
            ),
            "",
        ):
            pieces.setdefault(name, []).append(polygon)
        elif (
            fill == LABEL_FILL
            and d.get("color") is not None
            and max(d["color"]) < 0.3
            and (d.get("width") or 0) >= MIN_LABEL_OUTLINE_WIDTH
        ):
            labels.append(polygon)

    # y: fit each platform's centerline to the PCIP Phase 2 plan's.
    _plan, pcip_2_outlines = shapes_pcip_phase_2.read_plan()
    pcip_2 = dict(pcip_2_outlines)
    rotated_ys = []
    ft_ys = []
    for platform in Y_FIT_PLATFORMS:
        _, y0, _, y1 = unary_union(pieces[platform]).bounds
        rotated_ys.append((y0 + y1) / 2)
        _, fy0, _, fy1 = pcip_2[platform].bounds
        ft_ys.append((fy0 + fy1) / 2)
    y_per_unit, y_offset = np.polyfit(rotated_ys, ft_ys, 1)
    residuals = [f - (y_per_unit * y + y_offset) for y, f in zip(rotated_ys, ft_ys, strict=True)]
    print(
        f"{-y_per_unit:.4f} ft per PDF unit across the platforms, "
        f"centerlines off by up to {max(map(abs, residuals)):.1f} ft"
    )
    if abs(-y_per_unit / ft_per_unit - 1) > MAX_SCALE_DIFFERENCE:
        raise RuntimeError(f"the sheet's scales differ: {ft_per_unit:.4f} and {-y_per_unit:.4f}")
    if max(map(abs, residuals)) > MAX_Y_RESIDUAL_FT:
        raise RuntimeError("the platforms' centerlines don't fit the PCIP Phase 2 plan's")

    def rotated_to_ft(x: float, y: float) -> XY:
        return x_ft(x), float(y_per_unit * y + y_offset)

    def to_ft(point: pymupdf.Point) -> XY:
        rotated = point * rotate
        return rotated_to_ft(rotated.x, rotated.y)

    def in_ft(polygon: Polygon) -> Polygon:
        return Polygon([rotated_to_ft(x, y) for x, y in polygon.exterior.coords])

    # The diagonal platform is drawn in pieces, so it's their convex hull.
    outlines = [
        (
            name,
            in_ft(
                unary_union(polygons).convex_hull
                if name == DIAGONAL_PLATFORM
                else unary_union(polygons).buffer(0.5).buffer(-0.5)
            ),
        )
        for name, polygons in pieces.items()
    ]
    rows = {**PLATFORM_ROWS, **SHARED_ROWS}
    del rows["1/2"]
    rows_ft = {name: sorted(rotated_to_ft(0, y)[1] for y in ys) for name, ys in rows.items()}
    # Platforms not in a row are told apart by their outlines.
    shaped = [(name, o) for name, o in outlines if name == DIAGONAL_PLATFORM or name in fills]

    def platform_at(point: XY) -> str | None:
        on = next((name for name, o in shaped if o.contains(Point(point))), None)
        return on or next((p for p, (lo, hi) in rows_ft.items() if lo <= point[1] <= hi), None)

    return Plan(
        source=source,
        drawing=read_drawing(page, to_ft, ft_per_unit),
        outlines=outlines,
        platform_at=platform_at,
        hidden=unary_union([in_ft(box(*label.bounds)) for label in labels]),
        wall_grays=WALL_GRAYS,
        concourse_fill=CONCOURSE_FILL,
        names=names,
    )
