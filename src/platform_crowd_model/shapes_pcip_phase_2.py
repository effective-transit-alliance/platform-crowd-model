"""
Extract the shapes of platforms 1 to 8 and of what's on them, as `shapes` describes,
from NJT's PCIP Phase 2 existing concourse-level plan (sheet A-001), the sheet `vces` measures.
Platforms 1 and 2 share one outline, since the sheet draws them as one.

This plan defines the frame's y: feet north of platform 5's centerline on it.
x is registered to the Master Plan's as in `vces`, by the platforms' east ends.

Writes `data/shapes_pcip_phase_2.geojson` and `data/shapes_pcip_phase_2.latlon.geojson`.
"""

from shapely import Polygon, box, unary_union

from platform_crowd_model.paths import DATA_DIR
from platform_crowd_model.shapes import XY, Plan, read_drawing, write
from platform_crowd_model.vces import (
    MAX_PLATFORM_LABEL_OFFSET,
    PAGE,
    PDF_UNITS_PER_FOOT,
    SOURCE,
    master_plan_west_edge_x,
    pdf,
    platform_east_ends,
    platform_labels,
    platform_outlines,
)

OUT_GEOJSON = DATA_DIR / "shapes_pcip_phase_2.geojson"
LATLON_GEOJSON = DATA_DIR / "shapes_pcip_phase_2.latlon.geojson"

Y_ORIGIN_PLATFORM = 5

WALL_GRAYS = (0.4, 0.7)
"""The range of grays the sheet draws building elements in, e.g. walls."""

CONCOURSE_FILL = (0.97, 0.97, 0.88)
"""The sheet's fill color for existing concourses."""


def read_plan() -> tuple[Plan, list[tuple[str, Polygon]]]:
    """The plan, and its platforms' outlines, named e.g. `3`, or `1/2` for one they share."""
    page = pdf()[PAGE - 1]
    labels = platform_labels(page)
    outlines = platform_outlines(page)
    west_edge_x = master_plan_west_edge_x(platform_east_ends(outlines, labels))

    def label_y(platform: int) -> float:
        return (labels[platform].y0 + labels[platform].y1) / 2

    def outline_around(platform: int) -> list[float]:
        return next(
            ys
            for ys in ([p.y for p in o] for o in outlines)
            if min(ys) <= label_y(platform) <= max(ys)
        )

    ys = outline_around(Y_ORIGIN_PLATFORM)
    origin_y = (min(ys) + max(ys)) / 2

    def to_ft(x: float, y: float) -> XY:
        # The sheet's y points south, down the page, so it's flipped.
        return (x - west_edge_x) / PDF_UNITS_PER_FOOT, (origin_y - y) / PDF_UNITS_PER_FOOT

    named_outlines = []
    for outline in outlines:
        ys = [p.y for p in outline]
        platforms = sorted(p for p in labels if min(ys) <= label_y(p) <= max(ys))
        if platforms:
            polygon = Polygon([to_ft(p.x, p.y) for p in outline])
            named_outlines.append(("/".join(map(str, platforms)), polygon))
    label_ys = {str(p): to_ft(0, label_y(p))[1] for p in labels}
    max_offset = MAX_PLATFORM_LABEL_OFFSET / PDF_UNITS_PER_FOOT

    def platform_at(point: XY) -> str | None:
        """The platform whose label is nearest across the platforms, if it's near enough."""
        platform = min(label_ys, key=lambda p: abs(label_ys[p] - point[1]))
        return platform if abs(label_ys[platform] - point[1]) <= max_offset else None

    hidden = unary_union([box(*to_ft(r.x0, r.y1), *to_ft(r.x1, r.y0)) for r in labels.values()])
    plan = Plan(
        source=SOURCE,
        drawing=read_drawing(page, lambda p: to_ft(p.x, p.y), 1 / PDF_UNITS_PER_FOOT),
        outlines=named_outlines,
        platform_at=platform_at,
        hidden=hidden,
        wall_grays=WALL_GRAYS,
        concourse_fill=CONCOURSE_FILL,
    )
    return plan, named_outlines


def main() -> None:
    plan, named_outlines = read_plan()
    # This plan defines the frame, and `shapes.frame_outlines` reads it from `OUT_GEOJSON`.
    write(plan, OUT_GEOJSON, LATLON_GEOJSON, registration=named_outlines)
