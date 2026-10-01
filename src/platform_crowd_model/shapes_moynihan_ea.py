"""
Extract the shapes of the platforms' west ends and of what's on them, as `shapes` describes,
from the Moynihan Station EA's lower concourse plan (Figure 3-4, February 2010),
the plan `vces_moynihan_ea` measures,
with Moynihan Train Hall's escalators and the West End Concourse,
which the PCIP plans are from before, or cut off.
It's a design, from before the Train Hall was built, so it may differ from what was.

Its VCEs are as `vces_moynihan_ea` measures them, within their windows,
the extent of their treads along the platform, and their treads' median length across it,
leaving out the escalator taken as not built.

It's registered to the frame as in `vces_moynihan_ea`:
its scale from the West End Concourse's width, and x by its stairs.
Its y is fitted to platforms 3 to 8's centerlines on the PCIP Phase 2 plan,
just east of the West End Concourse, where both plans have them.
Its text is drawn as shapes, so its platforms are told apart by row there,
and further west, each piece of platform is named
by the PCIP Phase 1 plan's platform it overlaps most,
in `data/shapes_pcip_phase_1.geojson`, if at least half of it overlaps one, or else it's unnamed.

Writes `data/shapes_moynihan_ea.geojson` and `data/shapes_moynihan_ea.latlon.geojson`.
"""

import json

import numpy as np
import pymupdf
from shapely import GeometryCollection, Point, Polygon, unary_union

from platform_crowd_model.paths import DATA_DIR
from platform_crowd_model.shapes import XY, Plan, item_points, read_drawing, write
from platform_crowd_model.vces_moynihan_ea import (
    PAGE,
    SOURCE,
    VCES,
    measure,
    pdf,
    registration,
    treads,
)

PCIP_PHASE_1_GEOJSON = DATA_DIR / "shapes_pcip_phase_1.geojson"
PCIP_PHASE_2_GEOJSON = DATA_DIR / "shapes_pcip_phase_2.geojson"
OUT_GEOJSON = DATA_DIR / "shapes_moynihan_ea.geojson"
LATLON_GEOJSON = DATA_DIR / "shapes_moynihan_ea.latlon.geojson"

PLATFORM_FILL = (0.905, 0.908, 0.912)
"""The plan's fill color for platforms."""

EAST_OF_CONCOURSE = 630
"""The plan's x (PDF units) just east of the West End Concourse, where platforms are in rows."""

PLATFORM_ROWS = {
    "11": (108, 125),
    "10": (134, 159),
    "9": (159, 175),
    "8": (183, 200),
    "7": (207, 224),
    "6": (231, 246),
    "5": (254, 271),
    "4": (278, 295),
    "3": (302, 319),
}
"""Each platform's rows (PDF units) east of the West End Concourse."""

Y_FIT_PLATFORMS = [str(p) for p in range(3, 9)]
MAX_Y_RESIDUAL_FT = 3
MAX_SCALE_DIFFERENCE = 0.02

MIN_PLATFORM_AREA = 200
"""Smaller pieces (PDF units squared) of platform fill aren't platforms."""

WALL_GRAYS = (0.3, 0.6)
"""The range of grays the plan draws building elements in, e.g. walls."""

CONCOURSE_FILL = (0.976, 0.943, 0.765)
"""The West End Concourse's fill color."""


def outlines_in(path: str) -> dict[str, Polygon]:
    with open(path) as f:
        return {
            feature["properties"]["platform"]: Polygon(feature["geometry"]["coordinates"][0])
            for feature in json.load(f)["features"]
            if feature["properties"]["type"] == "platform"
        }


def main() -> None:
    page = pdf()[PAGE - 1]
    scale, x_offset = registration(page)

    pieces = []
    for d in page.get_drawings():
        fill = d.get("fill")
        if fill and all(abs(c - f) < 0.003 for c, f in zip(fill, PLATFORM_FILL, strict=True)):
            points = [p for item in d["items"] for p in item_points(item)]
            if len(points) >= 3:
                pieces.append(Polygon([(p.x, p.y) for p in points]).buffer(0))
    merged = unary_union([p.buffer(0.3) for p in pieces]).buffer(-0.3)
    pieces = [p for p in getattr(merged, "geoms", [merged]) if p.area >= MIN_PLATFORM_AREA]

    # y: fit the centerlines of the pieces just east of the West End Concourse to PCIP Phase 2's.
    pcip_2 = outlines_in(str(PCIP_PHASE_2_GEOJSON))
    plan_ys = []
    ft_ys = []
    for platform in Y_FIT_PLATFORMS:
        top, bottom = PLATFORM_ROWS[platform]
        (piece,) = [
            p for p in pieces if p.bounds[0] >= EAST_OF_CONCOURSE and top <= p.centroid.y <= bottom
        ]
        plan_ys.append((piece.bounds[1] + piece.bounds[3]) / 2)
        _, y0, _, y1 = pcip_2[platform].bounds
        ft_ys.append((y0 + y1) / 2)
    y_per_unit, y_offset = np.polyfit(plan_ys, ft_ys, 1)
    residuals = [f - (y_per_unit * y + y_offset) for y, f in zip(plan_ys, ft_ys, strict=True)]
    print(
        f"{-y_per_unit:.4f} ft per PDF unit across the platforms, "
        f"centerlines off by up to {max(map(abs, residuals)):.1f} ft"
    )
    if abs(-y_per_unit / scale - 1) > MAX_SCALE_DIFFERENCE:
        raise RuntimeError(f"the plan's scales differ: {scale:.4f} and {-y_per_unit:.4f}")
    if max(map(abs, residuals)) > MAX_Y_RESIDUAL_FT:
        raise RuntimeError("the platforms' centerlines don't fit the PCIP Phase 2 plan's")

    def to_ft(point: pymupdf.Point) -> XY:
        # Registered with one scale, the concourse's; the fit across the platforms only checks it.
        return point.x * scale + x_offset, float(y_per_unit * point.y + y_offset)

    def in_ft(polygon: Polygon) -> Polygon:
        return Polygon([to_ft(pymupdf.Point(x, y)) for x, y in polygon.exterior.coords])

    # Name each piece by its row east of the concourse, or else the PCIP Phase 1 plan's platform.
    pcip_1 = outlines_in(str(PCIP_PHASE_1_GEOJSON))
    named: dict[str, list[Polygon]] = {}
    for piece in pieces:
        feet = in_ft(piece)
        name = next(
            (
                p
                for p, (top, bottom) in PLATFORM_ROWS.items()
                if piece.bounds[0] >= EAST_OF_CONCOURSE and top <= piece.centroid.y <= bottom
            ),
            max(pcip_1, key=lambda p: pcip_1[p].intersection(feet).area),
        )
        if name not in PLATFORM_ROWS and pcip_1[name].intersection(feet).area < feet.area / 2:
            name = ""
        named.setdefault(name, []).append(feet)
    outlines = []
    for name, polygons in sorted(named.items()):
        for polygon in polygons:
            outlines.append((name, polygon))

    def platform_at(point: XY) -> str | None:
        return next((name for name, o in outlines if o.contains(Point(point))), None)

    # Its VCEs as `vces_moynihan_ea` measures them, since the Train Hall's escalators' treads
    # are drawn with two spacings, so they aren't found as one run.
    lines = treads(page)
    vce_boxes = []
    for window in VCES:
        if not window.built:
            continue
        x0, x1, width = measure(window, lines)
        a = to_ft(pymupdf.Point(x0, window.y - width / 2))
        b = to_ft(pymupdf.Point(x1, window.y + width / 2))
        corners = (min(a[0], b[0]), min(a[1], b[1]), max(a[0], b[0]), max(a[1], b[1]))
        vce_boxes.append((str(window.platform), window.type, corners))

    plan = Plan(
        source=SOURCE,
        drawing=read_drawing(page, to_ft, scale),
        outlines=outlines,
        platform_at=platform_at,
        hidden=GeometryCollection(),
        wall_grays=WALL_GRAYS,
        concourse_fill=CONCOURSE_FILL,
        vce_boxes=vce_boxes,
    )
    write(plan, OUT_GEOJSON, LATLON_GEOJSON)
