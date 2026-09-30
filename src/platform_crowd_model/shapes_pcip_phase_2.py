"""
Extract the shapes of platforms 1 to 8 and of what's on them
from NJT's PCIP Phase 2 existing concourse-level plan (sheet A-001), the sheet `vces` measures:
each platform's outline, each VCE's footprint, and the columns and elevators on the platforms.

- A VCE's footprint is the bounding box of its treads and the balustrade lines beside them,
  so it's a little wider than its treads, and T-shaped stairs' footprints include their corners.
  Its `vce_name` is its name in `data/vces.csv`.
- Columns are the small squares on the platforms.
- Elevators are the boxes with an X across them.
- Rooms, walls, and enclosures around the VCEs aren't extracted yet.
- Platforms 1 and 2 share one outline, since the sheet draws them as one.

Shapes are 2D, at the platform level.
Coordinates are in the Master Plan's frame, extended to 2D, in feet:
x is east of the Master Plan's plans' west edge, as in `data/vces.csv`,
and y is north of platform 5's centerline on this sheet,
where east and north are Manhattan's, along its street grid, about 29 degrees off true east.

Writes `data/shapes_pcip_phase_2.geojson` in that frame, and,
registered to OpenStreetMap's platform outlines, `data/shapes_pcip_phase_2_lonlat.geojson`
in longitude and latitude, as GeoJSON requires, so it can be viewed on a map, e.g. on GitHub.
Each has one feature per line.
"""

import csv
import json
from collections.abc import Callable
from typing import Any

import numpy as np
import pymupdf

from platform_crowd_model import platforms_osm
from platform_crowd_model.paths import DATA_DIR
from platform_crowd_model.vces import (
    MAX_PLATFORM_LABEL_OFFSET,
    PAGE,
    PDF_UNITS_PER_FOOT,
    SOURCE,
    Vce,
    flights,
    inside,
    master_plan_west_edge_x,
    pdf,
    platform_east_ends,
    platform_labels,
    platform_outlines,
    vces,
)

VCES_CSV = DATA_DIR / "vces.csv"
OUT_GEOJSON = DATA_DIR / "shapes_pcip_phase_2.geojson"
LONLAT_GEOJSON = DATA_DIR / "shapes_pcip_phase_2_lonlat.geojson"

Y_ORIGIN_PLATFORM = 5
"""The platform whose centerline is y = 0, one in the middle, on every plan."""

REGISTRATION_PLATFORMS = range(3, 9)
"""
Platforms whose east ends register this sheet to OpenStreetMap;
platforms 1 and 2 share one outline here.
"""

MAX_ACROSS_RESIDUAL_FT = 6
"""
How far OpenStreetMap's platforms can be from the sheet's across them.
Along them, their ends are only accurate to within tens of feet,
like their lengths (see `platforms_osm`), so they aren't checked.
"""

BALUSTRADE_REACH = 3
"""How far (PDF units, about 1.7 ft) beside a VCE's treads its balustrade lines can be."""

COLUMN_SIZE = (2, 12)
"""Range of a column's sides (PDF units, about 1.1 to 6.7 ft)."""

MIN_ELEVATOR_SIZE = 6
"""
Smaller boxes (PDF units, about 3.3 ft) with an X across them aren't elevators,
e.g. the hatching on the diagonal stair on platform 5.
"""

FT_DECIMALS = 1
LONLAT_DECIMALS = 7
"""About 1 cm."""

FRAME = (
    "feet: x east of the NY Penn Station Master Plan's plans' west edge, "
    f"y north of platform {Y_ORIGIN_PLATFORM}'s centerline on {SOURCE}, "
    "with east and north along Manhattan's street grid"
)


def axis_lines(page: pymupdf.Page) -> list[tuple[float, float, float, float]]:
    """Every horizontal or vertical line segment (x0, y0, x1, y1), with x0 <= x1 and y0 <= y1."""
    out = []
    for d in page.get_drawings():
        for item in d["items"]:
            if item[0] != "l":
                continue
            a, b = item[1], item[2]
            if abs(a.x - b.x) < 0.3 or abs(a.y - b.y) < 0.3:
                out.append((min(a.x, b.x), min(a.y, b.y), max(a.x, b.x), max(a.y, b.y)))
    return out


def footprint(vce: Vce, lines: list[tuple[float, float, float, float]]) -> pymupdf.Rect:
    """The bounding box of `vce`'s treads and the balustrade lines along them."""
    box = pymupdf.Rect(
        min(f.x0 for f in vce.flights),
        min(f.y0 for f in vce.flights),
        max(f.x1 for f in vce.flights),
        max(f.y1 for f in vce.flights),
    )
    out = pymupdf.Rect(box)
    along_x = vce.flights[0].vertical
    for x0, y0, x1, y1 in lines:
        if along_x and y1 - y0 < 0.3:
            length, overlap = x1 - x0, min(x1, box.x1) - max(x0, box.x0)
            near = box.y0 - BALUSTRADE_REACH <= y0 <= box.y1 + BALUSTRADE_REACH
        elif not along_x and x1 - x0 < 0.3:
            length, overlap = y1 - y0, min(y1, box.y1) - max(y0, box.y0)
            near = box.x0 - BALUSTRADE_REACH <= x0 <= box.x1 + BALUSTRADE_REACH
        else:
            continue
        span = box.width if along_x else box.height
        # Balustrades run alongside the treads, not along the whole platform.
        if near and overlap >= span / 2 and length <= span + 2 * BALUSTRADE_REACH:
            out |= pymupdf.Rect(x0, y0, x1, y1)
    return out


def columns(page: pymupdf.Page, outlines: list[list[pymupdf.Point]]) -> list[pymupdf.Rect]:
    """The small squares on the platforms."""
    out: list[pymupdf.Rect] = []
    for d in page.get_drawings():
        for item in d["items"]:
            if item[0] != "re":
                continue
            r = item[1]
            if not (
                COLUMN_SIZE[0] <= r.width <= COLUMN_SIZE[1]
                and COLUMN_SIZE[0] <= r.height <= COLUMN_SIZE[1]
                and abs(r.width - r.height) < 1
            ):
                continue
            if any(inside((r.tl + r.br) * 0.5, o) for o in outlines) and not any(
                abs(r.x0 - c.x0) < 0.5 and abs(r.y0 - c.y0) < 0.5 for c in out
            ):
                out.append(r)
    return out


def elevators(page: pymupdf.Page, outlines: list[list[pymupdf.Point]]) -> list[pymupdf.Rect]:
    """The boxes on the platforms with an X across them: pairs of crossing diagonals."""
    diagonals = []
    for d in page.get_drawings():
        for item in d["items"]:
            if item[0] == "l":
                a, b = item[1], item[2]
                if abs(a.x - b.x) > 3 and abs(a.y - b.y) > 3:
                    diagonals.append((a, b))
    out = []
    for i, (a, b) in enumerate(diagonals):
        box = pymupdf.Rect(a, b).normalize()
        for c, e in diagonals[i + 1 :]:
            other = pymupdf.Rect(c, e).normalize()
            same_box = all(
                abs(p - q) < 0.5 for p, q in zip(corners(box), corners(other), strict=True)
            )
            crossing = (b.x - a.x) * (b.y - a.y) * (e.x - c.x) * (e.y - c.y) < 0
            if (
                same_box
                and crossing
                and min(box.width, box.height) >= MIN_ELEVATOR_SIZE
                and any(inside((box.tl + box.br) * 0.5, o) for o in outlines)
                and not any(abs(box.x0 - r.x0) < 0.5 and abs(box.y0 - r.y0) < 0.5 for r in out)
            ):
                out.append(box)
    return out


def corners(rect: pymupdf.Rect) -> tuple[float, float, float, float]:
    return rect.x0, rect.y0, rect.x1, rect.y1


def ring(rect: pymupdf.Rect) -> list[pymupdf.Point]:
    return [rect.tl, rect.tr, rect.br, rect.bl]


def counterclockwise(points: list[tuple[float, float]]) -> list[tuple[float, float]]:
    """`points` as a closed ring, counterclockwise, as GeoJSON requires of outer rings."""
    area = sum(
        a[0] * b[1] - b[0] * a[1] for a, b in zip(points, points[1:] + points[:1], strict=True)
    )
    points = points if area > 0 else points[::-1]
    return [*points, points[0]]


def platform_of(y: float, labels: dict[int, pymupdf.Rect]) -> int | None:
    label_y = {p: (r.y0 + r.y1) / 2 for p, r in labels.items()}
    platform = min(label_y, key=lambda p: abs(label_y[p] - y))
    return platform if abs(label_y[platform] - y) <= MAX_PLATFORM_LABEL_OFFSET else None


def osm_platforms() -> dict[int, tuple[complex, complex]]:
    """
    Each platform's east end on its centerline in OpenStreetMap, and the direction it runs east,
    in `platforms_osm.project`'s plane (feet east and north), as complex numbers.
    """
    out = {}
    lat0 = (platforms_osm.BBOX[0] + platforms_osm.BBOX[2]) / 2
    for way in platforms_osm.fetch()["elements"]:
        tags = way.get("tags", {})
        if tags.get("level") != "-3" or not tags.get("ref", "").isdigit():
            continue
        points = np.array(platforms_osm.project(way["geometry"][:-1], lat0))
        center = points.mean(axis=0)
        # The platform's axis is its points' principal direction, pointing east.
        axis = np.linalg.svd(points - center)[2][0]
        axis = axis if axis[0] > 0 else -axis
        east_end = center + axis * ((points - center) @ axis).max()
        out[int(tags["ref"])] = (complex(*east_end), complex(*axis))
    return out


def to_lonlat(
    sheet_east_ends: dict[int, tuple[float, float]],
) -> Callable[[tuple[float, float]], tuple[float, float]]:
    """
    A function from the frame's feet to longitude and latitude, registered to OpenStreetMap:
    rotated to the platforms' average direction there,
    and offset to match their east ends on average,
    so it's only as accurate as OpenStreetMap's outlines, within tens of feet along the platforms.
    The sheet is to scale, so it isn't scaled.
    """
    osm = osm_platforms()
    platforms = [p for p in REGISTRATION_PLATFORMS if p in osm]
    directions = np.array([osm[p][1] for p in platforms])
    rotation = directions.mean() / abs(directions.mean())
    src = np.array([complex(*sheet_east_ends[p]) for p in platforms])
    dst = np.array([osm[p][0] for p in platforms])
    offset = (dst - rotation * src).mean()
    # In the frame's directions: along the platforms (east), and across them (north).
    residuals = (dst - (rotation * src + offset)) / rotation
    along, across = np.abs(residuals.real).max(), np.abs(residuals.imag).max()
    print(
        f"registered to OpenStreetMap, rotated {np.degrees(np.angle(rotation)):.1f} degrees, "
        f"with east ends off by up to {across:.1f} ft across the platforms "
        f"and {along:.1f} ft along them"
    )
    if across > MAX_ACROSS_RESIDUAL_FT:
        raise RuntimeError(f"OpenStreetMap's platforms are off by up to {across:.1f} ft across")
    lat0 = (platforms_osm.BBOX[0] + platforms_osm.BBOX[2]) / 2
    x_scale, y_scale = platforms_osm.scales(lat0)

    def convert(point: tuple[float, float]) -> tuple[float, float]:
        z = rotation * complex(*point) + offset
        return z.real / x_scale, z.imag / y_scale

    return convert


def write(path: Any, features: list[dict[str, Any]], frame: str | None) -> None:
    """Write `features` as a GeoJSON FeatureCollection with one feature per line."""
    header = {"type": "FeatureCollection"}
    if frame:
        # A foreign member: GeoJSON has no way to say coordinates aren't longitude and latitude.
        header["frame"] = frame
    head = json.dumps(header)[:-1]
    lines = ",\n".join(json.dumps(f, ensure_ascii=False) for f in features)
    path.write_text(f'{head}, "features": [\n{lines}\n]}}\n')


def main() -> None:
    page = pdf()[PAGE - 1]
    labels = platform_labels(page)
    outlines = platform_outlines(page)
    east_ends = platform_east_ends(outlines, labels)
    west_edge_x = master_plan_west_edge_x(east_ends)

    def centerline_y(platform: int) -> float:
        """The sheet's y of the middle of the outline around `platform`'s label."""
        label_y = (labels[platform].y0 + labels[platform].y1) / 2
        ys = next(
            ys for ys in ([p.y for p in o] for o in outlines) if min(ys) <= label_y <= max(ys)
        )
        return (min(ys) + max(ys)) / 2

    origin_y = centerline_y(Y_ORIGIN_PLATFORM)

    def ft(point: pymupdf.Point) -> tuple[float, float]:
        # The sheet's y points south, down the page, so it's flipped.
        return (
            (point.x - west_edge_x) / PDF_UNITS_PER_FOOT,
            (origin_y - point.y) / PDF_UNITS_PER_FOOT,
        )

    with VCES_CSV.open() as f:
        named = [row for row in csv.DictReader(f) if row["source"] == "pcip_phase_2"]

    shapes: list[tuple[dict[str, Any], list[pymupdf.Point]]] = []
    for outline in outlines:
        ys = [p.y for p in outline]
        platforms = sorted(p for p, r in labels.items() if min(ys) <= (r.y0 + r.y1) / 2 <= max(ys))
        if platforms:
            props = {"type": "platform", "platform": "/".join(map(str, platforms))}
            shapes.append((props, outline))
    lines = axis_lines(page)
    for vce in sorted(
        vces(flights(page), labels, outlines), key=lambda v: (v.platform, v.flights[0].x0)
    ):
        box = footprint(vce, lines)
        west = round(ft(pymupdf.Point(min(f.x0 for f in vce.flights), 0))[0])
        match = next(
            (
                row
                for row in named
                if int(row["platform"]) == vce.platform and int(row["west_end_ft"]) == west
            ),
            None,
        )
        if match is None:
            raise RuntimeError(f"no VCE in {VCES_CSV.name} at {west} ft on platform {vce.platform}")
        props = {"type": vce.type, "platform": str(vce.platform), "vce_name": match["vce_name"]}
        shapes.append((props, ring(box)))
    for kind, rects in (
        ("column", columns(page, outlines)),
        ("elevator", elevators(page, outlines)),
    ):
        for rect in sorted(rects, key=lambda r: (r.y0, r.x0)):
            platform = platform_of((rect.y0 + rect.y1) / 2, labels)
            props = {"type": kind, "platform": "" if platform is None else str(platform)}
            shapes.append((props, ring(rect)))

    sheet_east_ends = {
        p: ft(pymupdf.Point(east_ends[p], centerline_y(p))) for p in REGISTRATION_PLATFORMS
    }
    lonlat = to_lonlat(sheet_east_ends)
    features: list[dict[str, Any]] = []
    lonlat_features: list[dict[str, Any]] = []
    for props, points in shapes:
        feet = counterclockwise([ft(p) for p in points])
        props = {**props, "level": "platform", "source": SOURCE}
        features.append(
            {
                "type": "Feature",
                "properties": props,
                "geometry": {
                    "type": "Polygon",
                    "coordinates": [[[round(c, FT_DECIMALS) for c in p] for p in feet]],
                },
            }
        )
        lonlat_features.append(
            {
                "type": "Feature",
                "properties": props,
                "geometry": {
                    "type": "Polygon",
                    "coordinates": [[[round(c, LONLAT_DECIMALS) for c in lonlat(p)] for p in feet]],
                },
            }
        )
    write(OUT_GEOJSON, features, FRAME)
    write(LONLAT_GEOJSON, lonlat_features, None)
    counts: dict[str, int] = {}
    for props, _ in shapes:
        counts[props["type"]] = counts.get(props["type"], 0) + 1
    print(f"wrote {len(shapes)} shapes: {counts}")
