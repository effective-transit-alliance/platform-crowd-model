"""
Extract the shapes of platforms 1 to 8 and of what's on them
from NJT's PCIP Phase 2 existing concourse-level plan (sheet A-001), the sheet `vces` measures:
each platform's outline, each VCE's footprint, the columns, elevators, and walls on the platforms,
and the concourses above them.

- A VCE's footprint is the bounding box of its treads and the balustrade lines beside them,
  so it's a little wider than its treads, and T-shaped stairs' footprints include their corners.
  Its `vce_name` is its name in `data/vces.csv`.
- Columns are the small squares on the platforms, whether drawn as rectangles or as 4 lines.
- Elevators are the boxes with an X across them.
- Walls are the other thin gray lines on the platforms, e.g. of rooms and of enclosures around VCEs,
  joined end to end.
  Where they close, they're `enclosure` areas, but most have gaps, e.g. for doors,
  so they're `wall` lines, with their color, so they can be told apart later.
  Some are VCEs' details, e.g. break lines where a flight goes above the cut,
  and the curved stair to the Central Concourse on platform 5,
  whose treads aren't found as a flight.
- Concourses are the areas filled as existing concourse,
  at `level` `concourse`, above the platforms.
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
from shapely import LineString, Polygon, box, line_merge, unary_union

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

WALL_TOLERANCE = 0.3
"""How close (PDF units, about 2 in.) lines have to be to count as on something."""

MAX_WALL_LINE_WIDTH = 3
"""Thicker lines (PDF units) are the legend's streets and buildings above."""

WALL_GRAYS = (0.4, 0.7)
"""The range of grays the sheet draws building elements in."""

MIN_WALL_LENGTH = 1.8
"""Shorter lines (PDF units, about 1 ft), even joined end to end, are details, e.g. break lines."""

CONCOURSE_FILL = (0.97, 0.97, 0.88)
"""The sheet's fill color for existing concourses."""

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


def is_column(r: pymupdf.Rect) -> bool:
    """Whether `r` is the size and shape of a column: a small square."""
    return (
        COLUMN_SIZE[0] <= r.width <= COLUMN_SIZE[1]
        and COLUMN_SIZE[0] <= r.height <= COLUMN_SIZE[1]
        and abs(r.width - r.height) < 1
    )


def columns(page: pymupdf.Page, outlines: list[list[pymupdf.Point]]) -> list[pymupdf.Rect]:
    """The small squares on the platforms."""
    out: list[pymupdf.Rect] = []
    for d in page.get_drawings():
        for item in d["items"]:
            if item[0] != "re":
                continue
            r = item[1]
            if not is_column(r):
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


def item_points(item: tuple[Any, ...]) -> list[pymupdf.Point]:
    """A drawing item's points: a line's ends, a rectangle's or quad's corners, a curve's ends."""
    match item[0]:
        case "l":
            return [item[1], item[2]]
        case "re":
            r = item[1]
            return [r.tl, r.tr, r.br, r.bl, r.tl]
        case "qu":
            q = item[1]
            return [q.ul, q.ur, q.lr, q.ll, q.ul]
        case "c":
            return [item[1], item[4]]
    return []


def walls(
    page: pymupdf.Page,
    outlines: list[list[pymupdf.Point]],
    labels: dict[int, pymupdf.Rect],
    found: list[pymupdf.Rect],
) -> list[tuple[list[pymupdf.Point], tuple[float, ...]]]:
    """
    The other thin lines on the platforms, e.g. rooms' and enclosures' walls, joined end to end,
    with their color, leaving out the platforms' edges, what's under their labels,
    and the lines of what's already `found`: VCEs, columns, and elevators.
    Where walls have gaps, e.g. for doors, they don't close, so they're lines, not areas.
    """
    platforms = [Polygon([(p.x, p.y) for p in o]) for o in outlines]
    on_platforms = unary_union(platforms).buffer(WALL_TOLERANCE)
    edges = unary_union([p.exterior for p in platforms]).buffer(2 * WALL_TOLERANCE)
    hidden = unary_union([box(r.x0, r.y0, r.x1, r.y1) for r in labels.values()])
    known = unary_union([box(*corners(r)).buffer(2 * WALL_TOLERANCE) for r in found])
    by_color: dict[tuple[float, ...], list[LineString]] = {}
    for d in page.get_drawings():
        color = d.get("color")
        # Thicker lines are the streets and buildings above, per the legend,
        # and building elements are drawn in grays, unlike, e.g., the black section marker.
        if (
            color is None
            or (d.get("width") or 0) > MAX_WALL_LINE_WIDTH
            or not all(abs(c - color[0]) < 0.02 for c in color)
            or not WALL_GRAYS[0] <= color[0] <= WALL_GRAYS[1]
        ):
            continue
        for item in d["items"]:
            points = item_points(item)
            if not points:
                continue
            line = LineString([(p.x, p.y) for p in points])
            if (
                line.length >= WALL_TOLERANCE
                and on_platforms.contains(line)
                and not edges.contains(line)
                and not known.contains(line)
                and not hidden.intersects(line)
            ):
                by_color.setdefault(tuple(round(c, 2) for c in color), []).append(line)
    out = []
    for color, lines in sorted(by_color.items()):
        merged = line_merge(unary_union(lines))
        for line in getattr(merged, "geoms", [merged]):
            if line.length < MIN_WALL_LENGTH:
                continue
            out.append(([pymupdf.Point(x, y) for x, y in line.coords], color))
    return out


def concourses(page: pymupdf.Page) -> list[list[pymupdf.Point]]:
    """The outlines of the areas filled as existing concourse, above the platforms."""
    out = []
    for d in page.get_drawings():
        fill = d.get("fill")
        if fill and all(abs(c - f) < 0.01 for c, f in zip(fill, CONCOURSE_FILL, strict=True)):
            points = [p for item in d["items"] for p in item_points(item)]
            if len(points) >= 3:
                out.append(points)
    return out


def ring(rect: pymupdf.Rect) -> list[pymupdf.Point]:
    return [rect.tl, rect.tr, rect.br, rect.bl, rect.tl]


def counterclockwise(ring: list[tuple[float, float]]) -> list[tuple[float, float]]:
    """The closed `ring`, counterclockwise, as GeoJSON requires of outer rings."""
    area = sum(a[0] * b[1] - b[0] * a[1] for a, b in zip(ring, ring[1:], strict=False))
    return ring if area > 0 else ring[::-1]


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
    """Each shape's properties and points, closed for areas and open for lines."""
    for outline in outlines:
        ys = [p.y for p in outline]
        platforms = sorted(p for p, r in labels.items() if min(ys) <= (r.y0 + r.y1) / 2 <= max(ys))
        if platforms:
            props = {"type": "platform", "platform": "/".join(map(str, platforms))}
            shapes.append((props, [*outline, outline[0]]))
    lines = axis_lines(page)
    found = []
    for vce in sorted(
        vces(flights(page), labels, outlines), key=lambda v: (v.platform, v.flights[0].x0)
    ):
        rect = footprint(vce, lines)
        found.append(rect)
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
        shapes.append((props, ring(rect)))
    for kind, rects in (
        ("column", columns(page, outlines)),
        ("elevator", elevators(page, outlines)),
    ):
        found += rects
        for rect in sorted(rects, key=lambda r: (r.y0, r.x0)):
            platform = platform_of((rect.y0 + rect.y1) / 2, labels)
            props = {"type": kind, "platform": "" if platform is None else str(platform)}
            shapes.append((props, ring(rect)))

    for points, color in walls(page, outlines, labels, found):
        closed = len(points) > 3 and points[0] == points[-1]
        rect = pymupdf.Rect(
            min(p.x for p in points),
            min(p.y for p in points),
            max(p.x for p in points),
            max(p.y for p in points),
        )
        platform = platform_of(sum(p.y for p in points) / len(points), labels)
        kind = "column" if closed and is_column(rect) else "enclosure" if closed else "wall"
        props = {
            "type": kind,
            "platform": "" if platform is None else str(platform),
            "color": "#" + "".join(f"{round(c * 255):02x}" for c in color),
        }
        shapes.append((props, points))
    for outline in concourses(page):
        shapes.append(({"type": "concourse", "level": "concourse"}, [*outline, outline[0]]))

    sheet_east_ends = {
        p: ft(pymupdf.Point(east_ends[p], centerline_y(p))) for p in REGISTRATION_PLATFORMS
    }
    lonlat = to_lonlat(sheet_east_ends)
    features: list[dict[str, Any]] = []
    lonlat_features: list[dict[str, Any]] = []
    for props, points in shapes:
        feet = [ft(p) for p in points]
        area = len(points) > 3 and points[0] == points[-1]
        props = {"level": "platform", **props, "source": SOURCE}
        if area:
            feet = counterclockwise(feet)
        for out, points_out, decimals in (
            (features, feet, FT_DECIMALS),
            (lonlat_features, [lonlat(p) for p in feet], LONLAT_DECIMALS),
        ):
            coordinates = [[round(c, decimals) for c in p] for p in points_out]
            geometry = (
                {"type": "Polygon", "coordinates": [coordinates]}
                if area
                else {"type": "LineString", "coordinates": coordinates}
            )
            out.append({"type": "Feature", "properties": props, "geometry": geometry})
    write(OUT_GEOJSON, features, FRAME)
    write(LONLAT_GEOJSON, lonlat_features, None)
    counts: dict[str, int] = {}
    for props, _ in shapes:
        counts[props["type"]] = counts.get(props["type"], 0) + 1
    print(f"wrote {len(shapes)} shapes: {counts}")
