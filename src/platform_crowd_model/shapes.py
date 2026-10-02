"""
Extract the shapes of the platforms and of what's on them from a vector plan,
shared by each plan's `shapes_<source>` module.

Each module reads its plan's drawing into feet in the Master Plan's frame, extended to 2D
(see `FRAME`), so every threshold here is in feet, whatever the plan's scale,
and says where its platforms and their labels are.
From its lines and fills, this finds:

- each VCE's footprint: each flight's treads and the balustrade lines beside them,
  joined across landings, so a T-shaped stair's is a T.
  Each run of evenly spaced, equally long, parallel treads is a flight,
  and flights are joined into VCEs as in `vces`,
  with flights narrower than 42 in. being escalators.
  Its `vce_name` is that of the nearest VCE of the same type on its platform in `data/vces.csv`,
  within 15 ft, if there is one, and no nearer VCE has it.
- columns: small squares, whether drawn as rectangles or as 4 lines
- elevators: boxes with an X across them
- walls: the other thin gray lines on the platforms, e.g. of rooms and of enclosures around VCEs,
  joined end to end.
  Where they close, they're `enclosure` areas, but most have gaps, e.g. for doors,
  so they're `wall` lines, with their color, so they can be told apart later.
  Some are VCEs' details, e.g. break lines where a flight goes above the cut.
- curved or diagonal stairs and escalators: runs of evenly spaced, nearly parallel treads
  that aren't horizontal or vertical, whose footprint is the band their treads sweep
- concourses: the areas filled as existing concourse, at `level` `concourse`, above the platforms

What's under the platforms' labels is hidden, so it's left out.

Each plan's shapes are written as GeoJSON with one feature per line, in two files:
`data/shapes_<source>.geojson`, in feet in the frame,
and `data/shapes_<source>.latlon.geojson`, in longitude and latitude, as GeoJSON requires,
registered to OpenStreetMap's platform outlines, so it can be viewed on a map, e.g. on GitHub.
"""

import csv
import json
from collections.abc import Callable
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import numpy as np
import pymupdf
from shapely import LineString, Point, Polygon, box, line_merge, unary_union
from shapely.geometry.base import BaseGeometry

from platform_crowd_model import platforms_osm
from platform_crowd_model.paths import DATA_DIR

VCES_CSV = DATA_DIR / "vces.csv"

Y_ORIGIN_PLATFORM = 5
"""The platform whose centerline on the PCIP Phase 2 plan is y = 0, one in the middle."""

FRAME = (
    "feet: x east of the NY Penn Station Master Plan's plans' west edge, "
    f"y north of platform {Y_ORIGIN_PLATFORM}'s centerline on the PCIP Phase 2 existing plan, "
    "with east and north along Manhattan's street grid"
)

REGISTRATION_PLATFORMS = tuple(str(p) for p in range(3, 9))
"""Platforms whose east ends register each plan to OpenStreetMap."""

MAX_ACROSS_RESIDUAL_FT = 6
"""
How far OpenStreetMap's platforms can be from a plan's across them.
Along them, their ends are only accurate to within tens of feet,
like their lengths (see `platforms_osm`), so they aren't checked.
"""

AXIS_TOLERANCE_FT = 0.15
"""Lines within this far of horizontal or vertical along their length are."""

MIN_TREAD_FT = 1.65
TREAD_SPACING_FT = (0.44, 1.78)
"""Range of the spacing between consecutive treads of a flight, about 1 ft."""
MIN_TREADS = 5
TREAD_EXTENT_ROUNDING_FT = 0.55
"""Treads whose ends round to the same multiple of this are equally long and aligned."""

MAX_LANDING_FT = 6.7
"""Maximum gap between two flights of the same stair."""
MAX_JOIN_GAP_FT = 2.5
"""Maximum gap between flights of the same stair meeting at a landing."""
IN_LINE_TOLERANCE_FT = 0.55
MAX_ESCALATOR_WIDTH_IN = 42
"""Flights narrower than this are escalators; the Master Plan's narrowest stair is 44 in."""

MIN_CURVED_TREAD_FT = 3
"""
Shorter lines aren't a curved stair's treads,
but, e.g., pieces of a wall's two curved faces, which are side by side, too.
"""
DUPLICATE_TREAD_FT = 0.4
"""Lines whose midpoints are this close, at about the same angle, are the same line drawn twice."""
MAX_CURVED_TREAD_FT = 12
"""Longer lines that aren't horizontal or vertical aren't a curved stair's treads."""
MAX_CURVED_TREAD_TURN_DEGREES = 15
"""How much a curved stair's consecutive treads can turn."""
CURVED_TREAD_LENGTH_TOLERANCE = 0.2
MIN_TREAD_OFFSET_DEGREES = 45
"""How far off a curved stair's treads' direction the next tread has to be, not end to end."""
"""How much a curved stair's consecutive treads' lengths can differ, as a fraction."""

BALUSTRADE_REACH_FT = 1.7
"""How far beside a VCE's treads its balustrade lines can be."""

COLUMN_SIZE_FT = (1.1, 6.7)
SQUARE_TOLERANCE_FT = 0.55

MIN_ELEVATOR_SIZE_FT = 3.3
"""Smaller boxes with an X across them aren't elevators, e.g. hatching on a stair."""

SAME_BOX_TOLERANCE_FT = 0.28

WALL_TOLERANCE_FT = 0.17
"""How close (about 2 in.) lines have to be to count as on something."""
MAX_WALL_LINE_WIDTH_FT = 1.66
"""Lines drawn thicker are the plans' streets and buildings above."""
MIN_WALL_LENGTH_FT = 1
"""Shorter lines, even joined end to end, are details, e.g. break lines."""

SAME_VCE_TOLERANCE_FT = 15
"""A VCE in `VCES_CSV` of the same type within this many feet is the same VCE."""

OTHER_VCE_TYPE = {"stair": "escalator", "escalator": "stair"}
"""Each VCE type's other type, for naming a VCE a plan draws as the other."""

CURVE_SAMPLES = 8

FT_DECIMALS = 1
LONLAT_DECIMALS = 7
"""About 1 cm."""

type XY = tuple[float, float]


@dataclass(frozen=True)
class Stroke:
    """A stroked path of a plan's drawing, in feet in the frame."""

    points: tuple[XY, ...]
    color: tuple[float, ...]
    width_ft: float
    dashed: bool

    @property
    def closed(self) -> bool:
        return len(self.points) > 3 and self.points[0] == self.points[-1]


@dataclass(frozen=True)
class Fill:
    """A filled path of a plan's drawing, in feet in the frame."""

    points: tuple[XY, ...]
    color: tuple[float, ...]


@dataclass(frozen=True)
class Drawing:
    """A plan's lines and fills, in feet in the frame."""

    strokes: list[Stroke]
    fills: list[Fill]

    def segments(self) -> list[tuple[XY, XY]]:
        """Every straight segment of every stroke."""
        return [(a, b) for s in self.strokes for a, b in zip(s.points, s.points[1:], strict=False)]


def item_points(item: tuple[Any, ...]) -> list[pymupdf.Point]:
    """A drawing item's points: a line's ends, a rectangle's or quad's corners, or a curve's."""
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
            p0, p1, p2, p3 = item[1:5]
            return [
                p0 * (1 - t) ** 3 + p1 * 3 * (1 - t) ** 2 * t + p2 * 3 * (1 - t) * t**2 + p3 * t**3
                for t in np.linspace(0, 1, CURVE_SAMPLES)
            ]
    return []


def read_drawing(
    page: pymupdf.Page, to_ft: Callable[[pymupdf.Point], XY], ft_per_unit: float
) -> Drawing:
    """
    `page`'s drawing in feet, via `to_ft`, which takes the page's unrotated coordinates,
    as `get_drawings` gives them.
    """
    strokes: list[Stroke] = []
    fills = []
    for d in page.get_drawings():
        items = [[to_ft(p) for p in item_points(item)] for item in d["items"]]
        items = [points for points in items if points]
        if d.get("color") is not None:
            width = (d.get("width") or 0) * ft_per_unit
            color = tuple(d["color"])
            dashed = str(d.get("dashes") or "[] 0").strip() not in ("[] 0", "")
            strokes += [Stroke(tuple(points), color, width, dashed) for points in items]
        if d.get("fill") is not None:
            points = [p for item in items for p in item]
            if len(points) >= 3:
                fills.append(Fill(tuple(points), tuple(d["fill"])))
    return Drawing(strokes, fills)


@dataclass(frozen=True)
class Plan:
    """What a plan's module knows about it, in feet in the frame."""

    source: str
    """The plan, e.g. its report, sheet, date, and PDF page."""
    drawing: Drawing
    outlines: list[tuple[str, Polygon]]
    """Each platform's outline, named e.g. `3`, or `1/2` for one shared by platforms 1 and 2."""
    platform_at: Callable[[XY], str | None]
    """Which platform a point is on, if any."""
    hidden: BaseGeometry
    """The platforms' labels, which hide what's under them."""
    wall_grays: tuple[float, float]
    """The range of grays the plan draws building elements in, e.g. walls."""
    concourse_fill: tuple[float, ...] | None = None
    """The plan's fill color for existing concourses, if it has one."""
    vce_boxes: list[tuple[str, str, tuple[float, float, float, float]]] | None = None
    """
    Each VCE's platform, type, and footprint, where the plan's module measures them itself,
    in place of any found as runs of treads there.
    """
    names: Callable[[str, str], list[tuple[float, str]]] | None = None
    """
    Each known VCE of a type on a platform: its midpoint and name,
    by default from `data/vces.csv` (`vce_names`).
    """


@dataclass
class Flight:
    along_x: bool
    """Whether the treads are across the platform, so the flight runs east-west."""
    x0: float
    y0: float
    x1: float
    y1: float

    @property
    def center(self) -> XY:
        return (self.x0 + self.x1) / 2, (self.y0 + self.y1) / 2

    @property
    def width_in(self) -> float:
        """The treads' length."""
        return (self.y1 - self.y0 if self.along_x else self.x1 - self.x0) * 12

    @property
    def escalator(self) -> bool:
        return self.width_in < MAX_ESCALATOR_WIDTH_IN


@dataclass
class Vce:
    platform: str
    flights: list[Flight] = field(default_factory=list)

    @property
    def type(self) -> str:
        return "escalator" if all(f.escalator for f in self.flights) else "stair"

    @property
    def box(self) -> tuple[float, float, float, float]:
        return (
            min(f.x0 for f in self.flights),
            min(f.y0 for f in self.flights),
            max(f.x1 for f in self.flights),
            max(f.y1 for f in self.flights),
        )


def flights(drawing: Drawing) -> list[Flight]:
    """Runs of evenly spaced, equally long, parallel treads, as in `vces.flights`."""
    treads: set[tuple[bool, float, float, float]] = set()
    for a, b in drawing.segments():
        if abs(a[0] - b[0]) < AXIS_TOLERANCE_FT and abs(a[1] - b[1]) >= MIN_TREAD_FT:
            treads.add((True, round(a[0], 2), round(min(a[1], b[1]), 2), round(max(a[1], b[1]), 2)))
        elif abs(a[1] - b[1]) < AXIS_TOLERANCE_FT and abs(a[0] - b[0]) >= MIN_TREAD_FT:
            treads.add(
                (False, round(a[1], 2), round(min(a[0], b[0]), 2), round(max(a[0], b[0]), 2))
            )
    by_extent: dict[tuple[bool, int, int], list[tuple[bool, float, float, float]]] = {}
    for t in treads:
        key = (t[0], round(t[2] / TREAD_EXTENT_ROUNDING_FT), round(t[3] / TREAD_EXTENT_ROUNDING_FT))
        by_extent.setdefault(key, []).append(t)
    out = []
    for group in by_extent.values():
        group.sort(key=lambda t: t[1])
        runs = [[group[0]]]
        for t in group[1:]:
            gap = t[1] - runs[-1][-1][1]
            if gap < TREAD_SPACING_FT[0]:
                continue
            if gap > TREAD_SPACING_FT[1]:
                runs.append([])
            runs[-1].append(t)
        for run in runs:
            if len(run) < MIN_TREADS:
                continue
            along_x, first, last = run[0][0], run[0][1], run[-1][1]
            lo = min(t[2] for t in run)
            hi = max(t[3] for t in run)
            out.append(
                Flight(True, first, lo, last, hi) if along_x else Flight(False, lo, first, hi, last)
            )
    return out


def touching(a: Flight, b: Flight) -> bool:
    """Whether the flights are part of the same VCE, as in `vces.touching`."""
    if a.escalator != b.escalator:
        return False
    in_line = (
        a.along_x
        and b.along_x
        and abs(a.y0 - b.y0) < IN_LINE_TOLERANCE_FT
        and abs(a.y1 - b.y1) < IN_LINE_TOLERANCE_FT
    )
    if a.escalator and not in_line:
        return False
    gap_x = max(a.x0, b.x0) - min(a.x1, b.x1)
    gap_y = max(a.y0, b.y0) - min(a.y1, b.y1)
    if a.along_x and b.along_x and gap_y < IN_LINE_TOLERANCE_FT:
        return gap_x <= MAX_LANDING_FT
    return gap_x <= MAX_JOIN_GAP_FT and gap_y <= MAX_JOIN_GAP_FT


def vces(plan: Plan) -> list[Vce]:
    """The VCEs on the platforms, as in `vces.vces`."""
    by_platform: dict[str, list[Flight]] = {}
    for f in flights(plan.drawing):
        if plan.hidden.contains(Point(f.center)):
            continue
        platform = plan.platform_at(f.center)
        if platform is not None:
            by_platform.setdefault(platform, []).append(f)
    outlines = unary_union([o for _, o in plan.outlines])
    out = []
    for platform, fs in by_platform.items():
        remaining = list(fs)
        while remaining:
            component = [remaining.pop()]
            grew = True
            while grew:
                grew = False
                for f in list(remaining):
                    if any(touching(f, c) for c in component):
                        component.append(f)
                        remaining.remove(f)
                        grew = True
            if any(outlines.contains(Point(f.center)) for f in component):
                out.append(Vce(platform, component))
    return sorted(out, key=lambda v: (v.platform, v.box[0]))


def footprint(vce: Vce, segments: list[tuple[XY, XY]]) -> tuple[float, float, float, float]:
    """The bounding box of `vce`'s treads and the balustrade lines along them."""
    x0, y0, x1, y1 = vce.box
    out = [x0, y0, x1, y1]
    along_x = vce.flights[0].along_x
    reach = BALUSTRADE_REACH_FT
    for a, b in segments:
        if along_x and abs(a[1] - b[1]) < AXIS_TOLERANCE_FT:
            lo, hi, at = min(a[0], b[0]), max(a[0], b[0]), a[1]
            span, overlap = x1 - x0, min(hi, x1) - max(lo, x0)
            near = y0 - reach <= at <= y1 + reach
        elif not along_x and abs(a[0] - b[0]) < AXIS_TOLERANCE_FT:
            lo, hi, at = min(a[1], b[1]), max(a[1], b[1]), a[0]
            span, overlap = y1 - y0, min(hi, y1) - max(lo, y0)
            near = x0 - reach <= at <= x1 + reach
        else:
            continue
        # Balustrades run alongside the treads, not along the whole platform.
        if near and overlap >= span / 2 and hi - lo <= span + 2 * reach:
            out = [
                min(out[0], a[0], b[0]),
                min(out[1], a[1], b[1]),
                max(out[2], a[0], b[0]),
                max(out[3], a[1], b[1]),
            ]
    return out[0], out[1], out[2], out[3]


def vce_outline(vce: Vce, segments: list[tuple[XY, XY]]) -> Polygon:
    """
    `vce`'s footprint: each flight's treads and balustrades, joined by their landings,
    so a T-shaped stair's is a T, not a box including its corners.
    Flights in line are joined by the landing between them,
    and a flight meeting another at a right angle is extended across it.
    """
    parts = [box(*footprint(Vce(vce.platform, [f]), segments)) for f in vce.flights]
    for i, a in enumerate(vce.flights):
        for b in vce.flights[i + 1 :]:
            if not touching(a, b):
                continue
            if a.along_x == b.along_x:
                # The landing between them, across the width they share.
                x0, y0 = min(a.x0, b.x0), min(a.y0, b.y0)
                x1, y1 = max(a.x1, b.x1), max(a.y1, b.y1)
                if a.along_x:
                    y0, y1 = max(a.y0, b.y0), min(a.y1, b.y1)
                else:
                    x0, x1 = max(a.x0, b.x0), min(a.x1, b.x1)
                if x0 < x1 and y0 < y1:
                    parts.append(box(x0, y0, x1, y1))
            else:
                across, along = (b, a) if a.along_x else (a, b)
                # `across` runs north-south, so it's extended north or south across `along`.
                parts.append(
                    box(
                        across.x0,
                        min(across.y0, along.y0),
                        across.x1,
                        max(across.y1, along.y1),
                    )
                )
    # Closing slits where the parts' edges nearly meet.
    r = WALL_TOLERANCE_FT
    joined = unary_union(parts).buffer(r, join_style="mitre").buffer(-r, join_style="mitre")
    return joined if isinstance(joined, Polygon) else joined.convex_hull


def curved_vces(plan: Plan, platforms: BaseGeometry) -> list[tuple[str, str, Polygon]]:
    """
    Curved or diagonal stairs and escalators on the platforms:
    runs of evenly spaced, nearly parallel, equally long treads
    that aren't horizontal or vertical, so `flights` doesn't find them.
    Each one's footprint is the band its treads sweep.
    """
    treads = [
        (a, b)
        for a, b in plan.drawing.segments()
        if abs(a[0] - b[0]) >= AXIS_TOLERANCE_FT
        and abs(a[1] - b[1]) >= AXIS_TOLERANCE_FT
        and MIN_CURVED_TREAD_FT <= float(np.hypot(b[0] - a[0], b[1] - a[1])) <= MAX_CURVED_TREAD_FT
        and platforms.contains(Point((a[0] + b[0]) / 2, (a[1] + b[1]) / 2))
    ]
    # Lines drawn more than once, a little apart, aren't separate treads.
    unique: dict[tuple[int, int, int], tuple[XY, XY]] = {}
    for a, b in treads:
        angle = np.degrees(np.arctan2(b[1] - a[1], b[0] - a[0])) % 180
        key = (
            round((a[0] + b[0]) / 2 / DUPLICATE_TREAD_FT),
            round((a[1] + b[1]) / 2 / DUPLICATE_TREAD_FT),
            round(angle / MAX_CURVED_TREAD_TURN_DEGREES),
        )
        unique.setdefault(key, (a, b))
    treads = list(unique.values())
    if not treads:
        return []
    ends = np.array(treads)
    mids = ends.mean(axis=1)
    vectors = ends[:, 1] - ends[:, 0]
    lengths = np.hypot(vectors[:, 0], vectors[:, 1])
    angles = np.arctan2(vectors[:, 1], vectors[:, 0]) % np.pi
    neighbors: list[list[int]] = [[] for _ in treads]
    for i in range(len(treads)):
        offsets = mids - mids[i]
        distance = np.hypot(*offsets.T)
        turn = np.abs((angles - angles[i] + np.pi / 2) % np.pi - np.pi / 2)
        # Treads are side by side, not end to end, like the pieces of a curved wall.
        along = np.abs(offsets @ (vectors[i] / lengths[i]))
        close = (
            (along <= distance * np.cos(np.radians(MIN_TREAD_OFFSET_DEGREES)))
            & (TREAD_SPACING_FT[0] <= distance)
            & (distance <= TREAD_SPACING_FT[1])
            & (turn <= np.radians(MAX_CURVED_TREAD_TURN_DEGREES))
            & (np.abs(lengths / lengths[i] - 1) <= CURVED_TREAD_LENGTH_TOLERANCE)
        )
        neighbors[i] = [int(j) for j in np.nonzero(close)[0]]
    out = []
    seen: set[int] = set()
    for start in range(len(treads)):
        if start in seen:
            continue
        component = [start]
        seen.add(start)
        for i in component:
            for j in neighbors[i]:
                if j not in seen:
                    seen.add(j)
                    component.append(j)
        if len(component) < MIN_TREADS:
            continue
        band = unary_union(
            [
                LineString(treads[i]).buffer(TREAD_SPACING_FT[1] / 2, cap_style="flat")
                for i in component
            ]
        )
        outline = band if isinstance(band, Polygon) else band.convex_hull
        platform = plan.platform_at((outline.centroid.x, outline.centroid.y))
        if platform is None:
            continue
        width_in = float(np.median(lengths[component])) * 12
        out.append(
            (platform, "escalator" if width_in < MAX_ESCALATOR_WIDTH_IN else "stair", outline)
        )
    return out


def bounds(points: tuple[XY, ...] | list[XY]) -> tuple[float, float, float, float]:
    xs = [p[0] for p in points]
    ys = [p[1] for p in points]
    return min(xs), min(ys), max(xs), max(ys)


def is_column(b: tuple[float, float, float, float]) -> bool:
    """Whether a bounding box is the size and shape of a column: a small square."""
    width, height = b[2] - b[0], b[3] - b[1]
    return (
        COLUMN_SIZE_FT[0] <= width <= COLUMN_SIZE_FT[1]
        and COLUMN_SIZE_FT[0] <= height <= COLUMN_SIZE_FT[1]
        and abs(width - height) < SQUARE_TOLERANCE_FT
    )


def near_any(
    b: tuple[float, float, float, float], found: list[tuple[float, float, float, float]]
) -> bool:
    return any(
        all(abs(p - q) < SAME_BOX_TOLERANCE_FT for p, q in zip(b, f, strict=True)) for f in found
    )


def columns(plan: Plan, platforms: BaseGeometry) -> list[tuple[float, float, float, float]]:
    """The small squares drawn as rectangles on the platforms."""
    out: list[tuple[float, float, float, float]] = []
    for s in plan.drawing.strokes:
        b = bounds(s.points)
        if (
            len(s.points) == 5
            and s.closed
            and is_column(b)
            and platforms.contains(box(*b).centroid)
            and not plan.hidden.intersects(box(*b))
            and not near_any(b, out)
        ):
            out.append(b)
    return out


def elevators(plan: Plan, platforms: BaseGeometry) -> list[tuple[float, float, float, float]]:
    """The boxes on the platforms with an X across them: pairs of crossing diagonals."""
    diagonals = [
        (a, b)
        for a, b in plan.drawing.segments()
        if abs(a[0] - b[0]) > MIN_ELEVATOR_SIZE_FT / 2
        and abs(a[1] - b[1]) > MIN_ELEVATOR_SIZE_FT / 2
    ]
    out: list[tuple[float, float, float, float]] = []
    for i, (a, b) in enumerate(diagonals):
        bb = bounds([a, b])
        for c, e in diagonals[i + 1 :]:
            crossing = (b[0] - a[0]) * (b[1] - a[1]) * (e[0] - c[0]) * (e[1] - c[1]) < 0
            if (
                crossing
                and near_any(bb, [bounds([c, e])])
                and min(bb[2] - bb[0], bb[3] - bb[1]) >= MIN_ELEVATOR_SIZE_FT
                and platforms.contains(box(*bb).centroid)
                and not near_any(bb, out)
            ):
                out.append(bb)
    return out


def walls(plan: Plan, found: list[BaseGeometry]) -> list[tuple[list[XY], tuple[float, ...]]]:
    """
    The other thin, solid gray lines on the platforms, joined end to end, with their color,
    leaving out the platforms' edges, what's under their labels,
    and the lines of what's already `found`: VCEs, columns, and elevators.
    """
    outlines = [o for _, o in plan.outlines]
    on_platforms = unary_union(outlines).buffer(WALL_TOLERANCE_FT)
    edges = unary_union([o.exterior for o in outlines]).buffer(2 * WALL_TOLERANCE_FT)
    known = unary_union([g.buffer(2 * WALL_TOLERANCE_FT) for g in found])
    by_color: dict[tuple[float, ...], list[LineString]] = {}
    for s in plan.drawing.strokes:
        color = s.color
        # Thick or dashed lines are the plans' streets and buildings above.
        if (
            s.width_ft > MAX_WALL_LINE_WIDTH_FT
            or s.dashed
            or not all(abs(c - color[0]) < 0.02 for c in color)
            or not plan.wall_grays[0] <= color[0] <= plan.wall_grays[1]
        ):
            continue
        line = LineString(s.points)
        if (
            line.length >= WALL_TOLERANCE_FT
            and on_platforms.contains(line)
            and not edges.contains(line)
            and not known.contains(line)
            and not plan.hidden.intersects(line)
        ):
            by_color.setdefault(tuple(round(c, 2) for c in color), []).append(line)
    out = []
    for color, lines in sorted(by_color.items()):
        merged = line_merge(unary_union(lines))
        for line in getattr(merged, "geoms", [merged]):
            if line.length >= MIN_WALL_LENGTH_FT:
                out.append(([(float(x), float(y)) for x, y in line.coords], color))
    return out


def ring(b: tuple[float, float, float, float]) -> list[XY]:
    x0, y0, x1, y1 = b
    return [(x0, y0), (x1, y0), (x1, y1), (x0, y1), (x0, y0)]


def counterclockwise(points: list[XY]) -> list[XY]:
    """The closed ring `points`, counterclockwise, as GeoJSON requires of outer rings."""
    area = sum(a[0] * b[1] - b[0] * a[1] for a, b in zip(points, points[1:], strict=False))
    return points if area > 0 else points[::-1]


def hex_color(color: tuple[float, ...]) -> str:
    return "#" + "".join(f"{round(c * 255):02x}" for c in color)


def vce_names(platform: str, vce_type: str) -> list[tuple[float, str]]:
    """Each VCE of `vce_type` on `platform` in `VCES_CSV`: its midpoint and name."""
    with VCES_CSV.open() as f:
        return [
            ((int(row["west_end_ft"]) + int(row["east_end_ft"])) / 2, row["vce_name"])
            for row in csv.DictReader(f)
            if row["platform"] == platform and row["type"] == vce_type
        ]


def shapes(plan: Plan) -> list[tuple[dict[str, str], list[XY]]]:
    """Every shape: its properties, and its points, closed for areas and open for lines."""
    out: list[tuple[dict[str, str], list[XY]]] = []
    for name, outline in plan.outlines:
        out.append(
            ({"type": "platform", "platform": name}, [(x, y) for x, y in outline.exterior.coords])
        )
    platforms = unary_union([o for _, o in plan.outlines])
    segments = plan.drawing.segments()
    measured = [(p, t, box(*b)) for p, t, b in plan.vce_boxes or []]
    found: list[tuple[str, str, Polygon]] = list(measured)
    for platform, vce_type, outline in [
        *((vce.platform, vce.type, vce_outline(vce, segments)) for vce in vces(plan)),
        *curved_vces(plan, platforms),
    ]:
        if not any(outline.intersects(m) for _, _, m in measured):
            found.append((platform, vce_type, outline))
    # Each name goes to the nearest VCE of the same type, so no two VCEs get the same one,
    # and then the rest to the nearest VCE of the other type, taking its type,
    # e.g. an escalator whose balustrades are drawn as wide as a stair.
    names: dict[int, tuple[str, str]] = {}
    pairs = sorted(
        (other, abs(m - (g.bounds[0] + g.bounds[2]) / 2), i, name, named_type)
        for i, (platform, vce_type, g) in enumerate(found)
        for other, named_type in ((False, vce_type), (True, OTHER_VCE_TYPE[vce_type]))
        for m, name in (plan.names or vce_names)(platform, named_type)
    )
    for _, distance, i, name, named_type in pairs:
        if (
            distance <= SAME_VCE_TOLERANCE_FT
            and i not in names
            and name not in (n for n, _ in names.values())
        ):
            names[i] = name, named_type
    for i, (platform, vce_type, g) in enumerate(found):
        name, named_type = names.get(i, ("", vce_type))
        props = {"type": named_type, "platform": platform, "vce_name": name}
        out.append((props, [(float(x), float(y)) for x, y in g.exterior.coords]))
    shapes_found: list[BaseGeometry] = [g for _, _, g in found]
    for kind, boxes in (
        ("column", columns(plan, platforms)),
        ("elevator", elevators(plan, platforms)),
    ):
        shapes_found += [box(*b) for b in boxes]
        for b in sorted(boxes, key=lambda b: (-b[3], b[0])):
            center = ((b[0] + b[2]) / 2, (b[1] + b[3]) / 2)
            out.append(({"type": kind, "platform": plan.platform_at(center) or ""}, ring(b)))
    for points, color in walls(plan, shapes_found):
        b = bounds(points)
        closed = len(points) > 3 and points[0] == points[-1]
        kind = "column" if closed and is_column(b) else "enclosure" if closed else "wall"
        center = ((b[0] + b[2]) / 2, (b[1] + b[3]) / 2)
        props = {
            "type": kind,
            "platform": plan.platform_at(center) or "",
            "color": hex_color(color),
        }
        out.append((props, points))
    if plan.concourse_fill is not None:
        for f in plan.drawing.fills:
            # Leaving out the legend's swatch, which isn't over the platforms.
            if all(
                abs(c - d) < 0.01 for c, d in zip(f.color, plan.concourse_fill, strict=True)
            ) and platforms.intersects(Polygon(f.points)):
                out.append(({"type": "concourse", "level": "concourse"}, [*f.points, f.points[0]]))
    return out


def osm_platforms() -> dict[str, tuple[complex, complex]]:
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
        out[tags["ref"]] = (complex(*east_end), complex(*axis))
    return out


def east_ends(outlines: list[tuple[str, Polygon]]) -> dict[str, XY]:
    """Each platform's east end on its centerline, from its outline."""
    out = {}
    for name, outline in outlines:
        x0, y0, x1, y1 = outline.bounds
        out[name] = (x1, (y0 + y1) / 2)
    return out


def to_lonlat(plan_east_ends: dict[str, XY]) -> Callable[[XY], XY]:
    """
    A function from the frame's feet to longitude and latitude, registered to OpenStreetMap:
    rotated to the platforms' average direction there,
    and offset to match their east ends on average,
    so it's only as accurate as OpenStreetMap's outlines, within tens of feet along the platforms.
    The plans are to scale, so it isn't scaled.
    """
    osm = osm_platforms()
    platforms = [p for p in REGISTRATION_PLATFORMS if p in osm and p in plan_east_ends]
    directions = np.array([osm[p][1] for p in platforms])
    rotation = directions.mean() / abs(directions.mean())
    src = np.array([complex(*plan_east_ends[p]) for p in platforms])
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

    def convert(point: XY) -> XY:
        z = rotation * complex(*point) + offset
        return z.real / x_scale, z.imag / y_scale

    return convert


def write_geojson(path: Path, features: list[dict[str, Any]], frame: str | None) -> None:
    """Write `features` as a GeoJSON FeatureCollection with one feature per line."""
    header: dict[str, str] = {"type": "FeatureCollection"}
    if frame:
        # A foreign member: GeoJSON has no way to say coordinates aren't longitude and latitude.
        header["frame"] = frame
    head = json.dumps(header)[:-1]
    lines = ",\n".join(json.dumps(f, ensure_ascii=False) for f in features)
    path.write_text(f'{head}, "features": [\n{lines}\n]}}\n')


FRAME_GEOJSON = DATA_DIR / "shapes_pcip_phase_2.geojson"
"""The plan that defines the frame's y, whose platforms register every plan to OpenStreetMap."""


def frame_outlines() -> list[tuple[str, Polygon]]:
    """The platforms' outlines in `FRAME_GEOJSON`."""
    with FRAME_GEOJSON.open() as f:
        return [
            (feature["properties"]["platform"], Polygon(feature["geometry"]["coordinates"][0]))
            for feature in json.load(f)["features"]
            if feature["properties"]["type"] == "platform"
        ]


def write(
    plan: Plan,
    feet_path: Path,
    lonlat_path: Path,
    registration: list[tuple[str, Polygon]] | None = None,
) -> None:
    """Write `plan`'s shapes to `feet_path`, and to `lonlat_path` in longitude and latitude."""
    found = shapes(plan)
    # Every plan is in the same frame, so it's registered the same way, by `frame_outlines`.
    lonlat = to_lonlat(east_ends(frame_outlines() if registration is None else registration))
    features: list[dict[str, Any]] = []
    lonlat_features: list[dict[str, Any]] = []
    for props, points in found:
        area = len(points) > 3 and points[0] == points[-1]
        feet = counterclockwise(points) if area else points
        props = {"level": "platform", **props, "source": plan.source}
        for out, converted, decimals in (
            (features, feet, FT_DECIMALS),
            (lonlat_features, [lonlat(p) for p in feet], LONLAT_DECIMALS),
        ):
            coordinates = [[round(c, decimals) for c in p] for p in converted]
            geometry = (
                {"type": "Polygon", "coordinates": [coordinates]}
                if area
                else {"type": "LineString", "coordinates": coordinates}
            )
            out.append({"type": "Feature", "properties": props, "geometry": geometry})
    write_geojson(feet_path, features, FRAME)
    write_geojson(lonlat_path, lonlat_features, None)
    counts: dict[str, int] = {}
    for props, _ in found:
        counts[props["type"]] = counts.get(props["type"], 0) + 1
    print(f"wrote {len(found)} shapes: {counts}")
