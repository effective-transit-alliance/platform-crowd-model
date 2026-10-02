"""
Find the VCEs on a vector plan of the platforms,
shared by each plan's `shapes_<source>` module, which reads its plan,
and by `vces`, which takes `data/vces.csv`'s VCEs from them.

Each module reads its plan's drawing into feet in the Master Plan's frame, extended to 2D:
x east of the Master Plan's plans' west edge,
and y north of platform 5's centerline on the PCIP Phase 2 existing plan,
so every threshold here is in feet, whatever the plan's scale,
and says where its platforms and their labels are.
From its lines, this finds each VCE's footprint:
each flight's treads and the balustrade lines beside them,
joined across landings, so a T-shaped stair's is a T.
Each run of evenly spaced, equally long, parallel treads is a flight,
and flights are joined into VCEs as in `vces`,
with flights narrower than 42 in. being escalators.
Curved or diagonal stairs and escalators are runs of evenly spaced, nearly parallel treads
that aren't horizontal or vertical, whose footprint is the band their treads sweep.

What's under the platforms' labels is hidden, so it's left out.
"""

from collections.abc import Callable
from dataclasses import dataclass, field
from typing import Any

import numpy as np
import pymupdf
from shapely import LineString, Point, Polygon, box, unary_union
from shapely.geometry.base import BaseGeometry

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

MIN_CURVED_TREAD_FT = 2.4
"""
Shorter lines aren't a curved stair's treads,
but, e.g., pieces of a wall's two curved faces, which are side by side, too.
Long enough for a slightly skewed stair's half treads,
e.g. the PCIP Phase 1 plan's on platform 9, about 400' along it.
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

WALL_TOLERANCE_FT = 0.17
"""How close (about 2 in.) lines have to be to count as on something."""

CURVE_SAMPLES = 8

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
    by default from `data/vces.csv`.
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
    def width_in(self) -> float:
        """
        The width across the treads, at the narrowest flight.
        Flights side by side along the same stretch, e.g. treads drawn in two halves,
        are one flight, so their widths add up.
        A T-shaped stair's width is its upper flight's,
        the one perpendicular to its two flights along the platform.
        """
        flights = self.flights
        if len({f.along_x for f in flights}) == 2:
            flights = [f for f in flights if not f.along_x]
        runs: list[list[Flight]] = []
        for f in flights:
            for run in runs:
                if any(side_by_side(f, g) for g in run):
                    run.append(f)
                    break
            else:
                runs.append([f])
        widths = []
        for run in runs:
            along_x = run[0].along_x
            lo = min(f.y0 if along_x else f.x0 for f in run)
            hi = max(f.y1 if along_x else f.x1 for f in run)
            widths.append((hi - lo) * 12)
        return min(widths)

    @property
    def box(self) -> tuple[float, float, float, float]:
        return (
            min(f.x0 for f in self.flights),
            min(f.y0 for f in self.flights),
            max(f.x1 for f in self.flights),
            max(f.y1 for f in self.flights),
        )


def flights(drawing: Drawing) -> list[Flight]:
    """Runs of evenly spaced, equally long, parallel treads."""
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


def side_by_side(a: Flight, b: Flight) -> bool:
    """Whether the flights run along the same stretch, overlapping by at least half."""
    if a.along_x != b.along_x:
        return False
    a0, a1, b0, b1 = (a.x0, a.x1, b.x0, b.x1) if a.along_x else (a.y0, a.y1, b.y0, b.y1)
    overlap = min(a1, b1) - max(a0, b0)
    return overlap >= min(a1 - a0, b1 - b0) / 2


def touching(a: Flight, b: Flight) -> bool:
    """Whether the flights are part of the same VCE: touching, or across a short landing."""
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
    """
    The VCEs on the platforms, as connected flights,
    each on the platform whose row it's in, with at least one flight on a platform's outline,
    since some stairs' upper flights are drawn past the platform's edge.
    """
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


def curved_vces(plan: Plan, platforms: BaseGeometry) -> list[tuple[str, str, Polygon, float]]:
    """
    Curved or diagonal stairs and escalators on the platforms:
    runs of evenly spaced, nearly parallel, equally long treads
    that aren't horizontal or vertical, so `flights` doesn't find them.
    Each one's footprint is the band its treads sweep, and its width is its treads' median length.
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
        vce_type = "escalator" if width_in < MAX_ESCALATOR_WIDTH_IN else "stair"
        out.append((platform, vce_type, outline, width_in))
    return out


def found_vces(plan: Plan) -> list[tuple[str, str, Polygon, float | None]]:
    """
    Every VCE on `plan`'s platforms: its platform, type, footprint, and width (in.),
    from `plan.vce_boxes`, without widths, or else as runs of treads (`vces`),
    or curved (`curved_vces`).
    `data/vces.csv` takes its VCEs from these, too, so they match the shapes.
    """
    platforms = unary_union([o for _, o in plan.outlines])
    segments = plan.drawing.segments()
    measured: list[tuple[str, str, Polygon, float | None]] = [
        (p, t, box(*b), None) for p, t, b in plan.vce_boxes or []
    ]
    found = list(measured)
    for platform, vce_type, outline, width_in in [
        *((vce.platform, vce.type, vce_outline(vce, segments), vce.width_in) for vce in vces(plan)),
        *curved_vces(plan, platforms),
    ]:
        if not any(outline.intersects(m) for _, _, m, _ in measured):
            found.append((platform, vce_type, outline, width_in))
    return found
