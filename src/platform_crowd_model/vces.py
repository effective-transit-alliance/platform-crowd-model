"""
Estimate the widths and positions of every VCE on platforms 1 to 8
by measuring their treads on a scaled vector drawing,
and compare them with the Master Plan's (`data/vces_existing_master_plan.csv`).

NJT's PCIP Phase 2 drawings (November 2020) include an existing concourse-level plan
(sheet A-001, PDF page 45) drawn at 1" = 40', with each stair's and escalator's treads as lines,
so a tread's length is roughly a stair's width, or an escalator's step width.
It shows platforms 1 to 8, including the West End Concourse; no such plan of 9 to 11 was found.

Each run of evenly spaced, equally long, parallel treads is a flight.
Flights of a stair are merged when they touch or are separated by a short landing,
including stairs whose treads are drawn in two halves and T-shaped stairs.
A stair's width is the extent of its treads across it, at its narrowest flight.
Escalators are narrower than any stair, so flights under 42 in. are escalators,
and their treads are their steps, narrower than their balustrades.
Flights outside the platforms' outlines are skipped.
The platforms' labels cover an escalator the Master Plan has
about 230 ft along each of platforms 3 to 8,
whose treads the sheet still has under them, but clipped,
so their positions are the sheet's, but their type and width are the Master Plan's.

Positions are converted to the Master Plan's frame: feet east of its plans' west edge.
Its plans and this sheet register to within a few feet:
their platforms' east ends are all the same distance apart.
Each VCE is matched to a Master Plan VCE of the same type on the same platform within 15 ft.
Matched VCEs have the Master Plan's width,
and the rest have only this sheet's width, marked `estimated`.

Platforms 9 to 11 aren't on that plan, so their VCEs are estimated from
NJT's January 2022 station directory and the Master Plan instead;
see `vces_njt_directory`.
Their widths are the Master Plan's, or else typical of platforms 1 to 8, marked `typical`.

Writes `data/vces.csv`,
and `data/platform_east_ends.csv`: each platform's east end in the Master Plan's frame,
from this sheet's outlines on platforms 1 to 8,
and from the Master Plan's (`data/platform_east_ends_master_plan.csv`) on platforms 9 to 11.
"""

import bisect
import csv
import statistics
from collections import defaultdict
from collections.abc import Callable
from dataclasses import dataclass, field
from functools import cache
from typing import Any
from urllib.request import Request, urlopen

import pymupdf

from platform_crowd_model.paths import CACHE_DIR, DATA_DIR

MASTER_PLAN_CSV = DATA_DIR / "vces_existing_master_plan.csv"
DIRECTORY_CSV = DATA_DIR / "vces_njt_directory.csv"
MASTER_PLAN_EAST_ENDS_CSV = DATA_DIR / "platform_east_ends_master_plan.csv"
MOYNIHAN_EA_CSV = DATA_DIR / "vces_moynihan_ea.csv"
EAST_ENDS_CSV = DATA_DIR / "platform_east_ends.csv"
OUT_CSV = DATA_DIR / "vces.csv"
PDF_CACHE = CACHE_DIR / "pcip-2-conceptual-design-preliminary-drawings.pdf"
PDF_URL = "https://liamblank.com/wp-content/uploads/2026/09/pcip-2-conceptual-design-preliminary-drawings.pdf"
PAGE = 45
"""1-indexed PDF page of sheet A-001, "Existing Plan Overall"."""
SOURCE = f"PCIP Phase 2 Appendix A, sheet A-001, November 2020, PDF page {PAGE}"

PCIP_PHASE_2_VCE_SOURCE = "pcip_phase_2"
"""`data/vces.csv`'s `source` for VCEs measured on this sheet (`SOURCE`)."""

PCIP_PHASE_1_VCE_SOURCE = "pcip_phase_1"
"""
`data/vces.csv`'s `source` for platforms 9 to 11's VCEs measured on
NJT's PCIP Phase 1 existing plan (Appendix A, sheet A-001, July 2019).
"""

MOYNIHAN_EA_VCE_SOURCE = "moynihan_ea"
"""`data/vces.csv`'s `source` for VCEs from the Moynihan Station EA's plan, in `MOYNIHAN_EA_CSV`."""

PDF_UNITS_PER_FOOT = (507.2 - 290.9) / 120
"""From the centers of the scale bar's 0' and 120' labels."""

PLATFORM_FILL = (0.74, 0.74, 0.75)
"""The sheet's fill color for existing platforms."""

MIN_TREAD_LENGTH = 3
"""Shorter lines (PDF units) aren't treads."""

TREAD_SPACING = (0.8, 3.2)
"""Range of the spacing (PDF units) between consecutive treads of a flight, about 1 ft."""

MIN_TREADS = 5
"""Fewer parallel lines aren't a flight."""

MAX_LANDING = 12
"""Maximum gap (PDF units, about 6.7 ft) between two flights of the same stair."""

MAX_JOIN_GAP = 4.5
"""
Maximum gap (PDF units, about 2.5 ft) between flights of the same stair meeting at a landing,
since some flights' first treads aren't drawn.
"""

MAX_ESCALATOR_WIDTH_IN = 42
"""Flights narrower than this are escalators; the Master Plan's narrowest stair is 44 in."""

MIN_PLATFORM_LENGTH = 1000
"""Shorter shapes (PDF units) with the platforms' fill color aren't platforms."""

LABEL_PADDING = 8
"""Padding (PDF units) around a platform's label's words, to cover the box around it."""

MAX_PLATFORM_LABEL_OFFSET = 30
"""Maximum vertical distance (PDF units) between a flight and its platform's label."""

SAME_VCE_TOLERANCE_FT = 15
"""A Master Plan VCE of the same type within this many feet is the same VCE."""


@dataclass
class Flight:
    vertical: bool
    """Whether the treads are vertical lines, so the flight runs east-west, along the platform."""
    x0: float
    y0: float
    x1: float
    y1: float
    treads: int

    @property
    def center(self) -> pymupdf.Point:
        return pymupdf.Point((self.x0 + self.x1) / 2, (self.y0 + self.y1) / 2)

    @property
    def width(self) -> float:
        """The treads' length (PDF units)."""
        return self.y1 - self.y0 if self.vertical else self.x1 - self.x0

    @property
    def width_in(self) -> float:
        return self.width / PDF_UNITS_PER_FOOT * 12


@dataclass
class Vce:
    platform: int
    flights: list[Flight] = field(default_factory=list)
    hidden: bool = False
    """Whether its treads are under a platform's label, which clips them."""

    @property
    def type(self) -> str:
        if all(f.width_in < MAX_ESCALATOR_WIDTH_IN for f in self.flights):
            return "escalator"
        return "stair"

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
        if len({f.vertical for f in flights}) == 2:
            flights = [f for f in flights if not f.vertical]
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
            vertical = run[0].vertical
            lo = min(f.y0 if vertical else f.x0 for f in run)
            hi = max(f.y1 if vertical else f.x1 for f in run)
            widths.append((hi - lo) / PDF_UNITS_PER_FOOT * 12)
        return min(widths)

    @property
    def x0(self) -> float:
        return min(f.x0 for f in self.flights)

    @property
    def x1(self) -> float:
        return max(f.x1 for f in self.flights)


def pdf() -> pymupdf.Document:
    if not PDF_CACHE.exists():
        CACHE_DIR.mkdir(exist_ok=True)
        request = Request(
            PDF_URL, headers={"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) Firefox/130.0"}
        )
        with urlopen(request) as response:
            PDF_CACHE.write_bytes(response.read())
    return pymupdf.open(PDF_CACHE)


def flights(page: pymupdf.Page) -> list[Flight]:
    """Runs of evenly spaced, equally long, parallel lines."""
    lines: set[tuple[bool, float, float, float]] = set()
    for d in page.get_drawings():
        for item in d["items"]:
            if item[0] != "l":
                continue
            a, b = item[1], item[2]
            if abs(a.x - b.x) < 0.3 and abs(a.y - b.y) >= MIN_TREAD_LENGTH:
                lines.add((True, round(a.x, 1), round(min(a.y, b.y), 1), round(max(a.y, b.y), 1)))
            elif abs(a.y - b.y) < 0.3 and abs(a.x - b.x) >= MIN_TREAD_LENGTH:
                lines.add((False, round(a.y, 1), round(min(a.x, b.x), 1), round(max(a.x, b.x), 1)))
    # Group lines with the same orientation and extent, then split them into evenly spaced runs.
    by_extent: dict[tuple[bool, int, int], list[tuple[bool, float, float, float]]] = {}
    for line in lines:
        by_extent.setdefault((line[0], round(line[2]), round(line[3])), []).append(line)
    out = []
    for group in by_extent.values():
        group.sort(key=lambda line: line[1])
        runs = [[group[0]]]
        for line in group[1:]:
            gap = line[1] - runs[-1][-1][1]
            if gap < TREAD_SPACING[0]:
                continue
            if gap > TREAD_SPACING[1]:
                runs.append([])
            runs[-1].append(line)
        for run in runs:
            if len(run) < MIN_TREADS:
                continue
            vertical, first, last = run[0][0], run[0][1], run[-1][1]
            lo = min(line[2] for line in run)
            hi = max(line[3] for line in run)
            if vertical:
                out.append(Flight(True, first, lo, last, hi, len(run)))
            else:
                out.append(Flight(False, lo, first, hi, last, len(run)))
    return out


def platform_labels(page: pymupdf.Page) -> dict[int, pymupdf.Rect]:
    """Each platform's label, including the pill-shaped box around it."""
    words = page.get_text("words")
    out = {}
    for w, next_w in zip(words, words[1:], strict=False):
        if w[4] == "PLATFORM" and str(next_w[4]).isdigit() and int(next_w[4]) not in out:
            box = pymupdf.Rect(w[:4]) | pymupdf.Rect(next_w[:4])
            out[int(next_w[4])] = box + (
                -LABEL_PADDING,
                -LABEL_PADDING,
                LABEL_PADDING,
                LABEL_PADDING,
            )
    return out


def platform_outlines(page: pymupdf.Page) -> list[list[pymupdf.Point]]:
    """The polygons of the platforms' outlines; platforms 1 and 2 share one."""
    out = []
    for d in page.get_drawings():
        fill = d.get("fill")
        if fill is None or any(abs(a - b) > 0.01 for a, b in zip(fill, PLATFORM_FILL, strict=True)):
            continue
        if d["rect"].width > MIN_PLATFORM_LENGTH and all(item[0] == "l" for item in d["items"]):
            out.append([item[1] for item in d["items"]])
    return out


def inside(point: pymupdf.Point, polygon: list[pymupdf.Point]) -> bool:
    """Whether the point is inside the polygon, by ray casting."""
    result = False
    for a, b in zip(polygon, polygon[1:] + polygon[:1], strict=True):
        if (a.y > point.y) != (b.y > point.y):
            x = a.x + (point.y - a.y) / (b.y - a.y) * (b.x - a.x)
            if point.x < x:
                result = not result
    return result


def platform_east_ends(
    outlines: list[list[pymupdf.Point]], labels: dict[int, pymupdf.Rect]
) -> dict[int, float]:
    """The x coordinate of each platform's east end, from the outline around its label."""
    out = {}
    for platform, label in labels.items():
        for polygon in outlines:
            ys = [point.y for point in polygon]
            if min(ys) <= (label.y0 + label.y1) / 2 <= max(ys):
                out[platform] = max(point.x for point in polygon)
    return out


def side_by_side(a: Flight, b: Flight) -> bool:
    """Whether the flights run along the same stretch, overlapping by at least half."""
    if a.vertical != b.vertical:
        return False
    a0, a1, b0, b1 = (a.x0, a.x1, b.x0, b.x1) if a.vertical else (a.y0, a.y1, b.y0, b.y1)
    overlap = min(a1, b1) - max(a0, b0)
    return overlap >= min(a1 - a0, b1 - b0) / 2


def touching(a: Flight, b: Flight) -> bool:
    """Whether the flights are part of the same stair: touching, or across a short landing."""
    if (a.width_in < MAX_ESCALATOR_WIDTH_IN) != (b.width_in < MAX_ESCALATOR_WIDTH_IN):
        return False
    # Escalators side by side are separate escalators; only merge them along their length.
    in_line = a.vertical and b.vertical and abs(a.y0 - b.y0) < 1 and abs(a.y1 - b.y1) < 1
    if a.width_in < MAX_ESCALATOR_WIDTH_IN and not in_line:
        return False
    gap_x = max(a.x0, b.x0) - min(a.x1, b.x1)
    gap_y = max(a.y0, b.y0) - min(a.y1, b.y1)
    if a.vertical and b.vertical and gap_y < 1:
        return gap_x <= MAX_LANDING
    return gap_x <= MAX_JOIN_GAP and gap_y <= MAX_JOIN_GAP


def vces(
    all_flights: list[Flight],
    labels: dict[int, pymupdf.Rect],
    outlines: list[list[pymupdf.Point]],
) -> list[Vce]:
    by_platform: dict[int, list[Flight]] = {}
    for f in all_flights:
        center = f.center
        label_y = {p: (r.y0 + r.y1) / 2 for p, r in labels.items()}
        platform = min(label_y, key=lambda p: abs(label_y[p] - center.y))
        if abs(label_y[platform] - center.y) > MAX_PLATFORM_LABEL_OFFSET:
            continue
        by_platform.setdefault(platform, []).append(f)
    out = []
    for platform, fs in by_platform.items():
        # Connected components of touching flights.
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
            # Some stairs' upper flights are drawn past the platform's edge, so only require one
            # of a VCE's flights to be on a platform.
            if any(inside(f.center, polygon) for f in component for polygon in outlines):
                hidden = any(f.center in label for f in component for label in labels.values())
                out.append(Vce(platform, component, hidden))
    return out


def master_plan_west_edge_x(east_ends: dict[int, float]) -> float:
    """
    The sheet's x (PDF units) of the Master Plan's plans' west edge,
    registering each platform's east end on the sheet to the Master Plan's.
    """
    # Platforms 1 and 2 are drawn differently on both, so they aren't used to register them.
    with MASTER_PLAN_EAST_ENDS_CSV.open() as f:
        master_plan_east_ends = {
            int(row["platform"]): int(row["east_end_ft"]) for row in csv.DictReader(f)
        }
    offsets = [
        east_ends[p] - master_plan_east_ends[p] * PDF_UNITS_PER_FOOT
        for p in east_ends.keys() & master_plan_east_ends.keys()
    ]
    spread_ft = (max(offsets) - min(offsets)) / PDF_UNITS_PER_FOOT
    print(f"registered to the Master Plan to within {spread_ft:.1f} ft")
    return sum(offsets) / len(offsets)


def sheet_vces() -> tuple[list[dict[str, Any]], dict[int, int]]:
    """
    Every VCE on platforms 1 to 8 measured on the sheet, as rows of `OUT_CSV`,
    and each platform's east end, both in the Master Plan's frame.
    """
    page = pdf()[PAGE - 1]
    labels = platform_labels(page)
    outlines = platform_outlines(page)
    east_ends = platform_east_ends(outlines, labels)
    west_edge_x = master_plan_west_edge_x(east_ends)

    def ft(x: float) -> int:
        return round((x - west_edge_x) / PDF_UNITS_PER_FOOT)

    with MASTER_PLAN_CSV.open() as f:
        master_plan = list(csv.DictReader(f))
    found = sorted(vces(flights(page), labels, outlines), key=lambda v: (v.platform, v.x0))
    out = []
    numbers: dict[int, int] = {}
    for v in found:
        numbers[v.platform] = numbers.get(v.platform, 0) + 1
        mid = (ft(v.x0) + ft(v.x1)) / 2
        near = [
            m
            for m in master_plan
            if int(m["platform"]) == v.platform
            and abs(float(m["midpoint_ft"]) - mid) <= SAME_VCE_TOLERANCE_FT
        ]
        # A hidden VCE's treads are clipped, so its type and width are the Master Plan's.
        same_type = [m for m in near if v.hidden or m["type"] == v.type]
        match = min(same_type, key=lambda m: abs(float(m["midpoint_ft"]) - mid), default=None)
        vce_type = match["type"] if v.hidden and match else v.type
        west, east = ft(v.x0), ft(v.x1)
        source = PCIP_PHASE_2_VCE_SOURCE
        notes = []
        if v.hidden:
            ends = under_label(v.platform, west, east)
            if ends is not None:
                (west, east), source = ends, PCIP_PHASE_1_VCE_SOURCE
                notes.append(
                    "Its treads are under the PCIP Phase 2 plan's platform label, "
                    "which clips them, so its position is the PCIP Phase 1 plan's, "
                    "and its type and width are the Master Plan's."
                )
            else:
                notes.append(
                    "Its treads are under the sheet's platform label, which clips them, "
                    "so its type and width are the Master Plan's."
                )
        if match is None and near:
            other = min(near, key=lambda m: abs(float(m["midpoint_ft"]) - mid))
            notes.append(
                f"The Master Plan has a {other['width_in']} in. {other['type']} here instead."
            )
        if vce_type == "escalator" and not v.hidden:
            notes.append("An escalator's treads are its steps, narrower than its balustrades.")
        if v.platform == 4:
            notes.append("NJT replaced an escalator on tracks 7/8 with stairs in about 2021.")
        out.append(
            {
                "platform": v.platform,
                "vce_name": f"P{v.platform}-S{numbers[v.platform]}",
                "type": vce_type,
                "west_end_ft": west,
                "east_end_ft": east,
                "estimated_width_in": "" if v.hidden else round(v.width_in),
                "master_plan_width_in": match["width_in"] if match else "",
                "width_source": "master_plan" if match else "estimated",
                "source": source,
                "notes": " ".join(notes),
            }
        )
    return out, {p: ft(x) for p, x in east_ends.items()}


def moynihan_ea_vces() -> list[dict[str, Any]]:
    """The VCEs in `MOYNIHAN_EA_CSV` that were built, as rows of `OUT_CSV`."""
    with MOYNIHAN_EA_CSV.open() as f:
        return [
            {
                "platform": int(row["platform"]),
                "vce_name": "",
                "type": row["type"],
                "west_end_ft": int(row["west_end_ft"]),
                "east_end_ft": int(row["east_end_ft"]),
                "estimated_width_in": int(row["width_in"]),
                "master_plan_width_in": "",
                "width_source": "estimated",
                "source": MOYNIHAN_EA_VCE_SOURCE,
                "notes": row["notes"],
            }
            for row in csv.DictReader(f)
            if row["built"] == "yes"
        ]


def main() -> None:
    out, east_end_ft = sheet_vces()
    with MASTER_PLAN_CSV.open() as f:
        master_plan = list(csv.DictReader(f))
    with MASTER_PLAN_EAST_ENDS_CSV.open() as f:
        master_plan_east_ends = {
            int(row["platform"]): int(row["east_end_ft"]) for row in csv.DictReader(f)
        }
    # The directory is calibrated against the sheet's VCEs alone,
    # before the Moynihan EA's join them.
    out += with_pcip_phase_1(directory_vces(out, master_plan), master_plan)
    out += moynihan_ea_vces()
    out.sort(key=lambda v: (v["platform"], v["west_end_ft"]))
    numbers: dict[int, int] = {}
    for v in out:
        numbers[v["platform"]] = numbers.get(v["platform"], 0) + 1
        v["vce_name"] = f"P{v['platform']}-S{numbers[v['platform']]}"
    with EAST_ENDS_CSV.open("w", newline="") as f:
        writer = csv.writer(f, lineterminator="\n")
        writer.writerow(["platform", "east_end_ft", "source"])
        for p, x in sorted(east_end_ft.items()):
            writer.writerow([p, x, SOURCE])
        for p in DIRECTORY_PLATFORMS:
            writer.writerow([p, master_plan_east_ends[p], MASTER_PLAN_EAST_ENDS_SOURCE])
    with OUT_CSV.open("w", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=list(out[0]), lineterminator="\n")
        writer.writeheader()
        writer.writerows(out)
    matched = [v for v in out if v["master_plan_width_in"] and v["estimated_width_in"] != ""]
    errors = [
        int(str(v["estimated_width_in"])) - int(str(v["master_plan_width_in"]).split("/")[0])
        for v in matched
    ]
    print(
        f"wrote {len(out)} VCEs; {len(matched)} match the Master Plan, "
        f"with sheet widths off by {min(errors)} to {max(errors)} in."
    )


DIRECTORY_PLATFORMS = (9, 10, 11)
"""
Platforms the PCIP Phase 2 existing plan doesn't show,
whose VCEs are instead estimated from NJT's January 2022 station directory.
"""

MASTER_PLAN_EAST_ENDS_SOURCE = (
    "NY Penn Station Master Plan Alternatives Report, August 2020 draft, "
    "platform-level plans' platform outlines"
)

DIRECTORY_VCE_SOURCE = "njt_directory"
"""
`data/vces.csv`'s `source` for VCEs from NJT's Penn Station directory (January 2022),
positioned by calibrating its map against platforms 1 to 8.
"""

UNMATCHED_ICON_COST_FT = 60
"""
How far off (ft) a directory icon's position has to be
before it's left unmatched when calibrating, e.g. for a VCE the other source doesn't have.
"""

WRONG_TYPE_COST_FT = 30
"""Extra cost (ft) of matching a directory icon to a VCE of another type when calibrating."""

CALIBRATION_ITERATIONS = 8

INITIAL_CALIBRATIONS = {"lower": (0.28, 10), "upper": (0.225, 90)}
"""
Each directory level's approximate feet per map unit and offset (ft),
from fitting one line to all of its icons, as a starting point for the piecewise calibration.
"""


def align(
    icons: list[tuple[float, str]],
    known: list[tuple[float, str]],
    to_ft: Callable[[float], float],
) -> tuple[float, tuple[tuple[float, float], ...]]:
    """
    The cheapest matching of `icons` (map x, type) to `known` VCEs (ft, type) in order,
    and its cost, where each icon costs how far it is from its match, or else
    `UNMATCHED_ICON_COST_FT` if it's left unmatched.
    """
    icons = sorted(icons)
    known = sorted(known)

    @cache
    def best(i: int, j: int) -> tuple[float, tuple[tuple[float, float], ...]]:
        if i == len(icons):
            return 0.0, ()
        skip_cost, skip_pairs = best(i + 1, j)
        result = (UNMATCHED_ICON_COST_FT + skip_cost, skip_pairs)
        for k in range(j, len(known)):
            cost = abs(to_ft(icons[i][0]) - known[k][0])
            if icons[i][1] != known[k][1]:
                cost += WRONG_TYPE_COST_FT
            rest_cost, rest_pairs = best(i + 1, k + 1)
            if cost + rest_cost < result[0]:
                result = (cost + rest_cost, ((icons[i][0], known[k][0]), *rest_pairs))
        return result

    return best(0, 0)


def linear(scale: float, offset: float) -> Callable[[float], float]:
    """The function `offset + scale * x`."""

    def fit(x: float) -> float:
        return offset + scale * x

    return fit


def monotone_fit(pairs: list[tuple[float, float]]) -> Callable[[float], float]:
    """
    A piecewise-linear, nondecreasing function through `pairs` (x, y),
    after averaging nearby x's and pooling any that would decrease.
    """
    groups: defaultdict[int, list[tuple[float, float]]] = defaultdict(list)
    for x, y in pairs:
        groups[round(x / 40)].append((x, y))
    points = sorted(
        (statistics.mean(x for x, _ in g), statistics.mean(y for _, y in g))
        for g in groups.values()
    )
    # Pool adjacent points whose y's decrease, so the function never does.
    blocks = [[x, y, 1] for x, y in points]
    i = 0
    while i < len(blocks) - 1:
        a, b = blocks[i], blocks[i + 1]
        if a[1] > b[1]:
            n = a[2] + b[2]
            blocks[i] = [(a[0] * a[2] + b[0] * b[2]) / n, (a[1] * a[2] + b[1] * b[2]) / n, n]
            del blocks[i + 1]
            i = max(i - 1, 0)
        else:
            i += 1
    xs = [b[0] for b in blocks]
    ys = [b[1] for b in blocks]

    def fit(x: float) -> float:
        j = min(max(bisect.bisect(xs, x) - 1, 0), len(xs) - 2)
        return ys[j] + (ys[j + 1] - ys[j]) * (x - xs[j]) / (xs[j + 1] - xs[j])

    return fit


def directory_vces(
    estimated: list[dict[str, object]], master_plan: list[dict[str, str]]
) -> list[dict[str, object]]:
    """
    Every VCE on `DIRECTORY_PLATFORMS`, from NJT's station directory
    (`data/vces_njt_directory.csv`), whose map is schematic and not to scale.

    Each level's map is calibrated to feet with a piecewise-linear, nondecreasing fit,
    matching each platform's icons in order to its VCEs in `estimated` on platforms 1 to 8,
    and to the Master Plan's on platforms 9 and 10, and refitting until it settles.
    Icons of the same type within `SAME_VCE_TOLERANCE_FT` on both levels are one VCE.
    Each VCE has the Master Plan's width, if it has one of the same type within
    `SAME_VCE_TOLERANCE_FT`, or else the median width and length of that type on platforms 1 to 8.
    The directory doesn't show every VCE, so the Master Plan's that it doesn't have are added,
    though they're from before Moynihan Train Hall opened.
    """
    icons: defaultdict[tuple[str, int], list[tuple[float, str]]] = defaultdict(list)
    with DIRECTORY_CSV.open() as f:
        for row in csv.DictReader(f):
            if row["type"] != "elevator":
                icons[(row["level"], int(row["platform"]))].append(
                    (float(row["map_x"]), row["type"])
                )
    known: defaultdict[int, list[tuple[float, str]]] = defaultdict(list)
    for v in estimated:
        mid = (float(str(v["west_end_ft"])) + float(str(v["east_end_ft"]))) / 2
        known[int(str(v["platform"]))].append((mid, str(v["type"])))
    for m in master_plan:
        if int(m["platform"]) in DIRECTORY_PLATFORMS:
            known[int(m["platform"])].append((float(m["midpoint_ft"]), m["type"]))

    positions: defaultdict[int, list[tuple[float, str]]] = defaultdict(list)
    for level, (scale, offset) in INITIAL_CALIBRATIONS.items():
        to_ft = linear(scale, offset)
        platforms = [p for lv, p in icons if lv == level and p in known]
        for _ in range(CALIBRATION_ITERATIONS):
            pairs = [
                pair for p in platforms for pair in align(icons[(level, p)], known[p], to_ft)[1]
            ]
            to_ft = monotone_fit(pairs)
        for p in DIRECTORY_PLATFORMS:
            for x, type_ in icons.get((level, p), []):
                ft = to_ft(x)
                if not any(
                    t == type_ and abs(ft - other) <= SAME_VCE_TOLERANCE_FT
                    for other, t in positions[p]
                ):
                    positions[p].append((ft, type_))

    def typical(type_: str, key: str) -> float:
        return statistics.median(
            float(str(v[key]))
            if key == "estimated_width_in"
            else float(str(v["east_end_ft"])) - float(str(v["west_end_ft"]))
            for v in estimated
            # Hidden VCEs' treads are clipped, so they have no width or length of their own.
            if v["type"] == type_ and v["estimated_width_in"] != ""
        )

    # The directory doesn't show every VCE, so add the Master Plan's that it doesn't have.
    for m in master_plan:
        p = int(m["platform"])
        if p in DIRECTORY_PLATFORMS and not any(
            t == m["type"] and abs(float(m["midpoint_ft"]) - mid) <= SAME_VCE_TOLERANCE_FT
            for mid, t in positions[p]
        ):
            positions[p].append((float(m["midpoint_ft"]), m["type"]))

    out: list[dict[str, object]] = []
    for p in DIRECTORY_PLATFORMS:
        for n, (mid, type_) in enumerate(sorted(positions[p]), start=1):
            same = [
                m
                for m in master_plan
                if int(m["platform"]) == p
                and m["type"] == type_
                and abs(float(m["midpoint_ft"]) - mid) <= SAME_VCE_TOLERANCE_FT
            ]
            match = min(same, key=lambda m: abs(float(m["midpoint_ft"]) - mid), default=None)
            length = typical(type_, "length")
            out.append(
                {
                    "platform": p,
                    "vce_name": f"P{p}-S{n}",
                    "type": type_,
                    "west_end_ft": round(mid - length / 2),
                    "east_end_ft": round(mid + length / 2),
                    "estimated_width_in": ""
                    if match
                    else round(typical(type_, "estimated_width_in")),
                    # Where the Master Plan's alternatives disagree, e.g. 66/72, the narrowest.
                    "master_plan_width_in": match["width_in"].split("/")[0] if match else "",
                    "width_source": "master_plan" if match else "typical",
                    "source": DIRECTORY_VCE_SOURCE,
                    "notes": "Its position is only approximate; the directory isn't to scale.",
                }
            )
    return out


WIDEST_MASTER_PLAN_STAIR_IN = 108

WEST_END_CONCOURSE_EAST_FT = 20
"""
VCEs west of this (ft) on platforms 9 to 11 are the West End Concourse's,
which the Moynihan Station EA's plan has, so the PCIP Phase 1 plan's aren't added.
"""


@cache
def pcip_phase_1_found() -> list[tuple[int, str, float, float, float]]:
    """
    Every VCE on the numbered platforms that the PCIP Phase 1 existing plan draws treads for:
    its platform, type, west and east ends (ft), and width (in.),
    the width across its treads at its narrowest flight.
    """
    from platform_crowd_model import shapes, shapes_pcip_phase_1

    plan = shapes_pcip_phase_1.read_plan(
        shapes_pcip_phase_1.PAGE, shapes_pcip_phase_1.SOURCE, {}, fit_y=False
    )
    out: list[tuple[int, str, float, float, float]] = []
    for v in shapes.vces(plan):
        vce = (int(v.platform), v.type, v.box[0], v.box[2], min(f.width_in for f in v.flights))
        # Some VCEs are found twice, as flights drawn twice.
        if v.platform.isdigit() and vce not in out:
            out.append(vce)
    return out


def pcip_phase_1_vces() -> list[tuple[int, str, float, float, float]]:
    """
    Every VCE on `DIRECTORY_PLATFORMS` that the PCIP Phase 1 existing plan draws treads for,
    east of the West End Concourse, as in `pcip_phase_1_found`.
    """
    return [
        v
        for v in pcip_phase_1_found()
        if v[0] in DIRECTORY_PLATFORMS and v[2] >= WEST_END_CONCOURSE_EAST_FT
    ]


def under_label(platform: int, west: int, east: int) -> tuple[int, int] | None:
    """
    The west and east ends (ft) of the VCE on `platform` that the PCIP Phase 1 existing plan draws
    nearest a VCE between `west` and `east` under the PCIP Phase 2 plan's label,
    which clips its treads, if it draws one within `SAME_VCE_TOLERANCE_FT`.
    """
    mid = (west + east) / 2
    near = [v for v in pcip_phase_1_found() if v[0] == platform]
    best = min(near, key=lambda v: abs((v[2] + v[3]) / 2 - mid), default=None)
    if best is None or abs((best[2] + best[3]) / 2 - mid) > SAME_VCE_TOLERANCE_FT:
        return None
    return round(best[2]), round(best[3])


def with_pcip_phase_1(
    directory: list[dict[str, object]], master_plan: list[dict[str, str]]
) -> list[dict[str, object]]:
    """
    `directory`'s VCEs, from NJT's directory, which isn't to scale,
    at the positions the PCIP Phase 1 existing plan draws them, where it does,
    matching each platform's nearest first, within `UNMATCHED_ICON_COST_FT`,
    with the plan's width where the Master Plan has none,
    and the plan's VCEs that the directory doesn't show.
    The directory's VCEs the plan doesn't draw treads for keep the directory's positions.
    """
    measured = pcip_phase_1_vces()
    out: list[dict[str, object]] = []
    for p in DIRECTORY_PLATFORMS:
        mine = [v for v in directory if v["platform"] == p]
        theirs = [v for v in measured if v[0] == p]

        def mid(v: dict[str, object]) -> float:
            return (float(str(v["west_end_ft"])) + float(str(v["east_end_ft"]))) / 2

        # Matching nearest first, not in order, since some are side by side,
        # e.g. a stair beside an escalator, which the two sources may order differently.
        costs = sorted(
            (
                abs(mid(v) - (m[2] + m[3]) / 2) + (WRONG_TYPE_COST_FT if v["type"] != m[1] else 0),
                i,
                j,
            )
            for i, v in enumerate(mine)
            for j, m in enumerate(theirs)
        )
        matches: dict[int, int] = {}
        used: set[int] = set()
        for cost, i, j in costs:
            if cost <= UNMATCHED_ICON_COST_FT and i not in matches and j not in used:
                matches[i] = j
                used.add(j)
        for i, v in enumerate(mine):
            if i not in matches:
                notes = f"{v['notes']} The PCIP Phase 1 plan doesn't draw its treads."
                out.append({**v, "notes": notes})
                continue
            out.append(
                measured_row(theirs[matches[i]], master_plan, "NJT's directory shows it, too.")
            )
        for i, m in enumerate(theirs):
            if i not in used:
                out.append(measured_row(m, master_plan, "NJT's directory doesn't show it."))
    return out


def measured_row(
    vce: tuple[int, str, float, float, float], master_plan: list[dict[str, str]], note: str
) -> dict[str, object]:
    """A VCE measured on the PCIP Phase 1 plan, as a row of `OUT_CSV`."""
    platform, vce_type, west, east, width = vce
    mid = (west + east) / 2
    same = [
        m
        for m in master_plan
        if int(m["platform"]) == platform
        and m["type"] == vce_type
        and abs(float(m["midpoint_ft"]) - mid) <= SAME_VCE_TOLERANCE_FT
    ]
    match = min(same, key=lambda m: abs(float(m["midpoint_ft"]) - mid), default=None)
    return {
        "platform": platform,
        "vce_name": "",
        "type": vce_type,
        "west_end_ft": round(west),
        "east_end_ft": round(east),
        "estimated_width_in": round(width),
        # Where the Master Plan's alternatives disagree, e.g. 66/72, the narrowest.
        "master_plan_width_in": match["width_in"].split("/")[0] if match else "",
        "width_source": "master_plan" if match else "estimated",
        "source": PCIP_PHASE_1_VCE_SOURCE,
        "notes": f"Measured on the PCIP Phase 1 existing plan. {note}"
        + (
            " It's wider than any stair the Master Plan has, so it may be two side by side."
            if width > WIDEST_MASTER_PLAN_STAIR_IN and not match
            else ""
        ),
    }
