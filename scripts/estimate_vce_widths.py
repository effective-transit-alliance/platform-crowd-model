#!/usr/bin/env -S uv run --script
# /// script
# requires-python = ">=3.14"
# dependencies = ["pymupdf"]
# ///

"""
Estimate the widths of VCEs that no source lists widths for,
by measuring their stair treads on a scaled vector drawing.

The only such VCEs so far are the 2 stairs from the West End Concourse down to platform 3,
which the Master Plan's width tables don't include.
NJ Transit's PCIP Phase 2 drawings (November 2020) include an existing concourse-level plan
(sheet A-001, PDF page 45) drawn at 1" = 40' with each stair's treads as lines,
so a tread's length is roughly the stair's width.
On the same sheet, the treads of the stairs that seem to match the Master Plan's
measure within about 6 in.
of the Master Plan's widths (e.g. 66 vs. 69 in., 44 vs. 44 in.),
so these widths are estimates, not measurements.

Writes `data/estimated_vce_widths.csv`.
"""

import csv
import urllib.request
from dataclasses import dataclass
from pathlib import Path

import pymupdf

REPO = Path(__file__).resolve().parent.parent
OUT_CSV = REPO / "data" / "estimated_vce_widths.csv"
PDF_CACHE = REPO / ".cache" / "pcip-2-conceptual-design-preliminary-drawings.pdf"
PDF_URL = "https://liamblank.com/wp-content/uploads/2026/09/pcip-2-conceptual-design-preliminary-drawings.pdf"
PAGE = 45
"""1-indexed PDF page of sheet A-001, "Existing Plan Overall"."""

PDF_UNITS_PER_FOOT = (507.2 - 290.9) / 120
"""From the centers of the scale bar's 0' and 120' labels."""

PLATFORM_3_EAST_END_X = 1748
"""
x coordinate of platform 3's east end on the sheet, where its outline ends.
West is to the left.
"""

MIN_TREAD_LENGTH = 4
"""Shorter lines (PDF units) aren't treads."""


@dataclass
class Flight:
    """A run of a stair's treads, found within a box on the sheet."""

    name: str
    x0: float
    y0: float
    x1: float
    y1: float
    along_platform: bool
    """Whether the flight runs along the platform (east-west), so its treads are vertical lines."""


@dataclass
class Stair:
    platform: int
    name: str
    flights: list[Flight]
    notes: str


STAIRS = [
    Stair(
        platform=3,
        name="West End Concourse, west side",
        flights=[Flight("both flights", 295, 510, 345, 533, along_platform=True)],
        notes=(
            "A straight stair with a mid-landing, running along the platform. "
            "Its treads are drawn in two halves split at the platform's edge line, "
            "so its width is their sum."
        ),
    ),
    Stair(
        platform=3,
        name="West End Concourse, east side",
        flights=[Flight("upper flight", 438, 522, 461, 538, along_platform=False)],
        notes=(
            "A T-shaped stair: one flight down from the concourse to a landing, "
            "then two 60 in. flights east and west along the platform. "
            "Its width is the upper flight's, whose treads are drawn in two halves, "
            "which matches the two lower flights' combined width."
        ),
    ),
]


def pdf() -> pymupdf.Document:
    if not PDF_CACHE.exists():
        PDF_CACHE.parent.mkdir(exist_ok=True)
        request = urllib.request.Request(
            PDF_URL, headers={"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) Firefox/130.0"}
        )
        with urllib.request.urlopen(request) as response:
            PDF_CACHE.write_bytes(response.read())
    return pymupdf.open(PDF_CACHE)


def treads(page: pymupdf.Page, flight: Flight) -> list[tuple[pymupdf.Point, pymupdf.Point]]:
    """The tread lines within the flight's box, perpendicular to its direction."""
    box = pymupdf.Rect(flight.x0, flight.y0, flight.x1, flight.y1)
    out = []
    for d in page.get_drawings():
        for item in d["items"]:
            if item[0] != "l":
                continue
            a, b = item[1], item[2]
            if a not in box or b not in box:
                continue
            vertical = abs(a.x - b.x) < 0.3 and abs(a.y - b.y) >= MIN_TREAD_LENGTH
            horizontal = abs(a.y - b.y) < 0.3 and abs(a.x - b.x) >= MIN_TREAD_LENGTH
            if vertical if flight.along_platform else horizontal:
                out.append((a, b))
    return out


def width_in(lines: list[tuple[pymupdf.Point, pymupdf.Point]], along_platform: bool) -> float:
    """
    The flight's width: the extent of its treads across the flight,
    which includes treads drawn in several pieces.
    """
    across = [p.y if along_platform else p.x for line in lines for p in line]
    return (max(across) - min(across)) / PDF_UNITS_PER_FOOT * 12


def main() -> None:
    page = pdf()[PAGE - 1]
    out = []
    for stair in STAIRS:
        widths = []
        east_end = []
        west_end = []
        for flight in stair.flights:
            lines = treads(page, flight)
            if len(lines) < 5:
                raise SystemExit(f"{stair.name}, {flight.name}: found only {len(lines)} treads")
            widths.append(width_in(lines, flight.along_platform))
            xs = [p.x for line in lines for p in line]
            east_end.append((PLATFORM_3_EAST_END_X - max(xs)) / PDF_UNITS_PER_FOOT)
            west_end.append((PLATFORM_3_EAST_END_X - min(xs)) / PDF_UNITS_PER_FOOT)
        out.append(
            {
                "platform": stair.platform,
                "vce": stair.name,
                "type": "stair",
                "width_in": round(min(widths)),
                "width_status": "estimated",
                "east_end_ft": round(min(east_end)),
                "west_end_ft": round(max(west_end)),
                "source": f"PCIP Phase 2 Appendix A, sheet A-001, November 2020, PDF page {PAGE}",
                "notes": stair.notes,
            }
        )
    with OUT_CSV.open("w", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=list(out[0]), lineterminator="\n")
        writer.writeheader()
        writer.writerows(out)
    for v in out:
        print(f"{v['vce']}: {v['width_in']} in., {v['east_end_ft']} to {v['west_end_ft']} ft")


if __name__ == "__main__":
    main()
