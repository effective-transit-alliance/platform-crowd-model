"""
Measure the VCEs at the west ends of the platforms, around Moynihan Train Hall,
on the Moynihan Station Development Project EA's lower concourse plan (Figure 3-4, February 2010),
which the Master Plan's and PCIP Phase 2's plans cut off or leave out:
the pairs of escalators from the Train Hall down to Platforms 3 to 8,
and the West End Concourse's stairs down to Platforms 9 to 11.

The plan is a vector drawing, so each VCE is found as a run of treads,
short lines across the platform, within its window, `VCES`.
Its scale comes from the West End Concourse's width, dimensioned as 36'-3",
and it's registered to the Master Plan's frame (feet east of its plans' west edge)
by the West End Concourse's stairs down to Platforms 3 to 8,
as `vces` measures them on PCIP Phase 2's existing plan.

The plan is a design, from before the Train Hall was built,
but the Train Hall opened in 2021 with escalators to Platforms 3 to 8 (Tracks 5 to 16),
11 of them, one fewer than the plan's 12.
Platform 3 doesn't reach as far west as the others,
and its western escalator would run past its west end in `data/platform_west_ends_pcip_phase_1.csv`,
so it's taken as the one that wasn't built.
The escalators are taken to be 40 in. wide, KONE's 1,000 mm steps,
since the plan doesn't draw them consistently.

Writes `data/vces_moynihan_ea.csv`.
"""

import csv
import statistics
from dataclasses import dataclass
from urllib.request import Request, urlopen

import pymupdf

from platform_crowd_model.paths import CACHE_DIR, DATA_DIR
from platform_crowd_model.vces import sheet_vces

PDF_CACHE = CACHE_DIR / "moynihan-ea-figures-3-3-and-3-4.pdf"
PDF_URL = (
    "https://web.archive.org/web/2017id_/https://cdn.esd.ny.gov/subsidiaries_projects/msdc/"
    "Data/NEPA/03a%20Figure%203-3%20and%203-4.pdf"
)
PAGE = 2
"""1-indexed PDF page of Figure 3-4, "Moynihan Station - Lower Concourse"."""
SOURCE = f"Moynihan Station EA, Figure 3-4, February 2010, PDF page {PAGE}"

OUT_CSV = DATA_DIR / "vces_moynihan_ea.csv"

CONCOURSE_FILL = (0.976, 0.943, 0.765)
"""The West End Concourse's fill color."""

CONCOURSE_WIDTH_FT = 36 + 3 / 12
"""The West End Concourse's width where it's narrowest, as dimensioned on the plan."""

CONCOURSE_WEST_WALL = 611.6
"""The West End Concourse's west wall (PDF units), where its fill starts."""

CORRIDOR_WEST_WALL = 442.8
"""
The baggage and egress corridor's west wall (PDF units), where its fill starts,
whose width, dimensioned as 19', checks the scale.
"""

CORRIDOR_WIDTH_FT = 19

ESCALATOR_WIDTH_IN = 40
"""
Each Train Hall escalator's step width: KONE's 1,000 mm steps,
which Elevator World reports for its escalators at Moynihan.
"""

MAX_TREAD_LENGTH_PT = 6
"""Longest tread line (PDF units), to leave out longer lines like platform edges."""


@dataclass(frozen=True)
class VceWindow:
    platform: int
    type: str
    y: float
    """A row (PDF units) across the VCE's treads."""
    x0: float
    x1: float
    """The columns (PDF units) the VCE's treads are within."""
    built: bool = True
    notes: str = ""


ESCALATOR_NOTE = "Its width is KONE's 1,000 mm steps, not measured on the plan."

UNBUILT_ESCALATOR_NOTE = (
    "Taken as not built: the Train Hall opened with 11 escalators, not the plan's 12, "
    "and this one would run past the platform's west end."
)

TRAIN_HALL_ESCALATORS = [
    VceWindow(
        platform,
        "escalator",
        y,
        x0,
        x1,
        built=not unbuilt,
        notes=UNBUILT_ESCALATOR_NOTE if unbuilt else ESCALATOR_NOTE,
    )
    for platform, y in {8: 189.2, 7: 213.3, 6: 238.0, 5: 262.8, 4: 286.9, 3: 306.2}.items()
    for x0, x1 in ((533.9, 564.8), (564.8, 595.1))
    for unbuilt in [platform == 3 and x0 < 564.8]
]
"""The pairs of escalators down from the Train Hall, west then east."""

VCES = [
    *TRAIN_HALL_ESCALATORS,
    VceWindow(9, "stair", 165.4, 635.5, 649.8),
    VceWindow(9, "stair", 168.1, 592.5, 600.5),
    VceWindow(10, "stair", 146.7, 635.5, 649.8),
    VceWindow(11, "stair", 117.7, 644.0, 652.0),
]
"""Where each VCE is on the plan."""

CONCOURSE_STAIRS = [
    VceWindow(platform, "stair", y, 630.2, 659.0)
    for platform, y in {8: 189.2, 7: 213.3, 6: 238.0, 5: 262.8, 4: 286.9, 3: 306.2}.items()
]
"""The West End Concourse's stairs down to Platforms 3 to 8, which register the plan."""


def pdf() -> pymupdf.Document:
    if not PDF_CACHE.exists():
        CACHE_DIR.mkdir(exist_ok=True)
        request = Request(
            PDF_URL, headers={"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) Firefox/130.0"}
        )
        with urlopen(request) as response:
            PDF_CACHE.write_bytes(response.read())
    return pymupdf.open(PDF_CACHE)


def treads(page: pymupdf.Page) -> list[tuple[float, float, float]]:
    """Every short line across the tracks: its column, and its top and bottom rows."""
    out = []
    for drawing in page.get_drawings():
        for item in drawing["items"]:
            if item[0] != "l":
                continue
            a, b = item[1], item[2]
            if abs(a.x - b.x) < 0.05 and 0.3 < abs(a.y - b.y) < MAX_TREAD_LENGTH_PT:
                out.append((a.x, min(a.y, b.y), max(a.y, b.y)))
    return out


def measure(
    window: VceWindow, lines: list[tuple[float, float, float]]
) -> tuple[float, float, float]:
    """
    A VCE's west and east ends and its width (PDF units):
    the extent of its treads, which may be in more than one flight, and their median length,
    since the longest lines can be landings or balustrades.
    """
    found = [
        (x, y1 - y0)
        for x, y0, y1 in lines
        if window.x0 < x < window.x1 and y0 - 1 < window.y < y1 + 1
    ]
    xs = [x for x, _length in found]
    return min(xs), max(xs), statistics.median(length for _x, length in found)


def fill_width(page: pymupdf.Page, west_wall: float) -> float:
    """The narrowest width (PDF units) of the concourse fill starting at `west_wall`."""
    return min(
        drawing["rect"].width
        for drawing in page.get_drawings()
        if drawing.get("fill")
        and all(abs(c - f) < 0.002 for c, f in zip(drawing["fill"], CONCOURSE_FILL, strict=True))
        and drawing["rect"].height > 20
        and abs(drawing["rect"].x0 - west_wall) < 0.5
    )


def registration(page: pymupdf.Page) -> tuple[float, float]:
    """
    The plan's feet per PDF unit, from the West End Concourse's width,
    and the Master Plan frame's x at the plan's x = 0,
    from the West End Concourse's stairs as PCIP Phase 2's plan has them.
    """
    scale = CONCOURSE_WIDTH_FT / fill_width(page, CONCOURSE_WEST_WALL)
    corridor_ft = fill_width(page, CORRIDOR_WEST_WALL) * scale
    print(
        f"{scale:.4f} ft per PDF unit; the {CORRIDOR_WIDTH_FT} ft corridor is {corridor_ft:.2f} ft"
    )
    if abs(corridor_ft - CORRIDOR_WIDTH_FT) > 0.5:
        raise RuntimeError(f"the {CORRIDOR_WIDTH_FT} ft corridor measures {corridor_ft:.2f} ft")
    lines = treads(page)
    sheet, _east_ends = sheet_vces()
    offsets = []
    for window in CONCOURSE_STAIRS:
        x0, x1, _width = measure(window, lines)
        sheet_stair = min(
            (v for v in sheet if v["platform"] == window.platform and v["type"] == "stair"),
            key=lambda v: abs(v["west_end_ft"] + v["east_end_ft"]),
        )
        offsets.append(
            (sheet_stair["west_end_ft"] + sheet_stair["east_end_ft"]) / 2 - (x0 + x1) / 2 * scale
        )
    print(f"registered to PCIP Phase 2's stairs to within {max(offsets) - min(offsets):.1f} ft")
    return scale, statistics.mean(offsets)


def main() -> None:
    page = pdf()[PAGE - 1]
    scale, offset = registration(page)
    lines = treads(page)
    rows = []
    for window in VCES:
        x0, x1, width = measure(window, lines)
        rows.append(
            {
                "platform": window.platform,
                "type": window.type,
                "west_end_ft": round(x0 * scale + offset),
                "east_end_ft": round(x1 * scale + offset),
                "width_in": (
                    ESCALATOR_WIDTH_IN if window.type == "escalator" else round(width * scale * 12)
                ),
                "built": "yes" if window.built else "no",
                "source": SOURCE,
                "notes": window.notes,
            }
        )
    with OUT_CSV.open("w", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=list(rows[0]), lineterminator="\n")
        writer.writeheader()
        writer.writerows(rows)
    for row in rows:
        print(row)
