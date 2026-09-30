"""
Measure where each platform ends to the west
on NJT's PCIP Phase 1 existing track plan (Appendix A, sheet TK-003, July 2019),
which, unlike the Master Plan's and PCIP Phase 2's plans, doesn't cut the platforms off.

The sheet is a scan, so it's rendered at `DPI`,
and its orange existing platform edges are found by color.
Each platform's extent is where those edges are within its band of rows, `PLATFORM_ROWS`.
The sheet is converted to the Master Plan's frame (feet east of its plans' west edge)
by fitting a line to the platforms' east ends on it and in `data/platform_east_ends.csv`,
which comes out at 80.3 ft per inch of the sheet, matching its 1" = 80' scale.

Writes `data/platform_west_ends_pcip_phase_1.csv`.
"""

import csv
import urllib.request

import numpy as np
import pymupdf

from platform_crowd_model.paths import CACHE_DIR, DATA_DIR

PDF_CACHE = CACHE_DIR / "pcip-1-conceptual-design-preliminary-drawings.pdf"
PDF_URL = (
    "https://liamblank.com/wp-content/uploads/2026/07/"
    "penn-records-s-nj-transit-pcip-1-pcip1-final-report-c5015-01-262652-00-task-06-mem"
    "-final-report-draft-appendixa-drawings-copy.pdf"
)
PAGE = 8
"""1-indexed PDF page of sheet TK-003, "Existing Track Alignment, East Part Plan"."""
SOURCE = f"PCIP Phase 1 Appendix A, sheet TK-003, July 2019, PDF page {PAGE}"

EAST_ENDS_CSV = DATA_DIR / "platform_east_ends.csv"
OUT_CSV = DATA_DIR / "platform_west_ends_pcip_phase_1.csv"

DPI = 200

PLATFORM_ROWS = {
    11: (755, 816),
    10: (870, 985),
    9: (1000, 1058),
    8: (1112, 1178),
    7: (1226, 1316),
    6: (1340, 1400),
    5: (1440, 1516),
    4: (1564, 1628),
    3: (1680, 1744),
    2: (1796, 1910),
    1: (1906, 1972),
}
"""Each platform's band of rows (px at `DPI`) on the sheet, between its edges."""

DIAGONAL_PLATFORM_EAST_EDGES = {1: 3400, 2: 3400, 3: 3230}
"""
Columns (px) west of which platforms 1 to 3's rows are the Diagonal Platform's, which overlaps them,
though none of these platforms reach that far west.
"""

FIT_PLATFORMS = range(4, 12)
"""Platforms whose east ends calibrate the sheet, leaving out platforms 1 to 3's as a check."""

MAX_RESIDUAL_FT = 5
"""How far off a fitted east end can be before the measurement is considered broken."""


def pdf() -> pymupdf.Document:
    if not PDF_CACHE.exists():
        CACHE_DIR.mkdir(exist_ok=True)
        request = urllib.request.Request(
            PDF_URL, headers={"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) Firefox/130.0"}
        )
        with urllib.request.urlopen(request) as response:
            PDF_CACHE.write_bytes(response.read())
    return pymupdf.open(PDF_CACHE)


def platform_extents() -> dict[int, tuple[int, int]]:
    """Each platform's westmost and eastmost orange edge pixel's column on the sheet."""
    pix = pdf()[PAGE - 1].get_pixmap(dpi=DPI)
    rgb = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
    r, g, b = (rgb[..., i].astype(int) for i in range(3))
    orange = (r - b > 60) & (r > 180) & (g > 100) & (g < 220)
    extents = {}
    for platform, (top, bottom) in PLATFORM_ROWS.items():
        left = DIAGONAL_PLATFORM_EAST_EDGES.get(platform, 0)
        columns = np.nonzero(orange[top:bottom, left:].any(axis=0))[0] + left
        extents[platform] = (int(columns.min()), int(columns.max()))
    return extents


def main() -> None:
    extents = platform_extents()
    with EAST_ENDS_CSV.open() as f:
        east_ends = {int(row["platform"]): float(row["east_end_ft"]) for row in csv.DictReader(f)}
    ft_per_px, offset = np.polyfit(
        [extents[p][1] for p in FIT_PLATFORMS], [east_ends[p] for p in FIT_PLATFORMS], 1
    )
    print(f"{ft_per_px * DPI:.1f} ft per inch of the sheet")
    rows = []
    for platform in sorted(extents):
        west, east = (ft_per_px * px + offset for px in extents[platform])
        residual = east - east_ends[platform]
        print(
            f"platform {platform}: {west:.0f} to {east:.0f} ft, east end off by {residual:+.1f} ft"
        )
        if abs(residual) > MAX_RESIDUAL_FT:
            raise RuntimeError(f"platform {platform}'s east end is off by {residual:+.1f} ft")
        rows.append({"platform": platform, "west_end_ft": round(west), "source": SOURCE})
    with OUT_CSV.open("w", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=list(rows[0]), lineterminator="\n")
        writer.writeheader()
        writer.writerows(rows)


if __name__ == "__main__":
    main()
