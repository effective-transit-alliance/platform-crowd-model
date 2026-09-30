"""
Measure Penn Transformation's new VCEs and platform extensions
on the FRA's Penn Station Service Optimization Study (SOS) Phase I report (June 2026),
whose Figure 11 shows the "generalized locations" of up to 23 new VCEs,
19 stairs and 4 escalators, with at least one on every platform,
and the extensions of Platforms 1 to 3 to the west,
and whose Table 1 has each platform's circulation area before and after decluttering.

Figure 11 is an image, so each VCE's icon is found as a white disk ringed in navy,
and each extension as an orange box.
It's converted to the Master Plan's frame (feet east of its plans' west edge)
by fitting a line to where Platforms 3, 4, 7, and 8's drawn outlines end to the east
and Platform 8's to the west, and where they end in `data/platform_east_ends.csv`
and `data/platform_west_ends_pcip_phase_1.csv`.

The report doesn't say which VCEs are escalators, or how wide any of them are,
so the 4 escalators are taken to be the 4 in Moynihan at the platforms' west ends,
in line with the Train Hall's existing escalators,
and each VCE is taken to be as wide as Moynihan's:
stairs as wide as the West End Concourse's, "nominally 6 feet wide" in the Moynihan Station EA,
and escalators with 1,000 mm steps, like the Train Hall's in `data/vces_moynihan_ea.csv`.
Escalators are about as long as the Train Hall's, 50 ft,
and stairs as long as the existing stairs' median in `data/vces.csv`.
Together they add about 30% to the existing VCEs' total width,
close to the 32% more vertical circulation capacity Penn Transformation's designers report.

Writes `data/vces_transformation_fra_sos.csv` and `data/platforms_transformation_fra_sos.csv`.
"""

import csv
import re
import statistics
import urllib.request
from collections.abc import Callable

import numpy as np
import pymupdf

from platform_crowd_model.paths import CACHE_DIR, DATA_DIR

PDF_CACHE = CACHE_DIR / "fra-sos-phase-1.pdf"
PDF_URL = (
    "https://liamblank.com/wp-content/uploads/2026/07/"
    "2026.07.13_Penn-Station-SOS_Phase-I-Report_FINAL-1.pdf"
)
"""
Liam Blank's copy of the report, since the FRA's
(https://railroads.dot.gov/sites/fra.dot.gov/files/2026-07/2026.07.13_Penn%20Station%20SOS_Phase%20I%20Report_FINAL.pdf)
refuses scripted downloads.
"""
FIGURE_PAGE = 71
"""1-indexed PDF page of Figure 11, "Penn Transformation Improvements - Proposed VCE Additions"."""
TABLE_PAGE = 50
"""1-indexed PDF page of Table 1, "Results of Platform Decluttering Analysis"."""
SOURCE = f"FRA Penn Station SOS Phase I Report, June 2026, Figure 11 (PDF page {FIGURE_PAGE})"

EAST_ENDS_CSV = DATA_DIR / "platform_east_ends.csv"
WEST_ENDS_CSV = DATA_DIR / "platform_west_ends_pcip_phase_1.csv"
EXISTING_VCES_CSV = DATA_DIR / "vces.csv"
VCES_CSV = DATA_DIR / "vces_transformation_fra_sos.csv"
PLATFORMS_CSV = DATA_DIR / "platforms_transformation_fra_sos.csv"

PLATFORM_ROWS = {
    11: 228,
    10: 293,
    9: 337,
    8: 391,
    7: 438,
    6: 490,
    5: 540,
    4: 591,
    3: 645,
    2: 693,
    1: 746,
}
"""Each platform's row (px) on Figure 11's image."""

EAST_END_PLATFORMS = (3, 4, 7, 8)
"""Platforms whose drawn east ends are on the figure, not cut off by its edge."""

WEST_END_PLATFORMS = (8,)
"""Platforms whose drawn west ends are clear of other lines on the figure."""

EXTENDED_PLATFORMS = (1, 2, 3)

EXTENDED_MAX_CARS = 10
"""
Cars in the longest train on an extended platform:
"at least 10-car trains (not including the locomotive)",
though Track 5, on Platform 3, "will be capped at nine passenger cars".
"""

ESCALATOR_WEST_OF_FT = -200
"""New VCEs west of here (ft) are taken to be the 4 escalators, in Moynihan."""

STAIR_WIDTH_IN = 72
"""A new stair's width: the West End Concourse's stairs', "nominally 6 feet wide"."""

ESCALATOR_WIDTH_IN = 40
"""A new escalator's step width: the Train Hall's, KONE's 1,000 mm steps."""

ESCALATOR_LENGTH_FT = 50
"""A new escalator's length along the platform, about the Train Hall's."""


def pdf() -> pymupdf.Document:
    if not PDF_CACHE.exists():
        CACHE_DIR.mkdir(exist_ok=True)
        request = urllib.request.Request(
            PDF_URL, headers={"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) Firefox/130.0"}
        )
        with urllib.request.urlopen(request) as response:
            PDF_CACHE.write_bytes(response.read())
    return pymupdf.open(PDF_CACHE)


def figure(doc: pymupdf.Document) -> np.ndarray:
    """Figure 11's image, as rows of RGB pixels: the page's largest image but its background."""
    page = doc[FIGURE_PAGE - 1]
    info = min(page.get_image_info(xrefs=True), key=lambda i: i["width"] * i["height"])
    pix = pymupdf.Pixmap(doc, info["xref"])
    if pix.n != 3:
        pix = pymupdf.Pixmap(pymupdf.csRGB, pix)
    return np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, 3).astype(int)


def icons(rgb: np.ndarray) -> list[tuple[float, float]]:
    """Each VCE icon's center (px): a white disk with a navy ring 10 to 21 px from its center."""
    r, g, b = rgb[..., 0], rgb[..., 1], rgb[..., 2]
    navy = (b > 70) & (b - r > 35) & (r < 80) & (g < 100)
    white = rgb.min(axis=2) > 235
    height, width = navy.shape

    def shifted(dx: int, dy: int) -> np.ndarray:
        out = np.zeros_like(navy)
        ys = slice(max(0, -dy), height - max(0, dy))
        xs = slice(max(0, -dx), width - max(0, dx))
        out[ys, xs] = navy[max(0, dy) : height - max(0, -dy), max(0, dx) : width - max(0, -dx)]
        return out

    ringed = white.copy()
    for dx, dy in ((1, 0), (-1, 0), (0, 1), (0, -1)):
        near = np.zeros_like(navy)
        for k in range(10, 22):
            near |= shifted(dx * k, dy * k)
        ringed &= near
    clusters: list[list[tuple[int, int]]] = []
    for y, x in zip(*np.nonzero(ringed), strict=True):
        for cluster in clusters:
            cy, cx = cluster[0]
            if abs(cy - y) < 25 and abs(cx - x) < 25:
                cluster.append((y, x))
                break
        else:
            clusters.append([(y, x)])
    # A line through an icon can split its disk, so merge clusters whose centers are close.
    centers: list[list[tuple[int, int]]] = []
    for cluster in clusters:
        cy, cx = statistics.mean(y for y, _x in cluster), statistics.mean(x for _y, x in cluster)
        for merged in centers:
            my, mx = statistics.mean(y for y, _x in merged), statistics.mean(x for _y, x in merged)
            if abs(my - cy) < 25 and abs(mx - cx) < 25:
                merged.extend(cluster)
                break
        else:
            centers.append(list(cluster))
    return [
        (statistics.mean(x for _y, x in c), statistics.mean(y for y, _x in c))
        for c in centers
        if len(c) > 10
    ]


def platform_ends(rgb: np.ndarray) -> dict[int, tuple[int, int]]:
    """Each platform's drawn outline's westmost and eastmost dark pixel (px) near its row."""
    dark = rgb.max(axis=2) < 70
    out = {}
    for platform, row in PLATFORM_ROWS.items():
        columns = np.nonzero(dark[row - 6 : row + 7].any(axis=0))[0]
        out[platform] = (int(columns.min()), int(columns.max()))
    return out


def extensions(rgb: np.ndarray) -> dict[int, tuple[int, int]]:
    """Each extended platform's orange box's west and east edges (px)."""
    r, g, b = rgb[..., 0], rgb[..., 1], rgb[..., 2]
    orange = (r > 230) & (g > 190) & (g < 235) & (b < 170)
    out = {}
    for platform in EXTENDED_PLATFORMS:
        row = PLATFORM_ROWS[platform]
        columns = np.nonzero(orange[row - 12 : row + 13].any(axis=0))[0]
        out[platform] = (int(columns.min()), int(columns.max()))
    return out


def read_ends(path: str, column: str) -> dict[int, float]:
    with (DATA_DIR / path).open() as f:
        return {int(row["platform"]): float(row[column]) for row in csv.DictReader(f)}


def calibration(rgb: np.ndarray) -> Callable[[float], float]:
    """A function from the figure's columns (px) to feet in the Master Plan's frame."""
    ends = platform_ends(rgb)
    east = read_ends(EAST_ENDS_CSV.name, "east_end_ft")
    west = read_ends(WEST_ENDS_CSV.name, "west_end_ft")
    px = [ends[p][1] for p in EAST_END_PLATFORMS] + [ends[p][0] for p in WEST_END_PLATFORMS]
    ft = [east[p] for p in EAST_END_PLATFORMS] + [west[p] for p in WEST_END_PLATFORMS]
    ft_per_px, offset = np.polyfit(px, ft, 1)
    residuals = [f - (ft_per_px * x + offset) for x, f in zip(px, ft, strict=True)]
    print(f"{ft_per_px:.3f} ft per px, ends off by up to {max(map(abs, residuals)):.1f} ft")
    return lambda x: float(ft_per_px * x + offset)


def decluttering(doc: pymupdf.Document) -> dict[int, tuple[int, int]]:
    """Each platform's current and additional circulation area (sq ft), from Table 1."""
    text = str(doc[TABLE_PAGE - 1].get_text())
    table = text[text.index("Table 1: Results of Platform Decluttering") :]
    numbers = [int(n.replace(",", "")) for n in re.findall(r"^\s*([\d,]+)\s*$", table, re.M)]
    platforms, current, added = numbers[:11], numbers[11:22], numbers[22:33]
    if platforms != list(range(1, 12)):
        raise RuntimeError(f"Table 1's platforms are {platforms}")
    return {p: (c, a) for p, c, a in zip(platforms, current, added, strict=True)}


def median_stair_length_ft() -> float:
    with EXISTING_VCES_CSV.open() as f:
        return statistics.median(
            int(row["east_end_ft"]) - int(row["west_end_ft"])
            for row in csv.DictReader(f)
            if row["type"] == "stair"
        )


def main() -> None:
    doc = pdf()
    rgb = figure(doc)
    ft = calibration(rgb)
    stair_length = median_stair_length_ft()
    vces = []
    for x, y in icons(rgb):
        platform = min(PLATFORM_ROWS, key=lambda p: abs(PLATFORM_ROWS[p] - y))
        if abs(PLATFORM_ROWS[platform] - y) > 10:
            continue  # the legend's icon
        mid = ft(x)
        escalator = mid < ESCALATOR_WEST_OF_FT
        length = ESCALATOR_LENGTH_FT if escalator else stair_length
        vces.append(
            {
                "platform": platform,
                "type": "escalator" if escalator else "stair",
                "west_end_ft": round(mid - length / 2),
                "east_end_ft": round(mid + length / 2),
                "width_in": ESCALATOR_WIDTH_IN if escalator else STAIR_WIDTH_IN,
                "source": SOURCE,
            }
        )
    vces.sort(key=lambda v: (v["platform"], v["west_end_ft"]))
    types = [v["type"] for v in vces]
    print(f"{types.count('stair')} stairs and {types.count('escalator')} escalators")
    if (types.count("stair"), types.count("escalator")) != (19, 4):
        raise RuntimeError("expected 19 stairs and 4 escalators")
    with VCES_CSV.open("w", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=list(vces[0]), lineterminator="\n")
        writer.writeheader()
        writer.writerows(vces)

    boxes = extensions(rgb)
    areas = decluttering(doc)
    platforms = []
    for platform, (current, added) in sorted(areas.items()):
        box = boxes.get(platform)
        platforms.append(
            {
                "platform": platform,
                "west_end_ft": round(ft(box[0])) if box else "",
                "max_cars": EXTENDED_MAX_CARS if box else "",
                "circulation_area_sq_ft": current,
                "added_circulation_area_sq_ft": added,
                "source": SOURCE.replace("Figure 11", "Figure 11 and Table 1"),
            }
        )
    with PLATFORMS_CSV.open("w", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=list(platforms[0]), lineterminator="\n")
        writer.writeheader()
        writer.writerows(platforms)
    for row in vces + platforms:
        print(row)
