"""
Extract the stairs, escalators, and elevators to each platform
from NJT's January 2022 Penn Station directory,
the only source found that shows every platform's VCEs after Moynihan Train Hall opened.

The directory is a vector wayfinding map of the station's upper and lower concourse levels.
Each VCE to a platform is a circular icon beside a small box labeled with the platform's tracks.
Stair icons are a zigzag, escalator icons are a person on a curved escalator,
and elevator icons are filled arrows around a car.
Icons without a track label, e.g. those to another concourse level, are skipped.

The map is schematic and not to scale, so it has no widths, and positions are only approximate,
and it may show a VCE more than once, e.g. once on each level it passes.

Writes `data/vces_njt_directory.csv`.
"""

import csv
import urllib.request
from collections import Counter
from dataclasses import dataclass
from typing import Any

import pymupdf

from platform_crowd_model.paths import CACHE_DIR, DATA_DIR

OUT_CSV = DATA_DIR / "vces_njt_directory.csv"
PDF_CACHE = CACHE_DIR / "NY-Penn-Station-Directory_011022.pdf"
PDF_URL = (
    "https://content.njtransit.com/sites/default/files/NY%20Penn%20Station%20Directory_011022.pdf"
)

TRACKS_TO_PLATFORM = {
    "1/2": 1,
    "3/4": 2,
    "5/6": 3,
    "7/8": 4,
    "9/10": 5,
    "11/12": 6,
    "13/14": 7,
    "15/16": 8,
    "17": 9,
    "18/19": 10,
    "20/21": 11,
}
"""The map labels VCEs by the tracks of the platform they serve."""

LOWER_LEVEL_MIN_Y = 1450
"""The lower level's map is below this y coordinate, and the upper level's above it."""

ICON_SIZE = (22, 30)
"""Range of the width and height (PDF units) of an icon's circular outline."""

MAX_LABEL_GAP = 12
"""Maximum horizontal gap (PDF units) between an icon and its track label."""

MAX_LABEL_OFFSET_Y = 16
"""Maximum vertical offset (PDF units) between the centers of an icon and a label beside it."""

MAX_LABEL_GAP_ABOVE_OR_BELOW = 8
"""
Maximum vertical gap (PDF units) between an icon and its track label above or below it,
as in the NJT concourse and the LIRR's east end.
"""


@dataclass
class Icon:
    rect: pymupdf.Rect
    type: str


def pdf() -> pymupdf.Document:
    if not PDF_CACHE.exists():
        CACHE_DIR.mkdir(exist_ok=True)
        request = urllib.request.Request(
            PDF_URL, headers={"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) Firefox/130.0"}
        )
        with urllib.request.urlopen(request) as response:
            PDF_CACHE.write_bytes(response.read())
    return pymupdf.open(PDF_CACHE)


def is_dark(color: tuple[float, ...] | None) -> bool:
    return color is not None and max(color) < 0.2


def icon_type(inner: list[dict[str, Any]]) -> str:
    """Classify an icon by the paths inside its circular outline."""
    kinds = ["".join(item[0] for item in d["items"]) for d in inner]
    # A zigzag of straight lines, stroked, or filled with an arrow for a stair to another level.
    if any(set(k) == {"l"} and len(k) >= 7 for k in kinds):
        return "stair"
    if any(k.startswith("clclclcl") for k in kinds):
        return "elevator"
    # A person's head on top of a curved escalator.
    if any(k == "lccclccc" for k in kinds):
        return "escalator"
    return "other"


def icons(drawings: list[dict[str, Any]]) -> list[Icon]:
    rings = [
        d["rect"]
        for d in drawings
        if is_dark(d.get("fill"))
        and "".join(item[0] for item in d["items"]) == "cccccccc"
        and ICON_SIZE[0] <= d["rect"].width <= ICON_SIZE[1]
        and ICON_SIZE[0] <= d["rect"].height <= ICON_SIZE[1]
    ]
    out = []
    for ring in rings:
        inner = [
            d
            for d in drawings
            if d["rect"] in ring and d["rect"] != ring and d["rect"].width < ring.width - 4
        ]
        out.append(Icon(ring, icon_type(inner)))
    return out


def labels(page: pymupdf.Page, drawings: list[dict[str, Any]]) -> list[tuple[pymupdf.Rect, str]]:
    """The white, outlined boxes whose text is a platform's tracks."""
    out = []
    for d in drawings:
        r = d["rect"]
        if (
            d.get("fill") == (1.0, 1.0, 1.0)
            and is_dark(d.get("color"))
            and 8 < r.height < 18
            and r.width < 60
        ):
            # Only the words centered in the box, since `clip` also returns overlapping neighbors.
            words = [(pymupdf.Rect(w[:4]), str(w[4])) for w in page.get_text("words", clip=r)]
            text = " ".join(word for box, word in words if (box.tl + box.br) / 2 in r)
            if text in TRACKS_TO_PLATFORM:
                out.append((r, text))
    return out


def gap(icon: Icon, label: pymupdf.Rect) -> float:
    """The icon's gap from the label, if it's beside, above, or below it, or else infinity."""
    i = icon.rect
    beside = min(abs(i.x1 - label.x0), abs(label.x1 - i.x0))
    offset_y = abs((i.y0 + i.y1) / 2 - (label.y0 + label.y1) / 2)
    if offset_y <= MAX_LABEL_OFFSET_Y and beside <= MAX_LABEL_GAP:
        return beside
    if i.x0 <= (label.x0 + label.x1) / 2 <= i.x1:
        vertical = min(abs(label.y0 - i.y1), abs(i.y0 - label.y1))
        if vertical <= MAX_LABEL_GAP_ABOVE_OR_BELOW:
            return vertical
    return float("inf")


def main() -> None:
    page = pdf()[0]
    drawings = page.get_drawings()
    all_icons = icons(drawings)
    out = []
    for r, tracks in labels(page, drawings):
        cy = (r.y0 + r.y1) / 2
        near = [i for i in all_icons if gap(i, r) < float("inf")]
        if not near:
            print(f"no icon for label {tracks} at ({round(r.x0)}, {round(r.y0)})")
            continue
        icon = min(near, key=lambda i: gap(i, r))
        out.append(
            {
                "platform": TRACKS_TO_PLATFORM[tracks],
                "track_numbers": tracks,
                "level": "lower" if cy > LOWER_LEVEL_MIN_Y else "upper",
                "type": icon.type,
                "map_x": round((icon.rect.x0 + icon.rect.x1) / 2),
                "map_y": round((icon.rect.y0 + icon.rect.y1) / 2),
            }
        )
    out.sort(key=lambda v: (v["platform"], v["level"], v["map_x"]))
    with OUT_CSV.open("w", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=list(out[0]), lineterminator="\n")
        writer.writeheader()
        writer.writerows(out)
    counts = Counter((v["platform"], v["type"]) for v in out)
    for platform in sorted({v["platform"] for v in out}):
        print(
            platform,
            {
                t: counts[platform, t]
                for t in ("stair", "escalator", "elevator", "other")
                if counts[platform, t]
            },
        )


if __name__ == "__main__":
    main()
