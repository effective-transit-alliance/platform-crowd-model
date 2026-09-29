"""
Measure each of Penn Station's platforms from OpenStreetMap,
the data OpenRailwayMap draws, via the Overpass API.

Each platform is a way tagged `railway=platform` or `public_transport=platform`
around Penn Station, either an outline (a closed way) or a line along its middle.
Its length is its extent along the tracks,
the longest distance between any two of its points,
and an outline's area is the area inside it, which accounts for platforms tapering,
and its width is its area divided by its length, i.e. its average width.
Positions are projected onto a local plane in feet, which is accurate to well under a foot here.

The raw Overpass response is cached in `.cache/osm_platforms.json`;
delete it to fetch the current data.

Writes `data/osm_platforms.csv`.
"""

import csv
import json
import math
import urllib.parse
import urllib.request
from typing import Any

from platform_crowd_model.paths import CACHE_DIR, DATA_DIR

CACHE = CACHE_DIR / "osm_platforms.json"
OUT_CSV = DATA_DIR / "osm_platforms.csv"

OVERPASS_URL = "https://overpass-api.de/api/interpreter"

BBOX = (40.7480, -74.0000, 40.7530, -73.9905)
"""South, west, north, and east edges around Penn Station's platforms and tracks."""

QUERY = f"""
[out:json][timeout:60];
(
  way["railway"="platform"]{BBOX};
  way["public_transport"="platform"]{BBOX};
);
out tags geom;
"""

FEET_PER_METER = 1 / 0.3048

METERS_PER_DEGREE_LATITUDE = 111_320


def fetch() -> dict[str, Any]:
    if not CACHE.exists():
        CACHE_DIR.mkdir(exist_ok=True)
        request = urllib.request.Request(
            OVERPASS_URL,
            data=urllib.parse.urlencode({"data": QUERY}).encode(),
            headers={
                "User-Agent": "platform-crowd-model (https://github.com/effective-transit-alliance/platform-crowd-model)",
                "Accept": "application/json",
            },
        )
        with urllib.request.urlopen(request, timeout=120) as response:
            CACHE.write_bytes(response.read())
    return json.loads(CACHE.read_text())


def project(points: list[dict[str, float]], lat0: float) -> list[tuple[float, float]]:
    """`points`' longitudes and latitudes as x and y in feet on a local plane."""
    x_scale = METERS_PER_DEGREE_LATITUDE * math.cos(math.radians(lat0)) * FEET_PER_METER
    y_scale = METERS_PER_DEGREE_LATITUDE * FEET_PER_METER
    return [(p["lon"] * x_scale, p["lat"] * y_scale) for p in points]


def main() -> None:
    elements = fetch()["elements"]
    lat0 = (BBOX[0] + BBOX[2]) / 2
    rows = []
    for way in elements:
        tags = way.get("tags", {})
        points = project(way["geometry"], lat0)
        length = max(math.dist(a, b) for a in points for b in points)
        closed = way["geometry"][0] == way["geometry"][-1] and len(points) > 3
        area = (
            abs(sum(a[0] * b[1] - b[0] * a[1] for a, b in zip(points, points[1:], strict=False)))
            / 2
            if closed
            else None
        )
        rows.append(
            {
                "osm_way": way["id"],
                "ref": tags.get("ref", ""),
                "name": tags.get("name", ""),
                "railway": tags.get("railway", ""),
                "level": tags.get("level", tags.get("layer", "")),
                "outline": "yes" if closed else "no",
                "length_ft": round(length),
                "area_sq_ft": "" if area is None else round(area),
                "width_ft": "" if area is None else round(area / length, 1),
            }
        )
    rows.sort(key=lambda row: (row["ref"], row["name"], row["osm_way"]))
    with OUT_CSV.open("w", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=list(rows[0]), lineterminator="\n")
        writer.writeheader()
        writer.writerows(rows)
    for row in rows:
        print(row)


if __name__ == "__main__":
    main()
