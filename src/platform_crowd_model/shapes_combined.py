"""
Combine each plan's shapes into one set, taking each part of the station from its best source,
like `data/vces.csv` does for VCEs:

- Each VCE in `data/vces.csv` has its footprint from the plan its row is from:
  the PCIP Phase 2 plan's (`pcip_phase_2`) or the Moynihan Station EA's (`moynihan_ea`),
  or, where that plan doesn't draw it, e.g. the escalators under the PCIP Phase 2 plan's labels,
  or its row is from NJT's directory (`njt_directory`), the PCIP Phase 1 plan's.
- Everything else on the platforms, including VCEs not in `data/vces.csv`,
  e.g. curved stairs, is from the PCIP Phase 2 plan, the newest, where it has the platform,
  or else the PCIP Phase 1 plan, which has every platform, but cuts 5 to 8 off to the west,
  or else, further west, the Moynihan Station EA's.
  Each platform's outline is pieced together the same way.
- The concourses are the PCIP Phase 1 plan's, which has them all,
  and the Moynihan Station EA's that it doesn't have.

Writes `data/shapes.geojson` and `data/shapes.latlon.geojson`,
with each shape's `source` saying which plan it's from.
Run each plan's `shapes-*` command first.
"""

import csv
import json
from pathlib import Path
from typing import Any

from shapely import LineString, Polygon, unary_union
from shapely.geometry.base import BaseGeometry

from platform_crowd_model.paths import DATA_DIR
from platform_crowd_model.shapes import (
    FRAME,
    FT_DECIMALS,
    LONLAT_DECIMALS,
    VCES_CSV,
    counterclockwise,
    east_ends,
    frame_outlines,
    to_lonlat,
    write_geojson,
)

PCIP_PHASE_2 = DATA_DIR / "shapes_pcip_phase_2.geojson"
PCIP_PHASE_1 = DATA_DIR / "shapes_pcip_phase_1.geojson"
MOYNIHAN_EA = DATA_DIR / "shapes_moynihan_ea.geojson"
PLANS = [PCIP_PHASE_2, PCIP_PHASE_1, MOYNIHAN_EA]
"""Each plan, best first, for what's on the platforms."""

VCE_PLANS = {"pcip_phase_2": PCIP_PHASE_2, "moynihan_ea": MOYNIHAN_EA}
"""The plan each of `data/vces.csv`'s `source`s is from."""

OUT_GEOJSON = DATA_DIR / "shapes.geojson"
LATLON_GEOJSON = DATA_DIR / "shapes.latlon.geojson"

PLATFORM_REACH_FT = 3
"""How far off a platform's outline a shape on it can be, e.g. a railing along its edge."""

VCE_TYPES = ("stair", "escalator")


def features(path: Path) -> list[dict[str, Any]]:
    with path.open() as f:
        return json.load(f)["features"]


def geometry(feature: dict[str, Any]) -> BaseGeometry:
    g = feature["geometry"]
    if g["type"] == "Polygon":
        return Polygon(g["coordinates"][0])
    return LineString(g["coordinates"])


def main() -> None:
    plans = {path: features(path) for path in PLANS}

    # Where each plan has the platforms, west to east, that a better plan doesn't.
    # Shapes on the platforms can be a little off them, but their outlines' pieces have to meet.
    regions: dict[Path, BaseGeometry] = {}
    outline_regions: dict[Path, BaseGeometry] = {}
    covered: BaseGeometry = Polygon()
    outlines_covered: BaseGeometry = Polygon()
    for path in PLANS:
        outlines = unary_union(
            [geometry(f) for f in plans[path] if f["properties"]["type"] == "platform"]
        )
        regions[path] = outlines.buffer(PLATFORM_REACH_FT).difference(covered)
        covered = covered.union(regions[path])
        outline_regions[path] = outlines.difference(outlines_covered)
        outlines_covered = outlines_covered.union(outlines)

    def best(path: Path, shape: BaseGeometry) -> bool:
        """Whether `path` is the best plan for `shape`, by where its middle is."""
        return regions[path].contains(shape.representative_point())

    out: list[dict[str, Any]] = []

    # Platforms: each plan's outlines where it's the best, joined by platform.
    pieces: dict[str, list[BaseGeometry]] = {}
    sources: dict[str, list[str]] = {}
    for path in PLANS:
        for f in plans[path]:
            if f["properties"]["type"] == "platform" and f["properties"]["platform"]:
                piece = geometry(f).intersection(outline_regions[path])
                if piece.area > 1:
                    platform = f["properties"]["platform"]
                    pieces.setdefault(platform, []).append(piece)
                    if f["properties"]["source"] not in sources.setdefault(platform, []):
                        sources[platform].append(f["properties"]["source"])
    for platform, parts in sorted(pieces.items()):
        joined = unary_union([p.buffer(0.5) for p in parts]).buffer(-0.5)
        for polygon in getattr(joined, "geoms", [joined]):
            if polygon.area > 1:
                props = {
                    "type": "platform",
                    "platform": platform,
                    "level": "platform",
                    "source": "; ".join(sources[platform]),
                }
                out.append({"properties": props, "shape": polygon})

    # VCEs in `data/vces.csv`, from the plan each row is from, or else PCIP Phase 1's.
    with VCES_CSV.open() as f:
        rows = list(csv.DictReader(f))
    named = {
        path: {
            f["properties"]["vce_name"]: f
            for f in plans[path]
            if f["properties"]["type"] in VCE_TYPES and f["properties"].get("vce_name")
        }
        for path in PLANS
    }
    used: set[int] = set()
    for row in rows:
        name = row["vce_name"]
        for path in (VCE_PLANS.get(row["source"]), PCIP_PHASE_1):
            if path is not None and name in named[path]:
                feature = named[path][name]
                used.add(id(feature))
                out.append({"properties": feature["properties"], "shape": geometry(feature)})
                break

    # Everything else on the platforms, from the best plan there.
    for path in PLANS:
        for f in plans[path]:
            props = f["properties"]
            if props["type"] in ("platform", "concourse") or id(f) in used:
                continue
            # A VCE in `data/vces.csv` is only taken from its own plan, above.
            if props["type"] in VCE_TYPES and props.get("vce_name"):
                continue
            shape = geometry(f)
            if best(path, shape):
                out.append({"properties": props, "shape": shape})

    # Concourses: PCIP Phase 1's, and the EA's it doesn't have.
    pcip_1_concourses = [
        geometry(f) for f in plans[PCIP_PHASE_1] if f["properties"]["type"] == "concourse"
    ]
    for f in plans[PCIP_PHASE_1] + plans[MOYNIHAN_EA]:
        if f["properties"]["type"] != "concourse":
            continue
        shape = geometry(f)
        if f in plans[MOYNIHAN_EA] and any(
            shape.representative_point().within(c) for c in pcip_1_concourses
        ):
            continue
        out.append({"properties": f["properties"], "shape": shape})

    lonlat = to_lonlat(east_ends(frame_outlines()))
    feet_features: list[dict[str, Any]] = []
    latlon_features: list[dict[str, Any]] = []
    for item in out:
        shape = item["shape"]
        props = {"level": "platform", **item["properties"]}
        if isinstance(shape, Polygon):
            points = counterclockwise([(float(x), float(y)) for x, y in shape.exterior.coords])
        else:
            points = [(float(x), float(y)) for x, y in shape.coords]
        for target, converted, decimals in (
            (feet_features, points, FT_DECIMALS),
            (latlon_features, [lonlat(p) for p in points], LONLAT_DECIMALS),
        ):
            coordinates = [[round(c, decimals) for c in p] for p in converted]
            geom = (
                {"type": "Polygon", "coordinates": [coordinates]}
                if isinstance(shape, Polygon)
                else {"type": "LineString", "coordinates": coordinates}
            )
            target.append({"type": "Feature", "properties": props, "geometry": geom})
    write_geojson(OUT_GEOJSON, feet_features, FRAME)
    write_geojson(LATLON_GEOJSON, latlon_features, None)
    counts: dict[str, int] = {}
    for item in out:
        key = f"{item['properties']['type']} from {item['properties'].get('source', '')[:12]}"
        counts[key] = counts.get(key, 0) + 1
    print(f"wrote {len(out)} shapes")
    for key, n in sorted(counts.items()):
        print(f"  {n} {key}")
