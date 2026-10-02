"""Tests that `data/vces.csv` and the combined shapes have the same VCEs."""

import csv
import json

from platform_crowd_model import shapes_combined
from platform_crowd_model.vces import DIRECTORY_VCE_SOURCE, OUT_CSV

MAX_END_DIFFERENCE_FT = 1
"""Ends are rounded to the foot in `data/vces.csv`, and to 0.1 ft in the shapes."""


def test_shapes_match_vces() -> None:
    """
    Every VCE in `data/vces.csv` from a plan has a shape of the same name and type,
    with the same ends, and every VCE shape on a numbered platform is in `data/vces.csv`.
    NJT's directory's VCEs, which aren't on any plan, have no shapes.
    """
    with OUT_CSV.open() as f:
        rows = {row["vce_name"]: row for row in csv.DictReader(f)}
    with shapes_combined.OUT_GEOJSON.open() as f:
        features = [
            feature
            for feature in json.load(f)["features"]
            if feature["properties"]["type"] in shapes_combined.VCE_TYPES
        ]
    shapes = {}
    for feature in features:
        props = feature["properties"]
        if props.get("platform", "").isdigit():
            assert props.get("vce_name") in rows, f"unnamed {props['type']}: {props}"
        if props.get("vce_name"):
            xs = [x for x, _ in feature["geometry"]["coordinates"][0]]
            shapes[props["vce_name"]] = props["type"], min(xs), max(xs)
    for name, row in rows.items():
        if row["source"] == DIRECTORY_VCE_SOURCE:
            assert name not in shapes, f"{name} is only in NJT's directory, but has a shape"
            continue
        assert name in shapes, f"{name} has no shape"
        vce_type, west, east = shapes[name]
        assert vce_type == row["type"], name
        assert abs(west - int(row["west_end_ft"])) <= MAX_END_DIFFERENCE_FT, name
        assert abs(east - int(row["east_end_ft"])) <= MAX_END_DIFFERENCE_FT, name
