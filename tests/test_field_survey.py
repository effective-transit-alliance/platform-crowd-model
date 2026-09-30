"""Tests that regenerating the field survey sheet keeps what surveyors entered."""

import csv
import shutil
from pathlib import Path

import pytest

from platform_crowd_model import vces_field_survey


def read(path: Path) -> list[dict[str, str]]:
    with path.open() as f:
        return list(csv.DictReader(f))


def test_regenerating_keeps_survey_entries(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    """
    A generated row keeps its entries, even if its `vce_name` changed,
    and a row a surveyor added, or one that isn't generated anymore, is kept.
    """
    sheet = tmp_path / "vces_field_survey.csv"
    shutil.copy(vces_field_survey.OUT_CSV, sheet)
    monkeypatch.setattr(vces_field_survey, "OUT_CSV", sheet)
    rows = read(sheet)
    fields = list(rows[0])
    rows[0].update(found="yes", clear_width_in="71", vce_name="renamed")
    added = dict.fromkeys(fields, "")
    added.update(platform="3", found="yes", notes="a stair neither source lists")
    gone = {**rows[1], "midpoint_ft": "12345", "found": "no"}
    with sheet.open("w", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fields, lineterminator="\n")
        writer.writeheader()
        writer.writerows([*rows, added, gone])

    vces_field_survey.main()

    regenerated = read(sheet)
    first = regenerated[0]
    assert (first["found"], first["clear_width_in"]) == ("yes", "71")
    assert first["vce_name"] != "renamed"
    assert added in regenerated
    assert gone in regenerated
    assert len(regenerated) == len(rows) + 2
