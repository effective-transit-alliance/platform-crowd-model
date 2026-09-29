"""Snapshot tests of the model's output."""

import pytest

from platform_crowd_model import model


def test_readme_results_table_is_up_to_date(capsys: pytest.CaptureFixture[str]) -> None:
    """
    The `README.md`'s results table is what a run prints now,
    so a change to the results fails until `--update-readme` is run and committed.
    """
    model.main()
    readme = model.README.read_text()
    start = readme.index(model.RESULTS_START) + len(model.RESULTS_START)
    end = readme.index(model.RESULTS_END)
    assert capsys.readouterr().out.strip() == readme[start:end].strip()
