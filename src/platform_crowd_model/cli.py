"""The `platform-crowd-model` command."""

from typing import Annotated

import typer
from typer import Option

app = typer.Typer(no_args_is_help=True, add_completion=False)


@app.command()
def run(
    update_readme: Annotated[
        bool, Option(help="Replace the results table in the README with this run's.")
    ] = False,
    charts: Annotated[
        bool,
        Option(help="Also print each scenario's time series and save its CSVs and charts."),
    ] = False,
) -> None:
    """Run every scenario and print a table of their results."""
    from platform_crowd_model import model

    model.main(update_readme=update_readme, charts=charts)


@app.callback()
def callback() -> None:
    """Model platform crowding and alighting and boarding at NY Penn Station."""
