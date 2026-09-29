"""
The `platform-crowd-model` command.

`platform-crowd-model run` runs the model,
and `platform-crowd-model data ...` regenerates `data/` from each source.
Each data command imports its module only when run.
"""

from typing import Annotated

from typer import Option, Typer

app = Typer(no_args_is_help=True)

data_app = Typer(help="Regenerate `data/` from each source.", no_args_is_help=True)
app.add_typer(data_app, name="data")


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


@data_app.command()
def platforms_osm() -> None:
    """Measure each platform from OpenStreetMap, writing `data/platforms_osm.csv`."""
    from platform_crowd_model import platforms_osm

    platforms_osm.main()
