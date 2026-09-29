"""
The `platform-crowd-model` command.

With no command, it runs the model.
`platform-crowd-model data ...` regenerates `data/` from each source.
Each data command imports its module only when run,
so running the model doesn't import `pymupdf`.
"""

from typing import Annotated

import typer
from typer import Option

app = typer.Typer(no_args_is_help=False, add_completion=False)

data_app = typer.Typer(
    help="Regenerate `data/` from each source.",
    no_args_is_help=True,
    add_completion=False,
)
app.add_typer(data_app, name="data")


@app.callback(invoke_without_command=True)
def run(
    ctx: typer.Context,
    update_readme: Annotated[
        bool, Option(help="Replace the results table in the README with this run's.")
    ] = False,
    charts: Annotated[
        bool,
        Option(help="Also print each scenario's time series and save its CSVs and charts."),
    ] = False,
) -> None:
    """Run every scenario and print a table of their results."""
    if ctx.invoked_subcommand is not None:
        return
    from platform_crowd_model import model

    model.main(update_readme=update_readme, charts=charts)


@data_app.command("master-plan-vces")
def master_plan_vces() -> None:
    """
    Extract each VCE's position from the Master Plan's platform-level plans,
    writing `data/master_plan_vce_positions.csv`, `data/master_plan_existing_vces.csv`,
    and `data/master_plan_platform_east_ends.csv`.
    """
    from platform_crowd_model import master_plan_vces

    master_plan_vces.main()


@data_app.command("directory-vces")
def directory_vces() -> None:
    """
    Extract each platform's VCEs from NJ Transit's January 2022 station directory,
    writing `data/njt_directory_vces.csv`.
    """
    from platform_crowd_model import directory_vces

    directory_vces.main()


@data_app.command("estimated-vces")
def estimated_vces() -> None:
    """
    Estimate every VCE's width and position from the PCIP Phase 2 plan and the directory,
    writing `data/estimated_vce_widths.csv` and `data/platform_east_ends.csv`.
    Run `master-plan-vces` and `directory-vces` first.
    """
    from platform_crowd_model import estimated_vces

    estimated_vces.main()


@data_app.command("osm-platforms")
def osm_platforms() -> None:
    """Measure each platform from OpenStreetMap, writing `data/osm_platforms.csv`."""
    from platform_crowd_model import osm_platforms

    osm_platforms.main()


@data_app.command("field-survey")
def field_survey() -> None:
    """
    Make the field survey sheet for measuring every VCE in person,
    writing `data/field_survey.csv`.
    Run `estimated-vces` first.
    """
    from platform_crowd_model import field_survey

    field_survey.main()


@data_app.command("all")
def all_data() -> None:
    """Regenerate everything in `data/` that's generated, in order."""
    master_plan_vces()
    directory_vces()
    estimated_vces()
    osm_platforms()
    field_survey()
