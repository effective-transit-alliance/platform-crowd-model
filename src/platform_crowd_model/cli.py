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
def vce_positions_master_plan() -> None:
    """
    Extract each VCE's position from the Master Plan's platform-level plans,
    writing `data/vce_positions_master_plan.csv`, `data/vces_existing_master_plan.csv`,
    and `data/platform_east_ends_master_plan.csv`.
    """
    from platform_crowd_model import vce_positions_master_plan

    vce_positions_master_plan.main()


@data_app.command()
def vces_njt_directory() -> None:
    """
    Extract each platform's VCEs from NJT's January 2022 station directory,
    writing `data/vces_njt_directory.csv`.
    """
    from platform_crowd_model import vces_njt_directory

    vces_njt_directory.main()


@data_app.command()
def vces_moynihan_ea() -> None:
    """
    Measure the VCEs around Moynihan Train Hall on the Moynihan Station EA's plan,
    writing `data/vces_moynihan_ea.csv`.
    """
    from platform_crowd_model import vces_moynihan_ea

    vces_moynihan_ea.main()


@data_app.command()
def vces() -> None:
    """
    Estimate every VCE's width and position from the PCIP Phase 2 plan and the directory,
    writing `data/vces.csv` and `data/platform_east_ends.csv`.
    Run `vce-positions-master-plan` and `vces-njt-directory` first.
    """
    from platform_crowd_model import vces

    vces.main()


@data_app.command()
def platforms_osm() -> None:
    """Measure each platform from OpenStreetMap, writing `data/platforms_osm.csv`."""
    from platform_crowd_model import platforms_osm

    platforms_osm.main()


@data_app.command()
def platform_west_ends_pcip_phase_1() -> None:
    """
    Measure where each platform ends to the west on PCIP Phase 1's existing track plan,
    writing `data/platform_west_ends_pcip_phase_1.csv`.
    Run `vces` first.
    """
    from platform_crowd_model import platform_west_ends_pcip_phase_1

    platform_west_ends_pcip_phase_1.main()


@data_app.command()
def platform_a_pcip_phase_1() -> None:
    """
    Measure PCIP Phase 1's Platform A and its VCEs on its plan of Alternative 12,
    writing `data/platform_a_pcip_phase_1.csv` and `data/vces_platform_a_pcip_phase_1.csv`.
    Run `vces` first.
    """
    from platform_crowd_model import platform_a_pcip_phase_1

    platform_a_pcip_phase_1.main()


@data_app.command()
def vces_transformation_fra_sos() -> None:
    """
    Measure Penn Transformation's new VCEs and platform extensions on the FRA's SOS report,
    writing `data/vces_transformation_fra_sos.csv` and `data/platforms_transformation_fra_sos.csv`.
    Run `vces` and `platform-west-ends-pcip-phase-1` first.
    """
    from platform_crowd_model import vces_transformation_fra_sos

    vces_transformation_fra_sos.main()


@data_app.command("all")
def all_data() -> None:
    """Regenerate everything in `data/` that's generated, in order."""
    vce_positions_master_plan()
    vces_njt_directory()
    vces_moynihan_ea()
    vces()
    platform_west_ends_pcip_phase_1()
    vces_transformation_fra_sos()
    platform_a_pcip_phase_1()
    platforms_osm()
