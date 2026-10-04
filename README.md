# Platform Crowd Model

This repo is for modeling platform crowding and alighting and boarding of trains,
specifically at NY Penn Station.
[`src/platform_crowd_model/model.py`](./src/platform_crowd_model/model.py)
models a single platform with its tracks.

## Running

To run, we use the Python package manager [`uv`](https://github.com/astral-sh/uv),
which can be installed easily with

```sh
curl -LsSf https://astral.sh/uv/install.sh | sh
```

With `uv` installed, you can then just run the model:

```sh
uv run platform-crowd-model run
```

In doing so, `uv` will also install dependencies and set up a virtual environment.

Each run prints a table of headline results.
To also save each scenario's full time series as CSVs and its charts as an SVG in `output/`, run

```sh
uv run platform-crowd-model run --charts
```

To also replace the [results table](#current-results) below with that run's, run

```sh
uv run platform-crowd-model run --update-readme
```

Each platform's dimensions and VCEs are in [`data/`](./data).
The generated files there come from each source's commands,
which download and cache the sources in `.cache/`.
To regenerate them all, run

```sh
uv run platform-crowd-model data all
```

`uv run platform-crowd-model data --help` lists each source's command.

Everything the model assumes, like stair capacity, passenger loads, and LOS thresholds,
is in the `Assumptions` `dataclass` at the top of
[`src/platform_crowd_model/model.py`](./src/platform_crowd_model/model.py),
right below its constants,
with each assumption's units and source.
To test how sensitive the results are to one,
set a scenario's `Params.assumptions`, e.g. to `Assumptions(stair_capacity=15)`.

## Checking

We also use `ruff` for formatting and linting, `ty` and `pyrefly` for type checking,
and `pytest` for a snapshot test that fails if the [results table](#results) is out of date.
To run these, which are also checked in CI,
you can run

```sh
uv run ruff format # format
uv run ruff check # lint
uv run ty check # type check
uv run pyrefly check # type check
uv run pytest # test
```

These same checks also run as `pre-commit` hooks.
To install them, run

```sh
uv run pre-commit install
```

Every simulation also checks that it neither created nor lost passengers,
and stops with an error if it did.

## How the Model Works

The model is the one ETA's report
[Penn Station Can Handle the Load](https://www.etany.org/penn-station-can-handle-the-load) used,
as of the [`penn-station-can-handle-the-load`](https://github.com/effective-transit-alliance/platform-crowd-model/tree/penn-station-can-handle-the-load) tag.
Later changes to the model aren't reflected in the report.
Since then, several bugs have been fixed and improvements made,
each summarized in the [results' history](#history);
this describes how it works now.

The sources cited below are:

- [ETA's report](https://www.etany.org/penn-station-can-handle-the-load),
  which used this model as of the [`penn-station-can-handle-the-load`](https://github.com/effective-transit-alliance/platform-crowd-model/tree/penn-station-can-handle-the-load) tag.
- [Fruin, "Designing for Pedestrians: A Level-of-Service Concept"](https://onlinepubs.trb.org/Onlinepubs/hrr/1971/355/355-001.pdf),
  Highway Research Record 355 (1971).
- [The Transit Capacity and Quality of Service Manual (TCQSM), 3rd edition, chapter 10](https://onlinepubs.trb.org/onlinepubs/tcrp/tcrp_rpt_165ch-10.pdf),
  TCRP Report 165 (2013).

The model simulates one platform, served by two tracks or, on platform 9, one, one second at a time,
and tracks four groups of passengers:

- Passengers aboard each train
- Arriving passengers on the platform heading upstairs
- Departing passengers on the platform waiting to board each train
- Departing passengers still upstairs on the concourse

Each second, four things happen in order,
each using the passenger counts left by the one before.

1. **Alighting.**
   Once a train arrives, its passengers step off onto the platform
   at 1 pax/s per single-door equivalent.
   - Every train is as long as its platform's tracks and the platform itself allow, up to 12 cars:
     9 cars on platforms 1 and 2, 10 on platforms 3 and 9, and 12 on the rest.
     - The tracks' limits are from
       [a track map of unknown origin](https://www.railfanguides.us/ny/penntonewrochelle/PennStationLayout1.jpg)
       found at Railfan Guides of the U.S.,
       recorded in [`data/platform_max_cars_track_map.csv`](./data/platform_max_cars_track_map.csv).
     - The platforms' lengths are from the Moynihan Station environmental assessment
       ([Table 4.4-10, p. 4.4-22](https://web.archive.org/web/20241011135133/https://cdn.esd.ny.gov/subsidiaries_projects/msdc/Data/NEPA/04_4%20StationPedCirculation.pdf#page=22)),
       recorded in [`data/platform_lengths_moynihan_ea.csv`](./data/platform_lengths_moynihan_ea.csv).
       A train fits if its 85' cars overhang the platform by at most 15'.
   - Every car is a full, seated NJT MultiLevel car with 135 passengers,
     from the same EA, whose 12-car trains carry 1,620
     ([Table 4.4-19, p. 4.4-47](https://web.archive.org/web/20241011135133/https://cdn.esd.ny.gov/subsidiaries_projects/msdc/Data/NEPA/04_4%20StationPedCirculation.pdf#page=47)).
   - Every car has 4 doors, a NJT MultiLevel's, the worst case.
     (A LIRR car has more doors, and LIRR platforms are generally wider.)
2. **Going upstairs.**
   Arriving passengers leave the platform
   via the vertical circulation elements (VCEs), i.e. the stairs and escalators.
   - Each platform's VCEs are from [`data/vces.csv`](./data/vces.csv)
     (see [VCE Width Data](#vce-width-data)).
     They're all treated as stairs, except the widest escalator, which is left out,
     like the ETA report's one VCE per platform, e.g. an escalator running the other way.
   - They queue at the stairs, which discharge them at LOS E capacity,
     17 pax/min per foot of VCE width
     ([Fruin, p. 14](https://onlinepubs.trb.org/Onlinepubs/hrr/1971/355/355-001.pdf#page=14)),
     as long as anyone is queued
     ([TCQSM, p. 10-51](https://onlinepubs.trb.org/onlinepubs/tcrp/tcrp_rpt_165ch-10.pdf#page=55)).
   - So there's no gradual taper: the stairs stay at capacity until the platform is clear.
   - The report's "taper time" is when the remaining arriving passengers fit in the stair queues,
     20' of queue in front of the VCEs at 5 sq ft/pax
     (the TCQSM's stair queuing space).
     It's now always about 15 s before the clear time,
     but it's kept in the results table to compare with the report.
   - The few seconds of walking from the doors to the stairs are ignored.
3. **Coming downstairs.**
   Departing passengers queue upstairs and come down to the platform.
   - The trains' passengers split the VCE width
     in proportion to how many of each are still upstairs,
     and each train's share of the stairs carries the same share of the upward flow.
   - Nobody comes down while the upward flow is worse than LOS C, 10 pax/min/ft.
   - Otherwise, both directions share LOS E capacity, 17 pax/min/ft,
     so passengers come down with whatever their share of the upward flow leaves of that.
4. **Boarding.**
   Departing passengers on the platform board a train
   once every arriving passenger has alighted from it,
   at 1 pax/s per single-door equivalent.

### Crowding

- **On the platform**, the space per passenger is the usable platform area
  (75% of the platform's area) divided by everyone on the platform.
  It's graded with Fruin's LOS for queuing and waiting areas
  (A > 13, B > 10, C > 7, D > 3, E > 2 sq ft/pax, or else F;
  [TCQSM, Exhibit 10-32, p. 10-55](https://onlinepubs.trb.org/onlinepubs/tcrp/tcrp_rpt_165ch-10.pdf#page=59)),
  since most passengers on the platform are waiting, either to board or in the stair queues.
- **On the stairs**, the upward flow is graded with Fruin's stair LOS
  (A ≤ 5, B ≤ 7, C ≤ 10, D ≤ 13, E ≤ 17 pax/min/ft, or else F;
  [Fruin, pp. 12–14](https://onlinepubs.trb.org/Onlinepubs/hrr/1971/355/355-001.pdf#page=12)).

### Scenarios

- Each train has 400 departing passengers, all waiting upstairs,
  who start coming down 2 minutes before it's scheduled.
- On every platform, 4 trains are scheduled 0, 2, or 5 minutes apart,
  alternating between the platform's two tracks, or all on platform 9's one,
  so each train after the first on each track can only arrive
  once the train before it on its track has departed at the end of its dwell.
  0 minutes apart, the first two arrive together and the last two as soon as they depart,
  or on platform 9, each as soon as the one before it departs.
- Every platform, 1 to 11, is modeled,
  and platform 3 also with its VCEs after Penn Reconstruction, `(recon)`,
  with the ETA report's total width for them, 44'9".

### NFPA 130 Evacuation

NFPA 130 (2026 edition, 5.3.3.1) requires "sufficient egress capacity to evacuate the platform occupant load
… from the station platform in 4 minutes or less".
Its Annex C computes this as the platform occupant load divided by the egress capacity,
without walking time, which is part of its separate 6-minute limit to reach a point of safety (5.3.3.2).
NFPA 130's text isn't free to read, so these are from its 2007 edition,
with the changes since from NFPA's public [revision reports](https://docinfofiles.nfpa.org/files/AboutTheCodes/130/),
e.g. the [2023 edition's First Draft Report](https://docinfofiles.nfpa.org/files/AboutTheCodes/130/130_A2022_FKT_AAA_FRReport.pdf).

- NFPA 130's platform occupant load is, for each track, its train load plus its entraining load (5.3.2.5),
  each "per train headway factored to account for service disruptions and system reaction time",
  usually by doubling one headway (A.5.3.2.5(2)),
  with the train load at most the trains' capacity,
  and the results checked not to "exceed the system service capacity".
  At Penn Station, the North River Tunnels cap the headways,
  so the model instead checks the moment the most passengers are on the platform or aboard its trains
  in each simulated headway,
  counting each train from its arrival until it departs,
  so a disruption is the 0-minute headway scenario, with two trains arriving at once.
- Trains carry their seated capacity, like the rest of the Penn Station literature,
  not their "maximum passenger capacity" with standees (5.3.2.5(5)),
  though no more seated passengers have been measured through the North River Tunnels,
  24 trains per hour × 12 cars × 135 seats.
- Stairs and stopped escalators carry 1.41 pax/min per inch of width, i.e. 16.92 pax/min/ft (5.3.5.3).
- The widest escalator is out of service, as in the simulation,
  taken as the one "having the most adverse effect upon egress capacity" (5.3.5.4).
  NFPA 130 also lets escalators provide at most half of the egress capacity (5.3.5.6),
  but the model doesn't limit them yet, which overstates the exit capacity.

NFPA 130 also requires evacuating "from the most remote point on the platform to a point of safety in 6 minutes or less" (5.3.3.2).
The model checks this per its Annex C, at the same moment:

- The point of safety is the concourse,
  which NFPA 130 allows in an enclosed station like Penn Station
  only where its emergency ventilation protects the concourse from a train fire at the platform,
  as confirmed by an engineering analysis (5.3.3.4).
- The farthest occupant walks along the platform to their nearest VCE at 124 fpm (5.3.4.4).
  The model doesn't know where the VCEs are yet,
  so they walk the farthest NFPA 130 allows, 325' (5.3.3.5), taking 2:37.
- They wait there for the rest of the platform's flow time, if any (Annex C's W_p = F_p − T_p),
  so they reach the VCE at the longer of their walk and the 4-minute check's flow time.
- They climb 16'9¼", from the existing platforms to the existing concourse
  on the PCIP Phase 2 plan's north-south cross section (sheet A-213, PDF page 54),
  at a vertical speed of 48 fpm (5.3.5.3), taking 21 s.

## Assumptions

Beyond the numbers above, which are all in `Assumptions`, the model assumes the following.
Each is marked by which way it biases the results:
**optimistic** (less crowding than reality), **pessimistic** (more),
**neutral** (neither, e.g. it only simplifies the bookkeeping),
or **unclear** (it could go either way).

### Platform

- **Neutral:** The platform is a single island platform serving two tracks,
  except platform 9, which only serves track 17.
- **Neutral:** Its area is that of its outline in OpenStreetMap, which OpenRailwayMap draws,
  in [`data/platforms_osm.csv`](./data/platforms_osm.csv),
  written by `uv run platform-crowd-model data platforms-osm`
  ([`platforms_osm.py`](./src/platform_crowd_model/platforms_osm.py)),
  since platforms taper toward their ends.
  The outlines have no source, so they're only as accurate as whoever drew them.
- **Optimistic:** 75% of its area is usable, the rest taken by columns, stairs, and other obstructions
  (see [Usable Platform Area](#usable-platform-area)).
- **Optimistic:** Passengers are spread evenly over the whole usable area,
  so local crowding, e.g. at the foot of the stairs or at the doors, isn't modeled.
- **Neutral:** Space per passenger counts everyone on the platform:
  arriving passengers heading up and departing passengers waiting to board.

### Stairs and Escalators

- **Pessimistic:** All VCEs are treated as stairs, even escalators, which have higher capacities.
- **Pessimistic:** Each platform's widest escalator is excluded, e.g. as if it's running the other way.
- **Unclear:** Most VCEs' widths are estimated from a scaled plan,
  and platforms 9 to 11's positions are partly from a schematic map
  (see [VCE Width Data](#vce-width-data)).
- **Optimistic:** All VCEs act as one pooled queue:
  passengers spread across them in proportion to their widths,
  with no preference for any exit, e.g. toward 7th Avenue.
- **Unclear:** Stair capacity is linear in width,
  though the TCQSM notes capacity is really stepped by the number of pedestrian lanes
  ([p. 10-49](https://onlinepubs.trb.org/onlinepubs/tcrp/tcrp_rpt_165ch-10.pdf#page=53)).
- **Optimistic:** Stair capacity doesn't depend on the stair's rise,
  though long climbs slow people down ([p. 10-51](https://onlinepubs.trb.org/onlinepubs/tcrp/tcrp_rpt_165ch-10.pdf#page=55)),
  or on luggage, strollers, or wheelchairs.
- **Optimistic, by a few seconds:** Walking from the doors to the stairs takes no time;
  passengers can go upstairs the same second they alight.
- **Optimistic:** The concourse upstairs never backs up, so the stairs always discharge.

### Trains

- **Pessimistic:** Every train is as long as its platform allows, and every seat is full.
- **Unclear:** A train can overhang its platform by up to 15', which isn't from any source.
- **Pessimistic:** Every arriving passenger alights; nobody stays on a through train.
- **Optimistic:** Passengers alight and board at 1 pax/s per single-door equivalent,
  starting the second after the train arrives, with no delay for the doors to open.
- **Pessimistic:** Each train departs at the end of its dwell,
  once all of its departing passengers have boarded, so nobody misses a train.

### Departing Passengers

- **Optimistic:** Each train has 400 departing passengers, all waiting upstairs,
  who start coming down 2 minutes before it's scheduled, like when its track is announced.
  None arrive later, and the 2 minutes isn't from any source.
- **Pessimistic:** Departing passengers come downstairs as soon as they can,
  even before their train arrives, and wait for it on the platform, adding to the crowding.

### Simulation

- **Neutral:** It steps through time 1 s at a time, with fractional passengers,
  until the last train departs and the platform clears.
- **Neutral:** Each second, passengers alight, then go upstairs, then come downstairs, then board,
  each using the counts left by the previous step.
- **Neutral:** The stair LOS grades only the upward flow, not the downward flow.
- **Optimistic:** The NFPA 130 evacuation time uses the simulated peak occupants,
  without NFPA 130's accumulation of waiting passengers over a doubled headway, or standees.
- **Optimistic:** The NFPA 130 time to the concourse treats the concourse as a point of safety,
  which NFPA 130 only allows where an engineering analysis shows it's protected (5.3.3.4).
- **Pessimistic:** The NFPA 130 time to the concourse has the farthest occupant walk 325',
  the farthest NFPA 130 allows, instead of to their nearest VCE.
- **Optimistic:** The NFPA 130 evacuation time doesn't limit escalators to half of the exit capacity.

## Limitations

The model is much simpler than a pedestrian microsimulation,
like those typically used in detailed station planning,
and several of its simplifications overstate platform capacity.

### Comparison With the FRA's Pedestrian Simulation

The FRA's [New York Penn Station Service Optimization Study, Phase I Report](https://railroads.dot.gov/elibrary/new-york-penn-station-service-optimization-study-phase-i-report-final-june-2026)
(June 2026) simulated passengers alighting and boarding at Penn Station
with a pedestrian simulation, using VCE widths, platform widths, and obstructions
measured on site ([p. 3-28](https://railroads.dot.gov/sites/fra.dot.gov/files/2026-07/2026.07.13_Penn%20Station%20SOS_Phase%20I%20Report_FINAL.pdf#page=39)).
Its stress test is close to this model's platform 3 scenario with trains 2 minutes apart:
trains alighting and boarding on both tracks of the same platform, 2.5 minutes apart,
with up to 1,600 passengers per commuter train ([p. 4-40](https://railroads.dot.gov/sites/fra.dot.gov/files/2026-07/2026.07.13_Penn%20Station%20SOS_Phase%20I%20Report_FINAL.pdf#page=51)).

For NJ Transit trains that only alight ("drop and go"),
its baseline, i.e. today's VCEs,
needed up to 7.9 minutes of passenger service time,
the time for passengers to alight, cross the platform, and reach the VCEs
([Table 2, p. 4-43](https://railroads.dot.gov/sites/fra.dot.gov/files/2026-07/2026.07.13_Penn%20Station%20SOS_Phase%20I%20Report_FINAL.pdf#page=54)).
This model clears platform 3's arriving passengers in 4:29,
so it's likely substantially optimistic,
though the two aren't exactly comparable:
7.9 minutes is the worst case across all platforms and simulation runs.

The FRA attributes long clearance times to things this model leaves out
([p. 3-33](https://railroads.dot.gov/sites/fra.dot.gov/files/2026-07/2026.07.13_Penn%20Station%20SOS_Phase%20I%20Report_FINAL.pdf#page=44)):
queues at the base of VCEs, uneven use of VCEs, and platform clutter
reducing the usable width.

### Stairs Are One Pooled Queue

Like the TCQSM and NFPA 130 for platform clearance
([TCQSM, p. 10-79](https://onlinepubs.trb.org/onlinepubs/tcrp/tcrp_rpt_165ch-10.pdf#page=83)),
the model treats all of the VCEs as one queue discharging at capacity until it's empty.
This is optimistic for a whole platform:

- **Stairs empty unevenly.**
  Stairs near the ends of the platform, or far from the busiest doors,
  run out of passengers while others still have a queue,
  so the total flow drops below capacity before the platform clears.
  This is likely the model's largest optimistic bias.
- **Passengers prefer some exits**, e.g. toward 7th Avenue, as the ETA report notes,
  concentrating queues at fewer stairs.
- **Walking from the doors to the stairs takes no time.**
  This only shifts the results by a few seconds, but also ignores
  passengers crossing through crowds of waiting passengers.

Modeling each VCE's own queue with walking distances would fix these.
Each VCE's position and width are now in [`data/vces.csv`](./data/vces.csv)
(see [VCE Width Data](#vce-width-data)),
but the model only uses their widths so far.

### Platform Crowding Is Graded Against the Whole Platform

Though the model now grades the platform with the queuing LOS,
the TCQSM's platform sizing procedure ([p. 10-56](https://onlinepubs.trb.org/onlinepubs/tcrp/tcrp_rpt_165ch-10.pdf#page=60))
only applies the queuing LOS to passengers *waiting* to board, and adds separate areas for:

- Walkway width for arriving passengers, graded as a walkway
- Queue storage at the stairs
- A 1'6" buffer along each platform edge

The model lumps everyone into one area,
which is optimistic with the queuing LOS,
especially while arriving passengers are walking to the stairs.

### Usable Platform Area

The model uses 75% of the platform's area,
leaving 25% for columns, stairwells, and other obstructions.
But the TCQSM's 1'6" edge buffers alone take 3' of an 18' platform, about 17%,
leaving only 8% for everything else, which is likely too little on a narrow platform
with stairwells in it.
So the usable area, and the space per passenger, are likely overstated (optimistic).

### Total VCE Width

The ETA report's total VCE widths, whose source is unknown,
are only used for platform 3 with Penn Reconstruction, since each VCE now has its own width.
Its platforms 10 and 11's match the
[Moynihan Station EA's](https://web.archive.org/web/20241011135133/https://cdn.esd.ny.gov/subsidiaries_projects/msdc/Data/NEPA/04_4%20StationPedCirculation.pdf#page=22)
2008 per-platform stair capacities divided by 17 pax/min/ft,
but its platform 3's matches the EA's platform 1, not platform 3,
and they probably don't reflect the VCEs added since,
with the West End Concourse's expansion and Moynihan Train Hall.

## VCE Width Data

A stair-by-stair model needs each VCE's width and position on each platform.
The best public source of widths found so far is the NY Penn Station Master Plan's
tables of VCE widths for each platform under each reconstruction alternative,
transcribed in [`data/vce_widths_master_plan.csv`](./data/vce_widths_master_plan.csv):

- Alternatives 1 and 2 from the [August 2020 draft Alternatives Report](https://liamblank.com/wp-content/uploads/2026/07/R-Master-Plan-MTA-NJT-AMT-20_0814-PSMP-Alternatives-Report.pdf)
  ([p. 16](https://liamblank.com/wp-content/uploads/2026/07/R-Master-Plan-MTA-NJT-AMT-20_0814-PSMP-Alternatives-Report.pdf#page=22) and [p. 28](https://liamblank.com/wp-content/uploads/2026/07/R-Master-Plan-MTA-NJT-AMT-20_0814-PSMP-Alternatives-Report.pdf#page=34)).
    Its tables for Alternatives 3 and 4 ([p. 40](https://liamblank.com/wp-content/uploads/2026/07/R-Master-Plan-MTA-NJT-AMT-20_0814-PSMP-Alternatives-Report.pdf#page=46) and [p. 52](https://liamblank.com/wp-content/uploads/2026/07/R-Master-Plan-MTA-NJT-AMT-20_0814-PSMP-Alternatives-Report.pdf#page=58)) are identical.
- Alternatives 3 and 4 from the same draft.
- Alternative 3 from the April 2021 final report ([p. 117](https://esd.ny.gov/sites/default/files/CACWG-Meetings-8-9-Q-A-09-08-21.pdf#page=7)),
  as reproduced in the Empire Station Complex Q&A.

Each VCE is marked as a stair or escalator, and as new or existing.
Widths are in inches, listed in the tables' order, which is west to east along the platform.

`uv run platform-crowd-model data vce-positions-master-plan`
([`vce_positions_master_plan.py`](./src/platform_crowd_model/vce_positions_master_plan.py))
extracts each VCE's approximate position from the draft's platform-level plans
([p. 17](https://liamblank.com/wp-content/uploads/2026/07/R-Master-Plan-MTA-NJT-AMT-20_0814-PSMP-Alternatives-Report.pdf#page=23), [p. 29](https://liamblank.com/wp-content/uploads/2026/07/R-Master-Plan-MTA-NJT-AMT-20_0814-PSMP-Alternatives-Report.pdf#page=35),
[p. 41](https://liamblank.com/wp-content/uploads/2026/07/R-Master-Plan-MTA-NJT-AMT-20_0814-PSMP-Alternatives-Report.pdf#page=47), and [p. 53](https://liamblank.com/wp-content/uploads/2026/07/R-Master-Plan-MTA-NJT-AMT-20_0814-PSMP-Alternatives-Report.pdf#page=59)).
They're vector drawings with each VCE drawn as a rectangle color-coded by type and age,
and a scale bar to convert to feet.
The script matches the `n`th VCE from the west on each platform to the `n`th VCE in its table,
since the plans are too small to measure widths from,
and writes [`data/vce_positions_master_plan.csv`](./data/vce_positions_master_plan.csv),
with positions in feet east of the plans' west edge,
which cuts across every platform at the same place, under the West End Concourse.
The plans have west on the left:
the platforms' east ends on them match those on a scaled existing-conditions plan
in NJT's PCIP Phase 2 drawings to within about 4'
(see [Estimated Widths](#estimated-widths)).
The draft's plans and tables don't always agree:
often, they disagree on whether a VCE is new or existing, so both are recorded,
and on 16 of the 44 platforms, they list different stairs and escalators, so those are skipped.

Since each alternative keeps a different subset of the existing VCEs,
the script also combines the existing VCEs across alternatives by position into
[`data/vces_existing_master_plan.csv`](./data/vces_existing_master_plan.csv),
and measures each platform's east end from its outline, averaged across the alternatives, into
[`data/platform_east_ends_master_plan.csv`](./data/platform_east_ends_master_plan.csv).
It finds all 5 of the existing VCEs the Master Plan counts on platform 3,
matching both the final report's count and the 2 escalators and 3 stairs in the
[Moynihan Station EA's Table 4.4-10](https://web.archive.org/web/20241011135133/https://cdn.esd.ny.gov/subsidiaries_projects/msdc/Data/NEPA/04_4%20StationPedCirculation.pdf#page=22),
but only some of the other platforms' (none of platform 11's).
Those 5 may not be all of platform 3's VCEs, though.
The January 2022 NJT station directory shows 8 stairs and escalators to platform 3,
including 2 stairs down from the West End Concourse, under Moynihan Train Hall, west of 8th Ave.
The Master Plan's platform plans cut every platform off under the West End Concourse,
and show none of its stairs,
so the Master Plan probably doesn't include the West End Concourse's VCEs,
though it's also possible that the directory, a schematic wayfinding map,
shows some stairs more than once
(see [Sources and Their Dates](#sources-and-their-dates)).

This data couldn't replace the ETA report's total VCE widths on its own:

- **It doesn't fully describe the station today.**
  Each alternative replaces most existing VCEs,
  and the tables only list the ones each alternative keeps,
  so even combined, they're missing many existing VCEs.
  E.g. the final Alternative 3's table says platform 6 has 7 VCEs today,
  but keeps none of them.
- **The tables' printed totals are inconsistent.**
  Most exclude one escalator from the sum of the widths,
  like the ETA report's exclusion of one VCE per platform,
  but platforms 4 to 8 exclude nothing,
  and Alternative 2's platforms 2 and 3 exclude a stair.
- **It doesn't match the ETA report's total VCE widths.**
  E.g. platform 3's VCEs sum to 34'4" to 38'2" across the alternatives, even with their new VCEs,
  but the ETA report used 42'6" today,
  and platform 6's sum to 36'5" to 38'2", but the ETA report used 48'2".
  The source of the ETA report's widths is unknown.
  Platform 3's 5 existing VCEs in the Master Plan total 19'5":
  3 stairs (13'9") and 2 escalators (5'8").
  At 17 pax/min/ft for the stairs and typical escalator capacities,
  that's roughly the Moynihan Station EA's 437 pax/min for platform 3 in 2008,
  while the ETA report's 42'6" is 722 pax/min, the EA's figure for platform 1.
  But the Master Plan probably doesn't include the West End Concourse's VCEs,
  and the EA's data predates Moynihan Train Hall,
  so platform 3's total width today is probably more than 19'5".
  With the VCEs on NJT's scaled PCIP Phase 2 plan, whose widths are mostly
  [estimated](#estimated-widths), it's about 45'10", a little more than the ETA report's 42'6",
  or 43' with the Master Plan's widths where it has them, as the model now uses.
  Either way, a single total overstates how quickly a platform clears
  if some of that width is at its far west end, far from most of the train's doors.
- **Positions are approximate.**
  They're scaled from small drawings, so they're only accurate to within about 10'.

The FRA's Service Optimization Study measured today's VCEs on site,
but didn't publish the measurements.

### Counts from NJT's Station Directory

NJT's [January 2022 station directory](https://content.njtransit.com/sites/default/files/NY%20Penn%20Station%20Directory_011022.pdf),
which the ETA report links to,
is the only source found that shows every platform's VCEs after Moynihan Train Hall opened.
It's a vector wayfinding map of both concourse levels,
with an icon for each stair, escalator, and elevator to a platform,
labeled with the platform's tracks.
`uv run platform-crowd-model data vces-njt-directory`
([`vces_njt_directory.py`](./src/platform_crowd_model/vces_njt_directory.py))
extracts and classifies these icons into
[`data/vces_njt_directory.csv`](./data/vces_njt_directory.csv),
matching each of the map's 104 track labels to a distinct icon.
Compared with the Master Plan final report's counts of existing VCEs:

| Platform | Tracks | Stairs | Escalators | Elevators | Stairs and escalators | Master Plan's existing VCEs |
|---|---|---|---|---|---|---|
| 1 | 1/2 | 6 | 0 | 2 | 6 | 8 |
| 2 | 3/4 | 7 | 1 | 2 | 8 | 8 |
| 3 | 5/6 | 6 | 2 | 2 | 8 | 5 |
| 4 | 7/8 | 6 | 1 | 3 | 7 | 5 |
| 5 | 9/10 | 8 | 1 | 3 | 9 | 6 |
| 6 | 11/12 | 6 | 1 | 3 | 7 | 7 |
| 7 | 13/14 | 6 | 1 | 2 | 7 | 7 |
| 8 | 15/16 | 6 | 1 | 2 | 7 | 7 |
| 9 | 17 | 6 | 1 | 2 | 7 | 7 |
| 10 | 18/19 | 6 | 1 | 2 | 7 | 8 |
| 11 | 20/21 | 5 | 1 | 2 | 6 | 7 |

The two don't reconcile, in either direction:
the directory shows more on platforms 3 to 5, including the West End Concourse's,
but fewer on platforms 1, 10, and 11.
Neither is clearly complete:

- The directory is schematic and not to scale,
  so it has no widths, and its positions are only approximate.
- It may show a VCE more than once,
  e.g. a stair on each side of a concourse that crosses over a platform may be one stair or two,
  and may leave out VCEs, e.g. ones it doesn't label with tracks.
- It's unclear whether the Master Plan's counts include elevators.

So these counts are a check on the other sources, not a replacement for the FRA's measurements.

### Estimated Widths

Until the VCEs are measured, widths no source lists can be estimated from a scaled drawing.
NJT's PCIP Phase 2 drawings include an
[existing concourse-level plan](https://liamblank.com/wp-content/uploads/2026/09/pcip-2-conceptual-design-preliminary-drawings.pdf#page=45)
(sheet A-001, November 2020), a vector drawing at 1" = 40' of platforms 1 to 8,
including the West End Concourse, with each stair's and escalator's treads drawn as lines.
No such drawing of platforms 9 to 11 was found.
`uv run platform-crowd-model data vces`
([`vces.py`](./src/platform_crowd_model/vces.py))
finds every run of treads on the platforms, merges a stair's flights,
and measures each VCE's width and position into
[`data/vces.csv`](./data/vces.csv),
finding them with [`shapes.py`](./src/platform_crowd_model/shapes.py), including curved stairs:

- A VCE's position is its footprint's, including its landings and balustrades.

- Positions are in the Master Plan's frame, feet east of its plans' west edge.
  The two drawings register to within about 1'2":
  the platforms' east ends on them are all the same distance apart.
  Each platform's east end in that frame
  is in [`data/platform_east_ends.csv`](./data/platform_east_ends.csv):
  the sheet's on platforms 1 to 8, and the Master Plan's on platforms 9 to 11.
- VCEs matching a Master Plan VCE of the same type within 15'
  have `width_source` `master_plan`, with the Master Plan's width.
  The rest, 46 of 68, are `estimated`, with only the sheet's width.
- Their `source` is `pcip_phase_2`.
  Platforms 9 to 11's VCEs, which aren't on the sheet,
  have `source` `njt_directory`: they're from NJT's directory, as described below.
- For the 10 matched stairs, the sheet's widths differ from the Master Plan's by up to 1'2"
  (a median of 3½"), and for the 6 matched escalators, by up to 9"
- Escalators' treads are their steps, narrower than their balustrades.
- The platforms' labels cover an escalator the Master Plan has about 230' along each of platforms 3 to 8.
  The sheet still has their treads under the labels, but clipped,
  so their positions are the PCIP Phase 1 plan's, which draws them,
  with `source` `pcip_phase_1`, and their type and width are the Master Plan's.

| Platform | Stairs on the sheet | Escalators on the sheet | Their total width | Of which estimated | Master Plan VCEs not matched on the sheet |
|---|---|---|---|---|---|
| 1 | 6 | 2 | 35'4" | 16'10" | 3'8"/4'9" stair at 689'; 3'10" escalator at 774'; 4'4" stair at 776' |
| 2 | 5 | 4 | 35'7" | 22'8" | 4'4" stair at 775'; 3'10" escalator at 774' |
| 3 | 6 | 3 | 45'9" | 35'9" | 5'9" stair at 278'; 3'8" stair at 684' |
| 4 | 5 | 3 | 39'4" | 30'6" | 2'10" escalator at 404'; 3'8" stair at 683' |
| 5 | 5 | 5 | 47'2" | 41'6" | 3'8" stair at 683' |
| 6 | 3 | 5 | 33'8" | 23'10" | |
| 7 | 4 | 4 | 35'5" | 24'11" | |
| 8 | 5 | 3 | 39'2" | 31'6" | 2'10" escalator at 404' |

The unmatched Master Plan VCEs are a mix:
some, like platform 3's 3'8" stair at 684',
are probably the same VCEs as similar ones on the sheet 15' to 25' away;
and some disagree on the type, like platform 3's 5'9" stair at 278',
where the sheet has a 3'1" escalator.
So platform 3's VCEs total about 45'9" with the escalator under its label,
a little more than the ETA report's 42'6",
including the 2 West End Concourse stairs (8'11" and 10') and the Exit Concourse's 2 (5'1" and 5').

These are only estimates:

- A drawn tread isn't necessarily the clear width between handrails.
- The sheet's existing stairs aren't labeled,
  so some may be stairs between the upper and lower concourses drawn over a platform,
  though every platform's matches the directory's pattern,
  e.g. 2 stairs each for the West End and Exit Concourses.
- The sheet predates NJT's replacement of an escalator on tracks 7/8 (platform 4)
  with stairs in about 2021.

Platforms 9 to 11 aren't on that plan,
so their VCEs are from NJT's PCIP Phase 1 existing plan
([Appendix A, sheet A-001](https://liamblank.com/wp-content/uploads/2026/07/penn-records-s-nj-transit-pcip-1-pcip1-final-report-c5015-01-262652-00-task-06-mem-final-report-draft-appendixa-drawings-copy.pdf#page=24),
July 2019), with `source` `pcip_phase_1`, where it draws their treads,
measured the same way,
and otherwise from NJT's January 2022 station directory,
whose map is schematic and not to scale.
The directory's VCEs are matched to the plan's, nearest first, within 60',
counting a different type as 30' further.
Of 28 VCEs east of the West End Concourse, 18 are on both,
3 only on the plan, e.g. platform 11's stairs about 700' and 890' along,
and 7 only in the directory, which keep its positions, e.g. stairs the plan draws without treads.
Where the directory and the Master Plan agree on a VCE's type but the plan doesn't,
e.g. platform 9's 6' stair about 410' along, which the plan draws skewed with only half of each tread,
it has the plan's position but their type and width.
The plan's widths include a 17'9" stair on platform 10, wider than any the Master Plan has,
which may be two side by side.

For the directory's VCEs,
each level of its map is calibrated to feet by matching its icons on platforms 1 to 8,
and the Master Plan's on platforms 9 and 10, to their positions,
to within about 20' on average.
The directory doesn't show every VCE, so the Master Plan's existing VCEs it doesn't show are added,
though they're from before Moynihan Train Hall opened.
Each has the Master Plan's width where it has one,
or else the median width of that type on platforms 1 to 8, marked `typical`.

### Moynihan Train Hall and the West End Concourse

Neither of those shows the VCEs at the platforms' west ends:
the PCIP Phase 2 plan stops just west of the West End Concourse,
and the Master Plan's plans cut the platforms off there, too.
So the escalators down from Moynihan Train Hall were missing,
though the Moynihan Station Development Project's circulation study
([MSDC Attachments A–D, Table 13-19, p. 110](https://esd.ny.gov/sites/default/files/MSDC_GPP_attachmentsA_D.pdf#page=140))
counts 2 or 3 more escalators on each of platforms 3 to 8 than the other sources have,
as were the West End Concourse's stairs down to platforms 9 to 11.

The Moynihan Station EA's lower concourse plan
([Figure 3-4](https://web.archive.org/web/2017id_/https://cdn.esd.ny.gov/subsidiaries_projects/msdc/Data/NEPA/03a%20Figure%203-3%20and%203-4.pdf#page=2),
February 2010) draws them, as a vector drawing:
a pair of escalators, end to end, down from the Train Hall to each of platforms 3 to 8,
all at the same position, about 100' to 200' west of the West End Concourse,
and the West End Concourse's stairs down to platforms 9 to 11.
`uv run platform-crowd-model data vces-moynihan-ea`
([`vces_moynihan_ea.py`](./src/platform_crowd_model/vces_moynihan_ea.py))
finds each one's treads, scales the plan by the West End Concourse's dimensioned width of 36'-3",
which puts its 19' corridor within half an inch of 19',
and registers it by the West End Concourse's stairs down to platforms 3 to 8
to where the PCIP Phase 2 plan has them, to within about 2',
into [`data/vces_moynihan_ea.csv`](./data/vces_moynihan_ea.csv),
which `vces` adds to [`data/vces.csv`](./data/vces.csv) with `source` `moynihan_ea`,
except where the PCIP Phase 1 plan, which is newer, draws the same VCE within 20':
the West End Concourse's stairs to platforms 9 to 11,
and platform 9's second, which it draws as an escalator about 20' further east.
It predates the Train Hall's opening, so it doesn't draw its escalators, which stay the EA's.

The plan is a design from before the Train Hall was built, so:

- The Train Hall opened in 2021 with 11 escalators to platforms 3 to 8, not the plan's 12.
  Platform 3's western one is taken as the one that wasn't built,
  since it would run past the platform's west end.
- The plan doesn't draw the escalators consistently enough to measure their width,
  so they're taken to be 3'4" wide,
  the 1,000 mm steps [KONE](https://elevatorworld.com/article/let-there-be-light-and-accessibility/) reports for its escalators there.
- The stairs' widths are their drawn treads':
  6'6" on platforms 9 and 10, 3'8" on platform 11,
  and 6' for a second stair on platform 9, about 90' west of the West End Concourse,
  though the PCIP Phase 1 plan's are used instead: 4'8", 7'5", 4'4", and a 3'4" escalator.
- It also draws a stair down from the baggage and egress corridor
  to each of platforms 5 to 7, about 400' west of the West End Concourse, 3'5" to 3'6" wide.
  The PCIP plans cut those platforms off before them, so they're unconfirmed,
  and they're noted as such.

The PCIP Phase 1 plan also draws two VCEs no other source has,
a 4'8" stair on platform 7 about 65' west of the West End Concourse
and a 3'5" escalator on platform 8 about 73' west of it,
which are in `data/vces.csv`, too, noted as unconfirmed until they're checked in person.

### Sources and Their Dates

Penn Station's VCEs have changed over time,
most notably with the West End Concourse's expansion,
which opened in 2017,
and Moynihan Train Hall, which opened in 2021 above it,
so each source only describes the station as of its date.
From newest to oldest:

| Source | Date | Counts | Types | Widths | Positions | Includes West End Concourse |
|---|---|---|---|---|---|---|
| [FRA Service Optimization Study, Phase I](https://railroads.dot.gov/sites/fra.dot.gov/files/2026-07/2026.07.13_Penn%20Station%20SOS_Phase%20I%20Report_FINAL.pdf) | July 2026 report | Only of new VCEs | Only of new VCEs | Measured, but not published | Only of new VCEs | Yes |
| [Amtrak's *Doubling Trans-Hudson Train Capacity*](https://web.archive.org/web/20250307173712id_/http://pennstationcomplex.info/wp-content/uploads/2024/10/Doubling-of-Trans-Hudson-Train-Capacity-at-Penn-Station.pdf#page=34) | October 2024 | No | No | No | No | n/a |
| [NJT station directory](https://content.njtransit.com/sites/default/files/NY%20Penn%20Station%20Directory_011022.pdf) | January 2022 | Yes | Yes | No | Schematic | Yes |
| [Master Plan final report](https://esd.ny.gov/sites/default/files/CACWG-Meetings-8-9-Q-A-09-08-21.pdf#page=7), in the Empire Station Complex Q&A | April 2021 | Yes | Only of kept VCEs | Only of kept VCEs | No | Probably not |
| [Master Plan draft Alternatives Report](https://liamblank.com/wp-content/uploads/2026/07/R-Master-Plan-MTA-NJT-AMT-20_0814-PSMP-Alternatives-Report.pdf#page=22) | August 2020 | No | Only of kept VCEs | Only of kept VCEs | Yes, from vector plans | Probably not |
| [NJT PCIP Phase 1 and 2 reports and drawings](https://liamblank.com/wp-content/uploads/2026/09/penn-station-gateway-catalogue.csv) | 2019 to 2021 | Only of new VCEs | Only of new VCEs | Only of new VCEs, but existing ones can be estimated from a scaled plan | Only of new VCEs, but existing ones can be estimated from a scaled plan | Yes, on a scaled plan |
| [Moynihan Station EA, Table 4.4-10](https://web.archive.org/web/20241011135133/https://cdn.esd.ny.gov/subsidiaries_projects/msdc/Data/NEPA/04_4%20StationPedCirculation.pdf#page=22) | 2008 data | Yes | Yes | No, only capacities in ped/min | No | Before its expansion |
| ETA's `total_vce_width` | Unknown | No | No | Totals per platform | No | Unknown |

Notes:

- **FRA Service Optimization Study (July 2026).**
  The newest source, and the one whose pedestrian simulation this model is compared with.
  It collected "platform and VCE widths" on site
  ([p. 3-28](https://railroads.dot.gov/sites/fra.dot.gov/files/2026-07/2026.07.13_Penn%20Station%20SOS_Phase%20I%20Report_FINAL.pdf#page=39)),
  and says Moynihan "provides access to most platforms"
  ([p. 2-7](https://railroads.dot.gov/sites/fra.dot.gov/files/2026-07/2026.07.13_Penn%20Station%20SOS_Phase%20I%20Report_FINAL.pdf#page=18)),
  but its only per-VCE figure shows the up to 23 VCEs it proposes adding
  ([p. 5-60](https://railroads.dot.gov/sites/fra.dot.gov/files/2026-07/2026.07.13_Penn%20Station%20SOS_Phase%20I%20Report_FINAL.pdf#page=71)).
  Its measurements would be the best source for this model if released.
- **NJT station directory (January 2022).**
  The only source found that shows every platform's VCEs after Moynihan Train Hall opened,
  as icons for stairs, escalators, and elevators labeled with their tracks,
  on both concourse levels.
  The ETA report links to it.
  It's a schematic wayfinding map, not to scale, so it has no widths,
  and may show some VCEs more than once
  (see [Counts from NJT's Station Directory](#counts-from-njts-station-directory)).
- **Amtrak's *Doubling Trans-Hudson Train Capacity* (October 2024).**
  Lists an "NJT Track 7/8 Escalator Replacement with Stairs"
  among nearby projects, with an expected completion in 2021,
  a change on platform 4 that may postdate the Master Plan draft.
- **Master Plan (August 2020 draft and April 2021 final).**
  The only source found with per-VCE widths,
  but only for the existing VCEs each alternative keeps,
  and its platform plans seem to stop short of the West End Concourse (see above).
  The final report's "Existing Number of VCEs" counts per platform
  are the only complete per-platform counts in a table,
  but also probably exclude the West End Concourse.
- **NJT's PCIP reports (2019 to 2021).**
  From records requests, archived by Liam Blank.
  Their drawings only dimension proposed platforms, concourses, and VCEs, not existing ones,
  but PCIP Phase 2's existing concourse-level plan (November 2020) is a scaled vector drawing
  of platforms 1 to 8, including the West End Concourse,
  from which existing stairs' widths can be [estimated](#estimated-widths).
- **Moynihan Station EA (2008 data).**
  Per-platform counts of stairs and escalators, capacities, and clearance times,
  but from before the West End Concourse's expansion and Moynihan Train Hall.
- **ETA's `total_vce_width`.**
  Its source is unknown.
  Platforms 10 and 11's match the EA's capacities divided by 17 pax/min/ft,
  but platform 3's matches the EA's platform 1, not platform 3.

## Results

Times are in m:ss.
Headways, time at capacity, and dwells are durations, and the rest are times after the first train arrives.

- **Platform:** the platform's number, with `(recon)` for its VCEs after Penn Reconstruction.
- **NFPA 130 evacuation:** how long the platform's peak occupants, including those aboard its trains,
  take to evacuate at NFPA 130's exit capacity, marked ✗ if it's over its 4-minute limit.
- **NFPA 130 to concourse:** how long the farthest of those occupants takes to reach the concourse,
  marked ✗ if it's over NFPA 130's 6-minute limit to reach a point of safety.
- **Arrivals:** when each train arrives, in order.
  A train arrives later than scheduled if the train before it on its track hasn't departed.
- **Dwell:** each train's dwell, in the same order:
  from its arrival until all of its arriving passengers have alighted
  and all of its departing passengers have boarded.
- **Taper time:** the last second more arriving passengers are on the platform than fit in the stair queues.
- **Clear time:** when the last arriving passenger leaves the platform.
- **Boarded time:** when the last departing passenger boards.
- **Time at capacity:** how long the upstairs rate is at LOS E capacity.
- **Max pax on platform** and **Max density:** the most crowded moment on the platform,
  with its LOS, graded by its space per passenger in sq ft.
- **Max up rate:** the highest upstairs rate.

The ETA report quotes the max up rate, time at capacity, and taper time
for platform 3 with trains 2 minutes apart:
12.68 pax/s for 0:13, tapering at 5:55, with Penn Reconstruction,
and 12.04 pax/s for 0:32, tapering at 6:04, without it.

### Current Results

This table is generated by `uv run platform-crowd-model run --update-readme`:

<!-- results-table:start -->
| Platform | Headway | VCE width | NFPA 130 evacuation | NFPA 130 to concourse | Arrivals | Dwell | Taper time | Clear time | Boarded time | Time at capacity | Max pax on platform | Max density (pax/m²) | Max up rate (pax/s) |
|---|---|---|---|---|---|---|---|---|---|---|---|---|---|
| 1 | 0:00 | 31'6" | 6:36 ✗ | 6:57 ✗ | 0:00, 0:00, 5:31, 5:31 | 5:31, 5:31, 0:46, 0:46 | 9:49 | 10:04 | 6:17 | 9:04 | 3206 | 3.04 (D) | 8.93 |
| 2 | 0:00 | 32'7" | 6:27 ✗ | 6:47 ✗ | 0:00, 0:00, 5:16, 5:16 | 5:16, 5:16, 0:46, 0:46 | 9:25 | 9:40 | 6:02 | 8:46 | 3233 | 3.26 (D) | 9.23 |
| 3 | 0:00 | 45'9" | 5:31 ✗ | 5:52 ✓ | 0:00, 0:00, 3:31, 3:31 | 3:31, 3:31, 0:44, 0:44 | 6:45 | 7:00 | 4:15 | 6:56 | 3828 | 3.02 (D) | 12.96 |
| 3 (recon) | 0:00 | 44'9" | 5:36 ✗ | 5:57 ✓ | 0:00, 0:00, 3:38, 3:38 | 3:38, 3:38, 0:44, 0:44 | 6:56 | 7:11 | 7:12 | 7:04 | 3803 | 3.00 (D) | 12.68 |
| 4 | 0:00 | 42'8" | 6:31 ✗ | 6:52 ✗ | 0:00, 0:00, 4:40, 4:40 | 4:40, 4:40, 0:43, 0:43 | 8:53 | 9:08 | 5:23 | 8:56 | 4292 | 2.74 (D) | 12.09 |
| 5 | 0:00 | 53'11" | 7:16 ✗ | 7:37 ✗ | 0:00, 0:00, 0:43, 0:43 | 0:43, 0:43, 0:43, 0:43 | 6:50 | 7:05 | 1:26 | 7:04 | 6104 | 3.17 (D) | 15.28 |
| 6 | 0:00 | 40'6" | 6:46 ✗ | 7:07 ✗ | 0:00, 0:00, 5:01, 5:01 | 5:01, 5:01, 0:43, 0:43 | 9:29 | 9:44 | 5:44 | 9:24 | 4238 | 2.38 (D) | 11.47 |
| 7 | 0:00 | 46'11" | 8:28 ✗ | 8:49 ✗ | 0:00, 0:00, 0:43, 0:43 | 0:43, 0:43, 0:43, 0:43 | 7:53 | 8:08 | 1:26 | 8:07 | 6256 | 3.28 (D) | 13.29 |
| 8 | 0:00 | 45'10" | 6:13 ✗ | 6:34 ✗ | 0:00, 0:00, 4:12, 4:12 | 4:12, 4:12, 0:43, 0:43 | 8:07 | 8:22 | 4:55 | 8:18 | 4370 | 3.05 (D) | 12.99 |
| 9 | 0:00 | 47' | 5:05 ✗ | 5:26 ✓ | 0:00, 0:44, 1:28, 2:12 | 0:44, 0:44, 0:44, 0:44 | 6:31 | 6:46 | 2:56 | 6:45 | 3589 | 2.69 (D) | 13.32 |
| 10 | 0:00 | 86'6" | 4:16 ✗ | 4:37 ✓ | 0:00, 0:00, 0:43, 0:43 | 0:43, 0:43, 0:43, 0:43 | 4:10 | 4:25 | 1:26 | 4:24 | 5393 | 2.21 (D) | 24.51 |
| 11 | 0:00 | 41'8" | 6:38 ✗ | 6:59 ✗ | 0:00, 0:00, 4:49, 4:49 | 4:49, 4:49, 0:43, 0:43 | 9:09 | 9:24 | 5:32 | 9:08 | 4267 | 3.83 (E) | 11.81 |
| 1 | 2:00 | 31'6" | 3:03 ✓ | 3:24 ✓ | 0:00, 2:00, 4:00, 9:02 | 0:46, 7:02, 5:02, 0:46 | 11:04 | 11:19 | 9:48 | 9:04 | 1320 | 1.25 (C) | 8.93 |
| 2 | 2:00 | 32'7" | 2:57 ✓ | 3:18 ✓ | 0:00, 2:00, 4:00, 8:44 | 0:46, 6:44, 4:44, 0:46 | 10:41 | 10:56 | 9:30 | 8:45 | 1310 | 1.32 (C) | 9.23 |
| 3 | 2:00 | 45'9" | 2:18 ✓ | 2:59 ✓ | 0:00, 2:00, 4:00, 6:44 | 0:44, 4:44, 2:44, 0:44 | 8:14 | 8:29 | 7:28 | 6:56 | 1322 | 1.04 (B) | 12.96 |
| 3 (recon) | 2:00 | 44'9" | 2:20 ✓ | 2:59 ✓ | 0:00, 2:00, 4:00, 6:53 | 0:44, 4:53, 2:53, 0:44 | 8:25 | 8:40 | 7:37 | 7:04 | 1332 | 1.05 (B) | 12.68 |
| 4 | 2:00 | 42'8" | 2:49 ✓ | 3:10 ✓ | 0:00, 2:00, 4:00, 8:21 | 0:43, 6:21, 4:21, 0:43 | 10:20 | 10:35 | 9:04 | 8:56 | 1621 | 1.04 (B) | 12.09 |
| 5 | 2:00 | 53'11" | 2:16 ✓ | 2:59 ✓ | 0:00, 2:00, 4:00, 6:36 | 0:43, 4:36, 2:36, 0:43 | 8:07 | 8:22 | 7:19 | 7:04 | 1516 | 0.79 (A) | 15.28 |
| 6 | 2:00 | 40'6" | 3:06 ✓ | 3:27 ✓ | 0:00, 2:00, 4:00, 8:47 | 0:43, 6:47, 4:47, 0:43 | 10:54 | 11:09 | 11:09 | 9:24 | 1716 | 0.96 (B) | 11.47 |
| 7 | 2:00 | 46'11" | 2:34 ✓ | 2:59 ✓ | 0:00, 2:00, 4:00, 7:36 | 0:43, 5:36, 3:36, 0:43 | 9:23 | 9:38 | 8:19 | 8:06 | 1581 | 0.83 (B) | 13.29 |
| 8 | 2:00 | 45'10" | 2:38 ✓ | 2:59 ✓ | 0:00, 2:00, 4:00, 7:46 | 0:43, 5:46, 3:46, 0:43 | 9:36 | 9:51 | 8:29 | 8:18 | 1591 | 1.11 (C) | 12.99 |
| 9 | 2:00 | 47' | 2:46 ✓ | 3:07 ✓ | 0:00, 2:00, 4:52, 6:00 | 0:44, 2:52, 0:44, 0:44 | 8:00 | 8:15 | 6:44 | 6:44 | 1742 | 1.30 (C) | 13.32 |
| 10 | 2:00 | 86'6" | 1:24 ✓ | 2:59 ✓ | 0:00, 2:00, 4:00, 6:00 | 0:43, 0:43, 0:43, 0:43 | 6:51 | 7:07 | 6:43 | 4:24 | 1211 | 0.50 (A) | 24.51 |
| 11 | 2:00 | 41'8" | 2:54 ✓ | 3:15 ✓ | 0:00, 2:00, 4:00, 8:33 | 0:43, 6:33, 4:33, 0:43 | 10:36 | 10:51 | 9:16 | 9:08 | 1630 | 1.47 (C) | 11.81 |
| 1 | 5:00 | 31'6" | 3:02 ✓ | 3:23 ✓ | 0:00, 5:00, 10:00, 15:00 | 0:46, 0:46, 0:46, 0:46 | 17:02 | 17:17 | 15:46 | 9:04 | 1312 | 1.24 (C) | 8.93 |
| 2 | 5:00 | 32'7" | 2:56 ✓ | 3:17 ✓ | 0:00, 5:00, 10:00, 15:00 | 0:46, 0:46, 0:46, 0:46 | 16:57 | 17:12 | 15:46 | 8:44 | 1301 | 1.31 (C) | 9.23 |
| 3 | 5:00 | 45'9" | 2:16 ✓ | 2:59 ✓ | 0:00, 5:00, 10:00, 15:00 | 0:44, 0:44, 0:44, 0:44 | 16:30 | 16:45 | 15:44 | 6:56 | 1309 | 1.03 (B) | 12.96 |
| 3 (recon) | 5:00 | 44'9" | 2:19 ✓ | 2:59 ✓ | 0:00, 5:00, 10:00, 15:00 | 0:44, 0:44, 0:44, 0:44 | 16:32 | 16:47 | 15:44 | 7:04 | 1319 | 1.04 (B) | 12.68 |
| 4 | 5:00 | 42'8" | 2:48 ✓ | 3:09 ✓ | 0:00, 5:00, 10:00, 15:00 | 0:43, 0:43, 0:43, 0:43 | 16:59 | 17:14 | 15:43 | 8:56 | 1609 | 1.03 (B) | 12.09 |
| 5 | 5:00 | 53'11" | 2:13 ✓ | 2:59 ✓ | 0:00, 5:00, 10:00, 15:00 | 0:43, 0:43, 0:43, 0:43 | 16:31 | 16:46 | 15:43 | 7:04 | 1501 | 0.78 (A) | 15.28 |
| 6 | 5:00 | 40'6" | 2:57 ✓ | 3:18 ✓ | 0:00, 5:00, 10:00, 15:00 | 0:43, 0:43, 0:43, 0:43 | 17:07 | 17:22 | 15:43 | 9:24 | 1630 | 0.91 (B) | 11.47 |
| 7 | 5:00 | 46'11" | 2:33 ✓ | 2:59 ✓ | 0:00, 5:00, 10:00, 15:00 | 0:43, 0:43, 0:43, 0:43 | 16:47 | 17:02 | 15:43 | 8:04 | 1568 | 0.82 (A) | 13.29 |
| 8 | 5:00 | 45'10" | 2:37 ✓ | 2:59 ✓ | 0:00, 5:00, 10:00, 15:00 | 0:43, 0:43, 0:43, 0:43 | 16:50 | 17:05 | 15:43 | 8:16 | 1578 | 1.10 (C) | 12.99 |
| 9 | 5:00 | 47' | 2:13 ✓ | 2:59 ✓ | 0:00, 5:00, 10:00, 15:00 | 0:44, 0:44, 0:44, 0:44 | 16:27 | 16:42 | 15:44 | 6:44 | 1297 | 0.97 (B) | 13.32 |
| 10 | 5:00 | 86'6" | 1:23 ✓ | 2:59 ✓ | 0:00, 5:00, 10:00, 15:00 | 0:43, 0:43, 0:43, 0:43 | 15:51 | 16:07 | 15:43 | 4:24 | 1187 | 0.49 (A) | 24.51 |
| 11 | 5:00 | 41'8" | 2:52 ✓ | 3:13 ✓ | 0:00, 5:00, 10:00, 15:00 | 0:43, 0:43, 0:43, 0:43 | 17:03 | 17:18 | 15:43 | 9:08 | 1619 | 1.45 (C) | 11.81 |
<!-- results-table:end -->

### History

As bugs were fixed and improvements made, the results table above was updated,
and each fix's effect is summarized here.

- **Original model**, as of the [`penn-station-can-handle-the-load`](https://github.com/effective-transit-alliance/platform-crowd-model/tree/penn-station-can-handle-the-load) tag the ETA report used:
  reproduces the report's platform 3 numbers to within a few seconds.
- **Fixed the stair equation's units:**
  Fruin's `P` is pax/min per foot of width, but was used as pax/s across all stairs,
  scaling stair flow by 60 / VCE width, e.g. overstating it by about 1.4x on platform 3.
  Platform 3 with 2-minute headways now peaks at 10.13 pax/s (10.47 with Penn Reconstruction)
  and never reaches capacity, tapering at 411 s (392 s) instead of 361 s (353 s).
- **Replaced the platform-density taper with a stair queue:**
  Fruin's equation relates flow to the space per passenger on the stair, not on the platform,
  so it held the stairs below capacity while the platform was crowded,
  and slowed them as it emptied, so they never fully cleared.
  Now queued stairs discharge at LOS E capacity until empty.
  Platform 3 with 2-minute headways is at capacity for 269 s (255 s with Penn Reconstruction),
  tapers at 254 s (241 s), and clears at 269 s (256 s), where it never cleared before.
  The platform peaks at 1521 passengers (1496) instead of 2274 (2219).
- **Stopped double counting the upward flow against downward flow:**
  each train's share of the stairs had the whole upward flow subtracted from it.
  This barely matters: 0.1 more departing passengers come down on platform 3.
- **Stopped flow in both directions on stairs at 10 pax/min/ft, not 12,**
  per the report and the LOS C/D boundary.
  No effect yet, since the concourse-density limit on downward flow always binds first.
- **Used Fruin's descending stair equation for passengers coming downstairs,**
  instead of the ascending one, giving about 15% more downward flow.
  On platform 3, 379.0 of 400 departing passengers come down within 600 s instead of 368.9,
  and on platform 10, all now board, by 548 s.
- **Put the stair LOS C/D boundary at 10 pax/min/ft, not 9.5,** per the TRB paper and the report.
  No effect on results.
- **Graded platform crowding with Fruin's queuing LOS instead of his walkway LOS,**
  since most passengers on a platform are standing and waiting.
  At its most crowded, platform 3 is now LOS C instead of E, and platform 6 is D instead of F.
- **Stopped boarding until everyone has alighted,** per the report,
  instead of boarding with the doors left over in the last second of alighting.
  Boarding starts 1 s later, so the platform peaks at up to 26 more passengers.
- **Replaced the concourse-density limit on downward flow with a queue:**
  departing passengers upstairs come down with whatever stair capacity the upward flow leaves,
  instead of slowing as a fixed 5,000 sq ft concourse empties.
  Every scenario now finishes boarding: platform 3 with 2-minute headways by 325 s
  (309 s with Penn Reconstruction), where about 20 of 400 never boarded before.
- **Let passengers come down at up to LOS E capacity, 17 pax/min/ft, when few are going up,**
  since the report's 10 pax/min/ft rule only applies to flow in both directions.
  Platform 3 with 2-minute headways finishes boarding by 302 s instead of 325 s
  (287 s instead of 309 s with Penn Reconstruction).
- **Gave each train a 12-car train's 48 doors instead of a 10-car train's 40,**
  to match its 1,620 passengers, a seated 12-car train.
  Passengers alight sooner, so the platform peaks at up to 96 more passengers,
  e.g. 1623 instead of 1538 on platform 3 with trains 2 minutes apart,
  and departing passengers finish boarding up to 8 s sooner.
  Platform 3's tracks only fit 10 cars, though, so its trains are still too long.
- **Had each train's departing passengers start coming down 2 minutes before it's scheduled,**
  all 400 from upstairs, like when its track is announced,
  instead of every train's being there from the start, 200 of them already on the platform.
  The simulation now starts at -2:00, when the first train's passengers start coming down.
  With 5-minute headways, every train's passengers are on the platform by the time it arrives:
  on platform 3, every dwell is 0:43, instead of 3:20 for the first train,
  and it peaks at 1611 passengers instead of 2411.
  With 2-minute headways, the second train's passengers can't come down
  until the first train's arriving passengers have cleared the stairs,
  so its dwell is 6:23 instead of 3:35 on platform 3, but the first train's is 0:43 instead of 5:35.
  On platform 6, all 4 trains are scheduled at once,
  so all 1,600 departing passengers come down at -2:00, and it peaks at 6229 passengers (LOS F).
- **Took each platform's area from its outline in OpenStreetMap,**
  instead of its width times its length, which were hardcoded with no stated source.
  Only densities change: with trains 2 minutes apart,
  platform 6 peaks at 0.88 pax/m² (LOS B) instead of 1.37 (C),
  platform 11 at 1.45 instead of 1.17, and platform 3 at 1.28 instead of 1.44.
- **Made each train as long as its platform and its tracks allow, up to 12 cars,**
  with 135 passengers and 4 doors per car.
  Only platform 3 changes, to 10 cars' 1,350 passengers and 40 doors:
  with trains 2 minutes apart, it clears at 9:08 instead of 10:38,
  and peaks at 1353 passengers instead of 1623.
- **Checked NFPA 130's 6-minute limit to reach a point of safety,** taken as the concourse:
  the farthest occupant's walk along the platform, 325' for now, wait at the VCEs, and climb.
  With trains 2 or 5 minutes apart, every platform reaches the concourse within 6 minutes,
  and with two trains arriving at once, only platforms 3 with Penn Reconstruction and 10 do.
- **Took each platform's VCEs from the station's plans and NJT's station directory,**
  in [`data/vces.csv`](./data/vces.csv), instead of the ETA report's total widths,
  and modeled every platform, 1 to 11.
  Escalators are still treated as stairs, and each platform's widest is left out.
  With trains 2 minutes apart, platform 3's VCEs are 45'9" wide instead of 42'6",
  so it clears at 8:29 instead of 9:08,
  platform 6's are 40'6" instead of 48'2", so it clears at 11:09 instead of 9:23,
  platform 10's are 86'6" instead of 70'7" (7:07 instead of 7:21),
  and platform 11's are 41'8" instead of 43'7" (10:51 instead of 10:22).
  Platform 1, with the least, 31'6", clears last, at 11:19.
  With two trains arriving at once, platform 3 now reaches the concourse within NFPA 130's 6 minutes, too.
- **Ran platform 9's trains all on its one track, 17,** instead of alternating between two.
  Each train arrives once the one before it has departed,
  so with trains scheduled at once, they arrive at 0:00, 0:44, 1:28, and 2:12,
  and the platform peaks at 3589 passengers instead of 5161,
  now reaching the concourse within NFPA 130's 6 minutes, in 5:26.
  With trains 2 minutes apart, the third arrives at 4:52 instead of 4:00,
  and the platform peaks at 1742 passengers instead of 1311.
