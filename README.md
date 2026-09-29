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
uv run platform-crowd-model
```

In doing so, `uv` will also install dependencies and set up a virtual environment.

Each run prints a table of headline results.
To also save each scenario's full time series as CSVs and its charts as an SVG in `output/`, run

```sh
uv run platform-crowd-model --charts
```

To also replace the [results table](#current-results) below with that run's, run

```sh
uv run platform-crowd-model --update-readme
```

The generated files in [`data/`](./data) come from each source's commands,
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

## How the Model Works

The model is the one ETA's report
[Penn Station Can Handle the Load](https://www.etany.org/penn-station-can-handle-the-load) used,
as of the [`penn-station-can-handle-the-load`](https://github.com/effective-transit-alliance/platform-crowd-model/tree/penn-station-can-handle-the-load) tag.
Later changes to the model aren't reflected in the report.
Since then, several bugs have been fixed,
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
   - Every train is as long as its platform's tracks and the platform itself allow, up to 12 cars,
     per [this track map](https://www.railfanguides.us/ny/penntonewrochelle/PennStationLayout1.jpg),
     recorded in [`data/platform_max_cars_track_map.csv`](./data/platform_max_cars_track_map.csv):
     10 cars on platform 3, and 12 on platforms 6, 10, and 11.
   - Every car is a full, seated NJ Transit MultiLevel car with 135 passengers,
     from the Moynihan Station environmental assessment,
     whose 12-car trains carry 1,620
     ([Table 4.4-10, p. 4.4-22](https://web.archive.org/web/20241011135133/https://cdn.esd.ny.gov/subsidiaries_projects/msdc/Data/NEPA/04_4%20StationPedCirculation.pdf#page=22);
     [Table 4.4-19, p. 4.4-47](https://web.archive.org/web/20241011135133/https://cdn.esd.ny.gov/subsidiaries_projects/msdc/Data/NEPA/04_4%20StationPedCirculation.pdf#page=47)).
   - Every car has 4 doors and is 85 ft long, the worst case.
     (An LIRR car has more doors, and LIRR platforms are generally wider.)
2. **Going upstairs.**
   Arriving passengers leave the platform
   via the vertical circulation elements (VCEs), i.e. the stairs and escalators.
   - Each VCE has its own queue, which it discharges at its capacity
     as long as anyone is queued
     ([TCQSM, p. 10-51](https://onlinepubs.trb.org/onlinepubs/tcrp/tcrp_rpt_165ch-10.pdf#page=55)).
   - A stair's capacity is LOS E capacity, 17 pax/min per foot of its width
     ([Fruin, p. 14](https://onlinepubs.trb.org/Onlinepubs/hrr/1971/355/355-001.pdf#page=14)).
   - An escalator's is the TCQSM's nominal capacity at 90 ft/min:
     34 pax/min with treads narrower than 32 in., and 72 pax/min with wider treads,
     since 32 in. treads carry close to 40 in. treads' capacity
     ([TCQSM, Exhibit 10-31, p. 10-52](https://onlinepubs.trb.org/onlinepubs/tcrp/tcrp_rpt_165ch-10.pdf#page=56)).
   - Each VCE's width and position are from
     [`data/vces.csv`](./data/vces.csv)
     (see [Estimated Widths](#estimated-widths)).
     - The train's doors are spread evenly along it,
       and it stops wherever on the platform the longest of the trains' dwells is shortest,
       trying every position a car length (85 ft) apart,
       then every position 5 ft apart within a car length of the best of those.
     - Each second, each door's alighting passengers walk to the quickest VCE:
       the one with the least walking time plus waiting time
       for everyone already queued or walking there.
     - They walk to the VCE's nearest end at 250 ft/min,
       the TCQSM's design walking speed
       ([p. 10-20](https://onlinepubs.trb.org/onlinepubs/tcrp/tcrp_rpt_165ch-10.pdf#page=24)).
     - Stairs go both ways, but escalators go one way.
       A platform's only escalator goes up.
       With more, the easternmost, toward 7th Avenue, goes up,
       and the westernmost, toward 8th Avenue, goes down,
       matching the AM peak, when most passengers are heading toward 7th Avenue.
       Any others go up while the platform is alighting,
       then reverse to go down once fewer than 10% of a train's arriving passengers
       are left on the platform or aboard trains that have arrived,
       and nobody is queued at or walking to them.
   - The report's "taper time" is when the remaining arriving passengers fit in the stair queues,
     20 ft of queue in front of the VCEs at 5 sq ft/pax
     (the TCQSM's stair queuing space).
     It's kept in the results table to compare with the report.
3. **Coming downstairs.**
   Departing passengers queue upstairs and come down to the platform.
   - The trains' passengers split each VCE's width
     in proportion to how many of each are still upstairs,
     and each train's share of a VCE carries the same share of its upward flow.
   - Nobody comes down a VCE while its upward flow is worse than LOS C, 10 pax/min/ft.
   - Each walks at 250 ft/min to the nearest of their train's cars,
     unless it's close to full, i.e. 90% of its 135 seats are boarded, waiting, or walking to it,
     in which case they go to the nearest car that isn't.
   - Otherwise, both directions share LOS E capacity, 17 pax/min/ft,
     so passengers come down with whatever their share of the upward flow leaves of that.
4. **Boarding.**
   Departing passengers on the platform board a train
   once every arriving passenger has alighted from it,
   at 1 pax/s per single-door equivalent.
   Each car boards only its own waiting passengers, through its own 4 doors.

### Crowding

- **On the platform**, the space per passenger is the usable platform area
  (75% of the platform's area) divided by everyone on the platform.
  It's graded with Fruin's LOS for queuing and waiting areas
  (A > 13, B > 10, C > 7, D > 3, E > 2 sq ft/pax, or else F;
  [TCQSM, Exhibit 10-32, p. 10-55](https://onlinepubs.trb.org/onlinepubs/tcrp/tcrp_rpt_165ch-10.pdf#page=59)),
  since most passengers on the platform are waiting, either to board or in the stair queues.
- **On the stairs**, the upward flow on the most crowded VCE is graded with Fruin's stair LOS
  (A ≤ 5, B ≤ 7, C ≤ 10, D ≤ 13, E ≤ 17 pax/min/ft, or else F;
  [Fruin, pp. 12–14](https://onlinepubs.trb.org/Onlinepubs/hrr/1971/355/355-001.pdf#page=12)).

### Scenarios

- Each train has 400 departing passengers, all waiting upstairs,
  who start coming down 2 minutes before it's scheduled.
- On every platform, 4 trains are scheduled 0, 2, or 5 minutes apart,
  alternating between the platform's two tracks, or all on platform 9's one,
  so each train after the first on each track can only arrive
  once the train before it on its track has departed at the end of its dwell.
  0 minutes apart, the first two arrive together and the last two as soon as they depart.

### NFPA 130 Evacuation

NFPA 130 requires enough exit capacity to evacuate a platform's occupants,
including those on its trains, in 4 minutes or less
([TCQSM p. 10-3](https://onlinepubs.trb.org/onlinepubs/tcrp/tcrp_rpt_165ch-10.pdf#page=7)).
The model checks this at the moment the most passengers are on the platform or aboard its trains,
counting each train from its arrival until it departs.

- Stairs and stopped escalators carry 1.41 pax/min per inch of width, i.e. 16.92 pax/min/ft
  ([TCQSM p. 10-51](https://onlinepubs.trb.org/onlinepubs/tcrp/tcrp_rpt_165ch-10.pdf#page=55)).
- The widest escalator is out of service,
  and escalators provide at most half of the exit capacity
  ([TCQSM p. 10-52](https://onlinepubs.trb.org/onlinepubs/tcrp/tcrp_rpt_165ch-10.pdf#page=56)).
- It ignores walking time to the VCEs, and NFPA 130's separate 6-minute limit
  to reach a point of safety, so it's a lower bound.

## Assumptions

Beyond the numbers above, which are all in `Assumptions`, the model assumes the following.
Each is marked by which way it biases the results:
**optimistic** (less crowding than reality), **pessimistic** (more),
**neutral** (neither, e.g. it only simplifies the bookkeeping),
or **unclear** (it could go either way).

### Platform

- **Neutral:** The platform is a single island platform serving two tracks,
  except platform 9, which serves one.
- **Optimistic:** 75% of its area is usable, the rest taken by columns, stairs, and other obstructions
  (see [Usable Platform Area](#usable-platform-area)).
- **Optimistic:** Passengers are spread evenly over the whole usable area,
  so local crowding, e.g. at the foot of the stairs or at the doors, isn't modeled.
- **Neutral:** Space per passenger counts everyone on the platform:
  arriving passengers heading up and departing passengers waiting to board.

### Stairs and Escalators

- **Unclear:** Escalators' capacities are the TCQSM's nominal ones at 90 ft/min,
  but their speeds aren't known,
  and their tread widths are measured from a drawing, so they may be off by a few inches,
  which matters at the 32 in. boundary between 34 and 72 pax/min.
- **Unclear:** Which escalators run which way isn't known,
  so they follow the rules above.
  Which escalator goes down is based on AM peak demand toward 7th Avenue,
  and when the extra ones reverse isn't from any source.
- **Unclear:** Most VCEs' widths are estimated from a drawing,
  or on platforms 9 to 11, their positions from a schematic map,
  and some of their VCEs may be missing (pessimistic)
  (see [Estimated Widths](#estimated-widths)).
- **Optimistic:** Arriving passengers know which VCE is quickest,
  with no preference for any exit, e.g. toward 7th Avenue
  (see [Passengers Only Prefer the Quickest VCE](#passengers-only-prefer-the-quickest-vce)).
- **Optimistic:** Trains stop at the best position on the platform for their dwells,
  with their doors spread evenly along their length.
- **Unclear:** Stair capacity is linear in width,
  though the TCQSM notes capacity is really stepped by the number of pedestrian lanes
  ([p. 10-49](https://onlinepubs.trb.org/onlinepubs/tcrp/tcrp_rpt_165ch-10.pdf#page=53)).
- **Optimistic:** Stair capacity doesn't depend on the stair's rise,
  though long climbs slow people down ([p. 10-51](https://onlinepubs.trb.org/onlinepubs/tcrp/tcrp_rpt_165ch-10.pdf#page=55)),
  or on luggage, strollers, or wheelchairs.
- **Optimistic:** Arriving passengers walk to the VCEs at 250 ft/min,
  though the TCQSM notes people walk slower in crowds with less than 25 sq ft/pax
  ([Exhibit 10-10, p. 10-21](https://onlinepubs.trb.org/onlinepubs/tcrp/tcrp_rpt_165ch-10.pdf#page=25)).
- **Unclear:** Departing passengers coming downstairs
  walk to the nearest car until it's close to full, at 90% of its seats,
  which isn't from any source.
- **Optimistic:** The concourse upstairs never backs up, so the stairs always discharge.

### Trains

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
- **Optimistic:** The NFPA 130 evacuation time ignores walking time to the VCEs,
  and the 6-minute limit to reach a point of safety.

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
This model clears platform 3's arriving passengers in 5:10,
so it's likely still optimistic,
though the two aren't exactly comparable:
7.9 minutes is the worst case across all platforms and simulation runs.

The FRA attributes long clearance times to
queues at the base of VCEs, uneven use of VCEs, and platform clutter
reducing the usable width
([p. 3-33](https://railroads.dot.gov/sites/fra.dot.gov/files/2026-07/2026.07.13_Penn%20Station%20SOS_Phase%20I%20Report_FINAL.pdf#page=44)).
This model only captures the first two.

### Passengers Only Prefer the Quickest VCE

Each VCE has its own queue, with walking distances to it,
so stairs near the ends of the platform or far from the busiest doors can run dry
while others still have a queue.
But arriving passengers choose perfectly:
each second, each door's passengers know every VCE's queue, even hundreds of feet away,
and walk to the one they can go up soonest,
so every VCE going up is used, and their waits even out across the platform (optimistic).
Before each VCE had its own queue, e.g. on platform 3 with 2-minute headways,
one pooled queue cleared 44 s sooner.
Always walking to the nearest VCE is the opposite extreme (pessimistic),
and real passengers are somewhere in between:

- **They only see nearby queues.**
  From a door, passengers can see the nearest few VCEs,
  not one 500 ft down a crowded platform.
- **They don't all make the same choice.**
  Given similar options, people split between them unevenly,
  rather than all taking the best one each second.
- **They head for a destination.**
  Most are heading toward 7th Avenue, as the ETA report notes,
  and the West End Concourse leads toward 8th Avenue and Moynihan Train Hall,
  so many take a longer wait on a VCE toward where they're going,
  concentrating queues at fewer VCEs.
- **Regulars position themselves.**
  Commuters ride in the car nearest their usual exit,
  so arriving passengers aren't spread evenly across the doors.
- **They switch queues.**
  Some leave a queue that isn't moving for another,
  while the model's passengers stay with the VCE they first chose.
- **They only see who's queued.**
  The model counts passengers still walking to a VCE as queued ahead of them,
  which passengers can't see.

These could be modeled later, from simplest to most involved:

- **A visibility radius:** passengers only consider VCEs within some distance, or the nearest few,
  and take the quickest of those.
- **Logit choice:** passengers split across VCEs
  with probability proportional to e^(−θ × each VCE's time to go up),
  the standard approach to route choice in pedestrian and transit models.
  θ = 0 splits them evenly, and a large θ is the current model.
  Since the model already tracks fractional passengers,
  each second's alighting passengers can be split by those probabilities.
- **Destination preference:** adding each VCE's walk upstairs toward each destination,
  with a share of the passengers heading to each, e.g. 7th Avenue or the West End Concourse,
  which needs a source for that split.
- **Uneven doors:** more of each train's passengers at the doors nearest the busiest exits.
- **Queue switching:** each second, passengers queued at one VCE move to another
  if it's become much quicker.

Penn could also bring passengers closer to the current model's perfect choices,
e.g. with screens showing each VCE's queue in real time,
staff directing passengers,
or blocking off paths to some VCEs to spread passengers out.

### Platform Crowding Is Graded Against the Whole Platform

Though the model now grades the platform with the queuing LOS,
the TCQSM's platform sizing procedure ([p. 10-56](https://onlinepubs.trb.org/onlinepubs/tcrp/tcrp_rpt_165ch-10.pdf#page=60))
only applies the queuing LOS to passengers *waiting* to board, and adds separate areas for:

- Walkway width for arriving passengers, graded as a walkway
- Queue storage at the stairs
- An 18 in. buffer along each platform edge

The model lumps everyone into one area,
which is optimistic with the queuing LOS,
especially while arriving passengers are walking to the stairs.

### Usable Platform Area

The model uses 75% of the platform's area,
leaving 25% for columns, stairwells, and other obstructions.
But the TCQSM's 18 in. edge buffers alone take 3 ft of an 18 ft platform, about 17%,
leaving only 8% for everything else, which is likely too little on a narrow platform
with stairwells in it.
So the usable area, and the space per passenger, are likely overstated (optimistic).

### Total VCE Width

The ETA report's total VCE widths, whose source is unknown,
are only used for how much Penn Reconstruction widens platform 3's VCEs, 2.25 ft.
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

`uv run platform-crowd-model data master-plan-vces`
([`master_plan_vces.py`](./src/platform_crowd_model/master_plan_vces.py))
extracts each VCE's approximate position from the draft's platform-level plans
([p. 17](https://liamblank.com/wp-content/uploads/2026/07/R-Master-Plan-MTA-NJT-AMT-20_0814-PSMP-Alternatives-Report.pdf#page=23), [p. 29](https://liamblank.com/wp-content/uploads/2026/07/R-Master-Plan-MTA-NJT-AMT-20_0814-PSMP-Alternatives-Report.pdf#page=35),
[p. 41](https://liamblank.com/wp-content/uploads/2026/07/R-Master-Plan-MTA-NJT-AMT-20_0814-PSMP-Alternatives-Report.pdf#page=47), and [p. 53](https://liamblank.com/wp-content/uploads/2026/07/R-Master-Plan-MTA-NJT-AMT-20_0814-PSMP-Alternatives-Report.pdf#page=59)).
They're vector drawings with each VCE drawn as a rectangle color-coded by type and status,
and a scale bar to convert to feet.
The script matches the `n`th VCE from the west on each platform to the `n`th VCE in its table,
since the plans are too small to measure widths from,
and writes [`data/vce_positions_master_plan.csv`](./data/vce_positions_master_plan.csv),
with positions in feet east of the plans' west edge,
which cuts across every platform at the same place, under the West End Concourse.
The plans have west on the left:
the platforms' east ends on them match those on a scaled existing-conditions plan
in NJ Transit's PCIP Phase 2 drawings to within about 4 ft
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
The January 2022 NJ Transit station directory shows 8 stairs and escalators to platform 3,
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
  E.g. platform 3's VCEs sum to 412 to 458 in. across the alternatives, even with their new VCEs,
  but the ETA report used 42.5 ft (510 in.) today,
  and platform 6's sum to 437 to 458 in., but the ETA report used 48.168 ft (578 in.).
  The source of the ETA report's widths is unknown.
  Platform 3's 5 existing VCEs in the Master Plan total 233 in. (19.4 ft):
  3 stairs (165 in.) and 2 escalators (68 in.).
  At 17 pax/min/ft for the stairs and typical escalator capacities,
  that's roughly the Moynihan Station EA's 437 pax/min for platform 3 in 2008,
  while the ETA report's 42.5 ft is 722 pax/min, the EA's figure for platform 1.
  But the Master Plan probably doesn't include the West End Concourse's VCEs,
  and the EA's data predates Moynihan Train Hall,
  so platform 3's total width today is probably more than 19.4 ft.
  With the VCEs on NJ Transit's scaled PCIP Phase 2 plan, whose widths are mostly
  [estimated](#estimated-widths), it's about 550 in. (45.8 ft), a little more than the ETA report's 42.5 ft,
  or 516 in. (43 ft) with the Master Plan's widths where it has them, as the model now uses.
  Either way, a single total overstates how quickly a platform clears
  if some of that width is at its far west end, far from most of the train's doors.
- **Positions are approximate.**
  They're scaled from small drawings, so they're only accurate to within about 10 ft.

The FRA's Service Optimization Study measured today's VCEs on site,
but didn't publish the measurements.

### Counts from NJ Transit's Station Directory

NJ Transit's [January 2022 station directory](https://content.njtransit.com/sites/default/files/NY%20Penn%20Station%20Directory_011022.pdf),
which the ETA report links to,
is the only source found that shows every platform's VCEs after Moynihan Train Hall opened.
It's a vector wayfinding map of both concourse levels,
with an icon for each stair, escalator, and elevator to a platform,
labeled with the platform's tracks.
`uv run platform-crowd-model data directory-vces`
([`directory_vces.py`](./src/platform_crowd_model/directory_vces.py))
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
NJ Transit's PCIP Phase 2 drawings include an
[existing concourse-level plan](https://liamblank.com/wp-content/uploads/2026/09/pcip-2-conceptual-design-preliminary-drawings.pdf#page=45)
(sheet A-001, November 2020), a vector drawing at 1" = 40' of platforms 1 to 8,
including the West End Concourse, with each stair's and escalator's treads drawn as lines.
No such drawing of platforms 9 to 11 was found.
`uv run platform-crowd-model data estimated-vces`
([`estimated_vces.py`](./src/platform_crowd_model/estimated_vces.py))
finds every run of treads on the platforms, merges a stair's flights,
and measures each VCE's width and position into
[`data/vces.csv`](./data/vces.csv):

- Positions are in the Master Plan's frame, feet east of its plans' west edge.
  The two drawings register to within about 1.2 ft:
  the platforms' east ends on them are all the same distance apart.
  Each platform's east end in that frame, for the model's trains to stop against,
  is in [`data/platform_east_ends.csv`](./data/platform_east_ends.csv):
  the sheet's on platforms 1 to 8, and the Master Plan's on platforms 9 to 11.
- VCEs matching a Master Plan VCE of the same type within 15 ft
  have `width_status` `master plan`, with the Master Plan's width.
  The rest, 45 of 61, are `estimated`, with only the sheet's width.
- For the 10 matched stairs, the sheet's widths differ from the Master Plan's by up to 14 in.
  (a median of 4 in.), and for the 6 matched escalators, by up to 9 in.
- Escalators' treads are their steps, narrower than their balustrades.
- The platforms' labels hide what's under them,
  including an escalator the Master Plan has about 230 ft along each of platforms 3 to 8.

| Platform | Stairs on the sheet | Escalators on the sheet | Their total width | Of which estimated | Master Plan VCEs not matched on the sheet |
|---|---|---|---|---|---|
| 1 | 6 | 2 | 420 in. (35.0 ft) | 198 in. | 44/57 in. stair at 689 ft; 46 in. escalator at 774 ft; 52 in. stair at 776 ft |
| 2 | 5 | 4 | 431 in. (35.9 ft) | 276 in. | 52 in. stair at 775 ft; 46 in. escalator at 774 ft |
| 3 | 6 | 2 | 516 in. (43.0 ft) | 430 in. | 34 in. escalator at 230 ft; 69 in. stair at 278 ft; 44 in. stair at 684 ft |
| 4 | 5 | 2 | 441 in. (36.8 ft) | 369 in. | 34 in. escalator at 230 ft; 34 in. escalator at 404 ft; 44 in. stair at 683 ft |
| 5 | 4 | 4 | 438 in. (36.5 ft) | 404 in. | 34 in. escalator at 230 ft; 44 in. stair at 683 ft |
| 6 | 3 | 4 | 370 in. (30.8 ft) | 286 in. | 34 in. escalator at 230 ft |
| 7 | 4 | 3 | 393 in. (32.8 ft) | 301 in. | 34 in. escalator at 230 ft |
| 8 | 5 | 2 | 435 in. (36.2 ft) | 377 in. | 34 in. escalator at 230 ft; 34 in. escalator at 404 ft |

The unmatched Master Plan VCEs are a mix:
the escalators at 230 ft are hidden under the sheet's labels, so they should be added;
others, like platform 3's 44 in. stair at 684 ft,
are probably the same VCEs as similar ones on the sheet 15 to 25 ft away;
and some disagree on the type, like platform 3's 69 in. stair at 278 ft,
where the sheet has a 37 in. escalator.
So platform 3's VCEs total about 550 in. (45.8 ft) with its hidden escalator,
a little more than the model's 42.5 ft,
including the 2 West End Concourse stairs (107 and 120 in.) and the Exit Concourse's 2 (61 and 60 in.).

These are only estimates:

- A drawn tread isn't necessarily the clear width between handrails.
- The sheet's existing stairs aren't labeled,
  so some may be stairs between the upper and lower concourses drawn over a platform,
  though every platform's matches the directory's pattern,
  e.g. 2 stairs each for the West End and Exit Concourses.
- The sheet predates NJ Transit's replacement of an escalator on tracks 7/8 (platform 4)
  with stairs in about 2021.
- The model gives an escalator the TCQSM's capacity for its tread width,
  so a few inches' error can move it across 32 in., between 34 and 72 pax/min.

Platforms 9 to 11 aren't on that plan,
so their VCEs are estimated from NJ Transit's January 2022 station directory instead,
whose map is schematic and not to scale.
Each level of its map is calibrated to feet by matching its icons on platforms 1 to 8,
and the Master Plan's on platforms 9 and 10, to their positions,
to within about 20 ft on average.
The directory doesn't show every VCE, so the Master Plan's existing VCEs it doesn't show are added,
though they're from before Moynihan Train Hall opened.
Each has the Master Plan's width where it has one,
or else the median width of that type on platforms 1 to 8, marked `typical`.

### Field Survey

Since no public source has every VCE's width,
[`data/vces_field_survey.csv`](./data/vces_field_survey.csv) is a sheet for measuring them in person,
made by `uv run platform-crowd-model data field-survey`
([`field_survey.py`](./src/platform_crowd_model/field_survey.py)).
It lists each platform's VCEs expected from the PCIP Phase 2 plan (platforms 1 to 8),
the 2022 directory, and the Master Plan,
each sorted west to east, since they can't all be aligned reliably,
starting with platform 3, the ETA report's focus,
then platform 11, which has no width data.
Surveyors fill in the columns after `position_ft`, feet east of the Master Plan's plans' west edge:

- `found`: yes, no, or the `vce_name` of another row it duplicates.
  Add rows for VCEs neither source lists.
- `clear_width_in`: the width between the handrails at the platform end,
  which the capacity standards use.
- `escalator_step_width_in` and `escalator_direction_am_peak` and `_pm_peak`.
- `nearest_column_number` and `distance_from_east_end_ft`:
  platform columns' painted numbers give precise positions.
- `leads_to`: the concourse, which also shows whether the directory shows a VCE twice.
- `obstructions`: columns, benches, bins, or narrow landings near the bottom,
  which the FRA found also slow clearing.
- `photos` and `notes`.

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
| [NJ Transit station directory](https://content.njtransit.com/sites/default/files/NY%20Penn%20Station%20Directory_011022.pdf) | January 2022 | Yes | Yes | No | Schematic | Yes |
| [Master Plan final report](https://esd.ny.gov/sites/default/files/CACWG-Meetings-8-9-Q-A-09-08-21.pdf#page=7), in the Empire Station Complex Q&A | April 2021 | Yes | Only of kept VCEs | Only of kept VCEs | No | Probably not |
| [Master Plan draft Alternatives Report](https://liamblank.com/wp-content/uploads/2026/07/R-Master-Plan-MTA-NJT-AMT-20_0814-PSMP-Alternatives-Report.pdf#page=22) | August 2020 | No | Only of kept VCEs | Only of kept VCEs | Yes, from vector plans | Probably not |
| [NJ Transit PCIP Phase 1 and 2 reports and drawings](https://liamblank.com/wp-content/uploads/2026/09/penn-station-gateway-catalogue.csv) | 2019 to 2021 | Only of new VCEs | Only of new VCEs | Only of new VCEs, but existing ones can be estimated from a scaled plan | Only of new VCEs, but existing ones can be estimated from a scaled plan | Yes, on a scaled plan |
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
- **NJ Transit station directory (January 2022).**
  The only source found that shows every platform's VCEs after Moynihan Train Hall opened,
  as icons for stairs, escalators, and elevators labeled with their tracks,
  on both concourse levels.
  The ETA report links to it.
  It's a schematic wayfinding map, not to scale, so it has no widths,
  and may show some VCEs more than once
  (see [Counts from NJ Transit's Station Directory](#counts-from-nj-transits-station-directory)).
- **Amtrak's *Doubling Trans-Hudson Train Capacity* (October 2024).**
  Lists an "NJ Transit Track 7/8 Escalator Replacement with Stairs"
  among nearby projects, with an expected completion in 2021,
  a change on platform 4 that may postdate the Master Plan draft.
- **Master Plan (August 2020 draft and April 2021 final).**
  The only source found with per-VCE widths,
  but only for the existing VCEs each alternative keeps,
  and its platform plans seem to stop short of the West End Concourse (see above).
  The final report's "Existing Number of VCEs" counts per platform
  are the only complete per-platform counts in a table,
  but also probably exclude the West End Concourse.
- **NJ Transit's PCIP reports (2019 to 2021).**
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

### Platform Dimensions

Each platform's length is from the Moynihan Station EA's Table 4.4-10
([p. 4.4-22](https://web.archive.org/web/20241011135133/https://cdn.esd.ny.gov/subsidiaries_projects/msdc/Data/NEPA/04_4%20StationPedCirculation.pdf#page=22)),
in [`data/platform_lengths_moynihan_ea.csv`](./data/platform_lengths_moynihan_ea.csv),
the only source found with written platform lengths.
Its area is from its outline in OpenStreetMap, which OpenRailwayMap draws,
in [`data/platforms_osm.csv`](./data/platforms_osm.csv),
written by `uv run platform-crowd-model data osm-platforms`
([`osm_platforms.py`](./src/platform_crowd_model/osm_platforms.py)),
since platforms taper toward their ends.
The outlines have no source, but agree with the PCIP Phase 2 existing plan's widths
to within about 2 ft, and with the EA's lengths to within about 60 ft,
except platform 9's, which is 178 ft longer.
The PCIP Phase 2 and Master Plan drawings cut the platforms off at their west ends,
so they can't give lengths or areas.

Platform 11 is 1,007 ft long, a few feet short of a 12-car train at 85 ft per car,
but the EA has it take 12-car LIRR trains,
so trains can overhang their platform's west end by up to 15 ft.

## Results

Times are in m:ss.
Headways, time at capacity, and dwells are durations, and the rest are times after the first train arrives.

- **Platform:** the platform's number, with `(recon)` for its VCEs after Penn Reconstruction.
- **Arrivals:** when each train arrives, in order.
  A train arrives later than scheduled if the train before it on its track hasn't departed.
- **Dwell:** each train's dwell, in the same order:
  from its arrival until all of its arriving passengers have alighted
  and all of its departing passengers have boarded.
- **Taper time:** the last second more arriving passengers are on the platform than fit in the stair queues.
- **Clear time:** when the last arriving passenger leaves the platform.
- **Boarded time:** when the last departing passenger boards.
- **Time at capacity:** how long the upstairs rate is at LOS E capacity.
- **Max up rate:** the highest upstairs rate.
- **Max pax on platform** and **Max density:** the most crowded moment on the platform,
  with its LOS, graded by its space per passenger in sq ft.
- **NFPA 130 evacuation:** how long the platform's peak occupants, including those aboard its trains,
  take to evacuate at NFPA 130's exit capacity, marked ✗ if it's over its 4-minute limit.

The ETA report quotes the max up rate, time at capacity, and taper time
for platform 3 with trains 2 minutes apart:
12.68 pax/s for 0:13, tapering at 5:55, with Penn Reconstruction,
and 12.04 pax/s for 0:32, tapering at 6:04, without it.

### Current Results

This table is generated by `uv run platform-crowd-model --update-readme`.
Platforms 9 to 11's VCEs are only estimated from NJ Transit's station directory
(see [Estimated Widths](#estimated-widths)).
The directory doesn't show every VCE, and the Master Plan has none on platform 11,
so its VCEs total only 27.9 ft, much less than the ETA report's 43.58 ft,
which likely makes it look much worse than it is.
Penn Reconstruction's new VCEs on platform 3 aren't known,
so they're one stair in the middle of the platform,
as wide as the ETA report's total VCE width with it is wider than without it, 2.25 ft.


<!-- results-table:start -->
| Platform | Headway | VCE width | Arrivals | Dwell | Taper time | Clear time | Boarded time | Time at capacity | Max up rate (pax/s) | Max pax on platform | Max density (pax/m²) | NFPA 130 evacuation |
|---|---|---|---|---|---|---|---|---|---|---|---|---|
| 1 | 0:00 | 35 ft | 0:00, 0:00, 4:34, 4:34 | 4:34, 4:34, 1:04, 1:04 | 8:54 | 9:39 | 5:38 | 7:55 | 9.25 | 3284 | 3.12 (D) | 6:44 ✗ |
| 2 | 0:00 | 35.92 ft | 0:00, 0:00, 3:51, 3:51 | 3:51, 3:51, 1:05, 1:05 | 8:12 | 9:06 | 4:56 | 6:09 | 10.47 | 3386 | 3.42 (D) | 6:39 ✗ |
| 3 | 0:00 | 43 ft | 0:00, 0:00, 2:21, 2:21 | 2:21, 2:21, 0:56, 0:56 | 7:29 | 8:02 | 3:17 | 7:19 | 11.71 | 4163 | 3.28 (D) | 6:46 ✗ |
| 3 (recon) | 0:00 | 45.25 ft | 0:00, 0:00, 1:15, 1:15 | 1:15, 1:15, 0:56, 0:56 | 7:05 | 7:55 | 2:11 | 6:50 | 12.34 | 4869 | 3.84 (E) | 7:25 ✗ |
| 4 | 0:00 | 36.75 ft | 0:00, 0:00, 4:57, 4:57 | 4:57, 4:57, 1:05, 1:05 | 11:21 | 12:09 | 6:02 | 9:44 | 9.96 | 4162 | 2.66 (D) | 7:46 ✗ |
| 5 | 0:00 | 36.5 ft | 0:00, 0:00, 2:27, 2:27 | 2:27, 2:27, 1:05, 1:05 | 11:21 | 12:09 | 3:32 | 7:26 | 10.75 | 5679 | 2.95 (D) | 10:39 ✗ |
| 6 | 0:00 | 30.83 ft | 0:00, 0:00, 6:08, 6:08 | 6:08, 6:08, 1:05, 1:05 | 12:11 | 12:49 | 7:13 | 10:44 | 9.01 | 4073 | 2.28 (D) | 9:17 ✗ |
| 7 | 0:00 | 32.75 ft | 0:00, 0:00, 3:42, 3:42 | 3:42, 3:42, 1:05, 1:05 | 13:27 | 14:09 | 4:47 | 10:05 | 8.80 | 5301 | 2.78 (D) | 11:14 ✗ |
| 8 | 0:00 | 36.25 ft | 0:00, 0:00, 5:04, 5:04 | 5:04, 5:04, 1:05, 1:05 | 11:55 | 12:42 | 6:09 | 9:37 | 9.68 | 4151 | 2.89 (D) | 7:57 ✗ |
| 9 | 0:00 | 40.08 ft | 0:00, 2:27, 3:30, 4:33 | 2:27, 1:03, 1:03, 1:03 | 8:03 | 8:40 | 5:36 | 6:26 | 11.73 | 2598 | 1.95 (D) | 4:46 ✗ |
| 10 | 0:00 | 64.5 ft | 0:00, 0:00, 0:53, 0:53 | 0:53, 0:53, 0:53, 0:53 | 5:55 | 6:40 | 1:46 | 5:22 | 17.77 | 5784 | 2.37 (D) | 6:10 ✗ |
| 11 | 0:00 | 27.92 ft | 0:00, 0:00, 8:21, 8:21 | 8:21, 8:21, 1:05, 1:05 | 14:40 | 15:18 | 9:26 | 12:14 | 8.28 | 3823 | 3.44 (D) | 9:42 ✗ |
| 1 | 2:00 | 35 ft | 0:00, 2:00, 4:00, 7:58 | 0:50, 5:58, 3:59, 1:05 | 9:57 | 10:25 | 9:03 | 7:41 | 9.25 | 1365 | 1.30 (C) | 3:50 ✓ |
| 2 | 2:00 | 35.92 ft | 0:00, 2:00, 4:00, 7:21 | 1:01, 5:21, 3:22, 1:04 | 9:08 | 9:56 | 8:25 | 6:17 | 10.47 | 1322 | 1.33 (C) | 3:38 ✓ |
| 3 | 2:00 | 43 ft | 0:00, 2:00, 4:00, 6:50 | 0:58, 4:50, 2:52, 1:02 | 8:35 | 9:06 | 7:52 | 5:44 | 11.71 | 1429 | 1.13 (C) | 3:11 ✓ |
| 3 (recon) | 2:00 | 45.25 ft | 0:00, 2:00, 4:00, 6:32 | 0:57, 4:32, 2:32, 1:03 | 8:10 | 8:56 | 7:35 | 5:30 | 12.34 | 1403 | 1.11 (C) | 2:43 ✓ |
| 4 | 2:00 | 36.75 ft | 0:00, 2:00, 4:00, 9:19 | 1:03, 7:19, 5:21, 1:05 | 12:06 | 12:54 | 10:24 | 9:20 | 9.96 | 2270 | 1.45 (C) | 4:53 ✗ |
| 5 | 2:00 | 36.5 ft | 0:00, 2:00, 4:00, 8:52 | 1:01, 6:52, 4:50, 1:05 | 11:21 | 12:06 | 9:57 | 8:42 | 10.75 | 2053 | 1.07 (B) | 4:35 ✗ |
| 6 | 2:00 | 30.83 ft | 0:00, 2:00, 4:00, 10:13 | 1:05, 8:13, 6:11, 1:05 | 13:01 | 13:30 | 11:18 | 11:05 | 9.01 | 2530 | 1.42 (C) | 6:28 ✗ |
| 7 | 2:00 | 32.75 ft | 0:00, 2:00, 4:00, 10:18 | 1:05, 8:18, 6:18, 1:05 | 13:10 | 13:39 | 11:23 | 11:20 | 8.80 | 2585 | 1.35 (C) | 6:08 ✗ |
| 8 | 2:00 | 36.25 ft | 0:00, 2:00, 4:00, 9:31 | 1:04, 7:31, 5:32, 1:05 | 13:10 | 13:57 | 10:36 | 8:00 | 9.68 | 2433 | 1.70 (D) | 5:19 ✗ |
| 9 | 2:00 | 40.08 ft | 0:00, 2:00, 5:44, 6:47 | 1:01, 3:44, 1:03, 1:03 | 9:25 | 10:01 | 7:50 | 6:02 | 11.73 | 1987 | 1.49 (C) | 3:48 ✓ |
| 10 | 2:00 | 64.5 ft | 0:00, 2:00, 4:00, 6:00 | 0:53, 1:05, 1:05, 1:05 | 7:26 | 8:17 | 7:05 | 2:10 | 17.77 | 1648 | 0.68 (A) | 2:00 ✓ |
| 11 | 2:00 | 27.92 ft | 0:00, 2:00, 4:00, 12:40 | 1:05, 10:40, 8:40, 1:05 | 15:45 | 16:18 | 13:45 | 12:09 | 8.28 | 2601 | 2.34 (D) | 6:50 ✗ |
| 1 | 5:00 | 35 ft | 0:00, 5:00, 10:00, 15:00 | 0:50, 0:50, 0:50, 0:50 | 16:59 | 17:27 | 15:50 | 7:04 | 9.25 | 1311 | 1.24 (C) | 3:04 ✓ |
| 2 | 5:00 | 35.92 ft | 0:00, 5:00, 10:00, 15:00 | 0:57, 0:57, 0:57, 0:57 | 16:49 | 17:39 | 15:57 | 6:04 | 10.47 | 1268 | 1.28 (C) | 2:54 ✓ |
| 3 | 5:00 | 43 ft | 0:00, 5:00, 10:00, 15:00 | 0:58, 0:58, 0:58, 0:58 | 16:45 | 17:16 | 15:58 | 5:32 | 11.71 | 1368 | 1.08 (C) | 2:36 ✓ |
| 3 (recon) | 5:00 | 45.25 ft | 0:00, 5:00, 10:00, 15:00 | 0:57, 0:57, 0:57, 0:57 | 16:39 | 17:22 | 15:57 | 4:44 | 12.34 | 1343 | 1.06 (B) | 2:28 ✓ |
| 4 | 5:00 | 36.75 ft | 0:00, 5:00, 10:00, 15:00 | 1:03, 0:57, 0:57, 0:57 | 17:58 | 18:46 | 15:57 | 7:20 | 9.96 | 1700 | 1.09 (C) | 3:32 ✓ |
| 5 | 5:00 | 36.5 ft | 0:00, 5:00, 10:00, 15:00 | 1:01, 0:51, 0:51, 0:51 | 17:32 | 18:17 | 15:51 | 7:28 | 10.75 | 1667 | 0.87 (B) | 3:34 ✓ |
| 6 | 5:00 | 30.83 ft | 0:00, 5:00, 10:00, 15:00 | 1:05, 0:53, 0:53, 0:53 | 17:53 | 18:28 | 15:53 | 9:12 | 9.01 | 1723 | 0.97 (B) | 4:20 ✗ |
| 7 | 5:00 | 32.75 ft | 0:00, 5:00, 10:00, 15:00 | 1:05, 0:53, 0:53, 0:53 | 17:52 | 18:23 | 15:53 | 10:36 | 8.80 | 1732 | 0.91 (B) | 4:03 ✗ |
| 8 | 5:00 | 36.25 ft | 0:00, 5:00, 10:00, 15:00 | 1:04, 0:57, 0:57, 0:57 | 17:58 | 18:44 | 15:57 | 7:28 | 9.68 | 1703 | 1.19 (C) | 3:37 ✓ |
| 9 | 5:00 | 40.08 ft | 0:00, 5:00, 10:00, 15:00 | 1:01, 1:01, 1:01, 1:01 | 16:45 | 17:18 | 16:01 | 5:16 | 11.73 | 1376 | 1.03 (B) | 2:47 ✓ |
| 10 | 5:00 | 64.5 ft | 0:00, 5:00, 10:00, 15:00 | 0:53, 0:53, 0:53, 0:53 | 16:24 | 17:13 | 15:53 | 2:12 | 17.77 | 1510 | 0.62 (A) | 1:57 ✓ |
| 11 | 5:00 | 27.92 ft | 0:00, 5:00, 10:00, 15:00 | 1:05, 1:05, 1:05, 1:05 | 18:05 | 18:38 | 16:05 | 11:16 | 8.28 | 1749 | 1.57 (D) | 4:47 ✗ |
<!-- results-table:end -->

### History

As bugs were fixed, the results table above was updated,
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
- **Used each VCE's width from [`data/vces.csv`](./data/vces.csv)**
  on platforms 3 and 6, instead of the ETA report's totals,
  still with arriving passengers spread across the VCEs in proportion to their widths.
  Platform 3's VCEs total 43 ft instead of 42.5 ft,
  so with 2-minute headways it clears at 266 s instead of 269 s.
  Platform 6's total only 30.8 ft instead of 48.168 ft,
  so it clears at 371 s instead of 238 s, and bottoms out at 3.8 sq ft/pax instead of 4.0.
  Penn Reconstruction and platforms 10 and 11 have no per-VCE data, so they're unchanged.
- **Made each train as long as its platform's tracks allow, up to 12 cars,**
  with 135 passengers and 4 doors per car.
  Platform 3's tracks only fit 10 cars, so its trains carry 1,350 passengers on 40 doors,
  and with 2-minute headways it clears at 231 s instead of 266 s
  (227 s instead of 256 s with Penn Reconstruction).
  Platforms 6, 10, and 11 still get 12-car trains, so they're unchanged.
- **Sent each door's arriving passengers to the nearest VCE,**
  on platforms 3 and 6, instead of spreading them across the VCEs in proportion to their widths.
  Each train's doors are spread evenly along its cars, stopped flush with the platform's east end.
  Narrow VCEs near many doors get long queues while others run dry:
  on platform 3, the 3.1 ft escalator P3-S3 gets 8 doors' passengers,
  so neither platform clears within the 600 s simulated.
- **Added the time arriving passengers take to walk from the doors to the VCEs,**
  at the TCQSM's design walking speed, 250 ft/min,
  on platforms 3 and 6.
  The farthest walk takes 29 s, but most doors are near a VCE,
  so the platform peaks at up to 12 more passengers, and the time at capacity is up to 6 s shorter.
- **Sent arriving passengers to the quickest VCE instead of the nearest,**
  i.e. the one with the least walking time plus waiting time for everyone queued or walking there.
  Platform 3 with 2-minute headways now clears at 310 s,
  44 s later than with its VCEs as one pooled queue, where it never cleared with the nearest VCE,
  and platform 6 clears at 469 s, 98 s later than as one pooled queue.
- **Added the time departing passengers take to walk from the VCEs to their train's doors,**
  on platforms 3 and 6, spreading evenly across the doors at 250 ft/min.
  They walk about 50 to 100 s on average, depending on the VCE, and up to 197 s,
  so platform 3 with 2-minute headways finishes boarding at 472 s instead of 298 s,
  and platform 6 at 593 s instead of 406 s.
- **Ran escalators one way on platforms 3 and 6** instead of treating them as stairs going both ways:
  a platform's only escalator goes up, and with more,
  one goes up, one goes down, and the rest go up until the platform is nearly fully alighted.
  Platform 3's 2 escalators leave 2.8 ft less going up,
  so with 2-minute headways it clears at 325 s instead of 310 s,
  but the down escalator lets departing passengers board by 457 s instead of 472 s.
  Platform 6 clears at 510 s instead of 469 s, and finishes boarding by 556 s instead of 593 s.
  The time at capacity now counts only the VCEs going up.
- **Sent departing passengers to the nearest car that isn't close to full,**
  on platforms 3 and 6, instead of spreading them evenly across the doors,
  with each car boarding its own waiting passengers through its own doors.
  Platform 3 with 2-minute headways finishes boarding at 284 s instead of 457 s,
  but with 5-minute headways at 363 s instead of 351 s, since the busiest cars' doors hold them up.
  Platform 6 finishes boarding at 372 s instead of 556 s.
- **Ran the escalator toward 7th Avenue up and the one toward 8th Avenue down,**
  instead of the other way around, matching AM peak demand toward 7th Avenue.
  Platform 3 with 2-minute headways clears at 318 s instead of 325 s,
  and finishes boarding at 293 s instead of 284 s.
  Platform 6 clears at 497 s instead of 510 s,
  and finishes boarding at 383 s instead of 372 s.
- **Stopped trains where the arriving passengers clear the platform soonest,**
  on platforms 3 and 6, instead of flush with the platform's east end.
  Platform 3 with 2-minute headways clears at 267 s instead of 270 s,
  with its trains' east ends 45 ft west of the platform's
  (5 ft with 5-minute headways),
  and platform 6 at 416 s instead of 453 s, 45 ft west too.
- **Stopped trains where the longer of their dwells is shortest,**
  instead of where the arriving passengers clear the platform soonest.
  Platform 3 with 2-minute headways has dwells of 251 s and 131 s instead of 256 s and 136 s,
  with its trains flush with the platform's east end,
  but clears at 270 s instead of 267 s.
  Platform 6's trains also stop flush with its east end,
  with dwells of 411 s instead of 420 s, but clearing at 453 s instead of 416 s.
