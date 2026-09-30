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

Every simulation also checks that it neither created nor lost passengers,
and stops with an error if it did.

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
   - Every train is as long as its platform's tracks and the platform itself allow, up to 12 cars:
     10 cars on platform 3, and 12 on platforms 6, 10, and 11.
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
   - Each VCE has its own queue, which it discharges at its capacity
     as long as anyone is queued
     ([TCQSM, p. 10-51](https://onlinepubs.trb.org/onlinepubs/tcrp/tcrp_rpt_165ch-10.pdf#page=55)).
   - A stair's capacity is LOS E capacity, 17 pax/min per foot of its width
     ([Fruin, p. 14](https://onlinepubs.trb.org/Onlinepubs/hrr/1971/355/355-001.pdf#page=14)).
   - An escalator's is the TCQSM's nominal capacity at 90 ft/min:
     34 pax/min with treads narrower than 2'8", and 72 pax/min with wider treads,
     since 2'8" treads carry close to 3'4" treads' capacity
     ([TCQSM, Exhibit 10-31, p. 10-52](https://onlinepubs.trb.org/onlinepubs/tcrp/tcrp_rpt_165ch-10.pdf#page=56)).
   - Each VCE's width and position are from
     [`data/vces.csv`](./data/vces.csv)
     (see [Estimated Widths](#estimated-widths)).
     - The train's doors are spread evenly along it,
       and it stops wherever on the platform the longest of the trains' dwells is shortest,
       trying every position a car length (85') apart,
       then every position 5' apart within a car length of the best of those.
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
     20' of queue in front of the VCEs at 5 sq ft/pax
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
- Every platform is modeled both as it is and as Penn Transformation would leave it
  (see [Penn Transformation](#penn-transformation)):
  with its new VCEs, platforms 1 to 3 extended to the west for 10-car trains,
  and every platform decluttered, adding the FRA's percentage more area,
  and the extension's length at the platform's average width.
- Platform A, a new platform PCIP Phase 1 proposes south of platform 1,
  is modeled, too (see [Platform A](#platform-a)).

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
  which matters at the 2'8" boundary between 34 and 72 pax/min.
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

For NJT trains that only alight ("drop and go"),
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
  not one 500' down a crowded platform.
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
aren't used, since each VCE has its own width.
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
[`data/vces.csv`](./data/vces.csv):

- Positions are in the Master Plan's frame, feet east of its plans' west edge.
  The two drawings register to within about 1'2":
  the platforms' east ends on them are all the same distance apart.
  Each platform's east end in that frame, for the model's trains to stop against,
  is in [`data/platform_east_ends.csv`](./data/platform_east_ends.csv):
  the sheet's on platforms 1 to 8, and the Master Plan's on platforms 9 to 11.
- VCEs matching a Master Plan VCE of the same type within 15'
  have `width_source` `master_plan`, with the Master Plan's width.
  The rest, 45 of 61, are `estimated`, with only the sheet's width.
- Their `source` is `pcip_phase_2`.
  Platforms 9 to 11's VCEs, which aren't on the sheet,
  have `source` `njt_directory`: they're from NJT's directory, as described below.
- For the 10 matched stairs, the sheet's widths differ from the Master Plan's by up to 1'2"
  (a median of 4"), and for the 6 matched escalators, by up to 9"
- Escalators' treads are their steps, narrower than their balustrades.
- The platforms' labels hide what's under them,
  including an escalator the Master Plan has about 230' along each of platforms 3 to 8.

| Platform | Stairs on the sheet | Escalators on the sheet | Their total width | Of which estimated | Master Plan VCEs not matched on the sheet |
|---|---|---|---|---|---|
| 1 | 6 | 2 | 35' | 16'6" | 3'8"/4'9" stair at 689'; 3'10" escalator at 774'; 4'4" stair at 776' |
| 2 | 5 | 4 | 35'11" | 23' | 4'4" stair at 775'; 3'10" escalator at 774' |
| 3 | 6 | 2 | 43' | 35'10" | 2'10" escalator at 230'; 5'9" stair at 278'; 3'8" stair at 684' |
| 4 | 5 | 2 | 36'9" | 30'9" | 2'10" escalator at 230'; 2'10" escalator at 404'; 3'8" stair at 683' |
| 5 | 4 | 4 | 36'6" | 33'8" | 2'10" escalator at 230'; 3'8" stair at 683' |
| 6 | 3 | 4 | 30'10" | 23'10" | 2'10" escalator at 230' |
| 7 | 4 | 3 | 32'9" | 25'1" | 2'10" escalator at 230' |
| 8 | 5 | 2 | 36'3" | 31'5" | 2'10" escalator at 230'; 2'10" escalator at 404' |

The unmatched Master Plan VCEs are a mix:
the escalators at 230' are hidden under the sheet's labels, so they should be added;
others, like platform 3's 3'8" stair at 684',
are probably the same VCEs as similar ones on the sheet 15' to 25' away;
and some disagree on the type, like platform 3's 5'9" stair at 278',
where the sheet has a 3'1" escalator.
So platform 3's VCEs total about 45'10" with its hidden escalator,
a little more than the model's 42'6",
including the 2 West End Concourse stairs (8'11" and 10') and the Exit Concourse's 2 (5'1" and 5').

These are only estimates:

- A drawn tread isn't necessarily the clear width between handrails.
- The sheet's existing stairs aren't labeled,
  so some may be stairs between the upper and lower concourses drawn over a platform,
  though every platform's matches the directory's pattern,
  e.g. 2 stairs each for the West End and Exit Concourses.
- The sheet predates NJT's replacement of an escalator on tracks 7/8 (platform 4)
  with stairs in about 2021.
- The model gives an escalator the TCQSM's capacity for its tread width,
  so a few inches' error can move it across 2'8", between 34 and 72 pax/min.

Platforms 9 to 11 aren't on that plan,
so their VCEs are estimated from NJT's January 2022 station directory instead,
whose map is schematic and not to scale.
Each level of its map is calibrated to feet by matching its icons on platforms 1 to 8,
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
which `vces` adds to [`data/vces.csv`](./data/vces.csv) with `source` `moynihan_ea`.

The plan is a design from before the Train Hall was built, so:

- The Train Hall opened in 2021 with 11 escalators to platforms 3 to 8, not the plan's 12.
  Platform 3's western one is taken as the one that wasn't built,
  since it would run past the platform's west end.
- The plan doesn't draw the escalators consistently enough to measure their width,
  so they're taken to be 3'4" wide,
  the 1,000 mm steps [KONE](https://elevatorworld.com/article/let-there-be-light-and-accessibility/) reports for its escalators there.
- The stairs' widths are their drawn treads':
  6'6" on platforms 9 and 10, 3'8" on platform 11,
  and 6' for a second stair on platform 9, about 90' west of the West End Concourse.

### Penn Transformation

The FRA's Penn Station Service Optimization Study
([Phase I report](https://railroads.dot.gov/elibrary/new-york-penn-station-service-optimization-study-phase-i-report-final-june-2026), June 2026)
lays out Penn Transformation's improvements to the platforms:
up to 23 new VCEs, 19 stairs and 4 escalators, with at least one on every platform;
Platforms 1 to 3 extended west by about 200' to 350',
so 10-car NJT trains can open all of their doors;
and decluttered platforms, with 2% to 14% more circulation area (Table 1).
Its Figure 11 shows the new VCEs' "generalized locations" on an aerial image,
which `uv run platform-crowd-model data vces-transformation-fra-sos`
([`vces_transformation_fra_sos.py`](./src/platform_crowd_model/vces_transformation_fra_sos.py))
finds by their icons, along with the extensions,
and registers to the Master Plan's frame by the platforms' drawn ends to within about 3',
into [`data/vces_transformation_fra_sos.csv`](./data/vces_transformation_fra_sos.csv)
and [`data/platforms_transformation_fra_sos.csv`](./data/platforms_transformation_fra_sos.csv).
They agree with the report's breakdown:
5 in Moynihan, 2 on Platforms 1 and 2's extensions west of Eighth Avenue, and 16 in Penn Station,
mostly in two rows, about 175' and 535' east of the Master Plan's plans' west edge.

The report doesn't say which are escalators or how wide any are, so:

- The 4 escalators are taken to be the 4 in Moynihan, about 250' west of the West End Concourse,
  in line with the Train Hall's escalators.
- Stairs are as wide as the West End Concourse's, 6', and escalators as the Train Hall's, 3'4"
  Together, the 23 add about 30% to the existing VCEs' total width,
  close to the 32% more vertical circulation capacity
  [Penn Transformation's designers report](https://www.enr.com/articles/63127-penn-station-renderings-reveal-design-for-8b-reconstruction-beneath-madison-square-garden).
- Its locations are only general, and a few of them are within a few feet of existing VCEs,
  e.g. on platform 10, whose existing VCEs are only estimated from NJT's directory.
- Each extended platform's new west end is where its extension's box on Figure 11 ends,
  185' to 317' west of its end in PCIP Phase 1's plan.

### Platform A

NJT's PCIP Phase 1 study proposed, as its Alternative 12, a new Platform A
south of platform 1, under West 31st Street,
from under the 7th Avenue Subway west under the 8th Avenue Subway,
long enough for a 12-car train
([final report, pp. 95 to 97](https://liamblank.com/wp-content/uploads/2026/07/penn-records-s-nj-transit-pcip-1-pcip1-final-report-c5015-01-262652-00-task-06-mem-final-report-draft-final-rev1-copy.pdf#page=95)).
It would have 8 stairs, 6 escalators, and 4 elevators:
6 pairs of a stair and an escalator up to a new Concourse A above it,
and 2 stairs up to an extension of the West End Concourse,
where its west end curves too much for escalators.
`uv run platform-crowd-model data platform-a-pcip-phase-1`
([`platform_a_pcip_phase_1.py`](./src/platform_crowd_model/platform_a_pcip_phase_1.py))
measures it and its VCEs on its plan
([Appendix A, sheet A-021](https://liamblank.com/wp-content/uploads/2026/07/penn-records-s-nj-transit-pcip-1-pcip1-final-report-c5015-01-262652-00-task-06-mem-final-report-draft-appendixa-drawings-copy.pdf#page=31)),
registered to the Master Plan's frame by the existing platforms' east ends to within about 1',
into [`data/platform_a_pcip_phase_1.csv`](./data/platform_a_pcip_phase_1.csv)
and [`data/vces_platform_a_pcip_phase_1.csv`](./data/vces_platform_a_pcip_phase_1.csv):
1,013' long, from 90' west of the Master Plan's plans' west edge,
with 22,717 sq ft inside its outline.
The report doesn't give the VCEs' widths,
so stairs are taken to be the 5'-0" egress stairs it sizes its other alternatives' with,
and escalators to have 3'4" steps.
In the model, it's platform 0.

### Field Survey

Since no public source has every VCE's width,
[`data/vces_field_survey.csv`](./data/vces_field_survey.csv) is a sheet for measuring them in person,
made by `uv run platform-crowd-model data vces-field-survey`
([`vces_field_survey.py`](./src/platform_crowd_model/vces_field_survey.py)).
It lists each platform's VCEs expected from the PCIP Phase 2 plan (platforms 1 to 8),
the 2022 directory, and the Master Plan,
each sorted west to east, since they can't all be aligned reliably,
starting with platform 3, the ETA report's focus,
then platform 11, which has no width data.
Surveyors fill in the columns after `midpoint_ft`, feet east of the Master Plan's plans' west edge:

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

Regenerating the sheet keeps what surveyors have entered.
Each row that's still generated keeps its entries,
matched by its platform, source, type, and position rather than its `vce_name`,
which can change as VCEs are added.
Every other row with entries, like one a surveyor added, is kept at the end.

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
  (see [Counts from NJT's Station Directory](#counts-from-nj-transits-station-directory)).
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

### Platform Dimensions

Each platform's length is from the Moynihan Station EA's Table 4.4-10
([p. 4.4-22](https://web.archive.org/web/20241011135133/https://cdn.esd.ny.gov/subsidiaries_projects/msdc/Data/NEPA/04_4%20StationPedCirculation.pdf#page=22)),
in [`data/platform_lengths_moynihan_ea.csv`](./data/platform_lengths_moynihan_ea.csv),
the only source found with written platform lengths.
Its area is from its outline in OpenStreetMap, which OpenRailwayMap draws,
in [`data/platforms_osm.csv`](./data/platforms_osm.csv),
written by `uv run platform-crowd-model data platforms-osm`
([`platforms_osm.py`](./src/platform_crowd_model/platforms_osm.py)),
since platforms taper toward their ends.
The outlines have no source, but agree with the PCIP Phase 2 existing plan's widths
to within about 2', and with the EA's lengths to within about 60',
except platform 9's outline, which is 178' longer than the EA's length.
The PCIP Phase 2 and Master Plan drawings cut the platforms off at their west ends,
so they can't give lengths or areas.

NJT's PCIP Phase 1 existing track plan
([Appendix A, sheet TK-003](https://liamblank.com/wp-content/uploads/2026/07/penn-records-s-nj-transit-pcip-1-pcip1-final-report-c5015-01-262652-00-task-06-mem-final-report-draft-appendixa-drawings-copy.pdf#page=8),
July 2019) doesn't, though.
It's a scan at 1" = 80',
so `uv run platform-crowd-model data platform-west-ends-pcip-phase-1`
([`platform_west_ends_pcip_phase_1.py`](./src/platform_crowd_model/platform_west_ends_pcip_phase_1.py))
finds its orange platform edges by color
and fits them to the platforms' east ends in the Master Plan's frame,
which comes out at 80.3 ft per inch of the sheet, with every east end within about 1'.
Each platform's west end is in
[`data/platform_west_ends_pcip_phase_1.csv`](./data/platform_west_ends_pcip_phase_1.csv).
The lengths between them agree with the EA's to within about 40',
except platform 9's, which is 134' longer, like its outline in OpenStreetMap.

Platform 11 is 1,007' long, a few feet short of a 12-car train at 85' per car,
but the EA has it take 12-car LIRR trains,
so trains can overhang their platform's west end by up to 15'.

### Platform Shapes

The model doesn't use them yet, but the shapes of platforms 1 to 8 and of what's on them
are extracted from the PCIP Phase 2 existing plan
([sheet A-001](https://liamblank.com/wp-content/uploads/2026/09/pcip-2-conceptual-design-preliminary-drawings.pdf#page=45))
by `uv run platform-crowd-model data shapes-pcip-phase-2`
([`shapes_pcip_phase_2.py`](./src/platform_crowd_model/shapes_pcip_phase_2.py)):

- each platform's outline, with platforms 1 and 2 sharing one, as the sheet draws them
- each VCE's footprint: the bounding box of its treads and the balustrades beside them,
  named as in [`data/vces.csv`](./data/vces.csv)
- the columns, the small squares on the platforms
- the elevators, the boxes with an X across them

Rooms, walls, and the enclosures around VCEs aren't extracted yet.
The shapes are 2D, at the platform level,
since what matters on the platform is the space each VCE takes up there;
going up, a VCE's capacity is already its width.

They're GeoJSON, one feature per line, in two files:

- [`data/shapes_pcip_phase_2.geojson`](./data/shapes_pcip_phase_2.geojson),
  in feet in the Master Plan's frame, extended to 2D:
  x east of its plans' west edge, as in `data/vces.csv`,
  and y north of platform 5's centerline,
  where east and north are along Manhattan's street grid.
  GeoJSON requires longitude and latitude, so this is strictly not valid GeoJSON,
  but it's what the code reads.
- [`data/shapes_pcip_phase_2_lonlat.geojson`](./data/shapes_pcip_phase_2_lonlat.geojson),
  the same shapes in longitude and latitude, so GitHub can show them on a map.
  They're registered to the platforms' outlines in OpenStreetMap:
  rotated 29.2° to their average direction, the street grid's,
  and offset to match their east ends on average,
  which agree to within about 4' across the platforms, but only about 37' along them,
  like OpenStreetMap's lengths.

## Results

Times are in m:ss.
Headways, time at capacity, and dwells are durations, and the rest are times after the first train arrives.

- **Platform:** the platform's number, with `(transformation)` as Penn Transformation would leave it,
  or `A` for PCIP Phase 1's Platform A.
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

This table is generated by `uv run platform-crowd-model run --update-readme`.
Platforms 9 to 11's VCEs are only estimated from NJT's station directory
(see [Estimated Widths](#estimated-widths)).
The directory doesn't show every VCE, and the Master Plan has none on platform 11,
so its VCEs total only 31'7", much less than the ETA report's 43'7",
which likely makes it look much worse than it is.
Penn Transformation's new VCEs are only at the FRA's "generalized locations",
with widths assumed from Moynihan's
(see [Penn Transformation](#penn-transformation)).


<!-- results-table:start -->
| Platform | Headway | VCE width | Arrivals | Dwell | Taper time | Clear time | Boarded time | Time at capacity | Max up rate (pax/s) | Max pax on platform | Max density (pax/m²) | NFPA 130 evacuation |
|---|---|---|---|---|---|---|---|---|---|---|---|---|
| A | 0:00 | 60' | 0:00, 0:00, 0:50, 0:50 | 0:50, 0:50, 0:50, 0:50 | 6:08 | 6:44 | 1:40 | 5:36 | 17.33 | 5824 | 3.68 (E) | 6:42 ✗ |
| 1 | 0:00 | 35' | 0:00, 0:00, 4:34, 4:34 | 4:34, 4:34, 1:04, 1:04 | 8:54 | 9:39 | 5:38 | 7:55 | 9.25 | 3284 | 3.12 (D) | 6:44 ✗ |
| 1 (transformation) | 0:00 | 47' | 0:00, 0:00, 0:48, 0:48 | 0:48, 0:48, 0:48, 0:48 | 7:19 | 8:15 | 1:36 | 5:52 | 12.65 | 5180 | 3.78 (E) | 7:41 ✗ |
| 2 | 0:00 | 35'11" | 0:00, 0:00, 3:51, 3:51 | 3:51, 3:51, 1:05, 1:05 | 8:12 | 9:06 | 4:56 | 6:09 | 10.47 | 3386 | 3.42 (D) | 6:39 ✗ |
| 2 (transformation) | 0:00 | 47'11" | 0:00, 0:00, 0:52, 0:52 | 0:52, 0:52, 0:52, 0:52 | 6:21 | 6:48 | 1:44 | 5:42 | 13.87 | 5032 | 3.78 (E) | 7:15 ✗ |
| 3 | 0:00 | 46'4" | 0:00, 0:00, 0:56, 0:56 | 0:56, 0:56, 0:56, 0:56 | 6:45 | 7:20 | 1:52 | 6:43 | 12.91 | 5040 | 3.97 (E) | 7:32 ✗ |
| 3 (transformation) | 0:00 | 61'8" | 0:00, 0:00, 0:52, 0:52 | 0:52, 0:52, 0:52, 0:52 | 4:55 | 5:35 | 1:44 | 4:51 | 17.51 | 4695 | 2.68 (D) | 5:22 ✗ |
| 4 | 0:00 | 43'5" | 0:00, 0:00, 1:24, 1:24 | 1:24, 1:24, 0:56, 0:56 | 8:43 | 9:42 | 2:20 | 8:05 | 12.36 | 5825 | 3.73 (E) | 9:13 ✗ |
| 4 (transformation) | 0:00 | 55'5" | 0:00, 0:00, 0:54, 0:54 | 0:54, 0:54, 0:54, 0:54 | 7:01 | 8:09 | 1:48 | 6:08 | 15.76 | 5898 | 3.65 (E) | 7:19 ✗ |
| 5 | 0:00 | 43'2" | 0:00, 0:00, 0:56, 0:56 | 0:56, 0:56, 0:56, 0:56 | 8:05 | 8:44 | 1:52 | 7:45 | 13.15 | 6101 | 3.17 (D) | 9:43 ✗ |
| 5 (transformation) | 0:00 | 58'6" | 0:00, 0:00, 0:52, 0:52 | 0:52, 0:52, 0:52, 0:52 | 6:02 | 6:37 | 1:44 | 5:28 | 17.75 | 5755 | 2.87 (D) | 6:49 ✗ |
| 6 | 0:00 | 37'6" | 0:00, 0:00, 2:40, 2:40 | 2:40, 2:40, 1:05, 1:05 | 10:39 | 11:36 | 3:45 | 6:05 | 11.41 | 5332 | 2.99 (D) | 9:54 ✗ |
| 6 (transformation) | 0:00 | 49'6" | 0:00, 0:00, 0:53, 0:53 | 0:53, 0:53, 0:53, 0:53 | 7:09 | 8:21 | 1:47 | 6:57 | 14.81 | 5993 | 3.26 (D) | 8:20 ✗ |
| 7 | 0:00 | 39'5" | 0:00, 0:00, 1:19, 1:19 | 1:19, 1:19, 1:00, 1:00 | 10:27 | 11:23 | 2:19 | 7:38 | 11.20 | 6116 | 3.21 (D) | 10:39 ✗ |
| 7 (transformation) | 0:00 | 48'9" | 0:00, 0:00, 0:54, 0:54 | 0:54, 0:54, 0:54, 0:54 | 7:56 | 8:59 | 1:48 | 6:46 | 14.10 | 6043 | 3.04 (D) | 8:30 ✗ |
| 8 | 0:00 | 42'11" | 0:00, 0:00, 1:52, 1:52 | 1:52, 1:52, 0:56, 0:56 | 9:00 | 9:53 | 2:48 | 8:11 | 12.08 | 5521 | 3.85 (E) | 8:52 ✗ |
| 8 (transformation) | 0:00 | 52'3" | 0:00, 0:00, 0:53, 0:53 | 0:53, 0:53, 0:53, 0:53 | 7:02 | 7:38 | 1:46 | 6:50 | 14.98 | 5981 | 4.00 (E) | 7:51 ✗ |
| 9 | 0:00 | 52'7" | 0:00, 0:59, 1:58, 2:57 | 0:59, 0:59, 0:59, 0:59 | 6:18 | 7:19 | 3:56 | 2:37 | 15.27 | 2789 | 2.09 (D) | 3:57 ✓ |
| 9 (transformation) | 0:00 | 64'7" | 0:00, 1:04, 2:08, 3:12 | 1:04, 1:04, 1:04, 1:04 | 5:49 | 6:58 | 4:16 | 0:00 | 17.97 | 2470 | 1.82 (D) | 2:50 ✓ |
| 10 | 0:00 | 71' | 0:00, 0:00, 0:51, 0:51 | 0:51, 0:51, 0:51, 0:51 | 5:20 | 6:22 | 1:43 | 4:58 | 19.62 | 5626 | 2.31 (D) | 5:30 ✗ |
| 10 (transformation) | 0:00 | 77' | 0:00, 0:00, 0:56, 0:56 | 0:56, 0:56, 0:56, 0:56 | 4:59 | 6:06 | 1:52 | 4:22 | 21.32 | 5380 | 1.93 (D) | 4:54 ✗ |
| 11 | 0:00 | 31'7" | 0:00, 0:00, 7:10, 7:10 | 7:10, 7:10, 1:05, 1:05 | 12:46 | 13:25 | 8:15 | 10:44 | 9.32 | 3907 | 3.51 (D) | 8:43 ✗ |
| 11 (transformation) | 0:00 | 43'7" | 0:00, 0:00, 4:37, 4:37 | 4:37, 4:37, 1:05, 1:05 | 8:45 | 9:38 | 5:42 | 7:31 | 12.72 | 4202 | 3.70 (E) | 6:45 ✗ |
| A | 2:00 | 60' | 0:00, 2:00, 4:00, 6:00 | 0:50, 1:05, 1:05, 1:05 | 7:22 | 7:52 | 7:05 | 4:40 | 17.33 | 1493 | 0.94 (B) | 2:08 ✓ |
| 1 | 2:00 | 35' | 0:00, 2:00, 4:00, 7:58 | 0:50, 5:58, 3:59, 1:05 | 9:57 | 10:25 | 9:03 | 7:41 | 9.25 | 1365 | 1.30 (C) | 3:50 ✓ |
| 1 (transformation) | 2:00 | 47' | 0:00, 2:00, 4:00, 6:00 | 0:56, 1:05, 2:38, 1:05 | 8:07 | 9:03 | 7:05 | 3:33 | 12.65 | 1549 | 1.13 (C) | 3:06 ✓ |
| 2 | 2:00 | 35'11" | 0:00, 2:00, 4:00, 7:21 | 1:01, 5:21, 3:22, 1:04 | 9:08 | 9:56 | 8:25 | 6:17 | 10.47 | 1322 | 1.33 (C) | 3:38 ✓ |
| 2 (transformation) | 2:00 | 47'11" | 0:00, 2:00, 4:00, 6:00 | 0:52, 1:05, 1:05, 1:05 | 7:27 | 7:51 | 7:05 | 4:12 | 13.87 | 1394 | 1.05 (B) | 2:20 ✓ |
| 3 | 2:00 | 46'4" | 0:00, 2:00, 4:00, 6:17 | 0:56, 4:17, 2:20, 1:05 | 7:49 | 8:25 | 7:22 | 5:45 | 12.91 | 1389 | 1.09 (C) | 2:57 ✓ |
| 3 (transformation) | 2:00 | 61'8" | 0:00, 2:00, 4:00, 6:00 | 0:52, 1:02, 1:02, 1:02 | 7:07 | 7:41 | 7:02 | 1:08 | 17.51 | 1257 | 0.72 (A) | 1:48 ✓ |
| 4 | 2:00 | 43'5" | 0:00, 2:00, 4:00, 7:30 | 0:59, 5:30, 3:31, 1:05 | 9:35 | 10:25 | 8:35 | 7:04 | 12.36 | 1698 | 1.09 (C) | 3:35 ✓ |
| 4 (transformation) | 2:00 | 55'5" | 0:00, 2:00, 4:00, 6:00 | 0:54, 1:05, 1:05, 1:05 | 7:49 | 8:51 | 7:05 | 4:51 | 15.76 | 1642 | 1.02 (B) | 2:26 ✓ |
| 5 | 2:00 | 43'2" | 0:00, 2:00, 4:00, 7:16 | 0:56, 5:16, 3:18, 1:04 | 9:08 | 9:45 | 8:20 | 6:41 | 13.15 | 1628 | 0.85 (B) | 3:36 ✓ |
| 5 (transformation) | 2:00 | 58'6" | 0:00, 2:00, 4:00, 6:00 | 0:52, 1:05, 1:05, 1:05 | 7:21 | 7:46 | 7:05 | 4:20 | 17.75 | 1482 | 0.74 (A) | 2:11 ✓ |
| 6 | 2:00 | 37'6" | 0:00, 2:00, 4:00, 8:29 | 0:59, 6:29, 4:10, 1:05 | 10:41 | 11:18 | 9:34 | 8:32 | 11.41 | 1864 | 1.04 (B) | 4:12 ✗ |
| 6 (transformation) | 2:00 | 49'6" | 0:00, 2:00, 4:00, 6:23 | 0:53, 4:23, 2:33, 1:05 | 8:02 | 8:32 | 7:28 | 5:47 | 14.81 | 1582 | 0.86 (B) | 3:06 ✓ |
| 7 | 2:00 | 39'5" | 0:00, 2:00, 4:00, 8:20 | 0:58, 6:20, 4:20, 1:05 | 10:41 | 11:32 | 9:25 | 7:53 | 11.20 | 1937 | 1.02 (B) | 4:06 ✗ |
| 7 (transformation) | 2:00 | 48'9" | 0:00, 2:00, 4:00, 6:00 | 1:05, 1:14, 1:28, 1:32 | 9:39 | 10:41 | 7:32 | 0:13 | 14.10 | 2484 | 1.25 (C) | 3:40 ✓ |
| 8 | 2:00 | 42'11" | 0:00, 2:00, 4:00, 7:45 | 0:58, 5:45, 3:50, 1:05 | 10:35 | 11:36 | 8:50 | 4:56 | 12.08 | 2021 | 1.41 (C) | 4:02 ✗ |
| 8 (transformation) | 2:00 | 52'3" | 0:00, 2:00, 4:00, 6:07 | 0:53, 4:07, 2:53, 1:05 | 8:39 | 9:43 | 7:12 | 2:11 | 14.98 | 1894 | 1.27 (C) | 3:15 ✓ |
| 9 | 2:00 | 52'7" | 0:00, 2:00, 4:00, 6:00 | 0:59, 1:05, 1:05, 1:05 | 7:30 | 8:22 | 7:05 | 0:24 | 15.27 | 1464 | 1.10 (C) | 2:10 ✓ |
| 9 (transformation) | 2:00 | 64'7" | 0:00, 2:00, 4:00, 6:00 | 1:04, 1:05, 1:05, 1:05 | 7:19 | 8:19 | 7:05 | 0:00 | 16.97 | 1460 | 1.07 (B) | 1:45 ✓ |
| 10 | 2:00 | 71' | 0:00, 2:00, 4:00, 6:00 | 0:51, 1:04, 1:05, 1:04 | 7:13 | 8:12 | 7:04 | 2:36 | 19.62 | 1507 | 0.62 (A) | 1:48 ✓ |
| 10 (transformation) | 2:00 | 77' | 0:00, 2:00, 4:00, 6:00 | 0:56, 1:05, 1:05, 1:05 | 7:09 | 8:12 | 7:05 | 0:52 | 21.32 | 1468 | 0.53 (A) | 1:40 ✓ |
| 11 | 2:00 | 31'7" | 0:00, 2:00, 4:00, 11:15 | 1:05, 9:15, 7:15, 1:05 | 13:58 | 14:42 | 12:20 | 10:23 | 9.32 | 2314 | 2.08 (D) | 5:27 ✗ |
| 11 (transformation) | 2:00 | 43'7" | 0:00, 2:00, 4:00, 8:26 | 1:01, 6:26, 4:16, 1:05 | 10:37 | 11:22 | 9:31 | 4:51 | 12.72 | 1678 | 1.48 (C) | 3:19 ✓ |
| A | 5:00 | 60' | 0:00, 5:00, 10:00, 15:00 | 0:50, 0:50, 0:50, 0:50 | 16:22 | 16:52 | 15:50 | 4:40 | 17.33 | 1440 | 0.91 (B) | 2:07 ✓ |
| 1 | 5:00 | 35' | 0:00, 5:00, 10:00, 15:00 | 0:50, 0:50, 0:50, 0:50 | 16:59 | 17:27 | 15:50 | 7:04 | 9.25 | 1311 | 1.24 (C) | 3:04 ✓ |
| 1 (transformation) | 5:00 | 47' | 0:00, 5:00, 10:00, 15:00 | 0:48, 0:48, 0:48, 0:48 | 16:38 | 17:24 | 15:48 | 4:32 | 12.65 | 1348 | 0.98 (B) | 2:24 ✓ |
| 2 | 5:00 | 35'11" | 0:00, 5:00, 10:00, 15:00 | 0:57, 0:57, 0:57, 0:57 | 16:49 | 17:39 | 15:57 | 6:04 | 10.47 | 1268 | 1.28 (C) | 2:54 ✓ |
| 2 (transformation) | 5:00 | 47'11" | 0:00, 5:00, 10:00, 15:00 | 0:52, 0:52, 0:52, 0:52 | 16:27 | 16:51 | 15:52 | 4:12 | 13.87 | 1313 | 0.99 (B) | 2:19 ✓ |
| 3 | 5:00 | 46'4" | 0:00, 5:00, 10:00, 15:00 | 0:56, 0:56, 0:56, 0:56 | 16:31 | 17:07 | 15:56 | 6:08 | 12.91 | 1319 | 1.04 (B) | 2:25 ✓ |
| 3 (transformation) | 5:00 | 61'8" | 0:00, 5:00, 10:00, 15:00 | 0:52, 0:52, 0:52, 0:52 | 16:04 | 16:38 | 15:52 | 4:00 | 17.51 | 1169 | 0.67 (A) | 1:47 ✓ |
| 4 | 5:00 | 43'5" | 0:00, 5:00, 10:00, 15:00 | 0:57, 0:57, 0:57, 0:57 | 17:06 | 17:57 | 15:57 | 6:32 | 12.36 | 1607 | 1.03 (B) | 2:59 ✓ |
| 4 (transformation) | 5:00 | 55'5" | 0:00, 5:00, 10:00, 15:00 | 0:54, 0:54, 0:54, 0:54 | 16:37 | 17:34 | 15:54 | 4:48 | 15.76 | 1499 | 0.93 (B) | 2:18 ✓ |
| 5 | 5:00 | 43'2" | 0:00, 5:00, 10:00, 15:00 | 0:56, 0:56, 0:56, 0:56 | 16:53 | 17:27 | 15:56 | 6:48 | 13.15 | 1581 | 0.82 (A) | 3:00 ✓ |
| 5 (transformation) | 5:00 | 58'6" | 0:00, 5:00, 10:00, 15:00 | 0:52, 0:52, 0:52, 0:52 | 16:21 | 16:46 | 15:52 | 4:20 | 17.75 | 1426 | 0.71 (A) | 2:10 ✓ |
| 6 | 5:00 | 37'6" | 0:00, 5:00, 10:00, 15:00 | 0:59, 0:59, 0:59, 0:59 | 17:10 | 17:48 | 15:59 | 8:24 | 11.41 | 1638 | 0.92 (B) | 3:30 ✓ |
| 6 (transformation) | 5:00 | 49'6" | 0:00, 5:00, 10:00, 15:00 | 0:53, 0:53, 0:53, 0:53 | 16:38 | 17:27 | 15:53 | 5:44 | 14.81 | 1526 | 0.83 (B) | 2:36 ✓ |
| 7 | 5:00 | 39'5" | 0:00, 5:00, 10:00, 15:00 | 0:58, 0:58, 0:58, 0:58 | 17:12 | 17:47 | 15:58 | 8:40 | 11.20 | 1648 | 0.86 (B) | 3:19 ✓ |
| 7 (transformation) | 5:00 | 48'9" | 0:00, 5:00, 10:00, 15:00 | 0:54, 0:54, 0:54, 0:54 | 16:43 | 17:31 | 15:54 | 6:24 | 14.10 | 1549 | 0.78 (A) | 2:38 ✓ |
| 8 | 5:00 | 42'11" | 0:00, 5:00, 10:00, 15:00 | 0:58, 0:58, 0:58, 0:58 | 17:03 | 17:48 | 15:58 | 7:20 | 12.08 | 1619 | 1.13 (C) | 3:01 ✓ |
| 8 (transformation) | 5:00 | 52'3" | 0:00, 5:00, 10:00, 15:00 | 0:53, 0:53, 0:53, 0:53 | 16:35 | 16:59 | 15:53 | 6:12 | 14.98 | 1522 | 1.02 (B) | 2:27 ✓ |
| 9 | 5:00 | 52'7" | 0:00, 5:00, 10:00, 15:00 | 0:59, 0:59, 0:59, 0:59 | 16:25 | 17:15 | 15:59 | 0:24 | 15.27 | 1329 | 1.00 (B) | 2:05 ✓ |
| 9 (transformation) | 5:00 | 64'7" | 0:00, 5:00, 10:00, 15:00 | 1:04, 1:04, 1:04, 1:04 | 16:15 | 17:14 | 16:04 | 0:00 | 16.97 | 1270 | 0.93 (B) | 1:41 ✓ |
| 10 | 5:00 | 71' | 0:00, 5:00, 10:00, 15:00 | 0:51, 0:51, 0:51, 0:51 | 16:12 | 17:09 | 15:51 | 2:56 | 19.62 | 1396 | 0.57 (A) | 1:46 ✓ |
| 10 (transformation) | 5:00 | 77' | 0:00, 5:00, 10:00, 15:00 | 0:56, 0:56, 0:56, 0:56 | 16:08 | 17:09 | 15:56 | 1:08 | 21.32 | 1347 | 0.48 (A) | 1:38 ✓ |
| 11 | 5:00 | 31'7" | 0:00, 5:00, 10:00, 15:00 | 1:05, 1:05, 1:05, 1:05 | 17:43 | 18:27 | 16:05 | 9:08 | 9.32 | 1712 | 1.54 (D) | 4:10 ✗ |
| 11 (transformation) | 5:00 | 43'7" | 0:00, 5:00, 10:00, 15:00 | 1:01, 1:01, 1:01, 1:01 | 17:11 | 17:56 | 16:01 | 4:00 | 12.72 | 1597 | 1.41 (C) | 2:57 ✓ |
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
  Platform 3's VCEs total 43' instead of 42'6",
  so with 2-minute headways it clears at 266 s instead of 269 s.
  Platform 6's total only 30'10" instead of 48'2",
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
  on platform 3, the 3'1" escalator P3-S3 gets 8 doors' passengers,
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
  Platform 3's 2 escalators leave 2'10" less going up,
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
  with its trains' east ends 45' west of the platform's
  (5' with 5-minute headways),
  and platform 6 at 416 s instead of 453 s, 45' west too.
- **Stopped trains where the longer of their dwells is shortest,**
  instead of where the arriving passengers clear the platform soonest.
  Platform 3 with 2-minute headways has dwells of 251 s and 131 s instead of 256 s and 136 s,
  with its trains flush with the platform's east end,
  but clears at 270 s instead of 267 s.
  Platform 6's trains also stop flush with its east end,
  with dwells of 411 s instead of 420 s, but clearing at 453 s instead of 416 s.
- **Added the escalators down from Moynihan Train Hall and the West End Concourse's stairs to platforms 9 to 11,**
  from the Moynihan Station EA's lower concourse plan,
  which the other drawings leave out:
  2 escalators each on platforms 4 to 8 and 1 on platform 3, at their west ends,
  and 1 or 2 stairs each on platforms 9 to 11.
  Every platform but 1 and 2 clears sooner:
  with trains 2 minutes apart, platform 5 at 9:45 instead of 12:06,
  platform 8 at 11:36 instead of 13:57, and platform 9 at 8:22 instead of 10:01,
  and platforms 4 and 5 now evacuate within NFPA 130's 4 minutes.
- **Modeled every platform with Penn Transformation, instead of platform 3 with Penn Reconstruction,**
  whose new VCEs were only one stair standing in for the ETA report's extra 2'3" of VCEs.
  Penn Transformation's 23 new VCEs, platform extensions, and decluttering
  are from the FRA's Service Optimization Study.
  With trains 2 minutes apart, every platform but 10 clears sooner with it,
  e.g. platform 5 at 7:46 instead of 9:45 and platform 11 at 11:22 instead of 14:42,
  and every platform evacuates within NFPA 130's 4 minutes.
- **Modeled PCIP Phase 1's Platform A,** a new platform south of platform 1,
  with its 8 stairs and 6 escalators.
  With trains 2 minutes apart, it clears at 7:52, sooner than any platform as it is today,
  and evacuates within NFPA 130's 4 minutes.
