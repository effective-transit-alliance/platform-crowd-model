# Platform Crowd Model

This repo is for modeling platform crowding and alighting and boarding of trains,
specifically at NY Penn Station.
[`vce_two_trains_alight_and_board.py`](./vce_two_trains_alight_and_board.py)
models a single island platform with two tracks.

## Running

To run, we use the Python package manager [`uv`](https://github.com/astral-sh/uv),
which can be installed easily with

```sh
curl -LsSf https://astral.sh/uv/install.sh | sh
```

With `uv` installed, you can then just run the script directly:

```sh
./vce_two_trains_alight_and_board.py
```

In doing so, `uv` will also install dependencies and set up a virtual environment.

Each run prints a table of headline results.
To also write a spreadsheet per scenario with the full time series and charts, run

```sh
./vce_two_trains_alight_and_board.py --spreadsheets
```

To also replace the [results table](#current-results) below with that run's, run

```sh
./vce_two_trains_alight_and_board.py --update-readme
```

Everything the model assumes, like stair capacity, passenger loads, and LOS thresholds,
is in the `Assumptions` `dataclass` at the top of
[`vce_two_trains_alight_and_board.py`](./vce_two_trains_alight_and_board.py),
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

The model simulates one island platform served by two tracks, one second at a time,
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
   - Every train is as long as its platform's tracks allow, up to 12 cars,
     per [this track map](https://www.railfanguides.us/ny/penntonewrochelle/PennStationLayout1.jpg):
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
   - All VCEs are treated as stairs, except one per platform, which is excluded.
   - They queue at the stairs, which discharge them at LOS E capacity,
     17 pax/min per foot of VCE width
     ([Fruin, p. 14](https://onlinepubs.trb.org/Onlinepubs/hrr/1971/355/355-001.pdf#page=14)),
     as long as anyone is queued
     ([TCQSM, p. 10-51](https://onlinepubs.trb.org/onlinepubs/tcrp/tcrp_rpt_165ch-10.pdf#page=55)).
   - So there's no gradual taper: the stairs stay at capacity until the platform is clear.
   - The report's "taper time" is when the remaining arriving passengers fit in the stair queues,
     20 ft of queue in front of the VCEs at 5 sq ft/pax
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
  alternating between the platform's two tracks,
  so each train after the first two can only arrive
  once the train before it on its track has departed at the end of its dwell.
  0 minutes apart, the first two arrive together and the last two as soon as they depart.

### Emergency Egress Time

The model also computes an "emergency egress time":
the time for everyone on both trains to go upstairs at 19 pax/min/ft.

- That's Fruin's maximum ascending stair flow, 18.9 pax/min/ft
  ([p. 9](https://onlinepubs.trb.org/Onlinepubs/hrr/1971/355/355-001.pdf#page=9)),
  reached at about 3 sq ft/pax, the edge of LOS F.
- It's more than the 17 pax/min/ft LOS E capacity,
  and more than NFPA 130's 1.41 pax/in/min, i.e. 16.9 pax/min/ft,
  for evacuating up stairs ([TCQSM p. 10-79](https://onlinepubs.trb.org/onlinepubs/tcrp/tcrp_rpt_165ch-10.pdf#page=83)).
- It ignores walking time, so it's a lower bound.

## Assumptions

Beyond the numbers above, which are all in `Assumptions`, the model assumes the following.
Each is marked by which way it biases the results:
**optimistic** (less crowding than reality), **pessimistic** (more),
**neutral** (neither, e.g. it only simplifies the bookkeeping),
or **unclear** (it could go either way).

### Platform

- **Neutral:** The platform is a single island platform serving two tracks.
- **Optimistic:** 75% of its area is usable, the rest taken by columns, stairs, and other obstructions
  (see [Usable Platform Area](#usable-platform-area)).
- **Optimistic:** Passengers are spread evenly over the whole usable area,
  so local crowding, e.g. at the foot of the stairs or at the doors, isn't modeled.
- **Neutral:** Space per passenger counts everyone on the platform:
  arriving passengers heading up and departing passengers waiting to board.

### Stairs and Escalators

- **Pessimistic:** All VCEs are treated as stairs, even escalators, which have higher capacities.
- **Pessimistic:** One VCE per platform is excluded, e.g. an escalator running the other way.
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
- **Optimistic:** The emergency egress time only counts the passengers on both trains,
  not the departing passengers,
  and ignores walking time.

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

Modeling each VCE's own queue with walking distances would fix these,
but needs each VCE's position and width on each platform.

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

Each platform's `total_vce_width` comes from the ETA report, and its source is unknown.
Platforms 10 and 11's match the
[Moynihan Station EA's](https://web.archive.org/web/20241011135133/https://cdn.esd.ny.gov/subsidiaries_projects/msdc/Data/NEPA/04_4%20StationPedCirculation.pdf#page=22)
2008 per-platform stair capacities divided by 17 pax/min/ft,
but platform 3's matches the EA's platform 1, not platform 3,
and they probably don't reflect the VCEs added since,
with the West End Concourse's expansion and Moynihan Train Hall.

## VCE Width Data

A stair-by-stair model needs each VCE's width and position on each platform.
The best public source of widths found so far is the NY Penn Station Master Plan's
tables of VCE widths for each platform under each reconstruction alternative,
transcribed in [`data/master_plan_vce_widths.csv`](./data/master_plan_vce_widths.csv):

- Alternatives 1 and 2 from the [August 2020 draft Alternatives Report](https://liamblank.com/wp-content/uploads/2026/07/R-Master-Plan-MTA-NJT-AMT-20_0814-PSMP-Alternatives-Report.pdf)
  ([p. 16](https://liamblank.com/wp-content/uploads/2026/07/R-Master-Plan-MTA-NJT-AMT-20_0814-PSMP-Alternatives-Report.pdf#page=22) and [p. 28](https://liamblank.com/wp-content/uploads/2026/07/R-Master-Plan-MTA-NJT-AMT-20_0814-PSMP-Alternatives-Report.pdf#page=34)).
    Its tables for Alternatives 3 and 4 ([p. 40](https://liamblank.com/wp-content/uploads/2026/07/R-Master-Plan-MTA-NJT-AMT-20_0814-PSMP-Alternatives-Report.pdf#page=46) and [p. 52](https://liamblank.com/wp-content/uploads/2026/07/R-Master-Plan-MTA-NJT-AMT-20_0814-PSMP-Alternatives-Report.pdf#page=58)) are identical.
- Alternatives 3 and 4 from the same draft.
- Alternative 3 from the April 2021 final report ([p. 117](https://esd.ny.gov/sites/default/files/CACWG-Meetings-8-9-Q-A-09-08-21.pdf#page=7)),
  as reproduced in the Empire Station Complex Q&A.

Each VCE is marked as a stair or escalator, and as new or existing.
Widths are in inches, listed in the tables' order, which is west to east along the platform.

[`scripts/extract_vce_positions.py`](./scripts/extract_vce_positions.py)
extracts each VCE's approximate position from the draft's platform-level plans
([p. 17](https://liamblank.com/wp-content/uploads/2026/07/R-Master-Plan-MTA-NJT-AMT-20_0814-PSMP-Alternatives-Report.pdf#page=23), [p. 29](https://liamblank.com/wp-content/uploads/2026/07/R-Master-Plan-MTA-NJT-AMT-20_0814-PSMP-Alternatives-Report.pdf#page=35),
[p. 41](https://liamblank.com/wp-content/uploads/2026/07/R-Master-Plan-MTA-NJT-AMT-20_0814-PSMP-Alternatives-Report.pdf#page=47), and [p. 53](https://liamblank.com/wp-content/uploads/2026/07/R-Master-Plan-MTA-NJT-AMT-20_0814-PSMP-Alternatives-Report.pdf#page=59)).
They're vector drawings with each VCE drawn as a rectangle color-coded by type and status,
and a scale bar to convert to feet.
The script matches the `n`th VCE from the west on each platform to the `n`th VCE in its table,
since the plans are too small to measure widths from,
and writes [`data/master_plan_vce_positions.csv`](./data/master_plan_vce_positions.csv),
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
[`data/master_plan_existing_vces.csv`](./data/master_plan_existing_vces.csv).
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

This data can't yet replace `total_vce_width`:

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
- **It doesn't match `total_vce_width`.**
  E.g. platform 3's VCEs sum to 412 to 458 in. across the alternatives, even with their new VCEs,
  but the model uses 42.5 ft (510 in.) today,
  and platform 6's sum to 437 to 458 in., but the model uses 48.168 ft (578 in.).
  The source of the model's widths is unknown.
  Platform 3's 5 existing VCEs in the Master Plan total 233 in. (19.4 ft):
  3 stairs (165 in.) and 2 escalators (68 in.).
  At 17 pax/min/ft for the stairs and typical escalator capacities,
  that's roughly the Moynihan Station EA's 437 pax/min for platform 3 in 2008,
  while the model's 42.5 ft is 722 pax/min, the EA's figure for platform 1.
  But the Master Plan probably doesn't include the West End Concourse's VCEs,
  and the EA's data predates Moynihan Train Hall,
  so platform 3's total width today is probably more than 19.4 ft.
  With the VCEs on NJ Transit's scaled PCIP Phase 2 plan, whose widths are mostly
  [estimated](#estimated-widths), it's about 550 in. (45.8 ft), a little more than the model's 42.5 ft.
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
[`scripts/extract_directory_vces.py`](./scripts/extract_directory_vces.py)
extracts and classifies these icons into
[`data/njt_directory_vces.csv`](./data/njt_directory_vces.csv),
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
[`scripts/estimate_vce_widths.py`](./scripts/estimate_vce_widths.py)
finds every run of treads on the platforms, merges a stair's flights,
and measures each VCE's width and position into
[`data/estimated_vce_widths.csv`](./data/estimated_vce_widths.csv):

- Positions are in the Master Plan's frame, feet east of its plans' west edge.
  The two drawings register to within about 1.4 ft:
  the platforms' east ends on them are all the same distance apart.
- VCEs matching a Master Plan VCE of the same type within 15 ft
  have `width_status` `master plan`, with the Master Plan's width.
  The rest, 46 of 61, are `estimated`, with only the sheet's width.
- For the 10 matched stairs, the sheet's widths differ from the Master Plan's by up to 14 in.
  (a median of 4 in.), and for the 5 matched escalators, within 3 in.
- Escalators' treads are their steps, narrower than their balustrades.
- The platforms' labels hide what's under them,
  including an escalator the Master Plan has about 230 ft along each of platforms 3 to 8.

| Platform | Stairs on the sheet | Escalators on the sheet | Their total width | Of which estimated | Master Plan VCEs not matched on the sheet |
|---|---|---|---|---|---|
| 1 | 6 | 2 | 411 in. (34.2 ft) | 235 in. | 46 in. escalator at 688 ft; 44/57 in. stair at 689 ft; 46 in. escalator at 774 ft; 52 in. stair at 776 ft |
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
- Escalators are still counted by width, not by an escalator's capacity.

### Field Survey

Since no public source has every VCE's width,
[`data/field_survey.csv`](./data/field_survey.csv) is a sheet for measuring them in person,
made by [`scripts/make_field_survey.py`](./scripts/make_field_survey.py).
It lists each platform's VCEs expected from the PCIP Phase 2 plan (platforms 1 to 8),
the 2022 directory, and the Master Plan,
each sorted west to east, since they can't all be aligned reliably,
starting with platform 3, the ETA report's focus,
then platform 11, which has no width data.
Surveyors fill in the columns after `position_ft`, feet east of the Master Plan's plans' west edge:

- `found`: yes, no, or the `id` of another row it duplicates.
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

The ETA report quotes the max up rate, time at capacity, and taper time
for platform 3 with trains 2 minutes apart:
12.68 pax/s for 0:13, tapering at 5:55, with Penn Reconstruction,
and 12.04 pax/s for 0:32, tapering at 6:04, without it.

### Current Results

This table is generated by `./vce_two_trains_alight_and_board.py --update-readme`:

<!-- results-table:start -->
| Platform | Headway | VCE width | Arrivals | Dwell | Taper time | Clear time | Boarded time | Time at capacity | Max up rate (pax/s) | Max pax on platform | Max density (pax/m²) |
|---|---|---|---|---|---|---|---|---|---|---|---|
| 3 | 0:00 | 43 ft | 0:00, 0:00, 3:37, 3:37 | 3:37, 3:37, 0:44, 0:44 | 7:34 | 8:30 | 4:21 | 6:13 | 12.18 | 3774 | 3.34 (D) |
| 3 (recon) | 0:00 | 44.75 ft | 0:00, 0:00, 3:38, 3:38 | 3:38, 3:38, 0:44, 0:44 | 6:56 | 7:11 | 7:12 | 7:04 | 12.68 | 3803 | 3.37 (D) |
| 6 | 0:00 | 30.83 ft | 0:00, 0:00, 7:13, 7:13 | 7:13, 7:13, 0:43, 0:43 | 13:14 | 14:20 | 13:04 | 11:38 | 8.74 | 4000 | 3.48 (D) |
| 10 | 0:00 | 70.58 ft | 0:00, 0:00, 0:43, 0:43 | 0:43, 0:43, 0:43, 0:43 | 5:09 | 5:24 | 1:26 | 5:24 | 20.00 | 5740 | 1.78 (D) |
| 11 | 0:00 | 43.58 ft | 0:00, 0:00, 4:31, 4:31 | 4:31, 4:31, 0:43, 0:43 | 8:39 | 8:54 | 5:14 | 8:44 | 12.35 | 4314 | 3.13 (D) |
| 3 | 2:00 | 43 ft | 0:00, 2:00, 4:00, 7:12 | 0:44, 5:12, 3:12, 0:44 | 8:52 | 9:31 | 7:56 | 5:10 | 12.18 | 1367 | 1.21 (C) |
| 3 (recon) | 2:00 | 44.75 ft | 0:00, 2:00, 4:00, 6:53 | 0:44, 4:53, 2:53, 0:44 | 8:25 | 8:40 | 7:37 | 7:04 | 12.68 | 1332 | 1.18 (C) |
| 6 | 2:00 | 30.83 ft | 0:00, 2:00, 4:00, 11:33 | 0:43, 9:33, 7:33, 0:43 | 14:25 | 15:13 | 14:21 | 11:34 | 8.74 | 2470 | 2.15 (D) |
| 10 | 2:00 | 70.58 ft | 0:00, 2:00, 4:00, 6:00 | 0:43, 0:43, 0:43, 0:43 | 7:06 | 7:21 | 6:43 | 5:24 | 20.00 | 1360 | 0.42 (A) |
| 11 | 2:00 | 43.58 ft | 0:00, 2:00, 4:00, 8:10 | 0:43, 6:10, 4:10, 0:43 | 10:07 | 10:22 | 10:22 | 8:44 | 12.35 | 1613 | 1.17 (C) |
| 3 | 5:00 | 43 ft | 0:00, 5:00, 10:00, 15:00 | 0:44, 0:44, 0:44, 0:44 | 16:40 | 17:19 | 15:44 | 5:08 | 12.18 | 1348 | 1.19 (C) |
| 3 (recon) | 5:00 | 44.75 ft | 0:00, 5:00, 10:00, 15:00 | 0:44, 0:44, 0:44, 0:44 | 16:32 | 16:47 | 15:44 | 7:04 | 12.68 | 1319 | 1.17 (C) |
| 6 | 5:00 | 30.83 ft | 0:00, 5:00, 10:00, 15:00 | 0:43, 0:43, 0:43, 0:43 | 17:52 | 18:40 | 15:43 | 10:48 | 8.74 | 1727 | 1.50 (C) |
| 10 | 5:00 | 70.58 ft | 0:00, 5:00, 10:00, 15:00 | 0:43, 0:43, 0:43, 0:43 | 16:06 | 16:21 | 15:43 | 5:24 | 20.00 | 1340 | 0.42 (A) |
| 11 | 5:00 | 43.58 ft | 0:00, 5:00, 10:00, 15:00 | 0:43, 0:43, 0:43, 0:43 | 16:57 | 17:12 | 15:43 | 8:44 | 12.35 | 1600 | 1.16 (C) |
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
- **Used each VCE's width from [`data/estimated_vce_widths.csv`](./data/estimated_vce_widths.csv)**
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
