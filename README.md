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

Everything the model assumes, like stair capacity, passenger loads, and LOS thresholds,
is in the `Assumptions` dataclass at the top of
[`vce_two_trains_alight_and_board.py`](./vce_two_trains_alight_and_board.py),
with each assumption's units and source.
To test how sensitive the results are to one,
pass a scenario's `Params` different `assumptions`, e.g. `Assumptions(stair_capacity=15)`.

## Checking

We also use `ruff` for formatting and linting and `ty` and `pyrefly` for type checking.
To run these, which are also checked in CI,
you can run

```sh
uv run ruff format # format
uv run ruff check # lint
uv run ty check # type check
uv run pyrefly check # type check
```

These same checks also run as `pre-commit` hooks.
To install them, run

```sh
uv run pre-commit install
```

## How the Model Works

The sources cited below are:

- [ETA's report](https://www.etany.org/penn-station-can-handle-the-load), which used this model.
- [Fruin, "Designing for Pedestrians: A Level-of-Service Concept"](https://onlinepubs.trb.org/Onlinepubs/hrr/1971/355/355-001.pdf),
  Highway Research Record 355 (1971), cited in the code,
  with page numbers of the PDF.
- [The Transit Capacity and Quality of Service Manual (TCQSM), 3rd edition, chapter 10](https://onlinepubs.trb.org/onlinepubs/tcrp/tcrp_rpt_165ch-10.pdf),
  TCRP Report 165 (2013), with the manual's page numbers.

The model is the one used in ETA's report
[Penn Station Can Handle the Load](https://www.etany.org/penn-station-can-handle-the-load).
It simulates one island platform served by two tracks, one second at a time,
and tracks four groups of passengers:
passengers aboard each train,
arrived passengers on the platform heading upstairs,
departing passengers on the platform waiting to board each train,
and departing passengers still upstairs on the concourse.

Each second:

1. **Alighting.** Once a train arrives, its passengers step off at 1 pax/s per single-door equivalent.
   Every scenario uses 1,620 passengers on 40 doors,
   i.e. a crush-loaded 10-car NJ Transit MultiLevel EMU, the worst case.
   (A 12-car LIRR train has more doors, and LIRR platforms are generally wider.)
2. **Going upstairs.** Arrived passengers leave via the vertical circulation elements (VCEs),
   i.e. the stairs and escalators, all treated as stairs,
   with one VCE per platform excluded.
   They queue at the stairs, which discharge them at LOS E capacity,
   17 pax/min per foot of VCE width ([Fruin, p. 14](https://onlinepubs.trb.org/Onlinepubs/hrr/1971/355/355-001.pdf#page=14)), as long as anyone is queued,
   as in the TCQSM's stair queuing procedure ([p. 10-51](https://onlinepubs.trb.org/onlinepubs/tcrp/tcrp_rpt_165ch-10.pdf#page=55)).
   So there's no gradual taper: the stairs stay at capacity until the platform is clear.
   The report's "taper time" is when the remaining arrived passengers fit in the stair queues,
   20 ft of queue in front of the total VCE width at 5 sqft/pax
   (the TCQSM's stair queuing space, [p. 10-51](https://onlinepubs.trb.org/onlinepubs/tcrp/tcrp_rpt_165ch-10.pdf#page=55)),
   which is now always a fixed ~15 s before the clear time,
   but it's kept in the results table to compare with the report.
   The few seconds of walking from the doors to the stairs are ignored.
3. **Coming downstairs.** Departing passengers queue upstairs and come down with whatever stair capacity
   the upward flow leaves, with both directions sharing LOS E capacity, 17 pax/min/ft.
   But nobody comes down while the upward flow is worse than LOS C, 10 pax/min/ft,
   per the ETA report.
   The stairs are split between the two trains' departing passengers
   in proportion to how many are still upstairs.
4. **Boarding.** Departing passengers on the platform board a train that has arrived
   once every arriving passenger has alighted, at 1 pax/s per door.

Space per passenger is the usable platform area (75% of the platform's area)
divided by everyone on the platform,
graded with Fruin's LOS for queuing and waiting areas
(A > 13, B > 10, C > 7, D > 3, E > 2 sqft/pax, else F),
which the TCQSM applies to station platforms
([Exhibit 10-32, p. 10-55](https://onlinepubs.trb.org/onlinepubs/tcrp/tcrp_rpt_165ch-10.pdf#page=59)),
since most passengers on the platform are waiting, either to board or in the stair queues.
The upward stair flow is graded with Fruin's stair LOS
(A ≤ 5, B ≤ 7, C ≤ 10, D ≤ 13, E ≤ 17 pax/min/ft, else F;
[Fruin, pp. 12–14](https://onlinepubs.trb.org/Onlinepubs/hrr/1971/355/355-001.pdf#page=12)).

Every scenario has 200 departing passengers per train already on the platform
and another 200 per train upstairs.
The second train arrives 2 minutes (120 s) or 5 minutes (300 s) after the first,
except on platform 6, where both trains arrive at once.

The model also prints an "emergency egress time":
the time for everyone on both trains to go upstairs at 19 pax/min/ft,
the maximum ascending stair flow in Fruin's paper
(18.9 pax/min/ft, [p. 9](https://onlinepubs.trb.org/Onlinepubs/hrr/1971/355/355-001.pdf#page=9)),
reached at about 3 sqft/pax, the edge of LOS F.
This is more than the 17 pax/min/ft LOS E capacity used everywhere else,
and more than NFPA 130's 1.41 pax/in/min, i.e. 16.9 pax/min/ft,
for evacuating up stairs ([TCQSM p. 10-79](https://onlinepubs.trb.org/onlinepubs/tcrp/tcrp_rpt_165ch-10.pdf#page=83)),
and it ignores walking time, so it's a lower bound.

## Assumptions

Beyond the numbers above, the model assumes the following.
Where it's clear which way an assumption biases the results,
it's marked as optimistic (less crowding than reality) or pessimistic (more).

### Platform

- The platform is a single island platform serving two tracks.
- 75% of its area is usable, the rest taken by columns, stairs, and other obstructions.
- Passengers are spread evenly over the whole usable area,
  so local crowding, e.g. at the foot of the stairs or at the doors, isn't modeled (optimistic).
- Space per passenger counts everyone on the platform:
  arrived passengers heading up and departing passengers waiting to board.

### Stairs and Escalators

- All VCEs are treated as stairs, even escalators, which have higher capacities (pessimistic).
- One VCE per platform is excluded, e.g. an escalator running the other way (pessimistic).
- All VCEs act as one pooled queue:
  passengers spread across them in proportion to their widths,
  with no preference for any exit, e.g. toward 7th Avenue (optimistic).
- Stair capacity is linear in width, at 17 pax/min/ft,
  though the TCQSM notes capacity is really stepped by the number of pedestrian lanes
  ([p. 10-49](https://onlinepubs.trb.org/onlinepubs/tcrp/tcrp_rpt_165ch-10.pdf#page=53)).
- Stair capacity doesn't depend on the stair's rise,
  though long climbs slow people down ([p. 10-51](https://onlinepubs.trb.org/onlinepubs/tcrp/tcrp_rpt_165ch-10.pdf#page=55)) (optimistic),
  or on luggage, strollers, or wheelchairs (optimistic).
- Walking from the doors to the stairs takes no time;
  passengers join the stair queue the same second they alight (optimistic, by a few seconds).
- The concourse upstairs never backs up, so the stairs always discharge (optimistic).
- Both directions share the stairs' 17 pax/min/ft,
  but nobody comes down while more than 10 pax/min/ft is going up.

### Trains

- Every arriving passenger alights; nobody stays on a through train (pessimistic).
- Passengers alight and board at 1 pax/s per single-door equivalent,
  starting the second after the train arrives, with no delay for the doors to open (optimistic).
- Nobody boards until everyone has alighted.
- Trains don't depart within the 600 s simulated,
  so dwell times aren't modeled, and nobody misses a train.

### Departing Passengers

- Each train has 400 departing passengers, all present at the start:
  200 on the platform and 200 upstairs.
  None arrive during the simulation, even with 5-minute headways (optimistic).
- Departing passengers come downstairs even before their train arrives,
  and wait for it on the platform, adding to the crowding.
- The stairs are split between the two trains' departing passengers
  in proportion to how many of each are still upstairs.

### Simulation

- It steps through time 1 s at a time, with fractional passengers.
- Each second, passengers alight, then go upstairs, then come downstairs, then board,
  each using the counts left by the previous step.
- The stair LOS grades only the upward flow, not the downward flow.
- The emergency egress time only counts the passengers on both trains,
  not the departing passengers,
  and ignores walking time (optimistic).

## Limitations

The model is much simpler than a pedestrian microsimulation,
like those typically used in detailed station planning,
and several of its simplifications overstate platform capacity.
Treat its results as optimistic until these are addressed.

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
This model clears platform 3's arrived passengers in 4.5 minutes (269 s),
so it's likely substantially optimistic,
though the two aren't exactly comparable:
7.9 minutes is the worst case across all platforms and simulation runs.

The FRA attributes long clearance times to the same things this model leaves out
([p. 3-33](https://railroads.dot.gov/sites/fra.dot.gov/files/2026-07/2026.07.13_Penn%20Station%20SOS_Phase%20I%20Report_FINAL.pdf#page=44)):
queues at the base of VCEs, uneven use of VCEs, and platform clutter
reducing the usable width.

### Stairs Are One Pooled Queue

The model treats all of the VCEs as one queue that discharges at full capacity
until the last arrived passenger leaves.
This follows the TCQSM's stair queuing procedure ([p. 10-51](https://onlinepubs.trb.org/onlinepubs/tcrp/tcrp_rpt_165ch-10.pdf#page=55))
and NFPA 130's evacuation check ([p. 10-79](https://onlinepubs.trb.org/onlinepubs/tcrp/tcrp_rpt_165ch-10.pdf#page=83)),
which size stairs for a given demand, but it's optimistic for a whole platform:

- **Stairs empty unevenly.**
  Stairs near the ends of the platform, or far from the busiest doors,
  run out of queue while others still have one,
  so the total flow drops below capacity before the platform clears.
  This is likely the largest optimistic bias in the model.
- **Passengers prefer some exits**, e.g. toward 7th Avenue, as the ETA report notes,
  concentrating queues at fewer stairs.
- **Walking from the doors to the stairs takes no time.**
  This only shifts the results by a few seconds, but also ignores
  passengers crossing through crowds of waiting passengers.

Modeling each VCE's own queue with walking distances would fix these,
but needs each VCE's position and width on each platform.
Until then, running at a reduced effective stair capacity, e.g. 80–90%,
would give a rough bound.

### Platform Crowding Is Graded Against the Whole Platform

The model grades all passengers on the platform with Fruin's LOS for queuing areas
(TCQSM [Exhibit 10-32, p. 10-55](https://onlinepubs.trb.org/onlinepubs/tcrp/tcrp_rpt_165ch-10.pdf#page=59)).
But the TCQSM's platform sizing procedure ([p. 10-56](https://onlinepubs.trb.org/onlinepubs/tcrp/tcrp_rpt_165ch-10.pdf#page=60))
only applies that to passengers *waiting* to board, and adds separate areas for:

- walkway width for arriving passengers, graded as a walkway,
- queue storage at the stairs, and
- an 18 in. buffer along each platform edge.

The model lumps everyone into one area, so its crowding grades are optimistic,
especially while arriving passengers are walking to the stairs.
With Fruin's walkway LOS ([p. 7](https://onlinepubs.trb.org/Onlinepubs/hrr/1971/355/355-001.pdf#page=7)), which the model originally used,
platform 3's most crowded moment is LOS E instead of C.

### Usable Platform Area

The model uses 75% of the platform's area,
leaving 25% for columns, stairwells, and other obstructions.
But the TCQSM's 18 in. edge buffers alone take 3 ft of an 18 ft platform, about 17%,
leaving only 8% for everything else, which is likely too little on a narrow platform
with stairwells in it.
So the usable area, and the space per passenger, are likely overstated.

### The ETA Report's Results Aren't a Reliable Baseline

The original model reproduced the report's results,
but only because two bugs partly canceled out:
its stair flow was about 1.4x too high from a units error,
but also tapered off as the platform emptied,
from applying Fruin's stair equation to the platform's density.
Fixing only the units made the results much worse (see the history below),
and neither version modeled the stairs correctly,
so the report's results shouldn't be read as either conservative or optimistic.

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
  With the West End Concourse's 2 stairs' [estimated widths](#estimated-widths),
  it's about 464 in. (38.7 ft), a little less than the model's 42.5 ft.
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

Until the VCEs are measured, some widths no source lists can be estimated from a scaled drawing.
NJ Transit's PCIP Phase 2 drawings include an
[existing concourse-level plan](https://liamblank.com/wp-content/uploads/2026/09/pcip-2-conceptual-design-preliminary-drawings.pdf#page=45)
(sheet A-001, November 2020), a vector drawing at 1" = 40' showing platforms 1 to 8
and each stair's treads, so a tread's length is roughly its stair's width.
[`scripts/estimate_vce_widths.py`](./scripts/estimate_vce_widths.py)
measures them into [`data/estimated_vce_widths.csv`](./data/estimated_vce_widths.csv),
with every width marked `estimated`.

So far, it estimates the 2 stairs from the West End Concourse down to platform 3,
the ones the Master Plan's tables don't include:

| Stair | Estimated width | Position (ft west of platform 3's east end) |
|---|---|---|
| West End Concourse, west side | 109 in. | 780 to 802 |
| West End Concourse, east side | 122 in. | 715 to 726 |

The east side's is T-shaped, splitting into two 60 in. flights east and west along the platform.
With the Master Plan's 5, platform 3's VCEs total about 464 in. (38.7 ft).

These are only estimates:

- On the same sheet, the stairs that seem to match the Master Plan's measure within about 6 in.
  of its widths (e.g. 66 vs. 69 in. and 44 vs. 44 in.),
  but a drawn tread isn't necessarily the clear width between handrails.
- The sheet also shows more stairs over platform 3 than the Master Plan lists.
  Some are probably stairs between the upper and lower concourses drawn over the platform,
  but the sheet doesn't label them, so they aren't included.
- The escalators among the Master Plan's 5 are still counted by width,
  not by an escalator's capacity.

### Field Survey

Since no public source has every VCE's width,
[`data/field_survey.csv`](./data/field_survey.csv) is a sheet for measuring them in person,
made by [`scripts/make_field_survey.py`](./scripts/make_field_survey.py).
It lists each platform's VCEs expected from both the 2022 directory and the Master Plan,
each sorted west to east, since the two can't be aligned reliably,
starting with platform 3, the ETA report's focus,
then platform 11, which has no width data.
Surveyors fill in the columns after `master_plan_mid_ft`:

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

Each run writes a spreadsheet per scenario with the full time series and charts,
and prints this table of headline results.
Times are seconds after the first train arrives.

- **Max up rate:** the highest upstairs rate.
- **Time at capacity:** how long the upstairs rate is at LOS E capacity.
- **Taper time:** the last second more arrived passengers are on the platform than fit in the stair queues.
- **Clear time:** when the last arrived passenger leaves the platform, if within the 600 s simulated.
- **Boarded time:** when the last departing passenger boards, if within the 600 s simulated.
- **Max pax on platform** and **Min space/pax:** the most crowded moment on the platform, with its LOS.

The ETA report quotes the first three for platform 3 with trains 2 minutes apart:
12.68 pax/s for 13 s, tapering at 355 s, with Penn Reconstruction,
and 12.04 pax/s for 32 s, tapering at 364 s, without it.

### Current Results

| Platform | Headway | VCE width | Max up rate (pax/s) | Time at capacity | Taper time | Clear time | Boarded time | Max pax on platform | Min space/pax (sqft) |
|---|---|---|---|---|---|---|---|---|---|
| platform3 | 120 s | 42.5 ft | 12.04 | 269 s | 254 s | 269 s | 302 s | 1538 | 7.9 (C) |
| platform3 | 300 s | 42.5 ft | 12.04 | 268 s | 420 s | 435 s | 351 s | 1538 | 7.9 (C) |
| platform3_recon | 120 s | 44.75 ft | 12.68 | 255 s | 241 s | 256 s | 287 s | 1513 | 8.0 (C) |
| platform3_recon | 300 s | 44.75 ft | 12.68 | 254 s | 413 s | 428 s | 351 s | 1513 | 8.0 (C) |
| platform6 | 0 s | 48.168 ft | 13.65 | 237 s | 223 s | 238 s | 266 s | 3094 | 4.0 (D) |
| platform10 | 120 s | 70.58 ft | 20.00 | 162 s | 186 s | 201 s | 171 s | 1220 | 28.4 (A) |
| platform11 | 120 s | 43.58 ft | 12.35 | 262 s | 248 s | 263 s | 294 s | 1526 | 9.7 (C) |

### History

As bugs are fixed, the results table above is updated,
and each fix's effect is summarized here.

- **Original model** (as used in the ETA report):
  reproduces the report's platform 3 numbers to within a few seconds.
- **Fixed the stair equation's units:**
  Fruin's P is pax/min per foot of width, but was used as pax/s across all stairs,
  overstating stair flow by 60 / VCE width, about 1.4x.
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
  instead of slowing as a fixed 5,000 sqft concourse empties.
  Every scenario now finishes boarding: platform 3 with 2-minute headways by 325 s
  (309 s with Penn Reconstruction), where about 20 of 400 never boarded before.
- **Let passengers come down at up to LOS E capacity, 17 pax/min/ft, when few are going up,**
  since the report's 10 pax/min/ft rule only applies to flow in both directions.
  Platform 3 with 2-minute headways finishes boarding by 302 s instead of 325 s
  (287 s instead of 309 s with Penn Reconstruction).
