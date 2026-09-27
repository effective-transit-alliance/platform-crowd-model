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
This version gives the same results as that tag,
and has [several bugs](#known-bugs); this describes how it currently works, bugs included.

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
   - Every train carries 1,620 passengers,
     a seated 12-car NJ Transit train at 135 seats per car,
     from the Moynihan Station environmental assessment
     ([Table 4.4-10, p. 4.4-22](https://web.archive.org/web/20241011135133/https://cdn.esd.ny.gov/subsidiaries_projects/msdc/Data/NEPA/04_4%20StationPedCirculation.pdf#page=22);
     [Table 4.4-19, p. 4.4-47](https://web.archive.org/web/20241011135133/https://cdn.esd.ny.gov/subsidiaries_projects/msdc/Data/NEPA/04_4%20StationPedCirculation.pdf#page=47)).
   - It has 40 doors, as a 10-car train does,
     though its passengers fill a 12-car train
     (see [its known bug](#trains-have-a-12-car-trains-passengers-but-a-10-car-trains-doors)).
     (A 12-car LIRR train has more doors, and LIRR platforms are generally wider.)
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
   Departing passengers upstairs come down to the platform.
   - The trains' passengers split the VCE width
     in proportion to how many of each are still upstairs,
     and each train's share of the stairs carries the same share of the upward flow.
   - Stairs carry at most the LOS C/D boundary, 10 pax/min/ft, in both directions combined,
     so passengers come down with whatever their share of the upward flow leaves of that.
   - Their flow is also limited by Fruin's descending stair equation
     ([Fruin, p. 9](https://onlinepubs.trb.org/Onlinepubs/hrr/1971/355/355-001.pdf#page=9)),
     `P = (128M − 206)/M²` pax/min per foot of VCE width,
     where `M` is the space per departing passenger in a 5,000 sq ft concourse,
     minus their share of the upward flow.
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

- Each train has 400 departing passengers:
  200 already on the platform and 200 upstairs.
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
The [known bugs](#known-bugs) are listed separately.

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

- **Optimistic:** Each train has 400 departing passengers, all present at the start:
  200 on the platform and 200 upstairs.
  None arrive during the simulation, even for the fourth train with 5-minute headways,
  whose passengers wait 15 minutes for it.
  (See [its known bug](#departing-passengers-wait-on-the-platform-long-before-their-train).)
- **Pessimistic:** Departing passengers come downstairs even before their train arrives,
  and wait for it on the platform, adding to the crowding.

### Simulation

- **Neutral:** It steps through time 1 s at a time, with fractional passengers,
  until the last train departs and the platform clears.
- **Neutral:** Each second, passengers alight, then go upstairs, then come downstairs, then board,
  each using the counts left by the previous step.
- **Neutral:** The stair LOS grades only the upward flow, not the downward flow.
- **Optimistic:** The emergency egress time only counts the passengers on both trains,
  not the departing passengers,
  and ignores walking time.

## Known Bugs

These are errors in the model, as opposed to simplifications,
each checked against its source.
Some of them partly canceled out,
which is how the original model reproduced the ETA report's results despite them,
so neither those results nor the ones below should be read as either conservative or optimistic
until they're all fixed.

The bugs in the downward flow bias crowding and boarding in opposite directions:
slower downward flow keeps departing passengers upstairs longer,
so there's less crowding on the platform (optimistic),
but they board later (pessimistic).

### Downward flow is capped at 10 pax/min/ft even when nobody is going up

Once the upward flow is light, the report's rule about flow in both directions no longer applies,
so passengers should be able to come down at up to LOS E capacity, 17 pax/min/ft,
but the model still caps them at 10 (optimistic for crowding, pessimistic for boarding).

### Downward flow slows as a fixed concourse empties

The downward flow applies Fruin's equation to a fixed 5,000 sq ft concourse
divided by the departing passengers still upstairs.
So it's held down while many are waiting,
and slows as they come down, so the last ones take many minutes to
(optimistic for crowding, pessimistic for boarding).
The 5,000 sq ft isn't sourced,
and a queue at the top of the stairs, like the one at the bottom, would discharge at capacity.

### Departing passengers wait on the platform long before their train

Every train's departing passengers are there from the start,
200 on the platform and 200 upstairs, who come down as soon as the stairs allow,
even for the fourth train with 5-minute headways, which isn't due for 15 minutes.
The ETA report put them all there at the start, but with only 2 trains.
At Penn Station, passengers wait in the concourse until their track is announced,
so this fills the platform with passengers who wouldn't be there yet (pessimistic for crowding),
but has them already at the doors when later trains arrive,
so those trains' dwells are only as long as alighting and boarding take (optimistic for dwells).

### Trains have a 12-car train's passengers but a 10-car train's doors

Every train carries 1,620 passengers, a seated 12-car NJ Transit train,
but has only 40 doors, as a 10-car train does, instead of a 12-car train's 48,
so its passengers take longer to alight and board (pessimistic).
Platform 3's tracks only fit 10 cars,
so its trains should instead carry 1,350 passengers (pessimistic too).

## Limitations

Beyond its bugs,
the model is much simpler than a pedestrian microsimulation,
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
| 3 | 0:00 | 42.5 ft | 0:00, 0:00, 24:09, 24:09 | 24:09, 24:09, 0:51, 0:51 | 28:23 | 28:38 | 31:25 | 8:58 | 12.04 | 3550 | 3.14 (D) |
| 3 (recon) | 0:00 | 44.75 ft | 0:00, 0:00, 22:57, 22:57 | 22:57, 22:57, 0:51, 0:51 | 26:58 | 27:13 | 29:50 | 8:30 | 12.68 | 3524 | 3.12 (D) |
| 6 | 0:00 | 48.168 ft | 0:00, 0:00, 21:19, 21:19 | 21:19, 21:19, 0:51, 0:51 | 25:02 | 25:17 | 27:43 | 7:54 | 13.65 | 3484 | 3.03 (D) |
| 10 | 0:00 | 70.58 ft | 0:00, 0:00, 14:32, 14:32 | 14:32, 14:32, 0:51, 0:51 | 16:59 | 17:14 | 18:54 | 5:24 | 20.00 | 3226 | 1.00 (B) |
| 11 | 0:00 | 43.58 ft | 0:00, 0:00, 23:34, 23:34 | 23:34, 23:34, 0:51, 0:51 | 27:42 | 27:57 | 30:38 | 8:44 | 12.35 | 3537 | 2.56 (D) |
| 3 | 2:00 | 42.5 ft | 0:00, 2:00, 24:09, 24:09 | 24:09, 22:09, 0:51, 0:51 | 28:23 | 28:38 | 31:25 | 8:58 | 12.04 | 3544 | 3.14 (D) |
| 3 (recon) | 2:00 | 44.75 ft | 0:00, 2:00, 22:57, 22:57 | 22:57, 20:57, 0:51, 0:51 | 26:58 | 27:13 | 29:50 | 8:30 | 12.68 | 3518 | 3.12 (D) |
| 6 | 2:00 | 48.168 ft | 0:00, 2:00, 21:19, 21:19 | 21:19, 19:19, 0:51, 0:51 | 25:02 | 25:17 | 27:43 | 7:53 | 13.65 | 3478 | 3.03 (D) |
| 10 | 2:00 | 70.58 ft | 0:00, 2:00, 14:32, 14:32 | 14:32, 12:32, 0:51, 0:51 | 16:59 | 17:14 | 18:54 | 5:24 | 20.00 | 3218 | 1.00 (B) |
| 11 | 2:00 | 43.58 ft | 0:00, 2:00, 23:34, 23:34 | 23:34, 21:34, 0:51, 0:51 | 27:42 | 27:57 | 30:38 | 8:44 | 12.35 | 3532 | 2.56 (D) |
| 3 | 5:00 | 42.5 ft | 0:00, 5:00, 24:10, 24:10 | 24:10, 19:10, 0:51, 0:51 | 28:24 | 28:39 | 31:26 | 8:57 | 12.04 | 3544 | 3.14 (D) |
| 3 (recon) | 5:00 | 44.75 ft | 0:00, 5:00, 22:57, 22:57 | 22:57, 17:57, 0:51, 0:51 | 26:58 | 27:13 | 29:50 | 8:29 | 12.68 | 3518 | 3.12 (D) |
| 6 | 5:00 | 48.168 ft | 0:00, 5:00, 21:19, 21:19 | 21:19, 16:19, 0:51, 0:51 | 25:02 | 25:17 | 27:43 | 7:53 | 13.65 | 3478 | 3.03 (D) |
| 10 | 5:00 | 70.58 ft | 0:00, 5:00, 14:32, 15:00 | 14:32, 9:32, 0:51, 0:51 | 16:59 | 17:14 | 18:54 | 5:24 | 20.00 | 2259 | 0.70 (A) |
| 11 | 5:00 | 43.58 ft | 0:00, 5:00, 23:35, 23:35 | 23:35, 18:35, 0:51, 0:51 | 27:43 | 27:58 | 30:39 | 8:44 | 12.35 | 3532 | 2.56 (D) |
<!-- results-table:end -->

### History

As bugs are fixed, the results table above is updated,
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
