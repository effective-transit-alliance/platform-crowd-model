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

Building the model's Rust core also needs a Rust toolchain,
which can be installed with [`rustup`](https://rustup.rs).

With both installed, you can then just run the model:

```sh
uv run platform-crowd-model run
```

In doing so, `uv` will also install dependencies and set up a virtual environment.

Each run prints a table of headline results.
To also save each scenario's full time series as CSVs and its charts as an SVG in `output/`,
including each VCE's queue and upward flow each second,
and a list of its VCEs with when each one's queue last empties, run

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
     9 cars on platforms 1 and 2, 10 on platforms 3 and 9, and 12 on the rest,
     or with Penn Transformation's extensions, 10 on platforms 1 and 2.
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
     (see [VCE Width Data](#vce-width-data)).
     - The train's doors are spread evenly along it,
       and it stops wherever on the platform the most passengers on it at once is fewest,
       so it can be evacuated soonest under NFPA 130,
       of the positions where the longest of the trains' dwells is within 1:00 of the shortest it can be,
       and of those within 1 passenger of the fewest, where the longest and then the total dwell are shortest,
       trying every position a car length (85') apart,
       then every position 5' apart within a car length of the best of those.
     - Each second, each door's alighting passengers walk to the quickest VCE:
       the one they could go up soonest,
       i.e. the later of their walk there and when everyone already queued at or walking to it has gone up,
       since its queue drains while they walk.
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
  0 minutes apart, the first two arrive together and the last two as soon as they depart,
  or on platform 9, each as soon as the one before it departs.
- Every platform, 1 to 11, is modeled both as it is and as Penn Transformation would leave it
  (see [Penn Transformation](#penn-transformation)):
  with its new VCEs, platforms 1 to 3 extended to the west for 10-car trains,
  and every platform decluttered, adding the FRA's percentage more area,
  and the extension's length at the platform's average width.
- Platform A, a new platform PCIP Phase 1 proposes south of platform 1,
  is modeled, too (see [Platform A](#platform-a)).

### NFPA 130 Evacuation

NFPA 130 (2026 edition, 5.3.3.1) requires "sufficient egress capacity to evacuate the platform occupant load
… from the station platform in 4 minutes or less".
Its Annex C computes this as the platform occupant load divided by the egress capacity,
without walking time, which is part of its separate 6-minute limit to reach a point of safety (5.3.3.2).
NFPA 130's text isn't free to read, so these are from its 2007 edition,
with the changes since from NFPA's public [revision reports](https://docinfofiles.nfpa.org/files/AboutTheCodes/130/),
e.g. the [2023 edition's First Draft Report](https://docinfofiles.nfpa.org/files/AboutTheCodes/130/130_A2022_FKT_AAA_FRReport.pdf).
Where NFPA 130 follows a value with one in other units, "the first stated value shall be regarded as the requirement",
and the other "might be approximated" (1.5.2).
It states its metric values first, so the model uses those, converted exactly,
e.g. its travel distance limit is 100 m, which is 328'1", not the 325' it gives with it.
In 2009, its committee rejected changing that limit to 91.5 m to match the then 300',
since it "does not reflect the degree of precision intended by the standard".

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
- Stairs and stopped escalators carry 0.0555 pax/min per mm of width (16.92 pax/min/ft) (5.3.5.3).
- The widest escalator is out of service, taken as the one "having the most adverse effect upon egress capacity",
  and escalators provide at most half of the egress capacity (5.3.5.4, 5.3.5.6).
  Since every stair and escalator has the same capacity per width,
  the widest is always the most adverse for egress capacity, even with that half limit.
  No platform's escalators come close to half: they provide at most 44%, on platform 6.

NFPA 130 also requires evacuating "from the most remote point on the platform to a point of safety in 6 minutes or less" (5.3.3.2).
The model checks this per its Annex C, at the same moment:

- The point of safety is the concourse,
  which NFPA 130 allows in an enclosed station like Penn Station
  only where its emergency ventilation protects the concourse from a train fire at the platform,
  as confirmed by an engineering analysis (5.3.3.4).
- The farthest occupant walks along the platform to their nearest VCE at 37.8 m/min (124 fpm) (5.3.4.4),
  from either end of the platform or from halfway between two VCEs.
  That's at most 258', on platform 3 with Penn Transformation's extension,
  both with the escalator out of service (see below) and with every VCE in service,
  within NFPA 130's 100 m (328'1") limit on travel distance (5.3.3.5),
  which doesn't take an escalator out of service.
  The results table checks this, too.
- They wait there for the rest of the platform's flow time, if any (Annex C's W_p = F_p − T_p),
  so they reach the VCE at the longer of their walk and the 4-minute check's flow time.
- They climb 16'9¼", from the existing platforms to the existing concourse
  on the PCIP Phase 2 plan's north-south cross section (sheet A-213, PDF page 54),
  at a vertical speed of 14.63 m/min (48 fpm) (5.3.5.3), taking 21 s.
- The same escalator is out of service here as for the 4-minute check,
  the one "having the most adverse effect upon egress capacity" (5.3.5.4), i.e. the widest,
  both for the walk and for the flow time.
  NFPA 130 doesn't say which to choose if the widest tie,
  so the model chooses the one that gets the farthest occupant to the concourse soonest.
  This only matters on platform 3 with Penn Transformation,
  whose new westernmost escalator is as wide as its widest existing one:
  with the new one out of service, the farthest occupant would walk 381' instead of 258',
  reaching the concourse a minute later.

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

- **Unclear:** Escalators' capacities are the TCQSM's nominal ones at 90 ft/min,
  but their speeds aren't known,
  and their tread widths are measured from a drawing, so they may be off by a few inches,
  which matters at the 2'8" boundary between 34 and 72 pax/min.
- **Unclear:** Which escalators run which way isn't known,
  so they follow the rules above.
  Which escalator goes down is based on AM peak demand toward 7th Avenue,
  and when the extra ones reverse isn't from any source.
- **Unclear:** Most VCEs' widths are estimated from a scaled plan,
  and platforms 9 to 11's positions are partly from a schematic map
  (see [VCE Width Data](#vce-width-data)).
- **Unclear:** Penn Transformation's and Platform A's new VCEs' widths aren't published,
  so they're assumed (see [Penn Transformation](#penn-transformation) and [Platform A](#platform-a)).
- **Optimistic:** Arriving passengers know which VCE is quickest,
  with no preference for any exit, e.g. toward 7th Avenue
  (see [Passengers Only Prefer the Quickest VCE](#passengers-only-prefer-the-quickest-vce)).
- **Optimistic:** Trains stop at the best position on the platform for evacuating it,
  within 1:00 of their best dwells, with their doors spread evenly along their length.
- **Unclear:** That 1:00 isn't from any source.
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
- **Optimistic:** The NFPA 130 evacuation time uses the simulated peak occupants,
  without NFPA 130's accumulation of waiting passengers over a doubled headway, or standees.
- **Optimistic:** The NFPA 130 time to the concourse treats the concourse as a point of safety,
  which NFPA 130 only allows where an engineering analysis shows it's protected (5.3.3.4).
- **Unclear:** The NFPA 130 time to the concourse takes the VCEs' positions from plans and a schematic map,
  only accurate to within about 10' to 20' (see [VCE Width Data](#vce-width-data)).

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

### Passengers Only Prefer the Quickest VCE

Each VCE has its own queue, with walking distances to it,
so stairs near the ends of the platform or far from the busiest doors can run dry
while others still have a queue.
But arriving passengers choose perfectly:
each second, each door's passengers know every VCE's queue, even hundreds of feet away,
and walk to the one they can go up soonest,
so nearly every VCE is used, and their waits even out across the platform (optimistic).
Before arriving passengers walked to the VCEs, their VCEs acted as one pooled queue,
which cleared sooner, e.g. on platform 3 with trains 2 minutes apart, at 8:29.
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
[`data/vces.csv`](./data/vces.csv),
finding them the same way as [the shapes](#platform-shapes), including curved stairs,
so the two match:

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
- The model gives an escalator the TCQSM's capacity for its tread width,
  so a few inches' error can move it across 2'8", between 34 and 72 pax/min.

Platforms 9 to 11 aren't on that plan,
so their VCEs are from NJT's PCIP Phase 1 existing plan
([Appendix A, sheet A-001](https://liamblank.com/wp-content/uploads/2026/07/penn-records-s-nj-transit-pcip-1-pcip1-final-report-c5015-01-262652-00-task-06-mem-final-report-draft-appendixa-drawings-copy.pdf#page=24),
July 2019), with `source` `pcip_phase_1`, where it draws their treads,
measured as for the [platform shapes](#platform-shapes),
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
  e.g. on platform 10.
- Each extended platform's new west end is where its extension's box on Figure 11 ends,
  185' to 317' west of its end in PCIP Phase 1's plan.

Each platform's west end today is from NJT's PCIP Phase 1 existing track plan
([Appendix A, sheet TK-003](https://liamblank.com/wp-content/uploads/2026/07/penn-records-s-nj-transit-pcip-1-pcip1-final-report-c5015-01-262652-00-task-06-mem-final-report-draft-appendixa-drawings-copy.pdf#page=8),
July 2019), since the PCIP Phase 2 and Master Plan drawings cut the platforms off at their west ends.
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

### Platform Shapes

The model doesn't use them yet, but the shapes of the platforms and of what's on them
are extracted from four vector plans of the station,
by [`shapes.py`](./src/platform_crowd_model/shapes.py) and a module for each plan:

- NJT's PCIP Phase 1 existing plan
  ([Appendix A, sheet A-001](https://liamblank.com/wp-content/uploads/2026/07/penn-records-s-nj-transit-pcip-1-pcip1-final-report-c5015-01-262652-00-task-06-mem-final-report-draft-appendixa-drawings-copy.pdf#page=24),
  July 2019),
  by `uv run platform-crowd-model data shapes-pcip-phase-1`
  ([`shapes_pcip_phase_1.py`](./src/platform_crowd_model/shapes_pcip_phase_1.py)).
  It has every platform, 1 to 11 and the diagonal platform, though it cuts 5 to 8 off to the west.
  It's marked not to scale, but it's drawn to one:
  fitting its platforms' east ends and centerlines to the frame
  gives scales along and across the platforms within 0.1% of each other.
- NJT's PCIP Phase 2 existing plan
  ([sheet A-001](https://liamblank.com/wp-content/uploads/2026/09/pcip-2-conceptual-design-preliminary-drawings.pdf#page=45),
  November 2020),
  by `uv run platform-crowd-model data shapes-pcip-phase-2`
  ([`shapes_pcip_phase_2.py`](./src/platform_crowd_model/shapes_pcip_phase_2.py)).
  It only has platforms 1 to 8.
- NJT's PCIP Phase 1 plan of Alternative 12
  ([Appendix A, sheet A-021](https://liamblank.com/wp-content/uploads/2026/07/penn-records-s-nj-transit-pcip-1-pcip1-final-report-c5015-01-262652-00-task-06-mem-final-report-draft-appendixa-drawings-copy.pdf#page=31),
  July 2019), with [Platform A](#platform-a),
  by `uv run platform-crowd-model data shapes-platform-a-pcip-phase-1`
  ([`shapes_platform_a_pcip_phase_1.py`](./src/platform_crowd_model/shapes_platform_a_pcip_phase_1.py)).
  It's the existing plan with Alternative 12 added, so it's read the same way.
  Platform A's VCEs are named by their labels, e.g. `AP7`,
  and each of AP7 to AP12's stair and escalator side by side is one footprint,
  since they're drawn as one run of treads.
- The Moynihan Station EA's lower concourse plan
  ([Figure 3-4](https://web.archive.org/web/2017id_/https://cdn.esd.ny.gov/subsidiaries_projects/msdc/Data/NEPA/03a%20Figure%203-3%20and%203-4.pdf#page=2),
  February 2010),
  by `uv run platform-crowd-model data shapes-moynihan-ea`
  ([`shapes_moynihan_ea.py`](./src/platform_crowd_model/shapes_moynihan_ea.py)),
  with the platforms' west ends, Moynihan Train Hall's escalators,
  and the stairs down from the West End Concourse and the baggage and egress corridor.
  It's a design, from before the Train Hall was built,
  so its VCEs are as [`vces_moynihan_ea.py`](./src/platform_crowd_model/vces_moynihan_ea.py)
  measures them, leaving out the escalator taken as not built.
  Its platforms' west ends agree with the PCIP Phase 1 track plan's to within about 2'
  on platforms 3 to 9,
  though platforms 4 and 8's last 40', past the baggage and egress corridor,
  are drawn as slivers too thin to extract.
  It doesn't show platforms 1 and 2, or 10 and 11 west of the West End Concourse.
  Pieces of platform west of the West End Concourse are named by the PCIP Phase 1 plan's platform
  they're mostly on, if any.

The PCIP plans are from before Moynihan Train Hall opened, so they don't have its escalators.
From each plan, they're:

- each platform's outline, with platforms 1 and 2 sharing one, as the plans draw them
- each VCE's footprint: each flight's treads and the balustrades beside them,
  joined by their landings, so a T-shaped stair's is a T,
  named as the nearest VCE of the same type in [`data/vces.csv`](./data/vces.csv), if one's within 15'
- curved stairs, e.g. the Central Concourse's down to platforms 5 and 7,
  as the band their treads sweep, found as runs of evenly spaced treads
  that aren't horizontal or vertical
- the columns, the small squares on the platforms
- the elevators, the boxes with an X across them
- the walls, e.g. of rooms and of the enclosures around VCEs, as lines,
  since most have gaps, e.g. for doors, or as areas where they close
- the concourses above the platforms, at the concourse level

What's under the platforms' labels is left out of the shapes, since the labels clip it.
The two plans' VCEs on platforms 1 to 8 agree to within a median of 2'10" along the platforms
and 10" across them.
The PCIP Phase 1 plan also shows the escalator about 220' along each of platforms 3 to 8
that the PCIP Phase 2 plan's labels cover.

They're combined into [`data/shapes.geojson`](./data/shapes.geojson)
and [`data/shapes.latlon.geojson`](./data/shapes.latlon.geojson)
by `uv run platform-crowd-model data shapes-combined`
([`shapes_combined.py`](./src/platform_crowd_model/shapes_combined.py)),
taking each part of the station from its best source, like `data/vces.csv`:

- Each VCE in `data/vces.csv` is from the plan its row is from,
  or else, e.g. for those from NJT's directory, the PCIP Phase 1 plan.
  8 of platforms 9 to 11's, which the PCIP Phase 1 plan doesn't draw treads for, have no footprint,
  and nor does platform 3's escalator under the PCIP Phase 2 plan's label.
- Everything else on the platforms is from the PCIP Phase 2 plan where it has the platform,
  or else the PCIP Phase 1 plan, or else, further west, the Moynihan Station EA's,
  and each platform's outline is pieced together from them the same way.
- The concourses are the PCIP Phase 1 plan's, and the Moynihan Station EA's beyond it.

The shapes are 2D, at the platform level unless they say otherwise,
since what matters on the platform is the space each VCE takes up there;
going up, a VCE's capacity is already its width.

They're indented GeoJSON, with each geometry's coordinates on one line, in two files for each plan:

- `data/shapes_<plan>.geojson`, e.g.
  [`data/shapes_pcip_phase_1.geojson`](./data/shapes_pcip_phase_1.geojson),
  in feet in the Master Plan's frame, extended to 2D:
  x east of its plans' west edge, as in `data/vces.csv`,
  and y north of platform 5's centerline on the PCIP Phase 2 plan,
  where east and north are along Manhattan's street grid.
  GeoJSON requires longitude and latitude, so this is strictly not valid GeoJSON,
  but it's what the code reads.
- `data/shapes_<plan>.latlon.geojson`, e.g.
  [`data/shapes_pcip_phase_1.latlon.geojson`](./data/shapes_pcip_phase_1.latlon.geojson),
  the same shapes in longitude and latitude, so GitHub can show them on a map.
  Each point is still longitude first, as GeoJSON requires.
  Every plan's are converted the same way, since they're in the same frame,
  registered by the PCIP Phase 2 plan's platforms to their outlines in OpenStreetMap:
  rotated 29.2° to their average direction, the street grid's,
  and offset to match their east ends on average,
  which agree to within about 4' across the platforms, but only about 37' along them,
  like OpenStreetMap's lengths.

## Results

Times are in m:ss.
Headways, time at capacity, and dwells are durations, and the rest are times after the first train arrives.

- **Platform:** the platform's number, or `A` for Platform A,
  followed by `T` as Penn Transformation would leave it, e.g. `3T`.
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
- **NFPA 130 travel distance:** how far the farthest occupant walks along the platform to their nearest VCE,
  with every VCE in service, marked ✗ if it's over NFPA 130's 100 m (328'1") limit.
- **Unused VCEs:** the VCEs no arriving passenger goes up,
  e.g. because a nearer one is always quicker, or blank if every VCE is used,
  not counting escalators only going down.

The ETA report quotes the max up rate, time at capacity, and taper time
for platform 3 with trains 2 minutes apart:
12.68 pax/s for 0:13, tapering at 5:55, with Penn Reconstruction,
and 12.04 pax/s for 0:32, tapering at 6:04, without it.

### Current Results

This table is generated by `uv run platform-crowd-model run --update-readme`:

<!-- results-table:start -->
| Platform | Headway | VCE width | NFPA 130 evacuation | NFPA 130 to concourse | Arrivals | Dwell | Taper time | Clear time | Boarded time | Time at capacity | Max pax on platform | Max density (pax/m²) | Max up rate (pax/s) | NFPA 130 travel distance | Unused VCEs |
|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|
| A | 0:00 | 60' | 6:42 ✗ | 7:03 ✗ | 0:00, 0:00, 0:50, 0:50 | 0:50, 0:50, 0:50, 0:50 | 6:01 | 6:15 | 1:40 | 6:09 | 5832 | 3.68 (E) | 17.33 | 80'6" ✓ |  |
| 1 | 0:00 | 35'4" | 6:41 ✗ | 7:02 ✗ | 0:00, 0:00, 4:38, 4:38 | 4:38, 4:38, 1:04, 1:04 | 8:45 | 9:00 | 5:42 | 8:30 | 3292 | 3.12 (D) | 9.32 | 193' ✓ |  |
| 1T | 0:00 | 47'4" | 7:29 ✗ | 7:50 ✗ | 0:00, 0:00, 0:56, 0:56 | 0:56, 0:56, 0:56, 0:56 | 6:51 | 7:06 | 1:52 | 6:58 | 5069 | 3.70 (E) | 12.72 | 118' ✓ |  |
| 2 | 0:00 | 35'7" | 6:42 ✗ | 7:03 ✗ | 0:00, 0:00, 3:59, 3:59 | 3:59, 3:59, 1:01, 1:01 | 7:39 | 7:53 | 5:00 | 7:38 | 3380 | 3.41 (D) | 10.42 | 192' ✓ |  |
| 2T | 0:00 | 47'7" | 7:12 ✗ | 7:33 ✗ | 0:00, 0:00, 1:00, 1:00 | 1:00, 1:00, 1:00, 1:00 | 6:21 | 6:35 | 2:00 | 6:07 | 4955 | 3.72 (E) | 13.82 | 143' ✓ |  |
| 3 | 0:00 | 49'1" | 6:51 ✗ | 7:12 ✗ | 0:00, 0:00, 1:05, 1:05 | 1:05, 1:05, 1:05, 1:05 | 6:10 | 6:25 | 2:10 | 6:19 | 4812 | 3.79 (E) | 14.08 | 101'6" ✓ |  |
| 3T | 0:00 | 64'5" | 4:53 ✗ | 5:14 ✓ | 0:00, 0:00, 1:03, 1:03 | 1:03, 1:03, 1:03, 1:03 | 4:36 | 4:50 | 2:06 | 4:44 | 4394 | 2.50 (D) | 18.68 | 258' ✓ |  |
| 4 | 0:00 | 46' | 9:02 ✗ | 9:23 ✗ | 0:00, 0:00, 0:57, 0:57 | 0:57, 0:57, 0:57, 0:57 | 7:48 | 8:01 | 1:54 | 7:53 | 6057 | 3.87 (E) | 13.49 | 184' ✓ |  |
| 4T | 0:00 | 58' | 6:48 ✗ | 7:09 ✗ | 0:00, 0:00, 1:00, 1:00 | 1:00, 1:00, 1:00, 1:00 | 6:11 | 6:25 | 2:01 | 6:16 | 5699 | 3.53 (D) | 16.89 | 184' ✓ |  |
| 5 | 0:00 | 57'3" | 6:54 ✗ | 7:15 ✗ | 0:00, 0:00, 1:05, 1:05 | 1:05, 1:05, 1:05, 1:05 | 6:04 | 6:18 | 2:10 | 5:12 | 5684 | 2.95 (D) | 17.57 | 230' ✓ |  |
| 5T | 0:00 | 72'7" | 5:07 ✗ | 5:28 ✓ | 0:00, 0:00, 1:04, 1:04 | 1:04, 1:04, 1:04, 1:04 | 4:44 | 4:58 | 2:08 | 4:03 | 5222 | 2.60 (D) | 22.17 | 230' ✓ |  |
| 6 | 0:00 | 43'10" | 9:23 ✗ | 9:44 ✗ | 0:00, 0:00, 1:20, 1:20 | 1:20, 1:20, 1:05, 1:05 | 8:04 | 8:17 | 2:25 | 6:26 | 5997 | 3.36 (D) | 13.60 | 200' ✓ |  |
| 6T | 0:00 | 55'10" | 7:14 ✗ | 7:35 ✗ | 0:00, 0:00, 0:54, 0:54 | 0:54, 0:54, 0:54, 0:54 | 6:11 | 6:25 | 1:48 | 5:35 | 5835 | 3.17 (D) | 17.00 | 200' ✓ |  |
| 7 | 0:00 | 50'3" | 8:06 ✗ | 8:27 ✗ | 0:00, 0:00, 1:02, 1:02 | 1:02, 1:02, 1:02, 1:02 | 7:11 | 7:25 | 2:04 | 6:42 | 5917 | 3.10 (D) | 14.69 | 248' ✓ |  |
| 7T | 0:00 | 59'7" | 6:38 ✗ | 6:59 ✗ | 0:00, 0:00, 0:59, 0:59 | 0:59, 0:59, 0:59, 0:59 | 5:58 | 6:12 | 1:59 | 5:16 | 5701 | 2.87 (D) | 17.59 | 248' ✓ |  |
| 8 | 0:00 | 49'3" | 8:15 ✗ | 8:36 ✗ | 0:00, 0:00, 1:02, 1:02 | 1:02, 1:02, 1:02, 1:02 | 7:15 | 7:29 | 2:04 | 7:25 | 5896 | 4.11 (E) | 14.45 | 202' ✓ |  |
| 8T | 0:00 | 58'7" | 6:44 ✗ | 7:05 ✗ | 0:00, 0:00, 0:58, 0:58 | 0:58, 0:58, 0:58, 0:58 | 6:00 | 6:14 | 1:56 | 6:09 | 5685 | 3.80 (E) | 17.35 | 125' ✓ |  |
| 9 | 0:00 | 50'4" | 4:21 ✗ | 4:42 ✓ | 0:00, 0:58, 1:56, 2:54 | 0:58, 0:58, 0:58, 0:58 | 6:22 | 6:37 | 3:52 | 6:17 | 2987 | 2.24 (D) | 13.71 | 107'6" ✓ |  |
| 9T | 0:00 | 62'4" | 2:59 ✓ | 3:20 ✓ | 0:00, 0:59, 1:58, 2:57 | 0:59, 0:59, 0:59, 0:59 | 5:06 | 5:21 | 3:56 | 4:56 | 2444 | 1.80 (D) | 17.11 | 106' ✓ |  |
| 10 | 0:00 | 89'11" | 4:05 ✗ | 4:26 ✓ | 0:00, 0:00, 0:55, 0:55 | 0:55, 0:55, 0:55, 0:55 | 4:09 | 4:24 | 1:50 | 4:02 | 5118 | 2.10 (D) | 24.81 | 122' ✓ |  |
| 10T | 0:00 | 95'11" | 3:48 ✓ | 4:09 ✓ | 0:00, 0:00, 0:53, 0:53 | 0:53, 0:53, 0:53, 0:53 | 3:52 | 4:08 | 1:46 | 3:36 | 5041 | 1.81 (D) | 26.51 | 122' ✓ |  |
| 11 | 0:00 | 44'11" | 6:38 ✗ | 6:59 ✗ | 0:00, 0:00, 4:24, 4:24 | 4:24, 4:24, 0:48, 0:48 | 8:20 | 8:35 | 5:12 | 7:58 | 4248 | 3.82 (E) | 13.01 | 127'6" ✓ |  |
| 11T | 0:00 | 56'11" | 7:02 ✗ | 7:23 ✗ | 0:00, 0:00, 0:55, 0:55 | 0:55, 0:55, 0:55, 0:55 | 6:21 | 6:36 | 1:50 | 6:29 | 5825 | 5.13 (E) | 16.41 | 90' ✓ |  |
| A | 2:00 | 60' | 2:08 ✓ | 2:29 ✓ | 0:00, 2:00, 4:00, 6:00 | 0:50, 1:05, 1:05, 1:05 | 7:21 | 7:36 | 7:05 | 5:48 | 1530 | 0.97 (B) | 17.33 | 80'6" ✓ |  |
| 1 | 2:00 | 35'4" | 3:14 ✓ | 3:35 ✓ | 0:00, 2:00, 4:00, 8:03 | 1:04, 6:03, 3:52, 1:05 | 10:01 | 10:15 | 9:08 | 8:15 | 1361 | 1.29 (C) | 9.32 | 193' ✓ |  |
| 1T | 2:00 | 47'4" | 2:38 ✓ | 2:59 ✓ | 0:00, 2:00, 4:00, 7:24 | 0:56, 5:24, 3:10, 1:05 | 9:02 | 9:22 | 8:29 | 4:20 | 1482 | 1.08 (C) | 12.72 | 118' ✓ |  |
| 2 | 2:00 | 35'7" | 2:57 ✓ | 3:18 ✓ | 0:00, 2:00, 4:00, 7:15 | 0:59, 5:15, 3:14, 1:05 | 9:00 | 9:13 | 8:20 | 6:40 | 1325 | 1.34 (C) | 10.42 | 192' ✓ |  |
| 2T | 2:00 | 47'7" | 2:21 ✓ | 2:42 ✓ | 0:00, 2:00, 4:00, 6:00 | 0:52, 1:05, 1:05, 1:05 | 7:27 | 7:42 | 7:05 | 5:20 | 1396 | 1.05 (B) | 13.82 | 143' ✓ |  |
| 3 | 2:00 | 49'1" | 2:17 ✓ | 2:38 ✓ | 0:00, 2:00, 4:00, 6:00 | 0:54, 1:05, 1:05, 1:05 | 7:23 | 7:38 | 7:05 | 5:44 | 1343 | 1.06 (B) | 14.08 | 101'6" ✓ |  |
| 3T | 2:00 | 64'5" | 1:43 ✓ | 2:26 ✓ | 0:00, 2:00, 4:00, 6:00 | 0:56, 1:01, 1:01, 1:01 | 7:01 | 7:15 | 7:01 | 2:36 | 1231 | 0.70 (A) | 18.68 | 258' ✓ |  |
| 4 | 2:00 | 46' | 2:49 ✓ | 3:10 ✓ | 0:00, 2:00, 4:00, 6:58 | 0:55, 4:58, 2:58, 1:05 | 8:46 | 9:00 | 8:03 | 7:08 | 1626 | 1.04 (B) | 13.49 | 184' ✓ |  |
| 4T | 2:00 | 58' | 2:12 ✓ | 2:33 ✓ | 0:00, 2:00, 4:00, 6:00 | 0:52, 1:05, 1:05, 1:05 | 7:24 | 7:40 | 7:05 | 3:56 | 1536 | 0.95 (B) | 16.89 | 184' ✓ |  |
| 5 | 2:00 | 57'3" | 2:14 ✓ | 2:35 ✓ | 0:00, 2:00, 4:00, 6:00 | 0:55, 1:05, 1:05, 1:05 | 7:24 | 7:38 | 7:05 | 1:56 | 1567 | 0.81 (A) | 17.57 | 230' ✓ |  |
| 5T | 2:00 | 72'7" | 1:45 ✓ | 2:13 ✓ | 0:00, 2:00, 4:00, 6:00 | 0:58, 1:05, 1:05, 1:05 | 7:05 | 7:19 | 7:05 | 0:00 | 1456 | 0.73 (A) | 21.20 | 230' ✓ | P5-S1 |
| 6 | 2:00 | 43'10" | 2:58 ✓ | 3:19 ✓ | 0:00, 2:00, 4:00, 7:12 | 0:55, 5:12, 3:07, 0:59 | 9:04 | 9:24 | 8:11 | 5:12 | 1678 | 0.94 (B) | 13.60 | 200' ✓ |  |
| 6T | 2:00 | 55'10" | 2:18 ✓ | 2:39 ✓ | 0:00, 2:00, 4:00, 6:00 | 0:51, 1:05, 1:05, 1:05 | 7:26 | 7:42 | 7:05 | 3:44 | 1584 | 0.86 (B) | 17.00 | 200' ✓ |  |
| 7 | 2:00 | 50'3" | 3:04 ✓ | 3:24 ✓ | 0:00, 2:00, 4:00, 6:27 | 0:53, 4:27, 2:30, 1:00 | 8:06 | 8:20 | 7:27 | 6:40 | 1605 | 0.84 (B) | 14.69 | 248' ✓ |  |
| 7T | 2:00 | 59'7" | 2:09 ✓ | 2:30 ✓ | 0:00, 2:00, 4:00, 6:00 | 0:50, 1:05, 1:05, 1:05 | 7:22 | 7:38 | 7:05 | 3:24 | 1554 | 0.78 (A) | 17.59 | 248' ✓ |  |
| 8 | 2:00 | 49'3" | 2:39 ✓ | 3:00 ✓ | 0:00, 2:00, 4:00, 6:39 | 0:54, 4:39, 2:39, 0:59 | 8:21 | 8:34 | 7:38 | 5:48 | 1624 | 1.13 (C) | 14.45 | 202' ✓ |  |
| 8T | 2:00 | 58'7" | 2:11 ✓ | 2:32 ✓ | 0:00, 2:00, 4:00, 6:00 | 0:51, 1:05, 1:05, 1:05 | 7:21 | 7:37 | 7:05 | 5:28 | 1505 | 1.01 (B) | 17.35 | 125' ✓ |  |
| 9 | 2:00 | 50'4" | 2:14 ✓ | 2:35 ✓ | 0:00, 2:00, 4:00, 6:00 | 0:57, 1:05, 1:05, 1:05 | 7:28 | 7:44 | 7:05 | 5:16 | 1426 | 1.07 (B) | 13.71 | 107'6" ✓ |  |
| 9T | 2:00 | 62'4" | 1:47 ✓ | 2:08 ✓ | 0:00, 2:00, 4:00, 6:00 | 0:59, 1:05, 1:05, 1:05 | 7:11 | 7:35 | 7:05 | 1:24 | 1391 | 1.02 (B) | 17.11 | 106' ✓ |  |
| 10 | 2:00 | 89'11" | 1:25 ✓ | 1:46 ✓ | 0:00, 2:00, 4:00, 6:00 | 0:55, 1:05, 1:05, 1:05 | 6:58 | 7:38 | 7:05 | 0:00 | 1395 | 0.57 (A) | 24.61 | 122' ✓ |  |
| 10T | 2:00 | 95'11" | 1:19 ✓ | 1:40 ✓ | 0:00, 2:00, 4:00, 6:00 | 0:53, 1:00, 1:00, 1:00 | 6:56 | 7:38 | 7:00 | 0:00 | 1413 | 0.51 (A) | 26.31 | 122' ✓ |  |
| 11 | 2:00 | 44'11" | 2:53 ✓ | 3:14 ✓ | 0:00, 2:00, 4:00, 7:57 | 0:48, 5:57, 3:57, 0:48 | 9:49 | 10:04 | 8:45 | 7:13 | 1616 | 1.45 (C) | 13.01 | 127'6" ✓ |  |
| 11T | 2:00 | 56'11" | 2:40 ✓ | 3:01 ✓ | 0:00, 2:00, 4:00, 6:17 | 0:55, 4:17, 2:18, 0:55 | 7:43 | 7:59 | 7:12 | 5:24 | 1502 | 1.32 (C) | 16.41 | 90' ✓ |  |
| A | 5:00 | 60' | 2:07 ✓ | 2:28 ✓ | 0:00, 5:00, 10:00, 15:00 | 0:50, 0:50, 0:50, 0:50 | 16:21 | 16:36 | 15:50 | 5:48 | 1457 | 0.92 (B) | 17.33 | 80'6" ✓ |  |
| 1 | 5:00 | 35'4" | 3:02 ✓ | 3:23 ✓ | 0:00, 5:00, 10:00, 15:00 | 0:50, 0:50, 0:50, 0:50 | 16:58 | 17:12 | 15:50 | 7:44 | 1312 | 1.24 (C) | 9.32 | 193' ✓ |  |
| 1T | 5:00 | 47'4" | 2:23 ✓ | 2:44 ✓ | 0:00, 5:00, 10:00, 15:00 | 0:48, 0:48, 0:48, 0:48 | 16:34 | 16:50 | 15:48 | 6:12 | 1345 | 0.98 (B) | 12.72 | 118' ✓ |  |
| 2 | 5:00 | 35'7" | 2:56 ✓ | 3:17 ✓ | 0:00, 5:00, 10:00, 15:00 | 0:57, 0:57, 0:57, 0:57 | 16:44 | 16:58 | 15:57 | 7:00 | 1271 | 1.28 (C) | 10.42 | 192' ✓ |  |
| 2T | 5:00 | 47'7" | 2:20 ✓ | 2:41 ✓ | 0:00, 5:00, 10:00, 15:00 | 0:52, 0:52, 0:52, 0:52 | 16:27 | 16:42 | 15:52 | 5:20 | 1315 | 0.99 (B) | 13.82 | 143' ✓ |  |
| 3 | 5:00 | 49'1" | 2:16 ✓ | 2:37 ✓ | 0:00, 5:00, 10:00, 15:00 | 0:54, 0:54, 0:54, 0:54 | 16:23 | 16:38 | 15:54 | 5:44 | 1286 | 1.01 (B) | 14.08 | 101'6" ✓ |  |
| 3T | 5:00 | 64'5" | 1:42 ✓ | 2:26 ✓ | 0:00, 5:00, 10:00, 15:00 | 0:51, 0:51, 0:51, 0:51 | 16:00 | 16:15 | 15:51 | 4:04 | 1135 | 0.65 (A) | 18.68 | 258' ✓ |  |
| 4 | 5:00 | 46' | 2:48 ✓ | 3:09 ✓ | 0:00, 5:00, 10:00, 15:00 | 0:55, 0:55, 0:55, 0:55 | 16:48 | 17:02 | 15:55 | 7:08 | 1573 | 1.01 (B) | 13.49 | 184' ✓ |  |
| 4T | 5:00 | 58' | 2:12 ✓ | 2:33 ✓ | 0:00, 5:00, 10:00, 15:00 | 0:52, 0:52, 0:52, 0:52 | 16:24 | 16:40 | 15:52 | 3:56 | 1470 | 0.91 (B) | 16.89 | 184' ✓ |  |
| 5 | 5:00 | 57'3" | 2:13 ✓ | 2:34 ✓ | 0:00, 5:00, 10:00, 15:00 | 0:55, 0:55, 0:55, 0:55 | 16:24 | 16:38 | 15:55 | 1:56 | 1476 | 0.77 (A) | 17.57 | 230' ✓ |  |
| 5T | 5:00 | 72'7" | 1:44 ✓ | 2:13 ✓ | 0:00, 5:00, 10:00, 15:00 | 0:58, 0:58, 0:58, 0:58 | 16:05 | 16:19 | 15:58 | 0:00 | 1338 | 0.67 (A) | 21.20 | 230' ✓ | P5-S1 |
| 6 | 5:00 | 43'10" | 2:57 ✓ | 3:18 ✓ | 0:00, 5:00, 10:00, 15:00 | 0:55, 0:55, 0:55, 0:55 | 16:51 | 17:06 | 15:55 | 5:24 | 1608 | 0.90 (B) | 13.60 | 200' ✓ |  |
| 6T | 5:00 | 55'10" | 2:17 ✓ | 2:38 ✓ | 0:00, 5:00, 10:00, 15:00 | 0:51, 0:51, 0:51, 0:51 | 16:26 | 16:42 | 15:51 | 3:44 | 1494 | 0.81 (A) | 17.00 | 200' ✓ |  |
| 7 | 5:00 | 50'3" | 2:33 ✓ | 2:54 ✓ | 0:00, 5:00, 10:00, 15:00 | 0:53, 0:53, 0:53, 0:53 | 16:40 | 16:55 | 15:53 | 5:32 | 1560 | 0.82 (A) | 14.69 | 248' ✓ |  |
| 7T | 5:00 | 59'7" | 2:08 ✓ | 2:29 ✓ | 0:00, 5:00, 10:00, 15:00 | 0:50, 0:50, 0:50, 0:50 | 16:22 | 16:38 | 15:50 | 3:24 | 1465 | 0.74 (A) | 17.59 | 248' ✓ |  |
| 8 | 5:00 | 49'3" | 2:37 ✓ | 2:58 ✓ | 0:00, 5:00, 10:00, 15:00 | 0:54, 0:54, 0:54, 0:54 | 16:40 | 16:55 | 15:54 | 6:48 | 1546 | 1.08 (C) | 14.45 | 202' ✓ |  |
| 8T | 5:00 | 58'7" | 2:10 ✓ | 2:31 ✓ | 0:00, 5:00, 10:00, 15:00 | 0:51, 0:51, 0:51, 0:51 | 16:21 | 16:37 | 15:51 | 5:28 | 1446 | 0.97 (B) | 17.35 | 125' ✓ |  |
| 9 | 5:00 | 50'4" | 2:13 ✓ | 2:34 ✓ | 0:00, 5:00, 10:00, 15:00 | 0:54, 0:54, 0:54, 0:54 | 16:27 | 16:43 | 15:54 | 5:24 | 1327 | 0.99 (B) | 13.71 | 107'6" ✓ |  |
| 9T | 5:00 | 62'4" | 1:46 ✓ | 2:07 ✓ | 0:00, 5:00, 10:00, 15:00 | 0:59, 0:59, 0:59, 0:59 | 16:13 | 16:39 | 15:59 | 0:12 | 1263 | 0.93 (B) | 17.11 | 106' ✓ |  |
| 10 | 5:00 | 89'11" | 1:23 ✓ | 1:44 ✓ | 0:00, 5:00, 10:00, 15:00 | 0:55, 0:55, 0:55, 0:55 | 15:58 | 16:38 | 15:55 | 0:00 | 1280 | 0.53 (A) | 24.61 | 122' ✓ |  |
| 10T | 5:00 | 95'11" | 1:18 ✓ | 1:39 ✓ | 0:00, 5:00, 10:00, 15:00 | 0:53, 0:53, 0:53, 0:53 | 15:56 | 16:38 | 15:53 | 0:00 | 1259 | 0.45 (A) | 26.31 | 122' ✓ |  |
| 11 | 5:00 | 44'11" | 2:52 ✓ | 3:13 ✓ | 0:00, 5:00, 10:00, 15:00 | 0:48, 0:48, 0:48, 0:48 | 16:52 | 17:07 | 15:48 | 6:52 | 1600 | 1.44 (C) | 13.01 | 127'6" ✓ |  |
| 11T | 5:00 | 56'11" | 2:14 ✓ | 2:35 ✓ | 0:00, 5:00, 10:00, 15:00 | 0:55, 0:55, 0:55, 0:55 | 16:26 | 16:42 | 15:55 | 5:24 | 1483 | 1.31 (C) | 16.41 | 90' ✓ |  |
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
- **Modeled every platform as Penn Transformation would leave it, too,** e.g. `3T`,
  with the FRA's new VCEs, platforms 1 to 3's extensions, and decluttering,
  instead of only platform 3 with Penn Reconstruction, with the ETA report's total width for it.
  With trains 2 minutes apart, every platform clears sooner,
  e.g. platform 1 at 8:55 instead of 11:19, with 43'6" of VCEs instead of 31'6",
  and platform 3 at 7:18 instead of 8:29, with 61'1" instead of 45'9".
  With two trains arriving at once, platforms 9T and 10T pass NFPA 130's 4 minutes,
  and 2T, 3T, 5T, 9T, and 10T reach the concourse within 6 minutes.
  But 4T, 6T, 8T, and 11T take longer than they do today:
  their departing passengers board sooner, so the last two trains arrive at 0:43,
  and all four trains' passengers are on the platform at once.
- **Modeled PCIP Phase 1's Platform A,** a new platform south of platform 1,
  with 8 stairs and 6 escalators, 56'8" wide in all without its widest escalator.
  With trains 2 minutes apart, it clears at 7:57,
  and with two trains arriving at once, it fails both NFPA 130 checks, taking 6:53 and 7:14.
  The other platforms are unchanged.
- **Limited escalators to half of NFPA 130's exit capacity,** as NFPA 130 requires (5.3.5.6).
  No effect on results, since no platform's escalators provide more than 44% of it.
- **Had NFPA 130's farthest occupant walk to their nearest VCE,**
  from either end of the platform or from halfway between two VCEs,
  instead of the farthest NFPA 130 allows, 325',
  with the widest escalator out of service.
  The longest walk is 258', on platform 3 with Penn Transformation,
  and most are about 100' to 250'.
  With trains 2 or 5 minutes apart, platforms reach the concourse up to 1:20 sooner,
  e.g. platform 10 in 1:45 instead of 2:59.
  With two trains arriving at once, the flow time is longer than any walk,
  so those are unchanged.
- **Used NFPA 130's metric values, its requirements, instead of their approximate US equivalents** (1.5.2),
  e.g. 37.8 m/min instead of 124 fpm.
  A few NFPA 130 times to the concourse are 1 s longer.
- **Added NFPA 130's travel distance to the results table,**
  marked ✗ if it's over its 100 m (328'1") limit (5.3.3.5).
  Every platform is within it.
- **Gave each VCE its own queue,** which discharges at its own capacity,
  with departing passengers coming down each VCE with whatever capacity its own upward flow leaves,
  and the stair LOS grading the most crowded VCE.
  Arriving passengers still spread across the VCEs in proportion to their widths,
  so every queue empties at the same time, and no results change,
  except that platform 6's max up rate, exactly 11.475 pax/s, now rounds to 11.48 instead of 11.47.
- **Added the VCEs no arriving passenger goes up to the results table.**
  Arriving passengers spread across every VCE, so there are none yet.
- **Sent each door's arriving passengers to the quickest VCE,**
  i.e. the one with the least walking time plus waiting time for everyone queued or walking there,
  walking at the TCQSM's design walking speed, 250 ft/min,
  instead of spreading them across the VCEs in proportion to their widths.
  Trains stop with their east ends at the platform's, with their doors spread evenly along them.
  VCEs near the ends of the platform or far from the busiest doors now run dry
  while others still have a queue,
  so every platform clears later, e.g. with trains 2 minutes apart,
  platform 3 at 9:03 instead of 8:29, and platform 6 at 12:24 instead of 11:09,
  and more passengers crowd the platform at once,
  e.g. 1999 instead of 1716 on platform 6 and 1879 instead of 1494 on 7T,
  so NFPA 130's times are up to 33 s longer, though no platform passes or fails differently.
  With trains 5 minutes apart, 7T's `P7-S1`, about 215' west of its trains, is never the quickest, so nobody goes up it.
- **Sent departing passengers to the nearest car that isn't close to full,**
  walking there from their VCE at 250 ft/min,
  with each car boarding its own waiting passengers through its own doors,
  instead of spreading evenly across the doors as soon as they come down.
  Most trains' dwells are longer, e.g. 0:56 instead of 0:44 on platform 3 with trains 5 minutes apart,
  and with trains 2 minutes apart, the busiest cars' doors hold up their trains,
  so platform 3 finishes boarding at 8:14 instead of 7:29,
  and platform 6 at 11:20 instead of 9:28, clearing at 13:46 instead of 12:24.
  With trains arriving at once, platform 9's trains, all on its one track, arrive later,
  so it peaks at 2945 passengers instead of 3624, evacuating in 4:17 instead of 5:08.
  Elsewhere, longer dwells keep more passengers aboard,
  so NFPA 130's times are up to 44 s longer, e.g. platform 1's with trains 2 minutes apart, 3:47 instead of 3:03,
  though no platform passes or fails NFPA 130's checks differently.
- **Counted a VCE's queue as draining while arriving passengers walk to it,**
  so the quickest VCE is the one they could go up soonest,
  the later of their walk there and when everyone queued at or walking to it has gone up,
  instead of their walk plus that whole wait.
  Farther VCEs no longer look slower than they are, so passengers spread out more,
  and every platform clears sooner, by 9 s to 2:00,
  e.g. with trains 2 minutes apart, platform 3 at 8:36 instead of 9:28,
  and platform 6 at 12:54 instead of 13:46, peaking at 1808 passengers instead of 2015.
  NFPA 130's times change by up to 33 s, though no platform passes or fails differently,
  and 7T's `P7-S1` is now used with trains 5 minutes apart.
- **Ran escalators one way, at the TCQSM's escalator capacities,** instead of treating them as stairs,
  with the widest escalator left out:
  a platform's only escalator goes up, and with more,
  the easternmost goes up, the westernmost goes down,
  and the rest go up until the platform is nearly fully alighted.
  Escalators carry 34 or 72 pax/min by their tread width,
  and the VCE width counts every VCE, e.g. platform 3's is 49'1" instead of 45'9".
  With trains 2 minutes apart, most platforms clear about 50 s sooner,
  e.g. platform 6 at 10:28 instead of 12:54.
  With trains arriving at once, the escalator going down lets departing passengers board sooner,
  so the last two trains arrive while the platform is still crowded,
  e.g. on platform 6 at 0:59 instead of 5:00,
  and it peaks at 6082 passengers instead of 4312,
  so it takes 9:33 instead of 6:46 to evacuate under NFPA 130,
  and platforms 2T and 3 no longer reach the concourse within 6 minutes.
  With trains 2 or 5 minutes apart, 5T's `P5-S1`, at its west end, is never the quickest VCE going up,
  so nobody goes up it.
- **Stopped trains where the platform could be evacuated soonest,**
  i.e. with the fewest passengers on it at once,
  of the positions where the longest dwell is within 1:00 of the shortest it can be,
  instead of with their east ends at the platform's.
  NFPA 130's evacuation times are up to 35 s shorter,
  e.g. platform 6's with trains 2 minutes apart, 2:58 instead of 3:33,
  though no platform passes or fails differently.
  Clear times move by up to about a minute either way,
  e.g. with trains 2 minutes apart, platform 7's trains stop 240' further west,
  and it clears at 8:38 instead of 9:43,
  but 1T's stop 175' further west, crowding more passengers onto the platform,
  and it clears at 9:22 instead of 8:18.
  5T's `P5-S1` is now used with trains 2 minutes apart, but with trains 5 minutes apart,
  some trains stop far enough west to leave VCEs east of them that are never the quickest unused:
  3T's `P3-T3`, `P3-S9`, and `P3-S10`, and platform 5's `P5-S13`.
- **Counted stopping positions whose most occupants at once differ only by floating-point rounding as tied,**
  so the dwells choose between them instead of that rounding,
  e.g. from the last digits of `sum()`, which PyPy rounds differently from CPython.
  It moves 22 of 69 scenarios' stopping positions, all with trains 2 or 5 minutes apart,
  but no NFPA 130 time changes.
  Dwells are as short or shorter, e.g. 3T's with trains 5 minutes apart, 0:51 instead of 1:05,
  and it clears at 16:00 instead of 16:30.
  Every VCE is now used except 5T's `P5-S1`, with trains 2 or 5 minutes apart.
- **Counted stopping positions within 1 passenger of the fewest occupants at once as tied,**
  instead of only those within floating-point rounding,
  since a passenger barely changes NFPA 130's times, but dwells can differ by seconds.
  It moves 6 scenarios' stopping positions, but no NFPA 130 time changes.
  Platform 7's longest dwell with trains 2 minutes apart is 4:27 instead of 4:42,
  and it clears at 8:20 instead of 8:38.
- **Saved each VCE's queue and upward flow each second with `--charts`,**
  in `*_vces.csv`, and a list of the VCEs in `*_vce_list.csv`,
  with each one's type, role, width, position, and when its queue last empties.
  No results change.
- **Extracted the shapes of the platforms and of what's on them** from four plans
  into `data/shapes_<plan>.geojson`, in feet, and `data/shapes_<plan>.latlon.geojson`,
  with `uv run platform-crowd-model data shapes-*` for each plan
  (see [Platform Shapes](#platform-shapes)).
  The model doesn't read them yet. No results change.
- **Combined the plans' shapes into `data/shapes.geojson` and `data/shapes.latlon.geojson`,**
  taking each part of the station from its best source, like `data/vces.csv`,
  with `uv run platform-crowd-model data shapes-combined`,
  and tested that their VCEs match `data/vces.csv`'s.
  No results change.
- **Made a field survey sheet for measuring every VCE in person,**
  [`data/vces_field_survey.csv`](./data/vces_field_survey.csv),
  with `uv run platform-crowd-model data vces-field-survey`
  (see [Field Survey](#field-survey)).
  No results change.
