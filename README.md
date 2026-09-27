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
   17 pax/min per foot of VCE width, as long as anyone is queued.
   The stair queues hold 20 ft of queue in front of the total VCE width at 5 sqft/pax;
   once the remaining arrived passengers fit in them, they taper off.
   The few seconds of walking from the doors to the stairs are ignored.
3. **Coming downstairs.** Departing passengers come down with whatever stair capacity
   the upward flow leaves, up to the LOS C/D boundary, 10 pax/min/ft,
   and also limited by Fruin's descending stair equation, P = (128M − 206)/M² pax/min/ft,
   with M the space per departing passenger
   in a 5,000 sqft concourse,
   split between the two trains in proportion to how many are still upstairs.
4. **Boarding.** Departing passengers on the platform board a train that has arrived
   with whatever door capacity isn't being used for alighting.

Space per passenger is the usable platform area (75% of the platform's area)
divided by everyone on the platform,
graded with Fruin's walkway LOS (A > 35, B > 25, C > 15, D > 10, E > 5 sqft/pax, else F).

Every scenario has 200 departing passengers per train already on the platform
and another 200 per train upstairs.

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
| platform3 | 120 s | 42.5 ft | 12.04 | 269 s | 254 s | 269 s | never | 1522 | 8.0 (E) |
| platform3 | 300 s | 42.5 ft | 12.04 | 268 s | 420 s | 435 s | never | 1522 | 8.0 (E) |
| platform3_recon | 120 s | 44.75 ft | 12.68 | 255 s | 241 s | 256 s | never | 1496 | 8.1 (E) |
| platform3_recon | 300 s | 44.75 ft | 12.68 | 254 s | 413 s | 428 s | never | 1496 | 8.1 (E) |
| platform6 | 0 s | 48.168 ft | 13.65 | 237 s | 223 s | 238 s | never | 3058 | 4.0 (F) |
| platform10 | 120 s | 70.58 ft | 20.00 | 162 s | 186 s | 201 s | 548 s | 1206 | 28.7 (B) |
| platform11 | 120 s | 43.58 ft | 12.35 | 262 s | 248 s | 263 s | never | 1510 | 9.8 (E) |

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
