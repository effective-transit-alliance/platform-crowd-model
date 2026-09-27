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
   Stair flow is capped at LOS E capacity, 17 pax/min per foot of VCE width.
   The stair queues hold 20 ft of queue in front of the total VCE width at 5 sqft/pax.
   While more arrived passengers are waiting than fit in those queues,
   they flow up at no less than the LOS C/D boundary, 10 pax/min/ft.
   Otherwise, the flow comes from Fruin's stair equation, P = (111M − 162)/M²,
   where M is the platform space per arrived passenger,
   so it tapers off as the platform empties.
3. **Coming downstairs.** Departing passengers come down with whatever stair capacity
   the upward flow leaves, up to 12 pax/min/ft,
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
- **Max pax on platform** and **Min space/pax:** the most crowded moment on the platform, with its LOS.

The ETA report quotes the first three for platform 3 with trains 2 minutes apart:
12.68 pax/s for 13 s, tapering at 355 s, with Penn Reconstruction,
and 12.04 pax/s for 32 s, tapering at 364 s, without it.

### Current Results

| Platform | Headway | VCE width | Max up rate (pax/s) | Time at capacity | Taper time | Clear time | Max pax on platform | Min space/pax (sqft) |
|---|---|---|---|---|---|---|---|---|
| platform3 | 120 s | 42.5 ft | 12.04 | 32 s | 361 s | never | 2102 | 5.8 (E) |
| platform3 | 300 s | 42.5 ft | 10.53 | 0 s | 497 s | never | 1749 | 6.9 (E) |
| platform3_recon | 120 s | 44.75 ft | 12.68 | 15 s | 353 s | never | 2082 | 5.8 (E) |
| platform3_recon | 300 s | 44.75 ft | 10.48 | 0 s | 489 s | never | 1742 | 7.0 (E) |
| platform6 | 0 s | 48.168 ft | 13.65 | 74 s | 293 s | never | 3200 | 3.9 (F) |
| platform10 | 120 s | 70.58 ft | 11.76 | 0 s | 258 s | never | 1664 | 20.8 (C) |
| platform11 | 120 s | 43.58 ft | 11.99 | 0 s | 380 s | never | 2194 | 6.8 (E) |

### History

As bugs are fixed, the results table above is updated,
and each fix's effect is summarized here.

- **Original model** (as used in the ETA report):
  reproduces the report's platform 3 numbers to within a few seconds.
