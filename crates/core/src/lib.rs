//! The model's core: each second of `platform_crowd_model.model.simulate`,
//! and the search for where trains should stop.
//!
//! Python turns a scenario's `Params` into a `CoreScenario` of plain numbers.
//! From it, this lays out each stopping position's doors and cars,
//! simulates them, and returns a `CoreResult`,
//! which Python turns back into a `Summary` and a `TimeSeries`.

// Index loops keep each train's, car's, and VCE's parallel arrays in step.
#![allow(clippy::needless_range_loop)]

use pyo3::prelude::*;
use rayon::prelude::*;

const ROUNDING_TOLERANCE: f64 = 1e-9;

/// `total - pax`, or 0 if that's only left over from rounding.
fn subtract(total: f64, pax: f64) -> f64 {
    let remaining = total - pax;
    if remaining < ROUNDING_TOLERANCE {
        0.0
    } else {
        remaining
    }
}

/// What a VCE does, like Python's `Role`.
#[derive(Clone, Copy, PartialEq, Eq)]
enum Role {
    Stair,
    Up,
    Down,
    Reversible,
}

/// Which way a VCE runs now.
#[derive(Clone, Copy, PartialEq, Eq)]
enum Direction {
    Both,
    Up,
    Down,
}

impl<'a, 'py> FromPyObject<'a, 'py> for Role {
    type Error = PyErr;

    fn extract(obj: Borrowed<'a, 'py, PyAny>) -> PyResult<Self> {
        match obj.extract::<&str>()? {
            "stair" => Ok(Role::Stair),
            "up" => Ok(Role::Up),
            "down" => Ok(Role::Down),
            "reversible" => Ok(Role::Reversible),
            role => Err(pyo3::exceptions::PyValueError::new_err(format!(
                "unknown VCE role {role:?}"
            ))),
        }
    }
}

/// One scenario, as plain numbers, read from Python's `CoreScenario` by attribute.
/// Times are in whole seconds, and distances in feet along the platform.
#[derive(FromPyObject)]
struct Scenario {
    trains: usize,
    tracks: usize,
    headway: i64,
    /// When each train's departing passengers start coming down.
    release_times: Vec<i64>,
    arriving_pax_per_train: f64,
    departing_pax_per_train: f64,
    usable_area: f64,
    /// Each VCE's capacity in one direction (pax/s).
    vce_capacities: Vec<f64>,
    bidirectional_limits: Vec<f64>,
    roles: Vec<Role>,
    vce_west_ends: Vec<f64>,
    vce_east_ends: Vec<f64>,
    vce_choice_nearest: bool,
    /// Passengers (pax) at which a car counts as close to full.
    car_full: f64,
    /// Passengers (pax) still alighting below which reversible escalators go down.
    escalator_reversal_pax: f64,
    max_pax_in_stair_queues: f64,
    /// Steps after which the simulation gives up.
    max_steps: i64,

    platform_east_end: f64,
    platform_length: f64,
    train_length: f64,
    cars: usize,
    car_length: f64,
    doors_per_car: usize,
    doors_per_train: usize,
    /// Each door's flow rate (pax/s), alighting or boarding.
    door_flow_rate: f64,
    walking_speed: f64,

    /// Distance between the stopping positions `best_stopping_position` tries.
    stopping_position_step: f64,
    /// How many `stopping_position_step`s apart it tries them first.
    coarse_stopping_position_steps: usize,
    /// How much longer than the shortest a position's longest dwell can be.
    stopping_dwell_tolerance: i64,
    /// How far apart (pax) most occupants at once count as tied.
    stopping_occupants_tolerance: f64,
}

/// Where a train's doors and cars are, stopped at one position,
/// and how long each takes to walk to or from each VCE.
struct Layout {
    /// Each door's walking time to each VCE (s).
    door_walking_times: Vec<Vec<usize>>,
    /// Each car's maximum boarding rate (pax/s).
    car_board_rates: Vec<f64>,
    /// Each car's walking time from each VCE (s).
    car_walking_times: Vec<Vec<usize>>,
    /// Each VCE's cars, from the nearest to the farthest.
    nearest_cars: Vec<Vec<usize>>,
}

impl Scenario {
    /// Distance from `position` to the nearest end of VCE `vce`.
    fn distance_to(&self, vce: usize, position: f64) -> f64 {
        0.0_f64
            .max(self.vce_west_ends[vce] - position)
            .max(position - self.vce_east_ends[vce])
    }

    /// Time (s) to walk `distance`, rounded to the step.
    fn walking_time(&self, distance: f64) -> usize {
        (distance / self.walking_speed).round_ties_even() as usize
    }

    /// The train's doors spread evenly along it, stopped with its east end at `train_east_end`.
    fn layout(&self, train_east_end: f64) -> Layout {
        let train_west_end = train_east_end - self.train_length;
        let door_spacing = self.car_length / self.doors_per_car as f64;
        let doors: Vec<f64> = (0..self.doors_per_train)
            .map(|door| train_west_end + (door as f64 + 0.5) * door_spacing)
            .collect();
        let n_vces = self.vce_capacities.len();
        let door_walking_times = doors
            .iter()
            .map(|&door| {
                (0..n_vces)
                    .map(|vce| self.walking_time(self.distance_to(vce, door)))
                    .collect()
            })
            .collect();
        let car_doors: Vec<&[f64]> = doors.chunks(self.doors_per_car).take(self.cars).collect();
        let car_board_rates = car_doors
            .iter()
            .map(|doors| doors.len() as f64 * self.door_flow_rate)
            .collect();
        let car_walking_times: Vec<Vec<usize>> = car_doors
            .iter()
            .map(|doors| {
                (0..n_vces)
                    .map(|vce| {
                        let nearest = doors
                            .iter()
                            .map(|&door| self.distance_to(vce, door))
                            .fold(f64::INFINITY, f64::min);
                        self.walking_time(nearest)
                    })
                    .collect()
            })
            .collect();
        let nearest_cars = (0..n_vces)
            .map(|vce| {
                let mut cars: Vec<usize> = (0..car_doors.len()).collect();
                cars.sort_by_key(|&car| car_walking_times[car][vce]);
                cars
            })
            .collect();
        Layout {
            door_walking_times,
            car_board_rates,
            car_walking_times,
            nearest_cars,
        }
    }

    /// Every stopping position with the train on the platform, from its east end west,
    /// letting a train longer than the platform overhang its west end.
    fn stopping_positions(&self) -> Vec<f64> {
        let platform_west_end = self.platform_east_end - self.platform_length;
        let overhang = (self.train_length - self.platform_length).max(0.0);
        let mut positions = Vec::new();
        let mut train_east_end = self.platform_east_end;
        while train_east_end - self.train_length >= platform_west_end - overhang {
            positions.push(train_east_end);
            train_east_end -= self.stopping_position_step;
        }
        positions
    }

    /// The stopping position where the platform could be evacuated soonest under NFPA 130,
    /// i.e. with the fewest occupants at once,
    /// of the positions where the longest dwell is within `stopping_dwell_tolerance`
    /// of the shortest it can be, counting occupants within `stopping_occupants_tolerance` as tied,
    /// and if tied, where the longest and then the total dwell are shortest.
    /// It tries every `coarse_stopping_position_steps`th position first,
    /// then every position within that many of the best of those.
    fn best_stopping_position(&self) -> f64 {
        let positions = self.stopping_positions();
        let mut scores: Vec<Option<Score>> = vec![None; positions.len()];
        let coarse_step = self.coarse_stopping_position_steps;
        let coarse: Vec<usize> = (0..positions.len()).step_by(coarse_step).collect();
        self.score_all(&positions, &coarse, &mut scores);
        let best_coarse = self.best(&coarse, &scores);
        let nearby: Vec<usize> = (best_coarse.saturating_sub(coarse_step)
            ..(best_coarse + coarse_step + 1).min(positions.len()))
            .collect();
        self.score_all(&positions, &nearby, &mut scores);
        // Every position tried counts toward the shortest longest dwell,
        // and of any tied, the first tried wins, i.e. the coarse ones first.
        let tried: Vec<usize> = coarse.into_iter().chain(nearby).collect();
        positions[self.best(&tried, &scores)]
    }

    /// Score each of `indices`' `positions` not scored yet, in parallel.
    fn score_all(&self, positions: &[f64], indices: &[usize], scores: &mut [Option<Score>]) {
        let new: Vec<(usize, Score)> = indices
            .par_iter()
            .filter(|&&i| scores[i].is_none())
            .map(|&i| (i, self.score(positions[i])))
            .collect();
        for (i, score) in new {
            scores[i] = Some(score);
        }
    }

    fn score(&self, train_east_end: f64) -> Score {
        let result = simulate(self, train_east_end, false);
        let occupants = result.max_occupants;
        let dwells: Option<Vec<i64>> = result.dwells.iter().copied().collect();
        match dwells {
            Some(dwells) => Score {
                longest: dwells.iter().copied().max().unwrap_or(0),
                total: dwells.iter().sum(),
                occupants,
            },
            None => Score {
                longest: i64::MAX,
                total: i64::MAX,
                occupants,
            },
        }
    }

    /// The best of `of`, per `best_stopping_position`, the first of any tied.
    fn best(&self, of: &[usize], scores: &[Option<Score>]) -> usize {
        let score = |i: usize| scores[i].expect("every position is scored first");
        let shortest = of
            .iter()
            .map(|&i| score(i).longest)
            .min()
            .expect("a position");
        let within: Vec<usize> = of
            .iter()
            .copied()
            .filter(|&i| score(i).longest - self.stopping_dwell_tolerance <= shortest)
            .collect();
        let fewest = within
            .iter()
            .map(|&i| score(i).occupants)
            .fold(f64::INFINITY, f64::min);
        within
            .into_iter()
            .filter(|&i| score(i).occupants - self.stopping_occupants_tolerance <= fewest)
            .min_by_key(|&i| (score(i).longest, score(i).total))
            .expect("a position")
    }
}

/// How good one stopping position is, per `Scenario::best`.
#[derive(Clone, Copy)]
struct Score {
    longest: i64,
    total: i64,
    occupants: f64,
}

/// What `simulate_core` returns: the summary, the conservation totals,
/// and with `record_time_series`, each second's values, by column.
#[pyclass(get_all, frozen)]
#[derive(Default)]
struct CoreResult {
    /// Whether the last train departed and the platform cleared before `max_steps`.
    finished: bool,
    max_up_rate: f64,
    time_at_capacity: i64,
    taper_time: Option<i64>,
    clear_time: Option<i64>,
    arrival_times: Vec<Option<i64>>,
    dwells: Vec<Option<i64>>,
    boarded_time: Option<i64>,
    max_pax_on_platform: f64,
    max_occupants: f64,
    min_space_per_pax: f64,
    vce_gone_up: Vec<f64>,
    vce_empty_times: Vec<Option<i64>>,

    gone_up: f64,
    still_aboard: f64,
    arriving_pax_on_platform: f64,
    boarded: f64,
    still_upstairs: f64,
    still_walking_to_cars: f64,
    still_waiting_at_cars: f64,

    times: Vec<i64>,
    arriving_pax_waiting_on_platform: Vec<f64>,
    off_rate: Vec<f64>,
    on_rate: Vec<f64>,
    down_rate: Vec<f64>,
    departing_pax_on_platform: Vec<f64>,
    total_pax_on_platform: Vec<f64>,
    platform_crowding: Vec<f64>,
    up_rate: Vec<f64>,
    net_pax_flow_rate: Vec<f64>,
    /// Each second, each VCE's queue (pax) and upward flow (pax/s).
    vces: Vec<Vec<[f64; 2]>>,
    /// Each second, each train's `TRAIN_COLUMNS`.
    train_values: Vec<Vec<[f64; 4]>>,
}

fn calc_space_per_pax(pax_on_platform: f64, area: f64) -> f64 {
    if pax_on_platform > 0.0 {
        area / pax_on_platform
    } else {
        area
    }
}

fn choose_vce(
    nearest: bool,
    walking_times: &[usize],
    waits: &[f64],
    directions: &[Direction],
) -> usize {
    let mut best = 0;
    let mut best_time = f64::INFINITY;
    for (i, ((&walk, &wait), &direction)) in
        walking_times.iter().zip(waits).zip(directions).enumerate()
    {
        let walk = walk as f64;
        let time = if direction == Direction::Down {
            f64::INFINITY
        } else if nearest {
            walk
        } else {
            walk.max(wait)
        };
        // The first of any tied, i.e. the westernmost.
        if i == 0 || time < best_time {
            best = i;
            best_time = time;
        }
    }
    best
}

fn choose_car(nearest: &[usize], car_loads: &[f64], full: f64) -> usize {
    if nearest.len() == 1 {
        return nearest[0];
    }
    nearest
        .iter()
        .copied()
        .find(|&c| car_loads[c] < full)
        .unwrap_or(nearest[0])
}

/// Passengers walking toward somewhere, by the step they reach it,
/// in a ring of the next `len` steps.
struct Arrivals<T> {
    slots: Vec<Option<T>>,
}

impl<T> Arrivals<T> {
    fn new(longest_walk: usize) -> Self {
        Self {
            slots: (0..=longest_walk).map(|_| None).collect(),
        }
    }

    fn at(&mut self, step: usize, empty: impl FnOnce() -> T) -> &mut T {
        let len = self.slots.len();
        self.slots[step % len].get_or_insert_with(empty)
    }

    fn pop(&mut self, step: usize) -> Option<T> {
        let len = self.slots.len();
        self.slots[step % len].take()
    }
}

fn simulate(input: &Scenario, train_east_end: f64, record: bool) -> CoreResult {
    let layout = input.layout(train_east_end);
    let door_share = 1.0 / layout.door_walking_times.len() as f64;
    let n_vces = input.vce_capacities.len();
    let n_cars = layout.car_board_rates.len();
    let trains = input.trains;

    let mut arriving_pax_on_platform = 0.0_f64;
    let mut vce_queues = vec![0.0_f64; n_vces];
    let mut directions: Vec<Direction> = input
        .roles
        .iter()
        .map(|role| match role {
            Role::Stair => Direction::Both,
            Role::Down => Direction::Down,
            Role::Up | Role::Reversible => Direction::Up,
        })
        .collect();
    let door_rate = input.doors_per_train as f64 * input.door_flow_rate;
    let reversible: Vec<usize> = (0..n_vces)
        .filter(|&i| input.roles[i] == Role::Reversible)
        .collect();
    let longest_door_walk = layout
        .door_walking_times
        .iter()
        .flatten()
        .copied()
        .max()
        .unwrap_or(0);
    let mut walking: Arrivals<Vec<f64>> = Arrivals::new(longest_door_walk);
    let mut walking_totals = vec![0.0_f64; n_vces];
    let longest_car_walk = layout
        .car_walking_times
        .iter()
        .flatten()
        .copied()
        .max()
        .unwrap_or(0);
    let mut boarders_walking: Arrivals<Vec<Vec<f64>>> = Arrivals::new(longest_car_walk);
    let mut boarders_walking_totals = vec![0.0_f64; trains];
    let mut car_waiting = vec![vec![0.0_f64; n_cars]; trains];
    let mut car_loads = vec![vec![0.0_f64; n_cars]; trains];
    let mut arrival_times: Vec<Option<i64>> = (0..trains)
        .map(|train| (train < input.tracks).then(|| train as i64 * input.headway))
        .collect();
    let mut remaining_arrivals = vec![input.arriving_pax_per_train; trains];
    let mut new_pax = vec![0.0_f64; trains];
    let release_times = &input.release_times;
    let start_time = release_times.iter().copied().fold(0, i64::min);
    let last_release = release_times.iter().copied().max().unwrap_or(0);
    let mut boarders_upstairs = vec![0.0_f64; trains];
    let mut boarders_on_platform = vec![0.0_f64; trains];
    let mut total_pax_on_platform = 0.0_f64;
    let mut gone_up = 0.0_f64;
    let no_flow = vec![0.0_f64; n_vces];

    let mut summary = CoreResult {
        dwells: vec![None; trains],
        max_pax_on_platform: total_pax_on_platform,
        max_occupants: total_pax_on_platform,
        min_space_per_pax: calc_space_per_pax(total_pax_on_platform, input.usable_area),
        vce_gone_up: vec![0.0; n_vces],
        vce_empty_times: vec![None; n_vces],
        ..CoreResult::default()
    };

    let mut off_rates = vec![0.0_f64; trains];
    let mut vce_up_rates = vec![0.0_f64; n_vces];
    let mut waits = vec![0.0_f64; n_vces];
    let mut vce_down_capacities = vec![0.0_f64; n_vces];
    let mut down_rates = vec![0.0_f64; trains];
    let mut car_on_rates = vec![vec![0.0_f64; n_cars]; trains];
    let mut on_rates = vec![0.0_f64; trains];

    let mut step: usize = 0;
    loop {
        let time_after = start_time + step as i64;
        if time_after - start_time >= input.max_steps {
            break;
        }
        for train in 0..trains {
            if release_times[train] == time_after {
                boarders_upstairs[train] = input.departing_pax_per_train;
            }
        }
        for train in 0..trains {
            let off_rate = match arrival_times[train] {
                Some(arrival_time) if time_after > arrival_time => {
                    remaining_arrivals[train].min(door_rate)
                }
                _ => 0.0,
            };
            remaining_arrivals[train] -= off_rate;
            if remaining_arrivals[train] < 0.0 {
                remaining_arrivals[train] = 0.0;
            }
            off_rates[train] = off_rate;
        }
        total_pax_on_platform += off_rates.iter().sum::<f64>();
        arriving_pax_on_platform += off_rates.iter().sum::<f64>();
        if !reversible.is_empty() {
            let mut still_alighting = arriving_pax_on_platform;
            for train in 0..trains {
                if let Some(arrival_time) = arrival_times[train]
                    && time_after >= arrival_time
                {
                    still_alighting += remaining_arrivals[train];
                }
            }
            let nearly_alighted = still_alighting < input.escalator_reversal_pax;
            for &i in &reversible {
                // Once reversed, nobody walks to it, so it stays down until more alight.
                directions[i] = if !nearly_alighted || vce_queues[i] + walking_totals[i] > 0.0 {
                    Direction::Up
                } else {
                    Direction::Down
                };
            }
        }
        let up_rates: &[f64] = if arriving_pax_on_platform > 0.0 {
            // Each door's alighting passengers walk to a VCE,
            // though nobody does in a second nobody alights.
            let alighting = off_rates.iter().sum::<f64>();
            if alighting > 0.0 {
                for i in 0..n_vces {
                    waits[i] = (vce_queues[i] + walking_totals[i]) / input.vce_capacities[i];
                }
                for walking_times in &layout.door_walking_times {
                    let i =
                        choose_vce(input.vce_choice_nearest, walking_times, &waits, &directions);
                    walking.at(step + walking_times[i], || vec![0.0; n_vces])[i] +=
                        alighting * door_share;
                    walking_totals[i] += alighting * door_share;
                    waits[i] = (vce_queues[i] + walking_totals[i]) / input.vce_capacities[i];
                }
            }
            if let Some(reaching) = walking.pop(step) {
                for (i, reaching) in reaching.into_iter().enumerate() {
                    vce_queues[i] += reaching;
                    walking_totals[i] = subtract(walking_totals[i], reaching);
                }
            }
            for i in 0..n_vces {
                let queue = vce_queues[i];
                let vce_up_rate = queue.min(input.vce_capacities[i]);
                vce_queues[i] = queue - vce_up_rate;
                vce_up_rates[i] = vce_up_rate;
                summary.vce_gone_up[i] += vce_up_rate;
                if vce_up_rate > 0.0 && vce_queues[i] <= ROUNDING_TOLERANCE {
                    summary.vce_empty_times[i] = Some(time_after);
                }
            }
            &vce_up_rates
        } else {
            // Nobody's on the platform, so nobody can be walking to a VCE.
            walking.pop(step);
            &no_flow
        };
        let up_rate = up_rates.iter().sum::<f64>();
        gone_up += up_rate;
        arriving_pax_on_platform -= up_rate;
        if arriving_pax_on_platform < 0.0 {
            arriving_pax_on_platform = 0.0;
        }
        total_pax_on_platform -= up_rate;
        for i in 0..n_vces {
            vce_down_capacities[i] =
                if directions[i] == Direction::Up || up_rates[i] > input.bidirectional_limits[i] {
                    0.0
                } else {
                    input.vce_capacities[i] - up_rates[i]
                };
        }
        let down_capacity = vce_down_capacities.iter().sum::<f64>();
        let all_upstairs = boarders_upstairs.iter().sum::<f64>();
        for train in 0..trains {
            // A total of only rounding errors would give a share of nearly 1 / 0.
            let fraction = if all_upstairs > ROUNDING_TOLERANCE {
                boarders_upstairs[train] / all_upstairs
            } else {
                1.0
            };
            down_rates[train] = boarders_upstairs[train].min(down_capacity * fraction);
        }
        // Departing passengers come down each VCE in proportion to the capacity it has left,
        // and walk from it to a car.
        for train in 0..trains {
            if down_rates[train] <= 0.0 {
                continue;
            }
            for (i, &vce_down_capacity) in vce_down_capacities.iter().enumerate() {
                if vce_down_capacity <= 0.0 {
                    continue;
                }
                let rate = down_rates[train] * vce_down_capacity / down_capacity;
                let car = choose_car(&layout.nearest_cars[i], &car_loads[train], input.car_full);
                boarders_walking.at(step + layout.car_walking_times[car][i], || {
                    vec![vec![0.0; n_cars]; trains]
                })[train][car] += rate;
                boarders_walking_totals[train] += rate;
                car_loads[train][car] += rate;
            }
        }
        if let Some(reaching) = boarders_walking.pop(step) {
            for (train, reaching) in reaching.into_iter().enumerate() {
                for (car, pax) in reaching.into_iter().enumerate() {
                    if pax == 0.0 {
                        continue;
                    }
                    car_waiting[train][car] += pax;
                    boarders_walking_totals[train] = subtract(boarders_walking_totals[train], pax);
                }
            }
        }
        for train in 0..trains {
            boarders_on_platform[train] = car_waiting[train].iter().sum::<f64>();
            total_pax_on_platform += down_rates[train];
        }
        // Each car boards through its own doors,
        // once its train has arrived and everyone has alighted from it.
        for train in 0..trains {
            let boarding = matches!(arrival_times[train], Some(arrival_time) if arrival_time < time_after)
                && off_rates[train] == 0.0;
            for car in 0..n_cars {
                let waiting = car_waiting[train][car];
                car_on_rates[train][car] = if boarding && waiting > 0.0 {
                    layout.car_board_rates[car].min(waiting)
                } else {
                    0.0
                };
            }
        }
        for train in 0..trains {
            for car in 0..n_cars {
                car_waiting[train][car] -= car_on_rates[train][car];
            }
            on_rates[train] = car_on_rates[train].iter().sum::<f64>();
        }

        for train in 0..trains {
            boarders_on_platform[train] -= on_rates[train];
            total_pax_on_platform -= on_rates[train];
            boarders_upstairs[train] -= down_rates[train];
            new_pax[train] += on_rates[train];
        }

        let space_per_pax = calc_space_per_pax(total_pax_on_platform, input.usable_area);
        if total_pax_on_platform < 0.0 {
            total_pax_on_platform = 0.0;
        }
        for train in 0..trains {
            if boarders_on_platform[train] < 0.0 {
                boarders_on_platform[train] = 0.0;
            }
        }
        if arriving_pax_on_platform < 0.0 {
            arriving_pax_on_platform = 0.0;
        }
        summary.max_up_rate = summary.max_up_rate.max(up_rate);
        // Capacity of the VCEs going up now, not counting escalators going down.
        let capacity: f64 = input
            .vce_capacities
            .iter()
            .zip(&directions)
            .filter(|&(_, &direction)| direction != Direction::Down)
            .map(|(&capacity, _)| capacity)
            .sum();
        if up_rate >= capacity - 1e-9 {
            summary.time_at_capacity += 1;
        }
        if arriving_pax_on_platform > input.max_pax_in_stair_queues {
            summary.taper_time = Some(time_after);
        }
        if summary.clear_time.is_none()
            && arrival_times
                .iter()
                .all(|&arrival_time| matches!(arrival_time, Some(arrival_time) if time_after > arrival_time))
            && arriving_pax_on_platform < 1.0
        {
            summary.clear_time = Some(time_after);
        }
        if summary.boarded_time.is_none()
            && time_after >= last_release
            && boarders_upstairs.iter().sum::<f64>()
                + boarders_on_platform.iter().sum::<f64>()
                + boarders_walking_totals.iter().sum::<f64>()
                < 1.0
        {
            summary.boarded_time = Some(time_after);
        }
        for train in 0..trains {
            if summary.dwells[train].is_none()
                && let Some(arrival_time) = arrival_times[train]
                && time_after > arrival_time
                && remaining_arrivals[train] < 1.0
                && boarders_upstairs[train]
                    + boarders_walking_totals[train]
                    + boarders_on_platform[train]
                    < 1.0
            {
                summary.dwells[train] = Some(time_after - arrival_time);
                // The next train on its track arrives once it's scheduled and this one departs.
                let next = train + input.tracks;
                if next < trains {
                    arrival_times[next] = Some((next as i64 * input.headway).max(time_after));
                }
            }
        }
        summary.max_pax_on_platform = summary.max_pax_on_platform.max(total_pax_on_platform);
        let aboard: f64 = (0..trains)
            .filter_map(|train| match arrival_times[train] {
                Some(arrival_time)
                    if time_after >= arrival_time && summary.dwells[train].is_none() =>
                {
                    Some(remaining_arrivals[train] + new_pax[train])
                }
                _ => None,
            })
            .sum();
        summary.max_occupants = summary.max_occupants.max(total_pax_on_platform + aboard);
        summary.min_space_per_pax = summary.min_space_per_pax.min(space_per_pax);

        if record {
            let mut net_pax_flow_rate = 0.0_f64;
            for &rate in &down_rates {
                net_pax_flow_rate += rate;
            }
            for &rate in &off_rates {
                net_pax_flow_rate += rate;
            }
            net_pax_flow_rate -= up_rate;
            for &rate in &on_rates {
                net_pax_flow_rate -= rate;
            }
            summary.times.push(time_after);
            summary
                .arriving_pax_waiting_on_platform
                .push(arriving_pax_on_platform);
            summary.off_rate.push(off_rates.iter().sum::<f64>());
            summary.on_rate.push(on_rates.iter().sum::<f64>());
            summary.down_rate.push(down_rates.iter().sum::<f64>());
            summary.departing_pax_on_platform.push(
                boarders_on_platform.iter().sum::<f64>()
                    + boarders_walking_totals.iter().sum::<f64>(),
            );
            summary.total_pax_on_platform.push(total_pax_on_platform);
            summary.platform_crowding.push(space_per_pax);
            summary.up_rate.push(up_rate);
            summary.net_pax_flow_rate.push(net_pax_flow_rate);
            summary.vces.push(
                vce_queues
                    .iter()
                    .zip(up_rates)
                    .map(|(&queue, &rate)| [queue, rate])
                    .collect(),
            );
            summary.train_values.push(
                (0..trains)
                    .map(|train| {
                        [
                            remaining_arrivals[train] + new_pax[train],
                            off_rates[train],
                            on_rates[train],
                            boarders_on_platform[train] + boarders_walking_totals[train],
                        ]
                    })
                    .collect(),
            );
        }

        // Stop once the last train has departed and the platform has cleared.
        if summary.clear_time.is_some()
            && summary.boarded_time.is_some()
            && summary.dwells.iter().all(Option::is_some)
        {
            summary.finished = true;
            break;
        }
        step += 1;
    }
    summary.arrival_times = arrival_times;
    summary.gone_up = gone_up;
    summary.still_aboard = remaining_arrivals.iter().sum::<f64>();
    summary.arriving_pax_on_platform = arriving_pax_on_platform;
    summary.boarded = new_pax.iter().sum::<f64>();
    summary.still_upstairs = boarders_upstairs.iter().sum::<f64>();
    summary.still_walking_to_cars = boarders_walking_totals.iter().sum::<f64>();
    summary.still_waiting_at_cars = car_waiting.iter().flatten().sum();
    summary
}

#[pymodule]
mod _core {
    use super::*;

    /// Simulate `scenario` with its trains stopped with their east ends at `train_east_end`,
    /// without holding the GIL.
    #[pyfunction]
    fn simulate_core(
        py: Python<'_>,
        scenario: Scenario,
        train_east_end: f64,
        record_time_series: bool,
    ) -> CoreResult {
        py.detach(|| simulate(&scenario, train_east_end, record_time_series))
    }

    /// Where `scenario`'s trains should stop, per `Scenario::best_stopping_position`,
    /// trying positions in parallel without holding the GIL.
    #[pyfunction]
    fn best_stopping_position(py: Python<'_>, scenario: Scenario) -> f64 {
        py.detach(|| scenario.best_stopping_position())
    }
}
