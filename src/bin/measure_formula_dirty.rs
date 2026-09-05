//! Measurement-only formula dirty-propagation calibration.
//!
//! Emits p50/p95 wall-clock samples for dirty and forced full-plan recalculation,
//! and refuses to report success if both paths do not produce identical values.

use elixcee::types::Variant;
use elixcee::vm::{CalculationMode, Vm};
use std::hint::black_box;
use std::time::Instant;

const FORMULA_COUNTS: [u32; 2] = [100, 1_000];

#[cfg(unix)]
fn resource_stats() -> (u64, u64, u64) {
    let mut usage = std::mem::MaybeUninit::<libc::rusage>::zeroed();
    // SAFETY: `getrusage` initializes the supplied structure for RUSAGE_SELF.
    let result = unsafe { libc::getrusage(libc::RUSAGE_SELF, usage.as_mut_ptr()) };
    if result != 0 {
        return (0, 0, 0);
    }
    // SAFETY: getrusage returned success and initialized the structure.
    let usage = unsafe { usage.assume_init() };
    let user_us = (usage.ru_utime.tv_sec as u64)
        .saturating_mul(1_000_000)
        .saturating_add(usage.ru_utime.tv_usec as u64);
    let system_us = (usage.ru_stime.tv_sec as u64)
        .saturating_mul(1_000_000)
        .saturating_add(usage.ru_stime.tv_usec as u64);
    // macOS reports bytes; Linux and the other Unix targets report KiB.
    let peak_rss_bytes = if cfg!(target_os = "macos") {
        usage.ru_maxrss as u64
    } else {
        (usage.ru_maxrss as u64).saturating_mul(1024)
    };
    (peak_rss_bytes, user_us, system_us)
}

#[cfg(not(unix))]
fn resource_stats() -> (u64, u64, u64) {
    (0, 0, 0)
}

fn percentile(samples: &[f64], percentile: usize) -> f64 {
    let mut sorted = samples.to_vec();
    sorted.sort_by(f64::total_cmp);
    let index = ((sorted.len() * percentile).saturating_sub(1)) / 100;
    sorted[index]
}

fn chain_template(formula_count: u32) -> Vm {
    let mut vm = Vm::new();
    vm.write_rect("sheet1", (1, 1), &[vec![Variant::Integer(1)]]);
    for col in 2..=formula_count + 1 {
        let previous = column_name(col - 1);
        vm.set_cell_formula(1, col, &format!("={previous}1+1"))
            .expect("chain formula");
    }
    vm.recalculate_all().expect("warm formula plan");
    vm
}

fn column_name(mut column: u32) -> String {
    let mut out = String::new();
    while column > 0 {
        let remainder = (column - 1) % 26;
        out.push((b'A' + remainder as u8) as char);
        column = (column - 1) / 26;
    }
    out.chars().rev().collect()
}

fn measure(iterations: usize, formula_count: u32, template: &Vm) -> (Vec<f64>, Vec<f64>) {
    let mut dirty_samples = Vec::with_capacity(iterations);
    let mut full_samples = Vec::with_capacity(iterations);
    for iteration in 0..iterations {
        let mut dirty = template.clone();
        dirty.write_rect(
            "sheet1",
            (1, 1),
            &[vec![Variant::Integer(10 + iteration as i64)]],
        );
        let started = Instant::now();
        dirty.recalculate_all().expect("dirty recalculation");
        dirty_samples.push(black_box(started.elapsed().as_secs_f64() * 1000.0));

        let mut full = template.clone();
        full.write_rect(
            "sheet1",
            (1, 1),
            &[vec![Variant::Integer(10 + iteration as i64)]],
        );
        // Keep the graph unchanged while invalidating the persistent plan.
        full.set_cell_formula(
            1,
            formula_count + 1,
            &format!("={}1+1", column_name(formula_count)),
        )
        .expect("full-plan invalidation");
        let started = Instant::now();
        full.recalculate_all().expect("full recalculation");
        full_samples.push(black_box(started.elapsed().as_secs_f64() * 1000.0));

        for col in [1, 2, formula_count / 2, formula_count + 1] {
            assert_eq!(
                dirty.get_cell(1, col),
                full.get_cell(1, col),
                "dirty/full mismatch at column {col}"
            );
        }
    }
    (dirty_samples, full_samples)
}

fn measure_manual_transition(iterations: usize, template: &Vm) -> Vec<f64> {
    let mut samples = Vec::with_capacity(iterations);
    for iteration in 0..iterations {
        let mut vm = template.clone();
        vm.set_calc_mode(CalculationMode::Manual)
            .expect("manual calculation mode");
        vm.write_rect(
            "sheet1",
            (1, 1),
            &[vec![Variant::Integer(20 + iteration as i64)]],
        );
        let started = Instant::now();
        vm.set_calc_mode(CalculationMode::Automatic)
            .expect("automatic calculation mode");
        samples.push(black_box(started.elapsed().as_secs_f64() * 1000.0));
    }
    samples
}

fn measure_cycle(iterations: usize) -> Vec<f64> {
    let mut samples = Vec::with_capacity(iterations);
    for _ in 0..iterations {
        let mut vm = Vm::new();
        vm.set_cell_formula(1, 1, "=B1+1")
            .expect("first cycle formula");
        vm.set_cell_formula(1, 2, "=A1+1")
            .expect("second cycle formula");
        let started = Instant::now();
        vm.recalculate_all().expect("cycle recalculation");
        samples.push(black_box(started.elapsed().as_secs_f64() * 1000.0));
    }
    samples
}

fn main() {
    let mut args = std::env::args().skip(1);
    let iterations = args
        .next()
        .unwrap_or_else(|| usage("missing iteration count"))
        .parse::<usize>()
        .unwrap_or_else(|_| usage("iteration count must be a positive integer"));
    if iterations < 3 || args.next().is_some() {
        usage("iterations must be at least 3 and no extra arguments are allowed");
    }

    let started = Instant::now();
    let mut cases = Vec::with_capacity(FORMULA_COUNTS.len());
    let mut large_template = None;
    for formula_count in FORMULA_COUNTS {
        let template = chain_template(formula_count);
        let (dirty, full) = measure(iterations, formula_count, &template);
        if formula_count == 1_000 {
            large_template = Some(template);
        }
        cases.push(format!(
            "{{\"formula_count\":{},\"dirty_p50_ms\":{:.6},\"dirty_p95_ms\":{:.6},\"full_p50_ms\":{:.6},\"full_p95_ms\":{:.6}}}",
            formula_count,
            percentile(&dirty, 50),
            percentile(&dirty, 95),
            percentile(&full, 50),
            percentile(&full, 95),
        ));
    }
    let manual = measure_manual_transition(iterations, large_template.as_ref().unwrap());
    let cycle = measure_cycle(iterations);
    let (peak_rss_bytes, user_cpu_us, system_cpu_us) = resource_stats();
    println!(
        "{{\"iterations\":{},\"cases\":[{}],\"manual_to_automatic_p50_ms\":{:.6},\"manual_to_automatic_p95_ms\":{:.6},\"cycle_p50_ms\":{:.6},\"cycle_p95_ms\":{:.6},\"peak_rss_bytes\":{},\"user_cpu_us\":{},\"system_cpu_us\":{},\"resource_stats_supported\":{},\"wall_ms\":{:.3}}}",
        iterations,
        cases.join(","),
        percentile(&manual, 50),
        percentile(&manual, 95),
        percentile(&cycle, 50),
        percentile(&cycle, 95),
        peak_rss_bytes,
        user_cpu_us,
        system_cpu_us,
        cfg!(unix),
        started.elapsed().as_secs_f64() * 1000.0,
    );
}

fn usage(message: &str) -> ! {
    eprintln!("{message}\nusage: measure_formula_dirty <iterations>");
    std::process::exit(2);
}
