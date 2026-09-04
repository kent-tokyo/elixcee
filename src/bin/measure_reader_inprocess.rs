//! Measurement-only in-process reader soak. This binary is not part of the public API.

use std::hint::black_box;
use std::time::Instant;

#[cfg(target_os = "macos")]
#[repr(C)]
struct MallocStatistics {
    blocks_in_use: usize,
    size_in_use: usize,
    max_size_in_use: usize,
    size_allocated: usize,
}

#[cfg(target_os = "macos")]
#[repr(C)]
struct MallocZone {
    _private: [u8; 0],
}

#[cfg(target_os = "macos")]
unsafe extern "C" {
    fn malloc_default_zone() -> *mut MallocZone;
    fn malloc_zone_statistics(zone: *mut MallocZone, stats: *mut MallocStatistics);
}

#[cfg(target_os = "macos")]
fn allocator_stats() -> Option<MallocStatistics> {
    let mut stats = MallocStatistics {
        blocks_in_use: 0,
        size_in_use: 0,
        max_size_in_use: 0,
        size_allocated: 0,
    };
    unsafe { malloc_zone_statistics(malloc_default_zone(), &mut stats) };
    Some(stats)
}

#[cfg(not(target_os = "macos"))]
fn allocator_stats() -> Option<()> {
    None
}

fn json_string(value: &str) -> String {
    format!("\"{}\"", value.replace('\\', "\\\\").replace('"', "\\\""))
}

fn main() {
    let mut args = std::env::args().skip(1);
    let fixture = args.next().unwrap_or_else(|| usage("missing fixture path"));
    let iterations: usize = args
        .next()
        .unwrap_or_else(|| usage("missing iteration count"))
        .parse()
        .unwrap_or_else(|_| usage("iteration count must be a positive integer"));
    if iterations < 2 {
        usage("iteration count must be at least 2");
    }
    if args.next().is_some() {
        usage("too many arguments");
    }

    let bytes = std::fs::read(&fixture).unwrap_or_else(|error| {
        eprintln!("cannot read {fixture}: {error}");
        std::process::exit(2);
    });
    let input_bytes = bytes.len();
    let baseline = allocator_stats();
    let started = Instant::now();
    let mut successes = 0usize;
    let mut total_cells = 0usize;
    let mut observations = Vec::with_capacity(iterations);
    for iteration in 1..=iterations {
        let iteration_started = Instant::now();
        let workbook = elixcee::reader::read_workbook_from_bytes(&bytes).unwrap_or_else(|error| {
            eprintln!("iteration {iteration} failed: {error}");
            std::process::exit(1);
        });
        let cells: usize = workbook
            .sheets
            .iter()
            .map(|sheet| sheet.sheet.cells.len())
            .sum();
        total_cells += black_box(cells);
        successes += 1;
        drop(workbook);
        #[cfg(target_os = "macos")]
        let after = allocator_stats();
        #[cfg(target_os = "macos")]
        let allocator_json = after.map(|stats| {
            format!(
                "\"allocator_size_in_use_bytes\":{},\"allocator_size_allocated_bytes\":{},\"allocator_blocks_in_use\":{}",
                stats.size_in_use, stats.size_allocated, stats.blocks_in_use
            )
        });
        #[cfg(not(target_os = "macos"))]
        let allocator_json: Option<String> = None;
        let allocator_json =
            allocator_json.unwrap_or_else(|| "\"allocator_stats\":null".to_string());
        observations.push(format!(
            "{{\"iteration\":{},\"cells\":{},\"wall_ms\":{:.3},{} }}",
            iteration,
            cells,
            iteration_started.elapsed().as_secs_f64() * 1000.0,
            allocator_json
        ));
    }

    #[cfg(target_os = "macos")]
    let baseline_json = baseline.map(|stats| {
        format!(
            "\"baseline_allocator_size_in_use_bytes\":{},\"baseline_allocator_size_allocated_bytes\":{}",
            stats.size_in_use, stats.size_allocated
        )
    });
    #[cfg(not(target_os = "macos"))]
    let baseline_json: Option<String> = None;
    let baseline_json =
        baseline_json.unwrap_or_else(|| "\"allocator_stats_supported\":false".to_string());
    println!(
        "{{\"fixture\":{},\"input_bytes\":{},\"iterations\":{},\"successes\":{},\"total_cells\":{},\"wall_ms\":{:.3},{},\"observations\":[{}]}}",
        json_string(&fixture),
        input_bytes,
        iterations,
        successes,
        total_cells,
        started.elapsed().as_secs_f64() * 1000.0,
        baseline_json,
        observations.join(",")
    );
}

fn usage(message: &str) -> ! {
    eprintln!("{message}\nusage: measure_reader_inprocess <fixture.xlsx> <iterations>");
    std::process::exit(2);
}
