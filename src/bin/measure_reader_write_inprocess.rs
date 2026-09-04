//! Measurement-only in-process read/mutate/write soak.

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
fn allocator_stats() -> MallocStatistics {
    let mut stats = MallocStatistics {
        blocks_in_use: 0,
        size_in_use: 0,
        max_size_in_use: 0,
        size_allocated: 0,
    };
    unsafe { malloc_zone_statistics(malloc_default_zone(), &mut stats) };
    stats
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
    if iterations < 2 || args.next().is_some() {
        usage("iterations must be at least 2 and no extra arguments are allowed");
    }

    #[cfg(target_os = "macos")]
    let baseline = allocator_stats();
    let started = Instant::now();
    let mut observations = Vec::with_capacity(iterations);
    for iteration in 1..=iterations {
        let iteration_started = Instant::now();
        let mut vm = elixcee::vm::Vm::new();
        vm.load_workbook_file(&fixture).unwrap_or_else(|error| {
            eprintln!("iteration {iteration} load failed: {error}");
            std::process::exit(1);
        });
        let sheet = vm.sheet_names().into_iter().next().unwrap_or_else(|| {
            eprintln!("iteration {iteration}: workbook has no sheets");
            std::process::exit(1);
        });
        vm.write_rect(
            &sheet,
            (1, 1),
            &[vec![elixcee::types::Variant::Integer(iteration as i64)]],
        );
        vm.set_cell_formula(1, 2, "1+2").unwrap_or_else(|error| {
            eprintln!("iteration {iteration} formula mutation failed: {error}");
            std::process::exit(1);
        });
        let output = std::env::temp_dir().join(format!(
            "elixcee-inprocess-write-{}-{}.xlsx",
            std::process::id(),
            iteration
        ));
        elixcee::save_workbook(&vm, output.to_str().unwrap()).unwrap_or_else(|error| {
            eprintln!("iteration {iteration} save failed: {error}");
            std::process::exit(1);
        });
        drop(vm);
        let saved =
            elixcee::reader::read_workbook(output.to_str().unwrap()).unwrap_or_else(|error| {
                eprintln!("iteration {iteration} verification read failed: {error}");
                std::process::exit(1);
            });
        let cells = saved.iter().map(|sheet| sheet.cells.len()).sum::<usize>();
        let verified = saved.first().is_some_and(|sheet| {
            matches!(
                sheet.cells.get(&(1, 1)),
                Some(elixcee::reader::SheetCell::Integer(value)) if *value == iteration as i64
            ) && sheet.cells.contains_key(&(1, 2))
        });
        if !verified {
            eprintln!("iteration {iteration}: saved mutation did not round-trip");
            std::process::exit(1);
        }
        drop(saved);
        let _ = std::fs::remove_file(&output);
        #[cfg(target_os = "macos")]
        let stats = allocator_stats();
        #[cfg(target_os = "macos")]
        let allocator_json = format!(
            "\"allocator_size_in_use_bytes\":{},\"allocator_size_allocated_bytes\":{}",
            stats.size_in_use, stats.size_allocated
        );
        #[cfg(not(target_os = "macos"))]
        let allocator_json = "\"allocator_stats\":null".to_string();
        observations.push(format!(
            "{{\"iteration\":{},\"cells\":{},\"wall_ms\":{:.3},{} }}",
            iteration,
            black_box(cells),
            iteration_started.elapsed().as_secs_f64() * 1000.0,
            allocator_json
        ));
    }

    #[cfg(target_os = "macos")]
    let baseline_json = format!(
        "\"baseline_allocator_size_in_use_bytes\":{},\"baseline_allocator_size_allocated_bytes\":{}",
        baseline.size_in_use, baseline.size_allocated
    );
    #[cfg(not(target_os = "macos"))]
    let baseline_json = "\"allocator_stats_supported\":false".to_string();
    println!(
        "{{\"fixture\":{},\"iterations\":{},\"wall_ms\":{:.3},{},\"observations\":[{}]}}",
        json_string(&fixture),
        iterations,
        started.elapsed().as_secs_f64() * 1000.0,
        baseline_json,
        observations.join(",")
    );
}

fn usage(message: &str) -> ! {
    eprintln!("{message}\nusage: measure_reader_write_inprocess <fixture.xlsx> <iterations>");
    std::process::exit(2);
}
