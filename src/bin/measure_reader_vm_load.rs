//! Measurement-only whole reader-to-VM load timing.

use std::hint::black_box;
use std::time::Instant;

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

    let started = Instant::now();
    let mut observations = Vec::with_capacity(iterations);
    for iteration in 1..=iterations {
        let iteration_started = Instant::now();
        let mut vm = elixcee::vm::Vm::new();
        let names = vm.load_workbook_file(&fixture).unwrap_or_else(|error| {
            eprintln!("iteration {iteration} failed: {error}");
            std::process::exit(1);
        });
        let cells = names
            .iter()
            .map(|name| {
                vm.set_active_sheet(name).expect("loaded sheet must exist");
                vm.cells().len()
            })
            .sum::<usize>();
        black_box(&vm);
        observations.push(format!(
            "{{\"iteration\":{},\"sheets\":{},\"cells\":{},\"wall_ms\":{:.3}}}",
            iteration,
            names.len(),
            cells,
            iteration_started.elapsed().as_secs_f64() * 1000.0
        ));
    }

    println!(
        "{{\"fixture\":{},\"iterations\":{},\"wall_ms\":{:.3},\"observations\":[{}]}}",
        json_string(&fixture),
        iterations,
        started.elapsed().as_secs_f64() * 1000.0,
        observations.join(",")
    );
}

fn usage(message: &str) -> ! {
    eprintln!("{message}\nusage: measure_reader_vm_load <fixture.xlsx> <iterations>");
    std::process::exit(2);
}
