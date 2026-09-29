//! Measurement-only benchmark for the native VBA VM execution boundary.
//!
//! Parsing, workbook I/O, and JSON rendering are intentionally outside the
//! timed region. This binary is not a cross-library or Excel benchmark.

use elixcee::{parser, vm::Vm};
use std::time::Instant;

const SOURCE_TEMPLATE: &str =
    "Sub Main()\n    For i = 1 To {rows}\n        Cells(i, 1).Value = i * i\n    Next i\nEnd Sub\n";

fn percentile(sorted: &[f64], numerator: usize, denominator: usize) -> f64 {
    let index = ((sorted.len().saturating_sub(1) * numerator) + denominator / 2) / denominator;
    sorted[index]
}

fn parse_arg<T: std::str::FromStr>(args: &[String], name: &str, default: T) -> T {
    args.windows(2)
        .find(|pair| pair[0] == name)
        .and_then(|pair| pair[1].parse().ok())
        .unwrap_or(default)
}

fn json_number(value: f64) -> String {
    format!("{value:.6}")
}

fn main() {
    let args: Vec<String> = std::env::args().collect();
    let rows: u32 = parse_arg(&args, "--rows", 1_000);
    let warmup: usize = parse_arg(&args, "--warmup", 5);
    let repetitions: usize = parse_arg(&args, "--repetitions", 30);
    assert!(rows > 0, "--rows must be greater than zero");
    assert!(repetitions > 0, "--repetitions must be greater than zero");
    assert!(repetitions <= 10_000, "--repetitions is too large");

    let source = SOURCE_TEMPLATE.replace("{rows}", &rows.to_string());
    let program = parser::parse(&source).expect("benchmark source must parse");

    for _ in 0..warmup {
        let mut vm = Vm::new();
        vm.run_sub(&program, "Main")
            .expect("benchmark warmup must execute");
    }

    let mut samples = Vec::with_capacity(repetitions);
    for _ in 0..repetitions {
        let mut vm = Vm::new();
        let start = Instant::now();
        vm.run_sub(&program, "Main")
            .expect("benchmark sample must execute");
        let elapsed_ms = start.elapsed().as_secs_f64() * 1_000.0;
        assert_eq!(
            vm.get_cell(rows, 1),
            elixcee::types::Variant::Integer(i64::from(rows) * i64::from(rows))
        );
        samples.push(elapsed_ms);
    }

    let mut sorted = samples.clone();
    sorted.sort_by(f64::total_cmp);
    let p50 = percentile(&sorted, 1, 2);
    let p95 = percentile(&sorted, 19, 20);
    let sample_json = samples
        .iter()
        .map(|value| json_number(*value))
        .collect::<Vec<_>>()
        .join(",");
    println!(
        "{{\"schema\":\"elixcee.vba-vm-benchmark.v1\",\"version\":\"{}\",\"scope\":\"parse-excluded-vm-run-only\",\"rows\":{},\"warmup\":{},\"repetitions\":{},\"p50_ms\":{},\"p95_ms\":{},\"samples_ms\":[{}]}}",
        env!("CARGO_PKG_VERSION"),
        rows,
        warmup,
        repetitions,
        json_number(p50),
        json_number(p95),
        sample_json
    );
}
