//! Local benchmark worker; timings exclude startup and validation of untouched cells.
use elixcee::{reader, save_workbook, types::Variant, vm::Vm};
use std::time::Instant;

fn value(cell: &reader::SheetCell) -> Variant {
    match cell {
        reader::SheetCell::Integer(n) => Variant::Integer(*n),
        reader::SheetCell::Float(n) => Variant::Float(*n),
        reader::SheetCell::Str(s) => Variant::Str(s.clone()),
        reader::SheetCell::Bool(b) => Variant::Boolean(*b),
        reader::SheetCell::Error(e) => Variant::Error(e.clone()),
    }
}

fn main() {
    let args: Vec<String> = std::env::args().collect();
    assert_eq!(args.len(), 4, "fixture output iterations");
    let iterations: usize = args[3].parse().unwrap();
    assert!((1..=1000).contains(&iterations));
    let original = reader::read_workbook(&args[1]).unwrap();
    assert!(
        !original.is_empty(),
        "benchmark requires at least one sheet"
    );
    let mut samples = Vec::new();
    for _ in 0..iterations {
        let start = Instant::now();
        let mut vm = Vm::new();
        vm.load_workbook_file(&args[1]).unwrap();
        let sheet = vm.sheet_names()[0].clone();
        vm.set_active_sheet(&sheet).unwrap();
        let loaded = start.elapsed().as_secs_f64() * 1000.0;
        vm.write_rect(&sheet, (1, 1), &[vec![Variant::Integer(123)]]);
        vm.set_cell_formula(1, 2, "=1+2").unwrap();
        let mutated = start.elapsed().as_secs_f64() * 1000.0;
        save_workbook(&vm, &args[2]).unwrap();
        let saved = start.elapsed().as_secs_f64() * 1000.0;
        let result = reader::read_workbook(&args[2]).unwrap();
        let elapsed = start.elapsed().as_secs_f64() * 1000.0;
        assert_eq!(result.len(), original.len());
        assert!(matches!(
            result[0].cells.get(&(1, 1)),
            Some(reader::SheetCell::Integer(123))
        ));
        assert_eq!(
            result[0]
                .formulas
                .get(&(1, 2))
                .unwrap()
                .trim_start_matches('='),
            "1+2"
        );
        for (sheet_index, (before, after)) in original.iter().zip(&result).enumerate() {
            assert_eq!(before.name, after.name);
            for (position, original_value) in &before.cells {
                if sheet_index != 0 || (*position != (1, 1) && *position != (1, 2)) {
                    assert_eq!(
                        value(after.cells.get(position).unwrap()),
                        value(original_value)
                    );
                }
            }
            for (position, formula) in &before.formulas {
                if sheet_index != 0 || (*position != (1, 1) && *position != (1, 2)) {
                    assert_eq!(after.formulas.get(position), Some(formula));
                }
            }
        }
        samples.push(format!("{{\"load_ms\":{loaded},\"mutate_ms\":{},\"save_ms\":{},\"reload_ms\":{},\"total_ms\":{elapsed}}}", mutated-loaded, saved-mutated, elapsed-saved));
    }
    println!(
        "{{\"version\":\"{}\",\"samples\":[{}]}}",
        env!("CARGO_PKG_VERSION"),
        samples.join(",")
    );
}
