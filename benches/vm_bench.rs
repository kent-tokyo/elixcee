use criterion::{Criterion, black_box, criterion_group, criterion_main};
use elixcee::parser::parse;
use elixcee::vm::{CellContent, Variant, Vm};
use elixcee::formula::eval::evaluate;
use elixcee::formula::parser::parse as fparse;
use std::collections::HashMap;

// ── VBA macro benchmarks ───────────────────────────────────────────────────────

fn bench_vba_loop_1000(c: &mut Criterion) {
    let src = r#"
Sub FillSquares()
    For i = 1 To 1000
        Cells(i, 1).Value = i * i
    Next i
End Sub
"#;
    let prog = parse(src).unwrap();
    c.bench_function("vba_loop_1000", |b| {
        b.iter(|| {
            let mut vm = Vm::new();
            vm.run_sub(black_box(&prog), "FillSquares").unwrap();
        })
    });
}

fn bench_vba_loop_10000(c: &mut Criterion) {
    let src = r#"
Sub FillSquares()
    For i = 1 To 10000
        Cells(i, 1).Value = i * i
    Next i
End Sub
"#;
    let prog = parse(src).unwrap();
    c.bench_function("vba_loop_10000", |b| {
        b.iter(|| {
            let mut vm = Vm::new();
            vm.run_sub(black_box(&prog), "FillSquares").unwrap();
        })
    });
}

fn bench_vba_if_branch(c: &mut Criterion) {
    let src = r#"
Sub Classify()
    For i = 1 To 1000
        If Cells(i, 1).Value > 500 Then
            Cells(i, 2).Value = "high"
        ElseIf Cells(i, 1).Value > 100 Then
            Cells(i, 2).Value = "mid"
        Else
            Cells(i, 2).Value = "low"
        End If
    Next i
End Sub
"#;
    let prog = parse(src).unwrap();
    c.bench_function("vba_if_branch_1000", |b| {
        b.iter(|| {
            let mut vm = Vm::new();
            // pre-fill column A
            for row in 1u32..=1000 {
                let cells = vm.cells_mut();
                cells.insert((row, 1), CellContent { formula: None, value: Variant::Integer(row as i64) });
            }
            vm.run_sub(black_box(&prog), "Classify").unwrap();
        })
    });
}

fn bench_vba_parse_only(c: &mut Criterion) {
    let src = r#"
Sub FillSquares()
    For i = 1 To 1000
        Cells(i, 1).Value = i * i
    Next i
End Sub
"#;
    c.bench_function("vba_parse_only", |b| {
        b.iter(|| {
            parse(black_box(src)).unwrap()
        })
    });
}

// ── Cell write benchmarks ──────────────────────────────────────────────────────

fn bench_set_cell_10000(c: &mut Criterion) {
    c.bench_function("set_cell_direct_10000", |b| {
        b.iter(|| {
            let mut vm = Vm::new();
            let cells = vm.cells_mut();
            for row in 1u32..=10000 {
                cells.insert(
                    (row, 1),
                    CellContent { formula: None, value: Variant::Integer(black_box(row as i64)) },
                );
            }
        })
    });
}

// ── Formula evaluation benchmarks ─────────────────────────────────────────────

fn bench_formula_sum(c: &mut Criterion) {
    // SUM(A1:A1000) over 1000 cells
    let mut cells: HashMap<(u32, u32), CellContent> = HashMap::new();
    for row in 1u32..=1000 {
        cells.insert((row, 1), CellContent { formula: None, value: Variant::Integer(row as i64) });
    }
    let expr = fparse("=SUM(A1:A1000)").unwrap();
    c.bench_function("formula_sum_1000", |b| {
        b.iter(|| evaluate(black_box(&expr), black_box(&cells)).unwrap())
    });
}

fn bench_formula_sumif(c: &mut Criterion) {
    // SUMIF(A1:A1000, ">500", B1:B1000)
    let mut cells: HashMap<(u32, u32), CellContent> = HashMap::new();
    for row in 1u32..=1000 {
        cells.insert((row, 1), CellContent { formula: None, value: Variant::Integer(row as i64) });
        cells.insert((row, 2), CellContent { formula: None, value: Variant::Integer((row * 2) as i64) });
    }
    let expr = fparse("=SUMIF(A1:A1000,\">500\",B1:B1000)").unwrap();
    c.bench_function("formula_sumif_1000", |b| {
        b.iter(|| evaluate(black_box(&expr), black_box(&cells)).unwrap())
    });
}

fn bench_formula_vlookup(c: &mut Criterion) {
    // VLOOKUP on a 1000-row table
    let mut cells: HashMap<(u32, u32), CellContent> = HashMap::new();
    for row in 1u32..=1000 {
        cells.insert((row, 1), CellContent { formula: None, value: Variant::Integer(row as i64) });
        cells.insert((row, 2), CellContent { formula: None, value: Variant::Integer((row * 10) as i64) });
    }
    let expr = fparse("=VLOOKUP(750,A1:B1000,2,FALSE)").unwrap();
    c.bench_function("formula_vlookup_1000", |b| {
        b.iter(|| evaluate(black_box(&expr), black_box(&cells)).unwrap())
    });
}

fn bench_formula_dsum(c: &mut Criterion) {
    // DSUM over 1000-row database
    // Col A = category (even/odd), Col B = value
    // Criteria: category = "even"
    let mut cells: HashMap<(u32, u32), CellContent> = HashMap::new();
    cells.insert((1, 1), CellContent { formula: None, value: Variant::Str("Category".into()) });
    cells.insert((1, 2), CellContent { formula: None, value: Variant::Str("Value".into()) });
    for row in 2u32..=1001 {
        let cat = if row % 2 == 0 { "even" } else { "odd" };
        cells.insert((row, 1), CellContent { formula: None, value: Variant::Str(cat.into()) });
        cells.insert((row, 2), CellContent { formula: None, value: Variant::Integer(row as i64) });
    }
    // criteria at D1:D2
    cells.insert((1, 4), CellContent { formula: None, value: Variant::Str("Category".into()) });
    cells.insert((2, 4), CellContent { formula: None, value: Variant::Str("even".into()) });
    let expr = fparse("=DSUM(A1:B1001,\"Value\",D1:D2)").unwrap();
    c.bench_function("formula_dsum_1000", |b| {
        b.iter(|| evaluate(black_box(&expr), black_box(&cells)).unwrap())
    });
}

fn bench_formula_filter(c: &mut Criterion) {
    // FILTER(A1:A1000, B1:B1000>500)
    let mut cells: HashMap<(u32, u32), CellContent> = HashMap::new();
    for row in 1u32..=1000 {
        cells.insert((row, 1), CellContent { formula: None, value: Variant::Integer(row as i64) });
        cells.insert((row, 2), CellContent { formula: None, value: Variant::Integer(row as i64) });
    }
    let expr = fparse("=FILTER(A1:A1000,B1:B1000>500)").unwrap();
    c.bench_function("formula_filter_1000", |b| {
        b.iter(|| evaluate(black_box(&expr), black_box(&cells)).unwrap())
    });
}

fn bench_recalculate(c: &mut Criterion) {
    // recalculate_all on 100 SUM formulas referencing a shared range
    c.bench_function("recalculate_100_formulas", |b| {
        b.iter(|| {
            let mut vm = Vm::new();
            for row in 1u32..=100 {
                let cells = vm.cells_mut();
                cells.insert((row, 1), CellContent { formula: None, value: Variant::Integer(row as i64) });
            }
            for row in 1u32..=100 {
                vm.set_cell_formula(row, 2, &format!("=SUM(A1:A{})", row)).unwrap();
            }
        })
    });
}

criterion_group!(
    benches,
    bench_vba_parse_only,
    bench_vba_loop_1000,
    bench_vba_loop_10000,
    bench_vba_if_branch,
    bench_set_cell_10000,
    bench_formula_sum,
    bench_formula_sumif,
    bench_formula_vlookup,
    bench_formula_dsum,
    bench_formula_filter,
    bench_recalculate,
);
criterion_main!(benches);
