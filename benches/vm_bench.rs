use criterion::{Criterion, black_box, criterion_group, criterion_main};
use elixcee::formula::eval::evaluate;
use elixcee::formula::parser::parse as fparse;
use elixcee::parser::parse;
use elixcee::vm::{CellContent, Variant, Vm};
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
                cells.insert(
                    (row, 1),
                    CellContent {
                        formula: None,
                        value: Variant::Integer(row as i64),
                    },
                );
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
        b.iter(|| parse(black_box(src)).unwrap())
    });
}

fn bench_vba_user_function_calls(c: &mut Criterion) {
    let src = r#"
Sub Main()
    For i = 1 To 1000
        result = AddOne(i)
    Next i
End Sub
Function AddOne(value)
    AddOne = value + 1
End Function
"#;
    let prog = parse(src).unwrap();
    c.bench_function("vba_user_function_calls_1000", |b| {
        b.iter(|| {
            let mut vm = Vm::new();
            vm.run_sub(black_box(&prog), "Main").unwrap();
        })
    });
}

fn bench_variable_budget_large_array(c: &mut Criterion) {
    let prog =
        parse("Sub Work()\n    For i = 1 To 500\n        x = i\n    Next i\nEnd Sub\n").unwrap();
    c.bench_function("variable_budget_large_array_20000x500", |b| {
        b.iter_batched(
            || {
                let mut vm = Vm::new();
                vm.variables.insert(
                    "payload".into(),
                    Variant::Array(vec![Variant::Integer(0); 20_000]),
                );
                vm
            },
            |mut vm| vm.run_sub(black_box(&prog), "Work").unwrap(),
            criterion::BatchSize::SmallInput,
        )
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
                    CellContent {
                        formula: None,
                        value: Variant::Integer(black_box(row as i64)),
                    },
                );
            }
        })
    });
}

fn reference_append_row(vm: &mut Vm, value: Variant) {
    let target_row = vm
        .get_sheet_cells("sheet1")
        .expect("default sheet exists")
        .iter()
        .filter(|(_, content)| !matches!(&content.value, Variant::Empty))
        .map(|(&(row, _), _)| row)
        .max()
        .map_or(1, |row| row + 1);
    vm.write_rect("sheet1", (target_row, 1), &[vec![value]]);
}

fn bench_append_row_5000(c: &mut Criterion) {
    c.bench_function("append_row_reference_rescan_5000", |b| {
        b.iter(|| {
            let mut vm = Vm::new();
            for row in 1..=5000 {
                reference_append_row(&mut vm, Variant::Integer(black_box(row)));
            }
        })
    });
    c.bench_function("append_row_cached_5000", |b| {
        b.iter(|| {
            let mut vm = Vm::new();
            for row in 1..=5000 {
                vm.append_row_values("sheet1", vec![Variant::Integer(black_box(row))]);
            }
        })
    });
}

fn reference_read_rect(vm: &Vm, r1: u32, c1: u32, r2: u32, c2: u32) -> Vec<Vec<Variant>> {
    let empty = HashMap::new();
    let cells = vm.get_sheet_cells("sheet1").unwrap_or(&empty);
    (r1..=r2)
        .map(|row| {
            (c1..=c2)
                .map(|col| {
                    cells
                        .get(&(row, col))
                        .map(|content| content.value.clone())
                        .unwrap_or(Variant::Empty)
                })
                .collect()
        })
        .collect()
}

fn bench_read_rect_128(c: &mut Criterion) {
    let values: Vec<Vec<Variant>> = (1..=128)
        .map(|row| {
            (1..=128)
                .map(|col| Variant::Integer((row * 1000 + col) as i64))
                .collect()
        })
        .collect();
    let mut vm = Vm::new();
    vm.write_rect("sheet1", (1, 1), &values);
    vm.read_rect("sheet1", 1, 1, 128, 128);
    c.bench_function("read_rect_reference_dense_128", |b| {
        b.iter(|| black_box(reference_read_rect(black_box(&vm), 1, 1, 128, 128)))
    });
    c.bench_function("read_rect_tiled_dense_128", |b| {
        b.iter(|| black_box(vm.read_rect("sheet1", 1, 1, 128, 128)))
    });
}

fn bench_read_rect_tile_cache_lru_257(c: &mut Criterion) {
    c.bench_function("read_rect_tile_cache_lru_257", |b| {
        b.iter(|| {
            let vm = Vm::new();
            for tile_col in 0..257u32 {
                let col = tile_col * 32 + 1;
                black_box(vm.read_rect("sheet1", 1, col, 32, col + 31));
            }
        })
    });
}

fn bench_read_rect_tile_density(c: &mut Criterion) {
    for populated in [1usize, 32, 128, 129, 1024] {
        let mut vm = Vm::new();
        for position in 0..populated {
            vm.cells_mut().insert(
                (position as u32 / 32 + 1, position as u32 % 32 + 1),
                CellContent {
                    formula: None,
                    value: Variant::Integer(position as i64),
                },
            );
        }
        vm.read_rect("sheet1", 1, 1, 32, 32);
        let name = format!("read_rect_tile_density_{populated}");
        c.bench_function(&name, |b| {
            b.iter(|| black_box(vm.read_rect("sheet1", 1, 1, 32, 32)))
        });
    }
}

fn make_tile_density_vm(populated: usize) -> Vm {
    let mut vm = Vm::new();
    for position in 0..populated {
        vm.cells_mut().insert(
            (position as u32 / 32 + 1, position as u32 % 32 + 1),
            CellContent {
                formula: None,
                value: Variant::Integer(position as i64),
            },
        );
    }
    vm
}

fn bench_read_rect_tile_build_vs_warm(c: &mut Criterion) {
    for populated in [1usize, 1024] {
        let build_name = format!("read_rect_tile_build_{populated}");
        c.bench_function(&build_name, |b| {
            b.iter(|| {
                let vm = make_tile_density_vm(populated);
                black_box(vm.read_rect("sheet1", 1, 1, 32, 32));
            })
        });
        let vm = make_tile_density_vm(populated);
        vm.read_rect("sheet1", 1, 1, 32, 32);
        let warm_name = format!("read_rect_tile_warm_{populated}");
        c.bench_function(&warm_name, |b| {
            b.iter(|| black_box(vm.read_rect("sheet1", 1, 1, 32, 32)))
        });
    }
}

fn bench_incremental_write_cached_tile_100(c: &mut Criterion) {
    c.bench_function("write_rect_incremental_cached_tile_100", |b| {
        b.iter(|| {
            let mut vm = make_tile_density_vm(1024);
            vm.read_rect("sheet1", 1, 1, 32, 32);
            for offset in 0..100u32 {
                vm.write_rect(
                    "sheet1",
                    (offset % 32 + 1, offset / 32 + 1),
                    &[vec![Variant::Integer(offset as i64)]],
                );
                black_box(vm.read_rect("sheet1", 1, 1, 32, 32));
            }
        })
    });
    c.bench_function("write_rect_invalidate_rebuild_tile_100", |b| {
        b.iter(|| {
            let mut vm = make_tile_density_vm(1024);
            vm.read_rect("sheet1", 1, 1, 32, 32);
            for offset in 0..100u32 {
                vm.cells_mut().insert(
                    (offset % 32 + 1, offset / 32 + 1),
                    CellContent {
                        formula: None,
                        value: Variant::Integer(offset as i64),
                    },
                );
                black_box(vm.read_rect("sheet1", 1, 1, 32, 32));
            }
        })
    });
}

fn bench_write_rect_cached_tiles_128(c: &mut Criterion) {
    let values = vec![vec![Variant::Integer(7); 128]; 128];
    c.bench_function("write_rect_adaptive_cached_tiles_128", |b| {
        b.iter(|| {
            let mut vm = make_tile_density_vm(16_384);
            black_box(vm.read_rect("sheet1", 1, 1, 128, 128));
            vm.write_rect("sheet1", (1, 1), &values);
            black_box(vm.read_rect("sheet1", 1, 1, 128, 128));
        })
    });
    c.bench_function("write_rect_invalidate_rebuild_tiles_128", |b| {
        b.iter(|| {
            let mut vm = make_tile_density_vm(16_384);
            black_box(vm.read_rect("sheet1", 1, 1, 128, 128));
            {
                let cells = vm.cells_mut();
                for (row_offset, row) in values.iter().enumerate() {
                    for (col_offset, value) in row.iter().enumerate() {
                        cells.insert(
                            (row_offset as u32 + 1, col_offset as u32 + 1),
                            CellContent {
                                formula: None,
                                value: value.clone(),
                            },
                        );
                    }
                }
            }
            black_box(vm.read_rect("sheet1", 1, 1, 128, 128));
        })
    });
}

fn bench_write_rect_threshold_calibration(c: &mut Criterion) {
    for (name, height, width) in [
        ("256", 16usize, 16usize),
        ("512", 32, 16),
        ("528", 33, 16),
        ("1024", 32, 32),
    ] {
        let values = vec![vec![Variant::Integer(7); width]; height];
        let adaptive_name = format!("write_rect_adaptive_threshold_{name}");
        c.bench_function(&adaptive_name, |b| {
            b.iter(|| {
                let mut vm = make_tile_density_vm(1024);
                black_box(vm.read_rect("sheet1", 1, 1, 32, 32));
                vm.write_rect("sheet1", (1, 1), &values);
                black_box(vm.read_rect("sheet1", 1, 1, 32, 32));
            })
        });
        let reference_name = format!("write_rect_reference_threshold_{name}");
        c.bench_function(&reference_name, |b| {
            b.iter(|| {
                let mut vm = make_tile_density_vm(1024);
                black_box(vm.read_rect("sheet1", 1, 1, 32, 32));
                {
                    let cells = vm.cells_mut();
                    for (row_offset, row) in values.iter().enumerate() {
                        for (col_offset, value) in row.iter().enumerate() {
                            cells.insert(
                                (row_offset as u32 + 1, col_offset as u32 + 1),
                                CellContent {
                                    formula: None,
                                    value: value.clone(),
                                },
                            );
                        }
                    }
                }
                black_box(vm.read_rect("sheet1", 1, 1, 32, 32));
            })
        });
    }
}

// ── Formula evaluation benchmarks ─────────────────────────────────────────────

fn bench_formula_sum(c: &mut Criterion) {
    // SUM(A1:A1000) over 1000 cells
    let mut cells: HashMap<(u32, u32), CellContent> = HashMap::new();
    for row in 1u32..=1000 {
        cells.insert(
            (row, 1),
            CellContent {
                formula: None,
                value: Variant::Integer(row as i64),
            },
        );
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
        cells.insert(
            (row, 1),
            CellContent {
                formula: None,
                value: Variant::Integer(row as i64),
            },
        );
        cells.insert(
            (row, 2),
            CellContent {
                formula: None,
                value: Variant::Integer((row * 2) as i64),
            },
        );
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
        cells.insert(
            (row, 1),
            CellContent {
                formula: None,
                value: Variant::Integer(row as i64),
            },
        );
        cells.insert(
            (row, 2),
            CellContent {
                formula: None,
                value: Variant::Integer((row * 10) as i64),
            },
        );
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
    cells.insert(
        (1, 1),
        CellContent {
            formula: None,
            value: Variant::Str("Category".into()),
        },
    );
    cells.insert(
        (1, 2),
        CellContent {
            formula: None,
            value: Variant::Str("Value".into()),
        },
    );
    for row in 2u32..=1001 {
        let cat = if row % 2 == 0 { "even" } else { "odd" };
        cells.insert(
            (row, 1),
            CellContent {
                formula: None,
                value: Variant::Str(cat.into()),
            },
        );
        cells.insert(
            (row, 2),
            CellContent {
                formula: None,
                value: Variant::Integer(row as i64),
            },
        );
    }
    // criteria at D1:D2
    cells.insert(
        (1, 4),
        CellContent {
            formula: None,
            value: Variant::Str("Category".into()),
        },
    );
    cells.insert(
        (2, 4),
        CellContent {
            formula: None,
            value: Variant::Str("even".into()),
        },
    );
    let expr = fparse("=DSUM(A1:B1001,\"Value\",D1:D2)").unwrap();
    c.bench_function("formula_dsum_1000", |b| {
        b.iter(|| evaluate(black_box(&expr), black_box(&cells)).unwrap())
    });
}

fn bench_formula_filter(c: &mut Criterion) {
    // FILTER(A1:A1000, B1:B1000>500)
    let mut cells: HashMap<(u32, u32), CellContent> = HashMap::new();
    for row in 1u32..=1000 {
        cells.insert(
            (row, 1),
            CellContent {
                formula: None,
                value: Variant::Integer(row as i64),
            },
        );
        cells.insert(
            (row, 2),
            CellContent {
                formula: None,
                value: Variant::Integer(row as i64),
            },
        );
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
                cells.insert(
                    (row, 1),
                    CellContent {
                        formula: None,
                        value: Variant::Integer(row as i64),
                    },
                );
            }
            for row in 1u32..=100 {
                vm.set_cell_formula(row, 2, &format!("=SUM(A1:A{})", row))
                    .unwrap();
            }
        })
    });
}

fn bench_recalculate_large_range_dependencies(c: &mut Criterion) {
    let mut template = Vm::new();
    for row in 1u32..=50 {
        template
            .set_cell_formula(row, 2, "=IF(FALSE,SUM(A1:A100000),0)")
            .unwrap();
    }
    c.bench_function("recalculate_50_large_range_dependencies", |b| {
        b.iter_batched(
            || template.clone(),
            |mut vm| vm.recalculate_all().unwrap(),
            criterion::BatchSize::SmallInput,
        )
    });
}

fn bench_recalculate_many_formula_parses(c: &mut Criterion) {
    let mut template = Vm::new();
    for row in 1u32..=10_000 {
        template.set_cell_formula(row, 1, "=1+1").unwrap();
    }
    c.bench_function("recalculate_10000_formula_parses", |b| {
        b.iter_batched(
            || template.clone(),
            |mut vm| vm.recalculate_all().unwrap(),
            criterion::BatchSize::SmallInput,
        )
    });
}

fn make_dirty_propagation_template(formula_count: u32) -> Vm {
    let mut vm = Vm::new();
    vm.write_rect("sheet1", (1, 1), &[vec![Variant::Integer(1)]]);
    for col in 2..=formula_count + 1 {
        let previous = col - 1;
        vm.set_cell_formula(1, col, &format!("={}1+1", xlsx_col(previous)))
            .unwrap();
    }
    vm.recalculate_all().unwrap();
    vm
}

fn xlsx_col(mut col: u32) -> String {
    let mut out = String::new();
    while col > 0 {
        let rem = (col - 1) % 26;
        out.push((b'A' + rem as u8) as char);
        col = (col - 1) / 26;
    }
    out.chars().rev().collect()
}

fn bench_recalculate_dirty_propagation(c: &mut Criterion) {
    let template = make_dirty_propagation_template(1_000);
    c.bench_function("recalculate_dirty_chain_1000_single_input", |b| {
        b.iter_batched(
            || {
                let mut vm = template.clone();
                vm.write_rect("sheet1", (1, 1), &[vec![Variant::Integer(2)]]);
                vm
            },
            |mut vm| vm.recalculate_all().unwrap(),
            criterion::BatchSize::SmallInput,
        )
    });
    c.bench_function("recalculate_dirty_chain_1000_noop", |b| {
        b.iter_batched(
            || template.clone(),
            |mut vm| vm.recalculate_all().unwrap(),
            criterion::BatchSize::SmallInput,
        )
    });
    c.bench_function("recalculate_full_chain_1000_structure_rebuild", |b| {
        b.iter_batched(
            || {
                let mut vm = template.clone();
                // Replacing a formula invalidates the persistent plan and
                // represents the full-rebuild comparison case.
                vm.set_cell_formula(1, 1_001, "=ALL1+1").unwrap();
                vm
            },
            |mut vm| vm.recalculate_all().unwrap(),
            criterion::BatchSize::SmallInput,
        )
    });
    let mut independent = Vm::new();
    for row in 1..=1_000u32 {
        independent.write_rect("sheet1", (row, 1), &[vec![Variant::Integer(row as i64)]]);
        independent
            .set_cell_formula(row, 2, &format!("=A{row}+1"))
            .unwrap();
    }
    independent.recalculate_all().unwrap();
    c.bench_function("recalculate_dirty_independent_1000_all_inputs", |b| {
        b.iter_batched(
            || {
                let mut vm = independent.clone();
                let values = (1..=1_000u32)
                    .map(|row| vec![Variant::Integer((row + 1) as i64)])
                    .collect::<Vec<_>>();
                vm.write_rect("sheet1", (1, 1), &values);
                vm
            },
            |mut vm| vm.recalculate_all().unwrap(),
            criterion::BatchSize::SmallInput,
        )
    });
}

fn bench_formula_countif_wildcard(c: &mut Criterion) {
    // COUNTIF(A1:A1000,"*son") — suffix wildcard, no DP needed with fast path
    let mut cells: HashMap<(u32, u32), CellContent> = HashMap::new();
    let names = [
        "Jackson", "Mason", "Wilson", "Johnson", "Anderson", "Thompson", "Martin", "Cooper",
        "Evans", "Murphy",
    ];
    for row in 1u32..=1000 {
        let name = names[(row as usize - 1) % names.len()];
        cells.insert(
            (row, 1),
            CellContent {
                formula: None,
                value: Variant::Str(name.to_uppercase()),
            },
        );
    }
    let expr_suffix = fparse("=COUNTIF(A1:A1000,\"*SON\")").unwrap();
    let expr_contains = fparse("=COUNTIF(A1:A1000,\"*SON*\")").unwrap();
    c.bench_function("formula_countif_wildcard_suffix", |b| {
        b.iter(|| evaluate(black_box(&expr_suffix), black_box(&cells)).unwrap())
    });
    c.bench_function("formula_countif_wildcard_contains", |b| {
        b.iter(|| evaluate(black_box(&expr_contains), black_box(&cells)).unwrap())
    });
}

criterion_group!(
    benches,
    bench_vba_parse_only,
    bench_vba_user_function_calls,
    bench_variable_budget_large_array,
    bench_vba_loop_1000,
    bench_vba_loop_10000,
    bench_vba_if_branch,
    bench_set_cell_10000,
    bench_append_row_5000,
    bench_read_rect_128,
    bench_read_rect_tile_cache_lru_257,
    bench_read_rect_tile_density,
    bench_read_rect_tile_build_vs_warm,
    bench_incremental_write_cached_tile_100,
    bench_write_rect_cached_tiles_128,
    bench_write_rect_threshold_calibration,
    bench_formula_sum,
    bench_formula_sumif,
    bench_formula_vlookup,
    bench_formula_dsum,
    bench_formula_filter,
    bench_recalculate,
    bench_recalculate_large_range_dependencies,
    bench_recalculate_many_formula_parses,
    bench_recalculate_dirty_propagation,
    bench_formula_countif_wildcard,
);
criterion_main!(benches);
