/// Property-based tests for elixcee using proptest.
///
/// Each test verifies an invariant that must hold for *all* inputs in the
/// generated domain, not just the hand-picked cases in unit tests.
use elixcee::{
    formula,
    parser,
    vm::{CellContent, Variant, Vm},
};
use proptest::prelude::*;
use std::collections::HashMap;

// ── VBA Parser ────────────────────────────────────────────────────────────────

proptest! {
    /// parse() must never panic — must always return Ok or Err.
    #[test]
    fn prop_vba_parse_never_panics(s in "(?s).{0,200}") {
        let _ = parser::parse(&s);
    }

    /// parse() on a valid Sub/End Sub skeleton must not panic.
    #[test]
    fn prop_vba_parse_valid_sub(body in "[a-zA-Z0-9 =+\\-*/<>_()\n\t',.]{0,200}") {
        let code = format!("Sub Test()\n{}\nEnd Sub\n", body);
        let _ = parser::parse(&code);
    }

    /// Well-formed assignment statements must always parse successfully.
    #[test]
    fn prop_vba_assignment_parses(
        var in "[a-z][a-z0-9]{0,8}",
        val in -1_000_000i64..=1_000_000i64,
    ) {
        let code = format!("Sub MySub()\n    {} = {}\nEnd Sub\n", var, val);
        prop_assert!(parser::parse(&code).is_ok(), "expected Ok for: {}", code);
    }
}

// ── Formula Parser & Evaluator ────────────────────────────────────────────────

proptest! {
    /// formula::parse() must never panic.
    #[test]
    fn prop_formula_parse_never_panics(s in "(?s).{0,100}") {
        let _ = formula::parse(&s);
    }

    /// formula::evaluate() on a successfully parsed expression must never panic.
    #[test]
    fn prop_formula_eval_never_panics(
        s in "[A-Z0-9+\\-*/(),.\"' !=<>:A-Za-z]{0,80}"
    ) {
        if let Ok(expr) = formula::parse(&s) {
            let cells: HashMap<(u32, u32), CellContent> = HashMap::new();
            let _ = formula::evaluate(&expr, &cells);
        }
    }

    /// Numeric literals always evaluate to the expected value.
    #[test]
    fn prop_formula_numeric_literal(n in -1_000_000i64..=1_000_000i64) {
        let s = format!("={}", n);
        let cells: HashMap<(u32, u32), CellContent> = HashMap::new();
        if let Ok(expr) = formula::parse(&s) {
            match formula::evaluate(&expr, &cells) {
                Ok(Variant::Integer(v)) => prop_assert_eq!(v, n),
                Ok(Variant::Float(f))   => prop_assert!((f - n as f64).abs() < 1e-9),
                other => prop_assert!(false, "unexpected: {:?}", other),
            }
        }
    }

    /// SUM of a range equals the arithmetic sum of the individual values.
    #[test]
    fn prop_sum_matches_manual(
        values in proptest::collection::vec(-1000i64..=1000i64, 1..=10)
    ) {
        let mut cells: HashMap<(u32, u32), CellContent> = HashMap::new();
        let n = values.len() as u32;
        for (i, &v) in values.iter().enumerate() {
            cells.insert(
                (i as u32 + 1, 1),
                CellContent { formula: None, value: Variant::Integer(v) },
            );
        }
        let formula_str = format!("=SUM(A1:A{})", n);
        if let Ok(expr) = formula::parse(&formula_str) {
            if let Ok(result) = formula::evaluate(&expr, &cells) {
                let expected: i64 = values.iter().sum();
                match result {
                    Variant::Integer(v) => prop_assert_eq!(v, expected),
                    Variant::Float(f)   => prop_assert!((f - expected as f64).abs() < 1e-6),
                    other => prop_assert!(false, "unexpected SUM result: {:?}", other),
                }
            }
        }
    }

    /// COUNTIF with wildcard patterns must never panic (no exponential recursion).
    #[test]
    fn prop_countif_wildcard_never_panics(
        pattern in "[*?a-z]{0,20}",
        values  in proptest::collection::vec("[a-z]{0,10}", 1..=5),
    ) {
        let mut cells: HashMap<(u32, u32), CellContent> = HashMap::new();
        let n = values.len() as u32;
        for (i, v) in values.iter().enumerate() {
            cells.insert(
                (i as u32 + 1, 1),
                CellContent { formula: None, value: Variant::Str(v.clone()) },
            );
        }
        let formula_str = format!("=COUNTIF(A1:A{},\"{}\")", n, pattern);
        if let Ok(expr) = formula::parse(&formula_str) {
            let _ = formula::evaluate(&expr, &cells);
        }
    }
}

// ── VM / Cell Storage ─────────────────────────────────────────────────────────

proptest! {
    /// set_cell then get_cell roundtrip for integer values.
    #[test]
    fn prop_cell_integer_roundtrip(
        row in 1u32..=10_000u32,
        col in 1u32..=1_000u32,
        val in i64::MIN..=i64::MAX,
    ) {
        let mut vm = Vm::new();
        vm.cells_mut().insert(
            (row, col),
            CellContent { formula: None, value: Variant::Integer(val) },
        );
        prop_assert_eq!(vm.get_cell(row, col), Variant::Integer(val));
    }

    /// set_cell then get_cell roundtrip for string values.
    #[test]
    fn prop_cell_string_roundtrip(
        row in 1u32..=10_000u32,
        col in 1u32..=1_000u32,
        s in "[^\x00]{0,200}",
    ) {
        let mut vm = Vm::new();
        vm.cells_mut().insert(
            (row, col),
            CellContent { formula: None, value: Variant::Str(s.clone()) },
        );
        prop_assert_eq!(vm.get_cell(row, col), Variant::Str(s));
    }

    /// last_nonempty_row always returns the true maximum written row.
    #[test]
    fn prop_last_nonempty_row_is_max(
        rows in proptest::collection::vec(1u32..=10_000u32, 1..=30)
    ) {
        let mut vm = Vm::new();
        for &r in &rows {
            vm.cells_mut().insert(
                (r, 1),
                CellContent { formula: None, value: Variant::Integer(1) },
            );
        }
        let expected = *rows.iter().max().unwrap();
        prop_assert_eq!(vm.last_nonempty_row(1, 1_048_576), expected);
    }

    /// last_nonempty_col always returns the true maximum written column.
    #[test]
    fn prop_last_nonempty_col_is_max(
        cols in proptest::collection::vec(1u32..=1_000u32, 1..=20)
    ) {
        let mut vm = Vm::new();
        for &c in &cols {
            vm.cells_mut().insert(
                (1, c),
                CellContent { formula: None, value: Variant::Integer(1) },
            );
        }
        let expected = *cols.iter().max().unwrap();
        prop_assert_eq!(vm.last_nonempty_col(1, 16_384), expected);
    }

    /// Clearing all cells then querying last_nonempty_row returns the default (1).
    #[test]
    fn prop_clear_resets_row_index(
        rows in proptest::collection::vec(1u32..=1_000u32, 1..=10)
    ) {
        let mut vm = Vm::new();
        for &r in &rows {
            vm.cells_mut().insert(
                (r, 1),
                CellContent { formula: None, value: Variant::Integer(1) },
            );
        }
        vm.cells_mut().clear();
        // Empty column → default 1
        prop_assert_eq!(vm.last_nonempty_row(1, 1_048_576), 1u32);
    }

    /// Multiple writes to the same cell — only the last value survives.
    #[test]
    fn prop_overwrite_cell(
        row in 1u32..=1_000u32,
        col in 1u32..=100u32,
        v1  in i64::MIN..=i64::MAX,
        v2  in i64::MIN..=i64::MAX,
    ) {
        let mut vm = Vm::new();
        vm.cells_mut().insert((row, col), CellContent { formula: None, value: Variant::Integer(v1) });
        vm.cells_mut().insert((row, col), CellContent { formula: None, value: Variant::Integer(v2) });
        prop_assert_eq!(vm.get_cell(row, col), Variant::Integer(v2));
    }
}

// ── InStr safety ──────────────────────────────────────────────────────────────

proptest! {
    /// InStr must never panic — any haystack, needle, and start position is safe.
    #[test]
    fn prop_instr_never_panics(
        haystack in "[a-zA-Z0-9 ]{0,50}",
        needle   in "[a-zA-Z0-9]{0,10}",
        start    in 0usize..=60usize,
    ) {
        // Build and execute the VBA code; parse returning Err is also fine.
        let code = format!(
            "Sub MySub()\n    x = InStr({}, \"{}\", \"{}\")\nEnd Sub\n",
            start, haystack, needle,
        );
        if let Ok(prog) = parser::parse(&code) {
            let mut vm = Vm::new();
            let _ = vm.run_sub(&prog, "MySub");
        }
    }
}

// ── String comparison monotonicity ────────────────────────────────────────────

proptest! {
    /// VBA string < comparison must be consistent with lexicographic order.
    #[test]
    fn prop_string_lt_is_lexicographic(
        a in "[a-z]{1,8}",
        b in "[a-z]{1,8}",
    ) {
        let code = format!(
            concat!(
                "Sub MySub()\n",
                "    If {:?} < {:?} Then\n",
                "        result = 1\n",
                "    Else\n",
                "        result = 0\n",
                "    End If\n",
                "End Sub\n",
            ),
            a, b
        );
        if let Ok(prog) = parser::parse(&code) {
            let mut vm = Vm::new();
            if vm.run_sub(&prog, "MySub").is_ok() {
                let expected: i64 = if a.to_uppercase() < b.to_uppercase() { 1 } else { 0 };
                prop_assert_eq!(
                    vm.variables.get("result").cloned().unwrap_or(Variant::Integer(0)),
                    Variant::Integer(expected),
                    "a={:?} b={:?}", a, b
                );
            }
        }
    }

    /// String equality comparison must be symmetric.
    #[test]
    fn prop_string_eq_symmetric(
        a in "[a-zA-Z]{1,10}",
        b in "[a-zA-Z]{1,10}",
    ) {
        let code_ab = format!(
            "Sub MySub()\n    If {:?} = {:?} Then\n        r = 1\n    Else\n        r = 0\n    End If\nEnd Sub\n",
            a, b
        );
        let code_ba = format!(
            "Sub MySub()\n    If {:?} = {:?} Then\n        r = 1\n    Else\n        r = 0\n    End If\nEnd Sub\n",
            b, a
        );
        if let (Ok(p1), Ok(p2)) = (parser::parse(&code_ab), parser::parse(&code_ba)) {
            let mut vm1 = Vm::new();
            let mut vm2 = Vm::new();
            if vm1.run_sub(&p1, "MySub").is_ok() && vm2.run_sub(&p2, "MySub").is_ok() {
                prop_assert_eq!(
                    vm1.variables.get("r"),
                    vm2.variables.get("r"),
                    "equality not symmetric: {:?} vs {:?}", a, b
                );
            }
        }
    }
}
