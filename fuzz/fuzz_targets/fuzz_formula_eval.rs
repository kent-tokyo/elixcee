#![no_main]

use libfuzzer_sys::fuzz_target;
use elixcee::vm::{CellContent, Variant};
use std::collections::HashMap;

fuzz_target!(|data: &[u8]| {
    // parse then evaluate — neither step must panic.
    if let Ok(s) = std::str::from_utf8(data) {
        if let Ok(expr) = elixcee::formula::parse(s) {
            // Provide a small fixed cell environment so cell references resolve.
            let mut cells: HashMap<(u32, u32), CellContent> = HashMap::new();
            for r in 1u32..=5 {
                for c in 1u32..=5 {
                    cells.insert(
                        (r, c),
                        CellContent {
                            formula: None,
                            value: Variant::Integer((r * c) as i64),
                        },
                    );
                }
            }
            let _ = elixcee::formula::evaluate(&expr, &cells);
        }
    }
});
