//! Milestone B6d: the `diagnose-workbook` subcommand — reuses
//! `test-workbook`'s (Milestone B5a) generated-case search
//! (`testworkbook::run_fixture`, called with `strict: true`) and enriches
//! whichever failures are classifiable with `diagnose`'s (Milestones
//! B6a–B6c2) existing root-cause machinery, via the single `pub(crate)`
//! entry point `diagnose::root_causes_json`.
//!
//! Most root causes are **structural** (the 3 merge kinds, shape mismatch,
//! empty-clipboard paste, sheet protection) — they depend on the macro's own
//! text and the workbook's fixed layout, not on which boundary value a case
//! draws, so they fire identically on case 0 (or never) regardless of how
//! many cases run. This command earns its keep specifically for
//! **input-dependent** kinds (`ArrayIndexOutOfBounds` chief among them,
//! where a drawn value can flip an index in or out of bounds across cases) —
//! a single `diagnose` invocation already finds the structural kinds in one
//! shot. Every other runtime error (the vast majority — type mismatches,
//! division by zero, etc.) has no classification at all, matching
//! `diagnose`'s own permanent limitation: `root_causes` is `[]`, only the
//! bare message is reported, same as `test-workbook` already does today.

use crate::diagnose::{observations_json, root_causes_json};
use crate::diagnostics::{json_string, variant_to_json};
use crate::testworkbook::FixtureResult;

/// `test-workbook`'s existing failure JSON shape
/// (`schema_version`/`ok`/`seed`/`case_index`/`inputs`/`failure`) plus one
/// sibling field, `root_causes` — `"[]"` when the failure isn't one of
/// `diagnose`'s classified kinds (true for most runtime errors), a one-item
/// array when it is. Success shape is `test-workbook`'s, unchanged.
pub fn to_json(result: &FixtureResult) -> String {
    match result {
        FixtureResult::Passed { seed, cases_run, hidden_cells } => {
            // Milestone B7b: a sibling field, present only when `Some` —
            // same "never an always-present empty array" contract as
            // plain `diagnose`'s own `observations` field.
            let observations_field = if hidden_cells.is_some() {
                format!(",\"observations\":{}", observations_json(hidden_cells.as_deref()))
            } else {
                String::new()
            };
            format!(
                "{{\"schema_version\":1,\"ok\":true,\"seed\":{},\"cases_run\":{}{}}}",
                seed, cases_run, observations_field
            )
        }
        FixtureResult::Failed {
            seed,
            case_index,
            inputs_used,
            failure,
            resolution_kind,
            hidden_cells,
        } => {
            let inputs_json: Vec<String> = inputs_used
                .iter()
                .map(|iu| {
                    format!(
                        "{{\"address\":{},\"value\":{}}}",
                        json_string(&iu.address),
                        variant_to_json(&iu.value)
                    )
                })
                .collect();
            let mut fields = vec![format!("\"rule\":{}", json_string(&failure.rule))];
            if let Some(a) = &failure.address {
                fields.push(format!("\"address\":{}", json_string(a)));
            }
            if let Some(a) = &failure.actual {
                fields.push(format!("\"actual\":{}", json_string(a)));
            }
            if let Some(m) = &failure.message {
                fields.push(format!("\"message\":{}", json_string(m)));
            }
            let observations_field = if hidden_cells.is_some() {
                format!(",\"observations\":{}", observations_json(hidden_cells.as_deref()))
            } else {
                String::new()
            };
            format!(
                "{{\"schema_version\":1,\"ok\":false,\"seed\":{},\"case_index\":{},\"inputs\":[{}],\"failure\":{{{}}},\"root_causes\":{}{}}}",
                seed,
                case_index,
                inputs_json.join(","),
                fields.join(","),
                root_causes_json(resolution_kind.as_deref()),
                observations_field,
            )
        }
    }
}

/// Plain-text summary — `test-workbook`'s own rule/address/actual/message
/// line, plus a `root cause: <CODE>` line when classified. Full evidence
/// and suggestions are `--json`-only, matching every other subcommand's own
/// "plain text is a simplified view" convention (`snapshot`, `diagnose`).
pub fn to_plain_text(result: &FixtureResult) -> String {
    match result {
        FixtureResult::Passed { seed, cases_run, hidden_cells } => {
            let mut line = format!("ok: {} case(s) passed (seed {})", cases_run, seed);
            if let Some(obs) = hidden_cells {
                line.push_str(&format!(
                    "\n  observation: RANGE_CONTAINS_HIDDEN_CELLS ({} of {} cells visible in {})",
                    obs.visible_cells, obs.total_cells, obs.address
                ));
            }
            line
        }
        FixtureResult::Failed {
            seed,
            case_index,
            failure,
            resolution_kind,
            hidden_cells,
            ..
        } => {
            let mut line = format!(
                "FAIL: case {} (seed {}) - {}",
                case_index, seed, failure.rule
            );
            if let Some(a) = &failure.address {
                line.push_str(&format!(" at {}", a));
            }
            if let Some(a) = &failure.actual {
                line.push_str(&format!(": {}", a));
            }
            if let Some(m) = &failure.message {
                line.push_str(&format!(": {}", m));
            }
            if let Some(kind) = resolution_kind {
                let rc = crate::diagnose::RootCause::from_kind((**kind).clone());
                line.push_str(&format!("\n  root cause: {}", rc.code));
            }
            if let Some(obs) = hidden_cells {
                line.push_str(&format!(
                    "\n  observation: RANGE_CONTAINS_HIDDEN_CELLS ({} of {} cells visible in {})",
                    obs.visible_cells, obs.total_cells, obs.address
                ));
            }
            line
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::testworkbook::{FailureDetail, InputUsed};
    use crate::vm::{ResolutionFailureKind, Variant};

    #[test]
    fn to_json_success_shape_matches_test_workbook() {
        let json = to_json(&FixtureResult::Passed {
            seed: 42,
            cases_run: 100,
            hidden_cells: None,
        });
        assert!(json.contains("\"ok\":true"));
        assert!(json.contains("\"seed\":42"));
        assert!(json.contains("\"cases_run\":100"));
    }

    #[test]
    fn to_json_unclassified_failure_reports_empty_root_causes() {
        let result = FixtureResult::Failed {
            seed: 7,
            case_index: 3,
            inputs_used: vec![InputUsed {
                address: "Input!B2".to_string(),
                value: Variant::Integer(-1),
            }],
            failure: FailureDetail {
                rule: "no_runtime_error".to_string(),
                address: None,
                actual: None,
                message: Some("Division by zero".to_string()),
            },
            resolution_kind: None,
            hidden_cells: None,
        };
        let json = to_json(&result);
        assert!(json.contains("\"root_causes\":[]"));
        assert!(json.contains("\"message\":\"Division by zero\""));
    }

    #[test]
    fn to_json_classified_failure_reports_a_root_cause() {
        let result = FixtureResult::Failed {
            seed: 42,
            case_index: 17,
            inputs_used: vec![InputUsed {
                address: "Input!B2".to_string(),
                value: Variant::Integer(999_999_999),
            }],
            failure: FailureDetail {
                rule: "no_runtime_error".to_string(),
                address: None,
                actual: None,
                message: Some("array 'arr' index 999999999 out of bounds".to_string()),
            },
            resolution_kind: Some(Box::new(ResolutionFailureKind::ArrayIndexOutOfBounds {
                name: "arr".to_string(),
                index: 999_999_999,
                lower: 0,
                upper: 9,
            })),
            hidden_cells: None,
        };
        let json = to_json(&result);
        assert!(json.contains("\"code\":\"ARRAY_INDEX_OUT_OF_BOUNDS\""));
        assert!(json.contains("\"name\":\"arr\""));
        assert!(json.contains("\"index\":999999999"));
    }

    #[test]
    fn to_plain_text_appends_root_cause_code_when_classified() {
        let result = FixtureResult::Failed {
            seed: 42,
            case_index: 17,
            inputs_used: vec![],
            failure: FailureDetail {
                rule: "no_runtime_error".to_string(),
                address: None,
                actual: None,
                message: Some("array 'arr' index 999999999 out of bounds".to_string()),
            },
            resolution_kind: Some(Box::new(ResolutionFailureKind::ArrayIndexOutOfBounds {
                name: "arr".to_string(),
                index: 999_999_999,
                lower: 0,
                upper: 9,
            })),
            hidden_cells: None,
        };
        let text = to_plain_text(&result);
        assert!(text.contains("root cause: ARRAY_INDEX_OUT_OF_BOUNDS"));
    }

    #[test]
    fn to_plain_text_omits_root_cause_line_when_unclassified() {
        let result = FixtureResult::Failed {
            seed: 42,
            case_index: 17,
            inputs_used: vec![],
            failure: FailureDetail {
                rule: "no_runtime_error".to_string(),
                address: None,
                actual: None,
                message: Some("Division by zero".to_string()),
            },
            resolution_kind: None,
            hidden_cells: None,
        };
        let text = to_plain_text(&result);
        assert!(!text.contains("root cause"));
    }

    // ── Milestone B7b: hidden row/column evidence ────────────────────────────
    //
    // No end-to-end `run_fixture` test here: the observation is computed
    // purely from `Vm.clipboard`/`Vm.sheet_visibility`, and `sheet_visibility`
    // can only be populated from a real workbook file's hidden-row/column
    // metadata (the XLSX writer can't emit it — see the B7b plan's
    // grounding). So, same test-debt precedent as B6c2's merged cells:
    // `FixtureResult` is constructed directly here to test the rendering
    // layer, and the "case 0 and a later case report the same observation"
    // claim (Milestone B6d's honest-scope framing) holds by construction —
    // `Vm::hidden_cells_observation` never reads a case's drawn cell
    // values, only the copied range's geometry and the sheet's hidden-row/
    // column metadata, neither of which varies across cases.
    use crate::vm::{HiddenCellsObservation, Interval};

    fn sample_observation() -> HiddenCellsObservation {
        HiddenCellsObservation {
            sheet: "sheet1".to_string(),
            address: "A1:C10".to_string(),
            rows: 10,
            columns: 3,
            hidden_rows: vec![Interval { start: 3, end: 5 }],
            hidden_columns: vec![],
            total_cells: 30,
            visible_cells: 21,
        }
    }

    #[test]
    fn to_json_passing_fixture_includes_observations_when_present() {
        let json = to_json(&FixtureResult::Passed {
            seed: 42,
            cases_run: 20,
            hidden_cells: Some(Box::new(sample_observation())),
        });
        assert!(json.contains("\"ok\":true"));
        assert!(json.contains("\"observations\":[{\"code\":\"RANGE_CONTAINS_HIDDEN_CELLS\""));
        assert!(json.contains("\"visible_cells\":21"));
    }

    #[test]
    fn to_json_passing_fixture_omits_observations_when_absent() {
        let json = to_json(&FixtureResult::Passed {
            seed: 42,
            cases_run: 20,
            hidden_cells: None,
        });
        assert!(!json.contains("observations"));
    }

    #[test]
    fn to_json_failing_fixture_includes_observations_alongside_root_causes() {
        let result = FixtureResult::Failed {
            seed: 42,
            case_index: 3,
            inputs_used: vec![],
            failure: FailureDetail {
                rule: "no_runtime_error".to_string(),
                address: None,
                actual: None,
                message: Some("array 'arr' index 999999999 out of bounds".to_string()),
            },
            resolution_kind: Some(Box::new(ResolutionFailureKind::ArrayIndexOutOfBounds {
                name: "arr".to_string(),
                index: 999_999_999,
                lower: 0,
                upper: 9,
            })),
            hidden_cells: Some(Box::new(sample_observation())),
        };
        let json = to_json(&result);
        assert!(json.contains("\"code\":\"ARRAY_INDEX_OUT_OF_BOUNDS\""));
        assert!(json.contains("\"observations\":[{\"code\":\"RANGE_CONTAINS_HIDDEN_CELLS\""));
    }

    #[test]
    fn to_plain_text_passing_fixture_appends_the_observation_note() {
        let text = to_plain_text(&FixtureResult::Passed {
            seed: 42,
            cases_run: 20,
            hidden_cells: Some(Box::new(sample_observation())),
        });
        assert!(text.starts_with("ok: 20 case(s) passed"));
        assert!(text.contains("RANGE_CONTAINS_HIDDEN_CELLS"));
        assert!(text.contains("21 of 30"));
    }
}
