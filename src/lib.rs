pub mod check;
pub mod diagnose;
pub mod diagnoseworkbook;
pub mod diagnostics;
pub mod formula;
pub mod parser;
pub mod reader;
pub mod snapshot;
pub mod testworkbook;
pub mod vm;

/// Shared value types (`Variant`, `ExcelError`, `CellContent`, date-serial
/// math), physically defined in the `elixcee-types` crate — aliased here so
/// existing `crate::types::*` references (used internally by `vm`/`formula`)
/// resolve without every call site needing to know it's an external crate.
pub use elixcee_types as types;

#[cfg(any(feature = "python", test))]
use vm::CellContent;
#[cfg(any(feature = "python", test))]
use vm::{FillEdit, StyleAttrEdit};
use vm::{Variant, Vm, WorksheetOrigin};

#[cfg(feature = "python")]
use pyo3::prelude::*;
#[cfg(feature = "python")]
use pyo3::types::{PyDict, PyList};
#[cfg(feature = "python")]
use vm::{ExcelError, serial_to_display};

// ── ExcelError Python class ───────────────────────────────────────────────────

/// Represents an Excel cell error value (#N/A, #VALUE!, #DIV/0!, etc.).
/// Returned by ``get_cell`` and ``cells()`` for error cells, and accepted by
/// ``set_cell`` to store an error value.
// `from_py_object` opts in explicitly to the auto-derived `FromPyObject` impl pyo3 is
// deprecating as implicit-by-default for `Clone` pyclasses -- kept, not dropped, because
// `python_to_variant` (below) does `obj.extract::<PyExcelError>()` to accept an ExcelError
// a Python caller passes back into `set_cell`; `skip_from_py_object` would silently break
// that real, documented round-trip.
#[cfg(feature = "python")]
#[pyclass(name = "ExcelError", from_py_object)]
#[derive(Clone, Debug)]
pub struct PyExcelError {
    /// The error string, e.g. ``"#N/A"``, ``"#VALUE!"``, ``"#DIV/0!"``.
    #[pyo3(get)]
    pub code: String,
}

#[cfg(feature = "python")]
#[pymethods]
impl PyExcelError {
    #[new]
    fn new(code: String) -> Self {
        PyExcelError { code }
    }
    fn __repr__(&self) -> String {
        format!("ExcelError('{}')", self.code)
    }
    fn __str__(&self) -> String {
        self.code.clone()
    }
    fn __eq__(&self, other: &PyExcelError) -> bool {
        self.code == other.code
    }
    fn __hash__(&self) -> isize {
        self.code.len() as isize
    }
}

// ── Variant ↔ Python conversion ───────────────────────────────────────────────

#[cfg(feature = "python")]
fn variant_to_py(py: Python<'_>, v: &Variant) -> Py<PyAny> {
    match v {
        Variant::Integer(n) => (*n).into_pyobject(py).unwrap().into_any().unbind(),
        Variant::Float(f) => (*f).into_pyobject(py).unwrap().into_any().unbind(),
        Variant::Str(s) => s.as_str().into_pyobject(py).unwrap().into_any().unbind(),
        Variant::Boolean(b) => {
            let borrowed = (*b).into_pyobject(py).unwrap();
            <pyo3::Bound<'_, pyo3::types::PyBool> as Clone>::clone(&borrowed)
                .unbind()
                .into_any()
        }
        Variant::Date(s) => {
            let (y, m, d) = crate::types::serial_to_ymd(*s);
            pyo3::types::PyDate::new(py, y, m as u8, d as u8)
                .map(|dt| dt.into_any().unbind())
                .unwrap_or_else(|_| {
                    serial_to_display(*s)
                        .into_pyobject(py)
                        .unwrap()
                        .into_any()
                        .unbind()
                })
        }
        Variant::Error(e) => PyExcelError {
            code: e.as_str().to_string(),
        }
        .into_pyobject(py)
        .unwrap()
        .into_any()
        .unbind(),
        // VBA's `Null` crosses into Python as `None`, exactly as `Empty`
        // already does — neither has a Python value, and giving Null its own
        // Python representation would be a bindings-contract change. The
        // Empty-vs-Null distinction is observable through the VBA language
        // (`IsNull`, `TypeName`, `VarType`), not across this boundary.
        Variant::Empty | Variant::Null => py.None(),
        Variant::Array(a) => {
            let list =
                pyo3::types::PyList::new(py, a.iter().map(|x| variant_to_py(py, x))).unwrap();
            list.into_any().unbind()
        }
        Variant::VbaArray(a) => vba_array_to_py(py, a, &mut Vec::new()),
        Variant::Record(m) => {
            let dict = pyo3::types::PyDict::new(py);
            for (k, v) in m {
                dict.set_item(k, variant_to_py(py, v)).unwrap();
            }
            dict.into_any().unbind()
        }
    }
}

/// Reshapes a `VbaArray` into nested Python lists matching its real shape —
/// `Dim arr(1 To 2, 1 To 3)` crosses into Python as a 2-element list of
/// 3-element lists, not one flat 6-element list. `prefix` accumulates the
/// index down each recursion level; empty on the initial call.
#[cfg(feature = "python")]
fn vba_array_to_py(py: Python<'_>, arr: &vm::VbaArray, prefix: &mut Vec<i64>) -> Py<PyAny> {
    if prefix.len() == arr.bounds.len() {
        let v = arr.get(prefix).expect("prefix built from arr's own bounds");
        return variant_to_py(py, v);
    }
    let bound = arr.bounds[prefix.len()];
    let mut items = Vec::new();
    let mut i = bound.lower;
    while i <= bound.upper {
        prefix.push(i);
        items.push(vba_array_to_py(py, arr, prefix));
        prefix.pop();
        i += 1;
    }
    pyo3::types::PyList::new(py, items)
        .unwrap()
        .into_any()
        .unbind()
}

#[cfg(feature = "python")]
fn py_to_variant(obj: &Bound<'_, PyAny>) -> PyResult<Variant> {
    if obj.is_none() {
        return Ok(Variant::Empty);
    }
    // bool must come before int (Python bool is a subclass of int)
    if let Ok(b) = obj.extract::<bool>() {
        return Ok(Variant::Boolean(b));
    }
    if let Ok(n) = obj.extract::<i64>() {
        return Ok(Variant::Integer(n));
    }
    if let Ok(f) = obj.extract::<f64>() {
        return Ok(Variant::Float(f));
    }
    if let Ok(s) = obj.extract::<String>() {
        return Ok(Variant::Str(s));
    }
    if let Ok(e) = obj.extract::<PyExcelError>() {
        return Ok(Variant::Error(match e.code.as_str() {
            "#DIV/0!" => ExcelError::DivZero,
            "#N/A" => ExcelError::NA,
            "#VALUE!" => ExcelError::Value,
            "#REF!" => ExcelError::Ref,
            "#NAME?" => ExcelError::Name,
            "#NUM!" => ExcelError::Num,
            "#NULL!" => ExcelError::Null,
            _ => ExcelError::Value,
        }));
    }
    Err(PyErr::new::<pyo3::exceptions::PyTypeError, _>(
        "Unsupported cell value type",
    ))
}

// ── `set_style` (0.15.0-B) — dict-to-Edit-struct parsing ────────────────────────

/// Normalizes a caller-supplied color string to 8-hex-digit ARGB (real Excel's own
/// `rgb="..."` attribute shape) -- accepts a 6-digit RGB (alpha assumed fully opaque,
/// `"FF"`) or an already-8-digit ARGB, with or without a leading `#`. Case-insensitive
/// on input, always emits uppercase (matching Excel/openpyxl's own convention, confirmed
/// against a real openpyxl-authored fixture).
#[cfg(feature = "python")]
fn normalize_color_hex(input: &str) -> PyResult<String> {
    let s = input.trim().trim_start_matches('#');
    let is_hex = !s.is_empty() && s.chars().all(|c| c.is_ascii_hexdigit());
    match s.len() {
        6 if is_hex => Ok(format!("FF{}", s.to_uppercase())),
        8 if is_hex => Ok(s.to_uppercase()),
        _ => Err(PyErr::new::<pyo3::exceptions::PyValueError, _>(format!(
            "color must be a 6-digit RGB or 8-digit ARGB hex string, got {input:?}"
        ))),
    }
}

/// Extracts `dict[key]` as `String`, `TypeError` on a wrong-typed value, `None` if absent.
#[cfg(feature = "python")]
fn dict_str(dict: &Bound<'_, PyDict>, key: &str) -> PyResult<Option<String>> {
    match dict.get_item(key)? {
        Some(v) => Ok(Some(v.extract().map_err(|_| {
            PyErr::new::<pyo3::exceptions::PyTypeError, _>(format!("{key:?} must be a string"))
        })?)),
        None => Ok(None),
    }
}

#[cfg(feature = "python")]
fn dict_bool(dict: &Bound<'_, PyDict>, key: &str) -> PyResult<Option<bool>> {
    match dict.get_item(key)? {
        Some(v) => Ok(Some(v.extract().map_err(|_| {
            PyErr::new::<pyo3::exceptions::PyTypeError, _>(format!("{key:?} must be a bool"))
        })?)),
        None => Ok(None),
    }
}

#[cfg(feature = "python")]
fn dict_f64(dict: &Bound<'_, PyDict>, key: &str) -> PyResult<Option<f64>> {
    match dict.get_item(key)? {
        Some(v) => Ok(Some(v.extract().map_err(|_| {
            PyErr::new::<pyo3::exceptions::PyTypeError, _>(format!("{key:?} must be a number"))
        })?)),
        None => Ok(None),
    }
}

#[cfg(feature = "python")]
fn dict_u32(dict: &Bound<'_, PyDict>, key: &str) -> PyResult<Option<u32>> {
    match dict.get_item(key)? {
        Some(v) => Ok(Some(v.extract().map_err(|_| {
            PyErr::new::<pyo3::exceptions::PyTypeError, _>(format!(
                "{key:?} must be a non-negative integer"
            ))
        })?)),
        None => Ok(None),
    }
}

#[cfg(feature = "python")]
fn dict_color(dict: &Bound<'_, PyDict>, key: &str) -> PyResult<Option<String>> {
    match dict_str(dict, key)? {
        Some(s) => Ok(Some(normalize_color_hex(&s)?)),
        None => Ok(None),
    }
}

#[cfg(feature = "python")]
fn extract_font_edit(dict: &Bound<'_, PyDict>) -> PyResult<reader::FontEdit> {
    Ok(reader::FontEdit {
        bold: dict_bool(dict, "bold")?,
        italic: dict_bool(dict, "italic")?,
        underline: dict_bool(dict, "underline")?,
        strike: dict_bool(dict, "strike")?,
        size: dict_f64(dict, "size")?,
        color_argb: dict_color(dict, "color")?,
        name: dict_str(dict, "name")?,
    })
}

#[cfg(feature = "python")]
fn extract_fill_edit(dict: &Bound<'_, PyDict>) -> PyResult<FillEdit> {
    let fill_type = dict_str(dict, "type")?.unwrap_or_else(|| "solid".to_string());
    if fill_type != "solid" {
        return Err(PyErr::new::<pyo3::exceptions::PyValueError, _>(format!(
            "fill['type'] {fill_type:?} not supported yet -- only \"solid\" is implemented \
             (0.15.0-B); gradient/pattern fills are a future milestone"
        )));
    }
    let color_argb = dict_color(dict, "color")?.ok_or_else(|| {
        PyErr::new::<pyo3::exceptions::PyValueError, _>("fill requires a 'color'")
    })?;
    Ok(FillEdit { color_argb })
}

#[cfg(feature = "python")]
fn extract_border_side_edit(
    dict: &Bound<'_, PyDict>,
    side_name: &str,
) -> PyResult<Option<reader::BorderSideEdit>> {
    let Some(side_val) = dict.get_item(side_name)? else {
        return Ok(None);
    };
    let side_dict = side_val.cast::<PyDict>().map_err(|_| {
        PyErr::new::<pyo3::exceptions::PyTypeError, _>(format!(
            "border[{side_name:?}] must be a dict"
        ))
    })?;
    Ok(Some(reader::BorderSideEdit {
        style: dict_str(side_dict, "style")?,
        color_argb: dict_color(side_dict, "color")?,
    }))
}

#[cfg(feature = "python")]
fn extract_border_edit(dict: &Bound<'_, PyDict>) -> PyResult<reader::BorderEdit> {
    Ok(reader::BorderEdit {
        left: extract_border_side_edit(dict, "left")?,
        right: extract_border_side_edit(dict, "right")?,
        top: extract_border_side_edit(dict, "top")?,
        bottom: extract_border_side_edit(dict, "bottom")?,
        diagonal: extract_border_side_edit(dict, "diagonal")?,
    })
}

#[cfg(feature = "python")]
fn extract_alignment_edit(dict: &Bound<'_, PyDict>) -> PyResult<reader::AlignmentEdit> {
    Ok(reader::AlignmentEdit {
        horizontal: dict_str(dict, "horizontal")?,
        vertical: dict_str(dict, "vertical")?,
        wrap_text: dict_bool(dict, "wrap_text")?,
        indent: dict_u32(dict, "indent")?,
    })
}

#[cfg(feature = "python")]
fn extract_protection_edit(dict: &Bound<'_, PyDict>) -> PyResult<reader::ProtectionEdit> {
    Ok(reader::ProtectionEdit {
        locked: dict_bool(dict, "locked")?,
        hidden: dict_bool(dict, "hidden")?,
    })
}

// ── Bulk worksheet range/row API (R1) — address/shape validation ───────────────
//
// These two functions are deliberately NOT `#[cfg(feature = "python")]` — they
// take no pyo3 types at all, and `cargo check --features python --lib` (the
// only CI step that type-checks the gated PyVm code below) never *runs*
// anything, so keeping this pure validation logic ungated is what gives it
// any automated test coverage at all (via plain `cargo test --workspace`).

#[cfg_attr(not(feature = "python"), allow(dead_code))]
type RangeBounds = ((u32, u32), (u32, u32));

/// Per-sheet, per-cell style index -- `Vm::cell_style_indices`'s own shape, reused here
/// for `resolve_pending_number_formats`'s "effective" (edits applied) return value.
type StyleIndexMap = std::collections::HashMap<String, std::collections::HashMap<(u32, u32), u32>>;

/// `<cols>` emission's merge-by-exact-range accumulator (`build_xlsx_sheet`) --
/// `(hidden, width, style)`, one entry per exact `(min,max)` range.
type ColAttrsMap = std::collections::BTreeMap<(u32, u32), (bool, Option<f64>, Option<u32>)>;

/// Validates and parses a single-area A1 range address for the bulk-range
/// Python API (`get_range`/`set_range`). Deliberately NOT a new A1 parser —
/// delegates to `crate::types::parse_range_addr` for the actual grammar.
/// Adds only, as explicit errors instead of the shared parser's silent
/// `None`:
///   - `$`-stripping before delegating (Excel absolute-reference syntax;
///     `elixcee-types`'s column-letter parsing does an unchecked `u32`
///     subtraction that underflows on a leading `$` today — a real,
///     many-call-site, pre-existing gap in the shared parser that this
///     closes ONLY for calls that go through this wrapper, not
///     project-wide; see docs/openpyxl-gap-audit.md),
///   - multi-area (`,`-containing) rejection,
///   - reversed-range rejection (`start > end`), matching the precedent
///     `reader.rs`'s own `parse_dimension_ref` already sets for dimension
///     refs,
///   - row/col `0` rejection (`parse_cell_addr("A0")` succeeds as `(0,1)`
///     today — another disclosed, out-of-scope shared-parser gap).
///
/// Both this and `check_grid_shape` below are only ever called from the
/// `#[cfg(feature = "python")]` `PyVm` methods, but are deliberately left
/// ungated themselves (see the section comment above) — a plain,
/// feature-less build has no caller for them outside `#[cfg(test)]`, hence
/// the narrow, conditional `dead_code` allow rather than a broad one.
#[cfg_attr(not(feature = "python"), allow(dead_code))]
fn validate_range_addr(addr: &str) -> Result<RangeBounds, String> {
    if addr.contains(',') {
        return Err(format!("multi-area address not supported: {addr:?}"));
    }
    let stripped = addr.replace('$', "");
    let (start, end) = crate::types::parse_range_addr(&stripped)
        .ok_or_else(|| format!("invalid range address: {addr:?}"))?;
    if start.0 == 0 || start.1 == 0 || end.0 == 0 || end.1 == 0 {
        return Err(format!(
            "invalid range address (row/column must be >= 1): {addr:?}"
        ));
    }
    if start.0 > end.0 || start.1 > end.1 {
        return Err(format!("reversed range address: {addr:?}"));
    }
    Ok((start, end))
}

/// Validates a Python-supplied nested-list grid's shape against `set_range`'s
/// target rect, given each already-extracted outer-list row's length in
/// order. Two distinct failure messages: ragged input (row lengths disagree
/// with each other) vs. shape mismatch (rectangular, but wrong size) — both
/// surfaced as `ValueError` at the call site. Never indexes into the
/// original values, only their pre-collected lengths, so it can't panic on
/// empty input (`row_lens == []`: no ragged check fires, falls straight to
/// the 0x0-vs-expected shape-mismatch branch).
#[cfg_attr(not(feature = "python"), allow(dead_code))]
fn check_grid_shape(expected: (u32, u32), row_lens: &[usize]) -> Result<(), String> {
    if let Some(&first) = row_lens.first()
        && row_lens.iter().any(|&n| n != first)
    {
        return Err(format!(
            "ragged input: row lengths must all be equal, got {row_lens:?}"
        ));
    }
    let (expected_rows, expected_cols) = (expected.0 as usize, expected.1 as usize);
    let actual_rows = row_lens.len();
    let actual_cols = row_lens.first().copied().unwrap_or(0);
    if actual_rows != expected_rows || actual_cols != expected_cols {
        return Err(format!(
            "shape mismatch: range expects {expected_rows}x{expected_cols}, got {actual_rows}x{actual_cols}"
        ));
    }
    Ok(())
}

#[cfg(test)]
mod bulk_range_validation_tests {
    use super::*;

    #[test]
    fn validate_range_addr_accepts_a_normal_range() {
        assert_eq!(validate_range_addr("A1:C5").unwrap(), ((1, 1), (5, 3)));
    }

    #[test]
    fn validate_range_addr_accepts_a_bare_single_cell() {
        assert_eq!(validate_range_addr("B2").unwrap(), ((2, 2), (2, 2)));
    }

    #[test]
    fn validate_range_addr_strips_dollar_signs() {
        assert_eq!(validate_range_addr("$A$1:$C$5").unwrap(), ((1, 1), (5, 3)));
    }

    #[test]
    fn validate_range_addr_rejects_multi_area() {
        let err = validate_range_addr("A1:B2,D1:E2").unwrap_err();
        assert!(err.contains("multi-area"), "{err:?}");
    }

    #[test]
    fn validate_range_addr_rejects_malformed_input() {
        assert!(validate_range_addr("!!").is_err());
        assert!(validate_range_addr("").is_err());
    }

    #[test]
    fn validate_range_addr_rejects_a_reversed_range() {
        let err = validate_range_addr("C3:A1").unwrap_err();
        assert!(err.contains("reversed"), "{err:?}");
    }

    #[test]
    fn validate_range_addr_rejects_row_or_col_zero() {
        assert!(validate_range_addr("A0").is_err());
        assert!(validate_range_addr("A0:B1").is_err());
    }

    #[test]
    fn check_grid_shape_accepts_an_exact_match() {
        check_grid_shape((2, 3), &[3, 3]).unwrap();
    }

    #[test]
    fn check_grid_shape_rejects_ragged_input() {
        let err = check_grid_shape((2, 3), &[3, 2]).unwrap_err();
        assert!(err.contains("ragged"), "{err:?}");
    }

    #[test]
    fn check_grid_shape_rejects_a_shape_mismatch() {
        let err = check_grid_shape((2, 3), &[2, 2]).unwrap_err();
        assert!(err.contains("2x3"), "{err:?}");
        assert!(err.contains("2x2"), "{err:?}");
    }

    #[test]
    fn check_grid_shape_on_empty_input_does_not_panic() {
        let err = check_grid_shape((2, 3), &[]).unwrap_err();
        assert!(err.contains("0x0"), "{err:?}");
    }
}

// ── PyVm class ────────────────────────────────────────────────────────────────

/// VBA execution engine. Create one, pre-populate cells with ``set_cell``,
/// run a macro with ``run``, then read results via ``get_cell`` / ``cells``.
#[cfg(feature = "python")]
#[pyclass(name = "Vm")]
pub struct PyVm {
    inner: Vm,
}

#[cfg(feature = "python")]
#[pymethods]
impl PyVm {
    #[new]
    #[pyo3(signature = (on_msgbox = "skip"))]
    fn new(on_msgbox: &str) -> PyResult<Self> {
        let mut vm = Vm::new();
        vm.error_on_msgbox = on_msgbox == "error";
        Ok(PyVm { inner: vm })
    }

    /// Parse and execute *vba_code*, running the Sub named *macro_name*.
    fn run(&mut self, vba_code: &str, macro_name: &str) -> PyResult<()> {
        let prog = parser::parse(vba_code)
            .map_err(|e| PyErr::new::<pyo3::exceptions::PySyntaxError, _>(e.to_string()))?;
        self.inner
            .run_sub(&prog, macro_name)
            .map_err(PyErr::new::<pyo3::exceptions::PyRuntimeError, _>)
    }

    /// Write a value into a cell. ``row`` and ``col`` are 1-based (VBA convention).
    fn set_cell(&mut self, row: u32, col: u32, value: &Bound<'_, PyAny>) -> PyResult<()> {
        let v = py_to_variant(value)?;
        self.inner.cells_mut().insert(
            (row, col),
            CellContent {
                formula: None,
                value: v,
            },
        );
        Ok(())
    }

    /// Return the value of a cell (1-based row/col). Returns ``None`` for empty cells.
    fn get_cell(&self, py: Python<'_>, row: u32, col: u32) -> Py<PyAny> {
        variant_to_py(py, &self.inner.get_cell(row, col))
    }

    /// Return the active sheet's resolved number-format code for a cell (1-based
    /// row/col), e.g. ``"m/d/yyyy"`` for a date-formatted cell, or ``None`` for a cell
    /// with no format, the General format, or a sheet with no source-file styles.
    fn get_cell_number_format(&self, row: u32, col: u32) -> Option<&str> {
        self.inner.get_cell_number_format(row, col)
    }

    /// Return all non-empty cells as a dict: ``{(row, col): value}``.
    fn cells(&self, py: Python<'_>) -> PyResult<Py<PyAny>> {
        let dict = PyDict::new(py);
        for ((row, col), content) in self.inner.cells() {
            if !matches!(content.value, Variant::Empty) {
                let key = (*row, *col).into_pyobject(py)?.into_any().unbind();
                dict.set_item(key, variant_to_py(py, &content.value))?;
            }
        }
        Ok(dict.into_any().unbind())
    }

    /// Return all VBA variables as a dict: ``{name: value}``.
    fn variables(&self, py: Python<'_>) -> PyResult<Py<PyAny>> {
        let dict = PyDict::new(py);
        for (name, value) in &self.inner.variables {
            dict.set_item(name.as_str(), variant_to_py(py, value))?;
        }
        Ok(dict.into_any().unbind())
    }

    /// Store a formula string (e.g. ``"=SUM(A1:A3)"``) on a cell and evaluate it
    /// immediately against the current cell state.
    fn set_cell_formula(&mut self, row: u32, col: u32, formula: &str) -> PyResult<()> {
        self.inner
            .set_cell_formula(row, col, formula)
            .map_err(PyErr::new::<pyo3::exceptions::PyValueError, _>)
    }

    /// Re-evaluate all cells that have a stored formula.
    fn recalculate(&mut self) -> PyResult<()> {
        self.inner
            .recalculate_all()
            .map_err(PyErr::new::<pyo3::exceptions::PyRuntimeError, _>)
    }

    /// Set multiple cell formulas at once.
    /// ``formulas`` should be a dict mapping ``(row, col)`` tuples (1-based) to formula strings.
    fn set_cell_formula_batch(&mut self, formulas: &Bound<'_, PyDict>) -> PyResult<()> {
        for (key, val) in formulas.iter() {
            let (row, col): (u32, u32) = key.extract().map_err(|_| {
                PyErr::new::<pyo3::exceptions::PyTypeError, _>(
                    "keys must be (row, col) tuples of integers",
                )
            })?;
            let formula: String = val.extract().map_err(|_| {
                PyErr::new::<pyo3::exceptions::PyTypeError, _>("values must be formula strings")
            })?;
            self.inner
                .set_cell_formula(row, col, &formula)
                .map_err(PyErr::new::<pyo3::exceptions::PyValueError, _>)?;
        }
        Ok(())
    }

    /// Switch the active sheet. Creates the sheet if it does not exist.
    /// ``index`` (0-based) places a newly-created sheet at that position among the
    /// existing sheets instead of appending it at the end; ignored if ``name`` already
    /// exists, and clamped rather than erroring if it's past the current sheet count.
    #[pyo3(signature = (name, index = None))]
    fn set_sheet(&mut self, name: &str, index: Option<usize>) {
        self.inner.ensure_sheet_at(name, index);
        self.inner.active_sheet = name.to_lowercase();
    }

    /// Delete the sheet named ``name``. Raises ``ValueError`` if it doesn't exist.
    fn delete_sheet(&mut self, name: &str) -> PyResult<()> {
        self.inner
            .delete_sheet(name)
            .map_err(PyErr::new::<pyo3::exceptions::PyValueError, _>)
    }

    /// Rename a sheet.
    ///
    /// Every formula reference qualified with *old_name* (``='Old Name'!A1``),
    /// on ANY sheet in the workbook, is rewritten to name *new_name* instead —
    /// requoted/escaped as *new_name* itself requires, regardless of how the old
    /// reference was written. An unqualified reference is never touched, even on
    /// the renamed sheet itself (``=A1`` still means "this same sheet", whatever
    /// it's now called). A formula this parser can't parse at all (an external
    /// workbook reference, a 3D reference) is left completely untouched rather
    /// than partially rewritten. Does **not** rewrite ``<definedName>`` text that
    /// refers to this sheet by its old name — that mechanism stays out of scope;
    /// see docs/openpyxl-gap-audit.md.
    ///
    /// Parameters
    /// ----------
    /// old_name:
    ///     The sheet's current name (case-insensitive).
    /// new_name:
    ///     The new name. Renaming the active sheet is supported (it stays active
    ///     under the new name). Renaming a sheet to itself, or to a different
    ///     casing of its own name, succeeds — and still updates any formula
    ///     reference's qualifier to match the new casing.
    ///
    /// Raises ``ValueError`` if *old_name* doesn't exist, *new_name* is empty or
    /// whitespace-only, *new_name* (case-insensitively) already names a *different*
    /// existing sheet, or the sheet is protected.
    fn rename_sheet(&mut self, old_name: &str, new_name: &str) -> PyResult<()> {
        self.inner
            .rename_sheet(old_name, new_name)
            .map_err(PyErr::new::<pyo3::exceptions::PyValueError, _>)
    }

    /// Move a sheet to an absolute 0-based position among the workbook's sheets.
    ///
    /// Unlike openpyxl's ``Worksheet.move_sheet(offset)`` (a relative offset),
    /// *new_index* here is an absolute target position (0 = first), matching
    /// ``set_sheet``'s own ``index`` convention. Out-of-range values are clamped to
    /// the nearest end rather than raising. Does not check sheet protection --
    /// real Excel's per-sheet protection does not gate tab reordering.
    ///
    /// Raises ``ValueError`` if *name* doesn't exist.
    fn move_sheet(&mut self, name: &str, new_index: usize) -> PyResult<()> {
        self.inner
            .move_sheet(name, new_index)
            .map_err(PyErr::new::<pyo3::exceptions::PyValueError, _>)
    }

    /// Duplicate a sheet's cells, merges, hidden-row/col state, cell
    /// styles, cell number formats, whole-tab visibility state, and row
    /// heights/column widths into a brand-new sheet.
    ///
    /// Appended at the end of the workbook's sheets -- unlike openpyxl's own
    /// ``copy_worksheet`` (which places the copy immediately after the
    /// source), use :meth:`move_sheet` afterward if exact placement matters.
    /// Does not copy sheet protection status (the copy is always
    /// unprotected) and does not change the active sheet.
    ///
    /// Parameters
    /// ----------
    /// source_name:
    ///     The sheet to copy (case-insensitive).
    /// new_name:
    ///     The new sheet's name.
    ///
    /// Raises ``ValueError`` if *source_name* doesn't exist, or *new_name* is
    /// empty/whitespace-only or (case-insensitively) already names an
    /// existing sheet.
    fn copy_sheet(&mut self, source_name: &str, new_name: &str) -> PyResult<()> {
        self.inner
            .copy_sheet(source_name, new_name)
            .map_err(PyErr::new::<pyo3::exceptions::PyValueError, _>)
    }

    /// A sheet's whole-tab visibility: ``"visible"``, ``"hidden"``, or
    /// ``"veryHidden"`` -- matching openpyxl's own ``ws.sheet_state`` string
    /// vocabulary exactly, no translation needed.
    ///
    /// Read-only for now: there's no ``set_sheet_state`` yet, and a loaded
    /// file's hidden/veryHidden sheet currently reverts to visible on save
    /// regardless of what this reports (no writer support -- no real fixture
    /// in this project has a hidden sheet to validate the writer shape
    /// against; see docs/openpyxl-gap-audit.md).
    ///
    /// Parameters
    /// ----------
    /// name:
    ///     The sheet to query (case-insensitive).
    ///
    /// Raises ``ValueError`` if *name* doesn't exist. Unlike openpyxl's own
    /// property (which can't fail -- it just reads an attribute off an
    /// already-resolved ``Worksheet`` object), an unknown name is a real
    /// error here rather than silently returning ``"visible"``.
    fn sheet_state(&self, name: &str) -> PyResult<String> {
        self.inner
            .sheet_state(name)
            .map(|s| s.as_str().to_string())
            .map_err(PyErr::new::<pyo3::exceptions::PyValueError, _>)
    }

    /// Every workbook-level defined name as ``{name: raw_text}`` (e.g.
    /// ``{"MyRange": "Sheet1!$A$1:$A$3"}``).
    ///
    /// *raw_text* is the exact formula-text content, **not** resolved into a
    /// sheet+address — elixcee's formula engine has no cross-sheet reference
    /// syntax (``=Sheet2!A1``) to resolve it against. Sheet-scoped and
    /// workbook-scoped names are not distinguished; on a name collision
    /// across scopes, whichever the reader encounters last wins.
    ///
    /// Independent of VBA's own ``Range(addr).Name = "x"`` runtime names —
    /// this reads what the *originally-loaded source file on disk*
    /// declares, not the VM's in-memory named-range table. Re-read on every
    /// call rather than cached, but always from the path passed to
    /// ``load_workbook``/``Vm.load_workbook_file`` — a later
    /// ``save_workbook(new_path)`` does not change what this reports, so it
    /// will not reflect a save that dropped ``<definedNames>`` (e.g. after
    /// ``rename_sheet``/``move_sheet``) or was written to a different path.
    ///
    /// Returns ``{}`` if no workbook is loaded. Raises ``ValueError`` if a
    /// workbook WAS loaded but its source file is no longer readable (this
    /// method re-reads the file on every call rather than caching).
    fn defined_names(&self) -> PyResult<std::collections::HashMap<String, String>> {
        self.inner
            .defined_names()
            .map_err(PyErr::new::<pyo3::exceptions::PyValueError, _>)
    }

    /// Return the name of the currently active sheet.
    fn active_sheet(&self) -> &str {
        &self.inner.active_sheet
    }

    /// Return all sheet names.
    fn sheet_names(&self, py: Python<'_>) -> Py<PyAny> {
        let names = self.inner.sheet_names();
        names.into_pyobject(py).unwrap().into_any().unbind()
    }

    /// Return all non-empty cells in a specific sheet as ``{(row, col): value}``.
    fn get_sheet(&self, py: Python<'_>, name: &str) -> PyResult<Py<PyAny>> {
        let dict = PyDict::new(py);
        if let Some(sheet) = self.inner.get_sheet_cells(name) {
            for ((row, col), content) in sheet {
                if !matches!(content.value, Variant::Empty) {
                    let key = (*row, *col).into_pyobject(py)?.into_any().unbind();
                    dict.set_item(key, variant_to_py(py, &content.value))?;
                }
            }
        }
        Ok(dict.into_any().unbind())
    }

    /// Save all sheets to an .xlsx file. ``path`` should end with ``.xlsx``.
    fn save_workbook(&self, path: &str) -> PyResult<()> {
        save_workbook_impl(&self.inner, path).map_err(PyErr::new::<pyo3::exceptions::PyIOError, _>)
    }

    /// Return the active sheet's non-empty cells as a **pandas DataFrame**.
    ///
    /// Row indices and column indices are 1-based integers (matching VBA / Excel
    /// convention).  The DataFrame index is the row number; columns are column
    /// numbers.  Empty cells are represented as ``None`` (``pd.NA``).
    ///
    /// Raises ``ImportError`` if pandas is not installed.
    fn cells_df(&self, py: Python<'_>) -> PyResult<Py<PyAny>> {
        let pd = py.import("pandas").map_err(|_| {
            PyErr::new::<pyo3::exceptions::PyImportError, _>(
                "pandas is required for cells_df(); install it with: pip install pandas",
            )
        })?;

        let cells = self.inner.cells();
        if cells.is_empty() {
            return pd
                .getattr("DataFrame")?
                .call0()
                .map(|df| df.into_any().unbind());
        }

        let max_row = cells.keys().map(|(r, _)| *r).max().unwrap_or(1);
        let max_col = cells.keys().map(|(_, c)| *c).max().unwrap_or(1);

        // Build a list-of-lists (row-major), None for missing cells.
        let none = py.None();
        let rows_list = pyo3::types::PyList::empty(py);
        for r in 1..=max_row {
            let row_list = pyo3::types::PyList::empty(py);
            for c in 1..=max_col {
                match cells.get(&(r, c)) {
                    Some(cell) if !matches!(cell.value, Variant::Empty) => {
                        row_list.append(variant_to_py(py, &cell.value))?;
                    }
                    _ => row_list.append(&none)?,
                }
            }
            rows_list.append(row_list)?;
        }

        let col_index: Vec<u32> = (1..=max_col).collect();
        let row_index: Vec<u32> = (1..=max_row).collect();

        let kwargs = PyDict::new(py);
        kwargs.set_item("columns", col_index)?;
        kwargs.set_item("index", row_index)?;
        pd.getattr("DataFrame")?
            .call((rows_list,), Some(&kwargs))
            .map(|df| df.into_any().unbind())
    }

    /// Read a rectangular range (e.g. ``"A1:C5"``), 1-based A1 notation.
    ///
    /// Returns a row-major nested list, ``None`` for empty cells — same
    /// per-cell typing as ``get_cell``. Multi-area addresses (``"A1:B2,D1:E2"``)
    /// and malformed/reversed addresses raise ``ValueError``.
    ///
    /// Parameters
    /// ----------
    /// addr:
    ///     A single-area A1 range, e.g. ``"A1:C5"`` or a bare cell like ``"B2"``.
    /// sheet:
    ///     Sheet to read from. Defaults to the active sheet; does **not**
    ///     change the active sheet when given.
    #[pyo3(signature = (addr, sheet = None))]
    fn get_range(&self, py: Python<'_>, addr: &str, sheet: Option<&str>) -> PyResult<Py<PyAny>> {
        let (start, end) =
            validate_range_addr(addr).map_err(PyErr::new::<pyo3::exceptions::PyValueError, _>)?;
        let key = self
            .inner
            .resolve_sheet_key(sheet)
            .map_err(PyErr::new::<pyo3::exceptions::PyValueError, _>)?;
        let grid = self.inner.read_rect(&key, start.0, start.1, end.0, end.1);
        grid_to_py(py, &grid)
    }

    /// Write a rectangular range (e.g. ``"A1:C2"``), 1-based A1 notation.
    ///
    /// *values* must be a strictly rectangular (non-ragged) nested sequence
    /// whose shape exactly matches *addr*'s row×col shape, or ``ValueError``
    /// is raised naming both the expected and actual shape. ``None`` in the
    /// input means an empty cell. A string value starting with ``"="`` is
    /// stored literally, never promoted to a formula — use
    /// ``set_cell_formula``/``set_cell_formula_batch`` for that. Every value
    /// is converted and the shape is checked **before** any cell is
    /// touched — a validation failure leaves every existing cell unchanged.
    ///
    /// Writing into a non-anchor cell of a merged range, or into a protected
    /// sheet, is **not** blocked — this matches ``set_cell``'s existing
    /// behavior (see docs/openpyxl-gap-audit.md for why).
    ///
    /// Parameters
    /// ----------
    /// addr:
    ///     A single-area A1 range, e.g. ``"A1:C2"``.
    /// values:
    ///     A rectangular nested sequence matching *addr*'s shape.
    /// sheet:
    ///     Sheet to write to. Defaults to the active sheet; does **not**
    ///     change the active sheet when given.
    #[pyo3(signature = (addr, values, sheet = None))]
    fn set_range(
        &mut self,
        addr: &str,
        values: &Bound<'_, PyAny>,
        sheet: Option<&str>,
    ) -> PyResult<()> {
        let (start, end) =
            validate_range_addr(addr).map_err(PyErr::new::<pyo3::exceptions::PyValueError, _>)?;
        let key = self
            .inner
            .resolve_sheet_key(sheet)
            .map_err(PyErr::new::<pyo3::exceptions::PyValueError, _>)?;

        // The whole grid is converted into a scratch buffer first -- nothing
        // in `self.inner` is touched until every value has converted
        // successfully AND the shape has been confirmed exact. This is what
        // makes "no partial write on validation failure" structurally true,
        // not just tested.
        let mut grid: Vec<Vec<Variant>> = Vec::new();
        let mut row_lens: Vec<usize> = Vec::new();
        for row_obj in values.try_iter()? {
            let mut row: Vec<Variant> = Vec::new();
            for cell_obj in row_obj?.try_iter()? {
                row.push(py_to_variant(&cell_obj?)?);
            }
            row_lens.push(row.len());
            grid.push(row);
        }
        let expected = (end.0 - start.0 + 1, end.1 - start.1 + 1);
        check_grid_shape(expected, &row_lens)
            .map_err(PyErr::new::<pyo3::exceptions::PyValueError, _>)?;

        self.inner.write_rect(&key, start, &grid);
        Ok(())
    }

    /// Write one row just past the sheet's used range (row 1 if the sheet is
    /// empty/all-empty; uses the true max used row, so it's correct on a
    /// sparse sheet). Returns the 1-based row number written.
    ///
    /// Same validate-then-commit and active-sheet-preservation guarantees as
    /// ``set_range``. Raises ``ValueError`` if *values* is empty.
    ///
    /// Parameters
    /// ----------
    /// values:
    ///     The row's values, written starting at column 1.
    /// sheet:
    ///     Sheet to append to. Defaults to the active sheet; does **not**
    ///     change the active sheet when given.
    #[pyo3(signature = (values, sheet = None))]
    fn append_row(&mut self, values: &Bound<'_, PyAny>, sheet: Option<&str>) -> PyResult<u32> {
        let mut row: Vec<Variant> = Vec::new();
        for item in values.try_iter()? {
            row.push(py_to_variant(&item?)?);
        }
        if row.is_empty() {
            return Err(PyErr::new::<pyo3::exceptions::PyValueError, _>(
                "append_row: values must not be empty",
            ));
        }
        let key = self
            .inner
            .resolve_sheet_key(sheet)
            .map_err(PyErr::new::<pyo3::exceptions::PyValueError, _>)?;
        let target_row = self.inner.next_append_row(&key);
        self.inner.write_rect(&key, (target_row, 1), &[row]);
        Ok(target_row)
    }

    /// Values-only iteration over a rectangular region, 1-based bounds.
    ///
    /// ``max_row``/``max_col`` default to the sheet's used range; on a sheet
    /// with no non-empty cells at all **and** no explicit ``max_row``,
    /// returns ``[]`` rather than one row of ``None``\ s. Returns plain
    /// nested lists — this does **not** claim openpyxl ``Cell``-object
    /// compatibility (no ``.value``/``.style``/etc attached, just the values).
    ///
    /// Parameters
    /// ----------
    /// min_row, min_col:
    ///     1-based lower bounds (default 1).
    /// max_row, max_col:
    ///     1-based upper bounds. Default to the sheet's used range.
    /// sheet:
    ///     Sheet to read from. Defaults to the active sheet; does **not**
    ///     change the active sheet when given.
    #[pyo3(signature = (min_row = 1, max_row = None, min_col = 1, max_col = None, sheet = None))]
    #[allow(clippy::too_many_arguments)]
    fn iter_rows(
        &self,
        py: Python<'_>,
        min_row: u32,
        max_row: Option<u32>,
        min_col: u32,
        max_col: Option<u32>,
        sheet: Option<&str>,
    ) -> PyResult<Py<PyAny>> {
        if min_row == 0 || min_col == 0 || max_row == Some(0) || max_col == Some(0) {
            return Err(PyErr::new::<pyo3::exceptions::PyValueError, _>(
                "row/column numbers must be >= 1",
            ));
        }
        let key = self
            .inner
            .resolve_sheet_key(sheet)
            .map_err(PyErr::new::<pyo3::exceptions::PyValueError, _>)?;
        let grid = self
            .inner
            .iter_rows_values(&key, min_row, max_row, min_col, max_col);
        grid_to_py(py, &grid)
    }

    /// Values-only, column-major iteration over a rectangular region —
    /// the transposed sibling of ``iter_rows``. Each returned inner list is
    /// one column's values, top to bottom.
    ///
    /// ``max_row``/``max_col`` default to the sheet's used range; on a sheet
    /// with no non-empty cells at all **and** no explicit ``max_col``,
    /// returns ``[]`` rather than one column of ``None``\ s. Returns plain
    /// nested lists — this does **not** claim openpyxl ``Cell``-object
    /// compatibility (no ``.value``/``.style``/etc attached, just the values).
    ///
    /// Parameters
    /// ----------
    /// min_row, min_col:
    ///     1-based lower bounds (default 1).
    /// max_row, max_col:
    ///     1-based upper bounds. Default to the sheet's used range.
    /// sheet:
    ///     Sheet to read from. Defaults to the active sheet; does **not**
    ///     change the active sheet when given.
    #[pyo3(signature = (min_row = 1, max_row = None, min_col = 1, max_col = None, sheet = None))]
    #[allow(clippy::too_many_arguments)]
    fn iter_cols(
        &self,
        py: Python<'_>,
        min_row: u32,
        max_row: Option<u32>,
        min_col: u32,
        max_col: Option<u32>,
        sheet: Option<&str>,
    ) -> PyResult<Py<PyAny>> {
        if min_row == 0 || min_col == 0 || max_row == Some(0) || max_col == Some(0) {
            return Err(PyErr::new::<pyo3::exceptions::PyValueError, _>(
                "row/column numbers must be >= 1",
            ));
        }
        let key = self
            .inner
            .resolve_sheet_key(sheet)
            .map_err(PyErr::new::<pyo3::exceptions::PyValueError, _>)?;
        let grid = self
            .inner
            .iter_cols_values(&key, min_row, max_row, min_col, max_col);
        grid_to_py(py, &grid)
    }

    /// Highest used row number, or ``None`` for a sheet with zero non-empty
    /// cells (never ``0``).
    #[pyo3(signature = (sheet = None))]
    fn max_row(&self, sheet: Option<&str>) -> PyResult<Option<u32>> {
        let key = self
            .inner
            .resolve_sheet_key(sheet)
            .map_err(PyErr::new::<pyo3::exceptions::PyValueError, _>)?;
        Ok(self.inner.sheet_used_range(&key).map(|(_, (r2, _))| r2))
    }

    /// Highest used column number, or ``None`` for a sheet with zero
    /// non-empty cells (never ``0``).
    #[pyo3(signature = (sheet = None))]
    fn max_column(&self, sheet: Option<&str>) -> PyResult<Option<u32>> {
        let key = self
            .inner
            .resolve_sheet_key(sheet)
            .map_err(PyErr::new::<pyo3::exceptions::PyValueError, _>)?;
        Ok(self.inner.sheet_used_range(&key).map(|(_, (_, c2))| c2))
    }

    /// The used range as an A1-style string (e.g. ``"B2:D10"``), or ``None``
    /// for a sheet with zero non-empty cells (never ``"A1:A1"``).
    ///
    /// Min-anchored, not A1-anchored: if the only populated cell is C3, this
    /// returns ``"C3:C3"``, not ``"A1:C3"``. Always includes the ``:`` even
    /// for a single-cell range.
    #[pyo3(signature = (sheet = None))]
    fn calculate_dimension(&self, sheet: Option<&str>) -> PyResult<Option<String>> {
        let key = self
            .inner
            .resolve_sheet_key(sheet)
            .map_err(PyErr::new::<pyo3::exceptions::PyValueError, _>)?;
        Ok(self
            .inner
            .sheet_used_range(&key)
            .map(|((r1, c1), (r2, c2))| {
                format!(
                    "{}{}:{}{}",
                    xlsx_col_letters(c1),
                    r1,
                    xlsx_col_letters(c2),
                    r2
                )
            }))
    }

    /// Insert *amount* blank rows before 1-based row *idx*, shifting *idx* and
    /// everything below it down. Mirrors openpyxl's
    /// ``Worksheet.insert_rows(idx, amount=1)`` naming/value semantics.
    ///
    /// Formula cell-references naming *sheet*, workbook-wide, are updated to
    /// follow the shift — precisely:
    ///
    /// - an unqualified reference (``=A10``) shifts when its own cell lives ON
    ///   *sheet* (e.g. ``=A10`` becomes ``=A12`` when 2 rows are inserted
    ///   before row 5 on the sheet that formula is IN)
    /// - a sheet-qualified reference (``=Sheet2!A10``, ``='Sales 2026'!A10``)
    ///   shifts whenever it NAMES *sheet*, no matter which sheet hosts the
    ///   formula itself — a reference on Sheet3 to ``Sheet2!A10`` shifts when
    ///   *sheet* is Sheet2, even though it isn't Sheet3
    /// - a reference that lands inside a *deleted* band becomes ``#REF!`` on
    ///   :meth:`delete_rows` (the sheet qualifier, if any, is preserved —
    ///   only the coordinate becomes ``#REF!``)
    /// - **cross-sheet formula evaluation is a separate, still-unsupported
    ///   concern** — a qualified reference now parses and its coordinate
    ///   rewrites correctly, but the formula's cached VALUE is never
    ///   recomputed against another sheet's cells; :meth:`set_cell_formula`
    ///   still raises for any formula containing one, since it requires
    ///   evaluating immediately
    /// - a formula that couldn't be parsed at all (external workbook
    ///   references like ``[Book2.xlsx]Sheet1!A1``, 3D references like
    ///   ``Sheet1:Sheet3!A1``, and anything else outside same-workbook `A1`
    ///   syntax) is left exactly as-is rather than partially rewritten
    ///
    /// Also shifts merged ranges, hidden-row markers, per-cell
    /// styles/number formats, and row heights on *sheet* the same way real
    /// Excel does; range-move (:meth:`move_range`) deliberately does not
    /// carry hidden-row markers or row heights along, since those belong to
    /// the row itself, not to the moved cell content. Does not recompute any
    /// cached formula value — call :meth:`recalculate` afterwards if you need
    /// fresh values
    /// (note it recalculates the *active* sheet, so switch sheets first if
    /// the sheet you need refreshed isn't already active; a cross-sheet
    /// formula is skipped by recalculation entirely, same as one that fails
    /// to parse). See docs/openpyxl-gap-audit.md and ROADMAP.md's known gaps.
    ///
    /// Parameters
    /// ----------
    /// idx:
    ///     1-based row number to insert before.
    /// amount:
    ///     Number of rows to insert (default 1).
    /// sheet:
    ///     Sheet to modify. Defaults to the active sheet; never changes which
    ///     sheet is active.
    ///
    /// Raises ``ValueError`` if *idx*/*amount* is 0 or exceeds Excel's own grid
    /// limit (1,048,576 rows), or *sheet* is unknown.
    #[pyo3(signature = (idx, amount = 1, sheet = None))]
    fn insert_rows(&mut self, idx: u32, amount: u32, sheet: Option<&str>) -> PyResult<()> {
        const MAX_ROW: u32 = 1_048_576;
        if idx == 0 || amount == 0 || idx > MAX_ROW || amount > MAX_ROW {
            return Err(PyErr::new::<pyo3::exceptions::PyValueError, _>(
                "idx and amount must be between 1 and 1_048_576",
            ));
        }
        let key = self
            .inner
            .resolve_sheet_key(sheet)
            .map_err(PyErr::new::<pyo3::exceptions::PyValueError, _>)?;
        self.inner.insert_rows_on_sheet(&key, idx, amount);
        Ok(())
    }

    /// Delete *amount* rows starting at 1-based row *idx*, shifting everything
    /// below the deleted band up. Mirrors openpyxl's
    /// ``Worksheet.delete_rows(idx, amount=1)`` naming/value semantics.
    ///
    /// Formula references are updated workbook-wide the same way
    /// :meth:`insert_rows` documents in full — unqualified references on
    /// *sheet* itself, and any sheet-qualified reference naming *sheet* from
    /// anywhere in the workbook, both shift; a reference into the deleted
    /// band becomes ``#REF!``; cross-sheet formula evaluation remains
    /// unsupported. Also shifts merged ranges, hidden-row markers, per-cell
    /// styles/number formats, and row heights the same way :meth:`insert_rows`
    /// documents in full.
    ///
    /// Parameters
    /// ----------
    /// idx:
    ///     1-based row number to start deleting from.
    /// amount:
    ///     Number of rows to delete (default 1).
    /// sheet:
    ///     Sheet to modify. Defaults to the active sheet; never changes which
    ///     sheet is active.
    ///
    /// Raises ``ValueError`` if *idx*/*amount* is 0 or exceeds Excel's own grid
    /// limit (1,048,576 rows), or *sheet* is unknown.
    #[pyo3(signature = (idx, amount = 1, sheet = None))]
    fn delete_rows(&mut self, idx: u32, amount: u32, sheet: Option<&str>) -> PyResult<()> {
        const MAX_ROW: u32 = 1_048_576;
        if idx == 0 || amount == 0 || idx > MAX_ROW || amount > MAX_ROW {
            return Err(PyErr::new::<pyo3::exceptions::PyValueError, _>(
                "idx and amount must be between 1 and 1_048_576",
            ));
        }
        let key = self
            .inner
            .resolve_sheet_key(sheet)
            .map_err(PyErr::new::<pyo3::exceptions::PyValueError, _>)?;
        self.inner.delete_rows_on_sheet(&key, idx, amount);
        Ok(())
    }

    /// Insert *amount* blank columns before 1-based column *idx*, shifting *idx*
    /// and everything to its right, right. Mirrors openpyxl's
    /// ``Worksheet.insert_cols(idx, amount=1)`` naming/value semantics.
    ///
    /// Formula references are updated workbook-wide the same way
    /// :meth:`insert_rows` documents in full, on the column axis. Also shifts
    /// merged ranges, hidden-column markers, per-cell styles/number formats,
    /// and column widths on *sheet* the same way real Excel does; range-move
    /// (:meth:`move_range`) deliberately does not carry hidden-column markers
    /// or column widths along, since those belong to the column itself, not
    /// to the moved cell content.
    ///
    /// Parameters
    /// ----------
    /// idx:
    ///     1-based column number to insert before.
    /// amount:
    ///     Number of columns to insert (default 1).
    /// sheet:
    ///     Sheet to modify. Defaults to the active sheet; never changes which
    ///     sheet is active.
    ///
    /// Raises ``ValueError`` if *idx*/*amount* is 0 or exceeds Excel's own grid
    /// limit (16,384 columns, i.e. ``XFD``), or *sheet* is unknown.
    #[pyo3(signature = (idx, amount = 1, sheet = None))]
    fn insert_cols(&mut self, idx: u32, amount: u32, sheet: Option<&str>) -> PyResult<()> {
        const MAX_COL: u32 = 16_384;
        if idx == 0 || amount == 0 || idx > MAX_COL || amount > MAX_COL {
            return Err(PyErr::new::<pyo3::exceptions::PyValueError, _>(
                "idx and amount must be between 1 and 16_384",
            ));
        }
        let key = self
            .inner
            .resolve_sheet_key(sheet)
            .map_err(PyErr::new::<pyo3::exceptions::PyValueError, _>)?;
        self.inner.insert_cols_on_sheet(&key, idx, amount);
        Ok(())
    }

    /// Delete *amount* columns starting at 1-based column *idx*, shifting
    /// everything to the right of the deleted band left. Mirrors openpyxl's
    /// ``Worksheet.delete_cols(idx, amount=1)`` naming/value semantics.
    ///
    /// Formula references are updated workbook-wide the same way
    /// :meth:`insert_rows` documents in full, on the column axis (a reference
    /// into the deleted band becomes ``#REF!``). Also shifts merged ranges,
    /// hidden-column markers, per-cell styles/number formats, and column
    /// widths the same way :meth:`insert_cols` documents in full.
    ///
    /// Parameters
    /// ----------
    /// idx:
    ///     1-based column number to start deleting from.
    /// amount:
    ///     Number of columns to delete (default 1).
    /// sheet:
    ///     Sheet to modify. Defaults to the active sheet; never changes which
    ///     sheet is active.
    ///
    /// Raises ``ValueError`` if *idx*/*amount* is 0 or exceeds Excel's own grid
    /// limit (16,384 columns, i.e. ``XFD``), or *sheet* is unknown.
    #[pyo3(signature = (idx, amount = 1, sheet = None))]
    fn delete_cols(&mut self, idx: u32, amount: u32, sheet: Option<&str>) -> PyResult<()> {
        const MAX_COL: u32 = 16_384;
        if idx == 0 || amount == 0 || idx > MAX_COL || amount > MAX_COL {
            return Err(PyErr::new::<pyo3::exceptions::PyValueError, _>(
                "idx and amount must be between 1 and 16_384",
            ));
        }
        let key = self
            .inner
            .resolve_sheet_key(sheet)
            .map_err(PyErr::new::<pyo3::exceptions::PyValueError, _>)?;
        self.inner.delete_cols_on_sheet(&key, idx, amount);
        Ok(())
    }

    /// Return every merged range on a sheet as A1-style strings (e.g.
    /// ``["B1:C1"]``).
    ///
    /// Order matches source-file/insertion order (a stable list, never
    /// re-sorted) — do not assume alphabetical or row-major order.
    ///
    /// Raises ``ValueError`` if *sheet* is unknown.
    #[pyo3(signature = (sheet = None))]
    fn merged_cells(&self, sheet: Option<&str>) -> PyResult<Vec<String>> {
        let key = self
            .inner
            .resolve_sheet_key(sheet)
            .map_err(PyErr::new::<pyo3::exceptions::PyValueError, _>)?;
        Ok(self
            .inner
            .merged_ranges
            .get(&key)
            .map(|ranges| ranges.iter().map(merge_rect_to_a1).collect())
            .unwrap_or_default())
    }

    /// Every table defined on *sheet* (defaults to the active sheet), read-only
    /// (0.16.0-A1). Each table is a dict:
    ///
    /// ``{"name": ..., "display_name": ..., "ref": "A1:C4", "header_row_count": 1,
    /// "totals_row_count": 0, "totals_row_shown": False, "style_name": ...,
    /// "auto_filter_ref": "A1:C4" or None, "columns": [...]}``
    ///
    /// Each column is ``{"id": ..., "name": ..., "totals_row_function": ...,
    /// "totals_row_label": ..., "calculated_column_formula": ...}`` — string fields
    /// are ``None`` when the attribute is absent in the source file.
    ///
    /// Read-only: there's no create/edit/delete table API yet (separate future
    /// phases). Structured references (``Table1[@Qty]``) are entirely out of
    /// scope — ``calculated_column_formula`` is the raw, unparsed formula text,
    /// never evaluated. An unmodified table survives every save unchanged
    /// regardless of whether this method is ever called.
    ///
    /// Raises ``ValueError`` if *sheet* is unknown.
    #[pyo3(signature = (sheet = None))]
    fn tables(&self, py: Python<'_>, sheet: Option<&str>) -> PyResult<Py<PyAny>> {
        let key = self
            .inner
            .resolve_sheet_key(sheet)
            .map_err(PyErr::new::<pyo3::exceptions::PyValueError, _>)?;
        let list = PyList::empty(py);
        for t in self.inner.tables.get(&key).into_iter().flatten() {
            let dict = PyDict::new(py);
            dict.set_item("name", &t.name)?;
            dict.set_item("display_name", &t.display_name)?;
            dict.set_item("ref", merge_rect_to_a1(&t.ref_range))?;
            dict.set_item("header_row_count", t.header_row_count)?;
            dict.set_item("totals_row_count", t.totals_row_count)?;
            dict.set_item("totals_row_shown", t.totals_row_shown)?;
            dict.set_item("style_name", t.style_name.as_deref())?;
            dict.set_item(
                "auto_filter_ref",
                t.auto_filter_ref.as_ref().map(merge_rect_to_a1),
            )?;
            let cols = PyList::empty(py);
            for c in &t.columns {
                let cd = PyDict::new(py);
                cd.set_item("id", c.id.as_deref())?;
                cd.set_item("name", &c.name)?;
                cd.set_item("totals_row_function", c.totals_row_function.as_deref())?;
                cd.set_item("totals_row_label", c.totals_row_label.as_deref())?;
                cd.set_item(
                    "calculated_column_formula",
                    c.calculated_column_formula.as_deref(),
                )?;
                cols.append(cd)?;
            }
            dict.set_item("columns", cols)?;
            list.append(dict)?;
        }
        Ok(list.into_any().unbind())
    }

    /// Creates a merge over *addr*. Rejects a single-cell address (nothing
    /// would actually be merged) and rejects a merge that would overlap an
    /// existing one on the same sheet — two overlapping merges is invalid
    /// OOXML, not just a fidelity gap.
    ///
    /// Does **not** touch cell values — whatever is in the covered cells (if
    /// anything) stays exactly as it was.
    ///
    /// Raises ``ValueError`` on a bad, oversized, or single-cell address, an
    /// overlapping merge, or an unknown *sheet* name.
    #[pyo3(signature = (addr, sheet = None))]
    fn merge_cells(&mut self, addr: &str, sheet: Option<&str>) -> PyResult<()> {
        const MAX_ROW: u32 = 1_048_576;
        const MAX_COL: u32 = 16_384;
        let ((r1, c1), (r2, c2)) =
            validate_range_addr(addr).map_err(PyErr::new::<pyo3::exceptions::PyValueError, _>)?;
        // Same ceiling as sort_range -- an unbounded merge address writes a
        // real <mergeCell> spanning the whole sheet into the saved file.
        if r2 > MAX_ROW || c2 > MAX_COL {
            return Err(PyErr::new::<pyo3::exceptions::PyValueError, _>(format!(
                "range exceeds sheet bounds (max row {MAX_ROW}, max col {MAX_COL}), got row {r2}, col {c2}"
            )));
        }
        let key = self
            .inner
            .resolve_sheet_key(sheet)
            .map_err(PyErr::new::<pyo3::exceptions::PyValueError, _>)?;
        self.inner
            .merge_cells(&key, r1, c1, r2, c2)
            .map_err(PyErr::new::<pyo3::exceptions::PyValueError, _>)
    }

    /// Removes a merge whose range exactly matches *addr*. An inexact/
    /// partial match is rejected rather than silently no-opping.
    ///
    /// Raises ``ValueError`` on a bad or oversized address, no exact match,
    /// or an unknown *sheet* name.
    #[pyo3(signature = (addr, sheet = None))]
    fn unmerge_cells(&mut self, addr: &str, sheet: Option<&str>) -> PyResult<()> {
        const MAX_ROW: u32 = 1_048_576;
        const MAX_COL: u32 = 16_384;
        let ((r1, c1), (r2, c2)) =
            validate_range_addr(addr).map_err(PyErr::new::<pyo3::exceptions::PyValueError, _>)?;
        if r2 > MAX_ROW || c2 > MAX_COL {
            return Err(PyErr::new::<pyo3::exceptions::PyValueError, _>(format!(
                "range exceeds sheet bounds (max row {MAX_ROW}, max col {MAX_COL}), got row {r2}, col {c2}"
            )));
        }
        let key = self
            .inner
            .resolve_sheet_key(sheet)
            .map_err(PyErr::new::<pyo3::exceptions::PyValueError, _>)?;
        self.inner
            .unmerge_cells(&key, r1, c1, r2, c2)
            .map_err(PyErr::new::<pyo3::exceptions::PyValueError, _>)
    }

    /// Moves the range at *addr* by *rows* rows and *cols* columns (either
    /// may be negative, moving up/left) — 0.14.0-A4, same-sheet only.
    ///
    /// A reference (unqualified, or qualified naming this sheet) whose
    /// target cell falls inside *addr* follows the move to its new
    /// location, regardless of whether the referencing formula's own cell
    /// is inside or outside the moved range — matches real Excel's own
    /// "reference tracks cell identity" behavior, not a relative-offset
    /// rule. A range reference (e.g. inside a ``SUM``) with exactly one
    /// corner inside the moved area has unconfirmed behavior in real Excel
    /// (see ``internal_docs/range-move-0.14.0-a4-design.md``); rather than
    /// guess, the **whole** move is rejected and nothing changes.
    ///
    /// Source and destination may overlap — handled atomically, not with a
    /// naive cell-by-cell copy. A pre-existing cell at the destination that
    /// isn't itself part of the move is overwritten, matching real Excel's
    /// own paste behavior.
    ///
    /// Merged ranges, per-cell styles, and per-cell number formats inside
    /// *addr* move with it, using the same inside/outside/ambiguous-corner
    /// rules as formula references above (a merge with exactly one corner
    /// inside the moved area is rejected the same way an ambiguous range
    /// reference is). Does **not** move hidden-row/column markers, row
    /// heights, or column widths — those belong to the row/column itself,
    /// not to the moved cell content, matching real Excel's own paste
    /// behavior; see ROADMAP.md's 0.14.0-B. Cached values are left stale,
    /// same as every other structural edit — call recalculation yourself if
    /// you need fresh values.
    ///
    /// Cross-sheet moves are not supported this round (a qualified
    /// reference naming a *different* sheet is always left untouched, even
    /// if the mechanics above sound like they'd otherwise apply).
    ///
    /// Raises ``ValueError`` on a bad/oversized address, a destination that
    /// would fall outside the sheet, an unknown *sheet*, or the ambiguous
    /// range-reference case above.
    #[pyo3(signature = (addr, rows = 0, cols = 0, sheet = None))]
    fn move_range(
        &mut self,
        addr: &str,
        rows: i64,
        cols: i64,
        sheet: Option<&str>,
    ) -> PyResult<()> {
        const MAX_ROW: u32 = 1_048_576;
        const MAX_COL: u32 = 16_384;
        let ((r1, c1), (r2, c2)) =
            validate_range_addr(addr).map_err(PyErr::new::<pyo3::exceptions::PyValueError, _>)?;
        if r2 > MAX_ROW || c2 > MAX_COL {
            return Err(PyErr::new::<pyo3::exceptions::PyValueError, _>(format!(
                "range exceeds sheet bounds (max row {MAX_ROW}, max col {MAX_COL}), got row {r2}, col {c2}"
            )));
        }
        let dest_r1 = r1 as i64 + rows;
        let dest_c1 = c1 as i64 + cols;
        let dest_r2 = r2 as i64 + rows;
        let dest_c2 = c2 as i64 + cols;
        if dest_r1 < 1 || dest_c1 < 1 {
            return Err(PyErr::new::<pyo3::exceptions::PyValueError, _>(
                "move destination is out of bounds (row/col must stay >= 1)".to_string(),
            ));
        }
        if dest_r2 > MAX_ROW as i64 || dest_c2 > MAX_COL as i64 {
            return Err(PyErr::new::<pyo3::exceptions::PyValueError, _>(format!(
                "move destination exceeds sheet bounds (max row {MAX_ROW}, max col {MAX_COL})"
            )));
        }
        let key = self
            .inner
            .resolve_sheet_key(sheet)
            .map_err(PyErr::new::<pyo3::exceptions::PyValueError, _>)?;
        self.inner
            .move_range_on_sheet(
                &key,
                formula::MoveRect { r1, c1, r2, c2 },
                dest_r1 as u32,
                dest_c1 as u32,
            )
            .map_err(PyErr::new::<pyo3::exceptions::PyValueError, _>)
    }

    /// Sets every cell in *addr* to *format_code* (e.g. ``"#,##0.00"``, ``"m/d/yyyy"``,
    /// or a genuinely custom code like ``"0.00\"kg\""``) -- 0.15.0-A.
    ///
    /// An existing built-in (ids 0-49) or this file's own existing custom format is
    /// reused if *format_code* already matches one exactly; otherwise a new custom
    /// ``<numFmt>`` is minted on save. Every other style attribute on each cell (font,
    /// fill, border, alignment, protection) is preserved unchanged -- setting those is
    /// 0.15.0-B, not this method. Cells that already share a style index with cells
    /// OUTSIDE *addr* never have that shared index mutated in place: each touched cell
    /// gets a fresh or reused index of its own, so untouched cells elsewhere are
    /// unaffected, even if they happened to look identical before this call.
    ///
    /// Takes effect immediately for reads (:meth:`get_cell_number_format` reflects the
    /// new value right away), but the underlying ``xl/styles.xml`` write is deferred to
    /// :meth:`save_workbook` -- calling this without ever saving has no on-disk effect.
    ///
    /// A cell that never receives a value (via :meth:`set_cell` or otherwise) by the
    /// time you save has no persisted effect on disk, even though a read right after
    /// this call still reports the new format -- the writer only emits a ``<c>``
    /// element for a cell with actual content, a pre-existing limitation that predates
    /// this method (a loaded file's own genuinely empty, pre-formatted cell is dropped
    /// the same way on any save). See ROADMAP.md's known gaps.
    ///
    /// Parameters
    /// ----------
    /// addr:
    ///     A single-area A1 range, e.g. ``"B2:B100"``, or a single cell.
    /// format_code:
    ///     The number-format code to apply.
    /// sheet:
    ///     Sheet to modify. Defaults to the active sheet; does **not** change the
    ///     active sheet when given.
    ///
    /// Raises ``ValueError`` on a bad/oversized address or an unknown *sheet*.
    #[pyo3(signature = (addr, format_code, sheet = None))]
    fn set_number_format(
        &mut self,
        addr: &str,
        format_code: &str,
        sheet: Option<&str>,
    ) -> PyResult<()> {
        let ((r1, c1), (r2, c2)) =
            validate_range_addr(addr).map_err(PyErr::new::<pyo3::exceptions::PyValueError, _>)?;
        let key = self
            .inner
            .resolve_sheet_key(sheet)
            .map_err(PyErr::new::<pyo3::exceptions::PyValueError, _>)?;
        self.inner
            .set_number_format_on_sheet(&key, r1, c1, r2, c2, format_code);
        Ok(())
    }

    /// Sets font/fill/border/alignment/protection properties, and/or an existing named
    /// style, on every cell in *addr* -- 0.15.0-B/C1. At least one of *font*/*fill*/
    /// *border*/*alignment*/*protection*/*named_style* must be given. Each of the five
    /// dict arguments carries only the properties you want to change; anything a dict
    /// doesn't mention is left exactly as that cell's current style already has it --
    /// e.g. ``font={"bold": True}`` on a cell with an existing colored, sized font
    /// changes only boldness, the color/size survive untouched. Calling this more than
    /// once on the same cell before a save accumulates: a ``fill=...`` call after an
    /// earlier ``font=...`` call keeps both.
    ///
    /// ``font`` keys: ``bold``, ``italic``, ``underline``, ``strike`` (all ``bool``),
    /// ``size`` (number), ``color`` (hex string), ``name`` (string).
    ///
    /// ``fill`` keys: ``type`` (only ``"solid"`` is implemented — gradient/pattern fills
    /// are a future milestone), ``color`` (hex string, required). Unlike the other four,
    /// a fill request REPLACES the cell's whole fill record rather than merging onto it
    /// (matches real Excel's own single-color fill picker).
    ///
    /// ``border`` keys: any of ``left``/``right``/``top``/``bottom``/``diagonal``, each
    /// itself a dict with ``style`` (string, e.g. ``"thin"``/``"thick"``) and/or
    /// ``color`` (hex string). Only the sides you name are touched.
    ///
    /// ``alignment`` keys: ``horizontal``, ``vertical`` (strings, e.g. ``"center"``),
    /// ``wrap_text`` (``bool``), ``indent`` (non-negative integer).
    ///
    /// ``protection`` keys: ``locked``, ``hidden`` (both ``bool``).
    ///
    /// ``named_style`` (0.15.0-C1): a style name already defined in the LOADED file's own
    /// ``<cellStyles>`` (e.g. ``"Hyperlink"``, in whatever language the file itself uses
    /// for it — real Japanese-locale files spell it ``"ハイパーリンク"``). Bakes that
    /// style's font/fill/border/number-format/alignment/protection directly onto the
    /// cell, matching real Excel's own behavior — resolved FIRST, before any of the other
    /// five arguments on the same call, which then apply on top of it if also given.
    /// Raises ``OSError`` from :meth:`save_workbook` (not immediately — resolution is
    /// deferred like every other style edit, and :meth:`save_workbook` maps every save
    /// failure to ``OSError`` regardless of cause) if the name doesn't exist in this
    /// file. Defining a brand-new named style that doesn't already exist is not
    /// supported — only applying one the file already has.
    ///
    /// Color strings are a 6-digit RGB hex (``"4472C4"``, alpha assumed fully opaque) or
    /// an 8-digit ARGB hex (``"FF4472C4"``), with or without a leading ``#``. Only
    /// literal colors are supported — a theme-relative color (Excel's own theme palette)
    /// can't be set this way; cloning a font/fill/border that already uses one and not
    /// touching that specific property preserves it unchanged.
    ///
    /// Same safety/deferred-write model as :meth:`set_number_format`: never mutates a
    /// style record shared with cells outside *addr*, and the underlying
    /// ``xl/styles.xml`` write is deferred to :meth:`save_workbook`.
    ///
    /// Parameters
    /// ----------
    /// addr:
    ///     A single-area A1 range, e.g. ``"A1:D1"``, or a single cell.
    /// font, fill, border, alignment, protection:
    ///     Optional dicts, see above.
    /// named_style:
    ///     Optional existing style name, see above. At least one of the six arguments
    ///     must be given.
    /// sheet:
    ///     Sheet to modify. Defaults to the active sheet; does **not** change the
    ///     active sheet when given.
    ///
    /// Raises ``ValueError`` on a bad/oversized address, an unknown *sheet*, no
    /// arguments given, an unsupported ``fill['type']``, a missing required key
    /// (``fill`` needs ``color``), or a malformed color string. Raises ``TypeError`` on
    /// a wrong-typed value within a dict. An unknown *named_style* is NOT caught here —
    /// see above, it surfaces as ``OSError`` from :meth:`save_workbook` instead.
    #[pyo3(signature = (addr, font=None, fill=None, border=None, alignment=None, protection=None, named_style=None, sheet=None))]
    #[allow(clippy::too_many_arguments)]
    fn set_style(
        &mut self,
        addr: &str,
        font: Option<&Bound<'_, PyDict>>,
        fill: Option<&Bound<'_, PyDict>>,
        border: Option<&Bound<'_, PyDict>>,
        alignment: Option<&Bound<'_, PyDict>>,
        protection: Option<&Bound<'_, PyDict>>,
        named_style: Option<&str>,
        sheet: Option<&str>,
    ) -> PyResult<()> {
        if font.is_none()
            && fill.is_none()
            && border.is_none()
            && alignment.is_none()
            && protection.is_none()
            && named_style.is_none()
        {
            return Err(PyErr::new::<pyo3::exceptions::PyValueError, _>(
                "set_style requires at least one of font/fill/border/alignment/protection/named_style",
            ));
        }
        let ((r1, c1), (r2, c2)) =
            validate_range_addr(addr).map_err(PyErr::new::<pyo3::exceptions::PyValueError, _>)?;
        let key = self
            .inner
            .resolve_sheet_key(sheet)
            .map_err(PyErr::new::<pyo3::exceptions::PyValueError, _>)?;

        let edit = StyleAttrEdit {
            font: font.map(extract_font_edit).transpose()?,
            fill: fill.map(extract_fill_edit).transpose()?,
            border: border.map(extract_border_edit).transpose()?,
            alignment: alignment.map(extract_alignment_edit).transpose()?,
            protection: protection.map(extract_protection_edit).transpose()?,
            named_style: named_style.map(str::to_string),
        };
        self.inner.set_style_on_sheet(&key, r1, c1, r2, c2, &edit);
        Ok(())
    }

    /// Sets 1-based *row*'s DEFAULT style (0.15.0-C2) — applies to any cell in *row*
    /// that carries no explicit style of its own; a cell's own :meth:`set_style` always
    /// wins over its row's default. Same argument shapes, safety model, and deferred-
    /// write behavior as :meth:`set_style` (font/fill/border/alignment/protection/
    /// named_style, literal-RGB-only colors, resolved at :meth:`save_workbook` time) —
    /// see that method's own docstring for the full per-argument reference. Out of
    /// scope: theme-relative color minting, copy-style, named-style creation — same
    /// exclusions as :meth:`set_style`.
    ///
    /// Parameters
    /// ----------
    /// row:
    ///     1-based row number.
    /// font, fill, border, alignment, protection:
    ///     Optional dicts, see :meth:`set_style`.
    /// named_style:
    ///     Optional existing style name, see :meth:`set_style`.
    /// sheet:
    ///     Sheet to modify. Defaults to the active sheet; does **not** change the
    ///     active sheet when given.
    ///
    /// Raises ``ValueError`` if *row* is 0 or exceeds Excel's own grid limit
    /// (1,048,576 rows), an unknown *sheet*, no arguments given, or a malformed
    /// font/fill/border/alignment/protection value — same validation as
    /// :meth:`set_style`.
    #[pyo3(signature = (row, font=None, fill=None, border=None, alignment=None, protection=None, named_style=None, sheet=None))]
    #[allow(clippy::too_many_arguments)]
    fn set_row_style(
        &mut self,
        row: u32,
        font: Option<&Bound<'_, PyDict>>,
        fill: Option<&Bound<'_, PyDict>>,
        border: Option<&Bound<'_, PyDict>>,
        alignment: Option<&Bound<'_, PyDict>>,
        protection: Option<&Bound<'_, PyDict>>,
        named_style: Option<&str>,
        sheet: Option<&str>,
    ) -> PyResult<()> {
        const MAX_ROW: u32 = 1_048_576;
        if font.is_none()
            && fill.is_none()
            && border.is_none()
            && alignment.is_none()
            && protection.is_none()
            && named_style.is_none()
        {
            return Err(PyErr::new::<pyo3::exceptions::PyValueError, _>(
                "set_row_style requires at least one of font/fill/border/alignment/protection/named_style",
            ));
        }
        if row == 0 || row > MAX_ROW {
            return Err(PyErr::new::<pyo3::exceptions::PyValueError, _>(
                "row must be between 1 and 1_048_576",
            ));
        }
        let key = self
            .inner
            .resolve_sheet_key(sheet)
            .map_err(PyErr::new::<pyo3::exceptions::PyValueError, _>)?;
        let edit = StyleAttrEdit {
            font: font.map(extract_font_edit).transpose()?,
            fill: fill.map(extract_fill_edit).transpose()?,
            border: border.map(extract_border_edit).transpose()?,
            alignment: alignment.map(extract_alignment_edit).transpose()?,
            protection: protection.map(extract_protection_edit).transpose()?,
            named_style: named_style.map(str::to_string),
        };
        self.inner.set_row_style_on_sheet(&key, row, &edit);
        Ok(())
    }

    /// Column-axis mirror of :meth:`set_row_style` — sets 1-based *col*'s DEFAULT style
    /// (0.15.0-C2). See that method's own docstring for the full argument reference.
    ///
    /// Raises ``ValueError`` if *col* is 0 or exceeds Excel's own grid limit
    /// (16,384 columns), an unknown *sheet*, no arguments given, or a malformed
    /// font/fill/border/alignment/protection value.
    #[pyo3(signature = (col, font=None, fill=None, border=None, alignment=None, protection=None, named_style=None, sheet=None))]
    #[allow(clippy::too_many_arguments)]
    fn set_column_style(
        &mut self,
        col: u32,
        font: Option<&Bound<'_, PyDict>>,
        fill: Option<&Bound<'_, PyDict>>,
        border: Option<&Bound<'_, PyDict>>,
        alignment: Option<&Bound<'_, PyDict>>,
        protection: Option<&Bound<'_, PyDict>>,
        named_style: Option<&str>,
        sheet: Option<&str>,
    ) -> PyResult<()> {
        const MAX_COL: u32 = 16_384;
        if font.is_none()
            && fill.is_none()
            && border.is_none()
            && alignment.is_none()
            && protection.is_none()
            && named_style.is_none()
        {
            return Err(PyErr::new::<pyo3::exceptions::PyValueError, _>(
                "set_column_style requires at least one of font/fill/border/alignment/protection/named_style",
            ));
        }
        if col == 0 || col > MAX_COL {
            return Err(PyErr::new::<pyo3::exceptions::PyValueError, _>(
                "col must be between 1 and 16_384",
            ));
        }
        let key = self
            .inner
            .resolve_sheet_key(sheet)
            .map_err(PyErr::new::<pyo3::exceptions::PyValueError, _>)?;
        let edit = StyleAttrEdit {
            font: font.map(extract_font_edit).transpose()?,
            fill: fill.map(extract_fill_edit).transpose()?,
            border: border.map(extract_border_edit).transpose()?,
            alignment: alignment.map(extract_alignment_edit).transpose()?,
            protection: protection.map(extract_protection_edit).transpose()?,
            named_style: named_style.map(str::to_string),
        };
        self.inner.set_column_style_on_sheet(&key, col, &edit);
        Ok(())
    }

    /// Copies *source*'s complete style (font, fill, border, number format, alignment,
    /// and protection — everything, matching Excel's own "Format Painter") onto every
    /// cell in *dest* -- 0.15.0-C1. Picks up whatever *source*'s style resolves to at
    /// save time, including a `set_style`/`set_number_format` edit made on *source*
    /// itself earlier in the same session, even if not yet saved — no need to save
    /// between editing *source* and copying it. No new style record is ever created:
    /// *dest* cells simply point at the same underlying style as *source*, exactly what
    /// multiple cells sharing one style already means in the file format.
    ///
    /// A later `copy_style`/`set_style`/`set_number_format` call targeting the same
    /// destination cell before the next save always wins over an earlier one on that
    /// cell — but between two DIFFERENT features touching the same cell before one save,
    /// `copy_style` always takes effect last regardless of call order (it's resolved
    /// after every other pending style edit at save time).
    ///
    /// Parameters
    /// ----------
    /// source:
    ///     A single cell address, e.g. ``"A1"``.
    /// dest:
    ///     A single-area A1 range, e.g. ``"A2:A20"``, or a single cell.
    /// sheet:
    ///     Sheet both *source* and *dest* are on. Defaults to the active sheet; does
    ///     **not** change the active sheet when given.
    ///
    /// Raises ``ValueError`` if *source* isn't exactly one cell, on a bad/oversized
    /// *dest* address, or an unknown *sheet*.
    #[pyo3(signature = (source, dest, sheet = None))]
    fn copy_style(&mut self, source: &str, dest: &str, sheet: Option<&str>) -> PyResult<()> {
        let (src_start, src_end) =
            validate_range_addr(source).map_err(PyErr::new::<pyo3::exceptions::PyValueError, _>)?;
        if src_start != src_end {
            return Err(PyErr::new::<pyo3::exceptions::PyValueError, _>(
                "copy_style's source must be exactly one cell",
            ));
        }
        let ((r1, c1), (r2, c2)) =
            validate_range_addr(dest).map_err(PyErr::new::<pyo3::exceptions::PyValueError, _>)?;
        let key = self
            .inner
            .resolve_sheet_key(sheet)
            .map_err(PyErr::new::<pyo3::exceptions::PyValueError, _>)?;
        self.inner
            .copy_style_on_sheet(&key, src_start, r1, c1, r2, c2);
        Ok(())
    }

    /// Every hidden row number on a sheet, as a sorted list of 1-based row
    /// numbers (e.g. ``[5, 6, 9]``). Expanded, not interval-form.
    ///
    /// Raises ``ValueError`` if *sheet* is unknown.
    #[pyo3(signature = (sheet = None))]
    fn hidden_rows(&self, sheet: Option<&str>) -> PyResult<Vec<u32>> {
        let key = self
            .inner
            .resolve_sheet_key(sheet)
            .map_err(PyErr::new::<pyo3::exceptions::PyValueError, _>)?;
        Ok(self.inner.hidden_rows_on_sheet(&key))
    }

    /// Column-axis mirror of :meth:`hidden_rows`.
    #[pyo3(signature = (sheet = None))]
    fn hidden_columns(&self, sheet: Option<&str>) -> PyResult<Vec<u32>> {
        let key = self
            .inner
            .resolve_sheet_key(sheet)
            .map_err(PyErr::new::<pyo3::exceptions::PyValueError, _>)?;
        Ok(self.inner.hidden_columns_on_sheet(&key))
    }

    /// Hides or unhides a single row (1-based). Hiding an already-hidden row
    /// is a no-op; unhiding an already-visible row is a no-op.
    ///
    /// Raises ``ValueError`` if *row* is 0 or exceeds Excel's own grid limit
    /// (1,048,576 rows), or *sheet* is unknown.
    #[pyo3(signature = (row, hidden = true, sheet = None))]
    fn set_row_hidden(&mut self, row: u32, hidden: bool, sheet: Option<&str>) -> PyResult<()> {
        const MAX_ROW: u32 = 1_048_576;
        if row == 0 || row > MAX_ROW {
            return Err(PyErr::new::<pyo3::exceptions::PyValueError, _>(
                "row must be between 1 and 1_048_576",
            ));
        }
        let key = self
            .inner
            .resolve_sheet_key(sheet)
            .map_err(PyErr::new::<pyo3::exceptions::PyValueError, _>)?;
        self.inner.set_row_hidden_on_sheet(&key, row, hidden);
        Ok(())
    }

    /// Column-axis mirror of :meth:`set_row_hidden`.
    ///
    /// Raises ``ValueError`` if *col* is 0 or exceeds Excel's own grid limit
    /// (16,384 columns), or *sheet* is unknown.
    #[pyo3(signature = (col, hidden = true, sheet = None))]
    fn set_column_hidden(&mut self, col: u32, hidden: bool, sheet: Option<&str>) -> PyResult<()> {
        const MAX_COL: u32 = 16_384;
        if col == 0 || col > MAX_COL {
            return Err(PyErr::new::<pyo3::exceptions::PyValueError, _>(
                "col must be between 1 and 16_384",
            ));
        }
        let key = self
            .inner
            .resolve_sheet_key(sheet)
            .map_err(PyErr::new::<pyo3::exceptions::PyValueError, _>)?;
        self.inner.set_column_hidden_on_sheet(&key, col, hidden);
        Ok(())
    }

    /// A row's explicit height in points (P2), or ``None`` if it was never
    /// explicitly set (it uses the sheet's default row height, which this VM
    /// doesn't track as a queryable value). Read-only: there's no
    /// ``set_row_height`` yet, and a loaded file's row heights are dropped on
    /// EVERY save regardless of what this reports (the writer unconditionally
    /// regenerates ``<row>`` from hidden-row state alone) -- no real fixture
    /// has a custom row height to validate a writer shape against.
    ///
    /// Raises ``ValueError`` if *sheet* is unknown.
    #[pyo3(signature = (row, sheet = None))]
    fn row_height(&self, row: u32, sheet: Option<&str>) -> PyResult<Option<f64>> {
        let key = self
            .inner
            .resolve_sheet_key(sheet)
            .map_err(PyErr::new::<pyo3::exceptions::PyValueError, _>)?;
        Ok(self.inner.row_height_on_sheet(&key, row))
    }

    /// Column-axis mirror of :meth:`row_height` -- a column's explicit width
    /// in "characters" (the XLSX unit, font-relative), or ``None`` if never
    /// explicitly set. Same read-only caveat: no ``set_column_width`` yet, and
    /// a loaded file's column widths are dropped on every save today.
    ///
    /// Raises ``ValueError`` if *sheet* is unknown.
    #[pyo3(signature = (col, sheet = None))]
    fn column_width(&self, col: u32, sheet: Option<&str>) -> PyResult<Option<f64>> {
        let key = self
            .inner
            .resolve_sheet_key(sheet)
            .map_err(PyErr::new::<pyo3::exceptions::PyValueError, _>)?;
        Ok(self.inner.column_width_on_sheet(&key, col))
    }

    /// Python-native, single-key sort of a rectangular range, in place. Not
    /// from openpyxl (which has no sort primitive of its own) — this exposes
    /// the existing VBA ``Range(addr).Sort key:=, order:=, header:=``
    /// statement's exact behavior to Python.
    ///
    /// ``header=True`` excludes *addr*'s first row from the sort; it stays
    /// exactly where it is. Unlike the VBA statement (which silently clamps
    /// an out-of-range ``key_col``), this raises ``ValueError`` if *key_col*
    /// falls outside *addr*'s own column span.
    ///
    /// Does **not** check sheet protection — matches ``set_range``'s bulk
    /// cell-value-write precedent.
    ///
    /// Raises ``ValueError`` on a bad address, an out-of-bounds *key_col*, or
    /// an unknown *sheet* name.
    #[pyo3(signature = (addr, key_col, descending = false, header = false, sheet = None))]
    #[allow(clippy::too_many_arguments)]
    fn sort_range(
        &mut self,
        addr: &str,
        key_col: u32,
        descending: bool,
        header: bool,
        sheet: Option<&str>,
    ) -> PyResult<()> {
        const MAX_ROW: u32 = 1_048_576;
        const MAX_COL: u32 = 16_384;
        let ((r1, c1), (r2, c2)) =
            validate_range_addr(addr).map_err(PyErr::new::<pyo3::exceptions::PyValueError, _>)?;
        // Same ceiling as insert_rows/delete_rows -- validate_range_addr only
        // rejects 0 and reversed spans, not an absurdly large upper bound.
        // Unlike get_range/iter_rows (a large-but-harmless allocation), an
        // unbounded address here feeds a real write into the saved file.
        if r2 > MAX_ROW || c2 > MAX_COL {
            return Err(PyErr::new::<pyo3::exceptions::PyValueError, _>(format!(
                "range exceeds sheet bounds (max row {MAX_ROW}, max col {MAX_COL}), got row {r2}, col {c2}"
            )));
        }
        if key_col < c1 || key_col > c2 {
            return Err(PyErr::new::<pyo3::exceptions::PyValueError, _>(format!(
                "key_col {key_col} is outside the range's column span {c1}..={c2}"
            )));
        }
        let key = self
            .inner
            .resolve_sheet_key(sheet)
            .map_err(PyErr::new::<pyo3::exceptions::PyValueError, _>)?;
        self.inner
            .sort_range_on_sheet(&key, r1, c1, r2, c2, key_col, descending, header);
        Ok(())
    }
}

/// Shared row-major `Vec<Vec<Variant>>` -> Python nested-list conversion for
/// `get_range`/`iter_rows`.
#[cfg(feature = "python")]
fn grid_to_py(py: Python<'_>, grid: &[Vec<Variant>]) -> PyResult<Py<PyAny>> {
    let rows = pyo3::types::PyList::empty(py);
    for row in grid {
        let py_row = pyo3::types::PyList::empty(py);
        for v in row {
            py_row.append(variant_to_py(py, v))?;
        }
        rows.append(py_row)?;
    }
    Ok(rows.into_any().unbind())
}

// ── Module-level functions ────────────────────────────────────────────────────

/// Run a VBA macro string and return the resulting cells as ``{(row, col): value}``.
///
/// Parameters
/// ----------
/// vba_code : str
///     Full VBA source containing the target Sub.
/// macro_name : str
///     Name of the Sub to execute.
/// on_msgbox : str
///     ``"skip"`` (default) or ``"error"``.
#[cfg(feature = "python")]
#[pyfunction]
#[pyo3(signature = (vba_code, macro_name, on_msgbox = "skip"))]
fn run_macro(
    py: Python<'_>,
    vba_code: &str,
    macro_name: &str,
    on_msgbox: &str,
) -> PyResult<Py<PyAny>> {
    let mut vm = PyVm::new(on_msgbox)?;
    vm.run(vba_code, macro_name)?;
    vm.cells(py)
}

/// Load cell data from a spreadsheet file (.xlsx / .xlsm / .ods) into a new ``Vm``.
///
/// The VBA source code is **not** extracted from the file — pass it separately
/// to ``vm.run()``.
///
/// Parameters
/// ----------
/// path : str
///     Path to the spreadsheet file (.xlsx, .xlsm, or .ods).
/// sheet : str, optional
///     Sheet name to read. Defaults to the first sheet.
/// on_msgbox : str, optional
///     ``"skip"`` (default) or ``"error"``.
#[cfg(feature = "python")]
#[pyfunction]
#[pyo3(signature = (path, sheet = None, on_msgbox = "skip"))]
fn load_workbook(path: &str, sheet: Option<&str>, on_msgbox: &str) -> PyResult<PyVm> {
    let sheets =
        reader::read_workbook(path).map_err(PyErr::new::<pyo3::exceptions::PyIOError, _>)?;

    if sheets.is_empty() {
        return Err(PyErr::new::<pyo3::exceptions::PyValueError, _>(
            "Workbook has no sheets",
        ));
    }

    let mut vm = Vm::new();
    vm.error_on_msgbox = on_msgbox == "error";
    vm.populate_from_sheets(sheets);
    vm.loaded_workbook_path = Some(path.to_string());

    if let Some(s) = sheet {
        vm.set_active_sheet(&s.to_lowercase())
            .map_err(PyErr::new::<pyo3::exceptions::PyValueError, _>)?;
    }

    Ok(PyVm { inner: vm })
}

#[cfg(feature = "python")]
#[pyfunction]
fn hello() -> &'static str {
    "Hello from elixcee (Rust)!"
}

// ── save_workbook implementation ─────────────────────────────────────────────

/// Save all sheets in `vm` to a file. Supports `.xlsx` and `.ods`.
pub fn save_workbook(vm: &Vm, path: &str) -> Result<(), String> {
    save_workbook_impl(vm, path)
}

fn save_workbook_impl(vm: &Vm, path: &str) -> Result<(), String> {
    if path.to_lowercase().ends_with(".ods") {
        return save_ods_impl(vm, path);
    }
    save_xlsx_impl(vm, path)
}

/// 0.10.0-D, slice D1: one worksheet's complete set of output identifiers, computed once
/// per save by `plan_worksheet_output` rather than independently re-derived at each of the
/// several places that used to compute `sheet{i+1}.xml`/`sheetId`/`r:id` on their own
/// (`build_xlsx_content_types`, `build_xlsx_workbook`, `build_xlsx_workbook_rels`, and the
/// per-sheet write loop in `save_xlsx_impl`).
///
/// `output_part_name` is the load-bearing field this slice exists for: an EXISTING sheet
/// (one with a `WorksheetOrigin.original_part_name`) keeps that exact part name regardless
/// of its position in this save's `<sheets>` order, rather than being renumbered to
/// `sheet{i+1}.xml` every time. This is what lets a worksheet-level `.rels` file — which
/// already survives keyed by its ORIGINAL part path via the generic passthrough mechanism,
/// untouched by this slice — land back next to the worksheet content it actually belongs
/// to, instead of the two silently drifting apart on any save where sheets aren't in
/// exactly their original left-to-right order. Restoring the `r:id` REFERENCE inside that
/// worksheet content (so `check_source_references()`'s `SOURCE_REFERENCE_LOSS` actually
/// clears) is 0.10.0-D's later slice D2, not this one — D1 only makes sure content and
/// `.rels` land at the same path again.
///
/// A NEW sheet (no origin) gets a freshly allocated `sheetN.xml` that can't collide with
/// any part name that ever existed in the source file, including a deleted sheet's — see
/// `plan_worksheet_output`'s own doc comment for why that matters even before deleted-sheet
/// part cleanup exists (that's a later D slice too).
struct WorksheetOutputPlan {
    sheet_key: String,
    display_name: String,
    sheet_id: String,
    /// Positional (`rId{1-based index in this save's sheet order}`) — unrelated to a
    /// sheet's origin identity, purely internal to this writer's own
    /// `workbook.xml`/`workbook.xml.rels` pair (see `build_xlsx_workbook`'s original doc
    /// comment on this point, predating this struct).
    workbook_rel_id: String,
    output_part_name: String,
    /// The sheet's own `_rels/sheetN.xml.rels` path, derived from `output_part_name`.
    /// `save_xlsx_impl`'s `rels_survived` check uses this to confirm the `.rels` actually
    /// made it into `passthrough` before splicing any relationship-backed element back —
    /// see `OpaqueWorksheetFragments::table_parts`'s doc comment.
    output_rels_name: String,
    /// Has a `WorksheetOrigin` with a real `original_part_name` — false for a sheet
    /// created purely in-VBA (`Sheets.Add`), which can never have relationship-backed
    /// content to restore. `D4` will also use this to decide whether a deleted sheet's
    /// exclusively-reachable target parts need cleaning up.
    is_existing: bool,
}

/// Extracts `N` from a `xl/worksheets/sheetN.xml`-shaped part name (any prefix that
/// literally matches this pattern; a non-standard worksheet part name, which ECMA-376
/// technically permits but no fixture in this repo has ever shown, simply doesn't
/// contribute a reserved number — safe, since `plan_worksheet_output`'s existing-sheet
/// path never re-derives a number from a name at all, only new-sheet allocation does).
fn parse_sheet_part_number(name: &str) -> Option<u32> {
    name.strip_prefix("xl/worksheets/sheet")?
        .strip_suffix(".xml")?
        .parse()
        .ok()
}

/// `part` -> its own `_rels/<file>.rels` sibling path, the fixed relative-path rule every
/// OOXML part's own `.rels` file follows (a `_rels` sibling directory of the part, named
/// `<part-filename>.rels`) -- a `.rels` file is never itself a relationship TARGET, it's
/// discovered by this naming convention. Used both for a worksheet's own output rels name
/// (0.10.0-D1) and, generically, by `deleted_sheet_prunable_parts`'s reachability walk
/// (0.10.0-D4) -- the name predates the second use but the logic needed no change.
fn part_rels_name(part_name: &str) -> String {
    match part_name.rsplit_once('/') {
        Some((dir, file)) => format!("{dir}/_rels/{file}.rels"),
        None => format!("_rels/{part_name}.rels"),
    }
}

/// Builds one `WorksheetOutputPlan` per sheet in `sheet_names` (this save's real order,
/// `Vm::sheet_order`). `reserved_part_numbers` is every `sheetN.xml` number that ever
/// existed in the source file (see `save_xlsx_impl`'s own doc comment on
/// `reserved_sheet_part_numbers` for why deleted sheets' numbers must be included too) —
/// empty for a from-scratch `Vm` or an `.ods` source, in which case a new sheet's number
/// is derived from surviving origins' own part names instead (there's no raw source file
/// to scan, but existing sheets can still have origins from an earlier load).
fn plan_worksheet_output(
    sheet_names: &[String],
    origins: &std::collections::HashMap<String, WorksheetOrigin>,
    reserved_part_numbers: &[u32],
) -> Vec<WorksheetOutputPlan> {
    let mut reserved: Vec<u32> = reserved_part_numbers.to_vec();
    if reserved.is_empty() {
        reserved.extend(
            origins
                .values()
                .filter_map(|o| o.original_part_name.as_deref())
                .filter_map(parse_sheet_part_number),
        );
    }
    let mut next_fresh_part_n = reserved.into_iter().max().unwrap_or(0);

    let max_original_id: u32 = sheet_names
        .iter()
        .filter_map(|name| origins.get(name))
        .filter_map(|o| o.original_sheet_id.as_deref())
        .filter_map(|id| id.parse::<u32>().ok())
        .max()
        .unwrap_or(0);
    let mut next_fresh_id = max_original_id;

    sheet_names
        .iter()
        .enumerate()
        .map(|(i, sheet_key)| {
            let origin = origins.get(sheet_key);
            let output_part_name = match origin.and_then(|o| o.original_part_name.clone()) {
                Some(part) => part,
                None => {
                    next_fresh_part_n += 1;
                    format!("xl/worksheets/sheet{next_fresh_part_n}.xml")
                }
            };
            let sheet_id = match origin.and_then(|o| o.original_sheet_id.clone()) {
                Some(id) => id,
                None => {
                    next_fresh_id += 1;
                    next_fresh_id.to_string()
                }
            };
            let display_name = origin
                .and_then(|o| o.original_display_name.clone())
                .unwrap_or_else(|| sheet_key.clone());
            WorksheetOutputPlan {
                sheet_key: sheet_key.clone(),
                display_name,
                sheet_id,
                workbook_rel_id: format!("rId{}", i + 1),
                output_rels_name: part_rels_name(&output_part_name),
                is_existing: origin.and_then(|o| o.original_part_name.as_ref()).is_some(),
                output_part_name,
            }
        })
        .collect()
}

/// The directory a `.rels` file's own relative `Target` attributes are resolved against
/// -- e.g. `"xl/worksheets/_rels/sheet1.xml.rels"` -> `"xl/worksheets/"`, `"_rels/.rels"`
/// (the package root's own rels, no owning part) -> `""`. Mirrors
/// `compat/oracle-excel-com/mechanical_check.py`'s `_rels_target_dir`.
fn rels_target_dir(rels_name: &str) -> &str {
    match rels_name.rfind("/_rels/") {
        Some(idx) => &rels_name[..idx + 1],
        None => "",
    }
}

/// One-hop internal relationship targets declared by the `.rels` part named `rels_name`,
/// resolved to normalized part paths. Empty if `rels_name` is absent from `raw_entries`
/// or not valid UTF-8. Doesn't filter `TargetMode="External"` (the raw
/// `reader::workbook_rels_decls` parse doesn't carry that attribute at all) -- harmless
/// here, same as in `carry_over_rels`'s existing use of the same parse: an external
/// target (a URL, not a part path) never matches any real `raw_entries` key, so it just
/// never affects set membership against real parts.
fn direct_rel_targets(
    raw_entries: &std::collections::HashMap<String, Vec<u8>>,
    rels_name: &str,
) -> std::collections::HashSet<String> {
    let Some(text) = raw_entries
        .get(rels_name)
        .and_then(|b| String::from_utf8(b.clone()).ok())
    else {
        return Default::default();
    };
    let base = rels_target_dir(rels_name);
    reader::workbook_rels_decls(&text)
        .into_iter()
        .map(|(_, target)| normalize_part_path(&format!("{base}{target}")))
        .collect()
}

/// BFS over `raw_entries`' package relationship graph, starting from `roots` (each
/// reached via its own `part_rels_name()` sibling, per OPC convention). Returns every
/// part transitively reachable, INCLUDING the roots themselves and every `.rels` file
/// walked along the way -- a `.rels` file isn't a relationship target, but "belongs to"
/// the part that owns it for pruning purposes (see `deleted_sheet_prunable_parts`).
///
/// `exclude` is filtered at EVERY hop, not just the roots: a deleted sheet's own
/// worksheet part can be re-discovered via more than one path in the source's graph (its
/// own edge in `xl/_rels/workbook.xml.rels`, but also indirectly via `_rels/.rels` ->
/// `xl/workbook.xml` -> that same unfiltered `.rels`) -- filtering only the seed set
/// stops the first path but not the second, silently reintroducing the deleted sheet (and
/// everything reachable from it) as "reachable elsewhere". Mirrors
/// `compat/oracle-excel-com/mechanical_check.py`'s `_reachable_closure` exactly (verified
/// independently in Python before this was written -- see that file's Case N self-test,
/// which caught this exact bug first).
fn reachable_closure(
    raw_entries: &std::collections::HashMap<String, Vec<u8>>,
    roots: impl IntoIterator<Item = String>,
    exclude: &std::collections::HashSet<String>,
) -> std::collections::HashSet<String> {
    let mut seen: std::collections::HashSet<String> =
        roots.into_iter().filter(|p| !exclude.contains(p)).collect();
    let mut queue: Vec<String> = seen.iter().cloned().collect();
    while let Some(part) = queue.pop() {
        let rels_name = part_rels_name(&part);
        if raw_entries.contains_key(&rels_name) {
            seen.insert(rels_name.clone());
            for target in direct_rel_targets(raw_entries, &rels_name) {
                if exclude.contains(&target) || seen.contains(&target) {
                    continue;
                }
                seen.insert(target.clone());
                queue.push(target);
            }
        }
    }
    seen
}

/// 0.10.0-D4: the set of `raw_entries` part names that should be pruned from output
/// because they are reachable ONLY from a sheet no longer present in `sheet_order` --
/// computed purely from the source's own relationship graph, independent of anything
/// already decided about `passthrough`. Empty if no sheet was deleted.
///
/// A part reachable from a deleted sheet AND from anything else (a surviving sheet, or a
/// workbook-level relationship via `xl/_rels/workbook.xml.rels` or the root
/// `_rels/.rels`) is never included -- shared parts must survive regardless of how many
/// referencing sheets are gone. Closes ROADMAP.md Known gaps item 15: before this, a
/// deleted sheet's own `.rels` (and whatever it exclusively pointed at) survived as an
/// orphan, invisible to `mechanical_check.py`'s structural checks alone -- see
/// `compat/oracle-excel-com/mechanical_check.py`'s `check_deleted_sheet_cleanup`, the
/// dedicated checker written and self-test-verified before this function, per this
/// project's fixture/design-doc-recorded hard gate.
fn deleted_sheet_prunable_parts(
    raw_entries: &std::collections::HashMap<String, Vec<u8>>,
    worksheet_origins: &std::collections::HashMap<String, WorksheetOrigin>,
    sheet_order: &[String],
) -> std::collections::HashSet<String> {
    let deleted_parts: std::collections::HashSet<String> = worksheet_origins
        .iter()
        .filter(|(key, _)| !sheet_order.iter().any(|s| s == *key))
        .filter_map(|(_, origin)| origin.original_part_name.clone())
        .collect();
    if deleted_parts.is_empty() {
        return Default::default();
    }
    let surviving_parts: std::collections::HashSet<String> = worksheet_origins
        .iter()
        .filter(|(key, _)| sheet_order.iter().any(|s| s == *key))
        .filter_map(|(_, origin)| origin.original_part_name.clone())
        .collect();

    let mut elsewhere_roots: Vec<String> = surviving_parts.into_iter().collect();
    elsewhere_roots.extend(direct_rel_targets(
        raw_entries,
        "xl/_rels/workbook.xml.rels",
    ));
    elsewhere_roots.extend(direct_rel_targets(raw_entries, "_rels/.rels"));
    let reachable_elsewhere = reachable_closure(raw_entries, elsewhere_roots, &deleted_parts);

    let reachable_from_deleted = reachable_closure(
        raw_entries,
        deleted_parts.iter().cloned(),
        &Default::default(),
    );

    reachable_from_deleted
        .difference(&reachable_elsewhere)
        .filter(|p| raw_entries.contains_key(p.as_str()))
        .cloned()
        .collect()
}

/// Parts this writer always regenerates from `Vm` state — everything else read
/// from a passthrough source is copied through byte-for-byte (Milestone: safe
/// round-trip). Pattern-matched rather than checking against `sheet{i+1}.xml`
/// for `i in 0..sheet_names.len()`: a real workbook can have non-sequential
/// worksheet part names (e.g. sheets deleted, leaving `sheet2.xml`/
/// `sheet3.xml` as survivors) — keying exclusion off *this writer's own*
/// sequential naming would leave such a stale original part sitting alongside
/// the freshly-regenerated `sheet1.xml`. See `docs/xlsx-architecture.md`.
fn is_writer_owned_part(name: &str) -> bool {
    matches!(
        name,
        "[Content_Types].xml"
            | "_rels/.rels"
            | "xl/workbook.xml"
            | "xl/_rels/workbook.xml.rels"
            | "xl/sharedStrings.xml"
            | "xl/styles.xml"
    ) || (name.starts_with("xl/worksheets/")
        && name.ends_with(".xml")
        && !name["xl/worksheets/".len()..].contains('/'))
}

/// Parses `rels_part` out of `raw_entries` (a source's raw zip contents) and returns
/// every `(Type, Target)` relationship whose target both (a) survived into `passthrough`
/// and (b) isn't already one of `skip_types` -- the types this writer emits its own
/// relationship for elsewhere, which would otherwise get a duplicate. `target_base` is
/// prepended before resolving (`"xl/"` for `xl/_rels/workbook.xml.rels`, whose targets
/// are relative to `xl/`; `""` for `_rels/.rels`, whose targets are already relative to
/// the package root). See `save_xlsx_impl`'s `carried_rels`/`carried_root_rels` doc
/// comments for why this exists at all: a writer-owned `.rels` file that's regenerated
/// from a fixed template (not derived from the source) silently drops any relationship
/// the template doesn't know about, orphaning an otherwise-correctly-passed-through part
/// -- confirmed live against real Excel, which refuses to open the result outright.
fn carry_over_rels(
    raw_entries: &std::collections::HashMap<String, Vec<u8>>,
    rels_part: &str,
    target_base: &str,
    passthrough: &[(String, Vec<u8>)],
    skip_types: &[&str],
) -> Vec<(String, String)> {
    let Some(rels_xml) = raw_entries
        .get(rels_part)
        .and_then(|b| String::from_utf8(b.clone()).ok())
    else {
        return Vec::new();
    };
    reader::workbook_rels_decls(&rels_xml)
        .into_iter()
        .filter(|(ty, _)| !skip_types.contains(&ty.as_str()))
        .filter(|(_, target)| {
            let resolved = normalize_part_path(&format!("{}{}", target_base, target));
            passthrough.iter().any(|(name, _)| *name == resolved)
        })
        .collect()
}

/// Rewrites `container_xml` (the raw `<definedNames>...</definedNames>` blob already
/// cleared to pass through -- see the deletion/move gate right above this function's only
/// call site) for every sheet rename in `renames`, splicing each `<definedName>` element's
/// value via `formula::rewrite_defined_name_for_renames` instead of leaving stale text or
/// dropping the whole container. `CT_DefinedNames` has no attributes of its own (confirmed
/// against the real XSD -- just a sequence of `<definedName>` children), so the wrapper tag
/// is always emitted plain; `None` if nothing survives (every child individually dropped) or
/// if `container_xml` has no `<definedName>` children to rewrite (malformed/empty -- passed
/// through unchanged is wrong once a rename makes rewriting necessary, but there's nothing
/// here to rewrite, so the caller's fallback of just not using this path is equivalent).
///
/// `extract_defined_name_elements`'s spans are raw XML text, not the value a formula parser
/// understands (e.g. a concatenation `&` is `&amp;` on the wire) -- unescaped via
/// `reader::xml_unescape` before handing to `formula::rewrite_defined_name_for_renames`, and
/// a genuinely rewritten result is escaped back via `xml_escape` before splicing in. An
/// unchanged value is spliced back RAW (not round-tripped through unescape+escape), so a
/// source file's own escaping style (e.g. `&#38;` vs `&amp;`, both valid XML for the same
/// character) survives byte-for-byte when nothing about it needed to change.
fn rewrite_defined_names_xml(
    container_xml: &str,
    renames: &std::collections::HashMap<String, String>,
) -> Option<String> {
    let elements = reader::extract_defined_name_elements(container_xml);
    if elements.is_empty() {
        return None;
    }
    let mut rebuilt = String::new();
    for (el, (text_start, text_end)) in &elements {
        let raw_value = &el[*text_start..*text_end];
        let unescaped = reader::xml_unescape(raw_value);
        match formula::rewrite_defined_name_for_renames(&unescaped, renames) {
            formula::DefinedNameRewrite::Unchanged => rebuilt.push_str(el),
            formula::DefinedNameRewrite::Rewritten(new_value) => {
                rebuilt.push_str(&el[..*text_start]);
                rebuilt.push_str(&xml_escape(&new_value));
                rebuilt.push_str(&el[*text_end..]);
            }
            formula::DefinedNameRewrite::Drop => {}
        }
    }
    if rebuilt.is_empty() {
        None
    } else {
        Some(format!("<definedNames>{}</definedNames>", rebuilt))
    }
}

fn save_xlsx_impl(vm: &Vm, path: &str) -> Result<(), String> {
    use std::collections::HashMap;
    use std::io::{Cursor, Write};
    use zip::CompressionMethod;
    use zip::write::ZipWriter;

    // Real tab order, not `vm.sheet_names()`'s alphabetical order — a
    // workbook's physical sheet order (part naming, `<sheets>` order,
    // sheetId assignment) must match the source, not a sort. Found via a
    // synthetic "Zebra"/"Alpha" fixture that round-tripped as "Alpha"/
    // "Zebra". `sheet_order` is `Vm`'s parallel, insertion-ordered sheet
    // list; see its doc comment for how it's kept in sync with `sheets`.
    let sheet_names = vm.sheet_order.clone();

    // Collect shared strings (insertion-ordered, deduplicated)
    let mut str_index: HashMap<String, usize> = HashMap::new();
    let mut shared_strings: Vec<String> = Vec::new();
    for sheet_name in &sheet_names {
        if let Some(cells) = vm.get_sheet_cells(sheet_name) {
            let mut sorted: Vec<_> = cells.keys().collect();
            sorted.sort();
            for key in sorted {
                // Variant::Error is deliberately excluded: a real error cell is written as
                // t="e" with its literal error text in <v> (see xlsx_cell_xml below), never
                // shared-string indexed -- confirmed against real Excel-authored output,
                // which never puts e.g. "#VALUE!" in xl/sharedStrings.xml either.
                let s = match &cells[key].value {
                    Variant::Str(s) => s.as_str().to_string(),
                    _ => continue,
                };
                if !str_index.contains_key(&s) {
                    str_index.insert(s.clone(), shared_strings.len());
                    shared_strings.push(s);
                }
            }
        }
    }

    // ── Unknown-part passthrough (Milestone: safe round-trip) ───────────────
    // Only activates for an .xlsx/.xlsm source — never .ods, whose parts would
    // be meaningless (and wrong) inside an OOXML package. Re-reads the source
    // file fully into memory here, at save time, rather than caching it at
    // load time, so read-only paths (`check`/`snapshot`/`diagnose`) never pay
    // this cost. See `docs/xlsx-architecture.md`.
    let passthrough_source = vm.loaded_workbook_path.as_deref().filter(|p| {
        let l = p.to_lowercase();
        l.ends_with(".xlsx") || l.ends_with(".xlsm")
    });
    let is_xlsm_output = path.to_lowercase().ends_with(".xlsm");

    let mut passthrough: Vec<(String, Vec<u8>)> = Vec::new();
    let mut has_vba = false;
    let mut carried_overrides: Vec<(String, String)> = Vec::new();
    // Other workbook-level relationships (theme, calcChain, etc.) whose target part
    // survives as a passthrough entry -- see the loop below for why this is needed at
    // all: a part being copied through byte-for-byte does NOT mean anything still
    // references it. Excel found and confirmed this live (a real .xlsm's own
    // theme/theme1.xml, copied through correctly, refused to open at all once its
    // workbook.xml.rels relationship vanished -- an orphaned part, not just a stale
    // one). (Type, Target) pairs; Target is re-relativized against "xl/" by
    // build_xlsx_workbook_rels, matching how workbook.xml.rels' own Target values work.
    let mut carried_rels: Vec<(String, String)> = Vec::new();
    // Same idea as `carried_rels`, but for the root `_rels/.rels` file (docProps/core.xml,
    // docProps/app.xml, ...) -- see `carried_rels`'s own comment below for why this is
    // needed at all.
    let mut carried_root_rels: Vec<(String, String)> = Vec::new();
    // xl/styles.xml is writer-owned (never enters the generic `passthrough` loop
    // below, via `is_writer_owned_part`) but, unlike the other writer-owned parts,
    // its CONTENT is conditionally passed through rather than always regenerated
    // from the hardcoded `XLSX_STYLES` minimal stylesheet -- these raw bytes are the
    // starting point either way, byte-for-byte when nothing touched styles (still true
    // for every VBA statement -- `Range.Interior.Color =`/`.NumberFormat =` remain
    // explicit no-ops), or as `resolve_pending_number_formats`'s input when
    // `set_number_format` (0.15.0-A) has pending edits. See `docs/xlsx-architecture.md`.
    let mut passthrough_styles: Option<Vec<u8>> = None;
    // Original worksheet XML text, keyed by lowercased sheet name (matching
    // `worksheet_origins`'/`sheet_names()`'s own key space) — the source for
    // 0.10.0-B's opaque-fragment passthrough (see `build_xlsx_sheet`'s
    // `root_attrs`/`sheet_views` params). Only populated for sheets with a known
    // `WorksheetOrigin` whose `original_part_name` resolves to a real passthrough
    // entry; a new sheet (no origin) or an .ods source (no `raw_entries` at all)
    // falls back to `build_xlsx_sheet`'s hardcoded minimal defaults.
    let mut sheet_source_xml: HashMap<String, String> = HashMap::new();
    // Original xl/workbook.xml text -- the source for 0.10.0-C's opaque-fragment
    // passthrough (see `OpaqueWorkbookFragments`), same mechanism as `sheet_source_xml`
    // above but for the single, fixed-path workbook part rather than per-sheet parts.
    let mut workbook_source_xml: Option<String> = None;
    // Every "xl/worksheets/sheetN.xml" part number that ever existed in the SOURCE file
    // -- not just the sheets that survived to this save. A fresh worksheet part name
    // (0.10.0-D's WorksheetOutputPlan, see below) must never collide with one of these,
    // including a deleted sheet's original number -- a deliberate, permanent policy
    // (never reuse a freed number), not merely a stopgap made moot by D4's pruning below:
    // reuse would let a stale, not-yet-understood future reference collide with a
    // brand-new sheet, which pruning's own "reachable elsewhere" analysis cannot rule out
    // for parts this writer has never modeled at all.
    let mut reserved_sheet_part_numbers: Vec<u32> = Vec::new();

    if let Some(source_path) = passthrough_source {
        let raw_entries = reader::read_raw_zip_entries(source_path)?;
        has_vba = is_xlsm_output && raw_entries.keys().any(|n| n.starts_with("xl/vbaProject"));
        passthrough_styles = raw_entries.get("xl/styles.xml").cloned();
        workbook_source_xml = raw_entries
            .get("xl/workbook.xml")
            .and_then(|bytes| String::from_utf8(bytes.clone()).ok());
        reserved_sheet_part_numbers = raw_entries
            .keys()
            .filter_map(|name| parse_sheet_part_number(name))
            .collect();

        for (sheet_key, origin) in &vm.worksheet_origins {
            if let Some(part) = &origin.original_part_name
                && let Some(bytes) = raw_entries.get(part)
                && let Ok(text) = String::from_utf8(bytes.clone())
            {
                sheet_source_xml.insert(sheet_key.clone(), text);
            }
        }

        let (defaults, overrides) = raw_entries
            .get("[Content_Types].xml")
            .and_then(|b| String::from_utf8(b.clone()).ok())
            .map(|xml| reader::content_type_decls(&xml))
            .unwrap_or_default();

        // 0.10.0-D4: a deleted sheet's exclusively-reachable parts (its own worksheet
        // .rels, and whatever THAT points at, transitively) must never enter
        // `passthrough` at all -- computed here, before the loop below, so both this
        // loop's `carried_overrides` and the `carry_over_rels` calls after it (which
        // filter to targets present `&passthrough`) automatically treat a pruned part as
        // gone, with no separate cleanup pass needed.
        let prunable_parts =
            deleted_sheet_prunable_parts(&raw_entries, &vm.worksheet_origins, &sheet_names);

        for (name, bytes) in &raw_entries {
            if is_writer_owned_part(name) {
                continue;
            }
            // Excel's own "Save As .xlsx" behavior: a macro project never
            // survives into a workbook declared non-macro-enabled.
            if !is_xlsm_output && name.starts_with("xl/vbaProject") {
                continue;
            }
            if prunable_parts.contains(name) {
                continue;
            }
            passthrough.push((name.clone(), bytes.clone()));

            // Resolve this part's real declared content type from the source's
            // own [Content_Types].xml — exact Override first, then extension
            // Default — instead of guessing. xml/rels extensions are already
            // covered by this writer's own baseline <Default> entries below.
            let part_name = format!("/{}", name);
            let resolved = overrides
                .iter()
                .find(|(p, _)| p == &part_name)
                .map(|(_, ct)| ct.clone())
                .or_else(|| {
                    let ext = name.rsplit('.').next().unwrap_or("");
                    if ext == "xml" || ext == "rels" {
                        None
                    } else {
                        defaults
                            .iter()
                            .find(|(e, _)| e == ext)
                            .map(|(_, ct)| ct.clone())
                    }
                })
                .or_else(|| {
                    // Defensive fallback for a malformed/incomplete source —
                    // an Override, not a blanket <Default Extension="bin">,
                    // which would also mis-declare a sibling part like
                    // xl/printerSettings/printerSettings1.bin.
                    if name.starts_with("xl/vbaProject") {
                        Some("application/vnd.ms-office.vbaProject".to_string())
                    } else {
                        None
                    }
                });
            if let Some(ct) = resolved {
                carried_overrides.push((part_name, ct));
            }
        }

        // Any OTHER relationship (theme, calcChain, docProps, ...) whose target survived
        // as a passthrough part -- skip the types this writer already emits its own
        // relationship for, or every passthrough part would get a duplicate, conflicting
        // relationship pointing at the same target. Applied to BOTH rels files that
        // matter here: xl/_rels/workbook.xml.rels (targets relative to "xl/") and
        // _rels/.rels (targets relative to the package root, no prefix) -- the root rels
        // file is just as writer-owned/hardcoded as the workbook one (XLSX_ROOT_RELS),
        // and orphaned docProps/core.xml + docProps/app.xml relationships were found
        // missing the exact same way theme/calcChain were.
        carried_rels.extend(carry_over_rels(
            &raw_entries,
            "xl/_rels/workbook.xml.rels",
            "xl/",
            &passthrough,
            &[
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
                "http://schemas.microsoft.com/office/2006/relationships/vbaProject",
            ],
        ));
        carried_root_rels.extend(carry_over_rels(
            &raw_entries,
            "_rels/.rels",
            "",
            &passthrough,
            &["http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"],
        ));

        // Deterministic, reviewable output order.
        passthrough.sort_by(|a, b| a.0.cmp(&b.0));
        carried_overrides.sort_by(|a, b| a.0.cmp(&b.0));
        carried_rels.sort_by(|a, b| a.1.cmp(&b.1));
        carried_root_rels.sort_by(|a, b| a.1.cmp(&b.1));
    }

    let worksheet_plans = plan_worksheet_output(
        &sheet_names,
        &vm.worksheet_origins,
        &reserved_sheet_part_numbers,
    );

    let cursor = Cursor::new(Vec::<u8>::new());
    let mut zip = ZipWriter::new(cursor);
    let deflated =
        zip::write::SimpleFileOptions::default().compression_method(CompressionMethod::Deflated);

    zip.start_file("[Content_Types].xml", deflated)
        .map_err(|e| e.to_string())?;
    zip.write_all(
        build_xlsx_content_types(&worksheet_plans, is_xlsm_output, &carried_overrides).as_bytes(),
    )
    .map_err(|e| e.to_string())?;

    zip.start_file("_rels/.rels", deflated)
        .map_err(|e| e.to_string())?;
    zip.write_all(build_xlsx_root_rels(&carried_root_rels).as_bytes())
        .map_err(|e| e.to_string())?;

    // .and_then(ensure_r_prefix_bound): a source is free to bind the relationships
    // namespace to any prefix (e.g. `xmlns:rel="..."` + `rel:id="..."`, equally valid
    // OOXML) -- but build_xlsx_workbook's <sheet r:id="..."> below always hardcodes the
    // literal `r:` prefix. Carrying such a source's root attrs through unchanged would
    // leave `r:` unbound in the output, an XML error every strict consumer rejects. See
    // ensure_r_prefix_bound's own doc comment for the real report this fixes.
    let workbook_root_attrs = workbook_source_xml
        .as_deref()
        .and_then(|xml| reader::extract_root_attrs(xml, "workbook"))
        .and_then(|attrs| reader::ensure_r_prefix_bound(&attrs));
    let workbook_pr = workbook_source_xml
        .as_deref()
        .and_then(|xml| reader::extract_raw_element(xml, "workbookPr"));
    let book_views = workbook_source_xml
        .as_deref()
        .and_then(|xml| reader::extract_raw_element(xml, "bookViews"));
    let calc_pr = workbook_source_xml
        .as_deref()
        .and_then(|xml| reader::extract_raw_element(xml, "calcPr"));
    let ext_lst = workbook_source_xml
        .as_deref()
        .and_then(|xml| reader::extract_raw_element(xml, "extLst"));
    // A <definedName>'s localSheetId is a 0-based index into <sheets> -- if any sheet
    // present at load time is gone now (Sheets(...).Delete ran), every remaining
    // localSheetId could point at the wrong sheet, so the whole element is dropped
    // rather than carried through stale. See OpaqueWorkbookFragments' doc comment.
    // Also dropped if `defined_names_may_be_stale` is set: `move_sheet` reordering
    // `sheet_order` invalidates a positional `localSheetId` exactly like a deletion
    // does, even though every sheet is still present. (VBA's `Sheets.Add(before:=...)`
    // can shift positions the same way without tripping either check -- a narrower,
    // pre-existing gap this doesn't close; see ROADMAP.md's known gaps.)
    //
    // `rename_sheet` no longer forces this wholesale drop -- see
    // `vm.sheet_renames_since_load` and `rewrite_defined_names_xml` below, which rewrite
    // a <definedName>'s stale sheet-qualifier TEXT (e.g. "Sheet1!$F$5") in place instead.
    let no_sheet_was_deleted = vm
        .worksheet_origins
        .keys()
        .all(|original_key| vm.sheet_order.contains(original_key));
    let defined_names_raw = if no_sheet_was_deleted && !vm.defined_names_may_be_stale {
        workbook_source_xml
            .as_deref()
            .and_then(|xml| reader::extract_raw_element(xml, "definedNames"))
    } else {
        None
    };
    let defined_names = if vm.sheet_renames_since_load.is_empty() {
        defined_names_raw
    } else {
        defined_names_raw
            .as_deref()
            .and_then(|xml| rewrite_defined_names_xml(xml, &vm.sheet_renames_since_load))
    };
    let workbook_fragments = OpaqueWorkbookFragments {
        root_attrs: workbook_root_attrs.as_deref(),
        workbook_pr: workbook_pr.as_deref(),
        book_views: book_views.as_deref(),
        defined_names: defined_names.as_deref(),
        calc_pr: calc_pr.as_deref(),
        ext_lst: ext_lst.as_deref(),
    };

    zip.start_file("xl/workbook.xml", deflated)
        .map_err(|e| e.to_string())?;
    zip.write_all(build_xlsx_workbook(&worksheet_plans, &workbook_fragments).as_bytes())
        .map_err(|e| e.to_string())?;

    zip.start_file("xl/_rels/workbook.xml.rels", deflated)
        .map_err(|e| e.to_string())?;
    zip.write_all(build_xlsx_workbook_rels(&worksheet_plans, has_vba, &carried_rels).as_bytes())
        .map_err(|e| e.to_string())?;

    // `set_number_format` (0.15.0-A) resolution -- see `resolve_pending_number_formats`'s
    // own doc comment. `styles_source` is exactly the same starting document the final
    // `xl/styles.xml` write below always used; resolving here (before the per-sheet loop)
    // rather than at that write site is what lets each sheet's own `<c s="N">` emission
    // below see the EFFECTIVE (possibly-edited) index instead of `vm.cell_style_indices`'
    // original one.
    let styles_source = passthrough_styles
        .as_deref()
        .unwrap_or_else(|| XLSX_STYLES.as_bytes());
    let mut new_styles_bytes: Option<Vec<u8>> = None;
    let mut effective_style_indices: Option<StyleIndexMap> = None;
    if let Some((new_xml, indices)) = resolve_pending_number_formats(vm, styles_source) {
        new_styles_bytes = Some(new_xml.into_bytes());
        effective_style_indices = Some(indices);
    }
    // `set_style` (0.15.0-B) resolution, CHAINED after the number-format pass above (see
    // `resolve_pending_style_attrs`'s own doc comment for why an independent second pass
    // starting fresh from `styles_source` would silently discard whichever of the two
    // features ran first on a cell touched by both).
    {
        let chained_source: &[u8] = new_styles_bytes.as_deref().unwrap_or(styles_source);
        let chained_indices: &StyleIndexMap = effective_style_indices
            .as_ref()
            .unwrap_or(&vm.cell_style_indices);
        if let Some((new_xml, indices)) =
            resolve_pending_style_attrs(vm, chained_source, chained_indices)?
        {
            new_styles_bytes = Some(new_xml.into_bytes());
            effective_style_indices = Some(indices);
        }
    }
    // `copy_style` (0.15.0-C1) resolution, CHAINED after the two passes above -- pure
    // index aliasing (see `resolve_pending_style_copies`'s own doc comment). Never
    // produces new styles.xml bytes, so `new_styles_bytes` is deliberately left
    // untouched here.
    {
        let chained_indices: &StyleIndexMap = effective_style_indices
            .as_ref()
            .unwrap_or(&vm.cell_style_indices);
        if let Some(indices) = resolve_pending_style_copies(vm, chained_indices) {
            effective_style_indices = Some(indices);
        }
    }
    // `set_row_style`/`set_column_style` (0.15.0-C2) resolution, CHAINED LAST -- shares
    // the same font/fill/border/cellXf tables `resolve_pending_style_attrs` may have
    // already grown above (see `resolve_pending_row_column_styles`'s own doc comment).
    let mut effective_row_styles: Option<RowStyleIndexMap> = None;
    let mut effective_column_styles: Option<ColumnStyleRangeMap> = None;
    {
        let chained_source: &[u8] = new_styles_bytes.as_deref().unwrap_or(styles_source);
        if let Some((new_xml, rows, cols)) = resolve_pending_row_column_styles(
            vm,
            chained_source,
            &vm.row_styles,
            &vm.column_styles,
        )? {
            new_styles_bytes = Some(new_xml.into_bytes());
            effective_row_styles = Some(rows);
            effective_column_styles = Some(cols);
        }
    }

    for plan in &worksheet_plans {
        let sheet_name = &plan.sheet_key;
        let source_xml = sheet_source_xml.get(&sheet_name.to_lowercase());
        let style_override = effective_style_indices
            .as_ref()
            .and_then(|m| m.get(&sheet_name.to_lowercase()));
        let row_style_override = effective_row_styles
            .as_ref()
            .and_then(|m| m.get(&sheet_name.to_lowercase()));
        let column_style_override = effective_column_styles
            .as_ref()
            .and_then(|m| m.get(&sheet_name.to_lowercase()))
            .map(|v| v.as_slice());
        // .and_then(ensure_r_prefix_bound): same reasoning as workbook_root_attrs above --
        // this sheet's own tableParts/drawing/legacyDrawing/hyperlinks r:id restoration
        // below always hardcodes the literal `r:` prefix, so the source's root attrs must
        // guarantee that prefix is actually bound before being reused verbatim.
        let root_attrs = source_xml
            .and_then(|xml| reader::extract_root_attrs(xml, "worksheet"))
            .and_then(|attrs| reader::ensure_r_prefix_bound(&attrs));
        let sheet_pr = source_xml.and_then(|xml| reader::extract_raw_element(xml, "sheetPr"));
        let sheet_views = source_xml.and_then(|xml| reader::extract_raw_element(xml, "sheetViews"));
        let sheet_format_pr =
            source_xml.and_then(|xml| reader::extract_raw_element(xml, "sheetFormatPr"));
        let auto_filter = source_xml.and_then(|xml| reader::extract_raw_element(xml, "autoFilter"));
        let phonetic_pr = source_xml.and_then(|xml| reader::extract_raw_element(xml, "phoneticPr"));
        let conditional_formatting = source_xml
            .map(|xml| reader::extract_all_raw_elements(xml, "conditionalFormatting"))
            .unwrap_or_default();
        let data_validations =
            source_xml.and_then(|xml| reader::extract_raw_element(xml, "dataValidations"));
        let page_margins =
            source_xml.and_then(|xml| reader::extract_raw_element(xml, "pageMargins"));
        // <pageSetup> is unlike every other opaque fragment above: CT_PageSetup genuinely
        // CAN carry an r:id (referencing a printerSettings part) per the real XSD, even
        // though no fixture in this repo has ever shown that shape. A plain (no r:id)
        // pageSetup has no relationship dependency at all and is always safe to restore
        // verbatim, regardless of rels_survived; one WITH r:id is left unrestored until a
        // real fixture justifies wiring up the same rels_survived gate the elements below
        // use (this project's own hard gate: no writer code for a shape without fixture
        // evidence).
        let page_setup = source_xml
            .and_then(|xml| reader::extract_raw_element(xml, "pageSetup"))
            .filter(|el| !reader::root_tag_has_rid(el));
        // A relationship-backed element may only be restored when this sheet's own
        // .rels genuinely survived into THIS output -- `is_existing` alone only means
        // the sheet HAD an origin, not that its `.rels` specifically made it into
        // `passthrough` (worksheet-level `.rels` files are ordinary passthrough parts,
        // not writer-owned -- see `is_writer_owned_part`'s own doc comment). Splicing a
        // relationship-backed element back over a `.rels` that didn't survive would emit
        // a dangling `r:id`, a real Excel repair warning.
        let rels_survived = plan.is_existing
            && passthrough
                .iter()
                .any(|(name, _)| name == &plan.output_rels_name);
        // Location-only hyperlinks are always kept; r:id-bearing ones only when
        // rels_survived (see extract_hyperlinks' own doc comment).
        let hyperlinks = source_xml
            .map(|xml| reader::extract_hyperlinks(xml, rels_survived))
            .unwrap_or_default();
        let (table_parts, drawing, legacy_drawing) = if rels_survived {
            (
                source_xml.and_then(|xml| reader::extract_raw_element(xml, "tableParts")),
                source_xml.and_then(|xml| reader::extract_raw_element(xml, "drawing")),
                source_xml.and_then(|xml| reader::extract_raw_element(xml, "legacyDrawing")),
            )
        } else {
            (None, None, None)
        };
        let fragments = OpaqueWorksheetFragments {
            root_attrs: root_attrs.as_deref(),
            sheet_pr: sheet_pr.as_deref(),
            sheet_views: sheet_views.as_deref(),
            sheet_format_pr: sheet_format_pr.as_deref(),
            auto_filter: auto_filter.as_deref(),
            phonetic_pr: phonetic_pr.as_deref(),
            conditional_formatting: &conditional_formatting,
            data_validations: data_validations.as_deref(),
            hyperlinks: &hyperlinks,
            page_margins: page_margins.as_deref(),
            page_setup: page_setup.as_deref(),
            table_parts: table_parts.as_deref(),
            drawing: drawing.as_deref(),
            legacy_drawing: legacy_drawing.as_deref(),
        };

        zip.start_file(plan.output_part_name.as_str(), deflated)
            .map_err(|e| e.to_string())?;
        zip.write_all(
            build_xlsx_sheet(
                vm,
                sheet_name,
                &str_index,
                &fragments,
                style_override,
                row_style_override,
                column_style_override,
            )
            .as_bytes(),
        )
        .map_err(|e| e.to_string())?;
    }

    zip.start_file("xl/sharedStrings.xml", deflated)
        .map_err(|e| e.to_string())?;
    zip.write_all(build_xlsx_shared_strings(&shared_strings).as_bytes())
        .map_err(|e| e.to_string())?;

    zip.start_file("xl/styles.xml", deflated)
        .map_err(|e| e.to_string())?;
    zip.write_all(new_styles_bytes.as_deref().unwrap_or(styles_source))
        .map_err(|e| e.to_string())?;

    for (name, bytes) in &passthrough {
        zip.start_file(name.as_str(), deflated)
            .map_err(|e| e.to_string())?;
        zip.write_all(bytes).map_err(|e| e.to_string())?;
    }

    let data = zip.finish().map_err(|e| e.to_string())?.into_inner();
    std::fs::write(path, data).map_err(|e| e.to_string())?;
    Ok(())
}

fn build_xlsx_root_rels(carried_root_rels: &[(String, String)]) -> String {
    let mut out = String::from(concat!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n",
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n",
        "<Relationship Id=\"rId1\" ",
        "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" ",
        "Target=\"xl/workbook.xml\"/>\n",
    ));
    for (i, (ty, target)) in carried_root_rels.iter().enumerate() {
        out.push_str(&format!(
            "<Relationship Id=\"rId{}\" Type=\"{}\" Target=\"{}\"/>\n",
            i + 2,
            xml_escape(ty),
            xml_escape(target)
        ));
    }
    out.push_str("</Relationships>\n");
    out
}

const XLSX_STYLES: &str = concat!(
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n",
    "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n",
    "<fonts><font/></fonts>\n",
    "<fills><fill><patternFill/></fill><fill><patternFill patternType=\"gray125\"/></fill></fills>\n",
    "<borders><border/></borders>\n",
    "<cellStyleXfs><xf/></cellStyleXfs>\n",
    "<cellXfs><xf/></cellXfs>\n",
    "</styleSheet>\n",
);

/// Resolves `vm.pending_number_formats` (`set_number_format`, 0.15.0-A) against
/// `starting_styles` -- a loaded file's real `xl/styles.xml` bytes, or `XLSX_STYLES` for a
/// from-scratch `Vm()`; `save_xlsx_impl` picks whichever applies before calling this, the
/// same `.unwrap_or_else(...)` resolution it always used, just no longer assumed to be the
/// final answer. Returns `None` when there's nothing pending, so the caller can write
/// `starting_styles` back byte-for-byte unchanged, exactly as every save did before
/// `set_number_format` existed -- calling this at all, let alone its cost, is entirely
/// skipped in the common case.
///
/// Never mutates an existing `<xf>` record in place: every edited cell gets a fresh or
/// reused index found via `reader::with_num_fmt_id` + a byte-comparison dedup against
/// every `<xf>` seen so far (conservative -- can under-dedupe on attribute-order
/// differences, never over-dedupes two genuinely different records into one), since a
/// `cellXf` index can be shared by cells this call was never asked to touch ("no
/// shared-style mutation", 0.15.0's own safety requirement). `<fonts>`/`<fills>`/
/// `<borders>`/`<cellStyleXfs>` are never parsed or altered; only `<numFmts>` (appended
/// to, never rewritten) and `<cellXfs>` (existing entries copied byte-for-byte, new ones
/// appended) change.
///
/// Returns the effective per-sheet style-index map alongside the new document text
/// (a full clone of `vm.cell_style_indices` with just the touched cells overridden) --
/// `vm.cell_style_indices` itself is never mutated here (`save_xlsx_impl` takes `&Vm`,
/// not `&mut Vm`), so calling `save_workbook()` twice in a row re-resolves from the exact
/// same starting point both times and produces identical output.
fn resolve_pending_number_formats(
    vm: &Vm,
    starting_styles: &[u8],
) -> Option<(String, StyleIndexMap)> {
    if vm.pending_number_formats.values().all(|m| m.is_empty()) {
        return None;
    }
    let xml = String::from_utf8_lossy(starting_styles).into_owned();

    let cell_xfs_container = reader::extract_raw_element(&xml, "cellXfs");
    let mut xfs = reader::extract_cell_xfs(&xml);
    if xfs.is_empty() {
        // No <cellXfs> at all, or an empty/self-closing one -- every cell's `s` still
        // defaults to index 0, so synthesize the same bare record `XLSX_STYLES` itself
        // ships, rather than leaving index 0 with nothing to clone from.
        xfs.push("<xf/>".to_string());
    }

    let mut custom_formats = reader::custom_number_formats(&xml);
    let numfmts_container = reader::extract_raw_element(&xml, "numFmts");
    let mut new_numfmt_entries = String::new();

    let mut effective = vm.cell_style_indices.clone();

    for (sheet_key, edits) in &vm.pending_number_formats {
        if edits.is_empty() {
            continue;
        }
        let sheet_indices = effective.entry(sheet_key.clone()).or_default();
        for (&(row, col), format_code) in edits {
            let num_fmt_id = match reader::resolve_number_format_id(format_code, &custom_formats) {
                reader::ResolvedNumFmt::Existing(id) => id,
                reader::ResolvedNumFmt::New(id) => {
                    custom_formats.insert(id, format_code.clone());
                    new_numfmt_entries.push_str(&format!(
                        "<numFmt numFmtId=\"{id}\" formatCode=\"{}\"/>",
                        xml_escape(format_code)
                    ));
                    id
                }
            };
            let current_index = sheet_indices.get(&(row, col)).copied().unwrap_or(0) as usize;
            let current_xf = xfs
                .get(current_index)
                .cloned()
                .unwrap_or_else(|| "<xf/>".to_string());
            let candidate = reader::with_num_fmt_id(&current_xf, num_fmt_id);
            let new_index = match xfs.iter().position(|xf| xf == &candidate) {
                Some(i) => i as u32,
                None => {
                    xfs.push(candidate);
                    (xfs.len() - 1) as u32
                }
            };
            sheet_indices.insert((row, col), new_index);
        }
    }

    let new_cell_xfs = format!(
        "<cellXfs count=\"{}\">{}</cellXfs>",
        xfs.len(),
        xfs.concat()
    );
    let mut new_xml = match &cell_xfs_container {
        Some(container) => xml.replacen(container.as_str(), &new_cell_xfs, 1),
        // CT_StyleSheet requires <cellXfs> in every real file; defensive fallback only.
        None => xml.replacen("</styleSheet>", &format!("{new_cell_xfs}</styleSheet>"), 1),
    };

    if !new_numfmt_entries.is_empty() {
        // `custom_formats` now holds every pre-existing custom entry plus every one
        // minted this call -- its length is exactly the new <numFmts count="...">.
        let final_count = custom_formats.len();
        new_xml = match &numfmts_container {
            Some(container) => {
                let inner_start = container.find('>').map(|i| i + 1).unwrap_or(0);
                let inner_end = container.rfind("</numFmts>").unwrap_or(container.len());
                let inner = &container[inner_start..inner_end];
                let rebuilt = format!(
                    "<numFmts count=\"{final_count}\">{inner}{new_numfmt_entries}</numFmts>"
                );
                new_xml.replacen(container.as_str(), &rebuilt, 1)
            }
            None => {
                let block =
                    format!("<numFmts count=\"{final_count}\">{new_numfmt_entries}</numFmts>");
                // Schema order: numFmts is the very first child of <styleSheet>, before
                // <fonts> -- inserted right before it rather than appended at the end.
                new_xml.replacen("<fonts", &format!("{block}<fonts"), 1)
            }
        };
    }

    Some((new_xml, effective))
}

/// Resolves `vm.pending_style_attrs` (`set_style`, 0.15.0-B) against `starting_styles` and
/// `starting_indices` -- unlike `resolve_pending_number_formats`, these are NOT always
/// `vm.cell_style_indices`/the raw passthrough bytes: `save_xlsx_impl` chains this call
/// AFTER `resolve_pending_number_formats`, passing that pass's own `(new_xml, effective)`
/// output back in here, so a cell touched by BOTH `set_number_format` and `set_style`
/// before one save gets both changes rather than one silently overwriting the other (an
/// independent second pass starting fresh from the original bytes would only ever see
/// one). Returns `None` when there's nothing pending, so the caller writes
/// `starting_styles` straight through unchanged.
///
/// Same "no shared-style mutation" contract as the number-format pass, now spanning FOUR
/// tables instead of one: `<fonts>`/`<fills>`/`<borders>` are found-or-appended (via
/// `reader::with_font_edit`/`build_solid_fill`/`with_border_edit` + a byte-comparison
/// dedup, same conservative under-dedupe-never-over-dedupe bias), then `<cellXfs>` is
/// found-or-appended pointing at the resulting font/fill/border indices --
/// `<cellStyleXfs>` itself is never mutated, only READ (see the `named_style` case below).
/// Literal RGB/ARGB colors only; a font/fill/border that already carries a theme-relative
/// color (`<color theme="N"/>`) keeps it verbatim when that property isn't the one being
/// edited (0.15.0-C's job to mint a new one).
///
/// A pending `named_style` request (0.15.0-C1) is resolved FIRST in the per-cell loop,
/// before any of the other five fields on the same edit: it looks the name up against
/// `<cellStyles>`, then REPLACES the candidate `<xf>` wholesale with a clone of the
/// referenced `<cellStyleXfs>` entry (plus `xfId` set) -- matching real Excel's own
/// behavior of baking the named style's font/fill/border/numFmt/alignment/protection
/// directly onto the cell's own `<cellXfs>` entry, confirmed against a real fixture (see
/// `StyleAttrEdit::named_style`'s own doc comment). Any font/fill/border/alignment/
/// protection ALSO requested on the same edit then applies on top of that replacement, in
/// the usual order below. Returns `Err` (aborting the whole save) if the requested name
/// isn't in this file's `<cellStyles>` -- no minting fallback, matching 0.15.0-C's
/// decision to exclude named-style CREATE from this phase entirely.
/// Extracts the four style-table lists a `StyleAttrEdit` can touch, with the same
/// single-default-entry fallback `resolve_pending_style_attrs` always used when a table
/// is entirely absent (a from-scratch `Vm()`'s minimal `XLSX_STYLES`, or a genuinely
/// malformed source) -- shared by every resolve pass that mints/reuses font/fill/
/// border/cellXf records (`resolve_pending_style_attrs`,
/// `resolve_pending_row_column_styles`), so both always see identical starting tables.
fn extract_style_tables(xml: &str) -> (Vec<String>, Vec<String>, Vec<String>, Vec<String>) {
    let mut fonts = reader::extract_records(xml, "fonts", "font");
    if fonts.is_empty() {
        fonts.push("<font/>".to_string());
    }
    let mut fills = reader::extract_records(xml, "fills", "fill");
    if fills.is_empty() {
        fills.push("<fill><patternFill patternType=\"none\"/></fill>".to_string());
    }
    let mut borders = reader::extract_records(xml, "borders", "border");
    if borders.is_empty() {
        borders.push("<border><left/><right/><top/><bottom/><diagonal/></border>".to_string());
    }
    let mut xfs = reader::extract_cell_xfs(xml);
    if xfs.is_empty() {
        xfs.push("<xf/>".to_string());
    }
    (fonts, fills, borders, xfs)
}

/// Re-serializes the four style tables back into `xml`'s `<styleSheet>`, replacing each
/// existing `<fonts>`/`<fills>`/`<borders>`/`<cellXfs>` element wholesale (or inserting
/// one at the correct schema position if the source never had it at all -- defensive
/// fallback only, every real file and `XLSX_STYLES` itself always has all four). Shared
/// by every resolve pass that can grow these tables.
fn reserialize_style_tables(
    xml: String,
    fonts: &[String],
    fills: &[String],
    borders: &[String],
    xfs: &[String],
) -> String {
    let mut new_xml = xml;
    let new_fonts = format!(
        "<fonts count=\"{}\">{}</fonts>",
        fonts.len(),
        fonts.concat()
    );
    new_xml = match reader::extract_raw_element(&new_xml, "fonts") {
        Some(old) => new_xml.replacen(old.as_str(), &new_fonts, 1),
        // Schema order: fonts precedes fills. Defensive fallback only -- every real file
        // and XLSX_STYLES itself always has <fonts>.
        None => new_xml.replacen("<fills", &format!("{new_fonts}<fills"), 1),
    };
    let new_fills = format!(
        "<fills count=\"{}\">{}</fills>",
        fills.len(),
        fills.concat()
    );
    new_xml = match reader::extract_raw_element(&new_xml, "fills") {
        Some(old) => new_xml.replacen(old.as_str(), &new_fills, 1),
        None => new_xml.replacen("<borders", &format!("{new_fills}<borders"), 1),
    };
    let new_borders = format!(
        "<borders count=\"{}\">{}</borders>",
        borders.len(),
        borders.concat()
    );
    new_xml = match reader::extract_raw_element(&new_xml, "borders") {
        Some(old) => new_xml.replacen(old.as_str(), &new_borders, 1),
        None => new_xml.replacen("<cellStyleXfs", &format!("{new_borders}<cellStyleXfs"), 1),
    };
    let new_cell_xfs = format!(
        "<cellXfs count=\"{}\">{}</cellXfs>",
        xfs.len(),
        xfs.concat()
    );
    new_xml = match reader::extract_raw_element(&new_xml, "cellXfs") {
        Some(old) => new_xml.replacen(old.as_str(), &new_cell_xfs, 1),
        None => new_xml.replacen("</styleSheet>", &format!("{new_cell_xfs}</styleSheet>"), 1),
    };
    new_xml
}

/// Applies one `StyleAttrEdit` against `current_xf_index`'s existing `<xf>` record --
/// clones it, merges in only the requested font/fill/border/alignment/protection/
/// named-style properties, find-or-appends the result, and returns its index. Shared by
/// `resolve_pending_style_attrs` (per-cell, 0.15.0-B) and
/// `resolve_pending_row_column_styles` (per-row/-column, 0.15.0-C2) -- both need the
/// identical clone-merge-dedup logic; only WHERE the resulting index gets stored
/// afterward differs between the two callers.
#[allow(clippy::too_many_arguments)]
fn apply_style_edit_to_xf(
    current_xf_index: usize,
    edit: &vm::StyleAttrEdit,
    xfs: &mut Vec<String>,
    fonts: &mut Vec<String>,
    fills: &mut Vec<String>,
    borders: &mut Vec<String>,
    cell_style_xfs: &[String],
    xml: &str,
) -> Result<u32, String> {
    let mut current_xf = xfs
        .get(current_xf_index)
        .cloned()
        .unwrap_or_else(|| "<xf/>".to_string());

    if let Some(name) = &edit.named_style {
        let xf_id = reader::named_style_xf_id(xml, name)
            .ok_or_else(|| format!("named style '{name}' not found in this file's <cellStyles>"))?;
        let style_xf = cell_style_xfs
            .get(xf_id as usize)
            .cloned()
            .unwrap_or_else(|| "<xf/>".to_string());
        current_xf = reader::with_attr(&style_xf, "xfId", &xf_id.to_string());
    }

    if let Some(font_edit) = &edit.font {
        let font_id = reader::span_attr_u32(&current_xf, "fontId") as usize;
        let current_font = fonts
            .get(font_id)
            .cloned()
            .unwrap_or_else(|| "<font/>".to_string());
        let candidate = reader::with_font_edit(&current_font, font_edit);
        let new_font_id = match fonts.iter().position(|f| f == &candidate) {
            Some(i) => i as u32,
            None => {
                fonts.push(candidate);
                (fonts.len() - 1) as u32
            }
        };
        current_xf = reader::with_attr(&current_xf, "fontId", &new_font_id.to_string());
        current_xf = reader::with_attr(&current_xf, "applyFont", "1");
    }
    if let Some(fill_edit) = &edit.fill {
        let candidate = reader::build_solid_fill(&fill_edit.color_argb);
        let new_fill_id = match fills.iter().position(|f| f == &candidate) {
            Some(i) => i as u32,
            None => {
                fills.push(candidate);
                (fills.len() - 1) as u32
            }
        };
        current_xf = reader::with_attr(&current_xf, "fillId", &new_fill_id.to_string());
        current_xf = reader::with_attr(&current_xf, "applyFill", "1");
    }
    if let Some(border_edit) = &edit.border {
        let border_id = reader::span_attr_u32(&current_xf, "borderId") as usize;
        let current_border = borders.get(border_id).cloned().unwrap_or_else(|| {
            "<border><left/><right/><top/><bottom/><diagonal/></border>".to_string()
        });
        let candidate = reader::with_border_edit(&current_border, border_edit);
        let new_border_id = match borders.iter().position(|b| b == &candidate) {
            Some(i) => i as u32,
            None => {
                borders.push(candidate);
                (borders.len() - 1) as u32
            }
        };
        current_xf = reader::with_attr(&current_xf, "borderId", &new_border_id.to_string());
        current_xf = reader::with_attr(&current_xf, "applyBorder", "1");
    }
    if let Some(alignment_edit) = &edit.alignment {
        let new_alignment = reader::merged_alignment_span(&current_xf, alignment_edit);
        current_xf = reader::with_ordered_child(
            &current_xf,
            "alignment",
            &reader::XF_CHILD_ORDER,
            Some(&new_alignment),
        );
        current_xf = reader::with_attr(&current_xf, "applyAlignment", "1");
    }
    if let Some(protection_edit) = &edit.protection {
        let new_protection = reader::merged_protection_span(&current_xf, protection_edit);
        current_xf = reader::with_ordered_child(
            &current_xf,
            "protection",
            &reader::XF_CHILD_ORDER,
            Some(&new_protection),
        );
        current_xf = reader::with_attr(&current_xf, "applyProtection", "1");
    }

    Ok(match xfs.iter().position(|xf| xf == &current_xf) {
        Some(i) => i as u32,
        None => {
            xfs.push(current_xf);
            (xfs.len() - 1) as u32
        }
    })
}

fn resolve_pending_style_attrs(
    vm: &Vm,
    starting_styles: &[u8],
    starting_indices: &StyleIndexMap,
) -> Result<Option<(String, StyleIndexMap)>, String> {
    if vm.pending_style_attrs.values().all(|m| m.is_empty()) {
        return Ok(None);
    }
    let xml = String::from_utf8_lossy(starting_styles).into_owned();
    let (mut fonts, mut fills, mut borders, mut xfs) = extract_style_tables(&xml);
    let cell_style_xfs = reader::extract_records(&xml, "cellStyleXfs", "xf");

    let mut effective = starting_indices.clone();

    for (sheet_key, edits) in &vm.pending_style_attrs {
        if edits.is_empty() {
            continue;
        }
        let sheet_indices = effective.entry(sheet_key.clone()).or_default();
        for (&(row, col), edit) in edits {
            let current_index = sheet_indices.get(&(row, col)).copied().unwrap_or(0) as usize;
            let new_index = apply_style_edit_to_xf(
                current_index,
                edit,
                &mut xfs,
                &mut fonts,
                &mut fills,
                &mut borders,
                &cell_style_xfs,
                &xml,
            )?;
            sheet_indices.insert((row, col), new_index);
        }
    }

    let new_xml = reserialize_style_tables(xml, &fonts, &fills, &borders, &xfs);
    Ok(Some((new_xml, effective)))
}

type RowStyleIndexMap = std::collections::HashMap<String, std::collections::HashMap<u32, u32>>;
type ColumnStyleRangeMap = std::collections::HashMap<String, Vec<(u32, u32, u32)>>;

/// Resolves `vm.pending_row_styles`/`vm.pending_column_styles` (`set_row_style`/
/// `set_column_style`, 0.15.0-C2) -- chained after `resolve_pending_style_attrs` at save
/// time, sharing the SAME font/fill/border/cellXf tables (both mint into the same
/// `<cellXfs>`, so an independent pass starting fresh here would risk near-duplicate
/// records or index collisions with whatever the cell-level pass just appended -- same
/// "must chain, not run independently" reasoning as `resolve_pending_style_attrs`'s own
/// doc comment against `resolve_pending_number_formats`).
///
/// Row styles resolve like `pending_style_attrs`'s per-cell case (a plain per-row index
/// map). Column styles are range-shaped (`(min, max, style_index)`, like
/// `column_widths`) -- setting ONE column's style splits any existing range that
/// contains it into up to two remaining sub-ranges (before/after), then appends a fresh
/// singleton `(col, col, new_index)` for the touched column, mirroring how real Excel
/// itself fragments a `<cols>` run when an individual column's formatting changes.
fn resolve_pending_row_column_styles(
    vm: &Vm,
    starting_styles: &[u8],
    starting_row_styles: &RowStyleIndexMap,
    starting_column_styles: &ColumnStyleRangeMap,
) -> Result<Option<(String, RowStyleIndexMap, ColumnStyleRangeMap)>, String> {
    if vm.pending_row_styles.values().all(|m| m.is_empty())
        && vm.pending_column_styles.values().all(|m| m.is_empty())
    {
        return Ok(None);
    }
    let xml = String::from_utf8_lossy(starting_styles).into_owned();
    let (mut fonts, mut fills, mut borders, mut xfs) = extract_style_tables(&xml);
    let cell_style_xfs = reader::extract_records(&xml, "cellStyleXfs", "xf");

    let mut effective_rows = starting_row_styles.clone();
    let mut effective_cols = starting_column_styles.clone();

    for (sheet_key, edits) in &vm.pending_row_styles {
        if edits.is_empty() {
            continue;
        }
        let sheet_rows = effective_rows.entry(sheet_key.clone()).or_default();
        for (&row, edit) in edits {
            let current_index = sheet_rows.get(&row).copied().unwrap_or(0) as usize;
            let new_index = apply_style_edit_to_xf(
                current_index,
                edit,
                &mut xfs,
                &mut fonts,
                &mut fills,
                &mut borders,
                &cell_style_xfs,
                &xml,
            )?;
            sheet_rows.insert(row, new_index);
        }
    }

    for (sheet_key, edits) in &vm.pending_column_styles {
        if edits.is_empty() {
            continue;
        }
        let sheet_cols = effective_cols.entry(sheet_key.clone()).or_default();
        for (&col, edit) in edits {
            let current_index = sheet_cols
                .iter()
                .find(|&&(min, max, _)| min <= col && col <= max)
                .map(|&(_, _, s)| s)
                .unwrap_or(0) as usize;
            let new_index = apply_style_edit_to_xf(
                current_index,
                edit,
                &mut xfs,
                &mut fonts,
                &mut fills,
                &mut borders,
                &cell_style_xfs,
                &xml,
            )?;
            let mut remaining: Vec<(u32, u32, u32)> = Vec::with_capacity(sheet_cols.len() + 1);
            for &(min, max, style) in sheet_cols.iter() {
                if col < min || col > max {
                    remaining.push((min, max, style));
                    continue;
                }
                if min < col {
                    remaining.push((min, col - 1, style));
                }
                if col < max {
                    remaining.push((col + 1, max, style));
                }
            }
            remaining.push((col, col, new_index));
            *sheet_cols = remaining;
        }
    }

    let new_xml = reserialize_style_tables(xml, &fonts, &fills, &borders, &xfs);
    Ok(Some((new_xml, effective_rows, effective_cols)))
}

/// Resolves `vm.pending_style_copies` (`copy_style`, 0.15.0-C1) against `starting_indices`
/// -- chained LAST at save time, after BOTH `resolve_pending_number_formats` and
/// `resolve_pending_style_attrs`, so a `copy_style` call picks up whatever the source cell
/// resolved to from either of those passes automatically (see `Vm::pending_style_copies`'s
/// own doc comment). Pure index aliasing: a destination cell is pointed at EXACTLY the
/// same style index the source cell already resolves to -- multiple cells sharing one
/// index is exactly what a style table already supports, so this never mints a new
/// `<xf>`/font/fill/border/numFmt record and never touches the styles.xml bytes
/// themselves. Returns `None` when there's nothing pending, so the caller reuses
/// `starting_indices`/`starting_styles` unchanged.
///
/// Snapshots each sheet's indices before applying any copy in that sheet's batch, rather
/// than reading and writing the same live map -- so if one `copy_style` call's destination
/// happens to be another (same-batch) call's source, resolution isn't order-dependent on
/// `HashMap` iteration order.
fn resolve_pending_style_copies(
    vm: &Vm,
    starting_indices: &StyleIndexMap,
) -> Option<StyleIndexMap> {
    if vm.pending_style_copies.values().all(|m| m.is_empty()) {
        return None;
    }
    let mut effective = starting_indices.clone();
    for (sheet_key, copies) in &vm.pending_style_copies {
        if copies.is_empty() {
            continue;
        }
        let source_snapshot = effective.get(sheet_key).cloned().unwrap_or_default();
        let sheet_indices = effective.entry(sheet_key.clone()).or_default();
        for (&dest, &src) in copies {
            let src_index = source_snapshot.get(&src).copied().unwrap_or(0);
            sheet_indices.insert(dest, src_index);
        }
    }
    Some(effective)
}

fn build_xlsx_content_types(
    worksheet_plans: &[WorksheetOutputPlan],
    is_xlsm_output: bool,
    carried_overrides: &[(String, String)],
) -> String {
    let mut out = String::from(concat!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n",
        "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\n",
        "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>\n",
        "<Default Extension=\"xml\" ContentType=\"application/xml\"/>\n",
    ));
    // The macro-enabled content type is a property of the FILE FORMAT (.xlsm), not of
    // whether a VBA project happens to be present right now -- confirmed live against a
    // real Excel-authored .xlsm fixture with zero VBA content, which still declares
    // macroEnabled.main+xml for workbook.xml. Using has_vba here instead (as this
    // function did originally) declares the WRONG content type for the very common case
    // of an .xlsm workbook with no macros, which Excel treats as a fatal extension/format
    // mismatch and refuses to open at all -- not even a repair prompt.
    let workbook_ct = if is_xlsm_output {
        "application/vnd.ms-excel.sheet.macroEnabled.main+xml"
    } else {
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
    };
    out.push_str(&format!(
        "<Override PartName=\"/xl/workbook.xml\" ContentType=\"{}\"/>\n",
        workbook_ct
    ));
    for plan in worksheet_plans {
        out.push_str(&format!(
            "<Override PartName=\"/{}\" \
             ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>\n",
            plan.output_part_name
        ));
    }
    out.push_str(concat!(
        "<Override PartName=\"/xl/sharedStrings.xml\" ",
          "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>\n",
        "<Override PartName=\"/xl/styles.xml\" ",
          "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>\n",
    ));
    for (part_name, ct) in carried_overrides {
        out.push_str(&format!(
            "<Override PartName=\"{}\" ContentType=\"{}\"/>\n",
            xml_escape(part_name),
            xml_escape(ct)
        ));
    }
    out.push_str("</Types>\n");
    out
}

/// `origins` (0.10.0-A, `Vm::worksheet_origins`) lets a surviving sheet keep its
/// original `sheetId` across a save instead of always renumbering by current position --
/// see `vm::WorksheetOrigin`'s own doc comment for why (it's the one identifier
/// `snapshot.rs`'s `stable_id` already treats as cross-save-stable, and this writer was
/// the reason that promise didn't actually hold). `r:id="rIdN"` numbering stays purely
/// positional either way -- it's entirely internal to this writer's own
/// workbook.xml/workbook.xml.rels pair (see `build_xlsx_workbook_rels`), unrelated to a
/// sheet's `sheetId` identity, so it doesn't need to track a sheet's origin at all.
/// 0.10.0-C's opaque-fragment passthrough bundle for `xl/workbook.xml`, slices C1+C2:
/// raw text captured from the SOURCE workbook XML (see `save_xlsx_impl`'s
/// `workbook_source_xml`), re-emitted verbatim at the correct `CT_Workbook` schema
/// position (§8) rather than reconstructed. `None` for every field when there's no
/// passthrough source (new-from-scratch `Vm`, or an `.ods` source).
///
/// `book_views`'s `<workbookView>` can carry `activeTab`/`firstSheet` (sheet-position
/// indices) in principle, but every real fixture checked before adding this field
/// omits both (XSD default 0) -- a verbatim copy is correct against all currently
/// known evidence. See `check_workbook_elements()`'s docstring in
/// `compat/oracle-excel-com/mechanical_check.py` for the full reasoning on why no
/// gating logic was built for the unevidenced case.
///
/// `defined_names` is `None` whenever ANY sheet present at load time is no longer in
/// `Vm::sheets` (i.e. `Sheets(...).Delete` ran) -- a `<definedName>`'s `localSheetId` is
/// a 0-based index into `<sheets>`, so a deletion can leave every remaining
/// `localSheetId` pointing at a different sheet than it originally meant. Caller
/// (`save_xlsx_impl`) computes this gate, not `build_xlsx_workbook` itself, since it
/// needs direct access to `Vm::worksheet_origins`/`Vm::sheets`. Dropping the whole
/// element rather than remapping or pruning individual names is a deliberate
/// simplification -- see docs/xlsx-worksheet-preservation-0.10.0-design.md §10's C3
/// entry for why partial remapping is future work, not this slice's scope.
#[derive(Default)]
struct OpaqueWorkbookFragments<'a> {
    /// Source's root `<workbook ...>` tag's raw attribute string (namespace
    /// declarations, `mc:Ignorable`, ...) — replaces the hardcoded minimal
    /// `xmlns=".."`/`xmlns:r=".."` root tag when available. The caller
    /// (`save_xlsx_impl`) runs this through `reader::ensure_r_prefix_bound` before
    /// storing it here, which is what actually guarantees `r:` resolves correctly for
    /// the writer's own `r:id` emission below — a source binding the relationships
    /// namespace to some OTHER prefix (fully valid OOXML) is not, on its own, enough to
    /// guarantee that; see that function's own doc comment for the real report this
    /// fixed.
    root_attrs: Option<&'a str>,
    workbook_pr: Option<&'a str>,
    book_views: Option<&'a str>,
    defined_names: Option<&'a str>,
    calc_pr: Option<&'a str>,
    ext_lst: Option<&'a str>,
}

fn build_xlsx_workbook(
    worksheet_plans: &[WorksheetOutputPlan],
    fragments: &OpaqueWorkbookFragments,
) -> String {
    let mut out = String::from("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n");
    match fragments.root_attrs {
        Some(attrs) => {
            out.push_str("<workbook ");
            out.push_str(attrs);
            out.push_str(">\n");
        }
        None => out.push_str(concat!(
            "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" ",
            "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n",
        )),
    }
    // CT_Workbook order (§8): fileVersion, fileSharing, workbookPr, workbookProtection,
    // bookViews, sheets, functionGroups, externalReferences, definedNames, calcPr, ...
    for fragment in [fragments.workbook_pr, fragments.book_views]
        .into_iter()
        .flatten()
    {
        out.push_str(fragment);
        out.push('\n');
    }
    out.push_str("<sheets>\n");
    for plan in worksheet_plans {
        out.push_str(&format!(
            "<sheet name=\"{}\" sheetId=\"{}\" r:id=\"{}\"/>\n",
            xml_escape(&plan.display_name),
            plan.sheet_id,
            plan.workbook_rel_id
        ));
    }
    out.push_str("</sheets>\n");
    // CT_Workbook order (§8): ... sheets, functionGroups, externalReferences,
    // definedNames, calcPr, ...
    for fragment in [
        fragments.defined_names,
        fragments.calc_pr,
        fragments.ext_lst,
    ]
    .into_iter()
    .flatten()
    {
        out.push_str(fragment);
        out.push('\n');
    }
    out.push_str("</workbook>\n");
    out
}

/// Resolves a relationship `Target` (as written in some `.rels` file, e.g.
/// `"theme/theme1.xml"` or `"../customXml/item1.xml"`) against a zip-entry path,
/// collapsing `.`/`..` segments -- mirrors OOXML's own "relative to the part's own
/// directory" resolution rule (the same rule `compat/oracle-excel-com/
/// mechanical_check.py`'s `_normalize_part_path` implements independently, for the same
/// reason: a target can legitimately climb out of `xl/`).
fn normalize_part_path(joined: &str) -> String {
    let mut parts: Vec<&str> = Vec::new();
    for seg in joined.split('/') {
        match seg {
            "" | "." => {}
            ".." => {
                parts.pop();
            }
            _ => parts.push(seg),
        }
    }
    parts.join("/")
}

fn build_xlsx_workbook_rels(
    worksheet_plans: &[WorksheetOutputPlan],
    has_vba: bool,
    carried_rels: &[(String, String)],
) -> String {
    let mut out = String::from(concat!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n",
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n",
    ));
    for plan in worksheet_plans {
        // Target is relative to xl/ (this rels file's own part is xl/_rels/workbook.xml.rels),
        // matching how every other target below is already written relative to xl/.
        let target = plan
            .output_part_name
            .strip_prefix("xl/")
            .unwrap_or(&plan.output_part_name);
        out.push_str(&format!(
            "<Relationship Id=\"{}\" \
             Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" \
             Target=\"{}\"/>\n",
            plan.workbook_rel_id, target
        ));
    }
    let ss_id = worksheet_plans.len() + 1;
    let styles_id = worksheet_plans.len() + 2;
    out.push_str(&format!(
        "<Relationship Id=\"rId{}\" \
         Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" \
         Target=\"sharedStrings.xml\"/>\n",
        ss_id
    ));
    out.push_str(&format!(
        "<Relationship Id=\"rId{}\" \
         Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" \
         Target=\"styles.xml\"/>\n",
        styles_id
    ));
    let mut next_id = worksheet_plans.len() + 3;
    if has_vba {
        out.push_str(&format!(
            "<Relationship Id=\"rId{}\" \
             Type=\"http://schemas.microsoft.com/office/2006/relationships/vbaProject\" \
             Target=\"vbaProject.bin\"/>\n",
            next_id
        ));
        next_id += 1;
    }
    for (ty, target) in carried_rels {
        out.push_str(&format!(
            "<Relationship Id=\"rId{}\" Type=\"{}\" Target=\"{}\"/>\n",
            next_id,
            xml_escape(ty),
            xml_escape(target)
        ));
        next_id += 1;
    }
    out.push_str("</Relationships>\n");
    out
}

/// 0.10.0-B's opaque-fragment passthrough bundle for one worksheet: raw text captured
/// from the SOURCE worksheet XML (see `save_xlsx_impl`'s `sheet_source_xml`), re-emitted
/// verbatim at the correct `CT_Worksheet` schema position (§8) rather than reconstructed.
/// Every field is `None` for a sheet with no known `WorksheetOrigin` (new sheet, or an
/// `.ods` source), in which case `build_xlsx_sheet`'s behavior is unchanged from before
/// 0.10.0-B.
#[derive(Default)]
struct OpaqueWorksheetFragments<'a> {
    /// Source's root `<worksheet ...>` tag's raw attribute string (namespace
    /// declarations, `mc:Ignorable`, `xr:uid`, ...) — replaces the hardcoded minimal
    /// `xmlns=".."` root tag when available. See
    /// docs/xlsx-worksheet-preservation-0.10.0-design.md §8.
    root_attrs: Option<&'a str>,
    sheet_pr: Option<&'a str>,
    /// `<sheetViews>` (freeze panes via `<pane>`, active-cell `<selection>`).
    sheet_views: Option<&'a str>,
    sheet_format_pr: Option<&'a str>,
    /// `<autoFilter ref="..."><filterColumn .../>...</autoFilter>` -- unlike
    /// `table_parts`/`drawing`/`legacy_drawing` below, `CT_AutoFilter` has no `r:id` at
    /// all (confirmed against openpyxl's own writer, `worksheet/_writer.py`'s
    /// `write_tail` schema-order docstring and `write_filter`), so it needs no
    /// `rels_survived` gate -- same unconditional-passthrough treatment as
    /// `data_validations` below. Byte-preservation only: no structured `Vm` state, no
    /// create/remove/filter-type API (that's `0.16.0`, see ROADMAP.md's known gap 28).
    auto_filter: Option<&'a str>,
    phonetic_pr: Option<&'a str>,
    /// Every top-level `<conditionalFormatting sqref="...">...</conditionalFormatting>`
    /// block, in document order (`CT_Worksheet` allows more than one — one per distinct
    /// range/rule-set is the common real shape, see `fixture3`/`fixture4`). Unconditional,
    /// non-relationship-backed passthrough, same treatment as `data_validations` below —
    /// a `dxfId` a rule references stays valid regardless of any style mutation, since
    /// nothing in this codebase's `<cellXfs>`/`<fonts>`/`<fills>`/`<borders>` resolve
    /// passes ever touches `xl/styles.xml`'s `<dxfs>` (verified: no `dxf` reference
    /// anywhere in `resolve_pending_*`). Byte-preservation only: no structured `Vm`
    /// state, no create/edit API (that's `0.16.0`, a separate, much larger effort).
    conditional_formatting: &'a [String],
    data_validations: Option<&'a str>,
    /// Raw `<hyperlink .../>` spans (see `reader::extract_hyperlinks`) — NOT the whole
    /// source `<hyperlinks>` container. Location-only children are always included;
    /// r:id-bearing ones only when the caller passed `rels_survived` to
    /// `extract_hyperlinks`. `build_xlsx_sheet` synthesizes the `<hyperlinks>`/
    /// `</hyperlinks>` wrapper itself, based on whether this list is empty:
    /// `CT_Hyperlinks`' `<hyperlink>` child is `minOccurs="1"` (confirmed against the real
    /// XSD), so an empty container is invalid XML and must be omitted entirely rather
    /// than emitted as `<hyperlinks/>`.
    hyperlinks: &'a [String],
    page_margins: Option<&'a str>,
    /// Only ever `Some` for a `<pageSetup>` with NO `r:id` -- unlike `page_margins` above,
    /// `CT_PageSetup` genuinely can be relationship-backed, so the caller (`save_xlsx_impl`)
    /// filters those out via `reader::root_tag_has_rid` before this is ever populated.
    page_setup: Option<&'a str>,
    /// `<tableParts><tablePart r:id="..."/></tableParts>`, `<drawing r:id="..."/>`, and
    /// `<legacyDrawing r:id="..."/>` -- 0.10.0-D relationship-backed elements. Unlike every
    /// fragment above, the caller must only pass `Some` for these when the sheet's own
    /// `xl/worksheets/_rels/sheetN.xml.rels` genuinely survived into THIS save's output
    /// (see `save_xlsx_impl`'s `rels_survived` check) -- splicing one back over a `.rels`
    /// that didn't pass through would emit a dangling `r:id` reference, a real Excel
    /// repair warning and strictly worse than the prior silent inertness
    /// `check_source_references()`'s `SOURCE_REFERENCE_LOSS` was built to catch.
    table_parts: Option<&'a str>,
    drawing: Option<&'a str>,
    legacy_drawing: Option<&'a str>,
}

fn build_xlsx_sheet(
    vm: &Vm,
    sheet_name: &str,
    str_index: &std::collections::HashMap<String, usize>,
    fragments: &OpaqueWorksheetFragments,
    style_override: Option<&std::collections::HashMap<(u32, u32), u32>>,
    row_style_override: Option<&std::collections::HashMap<u32, u32>>,
    column_style_override: Option<&[(u32, u32, u32)]>,
) -> String {
    let mut out = String::from("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n");
    match fragments.root_attrs {
        Some(attrs) => {
            out.push_str("<worksheet ");
            out.push_str(attrs);
            out.push_str(">\n");
        }
        None => out.push_str(
            "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n",
        ),
    }
    // CT_Worksheet order (§8): sheetPr, dimension, sheetViews, sheetFormatPr, cols,
    // sheetData, ... — dimension is deliberately never emitted here (see design doc's
    // "dropped from 0.10.0-B's scope" note: it's derived from cell data, not opaque view
    // state, so carrying the source's stale value verbatim would be actively wrong after
    // a macro writes outside the original range).
    for fragment in [
        fragments.sheet_pr,
        fragments.sheet_views,
        fragments.sheet_format_pr,
    ]
    .into_iter()
    .flatten()
    {
        out.push_str(fragment);
        out.push('\n');
    }

    let sheet_key = sheet_name.to_lowercase();
    // `style_override` (0.15.0-A `set_number_format`) is a full clone of
    // `vm.cell_style_indices` with just the edited cells' indices replaced -- present
    // for EVERY sheet whenever any pending edit exists anywhere in the workbook (styles
    // are workbook-global), so a lookup miss here means this sheet genuinely has no
    // style-index entries at all, same as the direct `vm.cell_style_indices` fallback
    // would report.
    let style_indices = style_override.or_else(|| vm.cell_style_indices.get(&sheet_key));
    let visibility = vm.sheet_visibility.get(&sheet_key);
    let hidden_columns = visibility
        .map(|v| v.hidden_columns.as_slice())
        .unwrap_or(&[]);
    let hidden_rows = visibility.map(|v| v.hidden_rows.as_slice()).unwrap_or(&[]);
    let row_heights = vm.row_heights.get(&sheet_key);
    let column_widths = vm.column_widths.get(&sheet_key);
    // `row_style_override`/`column_style_override` (0.15.0-C2) follow the exact same
    // present-for-every-sheet-whenever-anything-is-pending-anywhere convention as
    // `style_override` above.
    let row_styles = row_style_override.or_else(|| vm.row_styles.get(&sheet_key));
    let column_styles =
        column_style_override.or_else(|| vm.column_styles.get(&sheet_key).map(|v| v.as_slice()));

    // <cols> — schema-ordered before <sheetData>. Merged by exact (min,max)
    // range rather than emitted as two independent passes: a source `<col>`
    // can carry `hidden="1"` and `customWidth="1" width="..."` together on
    // one element (the reader parses both attributes off the SAME element
    // into these two separate maps -- see `reader.rs`'s `"col" =>` arm), and
    // this is the inverse of that, restoring the combined shape when both
    // maps agree on the same range rather than splitting it into two
    // elements. A hidden interval and a width triple that DON'T share the
    // same exact range (real producers don't always coalesce ranges either,
    // confirmed via openpyxl -- see `internal_docs/openpyxl-gap-audit.md`'s
    // "Implementation notes for P2: row height / column width") land as
    // separate `<col>` elements, same as those independent, unmerged
    // real-world shapes.
    let mut col_attrs: ColAttrsMap = std::collections::BTreeMap::new();
    for iv in hidden_columns {
        col_attrs.entry((iv.start, iv.end)).or_default().0 = true;
    }
    if let Some(widths) = column_widths {
        for &(min, max, width) in widths {
            col_attrs.entry((min, max)).or_default().1 = Some(width);
        }
    }
    // `style=` (0.15.0-C2) merges onto the same exact-range convention as hidden/width
    // above -- a style range that doesn't share its exact (min,max) with a hidden/width
    // range lands as its own separate `<col>` element, same as the pre-existing
    // hidden/width mismatch case already does.
    if let Some(styles) = column_styles {
        for &(min, max, style) in styles {
            col_attrs.entry((min, max)).or_default().2 = Some(style);
        }
    }
    if !col_attrs.is_empty() {
        out.push_str("<cols>\n");
        for ((min, max), (hidden, width, style)) in col_attrs {
            let width_attr = width
                .map(|w| format!(" customWidth=\"1\" width=\"{w}\""))
                .unwrap_or_default();
            let style_attr = style.map(|s| format!(" style=\"{s}\"")).unwrap_or_default();
            let hidden_attr = if hidden { " hidden=\"1\"" } else { "" };
            out.push_str(&format!(
                "<col min=\"{min}\" max=\"{max}\"{width_attr}{style_attr}{hidden_attr}/>\n"
            ));
        }
        out.push_str("</cols>\n");
    }

    out.push_str("<sheetData>\n");

    if let Some(cells) = vm.get_sheet_cells(sheet_name) {
        // Group by row first to avoid O(max_row × total_cells) scanning.
        let mut by_row: std::collections::BTreeMap<u32, Vec<_>> = std::collections::BTreeMap::new();
        for (k @ &(r, c), v) in cells.iter() {
            if r > 0 && c > 0 {
                by_row.entry(r).or_default().push((k, v));
            }
        }
        // A hidden row (or one with an explicit height) with no cell data
        // still needs its own <row .../> element for the attribute to
        // actually apply to a real reader — hidden-ness/height are <row>
        // attributes, so an absent element is just default/visible/
        // default-height. Expanded per-row (not min/max like <col>) because
        // that's what a real <row>-element-per-row source already looks like.
        for iv in hidden_rows {
            for r in iv.start..=iv.end {
                by_row.entry(r).or_default();
            }
        }
        if let Some(heights) = row_heights {
            for &r in heights.keys() {
                by_row.entry(r).or_default();
            }
        }
        if let Some(styles) = row_styles {
            for &r in styles.keys() {
                by_row.entry(r).or_default();
            }
        }
        for (row, mut row_cells) in by_row {
            row_cells.sort_by_key(|&(&(_, c), _)| c);
            let row_hidden = hidden_rows
                .iter()
                .any(|iv| iv.start <= row && row <= iv.end);
            let hidden_attr = if row_hidden { " hidden=\"1\"" } else { "" };
            let height_attr = row_heights
                .and_then(|m| m.get(&row))
                .map(|ht| format!(" customHeight=\"1\" ht=\"{ht}\""))
                .unwrap_or_default();
            let style_attr = row_styles
                .and_then(|m| m.get(&row))
                .map(|s| format!(" s=\"{s}\" customFormat=\"1\""))
                .unwrap_or_default();

            out.push_str(&format!(
                "<row r=\"{row}\"{height_attr}{style_attr}{hidden_attr}>\n"
            ));
            for (&(r, c), content) in row_cells {
                let cell_ref = format!("{}{}", xlsx_col_letters(c), r);
                let style_idx = style_indices.and_then(|m| m.get(&(r, c)).copied());
                if let Some(xml) = xlsx_cell_xml(
                    &cell_ref,
                    &content.value,
                    str_index,
                    style_idx,
                    content.formula.as_deref(),
                ) {
                    out.push_str(&xml);
                    out.push('\n');
                }
            }
            out.push_str("</row>\n");
        }
    }

    out.push_str("</sheetData>\n");

    // <autoFilter> — schema-ordered after <sheetData>/before <mergeCells> (CT_Worksheet's
    // real sequence has sheetCalcPr/sheetProtection/protectedRanges/scenarios/autoFilter
    // between sheetData and mergeCells; none of the earlier ones are emitted here, but
    // autoFilter itself is byte-preserved verbatim — see `OpaqueWorksheetFragments::
    // auto_filter`'s doc comment). Verified against openpyxl's own writer order
    // (`worksheet/_writer.py`'s `write_tail`), not guessed.
    if let Some(af) = fragments.auto_filter {
        out.push_str(af);
        out.push('\n');
    }

    // <mergeCells> — schema-ordered after <sheetData>.
    if let Some(merges) = vm.merged_ranges.get(&sheet_key)
        && !merges.is_empty()
    {
        out.push_str(&format!("<mergeCells count=\"{}\">\n", merges.len()));
        for rect in merges {
            out.push_str(&format!(
                "<mergeCell ref=\"{}\"/>\n",
                merge_rect_to_a1(rect)
            ));
        }
        out.push_str("</mergeCells>\n");
    }

    // CT_Worksheet order (§8): mergeCells, phoneticPr, conditionalFormatting,
    // dataValidations, hyperlinks, printOptions, pageMargins, ... — printOptions has no
    // fixture evidence yet, still not emitted. conditionalFormatting's real position
    // (immediately after phoneticPr, before dataValidations) is confirmed against real
    // bytes, not just this comment's own claim — both fixture3 and fixture4 show exactly
    // this order. `<dxfs>` (which a `dxfId` here may reference) is never touched by any
    // `<cellXfs>`/`<fonts>`/`<fills>`/`<borders>` resolve pass, so a preserved rule's
    // `dxfId` stays valid regardless of any style mutation — see
    // `OpaqueWorksheetFragments::conditional_formatting`'s own doc comment.
    if let Some(pp) = fragments.phonetic_pr {
        out.push_str(pp);
        out.push('\n');
    }
    for cf in fragments.conditional_formatting {
        out.push_str(cf);
        out.push('\n');
    }
    if let Some(dv) = fragments.data_validations {
        out.push_str(dv);
        out.push('\n');
    }

    // <hyperlinks> — reconstructed from filtered children, not a single opaque blob
    // (unlike every other fragment above/below). `CT_Hyperlinks`' own `<hyperlink>`
    // child is `minOccurs="1"` (confirmed against the real XSD): if every hyperlink was
    // r:id-backed and got filtered out, the container itself must be omitted entirely --
    // an empty `<hyperlinks/>` would be invalid XML, not merely "nothing to show".
    if !fragments.hyperlinks.is_empty() {
        out.push_str("<hyperlinks>\n");
        for hyperlink in fragments.hyperlinks {
            out.push_str(hyperlink);
            out.push('\n');
        }
        out.push_str("</hyperlinks>\n");
    }

    if let Some(pm) = fragments.page_margins {
        out.push_str(pm);
        out.push('\n');
    }

    if let Some(ps) = fragments.page_setup {
        out.push_str(ps);
        out.push('\n');
    }

    // CT_Worksheet order (§8): ... pageMargins, pageSetup, headerFooter, rowBreaks,
    // colBreaks, customProperties, cellWatches, ignoredErrors, smartTags, drawing,
    // legacyDrawing, legacyDrawingHF, drawingHF, picture, oleObjects, controls,
    // webPublishItems, tableParts, extLst. headerFooter..smartTags aren't emitted yet, so
    // drawing/legacyDrawing's schema-correct position is still simply "right after
    // pageSetup" until one of those unemitted slots is restored too.
    if let Some(d) = fragments.drawing {
        out.push_str(d);
        out.push('\n');
    }
    if let Some(ld) = fragments.legacy_drawing {
        out.push_str(ld);
        out.push('\n');
    }

    if let Some(tp) = fragments.table_parts {
        out.push_str(tp);
        out.push('\n');
    }

    out.push_str("</worksheet>\n");
    out
}

// `style_idx` is the cell's original `s="N"` index (Milestone: safe round-trip,
// style-index preservation) -- `Some` when this cell survived from a passthrough
// source (see `Vm::cell_style_indices`), `None` for a cell built purely in-VBA.
// Always safe to re-emit unchanged: no VBA statement in this VM ever mutates a
// cell's style (see `Vm::cell_style_indices`'s own doc comment).
fn xlsx_cell_xml(
    cell_ref: &str,
    v: &Variant,
    str_index: &std::collections::HashMap<String, usize>,
    style_idx: Option<u32>,
    formula: Option<&str>,
) -> Option<String> {
    let s_attr = style_idx
        .map(|idx| format!(" s=\"{}\"", idx))
        .unwrap_or_default();
    // A loaded cell's formula text (no leading `=`, matching how <f> is written in the
    // file -- see WorkbookSheet::formulas' doc comment); a VBA-assigned formula
    // (Vm::set_cell_formula) may still carry one, so it's stripped here too rather than
    // trusting the source. Emitted before <v>, matching real Excel's own element order.
    let f_tag = formula
        .map(|f| format!("<f>{}</f>", xml_escape(f.trim().trim_start_matches('='))))
        .unwrap_or_default();
    match v {
        Variant::Integer(n) => Some(format!(
            "<c r=\"{}\"{}>{}<v>{}</v></c>",
            cell_ref, s_attr, f_tag, n
        )),
        Variant::Float(f) => Some(format!(
            "<c r=\"{}\"{}>{}<v>{}</v></c>",
            cell_ref, s_attr, f_tag, f
        )),
        Variant::Date(s) => Some(format!(
            "<c r=\"{}\"{}>{}<v>{}</v></c>",
            cell_ref, s_attr, f_tag, s
        )),
        Variant::Str(s) => {
            let idx = str_index[s.as_str()];
            Some(format!(
                "<c r=\"{}\"{} t=\"s\">{}<v>{}</v></c>",
                cell_ref, s_attr, f_tag, idx
            ))
        }
        // t="e" with the literal error text in <v>, never shared-string indexed -- matches
        // real Excel's own shape exactly (confirmed against fixture5_chart_image_freeze_
        // print.xlsm's D8, a real #VALUE! cell). Writing this as t="s" (the pre-fix
        // behavior) round-tripped the cell as an ordinary string, not an error.
        Variant::Error(e) => Some(format!(
            "<c r=\"{}\"{} t=\"e\">{}<v>{}</v></c>",
            cell_ref,
            s_attr,
            f_tag,
            e.as_str()
        )),
        Variant::Boolean(b) => Some(format!(
            "<c r=\"{}\"{} t=\"b\">{}<v>{}</v></c>",
            cell_ref,
            s_attr,
            f_tag,
            if *b { 1 } else { 0 }
        )),
        // A cell with no formula and no value carries no information worth writing --
        // omitted (`formula.map` -> `None`), as before. A FORMULA cell whose cached
        // value just happens to be Empty (freshly typed and not yet recalculated, or
        // -- 0.14.0-A2 -- a cross-sheet reference this engine doesn't evaluate) is NOT
        // the same as an empty cell and must never be silently dropped: "no cached
        // result" and "no formula" are different things. No <v> element is written at
        // all here, rather than fabricating a placeholder value -- valid per the OOXML
        // schema (<v> is optional) and avoids inventing a number that was never
        // computed. See `Vm::populate_from_sheets`'s matching reader-side fix -- a
        // formula-only cell now round-trips correctly on reload too, not just on save.
        Variant::Empty => formula.map(|_| format!("<c r=\"{}\"{}>{}</c>", cell_ref, s_attr, f_tag)),
        Variant::Null | Variant::Array(_) | Variant::VbaArray(_) | Variant::Record(_) => None,
    }
}

fn build_xlsx_shared_strings(strings: &[String]) -> String {
    let count = strings.len();
    let mut out = format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n\
         <sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" \
         count=\"{count}\" uniqueCount=\"{count}\">\n"
    );
    for s in strings {
        // Leading/trailing whitespace is only guaranteed to survive a round trip through a
        // spec-following XML consumer (real Excel included) when marked xml:space="preserve" —
        // the mirror-image, writer-side counterpart of the xml:space="preserve" handling
        // xlsx_sheet_cells already honors on read for `<v>` (see reader.rs).
        if s.trim() != s.as_str() {
            out.push_str(&format!(
                "<si><t xml:space=\"preserve\">{}</t></si>\n",
                xml_escape(s)
            ));
        } else {
            out.push_str(&format!("<si><t>{}</t></si>\n", xml_escape(s)));
        }
    }
    out.push_str("</sst>\n");
    out
}

#[cfg(test)]
mod shared_strings_tests {
    use super::build_xlsx_shared_strings;

    #[test]
    fn marks_leading_or_trailing_whitespace_as_xml_space_preserve() {
        let xml = build_xlsx_shared_strings(&[
            "plain".to_string(),
            " leading and trailing ".to_string(),
            "trailing ".to_string(),
        ]);
        assert!(xml.contains("<si><t>plain</t></si>"));
        assert!(xml.contains("<si><t xml:space=\"preserve\"> leading and trailing </t></si>"));
        assert!(xml.contains("<si><t xml:space=\"preserve\">trailing </t></si>"));
    }
}

fn xlsx_col_letters(mut col: u32) -> String {
    let mut bytes = Vec::new();
    while col > 0 {
        col -= 1;
        bytes.push(b'A' + (col % 26) as u8);
        col /= 26;
    }
    bytes.reverse();
    String::from_utf8(bytes).unwrap()
}

/// `((r1,c1),(r2,c2))` -> A1-style range string (e.g. `"B1:D1"`). Factored out
/// of `save_xlsx_impl`'s `<mergeCell ref="...">` writer since it's now also
/// used by `PyVm::merged_cells`.
fn merge_rect_to_a1(rect: &((u32, u32), (u32, u32))) -> String {
    let ((r1, c1), (r2, c2)) = *rect;
    format!(
        "{}{}:{}{}",
        xlsx_col_letters(c1),
        r1,
        xlsx_col_letters(c2),
        r2
    )
}

// ── ODS write ────────────────────────────────────────────────────────────────

fn save_ods_impl(vm: &Vm, path: &str) -> Result<(), String> {
    use std::io::{Cursor, Write};
    use zip::CompressionMethod;
    use zip::write::ZipWriter;

    let cursor = Cursor::new(Vec::<u8>::new());
    let mut zip = ZipWriter::new(cursor);
    let stored =
        zip::write::SimpleFileOptions::default().compression_method(CompressionMethod::Stored);
    let deflated =
        zip::write::SimpleFileOptions::default().compression_method(CompressionMethod::Deflated);

    // 1. mimetype (must be STORE and first entry per ODF spec)
    zip.start_file("mimetype", stored)
        .map_err(|e| e.to_string())?;
    zip.write_all(b"application/vnd.oasis.opendocument.spreadsheet")
        .map_err(|e| e.to_string())?;

    // 2. META-INF/manifest.xml
    let manifest = build_ods_manifest(vm);
    zip.start_file("META-INF/manifest.xml", deflated)
        .map_err(|e| e.to_string())?;
    zip.write_all(manifest.as_bytes())
        .map_err(|e| e.to_string())?;

    // 3. content.xml
    let content = build_ods_content(vm);
    zip.start_file("content.xml", deflated)
        .map_err(|e| e.to_string())?;
    zip.write_all(content.as_bytes())
        .map_err(|e| e.to_string())?;

    let data = zip.finish().map_err(|e| e.to_string())?.into_inner();
    std::fs::write(path, data).map_err(|e| e.to_string())?;
    Ok(())
}

fn build_ods_manifest(_vm: &Vm) -> String {
    let mut m = String::from(concat!(
        r#"<?xml version="1.0" encoding="UTF-8"?>"#,
        "\n",
        r#"<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2">"#,
        "\n",
        r#" <manifest:file-entry manifest:media-type="application/vnd.oasis.opendocument.spreadsheet" manifest:version="1.2" manifest:full-path="/"/>"#,
        "\n",
        r#" <manifest:file-entry manifest:media-type="text/xml" manifest:full-path="content.xml"/>"#,
        "\n",
    ));
    m.push_str("</manifest:manifest>\n");
    m
}

fn build_ods_content(vm: &Vm) -> String {
    let mut out = String::from(concat!(
        r#"<?xml version="1.0" encoding="UTF-8"?>"#,
        "\n",
        r#"<office:document-content"#,
        r#" xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0""#,
        r#" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0""#,
        r#" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0""#,
        r#" office:version="1.2">"#,
        "\n",
        r#"<office:body><office:spreadsheet>"#,
        "\n",
    ));

    for sheet_name in vm.sheet_names() {
        let escaped = xml_escape(&sheet_name);
        out.push_str(&format!("<table:table table:name=\"{}\">\n", escaped));

        if let Some(cells) = vm.get_sheet_cells(&sheet_name)
            && !cells.is_empty()
        {
            let max_row = cells.keys().map(|(r, _)| *r).max().unwrap_or(0);
            let max_col = cells.keys().map(|(_, c)| *c).max().unwrap_or(0);

            for r in 1..=max_row {
                out.push_str("<table:table-row>");
                for c in 1..=max_col {
                    let cell_xml = match cells.get(&(r, c)) {
                        None
                        | Some(vm::CellContent {
                            value: Variant::Empty,
                            ..
                        }) => "<table:table-cell/>".to_string(),
                        Some(content) => ods_cell_xml(&content.value),
                    };
                    out.push_str(&cell_xml);
                }
                out.push_str("</table:table-row>\n");
            }
        }

        out.push_str("</table:table>\n");
    }

    out.push_str("</office:spreadsheet></office:body>\n</office:document-content>\n");
    out
}

fn ods_cell_xml(v: &Variant) -> String {
    match v {
        Variant::Integer(n) => format!(
            r#"<table:table-cell office:value-type="float" office:value="{}"><text:p>{}</text:p></table:table-cell>"#,
            n, n
        ),
        Variant::Float(f) => format!(
            r#"<table:table-cell office:value-type="float" office:value="{}"><text:p>{}</text:p></table:table-cell>"#,
            f, f
        ),
        Variant::Str(s) => format!(
            r#"<table:table-cell office:value-type="string"><text:p>{}</text:p></table:table-cell>"#,
            xml_escape(s)
        ),
        Variant::Boolean(b) => {
            let bv = if *b { "true" } else { "false" };
            format!(
                r#"<table:table-cell office:value-type="boolean" office:boolean-value="{}"><text:p>{}</text:p></table:table-cell>"#,
                bv,
                if *b { "TRUE" } else { "FALSE" }
            )
        }
        Variant::Date(s) => format!(
            r#"<table:table-cell office:value-type="float" office:value="{}"><text:p>{}</text:p></table:table-cell>"#,
            s, s
        ),
        Variant::Error(e) => format!(
            r#"<table:table-cell office:value-type="string"><text:p>{}</text:p></table:table-cell>"#,
            xml_escape(e.as_str())
        ),
        Variant::Empty
        | Variant::Null
        | Variant::Array(_)
        | Variant::VbaArray(_)
        | Variant::Record(_) => "<table:table-cell/>".to_string(),
    }
}

pub(crate) fn xml_escape(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}

// ── Module ────────────────────────────────────────────────────────────────────

#[cfg(feature = "python")]
#[pymodule]
mod elixcee {
    #[pymodule_export]
    use super::{PyExcelError, PyVm, hello, load_workbook, run_macro};
}

#[cfg(test)]
mod tests {
    use super::*;
    use calamine::{Reader, Xlsx, open_workbook};

    #[test]
    // 3.14 is an arbitrary decimal test value for the save/load round trip, not π.
    #[allow(clippy::approx_constant)]
    fn test_save_workbook_roundtrip() {
        let mut vm = Vm::new();
        vm.cells_mut().insert(
            (1, 1),
            CellContent {
                formula: None,
                value: Variant::Integer(42),
            },
        );
        vm.cells_mut().insert(
            (2, 1),
            CellContent {
                formula: None,
                value: Variant::Str("hello".into()),
            },
        );
        vm.cells_mut().insert(
            (3, 1),
            CellContent {
                formula: None,
                value: Variant::Float(3.14),
            },
        );
        vm.cells_mut().insert(
            (4, 1),
            CellContent {
                formula: None,
                value: Variant::Boolean(true),
            },
        );

        let path = "/tmp/elixcee_test_roundtrip.xlsx";
        save_workbook_impl(&vm, path).expect("save should succeed");

        // Reload with calamine and verify
        let mut wb: Xlsx<_> = open_workbook(path).expect("open should succeed");
        let range = wb.worksheet_range("sheet1").expect("sheet1 should exist");
        let cells: Vec<_> = range.cells().collect();
        assert!(!cells.is_empty(), "saved file should have cells");
    }

    #[test]
    fn rewrite_defined_names_xml_rewrites_only_the_renamed_qualifier() {
        let xml = concat!(
            "<definedNames>",
            "<definedName name=\"MyRange\">Sheet1!$A$1:$A$3</definedName>",
            "<definedName name=\"Other\">Sheet2!$B$1</definedName>",
            "</definedNames>",
        );
        let renames: std::collections::HashMap<String, String> =
            [("sheet1".to_string(), "Renamed".to_string())]
                .into_iter()
                .collect();
        let out = rewrite_defined_names_xml(xml, &renames).unwrap();
        assert!(out.contains("<definedName name=\"MyRange\">Renamed!$A$1:$A$3</definedName>"));
        assert!(out.contains("<definedName name=\"Other\">Sheet2!$B$1</definedName>"));
    }

    #[test]
    fn rewrite_defined_names_xml_preserves_print_titles_full_row_and_column() {
        let xml = concat!(
            "<definedNames>",
            "<definedName name=\"_xlnm.Print_Titles\" localSheetId=\"0\">",
            "Sheet1!$1:$3,Sheet1!$A:$A</definedName>",
            "</definedNames>",
        );
        let renames: std::collections::HashMap<String, String> =
            [("sheet1".to_string(), "Data".to_string())]
                .into_iter()
                .collect();
        let out = rewrite_defined_names_xml(xml, &renames).unwrap();
        assert!(out.contains("Data!$1:$3,Data!$A:$A"));
    }

    #[test]
    fn rewrite_defined_names_xml_unescapes_and_reescapes_a_formula_value() {
        // Raw wire form of a value containing `&` -- must be unescaped before
        // `formula::rewrite_defined_name_for_renames` sees it (a literal "&amp;" isn't
        // valid formula syntax) and re-escaped in the rewritten result.
        let xml = concat!(
            "<definedNames>",
            "<definedName name=\"X\">Sheet1!$A$1 &amp; \"text\"</definedName>",
            "</definedNames>",
        );
        let renames: std::collections::HashMap<String, String> =
            [("sheet1".to_string(), "Renamed".to_string())]
                .into_iter()
                .collect();
        let out = rewrite_defined_names_xml(xml, &renames).unwrap();
        assert!(out.contains("Renamed!$A$1 &amp; &quot;text&quot;"));
    }

    #[test]
    fn rewrite_defined_names_xml_drops_only_the_unconfirmable_name() {
        let xml = concat!(
            "<definedNames>",
            "<definedName name=\"Dynamic\">OFFSET(Sheet1!$A$1,0,0,COUNTA(Sheet1!$A:$A),1)</definedName>",
            "<definedName name=\"Plain\">Sheet1!$B$1</definedName>",
            "</definedNames>",
        );
        let renames: std::collections::HashMap<String, String> =
            [("sheet1".to_string(), "Renamed".to_string())]
                .into_iter()
                .collect();
        let out = rewrite_defined_names_xml(xml, &renames).unwrap();
        assert!(!out.contains("Dynamic"));
        assert!(out.contains("<definedName name=\"Plain\">Renamed!$B$1</definedName>"));
    }

    #[test]
    fn rewrite_defined_names_xml_returns_none_when_every_name_is_dropped() {
        let xml = concat!(
            "<definedNames>",
            "<definedName name=\"Dynamic\">OFFSET(Sheet1!$A$1,0,0,COUNTA(Sheet1!$A:$A),1)</definedName>",
            "</definedNames>",
        );
        let renames: std::collections::HashMap<String, String> =
            [("sheet1".to_string(), "Renamed".to_string())]
                .into_iter()
                .collect();
        assert_eq!(rewrite_defined_names_xml(xml, &renames), None);
    }

    #[test]
    fn row_height_and_column_width_survive_a_save_and_reload() {
        // Was previously dropped on EVERY save -- xlsx_worksheet_xml's
        // <row>/<cols> emission was fully regenerated from Vm.sheet_visibility
        // alone, not passthrough (see ROADMAP.md's known gaps, item 26, and
        // internal_docs/openpyxl-gap-audit.md's "Implementation notes for P2:
        // row height / column width"). No real fixture in this repo has a
        // genuine custom row height/column width to build this test from (the
        // same reason write support -- set_row_height/set_column_width -- is
        // still deferred), so this constructs the Vm state directly rather
        // than loading a file, matching this project's other pub(crate)-field
        // unit tests. A row/column with BOTH a custom size AND hidden state
        // (matching how a real <col>/<row> source element commonly carries
        // both attributes together -- see reader.rs's "col" => arm) is
        // included so the merged-attribute emission path is exercised too,
        // not just the size-only path.
        let mut vm = Vm::new();
        vm.cells_mut().insert(
            (5, 1),
            CellContent {
                formula: None,
                value: Variant::Integer(1),
            },
        );
        vm.row_heights
            .entry("sheet1".to_string())
            .or_default()
            .insert(5, 30.5);
        vm.row_heights
            .entry("sheet1".to_string())
            .or_default()
            .insert(10, 20.0);
        vm.sheet_visibility
            .entry("sheet1".to_string())
            .or_default()
            .hidden_rows
            .push(vm::Interval { start: 10, end: 10 });
        vm.column_widths
            .entry("sheet1".to_string())
            .or_default()
            .push((2, 2, 12.5));

        let path = "/tmp/elixcee_test_row_height_column_width_roundtrip.xlsx";
        save_workbook_impl(&vm, path).expect("save should succeed");

        let mut reloaded = Vm::new();
        reloaded
            .load_workbook_file(path)
            .expect("reload should succeed");
        let key = reloaded.resolve_sheet_key(None).unwrap();
        assert_eq!(reloaded.row_height_on_sheet(&key, 5), Some(30.5));
        assert_eq!(reloaded.column_width_on_sheet(&key, 2), Some(12.5));
        // The row that's BOTH hidden and explicitly sized: both attributes
        // must survive on the same <row> element, not just one of them.
        assert_eq!(reloaded.row_height_on_sheet(&key, 10), Some(20.0));
        assert_eq!(reloaded.hidden_rows_on_sheet(&key), vec![10]);
        // A row/column with neither is still None, not a stray zero/default.
        assert_eq!(reloaded.row_height_on_sheet(&key, 1), None);
        assert_eq!(reloaded.column_width_on_sheet(&key, 1), None);

        let _ = std::fs::remove_file(path);
    }

    #[test]
    fn resolve_pending_number_formats_returns_none_with_no_pending_edits() {
        let vm = Vm::new();
        assert!(resolve_pending_number_formats(&vm, XLSX_STYLES.as_bytes()).is_none());
    }

    #[test]
    fn resolve_pending_number_formats_reuses_a_builtin_id_against_a_from_scratch_stylesheet() {
        let mut vm = Vm::new();
        vm.set_number_format_on_sheet("sheet1", 1, 1, 1, 1, "#,##0.00");
        let (new_xml, effective) =
            resolve_pending_number_formats(&vm, XLSX_STYLES.as_bytes()).expect("pending edit");
        // id 4 is BUILTIN_NUMBER_FORMATS' own "#,##0.00" -- no <numFmts> should be minted.
        assert!(!new_xml.contains("<numFmts"));
        assert!(new_xml.contains("numFmtId=\"4\""));
        let idx = effective["sheet1"][&(1, 1)];
        assert_ne!(idx, 0); // index 0 is the untouched bare <xf/>, General.
    }

    #[test]
    fn resolve_pending_number_formats_mints_a_new_custom_numfmt_when_no_source_has_none() {
        let mut vm = Vm::new();
        vm.set_number_format_on_sheet("sheet1", 1, 1, 1, 1, "0.00\"kg\"");
        let (new_xml, _) =
            resolve_pending_number_formats(&vm, XLSX_STYLES.as_bytes()).expect("pending edit");
        assert!(new_xml.contains("<numFmts count=\"1\">"));
        assert!(new_xml.contains("numFmtId=\"164\" formatCode=\"0.00&quot;kg&quot;\""));
        assert!(new_xml.contains("numFmtId=\"164\"")); // used on the minted <xf> too.
        // Schema order: <numFmts> must precede <fonts>.
        assert!(new_xml.find("<numFmts").unwrap() < new_xml.find("<fonts").unwrap());
    }

    #[test]
    fn resolve_pending_number_formats_appends_to_an_existing_numfmts_block_without_disturbing_it() {
        let source = concat!(
            "<styleSheet>",
            "<numFmts count=\"1\"><numFmt numFmtId=\"164\" formatCode=\"0.00&quot;kg&quot;\"/></numFmts>",
            "<fonts><font/></fonts><fills><fill/></fills><borders><border/></borders>",
            "<cellStyleXfs><xf/></cellStyleXfs>",
            "<cellXfs count=\"1\"><xf/></cellXfs>",
            "</styleSheet>",
        );
        let mut vm = Vm::new();
        vm.set_number_format_on_sheet("sheet1", 1, 1, 1, 1, "0.00\"lb\"");
        let (new_xml, _) =
            resolve_pending_number_formats(&vm, source.as_bytes()).expect("pending edit");
        // The pre-existing 164/"kg" entry must survive untouched, and the new one gets 165.
        assert!(new_xml.contains("numFmtId=\"164\" formatCode=\"0.00&quot;kg&quot;\""));
        assert!(new_xml.contains("numFmtId=\"165\" formatCode=\"0.00&quot;lb&quot;\""));
        assert!(new_xml.contains("<numFmts count=\"2\">"));
    }

    #[test]
    fn resolve_pending_number_formats_reuses_an_existing_xf_dedup_not_a_duplicate() {
        let mut vm = Vm::new();
        vm.set_number_format_on_sheet("sheet1", 1, 1, 1, 1, "#,##0.00");
        vm.set_number_format_on_sheet("sheet1", 5, 5, 5, 5, "#,##0.00");
        let (_, effective) =
            resolve_pending_number_formats(&vm, XLSX_STYLES.as_bytes()).expect("pending edit");
        // Two DIFFERENT cells asking for the SAME format must land on the SAME new index,
        // not two separate ones -- both requests clone from the same source (bare <xf/>,
        // index 0) with the same target numFmtId, so they're byte-identical candidates.
        assert_eq!(effective["sheet1"][&(1, 1)], effective["sheet1"][&(5, 5)]);
    }

    #[test]
    fn resolve_pending_number_formats_never_mutates_the_index_of_an_untouched_cell() {
        // Two cells sharing style index 0 (both bare/General) on a from-scratch stylesheet
        // -- only one is touched. The untouched one's effective index must stay 0 (its
        // <xf> record, shared by both, is never mutated in place).
        let mut vm = Vm::new();
        vm.set_number_format_on_sheet("sheet1", 1, 1, 1, 1, "0%");
        let (_, effective) =
            resolve_pending_number_formats(&vm, XLSX_STYLES.as_bytes()).expect("pending edit");
        assert_ne!(effective["sheet1"][&(1, 1)], 0);
        // (2,2) was never touched -- absent from the effective override map entirely,
        // meaning callers fall back to `cell_style_indices` (still empty/0 for it).
        assert!(!effective.get("sheet1").unwrap().contains_key(&(2, 2)));
    }

    #[test]
    fn set_number_format_on_a_from_scratch_vm_survives_a_save_and_reload() {
        let mut vm = Vm::new();
        vm.cells_mut().insert(
            (1, 1),
            CellContent {
                formula: None,
                value: Variant::Float(1234.5),
            },
        );
        vm.set_number_format_on_sheet("sheet1", 1, 1, 1, 1, "#,##0.00");

        let path = "/tmp/elixcee_test_set_number_format_from_scratch.xlsx";
        save_workbook_impl(&vm, path).expect("save should succeed");

        let mut reloaded = Vm::new();
        reloaded
            .load_workbook_file(path)
            .expect("reload should succeed");
        assert_eq!(reloaded.get_cell_number_format(1, 1), Some("#,##0.00"));

        let _ = std::fs::remove_file(path);
    }

    #[test]
    fn set_number_format_mints_a_custom_numfmt_that_survives_a_save_and_reload() {
        let mut vm = Vm::new();
        vm.cells_mut().insert(
            (1, 1),
            CellContent {
                formula: None,
                value: Variant::Float(2.5),
            },
        );
        vm.set_number_format_on_sheet("sheet1", 1, 1, 1, 1, "0.00\"kg\"");

        let path = "/tmp/elixcee_test_set_number_format_custom.xlsx";
        save_workbook_impl(&vm, path).expect("save should succeed");

        let mut reloaded = Vm::new();
        reloaded
            .load_workbook_file(path)
            .expect("reload should succeed");
        assert_eq!(reloaded.get_cell_number_format(1, 1), Some("0.00\"kg\""));

        let _ = std::fs::remove_file(path);
    }

    #[test]
    fn set_number_format_leaves_an_untouched_sibling_cell_unaffected_after_a_reload() {
        // The concrete "no shared-style mutation" regression: two cells start sharing
        // style index 0 (both General, from-scratch default). Only one gets a format.
        let mut vm = Vm::new();
        vm.cells_mut().insert(
            (1, 1),
            CellContent {
                formula: None,
                value: Variant::Float(1.0),
            },
        );
        vm.cells_mut().insert(
            (2, 1),
            CellContent {
                formula: None,
                value: Variant::Float(2.0),
            },
        );
        vm.set_number_format_on_sheet("sheet1", 1, 1, 1, 1, "0%");

        let path = "/tmp/elixcee_test_set_number_format_sibling_unaffected.xlsx";
        save_workbook_impl(&vm, path).expect("save should succeed");

        let mut reloaded = Vm::new();
        reloaded
            .load_workbook_file(path)
            .expect("reload should succeed");
        assert_eq!(reloaded.get_cell_number_format(1, 1), Some("0%"));
        // General has no format string to report at all -- must stay None, not "0%".
        assert_eq!(reloaded.get_cell_number_format(2, 1), None);

        let _ = std::fs::remove_file(path);
    }

    #[test]
    fn set_number_format_get_cell_number_format_reflects_the_edit_before_any_save() {
        let mut vm = Vm::new();
        assert_eq!(vm.get_cell_number_format(1, 1), None);
        vm.set_number_format_on_sheet("sheet1", 1, 1, 1, 1, "0.00%");
        assert_eq!(vm.get_cell_number_format(1, 1), Some("0.00%"));
    }

    // ── 0.15.0-B: resolve_pending_style_attrs ────────────────────────────────────

    fn style_attr_edit(edit: StyleAttrEdit) -> StyleAttrEdit {
        edit
    }

    #[test]
    fn resolve_pending_style_attrs_returns_none_with_no_pending_edits() {
        let vm = Vm::new();
        let indices = vm.cell_style_indices.clone();
        assert!(
            resolve_pending_style_attrs(&vm, XLSX_STYLES.as_bytes(), &indices)
                .unwrap()
                .is_none()
        );
    }

    #[test]
    fn resolve_pending_style_attrs_mints_a_new_font_preserving_other_properties() {
        let source = concat!(
            "<styleSheet>",
            "<fonts count=\"1\"><font><sz val=\"11\"/><color rgb=\"FF112233\"/><name val=\"Calibri\"/></font></fonts>",
            "<fills count=\"1\"><fill><patternFill patternType=\"none\"/></fill></fills>",
            "<borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders>",
            "<cellStyleXfs count=\"1\"><xf/></cellStyleXfs>",
            "<cellXfs count=\"1\"><xf/></cellXfs>",
            "</styleSheet>",
        );
        let mut vm = Vm::new();
        vm.set_style_on_sheet(
            "sheet1",
            1,
            1,
            1,
            1,
            &style_attr_edit(StyleAttrEdit {
                font: Some(reader::FontEdit {
                    bold: Some(true),
                    ..Default::default()
                }),
                ..Default::default()
            }),
        );
        let indices = vm.cell_style_indices.clone();
        let (new_xml, effective) = resolve_pending_style_attrs(&vm, source.as_bytes(), &indices)
            .unwrap()
            .expect("pending edit");
        assert!(new_xml.contains("<b val=\"1\"/>"));
        // Original font's other properties survive on the NEW cloned font record.
        assert!(new_xml.contains("<color rgb=\"FF112233\"/>"));
        assert!(new_xml.contains("<name val=\"Calibri\"/>"));
        assert_ne!(effective["sheet1"][&(1, 1)], 0);
    }

    #[test]
    fn resolve_pending_style_attrs_mints_a_solid_fill() {
        let mut vm = Vm::new();
        vm.set_style_on_sheet(
            "sheet1",
            1,
            1,
            1,
            1,
            &style_attr_edit(StyleAttrEdit {
                fill: Some(FillEdit {
                    color_argb: "FF4472C4".to_string(),
                }),
                ..Default::default()
            }),
        );
        let indices = vm.cell_style_indices.clone();
        let (new_xml, _) = resolve_pending_style_attrs(&vm, XLSX_STYLES.as_bytes(), &indices)
            .unwrap()
            .expect("pending edit");
        assert!(new_xml.contains("<fgColor rgb=\"FF4472C4\"/>"));
        assert!(new_xml.contains("<bgColor indexed=\"64\"/>"));
    }

    #[test]
    fn resolve_pending_style_attrs_reuses_a_font_dedup_not_a_duplicate() {
        let mut vm = Vm::new();
        vm.set_style_on_sheet(
            "sheet1",
            1,
            1,
            1,
            1,
            &style_attr_edit(StyleAttrEdit {
                font: Some(reader::FontEdit {
                    bold: Some(true),
                    ..Default::default()
                }),
                ..Default::default()
            }),
        );
        vm.set_style_on_sheet(
            "sheet1",
            5,
            5,
            5,
            5,
            &style_attr_edit(StyleAttrEdit {
                font: Some(reader::FontEdit {
                    bold: Some(true),
                    ..Default::default()
                }),
                ..Default::default()
            }),
        );
        let indices = vm.cell_style_indices.clone();
        let (_, effective) = resolve_pending_style_attrs(&vm, XLSX_STYLES.as_bytes(), &indices)
            .unwrap()
            .expect("pending edit");
        assert_eq!(effective["sheet1"][&(1, 1)], effective["sheet1"][&(5, 5)]);
    }

    #[test]
    fn resolve_pending_style_attrs_never_mutates_the_index_of_an_untouched_cell() {
        let mut vm = Vm::new();
        vm.set_style_on_sheet(
            "sheet1",
            1,
            1,
            1,
            1,
            &style_attr_edit(StyleAttrEdit {
                font: Some(reader::FontEdit {
                    bold: Some(true),
                    ..Default::default()
                }),
                ..Default::default()
            }),
        );
        let indices = vm.cell_style_indices.clone();
        let (_, effective) = resolve_pending_style_attrs(&vm, XLSX_STYLES.as_bytes(), &indices)
            .unwrap()
            .expect("pending edit");
        assert_ne!(effective["sheet1"][&(1, 1)], 0);
        assert!(!effective.get("sheet1").unwrap().contains_key(&(2, 2)));
    }

    #[test]
    fn resolve_pending_style_attrs_chains_onto_the_number_format_passs_own_output() {
        // The mandatory chaining fix: a cell edited by BOTH set_number_format and
        // set_style before one save must end up with both changes, not just
        // whichever pass ran last starting fresh from the original bytes.
        let mut vm = Vm::new();
        vm.set_number_format_on_sheet("sheet1", 1, 1, 1, 1, "#,##0.00");
        vm.set_style_on_sheet(
            "sheet1",
            1,
            1,
            1,
            1,
            &style_attr_edit(StyleAttrEdit {
                font: Some(reader::FontEdit {
                    bold: Some(true),
                    ..Default::default()
                }),
                ..Default::default()
            }),
        );
        let (numfmt_xml, numfmt_indices) =
            resolve_pending_number_formats(&vm, XLSX_STYLES.as_bytes()).expect("pending edit");
        let (final_xml, final_indices) =
            resolve_pending_style_attrs(&vm, numfmt_xml.as_bytes(), &numfmt_indices)
                .unwrap()
                .expect("pending edit");
        let idx = final_indices["sheet1"][&(1, 1)] as usize;
        let xfs = reader::extract_cell_xfs(&final_xml);
        let xf = &xfs[idx];
        assert!(xf.contains("numFmtId=\"4\""), "numFmtId lost: {xf}");
        let font_id = reader::span_attr_u32(xf, "fontId") as usize;
        let fonts = reader::extract_records(&final_xml, "fonts", "font");
        assert!(
            fonts[font_id].contains("<b val=\"1\"/>"),
            "font edit lost: {}",
            fonts[font_id]
        );
    }

    #[test]
    fn resolve_pending_style_attrs_alignment_merges_onto_a_real_fixture_shaped_xf() {
        let source = concat!(
            "<styleSheet>",
            "<fonts count=\"1\"><font/></fonts>",
            "<fills count=\"1\"><fill><patternFill patternType=\"none\"/></fill></fills>",
            "<borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders>",
            "<cellStyleXfs count=\"1\"><xf/></cellStyleXfs>",
            "<cellXfs count=\"1\"><xf><alignment vertical=\"center\"/></xf></cellXfs>",
            "</styleSheet>",
        );
        let mut vm = Vm::new();
        vm.set_style_on_sheet(
            "sheet1",
            1,
            1,
            1,
            1,
            &style_attr_edit(StyleAttrEdit {
                alignment: Some(reader::AlignmentEdit {
                    horizontal: Some("center".to_string()),
                    ..Default::default()
                }),
                ..Default::default()
            }),
        );
        let indices = vm.cell_style_indices.clone();
        let (new_xml, _) = resolve_pending_style_attrs(&vm, source.as_bytes(), &indices)
            .unwrap()
            .expect("pending edit");
        assert!(new_xml.contains("vertical=\"center\""));
        assert!(new_xml.contains("horizontal=\"center\""));
    }

    // ── 0.15.0-C1: named-style apply, copy_style ─────────────────────────────────

    /// `fixture4_hyperlink_comment_name.xlsm`'s real `xl/styles.xml` shape (font
    /// index/xfId/cellStyle name verified directly against that fixture).
    const FIXTURE4_SHAPED_STYLES: &str = concat!(
        "<styleSheet>",
        "<fonts count=\"3\">",
        "<font><sz val=\"12\"/><color theme=\"1\"/><name val=\"游ゴシック\"/></font>",
        "<font><sz val=\"6\"/><name val=\"游ゴシック\"/></font>",
        "<font><u/><sz val=\"12\"/><color theme=\"10\"/><name val=\"游ゴシック\"/></font>",
        "</fonts>",
        "<fills count=\"1\"><fill><patternFill patternType=\"none\"/></fill></fills>",
        "<borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders>",
        "<cellStyleXfs count=\"2\">",
        "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"><alignment vertical=\"center\"/></xf>",
        "<xf numFmtId=\"0\" fontId=\"2\" fillId=\"0\" borderId=\"0\" applyNumberFormat=\"0\" \
           applyFill=\"0\" applyBorder=\"0\" applyAlignment=\"0\" applyProtection=\"0\">\
           <alignment vertical=\"center\"/></xf>",
        "</cellStyleXfs>",
        "<cellXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\">\
           <alignment vertical=\"center\"/></xf></cellXfs>",
        "<cellStyles count=\"2\">",
        "<cellStyle name=\"ハイパーリンク\" xfId=\"1\" builtinId=\"8\"/>",
        "<cellStyle name=\"標準\" xfId=\"0\" builtinId=\"0\"/>",
        "</cellStyles>",
        "</styleSheet>",
    );

    #[test]
    fn resolve_pending_style_attrs_applies_a_real_fixture_shaped_named_style() {
        let mut vm = Vm::new();
        vm.set_style_on_sheet(
            "sheet1",
            1,
            1,
            1,
            1,
            &style_attr_edit(StyleAttrEdit {
                named_style: Some("ハイパーリンク".to_string()),
                ..Default::default()
            }),
        );
        let indices = vm.cell_style_indices.clone();
        let (new_xml, effective) =
            resolve_pending_style_attrs(&vm, FIXTURE4_SHAPED_STYLES.as_bytes(), &indices)
                .unwrap()
                .expect("pending edit");
        let idx = effective["sheet1"][&(1, 1)] as usize;
        let xfs = reader::extract_cell_xfs(&new_xml);
        let xf = &xfs[idx];
        // Matches real Excel's own behavior: the named style's fontId is baked directly
        // onto the cell's cellXfs entry, not left to xfId-based inheritance alone.
        assert!(xf.contains("xfId=\"1\""), "xfId not set: {xf}");
        assert_eq!(
            reader::span_attr_u32(xf, "fontId"),
            2,
            "fontId not baked in: {xf}"
        );
    }

    #[test]
    fn resolve_pending_style_attrs_errors_on_an_unknown_named_style() {
        let mut vm = Vm::new();
        vm.set_style_on_sheet(
            "sheet1",
            1,
            1,
            1,
            1,
            &style_attr_edit(StyleAttrEdit {
                named_style: Some("Bad".to_string()),
                ..Default::default()
            }),
        );
        let indices = vm.cell_style_indices.clone();
        let err = resolve_pending_style_attrs(&vm, FIXTURE4_SHAPED_STYLES.as_bytes(), &indices)
            .unwrap_err();
        assert!(
            err.contains("Bad"),
            "error should name the missing style: {err}"
        );
    }

    #[test]
    fn resolve_pending_style_attrs_applies_named_style_before_other_fields_on_the_same_edit() {
        // font=bold on the SAME edit as named_style must apply ON TOP of the named
        // style's own baked-in font, not be discarded by it.
        let mut vm = Vm::new();
        vm.set_style_on_sheet(
            "sheet1",
            1,
            1,
            1,
            1,
            &style_attr_edit(StyleAttrEdit {
                named_style: Some("ハイパーリンク".to_string()),
                font: Some(reader::FontEdit {
                    bold: Some(true),
                    ..Default::default()
                }),
                ..Default::default()
            }),
        );
        let indices = vm.cell_style_indices.clone();
        let (new_xml, effective) =
            resolve_pending_style_attrs(&vm, FIXTURE4_SHAPED_STYLES.as_bytes(), &indices)
                .unwrap()
                .expect("pending edit");
        let idx = effective["sheet1"][&(1, 1)] as usize;
        let xfs = reader::extract_cell_xfs(&new_xml);
        let xf = &xfs[idx];
        assert!(xf.contains("xfId=\"1\""), "named style lost: {xf}");
        let font_id = reader::span_attr_u32(xf, "fontId") as usize;
        let fonts = reader::extract_records(&new_xml, "fonts", "font");
        // Cloned from font index 2 (the hyperlink font: underlined, theme-colored) with
        // bold ADDED, not replaced by a from-scratch font.
        assert!(fonts[font_id].contains("<b/>") || fonts[font_id].contains("<b val=\"1\"/>"));
        assert!(
            fonts[font_id].contains("<u/>"),
            "underline lost: {}",
            fonts[font_id]
        );
        assert!(
            fonts[font_id].contains("theme=\"10\""),
            "theme color lost: {}",
            fonts[font_id]
        );
    }

    #[test]
    fn resolve_pending_style_copies_returns_none_with_no_pending_edits() {
        let vm = Vm::new();
        let indices = vm.cell_style_indices.clone();
        assert!(resolve_pending_style_copies(&vm, &indices).is_none());
    }

    #[test]
    fn resolve_pending_style_copies_points_dest_at_the_resolved_source_index() {
        let mut vm = Vm::new();
        vm.cell_style_indices.insert(
            "sheet1".to_string(),
            std::collections::HashMap::from([((1, 1), 7u32)]),
        );
        vm.copy_style_on_sheet("sheet1", (1, 1), 2, 2, 2, 2);
        let indices = vm.cell_style_indices.clone();
        let effective = resolve_pending_style_copies(&vm, &indices).expect("pending copy");
        assert_eq!(effective["sheet1"][&(2, 2)], 7);
        // Source cell itself is untouched.
        assert_eq!(effective["sheet1"][&(1, 1)], 7);
    }

    #[test]
    fn resolve_pending_style_copies_picks_up_a_pending_edit_on_the_source_not_just_its_resolved_index()
     {
        // The critical Finding-1-shaped case: set_style on the source, then copy_style
        // from it, before any save -- the copy must see the NEW style, not the stale
        // pre-edit one. Simulated here by chaining resolve_pending_style_attrs's own
        // output into resolve_pending_style_copies, exactly like save_xlsx_impl does.
        let mut vm = Vm::new();
        vm.set_style_on_sheet(
            "sheet1",
            1,
            1,
            1,
            1,
            &style_attr_edit(StyleAttrEdit {
                font: Some(reader::FontEdit {
                    bold: Some(true),
                    ..Default::default()
                }),
                ..Default::default()
            }),
        );
        vm.copy_style_on_sheet("sheet1", (1, 1), 2, 2, 2, 2);

        let starting_indices = vm.cell_style_indices.clone();
        let (_, style_attrs_effective) =
            resolve_pending_style_attrs(&vm, XLSX_STYLES.as_bytes(), &starting_indices)
                .unwrap()
                .expect("pending style edit");
        let final_effective =
            resolve_pending_style_copies(&vm, &style_attrs_effective).expect("pending copy");
        assert_eq!(
            final_effective["sheet1"][&(2, 2)],
            style_attrs_effective["sheet1"][&(1, 1)],
            "copy_style should point at the source's NEWLY resolved (bold) index, not a stale one"
        );
        assert_ne!(final_effective["sheet1"][&(2, 2)], 0);
    }

    #[test]
    fn resolve_pending_style_copies_never_mutates_an_untouched_cell() {
        let mut vm = Vm::new();
        vm.cell_style_indices.insert(
            "sheet1".to_string(),
            std::collections::HashMap::from([((1, 1), 7u32), ((3, 3), 9u32)]),
        );
        vm.copy_style_on_sheet("sheet1", (1, 1), 2, 2, 2, 2);
        let indices = vm.cell_style_indices.clone();
        let effective = resolve_pending_style_copies(&vm, &indices).expect("pending copy");
        assert_eq!(effective["sheet1"][&(3, 3)], 9);
    }

    // ── 0.15.0-C2: resolve_pending_row_column_styles ─────────────────────────────

    #[test]
    fn resolve_pending_row_column_styles_returns_none_with_no_pending_edits() {
        let vm = Vm::new();
        assert!(
            resolve_pending_row_column_styles(
                &vm,
                XLSX_STYLES.as_bytes(),
                &vm.row_styles,
                &vm.column_styles
            )
            .unwrap()
            .is_none()
        );
    }

    #[test]
    fn resolve_pending_row_column_styles_resolves_a_row_style() {
        let mut vm = Vm::new();
        vm.set_row_style_on_sheet(
            "sheet1",
            3,
            &style_attr_edit(StyleAttrEdit {
                font: Some(reader::FontEdit {
                    bold: Some(true),
                    ..Default::default()
                }),
                ..Default::default()
            }),
        );
        let (new_xml, rows, cols) = resolve_pending_row_column_styles(
            &vm,
            XLSX_STYLES.as_bytes(),
            &vm.row_styles,
            &vm.column_styles,
        )
        .unwrap()
        .expect("pending row edit");
        assert!(new_xml.contains("<b val=\"1\"/>"));
        assert_ne!(rows["sheet1"][&3], 0);
        assert!(cols.is_empty() || cols["sheet1"].is_empty());
    }

    #[test]
    fn resolve_pending_row_column_styles_resolves_a_column_style() {
        let mut vm = Vm::new();
        vm.set_column_style_on_sheet(
            "sheet1",
            2,
            &style_attr_edit(StyleAttrEdit {
                fill: Some(FillEdit {
                    color_argb: "FF4472C4".to_string(),
                }),
                ..Default::default()
            }),
        );
        let (new_xml, _, cols) = resolve_pending_row_column_styles(
            &vm,
            XLSX_STYLES.as_bytes(),
            &vm.row_styles,
            &vm.column_styles,
        )
        .unwrap()
        .expect("pending column edit");
        assert!(new_xml.contains("<fgColor rgb=\"FF4472C4\"/>"));
        assert_eq!(cols["sheet1"], vec![(2, 2, 1)]);
    }

    #[test]
    fn resolve_pending_row_column_styles_splits_an_existing_column_range_when_setting_one_column() {
        let mut vm = Vm::new();
        vm.column_styles
            .insert("sheet1".to_string(), vec![(1, 10, 3u32)]);
        vm.set_column_style_on_sheet(
            "sheet1",
            5,
            &style_attr_edit(StyleAttrEdit {
                font: Some(reader::FontEdit {
                    bold: Some(true),
                    ..Default::default()
                }),
                ..Default::default()
            }),
        );
        let (_, _, cols) = resolve_pending_row_column_styles(
            &vm,
            XLSX_STYLES.as_bytes(),
            &vm.row_styles,
            &vm.column_styles,
        )
        .unwrap()
        .expect("pending column edit");
        let mut ranges = cols["sheet1"].clone();
        ranges.sort();
        // The (1,10,3) range splits into (1,4,3) and (6,10,3), plus a fresh (5,5,new)
        // singleton for the touched column -- columns 1-4 and 6-10 keep the ORIGINAL
        // style index 3 unchanged, only column 5 gets the new one.
        assert_eq!(ranges.len(), 3);
        assert!(ranges.contains(&(1, 4, 3)));
        assert!(ranges.contains(&(6, 10, 3)));
        let touched = ranges
            .iter()
            .find(|&&(min, max, _)| min == 5 && max == 5)
            .expect("a (5,5,_) singleton for the touched column");
        assert_ne!(touched.2, 3);
    }

    #[test]
    fn resolve_pending_row_column_styles_chains_onto_the_style_attrs_passs_own_output() {
        // The mandatory chaining fix (0.15.0-B's own Finding 1, extended here): a cell
        // touched by set_style AND a row touched by set_row_style before one save must
        // resolve into the SAME cellXfs/fonts/fills/borders tables, not two independent
        // ones that could mint near-duplicate records or collide on index.
        let mut vm = Vm::new();
        vm.set_style_on_sheet(
            "sheet1",
            1,
            1,
            1,
            1,
            &style_attr_edit(StyleAttrEdit {
                font: Some(reader::FontEdit {
                    bold: Some(true),
                    ..Default::default()
                }),
                ..Default::default()
            }),
        );
        vm.set_row_style_on_sheet(
            "sheet1",
            9,
            &style_attr_edit(StyleAttrEdit {
                font: Some(reader::FontEdit {
                    bold: Some(true),
                    ..Default::default()
                }),
                ..Default::default()
            }),
        );
        let starting_indices = vm.cell_style_indices.clone();
        let (style_attrs_xml, style_attrs_effective) =
            resolve_pending_style_attrs(&vm, XLSX_STYLES.as_bytes(), &starting_indices)
                .unwrap()
                .expect("pending style edit");
        let (_, rows, _) = resolve_pending_row_column_styles(
            &vm,
            style_attrs_xml.as_bytes(),
            &vm.row_styles,
            &vm.column_styles,
        )
        .unwrap()
        .expect("pending row edit");
        // Both edits request the identical bold-only change, so they must dedup onto
        // the SAME cellXf index the cell-level pass already minted -- not a second,
        // near-duplicate record.
        assert_eq!(rows["sheet1"][&9], style_attrs_effective["sheet1"][&(1, 1)]);
    }

    #[test]
    fn copy_style_on_a_from_scratch_vm_survives_a_save_and_reload() {
        let mut vm = Vm::new();
        vm.cells_mut().insert(
            (1, 1),
            CellContent {
                formula: None,
                value: Variant::Str("styled".to_string()),
            },
        );
        vm.cells_mut().insert(
            (2, 2),
            CellContent {
                formula: None,
                value: Variant::Str("plain".to_string()),
            },
        );
        vm.set_style_on_sheet(
            "sheet1",
            1,
            1,
            1,
            1,
            &style_attr_edit(StyleAttrEdit {
                font: Some(reader::FontEdit {
                    bold: Some(true),
                    ..Default::default()
                }),
                ..Default::default()
            }),
        );
        vm.copy_style_on_sheet("sheet1", (1, 1), 2, 2, 2, 2);

        let path = "/tmp/elixcee_test_copy_style_from_scratch.xlsx";
        save_workbook_impl(&vm, path).expect("save should succeed");

        let bytes = std::fs::read(path).expect("read output");
        let mut zip = zip::ZipArchive::new(std::io::Cursor::new(bytes)).expect("open zip");
        let mut styles_xml = String::new();
        std::io::Read::read_to_string(
            &mut zip.by_name("xl/styles.xml").expect("styles.xml"),
            &mut styles_xml,
        )
        .expect("read styles.xml");
        assert!(styles_xml.contains("<b val=\"1\"/>"));

        let mut sheet_xml = String::new();
        std::io::Read::read_to_string(
            &mut zip.by_name("xl/worksheets/sheet1.xml").expect("sheet1.xml"),
            &mut sheet_xml,
        )
        .expect("read sheet1.xml");
        // Both cells must share the SAME non-default style index.
        let a1_style = sheet_xml
            .split("r=\"A1\"")
            .nth(1)
            .and_then(|s| s.split('>').next())
            .and_then(|s| s.split("s=\"").nth(1))
            .and_then(|s| s.split('"').next())
            .expect("A1 has an s= attribute");
        let b2_style = sheet_xml
            .split("r=\"B2\"")
            .nth(1)
            .and_then(|s| s.split('>').next())
            .and_then(|s| s.split("s=\"").nth(1))
            .and_then(|s| s.split('"').next())
            .expect("B2 has an s= attribute");
        assert_eq!(a1_style, b2_style);
        assert_ne!(a1_style, "0");

        let _ = std::fs::remove_file(path);
    }

    #[test]
    fn set_style_on_a_from_scratch_vm_survives_a_save_and_reload() {
        let mut vm = Vm::new();
        vm.cells_mut().insert(
            (1, 1),
            CellContent {
                formula: None,
                value: Variant::Str("hi".to_string()),
            },
        );
        vm.set_style_on_sheet(
            "sheet1",
            1,
            1,
            1,
            1,
            &style_attr_edit(StyleAttrEdit {
                font: Some(reader::FontEdit {
                    bold: Some(true),
                    ..Default::default()
                }),
                ..Default::default()
            }),
        );

        let path = "/tmp/elixcee_test_set_style_from_scratch.xlsx";
        save_workbook_impl(&vm, path).expect("save should succeed");

        let bytes = std::fs::read(path).expect("read output");
        let mut zip = zip::ZipArchive::new(std::io::Cursor::new(bytes)).expect("open zip");
        let mut styles_xml = String::new();
        std::io::Read::read_to_string(
            &mut zip.by_name("xl/styles.xml").expect("styles.xml"),
            &mut styles_xml,
        )
        .expect("read styles.xml");
        assert!(styles_xml.contains("<b val=\"1\"/>"));

        let _ = std::fs::remove_file(path);
    }

    #[test]
    fn set_row_style_and_set_column_style_survive_a_save_and_reload() {
        let mut vm = Vm::new();
        vm.set_row_style_on_sheet(
            "sheet1",
            3,
            &style_attr_edit(StyleAttrEdit {
                font: Some(reader::FontEdit {
                    bold: Some(true),
                    ..Default::default()
                }),
                ..Default::default()
            }),
        );
        vm.set_column_style_on_sheet(
            "sheet1",
            2,
            &style_attr_edit(StyleAttrEdit {
                fill: Some(FillEdit {
                    color_argb: "FF4472C4".to_string(),
                }),
                ..Default::default()
            }),
        );

        let path = "/tmp/elixcee_test_row_column_style_from_scratch.xlsx";
        save_workbook_impl(&vm, path).expect("save should succeed");

        let bytes = std::fs::read(path).expect("read output");
        let mut zip = zip::ZipArchive::new(std::io::Cursor::new(bytes)).expect("open zip");
        let mut styles_xml = String::new();
        std::io::Read::read_to_string(
            &mut zip.by_name("xl/styles.xml").expect("styles.xml"),
            &mut styles_xml,
        )
        .expect("read styles.xml");
        assert!(styles_xml.contains("<b val=\"1\"/>"));
        assert!(styles_xml.contains("<fgColor rgb=\"FF4472C4\"/>"));

        let mut sheet_xml = String::new();
        std::io::Read::read_to_string(
            &mut zip.by_name("xl/worksheets/sheet1.xml").expect("sheet1.xml"),
            &mut sheet_xml,
        )
        .expect("read sheet1.xml");
        assert!(sheet_xml.contains("customFormat=\"1\""));
        assert!(sheet_xml.contains("<row r=\"3\""));
        assert!(sheet_xml.contains("<col min=\"2\" max=\"2\""));

        let _ = std::fs::remove_file(path);
    }

    #[test]
    fn a_cells_own_set_style_wins_over_its_rows_default_style_on_reload() {
        // Real Excel precedence: a cell's own explicit s= always overrides its row's
        // customFormat default -- this project's job is only to persist both facts
        // independently (row_styles vs. cell_style_indices are separate fields), never
        // to resolve precedence itself.
        let mut vm = Vm::new();
        vm.cells_mut().insert(
            (3, 1),
            CellContent {
                formula: None,
                value: Variant::Str("own style".to_string()),
            },
        );
        vm.set_style_on_sheet(
            "sheet1",
            3,
            1,
            3,
            1,
            &style_attr_edit(StyleAttrEdit {
                font: Some(reader::FontEdit {
                    italic: Some(true),
                    ..Default::default()
                }),
                ..Default::default()
            }),
        );
        vm.set_row_style_on_sheet(
            "sheet1",
            3,
            &style_attr_edit(StyleAttrEdit {
                font: Some(reader::FontEdit {
                    bold: Some(true),
                    ..Default::default()
                }),
                ..Default::default()
            }),
        );

        let path = "/tmp/elixcee_test_row_style_vs_cell_style_precedence.xlsx";
        save_workbook_impl(&vm, path).expect("save should succeed");

        let bytes = std::fs::read(path).expect("read output");
        let mut zip = zip::ZipArchive::new(std::io::Cursor::new(bytes)).expect("open zip");
        let mut sheet_xml = String::new();
        std::io::Read::read_to_string(
            &mut zip.by_name("xl/worksheets/sheet1.xml").expect("sheet1.xml"),
            &mut sheet_xml,
        )
        .expect("read sheet1.xml");
        // The <row> element carries the row's own default style index...
        let row_style = sheet_xml
            .split("<row r=\"3\"")
            .nth(1)
            .and_then(|s| s.split('>').next())
            .and_then(|s| s.split("s=\"").nth(1))
            .and_then(|s| s.split('"').next())
            .expect("row 3 has an s= attribute")
            .to_string();
        // ...and the cell keeps its OWN, different style index, not the row's.
        let cell_style = sheet_xml
            .split("r=\"A3\"")
            .nth(1)
            .and_then(|s| s.split('>').next())
            .and_then(|s| s.split("s=\"").nth(1))
            .and_then(|s| s.split('"').next())
            .expect("A3 has an s= attribute");
        assert_ne!(row_style, cell_style);

        let _ = std::fs::remove_file(path);
    }

    #[test]
    fn insert_rows_on_sheet_shifts_row_styles() {
        let mut vm = Vm::new();
        vm.row_styles.insert(
            "sheet1".to_string(),
            std::collections::HashMap::from([(5u32, 3u32)]),
        );
        vm.insert_rows_on_sheet("sheet1", 1, 2);
        assert_eq!(vm.row_styles.get("sheet1").unwrap().get(&7), Some(&3));
        assert!(!vm.row_styles.get("sheet1").unwrap().contains_key(&5));
    }

    #[test]
    fn delete_rows_on_sheet_drops_a_row_style_inside_the_deleted_band() {
        let mut vm = Vm::new();
        vm.row_styles.insert(
            "sheet1".to_string(),
            std::collections::HashMap::from([(5u32, 3u32)]),
        );
        vm.delete_rows_on_sheet("sheet1", 4, 3);
        assert!(vm.row_styles.get("sheet1").unwrap().is_empty());
    }

    #[test]
    fn insert_cols_on_sheet_shifts_column_styles() {
        let mut vm = Vm::new();
        vm.column_styles
            .insert("sheet1".to_string(), vec![(3, 5, 3u32)]);
        vm.insert_cols_on_sheet("sheet1", 1, 2);
        assert_eq!(vm.column_styles.get("sheet1").unwrap(), &vec![(5, 7, 3)]);
    }

    #[test]
    fn rename_sheet_re_keys_row_styles_and_column_styles() {
        let mut vm = Vm::new();
        vm.row_styles.insert(
            "sheet1".to_string(),
            std::collections::HashMap::from([(3u32, 5u32)]),
        );
        vm.column_styles
            .insert("sheet1".to_string(), vec![(1, 2, 7)]);
        vm.pending_row_styles.insert(
            "sheet1".to_string(),
            std::collections::HashMap::from([(9u32, StyleAttrEdit::default())]),
        );
        vm.rename_sheet("Sheet1", "Renamed").unwrap();
        assert!(!vm.row_styles.contains_key("sheet1"));
        assert_eq!(vm.row_styles.get("renamed").unwrap().get(&3), Some(&5));
        assert_eq!(vm.column_styles.get("renamed").unwrap(), &vec![(1, 2, 7)]);
        assert!(
            vm.pending_row_styles
                .get("renamed")
                .unwrap()
                .contains_key(&9)
        );
    }

    #[test]
    fn remove_sheet_clears_row_styles_and_column_styles() {
        // delete_sheet no-ops on the ACTIVE sheet (matches every other remove_sheet
        // test in this codebase) -- "Sheet1" stays active by construction, so create
        // and delete a second, non-active sheet instead.
        let mut vm = Vm::new();
        vm.ensure_sheet("Sheet2");
        vm.row_styles.insert(
            "sheet2".to_string(),
            std::collections::HashMap::from([(3u32, 5u32)]),
        );
        vm.column_styles
            .insert("sheet2".to_string(), vec![(1, 2, 7)]);
        vm.delete_sheet("Sheet2").unwrap();
        assert!(!vm.row_styles.contains_key("sheet2"));
        assert!(!vm.column_styles.contains_key("sheet2"));
    }

    #[test]
    fn test_save_ods_roundtrip() {
        use calamine::{Reader, open_workbook_auto};

        let mut vm = Vm::new();
        vm.cells_mut().insert(
            (1, 1),
            CellContent {
                formula: None,
                value: Variant::Integer(42),
            },
        );
        vm.cells_mut().insert(
            (1, 2),
            CellContent {
                formula: None,
                value: Variant::Str("hello".into()),
            },
        );
        vm.cells_mut().insert(
            (2, 1),
            CellContent {
                formula: None,
                value: Variant::Boolean(true),
            },
        );

        let path = "/tmp/elixcee_test_ods.ods";
        save_workbook_impl(&vm, path).expect("ODS save should succeed");

        // Reload with calamine
        let mut wb = open_workbook_auto(path).expect("ODS open should succeed");
        let range = wb.worksheet_range("sheet1").expect("sheet1 should exist");
        let cells: Vec<_> = range.cells().collect();
        assert!(!cells.is_empty(), "ODS file should have cells");
    }

    #[test]
    fn test_save_workbook_multi_sheet() {
        let mut vm = Vm::new();
        vm.cells_mut().insert(
            (1, 1),
            CellContent {
                formula: None,
                value: Variant::Integer(1),
            },
        );
        vm.ensure_sheet("sheet2");
        let prev = vm.active_sheet.clone();
        vm.active_sheet = "sheet2".into();
        vm.cells_mut().insert(
            (1, 1),
            CellContent {
                formula: None,
                value: Variant::Integer(2),
            },
        );
        vm.active_sheet = prev;

        let path = "/tmp/elixcee_test_multisheet.xlsx";
        save_workbook_impl(&vm, path).expect("save should succeed");

        let mut wb: Xlsx<_> = open_workbook(path).expect("open should succeed");
        assert!(wb.worksheet_range("sheet1").is_ok(), "sheet1 should exist");
        assert!(wb.worksheet_range("sheet2").is_ok(), "sheet2 should exist");
    }

    // 0.10.0-A: build_xlsx_workbook must preserve a surviving sheet's original sheetId
    // (see WorksheetOrigin's own doc comment for why -- this is the one identifier
    // snapshot.rs's stable_id already treats as cross-save-stable) instead of always
    // renumbering positionally, while a sheet with no known origin still gets a fresh,
    // non-colliding id.
    #[test]
    fn build_xlsx_workbook_preserves_original_sheet_ids_and_assigns_fresh_ones_for_new_sheets() {
        let mut origins = std::collections::HashMap::new();
        origins.insert(
            "sheet1".to_string(),
            WorksheetOrigin {
                original_sheet_id: Some("7".to_string()),
                original_workbook_rel_id: Some("rId3".to_string()),
                original_part_name: Some("xl/worksheets/sheet2.xml".to_string()),
                original_display_name: Some("Sheet1".to_string()),
            },
        );
        origins.insert(
            "sheet2".to_string(),
            WorksheetOrigin {
                original_sheet_id: Some("2".to_string()),
                original_workbook_rel_id: None,
                original_part_name: None,
                original_display_name: None,
            },
        );
        // "newsheet" is deliberately absent from `origins` -- a sheet created purely
        // in-VBA, or one populate_from_sheets never saw.
        let plans = plan_worksheet_output(
            &[
                "sheet1".to_string(),
                "sheet2".to_string(),
                "newsheet".to_string(),
            ],
            &origins,
            &[],
        );
        let xml = build_xlsx_workbook(&plans, &OpaqueWorkbookFragments::default());

        assert!(
            xml.contains("<sheet name=\"Sheet1\" sheetId=\"7\" r:id=\"rId1\"/>"),
            "expected sheet1 to keep its original sheetId 7 and original-case display name: {xml}"
        );
        assert!(
            xml.contains("<sheet name=\"sheet2\" sheetId=\"2\" r:id=\"rId2\"/>"),
            "expected sheet2 (no display name recorded) to fall back to its lookup key: {xml}"
        );
        // Fresh id must not collide with any preserved original id (max is 7) --
        // asserting it's strictly greater than 7 rather than a specific number, since
        // the exact fresh value is an implementation detail this test shouldn't pin down.
        let newsheet_id: u32 = xml
            .lines()
            .find(|l| l.contains("name=\"newsheet\""))
            .and_then(|l| l.split("sheetId=\"").nth(1))
            .and_then(|rest| rest.split('"').next())
            .and_then(|id| id.parse().ok())
            .expect("newsheet should have a numeric sheetId");
        assert!(
            newsheet_id > 7,
            "fresh sheetId {newsheet_id} must not collide with the highest preserved original id (7)"
        );
        assert!(
            xml.contains("r:id=\"rId3\""),
            "newsheet's r:id must still be positional: {xml}"
        );

        // 0.10.0-D, D1: sheet1's output part name is its ORIGIN's part name
        // (sheet2.xml), not a position-derived sheet1.xml.
        assert_eq!(plans[0].output_part_name, "xl/worksheets/sheet2.xml");
        assert!(plans[0].is_existing);
        // sheet2 has a sheetId but no original_part_name (e.g. an .ods-sourced origin) --
        // treated as needing a fresh part name, same as a brand-new sheet.
        assert!(!plans[1].is_existing);
        // newsheet's fresh part name must not collide with sheet1's real original
        // (sheet2.xml) -- it starts counting from that number, not from 0.
        assert_ne!(plans[2].output_part_name, "xl/worksheets/sheet2.xml");
        assert_ne!(plans[1].output_part_name, plans[2].output_part_name);
    }

    #[test]
    fn build_xlsx_workbook_assigns_sequential_ids_when_no_sheet_has_a_known_origin() {
        let origins = std::collections::HashMap::new();
        let plans = plan_worksheet_output(&["a".to_string(), "b".to_string()], &origins, &[]);
        let xml = build_xlsx_workbook(&plans, &OpaqueWorkbookFragments::default());
        assert!(
            xml.contains("<sheet name=\"a\" sheetId=\"1\" r:id=\"rId1\"/>"),
            "{xml}"
        );
        assert!(
            xml.contains("<sheet name=\"b\" sheetId=\"2\" r:id=\"rId2\"/>"),
            "{xml}"
        );
    }

    #[test]
    fn merge_rect_to_a1_formats_a_multi_cell_range() {
        assert_eq!(merge_rect_to_a1(&((1, 2), (1, 4))), "B1:D1");
    }

    #[test]
    fn merge_rect_to_a1_formats_a_single_cell_range() {
        assert_eq!(merge_rect_to_a1(&((3, 3), (3, 3))), "C3:C3");
    }
}

// ── Differential read tests: calamine (oracle) vs hand-written reader ─────────
#[cfg(test)]
mod diff_reader_tests {
    use super::*;
    use crate::reader::{SheetCell, read_workbook as rd};
    use calamine::{Data, Reader, Xlsx, open_workbook, open_workbook_auto};

    fn calamine_cell_to_variant(d: &Data) -> Option<Variant> {
        match d {
            Data::String(s) => Some(Variant::Str(s.clone())),
            Data::Float(f) => {
                if f.fract() == 0.0 && *f >= i64::MIN as f64 && *f <= i64::MAX as f64 {
                    Some(Variant::Integer(*f as i64))
                } else {
                    Some(Variant::Float(*f))
                }
            }
            Data::Bool(b) => Some(Variant::Boolean(*b)),
            _ => None,
        }
    }

    fn rd_cell_to_variant(c: &SheetCell) -> Variant {
        match c {
            SheetCell::Integer(n) => Variant::Integer(*n),
            SheetCell::Float(f) => Variant::Float(*f),
            SheetCell::Str(s) => Variant::Str(s.clone()),
            SheetCell::Bool(b) => Variant::Boolean(*b),
            SheetCell::Error(e) => Variant::Error(e.clone()),
        }
    }

    fn calamine_xlsx_cells(
        path: &str,
        sheet: &str,
    ) -> std::collections::HashMap<(u32, u32), Variant> {
        let mut wb: Xlsx<_> = open_workbook(path).unwrap();
        let range = wb.worksheet_range(sheet).unwrap();
        let (sr, sc) = range.start().unwrap_or((0, 0));
        range
            .cells()
            .filter_map(|(r, c, d)| {
                calamine_cell_to_variant(d).map(|v| ((r as u32 + sr + 1, c as u32 + sc + 1), v))
            })
            .collect()
    }

    fn rd_xlsx_cells(path: &str, sheet: &str) -> std::collections::HashMap<(u32, u32), Variant> {
        rd(path)
            .unwrap()
            .into_iter()
            .find(|s| s.name == sheet)
            .unwrap()
            .cells
            .iter()
            .map(|(&k, v)| (k, rd_cell_to_variant(v)))
            .collect()
    }

    fn calamine_ods_cells(
        path: &str,
        sheet: &str,
    ) -> std::collections::HashMap<(u32, u32), Variant> {
        let mut wb = open_workbook_auto(path).unwrap();
        let range = wb.worksheet_range(sheet).unwrap();
        let (sr, sc) = range.start().unwrap_or((0, 0));
        range
            .cells()
            .filter_map(|(r, c, d)| {
                calamine_cell_to_variant(d).map(|v| ((r as u32 + sr + 1, c as u32 + sc + 1), v))
            })
            .collect()
    }

    fn rd_ods_cells(path: &str, sheet: &str) -> std::collections::HashMap<(u32, u32), Variant> {
        rd(path)
            .unwrap()
            .into_iter()
            .find(|s| s.name == sheet)
            .unwrap()
            .cells
            .iter()
            .map(|(&k, v)| (k, rd_cell_to_variant(v)))
            .collect()
    }

    #[test]
    // 3.14 is an arbitrary decimal test value covering a float cell type, not π.
    #[allow(clippy::approx_constant)]
    fn diff_xlsx_all_types() {
        let mut vm = Vm::new();
        vm.cells_mut().insert(
            (1, 1),
            CellContent {
                formula: None,
                value: Variant::Integer(42),
            },
        );
        vm.cells_mut().insert(
            (2, 1),
            CellContent {
                formula: None,
                value: Variant::Str("hello".into()),
            },
        );
        vm.cells_mut().insert(
            (3, 1),
            CellContent {
                formula: None,
                value: Variant::Float(3.14),
            },
        );
        vm.cells_mut().insert(
            (4, 1),
            CellContent {
                formula: None,
                value: Variant::Boolean(true),
            },
        );
        vm.cells_mut().insert(
            (5, 1),
            CellContent {
                formula: None,
                value: Variant::Str(" leading and trailing ".into()),
            },
        );

        let path = "/tmp/elixcee_diff_xlsx.xlsx";
        save_workbook_impl(&vm, path).unwrap();

        let cal = calamine_xlsx_cells(path, "sheet1");
        let mine = rd_xlsx_cells(path, "sheet1");
        assert_eq!(cal, mine, "XLSX diff failed");
    }

    #[test]
    fn diff_xlsx_multi_sheet() {
        let mut vm = Vm::new();
        vm.cells_mut().insert(
            (1, 1),
            CellContent {
                formula: None,
                value: Variant::Integer(1),
            },
        );
        vm.ensure_sheet("sheet2");
        let prev = vm.active_sheet.clone();
        vm.active_sheet = "sheet2".into();
        vm.cells_mut().insert(
            (2, 3),
            CellContent {
                formula: None,
                value: Variant::Str("s2".into()),
            },
        );
        vm.active_sheet = prev;

        let path = "/tmp/elixcee_diff_multi.xlsx";
        save_workbook_impl(&vm, path).unwrap();

        for sheet in &["sheet1", "sheet2"] {
            let cal = calamine_xlsx_cells(path, sheet);
            let mine = rd_xlsx_cells(path, sheet);
            assert_eq!(cal, mine, "XLSX multi-sheet diff failed for {}", sheet);
        }
    }

    // Every test above round-trips a file elixcee itself wrote — it proves the
    // writer and reader agree with each other via calamine, not that either
    // agrees with a real, independent producer. `source.xlsx`/`source.ods`
    // are generated by real LibreOffice (see tests/fixtures/e2e/README.md),
    // exercising real-world encoding elixcee's own writer never produces:
    // ODS `number-columns-repeated`, non-contiguous XLSX `<row r="...">`
    // numbering from a dropped blank row, and a shared string split across
    // multiple `<r><t>` rich-text runs.
    fn e2e_fixture(name: &str) -> String {
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("tests/fixtures/e2e")
            .join(name)
            .to_str()
            .unwrap()
            .to_string()
    }

    #[test]
    fn diff_real_producer_xlsx() {
        let path = e2e_fixture("source.xlsx");
        let cal = calamine_xlsx_cells(&path, "source");
        let mine = rd_xlsx_cells(&path, "source");
        assert_eq!(cal, mine, "real-LibreOffice-produced XLSX diff failed");
        assert_eq!(
            mine.get(&(2, 4)),
            Some(&Variant::Str("quote \" amp & lt < gt >".into())),
            "named-entity decoding on a real producer's sharedStrings.xml"
        );
        assert_eq!(
            mine.get(&(3, 4)),
            Some(&Variant::Str("unicode: café ★ 日本語".into())),
            "multi-run <si> (split across two <r><t> runs by a real producer) must concatenate"
        );
        assert!(
            !mine.contains_key(&(4, 1)),
            "row 4 is blank and dropped entirely from <sheetData> by the real producer"
        );
        assert_eq!(
            mine.get(&(5, 4)),
            Some(&Variant::Str("after-column-gap".into())),
            "real content after 3 leading blank columns"
        );
        assert_eq!(
            mine.get(&(9, 1)),
            Some(&Variant::Str("Carol".into())),
            "row numbering after 4 dropped blank rows (4,6,7,8) stays non-contiguous, not shifted"
        );
    }

    #[test]
    fn diff_real_producer_ods() {
        let path = e2e_fixture("source.ods");
        let cal = calamine_ods_cells(&path, "source");
        let mine = rd_ods_cells(&path, "source");
        assert_eq!(cal, mine, "real-LibreOffice-produced ODS diff failed");
        assert_eq!(
            mine.get(&(2, 4)),
            Some(&Variant::Str("quote \" amp & lt < gt >".into()))
        );
        assert_eq!(
            mine.get(&(5, 4)),
            Some(&Variant::Str("after-column-gap".into())),
            "table:number-columns-repeated=\"3\" followed by real content in the same row must not shift its column"
        );
        assert_eq!(
            mine.get(&(9, 1)),
            Some(&Variant::Str("Carol".into())),
            "table:number-rows-repeated=\"3\" must advance the row counter by 3, not 1"
        );
    }

    #[test]
    fn diff_ods_all_types() {
        let mut vm = Vm::new();
        vm.cells_mut().insert(
            (1, 1),
            CellContent {
                formula: None,
                value: Variant::Integer(42),
            },
        );
        vm.cells_mut().insert(
            (1, 2),
            CellContent {
                formula: None,
                value: Variant::Str("hello".into()),
            },
        );
        vm.cells_mut().insert(
            (2, 1),
            CellContent {
                formula: None,
                value: Variant::Boolean(true),
            },
        );
        vm.cells_mut().insert(
            (3, 1),
            CellContent {
                formula: None,
                value: Variant::Float(1.5),
            },
        );
        vm.cells_mut().insert(
            (4, 1),
            CellContent {
                formula: None,
                value: Variant::Str(" padded ".into()),
            },
        );

        let path = "/tmp/elixcee_diff_ods.ods";
        save_workbook_impl(&vm, path).unwrap();

        let cal = calamine_ods_cells(path, "sheet1");
        let mine = rd_ods_cells(path, "sheet1");
        assert_eq!(cal, mine, "ODS diff failed");
    }
}
