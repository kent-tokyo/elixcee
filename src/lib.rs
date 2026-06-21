pub mod formula;
pub mod parser;
pub mod vm;

use pyo3::prelude::*;
use pyo3::types::PyDict;
use vm::{serial_to_display, CellContent, Variant, Vm};

// ── Variant ↔ Python conversion ───────────────────────────────────────────────

fn variant_to_py(py: Python<'_>, v: &Variant) -> Py<PyAny> {
    match v {
        Variant::Integer(n) => (*n).into_pyobject(py).unwrap().into_any().unbind(),
        Variant::Float(f)   => (*f).into_pyobject(py).unwrap().into_any().unbind(),
        Variant::Str(s)     => s.as_str().into_pyobject(py).unwrap().into_any().unbind(),
        Variant::Boolean(b) => {
            let borrowed = (*b).into_pyobject(py).unwrap();
            <pyo3::Bound<'_, pyo3::types::PyBool> as Clone>::clone(&borrowed)
                .unbind()
                .into_any()
        }
        Variant::Date(s)    => serial_to_display(*s).into_pyobject(py).unwrap().into_any().unbind(),
        Variant::Error(e)   => e.as_str().into_pyobject(py).unwrap().into_any().unbind(),
        Variant::Empty      => py.None(),
    }
}

fn py_to_variant(obj: &Bound<'_, PyAny>) -> PyResult<Variant> {
    if obj.is_none() { return Ok(Variant::Empty); }
    // bool must come before int (Python bool is a subclass of int)
    if let Ok(b) = obj.extract::<bool>()   { return Ok(Variant::Boolean(b)); }
    if let Ok(n) = obj.extract::<i64>()    { return Ok(Variant::Integer(n)); }
    if let Ok(f) = obj.extract::<f64>()    { return Ok(Variant::Float(f)); }
    if let Ok(s) = obj.extract::<String>() { return Ok(Variant::Str(s)); }
    Err(PyErr::new::<pyo3::exceptions::PyTypeError, _>("Unsupported cell value type"))
}

// ── PyVm class ────────────────────────────────────────────────────────────────

/// VBA execution engine. Create one, pre-populate cells with ``set_cell``,
/// run a macro with ``run``, then read results via ``get_cell`` / ``cells``.
#[pyclass(name = "Vm")]
pub struct PyVm {
    inner: Vm,
}

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
        let prog = parser::parse(vba_code).map_err(|e| {
            PyErr::new::<pyo3::exceptions::PySyntaxError, _>(e.to_string())
        })?;
        self.inner.run_sub(&prog, macro_name).map_err(|e| {
            PyErr::new::<pyo3::exceptions::PyRuntimeError, _>(e)
        })
    }

    /// Write a value into a cell. ``row`` and ``col`` are 1-based (VBA convention).
    fn set_cell(&mut self, row: u32, col: u32, value: &Bound<'_, PyAny>) -> PyResult<()> {
        let v = py_to_variant(value)?;
        self.inner.cells.insert((row, col), CellContent { formula: None, value: v });
        Ok(())
    }

    /// Return the value of a cell (1-based row/col). Returns ``None`` for empty cells.
    fn get_cell(&self, py: Python<'_>, row: u32, col: u32) -> Py<PyAny> {
        variant_to_py(py, &self.inner.get_cell(row, col))
    }

    /// Return all non-empty cells as a dict: ``{(row, col): value}``.
    fn cells(&self, py: Python<'_>) -> PyResult<Py<PyAny>> {
        let dict = PyDict::new(py);
        for ((row, col), content) in &self.inner.cells {
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
        self.inner.set_cell_formula(row, col, formula).map_err(|e| {
            PyErr::new::<pyo3::exceptions::PyValueError, _>(e)
        })
    }

    /// Re-evaluate all cells that have a stored formula.
    fn recalculate(&mut self) -> PyResult<()> {
        self.inner.recalculate_all().map_err(|e| {
            PyErr::new::<pyo3::exceptions::PyRuntimeError, _>(e)
        })
    }
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

/// Load cell data from an Excel file (.xlsx / .xlsm) into a new ``Vm``.
///
/// The VBA source code is **not** extracted from the file — pass it separately
/// to ``vm.run()``.
///
/// Parameters
/// ----------
/// path : str
///     Path to the .xlsx or .xlsm file.
/// sheet : str, optional
///     Sheet name to read. Defaults to the first sheet.
/// on_msgbox : str, optional
///     ``"skip"`` (default) or ``"error"``.
#[pyfunction]
#[pyo3(signature = (path, sheet = None, on_msgbox = "skip"))]
fn load_workbook(path: &str, sheet: Option<&str>, on_msgbox: &str) -> PyResult<PyVm> {
    use calamine::{open_workbook, Data, Reader, Xlsx};

    // Annotate the result type so Rust infers the workbook type correctly.
    let wb_result: Result<Xlsx<_>, _> = open_workbook(path);
    let mut wb = wb_result.map_err(|e| {
        PyErr::new::<pyo3::exceptions::PyIOError, _>(e.to_string())
    })?;

    let sheet_name: String = match sheet {
        Some(s) => s.to_string(),
        None => wb.sheet_names()
            .first()
            .ok_or_else(|| {
                PyErr::new::<pyo3::exceptions::PyValueError, _>("Workbook has no sheets")
            })?
            .clone(),
    };

    let range = wb.worksheet_range(&sheet_name).map_err(|e| {
        PyErr::new::<pyo3::exceptions::PyValueError, _>(e.to_string())
    })?;

    let mut vm = Vm::new();
    vm.error_on_msgbox = on_msgbox == "error";

    for (row, col, cell) in range.cells() {
        // cell: &Data; f/b inside are &f64 / &bool in Rust 2024 edition
        let v = match cell {
            Data::String(s) => Variant::Str(s.clone()),
            Data::Float(f)  => {
                // *f dereferences &f64 → f64
                if f.fract() == 0.0 && *f >= i64::MIN as f64 && *f <= i64::MAX as f64 {
                    Variant::Integer(*f as i64)
                } else {
                    Variant::Float(*f)
                }
            }
            Data::Bool(b)   => Variant::Boolean(*b),
            Data::Empty | Data::Error(_) => continue,
            _ => continue,
        };
        // calamine is 0-based; elixcee/VBA is 1-based
        vm.cells.insert((row as u32 + 1, col as u32 + 1), CellContent { formula: None, value: v });
    }

    Ok(PyVm { inner: vm })
}

#[pyfunction]
fn hello() -> &'static str {
    "Hello from elixcee (Rust)!"
}

// ── Module ────────────────────────────────────────────────────────────────────

#[pymodule]
mod elixcee {
    #[pymodule_export]
    use super::{hello, run_macro, load_workbook, PyVm};
}
