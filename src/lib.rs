pub mod formula;
pub mod parser;
pub mod reader;
pub mod vm;

use pyo3::prelude::*;
use pyo3::types::PyDict;
use vm::{serial_to_display, CellContent, ExcelError, Variant, Vm};

// ── ExcelError Python class ───────────────────────────────────────────────────

/// Represents an Excel cell error value (#N/A, #VALUE!, #DIV/0!, etc.).
/// Returned by ``get_cell`` and ``cells()`` for error cells, and accepted by
/// ``set_cell`` to store an error value.
#[pyclass(name = "ExcelError")]
#[derive(Clone, Debug)]
pub struct PyExcelError {
    /// The error string, e.g. ``"#N/A"``, ``"#VALUE!"``, ``"#DIV/0!"``.
    #[pyo3(get)]
    pub code: String,
}

#[pymethods]
impl PyExcelError {
    #[new]
    fn new(code: String) -> Self { PyExcelError { code } }
    fn __repr__(&self) -> String { format!("ExcelError('{}')", self.code) }
    fn __str__(&self)  -> String { self.code.clone() }
    fn __eq__(&self, other: &PyExcelError) -> bool { self.code == other.code }
    fn __hash__(&self) -> isize { self.code.len() as isize }
}

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
        Variant::Date(s) => {
            let (y, m, d) = crate::formula::eval::serial_to_ymd_pub(*s);
            pyo3::types::PyDate::new(py, y, m as u8, d as u8)
                .map(|dt| dt.into_any().unbind())
                .unwrap_or_else(|_| serial_to_display(*s).into_pyobject(py).unwrap().into_any().unbind())
        }
        Variant::Error(e) => PyExcelError { code: e.as_str().to_string() }
            .into_pyobject(py).unwrap().into_any().unbind(),
        Variant::Empty      => py.None(),
        Variant::Array(a)   => {
            let list = pyo3::types::PyList::new(py, a.iter().map(|x| variant_to_py(py, x))).unwrap();
            list.into_any().unbind()
        }
        Variant::Record(m) => {
            let dict = pyo3::types::PyDict::new(py);
            for (k, v) in m {
                dict.set_item(k, variant_to_py(py, v)).unwrap();
            }
            dict.into_any().unbind()
        }
    }
}

fn py_to_variant(obj: &Bound<'_, PyAny>) -> PyResult<Variant> {
    if obj.is_none() { return Ok(Variant::Empty); }
    // bool must come before int (Python bool is a subclass of int)
    if let Ok(b) = obj.extract::<bool>()   { return Ok(Variant::Boolean(b)); }
    if let Ok(n) = obj.extract::<i64>()    { return Ok(Variant::Integer(n)); }
    if let Ok(f) = obj.extract::<f64>()    { return Ok(Variant::Float(f)); }
    if let Ok(s) = obj.extract::<String>() { return Ok(Variant::Str(s)); }
    if let Ok(e) = obj.extract::<PyExcelError>() {
        return Ok(Variant::Error(match e.code.as_str() {
            "#DIV/0!" => ExcelError::DivZero,
            "#N/A"    => ExcelError::NA,
            "#VALUE!" => ExcelError::Value,
            "#REF!"   => ExcelError::Ref,
            "#NAME?"  => ExcelError::Name,
            "#NUM!"   => ExcelError::Num,
            "#NULL!"  => ExcelError::Null,
            _         => ExcelError::Value,
        }));
    }
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
        self.inner.cells_mut().insert((row, col), CellContent { formula: None, value: v });
        Ok(())
    }

    /// Return the value of a cell (1-based row/col). Returns ``None`` for empty cells.
    fn get_cell(&self, py: Python<'_>, row: u32, col: u32) -> Py<PyAny> {
        variant_to_py(py, &self.inner.get_cell(row, col))
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

    /// Set multiple cell formulas at once.
    /// ``formulas`` should be a dict mapping ``(row, col)`` tuples (1-based) to formula strings.
    fn set_cell_formula_batch(&mut self, formulas: &Bound<'_, PyDict>) -> PyResult<()> {
        for (key, val) in formulas.iter() {
            let (row, col): (u32, u32) = key.extract()
                .map_err(|_| PyErr::new::<pyo3::exceptions::PyTypeError, _>(
                    "keys must be (row, col) tuples of integers"))?;
            let formula: String = val.extract()
                .map_err(|_| PyErr::new::<pyo3::exceptions::PyTypeError, _>(
                    "values must be formula strings"))?;
            self.inner.set_cell_formula(row, col, &formula)
                .map_err(|e| PyErr::new::<pyo3::exceptions::PyValueError, _>(e))?;
        }
        Ok(())
    }

    /// Switch the active sheet. Creates the sheet if it does not exist.
    fn set_sheet(&mut self, name: &str) {
        self.inner.ensure_sheet(name);
        self.inner.active_sheet = name.to_lowercase();
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
        save_workbook_impl(&self.inner, path)
            .map_err(|e| PyErr::new::<pyo3::exceptions::PyIOError, _>(e))
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
            return pd.getattr("DataFrame")?.call0().map(|df| df.into_any().unbind());
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
#[pyfunction]
#[pyo3(signature = (path, sheet = None, on_msgbox = "skip"))]
fn load_workbook(path: &str, sheet: Option<&str>, on_msgbox: &str) -> PyResult<PyVm> {
    let sheets = reader::read_workbook(path).map_err(|e| {
        PyErr::new::<pyo3::exceptions::PyIOError, _>(e)
    })?;

    if sheets.is_empty() {
        return Err(PyErr::new::<pyo3::exceptions::PyValueError, _>("Workbook has no sheets"));
    }

    let mut vm = Vm::new();
    vm.error_on_msgbox = on_msgbox == "error";

    for sheet_data in &sheets {
        let key = sheet_data.name.clone();
        vm.ensure_sheet(&key);
        let prev = vm.active_sheet.clone();
        vm.active_sheet = key;
        for (&(row, col), cell) in &sheet_data.cells {
            let value = match cell {
                reader::SheetCell::Integer(n) => Variant::Integer(*n),
                reader::SheetCell::Float(f)   => Variant::Float(*f),
                reader::SheetCell::Str(s)     => Variant::Str(s.clone()),
                reader::SheetCell::Bool(b)    => Variant::Boolean(*b),
            };
            vm.cells_mut().insert((row, col), CellContent { formula: None, value });
        }
        vm.active_sheet = prev;
    }

    let active = match sheet {
        Some(s) => s.to_lowercase(),
        None    => sheets[0].name.clone(),
    };
    vm.set_active_sheet(&active).map_err(|e| {
        PyErr::new::<pyo3::exceptions::PyValueError, _>(e)
    })?;

    Ok(PyVm { inner: vm })
}

#[pyfunction]
fn hello() -> &'static str {
    "Hello from elixcee (Rust)!"
}

// ── save_workbook implementation ─────────────────────────────────────────────

fn save_workbook_impl(vm: &Vm, path: &str) -> Result<(), String> {
    if path.to_lowercase().ends_with(".ods") {
        return save_ods_impl(vm, path);
    }
    save_xlsx_impl(vm, path)
}

fn save_xlsx_impl(vm: &Vm, path: &str) -> Result<(), String> {
    use zip::write::ZipWriter;
    use zip::CompressionMethod;
    use std::io::{Write, Cursor};
    use std::collections::HashMap;

    let sheet_names = vm.sheet_names();

    // Collect shared strings (insertion-ordered, deduplicated)
    let mut str_index: HashMap<String, usize> = HashMap::new();
    let mut shared_strings: Vec<String> = Vec::new();
    for sheet_name in &sheet_names {
        if let Some(cells) = vm.get_sheet_cells(sheet_name) {
            let mut sorted: Vec<_> = cells.keys().collect();
            sorted.sort();
            for key in sorted {
                let s = match &cells[key].value {
                    Variant::Str(s)  => s.as_str().to_string(),
                    Variant::Error(e) => e.as_str().to_string(),
                    _ => continue,
                };
                if !str_index.contains_key(&s) {
                    str_index.insert(s.clone(), shared_strings.len());
                    shared_strings.push(s);
                }
            }
        }
    }

    let cursor = Cursor::new(Vec::<u8>::new());
    let mut zip = ZipWriter::new(cursor);
    let deflated = zip::write::SimpleFileOptions::default()
        .compression_method(CompressionMethod::Deflated);

    zip.start_file("[Content_Types].xml", deflated).map_err(|e| e.to_string())?;
    zip.write_all(build_xlsx_content_types(&sheet_names).as_bytes()).map_err(|e| e.to_string())?;

    zip.start_file("_rels/.rels", deflated).map_err(|e| e.to_string())?;
    zip.write_all(XLSX_ROOT_RELS.as_bytes()).map_err(|e| e.to_string())?;

    zip.start_file("xl/workbook.xml", deflated).map_err(|e| e.to_string())?;
    zip.write_all(build_xlsx_workbook(&sheet_names).as_bytes()).map_err(|e| e.to_string())?;

    zip.start_file("xl/_rels/workbook.xml.rels", deflated).map_err(|e| e.to_string())?;
    zip.write_all(build_xlsx_workbook_rels(&sheet_names).as_bytes()).map_err(|e| e.to_string())?;

    for (i, sheet_name) in sheet_names.iter().enumerate() {
        zip.start_file(format!("xl/worksheets/sheet{}.xml", i + 1), deflated).map_err(|e| e.to_string())?;
        zip.write_all(build_xlsx_sheet(vm, sheet_name, &str_index).as_bytes()).map_err(|e| e.to_string())?;
    }

    zip.start_file("xl/sharedStrings.xml", deflated).map_err(|e| e.to_string())?;
    zip.write_all(build_xlsx_shared_strings(&shared_strings).as_bytes()).map_err(|e| e.to_string())?;

    zip.start_file("xl/styles.xml", deflated).map_err(|e| e.to_string())?;
    zip.write_all(XLSX_STYLES.as_bytes()).map_err(|e| e.to_string())?;

    let data = zip.finish().map_err(|e| e.to_string())?.into_inner();
    std::fs::write(path, data).map_err(|e| e.to_string())?;
    Ok(())
}

const XLSX_ROOT_RELS: &str = concat!(
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n",
    "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n",
    "<Relationship Id=\"rId1\" ",
      "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" ",
      "Target=\"xl/workbook.xml\"/>\n",
    "</Relationships>\n",
);

const XLSX_STYLES: &str = concat!(
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n",
    "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n",
    "<fonts><font/></fonts>\n",
    "<fills><fill/><fill/></fills>\n",
    "<borders><border/></borders>\n",
    "<cellStyleXfs><xf/></cellStyleXfs>\n",
    "<cellXfs><xf/></cellXfs>\n",
    "</styleSheet>\n",
);

fn build_xlsx_content_types(sheet_names: &[String]) -> String {
    let mut out = String::from(concat!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n",
        "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\n",
        "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>\n",
        "<Default Extension=\"xml\" ContentType=\"application/xml\"/>\n",
        "<Override PartName=\"/xl/workbook.xml\" ",
          "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>\n",
    ));
    for (i, _) in sheet_names.iter().enumerate() {
        out.push_str(&format!(
            "<Override PartName=\"/xl/worksheets/sheet{}.xml\" \
             ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>\n",
            i + 1
        ));
    }
    out.push_str(concat!(
        "<Override PartName=\"/xl/sharedStrings.xml\" ",
          "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>\n",
        "<Override PartName=\"/xl/styles.xml\" ",
          "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>\n",
        "</Types>\n",
    ));
    out
}

fn build_xlsx_workbook(sheet_names: &[String]) -> String {
    let mut out = String::from(concat!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n",
        "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" ",
          "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n",
        "<sheets>\n",
    ));
    for (i, name) in sheet_names.iter().enumerate() {
        let n = i + 1;
        out.push_str(&format!(
            "<sheet name=\"{}\" sheetId=\"{}\" r:id=\"rId{}\"/>\n",
            xml_escape(name), n, n
        ));
    }
    out.push_str("</sheets>\n</workbook>\n");
    out
}

fn build_xlsx_workbook_rels(sheet_names: &[String]) -> String {
    let mut out = String::from(concat!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n",
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n",
    ));
    for (i, _) in sheet_names.iter().enumerate() {
        let n = i + 1;
        out.push_str(&format!(
            "<Relationship Id=\"rId{}\" \
             Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" \
             Target=\"worksheets/sheet{}.xml\"/>\n",
            n, n
        ));
    }
    let ss_id = sheet_names.len() + 1;
    let styles_id = sheet_names.len() + 2;
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
    out.push_str("</Relationships>\n");
    out
}

fn build_xlsx_sheet(vm: &Vm, sheet_name: &str, str_index: &std::collections::HashMap<String, usize>) -> String {
    let mut out = String::from(concat!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n",
        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n",
        "<sheetData>\n",
    ));

    if let Some(cells) = vm.get_sheet_cells(sheet_name) {
        if !cells.is_empty() {
            // Group by row first to avoid O(max_row × total_cells) scanning.
            let mut by_row: std::collections::BTreeMap<u32, Vec<_>> = std::collections::BTreeMap::new();
            for (k @ &(r, c), v) in cells.iter() {
                if r > 0 && c > 0 { by_row.entry(r).or_default().push((k, v)); }
            }
            for (row, mut row_cells) in by_row {
                row_cells.sort_by_key(|&(&(_, c), _)| c);

                out.push_str(&format!("<row r=\"{}\">\n", row));
                for (&(r, c), content) in row_cells {
                    let cell_ref = format!("{}{}", xlsx_col_letters(c), r);
                    if let Some(xml) = xlsx_cell_xml(&cell_ref, &content.value, str_index) {
                        out.push_str(&xml);
                        out.push('\n');
                    }
                }
                out.push_str("</row>\n");
            }
        }
    }

    out.push_str("</sheetData>\n</worksheet>\n");
    out
}

fn xlsx_cell_xml(cell_ref: &str, v: &Variant, str_index: &std::collections::HashMap<String, usize>) -> Option<String> {
    match v {
        Variant::Integer(n) => Some(format!("<c r=\"{}\"><v>{}</v></c>", cell_ref, n)),
        Variant::Float(f)   => Some(format!("<c r=\"{}\"><v>{}</v></c>", cell_ref, f)),
        Variant::Date(s)    => Some(format!("<c r=\"{}\"><v>{}</v></c>", cell_ref, s)),
        Variant::Str(s) => {
            let idx = str_index[s.as_str()];
            Some(format!("<c r=\"{}\" t=\"s\"><v>{}</v></c>", cell_ref, idx))
        }
        Variant::Error(e) => {
            let idx = str_index[e.as_str()];
            Some(format!("<c r=\"{}\" t=\"s\"><v>{}</v></c>", cell_ref, idx))
        }
        Variant::Boolean(b) => Some(format!(
            "<c r=\"{}\" t=\"b\"><v>{}</v></c>", cell_ref, if *b { 1 } else { 0 }
        )),
        Variant::Empty | Variant::Array(_) | Variant::Record(_) => None,
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
        out.push_str(&format!("<si><t>{}</t></si>\n", xml_escape(s)));
    }
    out.push_str("</sst>\n");
    out
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

// ── ODS write ────────────────────────────────────────────────────────────────

fn save_ods_impl(vm: &Vm, path: &str) -> Result<(), String> {
    use zip::write::ZipWriter;
    use zip::CompressionMethod;
    use std::io::{Write, Cursor};

    let cursor = Cursor::new(Vec::<u8>::new());
    let mut zip = ZipWriter::new(cursor);
    let stored = zip::write::SimpleFileOptions::default()
        .compression_method(CompressionMethod::Stored);
    let deflated = zip::write::SimpleFileOptions::default()
        .compression_method(CompressionMethod::Deflated);

    // 1. mimetype (must be STORE and first entry per ODF spec)
    zip.start_file("mimetype", stored).map_err(|e| e.to_string())?;
    zip.write_all(b"application/vnd.oasis.opendocument.spreadsheet")
        .map_err(|e| e.to_string())?;

    // 2. META-INF/manifest.xml
    let manifest = build_ods_manifest(vm);
    zip.start_file("META-INF/manifest.xml", deflated).map_err(|e| e.to_string())?;
    zip.write_all(manifest.as_bytes()).map_err(|e| e.to_string())?;

    // 3. content.xml
    let content = build_ods_content(vm);
    zip.start_file("content.xml", deflated).map_err(|e| e.to_string())?;
    zip.write_all(content.as_bytes()).map_err(|e| e.to_string())?;

    let data = zip.finish().map_err(|e| e.to_string())?.into_inner();
    std::fs::write(path, data).map_err(|e| e.to_string())?;
    Ok(())
}

fn build_ods_manifest(vm: &Vm) -> String {
    let mut m = String::from(concat!(
        r#"<?xml version="1.0" encoding="UTF-8"?>"#, "\n",
        r#"<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2">"#, "\n",
        r#" <manifest:file-entry manifest:media-type="application/vnd.oasis.opendocument.spreadsheet" manifest:version="1.2" manifest:full-path="/"/>"#, "\n",
        r#" <manifest:file-entry manifest:media-type="text/xml" manifest:full-path="content.xml"/>"#, "\n",
    ));
    m.push_str("</manifest:manifest>\n");
    m
}

fn build_ods_content(vm: &Vm) -> String {
    let mut out = String::from(concat!(
        r#"<?xml version="1.0" encoding="UTF-8"?>"#, "\n",
        r#"<office:document-content"#,
        r#" xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0""#,
        r#" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0""#,
        r#" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0""#,
        r#" office:version="1.2">"#, "\n",
        r#"<office:body><office:spreadsheet>"#, "\n",
    ));

    for sheet_name in vm.sheet_names() {
        let escaped = xml_escape(&sheet_name);
        out.push_str(&format!("<table:table table:name=\"{}\">\n", escaped));

        if let Some(cells) = vm.get_sheet_cells(&sheet_name) {
            if !cells.is_empty() {
                let max_row = cells.keys().map(|(r,_)| *r).max().unwrap_or(0);
                let max_col = cells.keys().map(|(_,c)| *c).max().unwrap_or(0);

                for r in 1..=max_row {
                    out.push_str("<table:table-row>");
                    for c in 1..=max_col {
                        let cell_xml = match cells.get(&(r, c)) {
                            None | Some(vm::CellContent { value: Variant::Empty, .. }) => {
                                "<table:table-cell/>".to_string()
                            }
                            Some(content) => ods_cell_xml(&content.value),
                        };
                        out.push_str(&cell_xml);
                    }
                    out.push_str("</table:table-row>\n");
                }
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
                bv, if *b { "TRUE" } else { "FALSE" }
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
        Variant::Empty | Variant::Array(_) | Variant::Record(_) => "<table:table-cell/>".to_string(),
    }
}

fn xml_escape(s: &str) -> String {
    s.replace('&', "&amp;")
     .replace('<', "&lt;")
     .replace('>', "&gt;")
     .replace('"', "&quot;")
     .replace('\'', "&apos;")
}

// ── Module ────────────────────────────────────────────────────────────────────

#[pymodule]
mod elixcee {
    #[pymodule_export]
    use super::{hello, run_macro, load_workbook, PyVm, PyExcelError};
}

#[cfg(test)]
mod tests {
    use super::*;
    use calamine::{open_workbook, Reader, Xlsx};

    #[test]
    fn test_save_workbook_roundtrip() {
        let mut vm = Vm::new();
        vm.cells_mut().insert((1, 1), CellContent { formula: None, value: Variant::Integer(42) });
        vm.cells_mut().insert((2, 1), CellContent { formula: None, value: Variant::Str("hello".into()) });
        vm.cells_mut().insert((3, 1), CellContent { formula: None, value: Variant::Float(3.14) });
        vm.cells_mut().insert((4, 1), CellContent { formula: None, value: Variant::Boolean(true) });

        let path = "/tmp/elixcee_test_roundtrip.xlsx";
        save_workbook_impl(&vm, path).expect("save should succeed");

        // Reload with calamine and verify
        let mut wb: Xlsx<_> = open_workbook(path).expect("open should succeed");
        let range = wb.worksheet_range("sheet1").expect("sheet1 should exist");
        let cells: Vec<_> = range.cells().collect();
        assert!(!cells.is_empty(), "saved file should have cells");
    }

    #[test]
    fn test_save_ods_roundtrip() {
        use calamine::{open_workbook_auto, Reader};

        let mut vm = Vm::new();
        vm.cells_mut().insert((1, 1), CellContent { formula: None, value: Variant::Integer(42) });
        vm.cells_mut().insert((1, 2), CellContent { formula: None, value: Variant::Str("hello".into()) });
        vm.cells_mut().insert((2, 1), CellContent { formula: None, value: Variant::Boolean(true) });

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
        vm.cells_mut().insert((1, 1), CellContent { formula: None, value: Variant::Integer(1) });
        vm.ensure_sheet("sheet2");
        let prev = vm.active_sheet.clone();
        vm.active_sheet = "sheet2".into();
        vm.cells_mut().insert((1, 1), CellContent { formula: None, value: Variant::Integer(2) });
        vm.active_sheet = prev;

        let path = "/tmp/elixcee_test_multisheet.xlsx";
        save_workbook_impl(&vm, path).expect("save should succeed");

        let mut wb: Xlsx<_> = open_workbook(path).expect("open should succeed");
        assert!(wb.worksheet_range("sheet1").is_ok(), "sheet1 should exist");
        assert!(wb.worksheet_range("sheet2").is_ok(), "sheet2 should exist");
    }
}

// ── Differential read tests: calamine (oracle) vs hand-written reader ─────────
#[cfg(test)]
mod diff_reader_tests {
    use super::*;
    use calamine::{Data, Reader, open_workbook, open_workbook_auto, Xlsx};
    use crate::reader::{read_workbook as rd, SheetCell};

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
            SheetCell::Float(f)   => Variant::Float(*f),
            SheetCell::Str(s)     => Variant::Str(s.clone()),
            SheetCell::Bool(b)    => Variant::Boolean(*b),
        }
    }

    fn calamine_xlsx_cells(path: &str, sheet: &str) -> std::collections::HashMap<(u32,u32), Variant> {
        let mut wb: Xlsx<_> = open_workbook(path).unwrap();
        let range = wb.worksheet_range(sheet).unwrap();
        let (sr, sc) = range.start().unwrap_or((0, 0));
        range.cells()
            .filter_map(|(r, c, d)| calamine_cell_to_variant(d)
                .map(|v| ((r as u32 + sr + 1, c as u32 + sc + 1), v)))
            .collect()
    }

    fn rd_xlsx_cells(path: &str, sheet: &str) -> std::collections::HashMap<(u32,u32), Variant> {
        rd(path).unwrap().into_iter()
            .find(|s| s.name == sheet).unwrap()
            .cells.iter()
            .map(|(&k, v)| (k, rd_cell_to_variant(v)))
            .collect()
    }

    fn calamine_ods_cells(path: &str, sheet: &str) -> std::collections::HashMap<(u32,u32), Variant> {
        let mut wb = open_workbook_auto(path).unwrap();
        let range = wb.worksheet_range(sheet).unwrap();
        let (sr, sc) = range.start().unwrap_or((0, 0));
        range.cells()
            .filter_map(|(r, c, d)| calamine_cell_to_variant(d)
                .map(|v| ((r as u32 + sr + 1, c as u32 + sc + 1), v)))
            .collect()
    }

    fn rd_ods_cells(path: &str, sheet: &str) -> std::collections::HashMap<(u32,u32), Variant> {
        rd(path).unwrap().into_iter()
            .find(|s| s.name == sheet).unwrap()
            .cells.iter()
            .map(|(&k, v)| (k, rd_cell_to_variant(v)))
            .collect()
    }

    #[test]
    fn diff_xlsx_all_types() {
        let mut vm = Vm::new();
        vm.cells_mut().insert((1, 1), CellContent { formula: None, value: Variant::Integer(42) });
        vm.cells_mut().insert((2, 1), CellContent { formula: None, value: Variant::Str("hello".into()) });
        vm.cells_mut().insert((3, 1), CellContent { formula: None, value: Variant::Float(3.14) });
        vm.cells_mut().insert((4, 1), CellContent { formula: None, value: Variant::Boolean(true) });
        vm.cells_mut().insert((5, 1), CellContent { formula: None, value: Variant::Str(" leading and trailing ".into()) });

        let path = "/tmp/elixcee_diff_xlsx.xlsx";
        save_workbook_impl(&vm, path).unwrap();

        let cal = calamine_xlsx_cells(path, "sheet1");
        let mine = rd_xlsx_cells(path, "sheet1");
        assert_eq!(cal, mine, "XLSX diff failed");
    }

    #[test]
    fn diff_xlsx_multi_sheet() {
        let mut vm = Vm::new();
        vm.cells_mut().insert((1, 1), CellContent { formula: None, value: Variant::Integer(1) });
        vm.ensure_sheet("sheet2");
        let prev = vm.active_sheet.clone();
        vm.active_sheet = "sheet2".into();
        vm.cells_mut().insert((2, 3), CellContent { formula: None, value: Variant::Str("s2".into()) });
        vm.active_sheet = prev;

        let path = "/tmp/elixcee_diff_multi.xlsx";
        save_workbook_impl(&vm, path).unwrap();

        for sheet in &["sheet1", "sheet2"] {
            let cal = calamine_xlsx_cells(path, sheet);
            let mine = rd_xlsx_cells(path, sheet);
            assert_eq!(cal, mine, "XLSX multi-sheet diff failed for {}", sheet);
        }
    }

    #[test]
    fn diff_ods_all_types() {
        let mut vm = Vm::new();
        vm.cells_mut().insert((1, 1), CellContent { formula: None, value: Variant::Integer(42) });
        vm.cells_mut().insert((1, 2), CellContent { formula: None, value: Variant::Str("hello".into()) });
        vm.cells_mut().insert((2, 1), CellContent { formula: None, value: Variant::Boolean(true) });
        vm.cells_mut().insert((3, 1), CellContent { formula: None, value: Variant::Float(1.5) });
        vm.cells_mut().insert((4, 1), CellContent { formula: None, value: Variant::Str(" padded ".into()) });

        let path = "/tmp/elixcee_diff_ods.ods";
        save_workbook_impl(&vm, path).unwrap();

        let cal = calamine_ods_cells(path, "sheet1");
        let mine = rd_ods_cells(path, "sheet1");
        assert_eq!(cal, mine, "ODS diff failed");
    }
}
