//! Forward-only row streaming for large XLSX files (0.25.0).
//!
//! The normal `Vm` intentionally remains a random-access, fully mutable model. This
//! module provides a separate pipeline API whose worker owns the ZIP entry and sends
//! one decoded row at a time, so callers do not need to materialize a workbook.

use std::fs::File;
use std::io::{BufRead, BufReader};
use std::sync::{
    Mutex,
    mpsc::{self, Receiver, RecvTimeoutError},
};
use std::time::Duration;

use pyo3::prelude::*;
use pyo3::types::{PyAny, PyList};

use crate::reader::{self, SheetCell};
use crate::{Variant, variant_to_py};

const STREAM_BUFFER_BYTES: usize = 64 * 1024;
/// Refuse an unterminated or hostile worksheet row before it can grow without bound.
/// This keeps the forward-only API bounded even when given malformed XML.
const MAX_STREAM_ROW_BYTES: usize = 16 * 1024 * 1024;
/// Bound the pending rows retained by the append-only writer before `close()`
/// materializes them into the normal workbook writer.
const MAX_STREAM_WRITER_BYTES: usize = 64 * 1024 * 1024;

fn estimated_variant_bytes(value: &Variant) -> usize {
    match value {
        Variant::Str(text) => text.len(),
        Variant::Array(values) => values.iter().fold(0, |total, value| {
            total.saturating_add(estimated_variant_bytes(value))
        }),
        Variant::VbaArray(values) => values.elements.iter().fold(0, |total, value| {
            total.saturating_add(estimated_variant_bytes(value))
        }),
        Variant::Record(values) => values.iter().fold(0, |total, (key, value)| {
            total
                .saturating_add(key.len())
                .saturating_add(estimated_variant_bytes(value))
        }),
        _ => std::mem::size_of_val(value),
    }
}

fn append_row_token(
    row_buf: &mut Vec<u8>,
    token: &[u8],
    max_row_bytes: usize,
) -> Result<(), String> {
    if row_buf.len().saturating_add(token.len()) > max_row_bytes {
        return Err(format!(
            "worksheet row exceeds the streaming limit of {max_row_bytes} bytes"
        ));
    }
    row_buf.extend_from_slice(token);
    Ok(())
}

/// Resolve a worksheet relationship target relative to `xl/`, as required by
/// OOXML package relationships. ZIP entry names always use `/`, so doing this
/// explicitly also keeps `..` from producing a name that can never be found.
fn resolve_xlsx_target(target: &str) -> Result<String, String> {
    let base = if target.starts_with('/') {
        target.trim_start_matches('/').to_string()
    } else {
        format!("xl/{target}")
    };
    let mut parts = Vec::new();
    for part in base.split('/') {
        match part {
            "" | "." => {}
            ".." => {
                if parts.pop().is_none() {
                    return Err(format!("worksheet relationship escapes ZIP root: {target}"));
                }
            }
            part => parts.push(part),
        }
    }
    if parts.is_empty() {
        return Err(format!(
            "worksheet relationship has an empty target: {target}"
        ));
    }
    Ok(parts.join("/"))
}

fn sheet_target(path: &str, requested: Option<&str>) -> Result<(String, Vec<String>), String> {
    let file = File::open(path).map_err(|e| e.to_string())?;
    let mut archive = zip::ZipArchive::new(file).map_err(|e| e.to_string())?;
    reader::validate_zip_archive_for_stream(&mut archive)?;
    let workbook = reader::zip_read_text_for_stream(&mut archive, "xl/workbook.xml")?;
    let refs = reader::xlsx_workbook_sheets_for_stream(&workbook);
    let rels_xml = reader::zip_read_text_for_stream(&mut archive, "xl/_rels/workbook.xml.rels")?;
    let rels = reader::xlsx_rels_for_stream(&rels_xml, "/worksheet");
    let names = refs.iter().map(|r| r.0.clone()).collect::<Vec<_>>();
    let chosen = requested
        .map(|name| {
            refs.iter()
                .find(|r| r.0.eq_ignore_ascii_case(name))
                .ok_or_else(|| format!("unknown sheet: {name}"))
        })
        .unwrap_or_else(|| {
            refs.first()
                .ok_or_else(|| "Workbook has no sheets".to_string())
        })?;
    let target = rels
        .get(&chosen.1)
        .ok_or_else(|| format!("worksheet relationship is missing for {}", chosen.0))?;
    let zip_path = resolve_xlsx_target(target)?;
    Ok((zip_path, names))
}

fn cell_variant(cell: &SheetCell) -> Variant {
    match cell {
        SheetCell::Integer(v) => Variant::Integer(*v),
        SheetCell::Float(v) => Variant::Float(*v),
        SheetCell::Str(v) => Variant::Str(v.clone()),
        SheetCell::Bool(v) => Variant::Boolean(*v),
        SheetCell::Error(v) => Variant::Error(v.clone()),
    }
}

fn row_from_xml_with_limit(
    xml: &str,
    shared: &[String],
    max_columns: usize,
) -> Result<Option<(u32, Vec<Variant>)>, String> {
    let parsed = reader::xlsx_sheet_cells_for_stream(xml, shared);
    let Some(row) = parsed.first_row else {
        return Ok(None);
    };
    let max_col = parsed.cells.keys().map(|(_, col)| *col).max().unwrap_or(0);
    if max_col as usize > max_columns {
        return Err(format!(
            "worksheet row exceeds the streaming limit of {max_columns} columns"
        ));
    }
    let mut out = vec![Variant::Empty; max_col as usize];
    for ((_, col), cell) in parsed.cells {
        if col > 0 {
            out[(col - 1) as usize] = cell_variant(&cell);
        }
    }
    Ok(Some((row, out)))
}

#[cfg(test)]
fn row_from_xml(xml: &str, shared: &[String]) -> Option<(u32, Vec<Variant>)> {
    row_from_xml_with_limit(xml, shared, usize::MAX)
        .ok()
        .flatten()
}

fn is_row_close(token: &[u8]) -> bool {
    let token = token.trim_ascii_start();
    let Some(rest) = token.strip_prefix(b"</row") else {
        return false;
    };
    rest.iter().all(|b| b.is_ascii_whitespace() || *b == b'>') && rest.contains(&b'>')
}

fn parse_stream_row(
    row_buf: &[u8],
    shared: &[String],
    max_columns: usize,
    wrap_in_sheet_data: bool,
) -> Result<Option<(u32, Vec<Variant>)>, String> {
    let xml =
        std::str::from_utf8(row_buf).map_err(|_| "worksheet row is not valid UTF-8".to_string())?;
    reader::validate_shared_string_refs_for_stream(xml, shared)?;
    if wrap_in_sheet_data {
        let wrapped = format!("<worksheet><sheetData>{xml}</sheetData></worksheet>");
        row_from_xml_with_limit(&wrapped, shared, max_columns)
    } else {
        row_from_xml_with_limit(xml, shared, max_columns)
    }
}

fn stream_rows(
    path: String,
    sheet: Option<String>,
    max_row_bytes: usize,
    max_columns: usize,
) -> Result<Receiver<Result<(u32, Vec<Variant>), String>>, String> {
    let (zip_path, _) = sheet_target(&path, sheet.as_deref())?;
    let (tx, rx) = mpsc::sync_channel(2);
    std::thread::spawn(move || {
        let result = (|| -> Result<(), String> {
            let file = File::open(&path).map_err(|e| e.to_string())?;
            let mut archive = zip::ZipArchive::new(file).map_err(|e| e.to_string())?;
            reader::validate_zip_archive_for_stream(&mut archive)?;
            let shared_xml = if archive
                .file_names()
                .any(|name| name == "xl/sharedStrings.xml")
            {
                reader::zip_read_text_for_stream(&mut archive, "xl/sharedStrings.xml")?
            } else {
                String::new()
            };
            let shared = reader::xlsx_shared_strings_for_stream(&shared_xml);
            reader::validate_shared_strings_for_stream(&shared)?;
            let entry = archive.by_name(&zip_path).map_err(|e| e.to_string())?;
            let mut input = BufReader::with_capacity(STREAM_BUFFER_BYTES, entry);
            let mut row_buf = Vec::with_capacity(128 * 1024);
            let mut token = Vec::with_capacity(1024);
            let mut in_row = false;
            loop {
                token.clear();
                if input
                    .read_until(b'>', &mut token)
                    .map_err(|e| e.to_string())?
                    == 0
                {
                    break;
                }
                if !in_row {
                    let trimmed = token.trim_ascii_start();
                    if trimmed.starts_with(b"<row")
                        && trimmed
                            .get(4)
                            .is_some_and(|b| *b == b' ' || *b == b'>' || *b == b'/')
                    {
                        in_row = true;
                        row_buf.clear();
                        append_row_token(&mut row_buf, &token, max_row_bytes)?;
                        if trimmed.ends_with(b"/>") {
                            if let Some((row_number, row)) =
                                parse_stream_row(&row_buf, &shared, max_columns, false)?
                                && tx.send(Ok((row_number, row))).is_err()
                            {
                                return Ok(());
                            }
                            in_row = false;
                        }
                    }
                } else {
                    append_row_token(&mut row_buf, &token, max_row_bytes)?;
                    if is_row_close(&token) {
                        if let Some((row_number, row)) =
                            parse_stream_row(&row_buf, &shared, max_columns, true)?
                            && tx.send(Ok((row_number, row))).is_err()
                        {
                            return Ok(());
                        }
                        in_row = false;
                    }
                }
            }
            if in_row {
                return Err("worksheet row is unterminated".to_string());
            }
            Ok(())
        })();
        if let Err(err) = result {
            let _ = tx.send(Err(err));
        }
    });
    Ok(rx)
}

fn validate_max_rows(max_rows: Option<usize>) -> Result<(), &'static str> {
    if max_rows == Some(0) {
        Err("max_rows must be greater than zero")
    } else {
        Ok(())
    }
}

fn validate_max_row_bytes(max_row_bytes: Option<usize>) -> Result<(), &'static str> {
    if max_row_bytes == Some(0) {
        Err("max_row_bytes must be greater than zero")
    } else {
        Ok(())
    }
}

fn validate_max_columns(max_columns: Option<usize>) -> Result<(), &'static str> {
    if max_columns == Some(0) {
        Err("max_columns must be greater than zero")
    } else {
        Ok(())
    }
}

fn validate_max_writer_rows(max_rows: Option<usize>) -> Result<(), &'static str> {
    if max_rows == Some(0) {
        Err("max_rows must be greater than zero")
    } else {
        Ok(())
    }
}

fn validate_writer_row_columns(row_len: usize, max_columns: Option<usize>) -> Result<(), String> {
    if max_columns.is_some_and(|max_columns| row_len > max_columns) {
        let max_columns = max_columns.expect("checked above");
        Err(format!(
            "stream writer row exceeds the limit of {max_columns} columns"
        ))
    } else {
        Ok(())
    }
}

fn validate_timeout_ms(timeout_ms: Option<u64>) -> Result<(), &'static str> {
    if timeout_ms == Some(0) {
        Err("timeout_ms must be greater than zero")
    } else {
        Ok(())
    }
}

#[pyclass(name = "StreamReader")]
pub struct PyStreamReader {
    receiver: Option<RowReceiver>,
    include_row_numbers: bool,
    max_rows: Option<usize>,
    timeout_ms: Option<u64>,
    rows_read: usize,
}

type RowReceiver = Mutex<Receiver<Result<(u32, Vec<Variant>), String>>>;

pub(crate) fn stream_reader_from_path(
    path: &str,
    sheet: Option<&str>,
    include_row_numbers: bool,
    max_rows: Option<usize>,
    max_row_bytes: Option<usize>,
    max_columns: Option<usize>,
    timeout_ms: Option<u64>,
) -> PyResult<PyStreamReader> {
    validate_max_rows(max_rows).map_err(PyErr::new::<pyo3::exceptions::PyValueError, _>)?;
    validate_max_row_bytes(max_row_bytes)
        .map_err(PyErr::new::<pyo3::exceptions::PyValueError, _>)?;
    validate_max_columns(max_columns).map_err(PyErr::new::<pyo3::exceptions::PyValueError, _>)?;
    validate_timeout_ms(timeout_ms).map_err(PyErr::new::<pyo3::exceptions::PyValueError, _>)?;
    let receiver = stream_rows(
        path.to_string(),
        sheet.map(str::to_string),
        max_row_bytes.unwrap_or(MAX_STREAM_ROW_BYTES),
        max_columns.unwrap_or(16_384),
    )
    .map_err(PyErr::new::<pyo3::exceptions::PyIOError, _>)?;
    Ok(PyStreamReader {
        receiver: Some(Mutex::new(receiver)),
        include_row_numbers,
        max_rows,
        timeout_ms,
        rows_read: 0,
    })
}

#[pymethods]
impl PyStreamReader {
    #[new]
    #[pyo3(signature = (path, sheet = None, include_row_numbers = false, max_rows = None, max_row_bytes = None, max_columns = None, timeout_ms = None))]
    fn new(
        path: &str,
        sheet: Option<&str>,
        include_row_numbers: bool,
        max_rows: Option<usize>,
        max_row_bytes: Option<usize>,
        max_columns: Option<usize>,
        timeout_ms: Option<u64>,
    ) -> PyResult<Self> {
        stream_reader_from_path(
            path,
            sheet,
            include_row_numbers,
            max_rows,
            max_row_bytes,
            max_columns,
            timeout_ms,
        )
    }
    fn __enter__(slf: PyRef<'_, Self>) -> PyRef<'_, Self> {
        slf
    }
    fn __exit__(
        &mut self,
        _ty: Option<&Bound<'_, PyAny>>,
        _value: Option<&Bound<'_, PyAny>>,
        _tb: Option<&Bound<'_, PyAny>>,
    ) -> bool {
        self.close();
        false
    }
    /// Stop reading and release the worker receiver.
    fn close(&mut self) {
        self.receiver = None;
    }
    /// Whether this reader has been explicitly closed or exhausted.
    #[getter]
    fn closed(&self) -> bool {
        self.receiver.is_none()
    }
    fn __iter__(slf: PyRef<'_, Self>) -> PyRef<'_, Self> {
        slf
    }
    fn __next__(&mut self, py: Python<'_>) -> PyResult<Option<Py<PyAny>>> {
        let Some(receiver) = &self.receiver else {
            return Ok(None);
        };
        let next = {
            let receiver = receiver.lock().expect("stream reader mutex poisoned");
            match self.timeout_ms {
                Some(timeout_ms) => receiver.recv_timeout(Duration::from_millis(timeout_ms)),
                None => receiver.recv().map_err(|_| RecvTimeoutError::Disconnected),
            }
        };
        match next {
            Ok(Ok((row_number, row))) => {
                let values = PyList::new(py, row.iter().map(|v| variant_to_py(py, v)))?
                    .into_any()
                    .unbind();
                self.rows_read += 1;
                if self
                    .max_rows
                    .is_some_and(|max_rows| self.rows_read >= max_rows)
                {
                    self.receiver = None;
                }
                if self.include_row_numbers {
                    Ok(Some(
                        (row_number, values).into_pyobject(py)?.into_any().unbind(),
                    ))
                } else {
                    Ok(Some(values))
                }
            }
            Ok(Err(err)) => Err(PyErr::new::<pyo3::exceptions::PyIOError, _>(err)),
            Err(RecvTimeoutError::Timeout) => {
                Err(PyErr::new::<pyo3::exceptions::PyTimeoutError, _>(format!(
                    "stream reader timed out after {} ms",
                    self.timeout_ms.expect("timeout error requires a timeout")
                )))
            }
            Err(RecvTimeoutError::Disconnected) => {
                self.receiver = None;
                Ok(None)
            }
        }
    }
}

#[pyclass(name = "StreamWriter")]
pub struct PyStreamWriter {
    path: String,
    rows: Vec<Vec<Variant>>,
    pending_bytes: usize,
    max_pending_bytes: usize,
    max_rows: Option<usize>,
    max_columns: Option<usize>,
    active: bool,
}

pub(crate) fn stream_writer_from_path(
    path: &str,
    max_pending_bytes: Option<usize>,
    max_rows: Option<usize>,
    max_columns: Option<usize>,
) -> PyResult<PyStreamWriter> {
    let max_pending_bytes = max_pending_bytes.unwrap_or(MAX_STREAM_WRITER_BYTES);
    if max_pending_bytes == 0 {
        return Err(PyErr::new::<pyo3::exceptions::PyValueError, _>(
            "max_pending_bytes must be greater than zero",
        ));
    }
    validate_max_writer_rows(max_rows).map_err(PyErr::new::<pyo3::exceptions::PyValueError, _>)?;
    validate_max_columns(max_columns).map_err(PyErr::new::<pyo3::exceptions::PyValueError, _>)?;
    Ok(PyStreamWriter {
        path: path.to_string(),
        rows: Vec::new(),
        pending_bytes: 0,
        max_pending_bytes,
        max_rows,
        max_columns,
        active: true,
    })
}

#[pymethods]
impl PyStreamWriter {
    #[new]
    #[pyo3(signature = (path, max_pending_bytes = None, max_rows = None, max_columns = None))]
    fn new(
        path: &str,
        max_pending_bytes: Option<usize>,
        max_rows: Option<usize>,
        max_columns: Option<usize>,
    ) -> PyResult<Self> {
        stream_writer_from_path(path, max_pending_bytes, max_rows, max_columns)
    }
    fn __enter__(slf: PyRef<'_, Self>) -> PyRef<'_, Self> {
        slf
    }
    fn __exit__(
        &mut self,
        _ty: Option<&Bound<'_, PyAny>>,
        _value: Option<&Bound<'_, PyAny>>,
        _tb: Option<&Bound<'_, PyAny>>,
    ) -> PyResult<bool> {
        self.close()?;
        Ok(false)
    }

    /// Whether this writer has already materialized its rows and closed.
    #[getter]
    fn closed(&self) -> bool {
        !self.active
    }

    /// Number of rows accepted and retained until `close()`.
    #[getter]
    fn row_count(&self) -> usize {
        self.rows.len()
    }

    /// Estimated bytes retained by pending rows.
    #[getter]
    fn pending_bytes(&self) -> usize {
        self.pending_bytes
    }

    /// Maximum estimated pending-row budget for this writer.
    #[getter]
    fn max_pending_bytes(&self) -> usize {
        self.max_pending_bytes
    }

    /// Maximum number of pending rows accepted by this writer.
    #[getter]
    fn max_rows(&self) -> Option<usize> {
        self.max_rows
    }

    /// Maximum number of columns accepted in each pending row.
    #[getter]
    fn max_columns(&self) -> Option<usize> {
        self.max_columns
    }

    fn append(&mut self, values: &Bound<'_, PyAny>) -> PyResult<()> {
        if !self.active {
            return Err(PyErr::new::<pyo3::exceptions::PyValueError, _>(
                "stream writer is closed",
            ));
        }
        let mut row = Vec::new();
        for item in values.try_iter()? {
            row.push(crate::py_to_variant(&item?)?);
        }
        if row.is_empty() {
            return Err(PyErr::new::<pyo3::exceptions::PyValueError, _>(
                "row must not be empty",
            ));
        }
        validate_writer_row_columns(row.len(), self.max_columns)
            .map_err(PyErr::new::<pyo3::exceptions::PyMemoryError, _>)?;
        if self
            .max_rows
            .is_some_and(|max_rows| self.rows.len() >= max_rows)
        {
            let max_rows = self.max_rows.expect("checked above");
            return Err(PyErr::new::<pyo3::exceptions::PyMemoryError, _>(format!(
                "stream writer pending rows exceed the limit of {max_rows} rows"
            )));
        }
        let row_bytes = row.iter().fold(0usize, |total, value| {
            total.saturating_add(estimated_variant_bytes(value))
        });
        if self.pending_bytes.saturating_add(row_bytes) > self.max_pending_bytes {
            return Err(PyErr::new::<pyo3::exceptions::PyMemoryError, _>(format!(
                "stream writer pending rows exceed the limit of {} bytes",
                self.max_pending_bytes
            )));
        }
        self.pending_bytes = self.pending_bytes.saturating_add(row_bytes);
        self.rows.push(row);
        Ok(())
    }
    fn close(&mut self) -> PyResult<()> {
        if !self.active {
            return Ok(());
        }
        let mut vm = crate::vm::Vm::new();
        // Keep the pending rows until the output has been saved successfully.
        // A failed save must be retryable without silently turning the next
        // attempt into an empty workbook.
        for (r, row) in self.rows.iter().enumerate() {
            vm.write_rect("sheet1", ((r + 1) as u32, 1), std::slice::from_ref(row));
        }
        crate::save_workbook(&vm, &self.path)
            .map_err(PyErr::new::<pyo3::exceptions::PyIOError, _>)?;
        self.rows.clear();
        self.pending_bytes = 0;
        self.active = false;
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use std::collections::HashMap;

    use super::{
        MAX_STREAM_ROW_BYTES, MAX_STREAM_WRITER_BYTES, append_row_token, estimated_variant_bytes,
        parse_stream_row, resolve_xlsx_target, row_from_xml, row_from_xml_with_limit,
        stream_writer_from_path, validate_max_columns, validate_max_row_bytes, validate_max_rows,
        validate_max_writer_rows, validate_timeout_ms, validate_writer_row_columns,
    };
    use crate::Variant;

    #[test]
    fn row_parser_preserves_positions_and_shared_strings() {
        let xml = r#"<worksheet><sheetData>
            <row r="7"><c r="B7" t="s"><v>0</v></c><c r="D7"><v>42</v></c></row>
        </sheetData></worksheet>"#;
        let (_, row) = row_from_xml(xml, &["shared".to_string()]).expect("row");
        assert_eq!(row.len(), 4);
        assert!(matches!(row[0], Variant::Empty));
        assert!(matches!(&row[1], Variant::Str(value) if value == "shared"));
        assert!(matches!(row[2], Variant::Empty));
        assert!(matches!(row[3], Variant::Integer(42)));
    }

    #[test]
    fn row_parser_handles_a_row_with_inline_string() {
        let xml = r#"<worksheet><sheetData><row r="2"><c r="A2" t="inlineStr"><is><t>hello</t></is></c></row></sheetData></worksheet>"#;
        let (_, row) = row_from_xml(xml, &[]).expect("row");
        assert!(matches!(&row[0], Variant::Str(value) if value == "hello"));
    }

    #[test]
    fn row_parser_keeps_an_empty_row() {
        let xml = r#"<worksheet><sheetData><row r="9"/></sheetData></worksheet>"#;
        assert_eq!(row_from_xml(xml, &[]), Some((9, vec![])));
    }

    #[test]
    fn row_parser_accepts_whitespace_in_row_close_tag() {
        let xml = r#"<worksheet><sheetData><row r="2"><c r="A2"><v>7</v></c></row ></sheetData></worksheet>"#;
        let (_, row) = row_from_xml(xml, &[]).expect("row");
        assert!(matches!(row.first(), Some(Variant::Integer(7))));
    }

    #[test]
    fn worksheet_relationship_targets_are_normalized() {
        assert_eq!(
            resolve_xlsx_target("worksheets/sheet1.xml").unwrap(),
            "xl/worksheets/sheet1.xml"
        );
        assert_eq!(
            resolve_xlsx_target("../worksheets/sheet1.xml").unwrap(),
            "worksheets/sheet1.xml"
        );
        assert_eq!(
            resolve_xlsx_target("/xl/worksheets/sheet1.xml").unwrap(),
            "xl/worksheets/sheet1.xml"
        );
        assert!(resolve_xlsx_target("../../outside.xml").is_err());
    }

    #[test]
    fn streaming_row_buffer_rejects_unbounded_rows() {
        let mut row = Vec::new();
        assert!(append_row_token(&mut row, &[b'x'; 1024], MAX_STREAM_ROW_BYTES).is_ok());
        let remaining = MAX_STREAM_ROW_BYTES - row.len();
        assert!(append_row_token(&mut row, &vec![b'x'; remaining], MAX_STREAM_ROW_BYTES).is_ok());
        assert!(append_row_token(&mut row, b"x", MAX_STREAM_ROW_BYTES).is_err());
    }

    #[test]
    fn streaming_writer_estimate_accounts_for_nested_values() {
        let value = Variant::Array(vec![
            Variant::Str("hello".to_string()),
            Variant::Integer(42),
        ]);
        assert!(estimated_variant_bytes(&value) >= 5);
        assert!(MAX_STREAM_WRITER_BYTES > estimated_variant_bytes(&value));
    }

    #[test]
    fn streaming_writer_estimate_saturates_nested_totals() {
        let value = Variant::Array(vec![
            Variant::Record(HashMap::from([(
                "key".to_string(),
                Variant::Str("value".to_string()),
            )])),
            Variant::VbaArray(crate::types::VbaArray {
                bounds: vec![crate::types::ArrayBound { lower: 0, upper: 0 }],
                elements: vec![Variant::Str("value".to_string())],
            }),
        ]);
        assert_eq!(estimated_variant_bytes(&value), 3 + 5 + 5);
    }

    #[test]
    fn streaming_writer_lifecycle_is_explicit_and_idempotent() {
        let path = std::env::temp_dir().join("elixcee_stream_writer_lifecycle.xlsx");
        let mut writer =
            stream_writer_from_path(path.to_str().unwrap(), Some(1024), Some(3), Some(4)).unwrap();
        assert!(!writer.closed());
        assert_eq!(writer.row_count(), 0);
        assert_eq!(writer.pending_bytes(), 0);
        assert_eq!(writer.max_pending_bytes(), 1024);
        assert_eq!(writer.max_rows(), Some(3));
        assert_eq!(writer.max_columns(), Some(4));

        writer.close().unwrap();
        assert!(writer.closed());
        writer.close().unwrap();
        let _ = std::fs::remove_file(path);
    }

    #[test]
    fn streaming_writer_preserves_rows_after_a_failed_save() {
        let path = std::env::temp_dir()
            .join(format!(
                "elixcee_stream_writer_missing_parent-{}",
                std::process::id()
            ))
            .join("output.xlsx");
        let mut writer =
            stream_writer_from_path(path.to_str().unwrap(), Some(1024), Some(3), Some(4)).unwrap();
        writer.rows.push(vec![Variant::Integer(42)]);
        writer.pending_bytes = writer.rows[0].iter().fold(0usize, |total, value| {
            total.saturating_add(estimated_variant_bytes(value))
        });

        let _error = writer
            .close()
            .expect_err("missing parent must make save fail");
        assert!(!writer.closed());
        assert_eq!(writer.row_count(), 1);
        assert!(writer.pending_bytes() > 0);
    }

    #[test]
    fn streaming_reader_rejects_a_zero_row_budget_before_opening_the_file() {
        let err = validate_max_rows(Some(0)).expect_err("zero row budget must be rejected");
        assert_eq!(err, "max_rows must be greater than zero");
    }

    #[test]
    fn streaming_writer_rejects_a_zero_row_budget_before_opening_the_file() {
        let err = validate_max_writer_rows(Some(0)).expect_err("zero row budget must be rejected");
        assert_eq!(err, "max_rows must be greater than zero");
    }

    #[test]
    fn streaming_writer_rejects_rows_wider_than_the_column_budget() {
        assert!(validate_writer_row_columns(4, Some(4)).is_ok());
        assert_eq!(
            validate_writer_row_columns(5, Some(4)).unwrap_err(),
            "stream writer row exceeds the limit of 4 columns"
        );
    }

    #[test]
    fn streaming_reader_rejects_a_zero_timeout_before_opening_the_file() {
        let err = validate_timeout_ms(Some(0)).expect_err("zero timeout must be rejected");
        assert_eq!(err, "timeout_ms must be greater than zero");
    }

    #[test]
    fn streaming_reader_rejects_a_zero_row_byte_budget_before_opening_the_file() {
        let err = validate_max_row_bytes(Some(0)).expect_err("zero byte budget must be rejected");
        assert_eq!(err, "max_row_bytes must be greater than zero");
    }

    #[test]
    fn streaming_row_buffer_honors_a_custom_byte_budget() {
        let mut row = Vec::new();
        assert!(append_row_token(&mut row, b"1234", 4).is_ok());
        assert!(append_row_token(&mut row, b"5", 4).is_err());
    }

    #[test]
    fn streaming_reader_rejects_a_zero_column_budget_before_opening_the_file() {
        let err = validate_max_columns(Some(0)).expect_err("zero column budget must be rejected");
        assert_eq!(err, "max_columns must be greater than zero");
    }

    #[test]
    fn streaming_row_parser_honors_a_custom_column_budget() {
        let xml = r#"<worksheet><sheetData><row r="1"><c r="C1"><v>7</v></c></row></sheetData></worksheet>"#;
        assert!(row_from_xml_with_limit(xml, &[], 3).is_ok());
        assert!(row_from_xml_with_limit(xml, &[], 2).is_err());
    }

    #[test]
    fn streaming_row_parser_rejects_invalid_utf8_instead_of_dropping_the_row() {
        let error = parse_stream_row(b"<row r=\"1\">\xff</row>", &[], 4, true)
            .expect_err("invalid UTF-8 must be reported");
        assert_eq!(error, "worksheet row is not valid UTF-8");
    }
}
