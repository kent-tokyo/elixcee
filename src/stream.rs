//! Forward-only row streaming for large XLSX files (0.25.0).
//!
//! The normal `Vm` intentionally remains a random-access, fully mutable model. This
//! module provides a separate pipeline API whose worker owns the ZIP entry and sends
//! one decoded row at a time, so callers do not need to materialize a workbook.

use std::fs::File;
use std::io::{BufRead, BufReader};
use std::sync::{
    Mutex,
    mpsc::{self, Receiver},
};

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
        Variant::Array(values) => values.iter().map(estimated_variant_bytes).sum(),
        Variant::VbaArray(values) => values.elements.iter().map(estimated_variant_bytes).sum(),
        Variant::Record(values) => values
            .iter()
            .map(|(key, value)| key.len() + estimated_variant_bytes(value))
            .sum(),
        _ => std::mem::size_of_val(value),
    }
}

fn append_row_token(row_buf: &mut Vec<u8>, token: &[u8]) -> Result<(), String> {
    if row_buf.len().saturating_add(token.len()) > MAX_STREAM_ROW_BYTES {
        return Err(format!(
            "worksheet row exceeds the streaming limit of {MAX_STREAM_ROW_BYTES} bytes"
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

fn row_from_xml(xml: &str, shared: &[String]) -> Option<(u32, Vec<Variant>)> {
    let parsed = reader::xlsx_sheet_cells_for_stream(xml, shared);
    let row = parsed.first_row?;
    let max_col = parsed.cells.keys().map(|(_, col)| *col).max().unwrap_or(0);
    let mut out = vec![Variant::Empty; max_col as usize];
    for ((_, col), cell) in parsed.cells {
        if col > 0 {
            out[(col - 1) as usize] = cell_variant(&cell);
        }
    }
    Some((row, out))
}

fn is_row_close(token: &[u8]) -> bool {
    let token = token.trim_ascii_start();
    let Some(rest) = token.strip_prefix(b"</row") else {
        return false;
    };
    rest.iter().all(|b| b.is_ascii_whitespace() || *b == b'>') && rest.contains(&b'>')
}

fn stream_rows(
    path: String,
    sheet: Option<String>,
) -> Result<Receiver<Result<(u32, Vec<Variant>), String>>, String> {
    let (zip_path, _) = sheet_target(&path, sheet.as_deref())?;
    let (tx, rx) = mpsc::sync_channel(2);
    std::thread::spawn(move || {
        let result = (|| -> Result<(), String> {
            let file = File::open(&path).map_err(|e| e.to_string())?;
            let mut archive = zip::ZipArchive::new(file).map_err(|e| e.to_string())?;
            reader::validate_zip_archive_for_stream(&mut archive)?;
            let shared_xml = reader::zip_read_text_for_stream(&mut archive, "xl/sharedStrings.xml")
                .unwrap_or_default();
            let shared = reader::xlsx_shared_strings_for_stream(&shared_xml);
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
                        append_row_token(&mut row_buf, &token)?;
                        if trimmed.ends_with(b"/>") {
                            if let Ok(xml) = std::str::from_utf8(&row_buf)
                                && let Some((row_number, row)) = row_from_xml(xml, &shared)
                                && tx.send(Ok((row_number, row))).is_err()
                            {
                                return Ok(());
                            }
                            in_row = false;
                        }
                    }
                } else {
                    append_row_token(&mut row_buf, &token)?;
                    if is_row_close(&token) {
                        if let Ok(xml) = std::str::from_utf8(&row_buf) {
                            let wrapped =
                                format!("<worksheet><sheetData>{xml}</sheetData></worksheet>");
                            if let Some((row_number, row)) = row_from_xml(&wrapped, &shared)
                                && tx.send(Ok((row_number, row))).is_err()
                            {
                                return Ok(());
                            }
                        }
                        in_row = false;
                    }
                }
            }
            Ok(())
        })();
        if let Err(err) = result {
            let _ = tx.send(Err(err));
        }
    });
    Ok(rx)
}

#[pyclass(name = "StreamReader")]
pub struct PyStreamReader {
    receiver: Option<RowReceiver>,
    include_row_numbers: bool,
}

type RowReceiver = Mutex<Receiver<Result<(u32, Vec<Variant>), String>>>;

pub(crate) fn stream_reader_from_path(
    path: &str,
    sheet: Option<&str>,
    include_row_numbers: bool,
) -> PyResult<PyStreamReader> {
    let receiver = stream_rows(path.to_string(), sheet.map(str::to_string))
        .map_err(PyErr::new::<pyo3::exceptions::PyIOError, _>)?;
    Ok(PyStreamReader {
        receiver: Some(Mutex::new(receiver)),
        include_row_numbers,
    })
}

#[pymethods]
impl PyStreamReader {
    #[new]
    #[pyo3(signature = (path, sheet = None, include_row_numbers = false))]
    fn new(path: &str, sheet: Option<&str>, include_row_numbers: bool) -> PyResult<Self> {
        stream_reader_from_path(path, sheet, include_row_numbers)
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
        self.receiver = None;
        false
    }
    fn __iter__(slf: PyRef<'_, Self>) -> PyRef<'_, Self> {
        slf
    }
    fn __next__(&mut self, py: Python<'_>) -> PyResult<Option<Py<PyAny>>> {
        let Some(receiver) = &self.receiver else {
            return Ok(None);
        };
        let next = receiver
            .lock()
            .expect("stream reader mutex poisoned")
            .recv();
        match next {
            Ok(Ok((row_number, row))) => {
                let values = PyList::new(py, row.iter().map(|v| variant_to_py(py, v)))?
                    .into_any()
                    .unbind();
                if self.include_row_numbers {
                    Ok(Some(
                        (row_number, values).into_pyobject(py)?.into_any().unbind(),
                    ))
                } else {
                    Ok(Some(values))
                }
            }
            Ok(Err(err)) => Err(PyErr::new::<pyo3::exceptions::PyIOError, _>(err)),
            Err(_) => {
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
    active: bool,
}

pub(crate) fn stream_writer_from_path(
    path: &str,
    max_pending_bytes: Option<usize>,
) -> PyResult<PyStreamWriter> {
    let max_pending_bytes = max_pending_bytes.unwrap_or(MAX_STREAM_WRITER_BYTES);
    if max_pending_bytes == 0 {
        return Err(PyErr::new::<pyo3::exceptions::PyValueError, _>(
            "max_pending_bytes must be greater than zero",
        ));
    }
    Ok(PyStreamWriter {
        path: path.to_string(),
        rows: Vec::new(),
        pending_bytes: 0,
        max_pending_bytes,
        active: true,
    })
}

#[pymethods]
impl PyStreamWriter {
    #[new]
    #[pyo3(signature = (path, max_pending_bytes = None))]
    fn new(path: &str, max_pending_bytes: Option<usize>) -> PyResult<Self> {
        stream_writer_from_path(path, max_pending_bytes)
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
        let row_bytes = row.iter().map(estimated_variant_bytes).sum::<usize>();
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
        self.pending_bytes = 0;
        for (r, row) in self.rows.drain(..).enumerate() {
            vm.write_rect("sheet1", ((r + 1) as u32, 1), &[row]);
        }
        crate::save_workbook(&vm, &self.path)
            .map_err(PyErr::new::<pyo3::exceptions::PyIOError, _>)?;
        self.active = false;
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::{
        MAX_STREAM_ROW_BYTES, MAX_STREAM_WRITER_BYTES, append_row_token, estimated_variant_bytes,
        resolve_xlsx_target, row_from_xml,
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
        assert!(append_row_token(&mut row, &[b'x'; 1024]).is_ok());
        let remaining = MAX_STREAM_ROW_BYTES - row.len();
        assert!(append_row_token(&mut row, &vec![b'x'; remaining]).is_ok());
        assert!(append_row_token(&mut row, b"x").is_err());
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
}
