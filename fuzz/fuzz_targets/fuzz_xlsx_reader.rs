#![no_main]

use libfuzzer_sys::fuzz_target;
use std::io::Write;

fuzz_target!(|data: &[u8]| {
    // Write the fuzz bytes to a temp file and try to parse it as XLSX/ODS.
    // The reader must never panic — only return Ok or Err.
    let mut tmp = match tempfile::NamedTempFile::new() {
        Ok(f) => f,
        Err(_) => return,
    };
    if tmp.write_all(data).is_err() { return; }
    let path = match tmp.path().to_str() {
        Some(p) => p.to_owned(),
        None => return,
    };
    // Try as .xlsx — the reader dispatches by file extension,
    // so use a suffixed copy.
    let xlsx_path = format!("{}.xlsx", path);
    if std::fs::copy(tmp.path(), &xlsx_path).is_ok() {
        let _ = elixcee::reader::read_workbook(&xlsx_path);
        let _ = std::fs::remove_file(&xlsx_path);
    }
});
