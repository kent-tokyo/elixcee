#![no_main]

use libfuzzer_sys::fuzz_target;

fuzz_target!(|data: &[u8]| {
    // Accept arbitrary bytes; parse() must never panic regardless of input.
    if let Ok(s) = std::str::from_utf8(data) {
        let _ = elixcee::parser::parse(s);
    }
});
