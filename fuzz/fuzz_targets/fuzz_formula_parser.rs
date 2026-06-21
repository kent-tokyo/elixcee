#![no_main]

use libfuzzer_sys::fuzz_target;

fuzz_target!(|data: &[u8]| {
    // formula::parse() must never panic on arbitrary input.
    if let Ok(s) = std::str::from_utf8(data) {
        let _ = elixcee::formula::parse(s);
    }
});
