use std::{env, fs, process};

use elixcee::{
    check,
    diagnostics::{self, ElixceeError},
    parser,
    reader::{self, SheetCell},
    save_workbook, snapshot,
    vm::{serial_to_display, CellContent, Variant, Vm},
};

fn usage() -> ! {
    eprintln!(
        "Usage: elixcee <vba_file>... <MacroName> [OPTIONS]\n\
         \n\
         Arguments:\n\
           <vba_file>...  One or more VBA source files (.vbs / .bas / .txt).\n\
         \x20              With more than one, Sub/Function names are shared\n\
         \x20              project-wide; use Module.Sub to disambiguate.\n\
           <MacroName>    Name of the Sub to execute (last argument; may be\n\
         \x20              bare or Module.Sub-qualified)\n\
         \n\
         Options:\n\
           --file <path>    Load cell data from spreadsheet (.xlsx / .xlsm / .ods)\n\
           --sheet <name>   Active sheet name (default: first sheet in --file)\n\
           --output <path>  Save result cells to spreadsheet (.xlsx / .ods)\n\
           --json           Emit a single JSON object (result or error) instead of plain text\n\
         \n\
         Subcommands:\n\
           elixcee check <vba_file>... [--entry <MacroName>] [--json]\n\
         \x20   Static analysis — parse + optional entrypoint check + interactive-call\n\
         \x20   detection, without executing the macro. All positional arguments\n\
         \x20   are files; the entrypoint (if any) is always given via --entry.\n\
           elixcee snapshot <file> [--json]\n\
         \x20   Reads a .xlsx/.ods file directly (no VBA execution) and prints every\n\
         \x20   sheet's non-empty cells — Markdown by default, JSON with --json."
    );
    process::exit(1);
}

/// A parsed VBA module ready for multi-file resolution: its derived name
/// (from `Attribute VB_Name` if present, else the file stem — lowercased),
/// its source path and text (needed for diagnostic/error locations), and
/// the parsed `Program` itself.
struct LoadedModule {
    name: String,
    path: String,
    source: String,
    program: parser::Program,
}

enum LoadModuleError {
    Io(String),
    Parse { message: String, span: parser::SourceSpan, source: String },
}

fn derive_module_name(path: &str) -> String {
    std::path::Path::new(path)
        .file_stem()
        .and_then(|s| s.to_str())
        .unwrap_or(path)
        .to_string()
}

/// Read and parse one `.bas` file into a `LoadedModule`. Shared by both
/// run-mode (which exits immediately on failure — execution can't proceed
/// with a bad module) and `check` mode (which collects a diagnostic and
/// keeps checking the other modules instead).
fn load_one_module(path: &str) -> Result<LoadedModule, LoadModuleError> {
    let code = fs::read_to_string(path)
        .map_err(|e| LoadModuleError::Io(format!("cannot read '{}': {}", path, e)))?;
    let program = parser::parse_with_span(&code)
        .map_err(|e| LoadModuleError::Parse { message: e.message, span: e.span, source: code.clone() })?;
    let name = program.module_name.clone()
        .unwrap_or_else(|| derive_module_name(path))
        .to_lowercase();
    Ok(LoadedModule { name, path: path.to_string(), source: code, program })
}

/// Run-mode's module loader: reads and parses every file in `paths`,
/// exiting the process on the first read/parse failure or on a duplicate
/// module name (execution can't proceed with an incomplete or ambiguous
/// project). Mirrors the single-file io_error/parse_error paths exactly
/// when `paths.len() == 1`.
fn load_modules(paths: &[String], json: bool) -> Vec<LoadedModule> {
    let mut modules: Vec<LoadedModule> = Vec::new();
    for path in paths {
        let module = match load_one_module(path) {
            Ok(m) => m,
            Err(LoadModuleError::Io(msg)) => {
                if json { fail_json(ElixceeError::io_error(msg), &[]) } else { die(&msg) }
            }
            Err(LoadModuleError::Parse { message, span, source }) => {
                let location = diagnostics::locate(&source, path, span);
                if json {
                    fail_json(ElixceeError::parse_error(message).with_location(Some(location)), &[])
                } else {
                    die(&format!("parse error in '{}': {}", path, message))
                }
            }
        };
        if modules.iter().any(|m: &LoadedModule| m.name == module.name) {
            let msg = format!(
                "duplicate module name '{}' (from '{}') — every module in a project needs a unique name",
                module.name, path,
            );
            if json { fail_json(ElixceeError::io_error(msg), &[]) } else { die(&msg) }
        }
        modules.push(module);
    }
    modules
}

/// Read and parse each file for `check` mode. Unlike run-mode's loader
/// (which exits on the first read/parse failure, since execution can't
/// proceed at all), a read/parse failure in one module doesn't stop
/// `check` from reporting diagnostics for the others — check's job is a
/// batch of findings, not an all-or-nothing gate. Returns the successfully
/// loaded modules plus any io/parse-error diagnostics collected so far.
fn check_load_modules(paths: &[String]) -> (Vec<LoadedModule>, Vec<check::Diagnostic>) {
    let mut modules = Vec::new();
    let mut diags = Vec::new();
    for path in paths {
        match load_one_module(path) {
            Ok(module) => {
                if modules.iter().any(|m: &LoadedModule| m.name == module.name) {
                    diags.push(check::Diagnostic {
                        severity: "error",
                        code: "E1006",
                        kind: "duplicate_module_name",
                        message: format!(
                            "duplicate module name '{}' (from '{}') — every module in a project needs a unique name",
                            module.name, path,
                        ),
                        location: None,
                    });
                }
                modules.push(module);
            }
            Err(LoadModuleError::Io(msg)) => diags.push(check::Diagnostic {
                severity: "error",
                code: "E3001",
                kind: "io_error",
                message: msg,
                location: None,
            }),
            Err(LoadModuleError::Parse { message, span, source }) => diags.push(check::Diagnostic {
                severity: "error",
                code: "E2001",
                kind: "parse_error",
                message,
                location: Some(diagnostics::locate(&source, path, span)),
            }),
        }
    }
    (modules, diags)
}

/// `elixcee check <vba_file>... [--entry <MacroName>] [--json]` — static
/// analysis, no execution. Kept as a separate small entry point rather than
/// folding into the run-mode arg loop above: it has its own (looser)
/// argument shape (every positional is a file; the entrypoint, if any, is
/// always `--entry`, never positional — unlike run-mode, `check`'s
/// entrypoint is optional, so a positional macro name would be ambiguous
/// against a project with 2+ files and no desired entrypoint check) and its
/// own output shape (a diagnostics list, not a single result/error).
fn run_check_command(args: &[String]) -> ! {
    let mut vba_files: Vec<String> = Vec::new();
    let mut macro_name: Option<String> = None;
    let mut json = false;

    let mut i = 0;
    while i < args.len() {
        match args[i].as_str() {
            "--json" => json = true,
            "--entry" => { i += 1; macro_name = args.get(i).cloned().or_else(|| die("--entry requires a name")); }
            a if a.starts_with('-') => die(&format!("unknown option: {}", a)),
            _ => vba_files.push(args[i].clone()),
        }
        i += 1;
    }
    if vba_files.is_empty() { usage(); }

    let (modules, mut diags) = check_load_modules(&vba_files);

    if modules.len() > 1 {
        let project: Vec<(String, parser::Program)> =
            modules.iter().map(|m| (m.name.clone(), m.program.clone())).collect();

        for (name, mods) in parser::find_cross_module_sub_collisions(&project) {
            diags.push(check::Diagnostic {
                severity: "error",
                code: "E1005",
                kind: "duplicate_sub_or_function",
                message: format!(
                    "duplicate Sub '{}' across modules '{}'",
                    name,
                    mods.join("', '")
                ),
                location: None,
            });
        }
        for (name, mods) in parser::find_cross_module_func_collisions(&project) {
            diags.push(check::Diagnostic {
                severity: "error",
                code: "E1005",
                kind: "duplicate_sub_or_function",
                message: format!(
                    "duplicate Function '{}' across modules '{}'",
                    name,
                    mods.join("', '")
                ),
                location: None,
            });
        }

        for m in &modules {
            let mut others: std::collections::HashSet<String> = std::collections::HashSet::new();
            for other in &modules {
                if other.name != m.name {
                    others.extend(other.program.subs.iter().map(|s| s.name.clone()));
                    others.extend(other.program.funcs.iter().map(|f| f.name.clone()));
                }
            }
            diags.extend(check::run_check_in_project(
                &m.source, &m.path, None, &others,
            ));
        }

        if let Some(ref name) = macro_name
            && matches!(
                parser::resolve_entrypoint(&project, name),
                parser::EntrypointResolution::NotFound
            )
        {
            diags.push(check::Diagnostic {
                severity: "error",
                code: "E1002",
                kind: "undefined_sub_or_function",
                message: format!("Sub '{}' not found", name),
                location: None,
            });
        }
    } else if let Some(m) = modules.first() {
        diags.extend(check::run_check(&m.source, &m.path, macro_name.as_deref()));
    }

    let ok = check::all_ok(&diags);

    if json {
        println!("{}", check::diagnostics_to_json(&diags));
    } else if diags.is_empty() {
        println!("ok");
    } else {
        for d in &diags {
            let loc = d
                .location
                .as_ref()
                .map(|l| format!(" ({}:{}:{})", l.file, l.line, l.column))
                .unwrap_or_default();
            println!("{} {} {}: {}{}", d.severity, d.code, d.kind, d.message, loc);
        }
    }
    process::exit(if ok { 0 } else { 1 });
}

/// `elixcee snapshot <file> [--json]` — reads a `.xlsx`/`.ods` workbook file
/// directly (no VBA execution, same "inspect, don't execute" posture as
/// `check`) and renders every sheet's non-empty cells as JSON (authoritative)
/// or Markdown (default, for display). See `elixcee::snapshot` for the
/// output-shape details and the `stable_id`/`sheet_id` design rationale.
fn run_snapshot_command(args: &[String]) -> ! {
    let mut path: Option<String> = None;
    let mut json = false;

    for arg in args {
        match arg.as_str() {
            "--json" => json = true,
            a if a.starts_with('-') => die(&format!("unknown option: {}", a)),
            _ if path.is_none() => path = Some(arg.clone()),
            _ => die("snapshot takes exactly one file"),
        }
    }
    let Some(path) = path else { usage() };

    match reader::read_workbook(&path) {
        Ok(sheets) => {
            if json {
                println!("{}", snapshot::to_json(&path, &sheets));
            } else {
                println!("{}", snapshot::to_markdown(&path, &sheets));
            }
            process::exit(0);
        }
        Err(msg) => {
            if json {
                fail_json(ElixceeError::io_error(msg), &[]);
            } else {
                die(&msg);
            }
        }
    }
}

fn die(msg: &str) -> ! {
    eprintln!("error: {}", msg);
    process::exit(1);
}

/// Print the `--json` error object to stdout and exit(1). Kept separate from
/// `die()` (which writes to stderr) so a `--json` run always emits exactly
/// one JSON object on stdout, success or failure. `messages` should be
/// `vm.take_messages()` for failures that happen after the macro started
/// running (so any MsgBox shown before the failure isn't silently lost),
/// and `&[]` for failures that happen before that (nothing could have
/// fired yet).
fn fail_json(err: ElixceeError, messages: &[String]) -> ! {
    println!("{}", err.to_json(messages));
    process::exit(1);
}

fn col_to_letters(mut col: u32) -> String {
    let mut bytes = Vec::new();
    while col > 0 {
        col -= 1;
        bytes.push(b'A' + (col % 26) as u8);
        col /= 26;
    }
    bytes.reverse();
    String::from_utf8(bytes).unwrap()
}

fn format_variant(v: &Variant) -> String {
    match v {
        Variant::Integer(n) => n.to_string(),
        Variant::Float(f)   => f.to_string(),
        Variant::Str(s)     => s.clone(),
        Variant::Boolean(b) => if *b { "TRUE".into() } else { "FALSE".into() },
        Variant::Date(s)    => serial_to_display(*s),
        Variant::Error(e)   => e.as_str().to_string(),
        Variant::Empty      => String::new(),
        Variant::Array(_)   => "[array]".into(),
        Variant::Record(_)  => "[record]".into(),
    }
}

/// Build the `"cells"` JSON array from the active sheet's non-empty cells —
/// same selection the plain-text TSV output uses.
fn cells_to_json(vm: &Vm) -> String {
    let mut cells: Vec<_> = vm.cells().iter()
        .filter(|(_, c)| !matches!(c.value, Variant::Empty))
        .collect();
    cells.sort_by_key(|&(&(r, c), _)| (r, c));

    let sheet = diagnostics::json_string(&vm.active_sheet);
    let items: Vec<String> = cells.iter().map(|&(&(row, col), content)| {
        let address = diagnostics::json_string(&format!("{}{}", col_to_letters(col), row));
        format!(
            "{{\"sheet\":{},\"address\":{},\"value\":{}}}",
            sheet, address, diagnostics::variant_to_json(&content.value),
        )
    }).collect();
    format!("[{}]", items.join(","))
}

fn messages_to_json(messages: &[String]) -> String {
    let items: Vec<String> = messages.iter().map(|m| diagnostics::json_string(m)).collect();
    format!("[{}]", items.join(","))
}

fn main() {
    let args: Vec<String> = env::args().collect();

    if args.get(1).map(String::as_str) == Some("check") {
        run_check_command(&args[2..]);
    }
    if args.get(1).map(String::as_str) == Some("snapshot") {
        run_snapshot_command(&args[2..]);
    }

    let mut positionals: Vec<String> = Vec::new();
    let mut xlsx_file:  Option<String> = None;
    let mut sheet_name: Option<String> = None;
    let mut output:     Option<String> = None;
    let mut json = false;

    let mut i = 1;
    while i < args.len() {
        match args[i].as_str() {
            "--file"   => { i += 1; xlsx_file  = args.get(i).cloned().or_else(|| die("--file requires a path")); }
            "--sheet"  => { i += 1; sheet_name = args.get(i).cloned().or_else(|| die("--sheet requires a name")); }
            "--output" => { i += 1; output     = args.get(i).cloned().or_else(|| die("--output requires a path")); }
            "--json"   => { json = true; }
            "--help" | "-h" => usage(),
            arg if arg.starts_with('-') => die(&format!("unknown option: {}", arg)),
            _ => positionals.push(args[i].clone()),
        }
        i += 1;
    }

    // Macro name is mandatory in run mode (unlike `check`), so it's always
    // unambiguous: the last positional is the entrypoint, everything before
    // it is a source file. A single file + single macro name — today's only
    // shape until now — parses identically to before.
    if positionals.is_empty() { usage(); }
    let macro_name = positionals.pop().unwrap();
    if positionals.is_empty() { usage(); }
    let vba_paths = positionals;

    let modules = load_modules(&vba_paths, json);

    let mut vm = Vm::new();
    vm.print_msgbox = !json;

    // Load spreadsheet data if provided
    if let Some(ref path) = xlsx_file {
        let sheets = match reader::read_workbook(path) {
            Ok(s) => s,
            Err(e) => {
                let msg = format!("cannot read '{}': {}", path, e);
                if json { fail_json(ElixceeError::io_error(msg), &[]) } else { die(&msg) }
            }
        };
        if sheets.is_empty() {
            if json { fail_json(ElixceeError::sheet_setup_error("workbook has no sheets".into()), &[]) }
            else { die("workbook has no sheets") }
        }

        for sheet_data in &sheets {
            vm.ensure_sheet(&sheet_data.name);
            let prev = vm.active_sheet.clone();
            vm.active_sheet = sheet_data.name.clone();
            for (&(row, col), cell) in &sheet_data.cells {
                let value = match cell {
                    SheetCell::Integer(n) => Variant::Integer(*n),
                    SheetCell::Float(f)   => Variant::Float(*f),
                    SheetCell::Str(s)     => Variant::Str(s.clone()),
                    SheetCell::Bool(b)    => Variant::Boolean(*b),
                };
                vm.cells_mut().insert((row, col), CellContent { formula: None, value });
            }
            vm.active_sheet = prev;
        }

        let active = sheet_name.as_deref().unwrap_or(&sheets[0].name).to_string();
        if let Err(e) = vm.set_active_sheet(&active) {
            if json { fail_json(ElixceeError::sheet_setup_error(e), &[]) } else { die(&e) }
        }
    } else if let Some(ref name) = sheet_name {
        if let Err(e) = vm.set_active_sheet(name) {
            if json { fail_json(ElixceeError::sheet_setup_error(e), &[]) } else { die(&e) }
        }
    }

    let start = std::time::Instant::now();
    let run_result = if modules.len() == 1 {
        vm.run_sub(&modules[0].program, &macro_name)
    } else {
        let project: Vec<(String, parser::Program)> =
            modules.iter().map(|m| (m.name.clone(), m.program.clone())).collect();
        vm.run_sub_multi(&project, &macro_name)
    };
    let duration_ms = start.elapsed().as_secs_f64() * 1000.0;
    if let Err(e) = run_result {
        // MsgBox text shown before the failure must still reach the agent —
        // take_messages() before fail_json, not after (fail_json never returns).
        if json {
            // A runtime-error span is a char offset into whichever module's
            // source was executing — but `SourceSpan` carries no module id
            // (deliberately deferred back in Milestone A.5, when there was
            // always exactly one source per run). Reusing it against the
            // wrong module's source would print a confidently wrong
            // location, so multi-module runs report `location: None`
            // instead — single-module runs are completely unaffected and
            // keep today's exact precise location.
            let location = if modules.len() == 1 {
                vm.current_span().map(|span| diagnostics::locate(&modules[0].source, &modules[0].path, span))
            } else {
                None
            };
            fail_json(ElixceeError::runtime_error(e).with_location(location), &vm.take_messages())
        } else {
            die(&format!("runtime error: {}", e))
        }
    }

    if json {
        // Do the (optional) save first so a write failure doesn't leave a
        // success object already printed — --json must emit exactly one
        // JSON object on stdout.
        if let Some(ref path) = output {
            if let Err(e) = save_workbook(&vm, path) {
                // The macro already ran successfully — don't drop any MsgBox
                // text it showed just because the save step failed after.
                fail_json(ElixceeError::io_error(format!("cannot write '{}': {}", path, e)), &vm.take_messages());
            }
        }
        let messages = vm.take_messages();
        println!(
            "{{\"schema_version\":1,\"ok\":true,\"entrypoint\":{},\"duration_ms\":{:.3},\"cells\":{},\"messages\":{}}}",
            diagnostics::json_string(&macro_name), duration_ms, cells_to_json(&vm), messages_to_json(&messages),
        );
        return;
    }

    // Print non-empty cells sorted by (row, col)
    let mut cells: Vec<_> = vm.cells().iter()
        .filter(|(_, c)| !matches!(c.value, Variant::Empty))
        .collect();
    cells.sort_by_key(|&(&(r, c), _)| (r, c));

    for &(&(row, col), content) in &cells {
        println!("{}{}\t{}", col_to_letters(col), row, format_variant(&content.value));
    }

    // Save output file if requested
    if let Some(ref path) = output {
        save_workbook(&vm, path)
            .unwrap_or_else(|e| die(&format!("cannot write '{}': {}", path, e)));
    }
}
