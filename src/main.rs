use std::{env, fs, process};

use elixcee::{
    parser,
    reader::{self, SheetCell},
    save_workbook,
    vm::{serial_to_display, CellContent, Variant, Vm},
};

fn usage() -> ! {
    eprintln!(
        "Usage: elixcee <vba_file> <MacroName> [OPTIONS]\n\
         \n\
         Arguments:\n\
           <vba_file>    Path to VBA source file (.vbs / .bas / .txt)\n\
           <MacroName>   Name of the Sub to execute\n\
         \n\
         Options:\n\
           --file <path>    Load cell data from spreadsheet (.xlsx / .xlsm / .ods)\n\
           --sheet <name>   Active sheet name (default: first sheet in --file)\n\
           --output <path>  Save result cells to spreadsheet (.xlsx / .ods)"
    );
    process::exit(1);
}

fn die(msg: &str) -> ! {
    eprintln!("error: {}", msg);
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

fn main() {
    let args: Vec<String> = env::args().collect();

    let mut vba_file:   Option<String> = None;
    let mut macro_name: Option<String> = None;
    let mut xlsx_file:  Option<String> = None;
    let mut sheet_name: Option<String> = None;
    let mut output:     Option<String> = None;

    let mut i = 1;
    while i < args.len() {
        match args[i].as_str() {
            "--file"   => { i += 1; xlsx_file  = args.get(i).cloned().or_else(|| die("--file requires a path")); }
            "--sheet"  => { i += 1; sheet_name = args.get(i).cloned().or_else(|| die("--sheet requires a name")); }
            "--output" => { i += 1; output     = args.get(i).cloned().or_else(|| die("--output requires a path")); }
            "--help" | "-h" => usage(),
            arg if arg.starts_with('-') => die(&format!("unknown option: {}", arg)),
            _ if vba_file.is_none()   => { vba_file   = Some(args[i].clone()); }
            _ if macro_name.is_none() => { macro_name = Some(args[i].clone()); }
            _ => usage(),
        }
        i += 1;
    }

    let vba_path   = vba_file.unwrap_or_else(|| usage());
    let macro_name = macro_name.unwrap_or_else(|| usage());

    let vba_code = fs::read_to_string(&vba_path)
        .unwrap_or_else(|e| die(&format!("cannot read '{}': {}", vba_path, e)));

    let prog = parser::parse(&vba_code)
        .unwrap_or_else(|e| die(&format!("parse error: {}", e)));

    let mut vm = Vm::new();
    vm.print_msgbox = true;

    // Load spreadsheet data if provided
    if let Some(ref path) = xlsx_file {
        let sheets = reader::read_workbook(path)
            .unwrap_or_else(|e| die(&format!("cannot read '{}': {}", path, e)));
        if sheets.is_empty() {
            die("workbook has no sheets");
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
        vm.set_active_sheet(&active)
            .unwrap_or_else(|e| die(&e));
    } else if let Some(ref name) = sheet_name {
        vm.set_active_sheet(name)
            .unwrap_or_else(|e| die(&e));
    }

    vm.run_sub(&prog, &macro_name)
        .unwrap_or_else(|e| die(&format!("runtime error: {}", e)));

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
