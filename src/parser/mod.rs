pub mod ast;
pub use ast::*;

// ── Tokenizer ─────────────────────────────────────────────────────────────────

#[derive(Debug, Clone, PartialEq)]
enum Tok {
    Int(i64),
    Float(f64),
    Str(String),
    Ident(String), // always lowercase
    // Comparison operators
    Eq, Ne, Lt, Le, Gt, Ge,
    // Arithmetic / string
    Plus, Minus, Star, Slash, Amp,
    // Punctuation
    LParen, RParen, Comma, Dot, ColonEq, Colon,
    // End of line
    Newline,
    Eof,
}

fn tokenize(input: &str) -> (Vec<Tok>, Vec<(u32, u32)>) {
    let chars: Vec<char> = input.chars().collect();
    let mut pos = 0;
    let mut toks: Vec<Tok> = Vec::new();
    // Parallel (start, end) char-offset span per token in `toks`.
    let mut spans: Vec<(u32, u32)> = Vec::new();

    macro_rules! ch { () => { chars[pos] }; }

    while pos < chars.len() {
        let tok_start = pos;
        match chars[pos] {
            ' ' | '\t' => { pos += 1; }
            '\'' => {
                while pos < chars.len() && chars[pos] != '\n' && chars[pos] != '\r' { pos += 1; }
            }
            '\r' => {
                pos += 1;
                if pos < chars.len() && chars[pos] == '\n' { pos += 1; }
                push_nl(&mut toks);
            }
            '\n' => { pos += 1; push_nl(&mut toks); }
            '"' => {
                pos += 1;
                let mut s = String::new();
                loop {
                    if pos >= chars.len() { break; }
                    if chars[pos] == '"' {
                        pos += 1;
                        if pos < chars.len() && chars[pos] == '"' { s.push('"'); pos += 1; }
                        else { break; }
                    } else {
                        s.push(chars[pos]); pos += 1;
                    }
                }
                toks.push(Tok::Str(s));
            }
            '<' => {
                pos += 1;
                if pos < chars.len() && ch!() == '>' { pos += 1; toks.push(Tok::Ne); }
                else if pos < chars.len() && ch!() == '=' { pos += 1; toks.push(Tok::Le); }
                else { toks.push(Tok::Lt); }
            }
            '>' => {
                pos += 1;
                if pos < chars.len() && ch!() == '=' { pos += 1; toks.push(Tok::Ge); }
                else { toks.push(Tok::Gt); }
            }
            '=' => { pos += 1; toks.push(Tok::Eq); }
            '+' => { pos += 1; toks.push(Tok::Plus); }
            '-' => { pos += 1; toks.push(Tok::Minus); }
            '*' => { pos += 1; toks.push(Tok::Star); }
            '/' => { pos += 1; toks.push(Tok::Slash); }
            '&' => { pos += 1; toks.push(Tok::Amp); }
            '(' => { pos += 1; toks.push(Tok::LParen); }
            ')' => { pos += 1; toks.push(Tok::RParen); }
            ',' => { pos += 1; toks.push(Tok::Comma); }
            '.' => { pos += 1; toks.push(Tok::Dot); }
            ':' => {
                pos += 1;
                if pos < chars.len() && ch!() == '=' { pos += 1; toks.push(Tok::ColonEq); }
                else { toks.push(Tok::Colon); }
            }
            '_' => {
                // Line continuation: _ at end of line
                pos += 1;
                while pos < chars.len() && (chars[pos] == ' ' || chars[pos] == '\t') { pos += 1; }
                if pos < chars.len() && (chars[pos] == '\n' || chars[pos] == '\r') {
                    if chars[pos] == '\r' { pos += 1; }
                    if pos < chars.len() && chars[pos] == '\n' { pos += 1; }
                    // continuation: don't emit Newline, keep parsing next line
                }
            }
            c if c.is_ascii_digit() => {
                let start = pos;
                while pos < chars.len() && chars[pos].is_ascii_digit() { pos += 1; }
                if pos < chars.len() && chars[pos] == '.'
                    && pos + 1 < chars.len() && chars[pos + 1].is_ascii_digit()
                {
                    pos += 1;
                    while pos < chars.len() && chars[pos].is_ascii_digit() { pos += 1; }
                    let s: String = chars[start..pos].iter().collect();
                    toks.push(Tok::Float(s.parse().unwrap()));
                } else {
                    let s: String = chars[start..pos].iter().collect();
                    toks.push(Tok::Int(s.parse().unwrap()));
                }
            }
            c if c.is_ascii_alphabetic() => {
                let start = pos;
                while pos < chars.len() && (chars[pos].is_ascii_alphanumeric() || chars[pos] == '_') {
                    pos += 1;
                }
                let s: String = chars[start..pos].iter().collect::<String>().to_lowercase();
                toks.push(Tok::Ident(s));
            }
            _ => { pos += 1; }
        }
        // The match arm above pushed 0 or 1 tokens (0 for whitespace/comments/
        // line continuations) — record the same (tok_start, pos) span for
        // however many it actually pushed, without touching any arm above.
        while spans.len() < toks.len() {
            spans.push((tok_start as u32, pos as u32));
        }
    }
    toks.push(Tok::Eof);
    spans.push((pos as u32, pos as u32));
    (toks, spans)
}

// Only push Newline if last token isn't already one (collapse runs)
fn push_nl(toks: &mut Vec<Tok>) {
    if !matches!(toks.last(), Some(Tok::Newline)) {
        toks.push(Tok::Newline);
    }
}

// ── Parser ────────────────────────────────────────────────────────────────────

struct Parser {
    tokens: Vec<Tok>,
    /// Parallel to `tokens`: (start, end) char-offset span of each token.
    spans: Vec<(u32, u32)>,
    pos: usize,
    /// Variable name currently targeted by `With p` (None outside a With block).
    with_target: Option<String>,
}

impl Parser {
    fn new(tokens: Vec<Tok>, spans: Vec<(u32, u32)>) -> Self {
        Parser { tokens, spans, pos: 0, with_target: None }
    }

    fn peek(&self) -> &Tok {
        self.tokens.get(self.pos).unwrap_or(&Tok::Eof)
    }

    /// Span of the token at the current position (clamped to the last
    /// recorded span — the EOF sentinel — if past the end).
    fn peek_span(&self) -> SourceSpan {
        let &(start, end) = self.spans.get(self.pos)
            .unwrap_or_else(|| self.spans.last().expect("tokenize always emits at least an EOF span"));
        SourceSpan { start, end }
    }

    fn peek_at(&self, offset: usize) -> &Tok {
        self.tokens.get(self.pos + offset).unwrap_or(&Tok::Eof)
    }

    fn advance(&mut self) -> Tok {
        let t = self.tokens.get(self.pos).cloned().unwrap_or(Tok::Eof);
        if self.pos < self.tokens.len() { self.pos += 1; }
        t
    }

    fn is_ident(&self, name: &str) -> bool {
        matches!(self.peek(), Tok::Ident(s) if s == name)
    }

    fn is_ident_at(&self, offset: usize, name: &str) -> bool {
        matches!(self.peek_at(offset), Tok::Ident(s) if s == name)
    }

    fn expect_ident(&mut self, name: &str) -> Result<(), String> {
        match self.peek() {
            Tok::Ident(s) if s == name => { self.advance(); Ok(()) }
            t => Err(format!("expected '{}', got {:?}", name, t)),
        }
    }

    fn expect_tok(&mut self, expected: Tok) -> Result<(), String> {
        if *self.peek() == expected {
            self.advance();
            Ok(())
        } else {
            Err(format!("expected {:?}, got {:?}", expected, self.peek()))
        }
    }

    fn consume_ident(&mut self) -> Result<String, String> {
        match self.advance() {
            Tok::Ident(s) => Ok(s),
            t => Err(format!("expected identifier, got {:?}", t)),
        }
    }

    fn consume_str(&mut self) -> Result<String, String> {
        match self.advance() {
            Tok::Str(s) => Ok(s),
            t => Err(format!("expected string literal, got {:?}", t)),
        }
    }

    fn skip_nl(&mut self) {
        while *self.peek() == Tok::Newline { self.advance(); }
    }

    fn eat_eol(&mut self) -> Result<(), String> {
        match self.peek() {
            Tok::Newline => { self.advance(); Ok(()) }
            Tok::Eof     => Ok(()),
            t => Err(format!("expected newline, got {:?}", t)),
        }
    }

    // Consume to end of line (inclusive of the newline token)
    fn skip_to_eol(&mut self) {
        while !matches!(self.peek(), Tok::Newline | Tok::Eof) { self.advance(); }
        if *self.peek() == Tok::Newline { self.advance(); }
    }

    fn is_end_kw(&self, kw: &str) -> bool {
        self.is_ident_at(0, "end") && self.is_ident_at(1, kw)
    }

    fn consume_end_kw(&mut self, kw: &str) -> Result<(), String> {
        self.expect_ident("end")?;
        self.expect_ident(kw)
    }

    fn is_elseif(&self) -> bool {
        self.is_ident_at(0, "elseif")
            || (self.is_ident_at(0, "else") && self.is_ident_at(1, "if"))
    }

    fn consume_elseif(&mut self) {
        if self.is_ident_at(0, "elseif") {
            self.advance();
        } else {
            self.advance(); // else
            self.advance(); // if
        }
    }

    // Parse a body of statements until `at_end` returns true or EOF.
    // Caller is responsible for consuming the terminator.
    fn parse_stmts<F: Fn(&Self) -> bool>(&mut self, at_end: F) -> Result<Vec<SpannedStmt>, String> {
        let mut stmts = vec![];
        loop {
            self.skip_nl();
            if *self.peek() == Tok::Eof || at_end(self) { break; }
            let start = self.peek_span().start;
            if let Some(s) = self.parse_stmt()? {
                let end = self.peek_span().start;
                stmts.push(SpannedStmt { stmt: s, span: SourceSpan { start, end } });
            }
        }
        Ok(stmts)
    }

    // ── Top-level ──────────────────────────────────────────────────────────────

    fn parse_program(&mut self) -> Result<Program, String> {
        self.skip_nl();
        let mut subs      = vec![];
        let mut funcs     = vec![];
        let mut type_defs = vec![];
        let mut module_diagnostics: Vec<(String, SourceSpan)> = vec![];
        let mut module_name: Option<String> = None;
        while *self.peek() != Tok::Eof {
            // Module-level Option declarations → no-op
            if self.is_ident("option") {
                self.skip_to_eol();
                continue;
            }
            // `Attribute VB_Name = "..."` names the module, as real VBA
            // does — captured for multi-module CLI use. Every other
            // Attribute line is still a no-op, same as before.
            if self.is_ident("attribute") {
                if self.is_ident_at(1, "vb_name") && *self.peek_at(2) == Tok::Eq {
                    self.advance(); // attribute
                    self.advance(); // vb_name
                    self.advance(); // =
                    if let Ok(name) = self.consume_str() {
                        module_name = Some(name);
                    }
                }
                self.skip_to_eol();
                continue;
            }
            // Access/scope modifiers before Sub, Function, or Type
            if self.is_ident("public") || self.is_ident("private")
                || self.is_ident("friend") || self.is_ident("static")
            {
                let start = self.peek_span().start;
                self.advance();
                if !self.is_ident("sub") && !self.is_ident("function") && !self.is_ident("type") {
                    // Module-level `Const` never gets its value evaluated
                    // anywhere (unlike inside a Sub) — a real gap, worth
                    // flagging. A plain `Public x As Long`/`Static y` etc.
                    // is a harmless no-op (no separate module scope exists;
                    // `Vm::variables` is one flat namespace) — same as
                    // plain `Dim` inside a Sub, left unflagged.
                    if self.is_ident("const") {
                        self.skip_to_eol();
                        let end = self.peek_span().start;
                        module_diagnostics.push((
                            "Module-level 'Const' is not evaluated (module-level constants aren't supported outside a Sub/Function) and was skipped".to_string(),
                            SourceSpan { start, end },
                        ));
                    } else {
                        self.skip_to_eol(); // module-level declaration (Dim, etc.) → skip
                    }
                    continue;
                }
            }
            if self.is_ident("sub") {
                subs.push(self.parse_sub()?);
            } else if self.is_ident("function") {
                funcs.push(self.parse_func()?);
            } else if self.is_ident("type") {
                type_defs.push(self.parse_type_def()?);
            } else if *self.peek() == Tok::Newline {
                self.advance();
            } else if self.is_ident("const") {
                // Bare module-level `Const` (no modifier) — same gap as above.
                let start = self.peek_span().start;
                self.skip_to_eol();
                let end = self.peek_span().start;
                module_diagnostics.push((
                    "Module-level 'Const' is not evaluated (module-level constants aren't supported outside a Sub/Function) and was skipped".to_string(),
                    SourceSpan { start, end },
                ));
            } else if self.is_ident("dim") {
                // Bare module-level `Dim` (no modifier) — harmless, same as Group A above.
                self.skip_to_eol();
            } else {
                // Unknown module-level line → genuinely unrecognized construct.
                let start = self.peek_span().start;
                let reason = if let Tok::Ident(name) = self.peek().clone() {
                    format!(
                        "Module-level statement starting with '{}' is not recognized and was skipped",
                        name
                    )
                } else {
                    "Module-level statement is not recognized and was skipped".to_string()
                };
                self.skip_to_eol();
                let end = self.peek_span().start;
                module_diagnostics.push((reason, SourceSpan { start, end }));
            }
        }
        Ok(Program {
            subs,
            funcs,
            type_defs,
            module_diagnostics,
            module_name,
        })
    }

    /// Parse a `Type Name ... End Type` block.
    fn parse_type_def(&mut self) -> Result<TypeDef, String> {
        self.expect_ident("type")?;
        let name = self.consume_ident()?.to_lowercase();
        self.eat_eol()?;
        let mut fields = vec![];
        loop {
            self.skip_nl();
            if self.is_end_kw("type") || *self.peek() == Tok::Eof { break; }
            // Each line: FieldName As TypeName  (or blank/comment)
            if let Tok::Ident(_) = self.peek().clone() {
                let field_name = self.consume_ident()?.to_lowercase();
                let vba_type = if self.is_ident("as") {
                    self.advance();
                    self.consume_ident()?.to_lowercase()
                } else {
                    "variant".into()
                };
                fields.push((field_name, vba_type));
            }
            self.skip_to_eol();
        }
        self.consume_end_kw("type")?;
        self.skip_nl();
        Ok(TypeDef { name, fields })
    }

    fn parse_sub(&mut self) -> Result<SubDef, String> {
        self.expect_ident("sub")?;
        let name = self.consume_ident()?;
        self.expect_tok(Tok::LParen)?;
        let params = self.parse_params()?;
        self.expect_tok(Tok::RParen)?;
        self.eat_eol()?;
        let body = self.parse_stmts(|p| p.is_end_kw("sub"))?;
        self.consume_end_kw("sub")?;
        self.skip_nl();
        Ok(SubDef { name, params, body })
    }

    fn parse_func(&mut self) -> Result<FuncDef, String> {
        self.expect_ident("function")?;
        let name = self.consume_ident()?;
        self.expect_tok(Tok::LParen)?;
        let params = self.parse_params()?;
        self.expect_tok(Tok::RParen)?;
        self.eat_eol()?;
        let body = self.parse_stmts(|p| p.is_end_kw("function"))?;
        self.consume_end_kw("function")?;
        self.skip_nl();
        Ok(FuncDef { name, params, body })
    }

    fn parse_params(&mut self) -> Result<Vec<String>, String> {
        let mut params = vec![];
        while !matches!(self.peek(), Tok::RParen | Tok::Eof) {
            let name = self.consume_ident()?;
            params.push(name);
            // optional: As <type>
            if self.is_ident("as") {
                self.advance();
                self.consume_ident()?; // type name
            }
            if *self.peek() == Tok::Comma { self.advance(); }
        }
        Ok(params)
    }

    // ── Statement dispatch ─────────────────────────────────────────────────────

    fn parse_stmt(&mut self) -> Result<Option<Stmt>, String> {
        // The tok at this point is not a Newline (caller skips those)
        let first = match self.peek() {
            Tok::Ident(s) => s.clone(),
            Tok::Eof | Tok::Newline => return Ok(None),
            _ => return Err(format!("unexpected token starting statement: {:?}", self.peek())),
        };

        // A bare `name = ...` is always a plain assignment, even when `name`
        // collides with one of the statement keywords below (e.g. `do = 0`,
        // `select = 1`) — no VBA statement keyword's grammar puts `=`
        // immediately after itself (Dim/Const/For/etc. all require a name or
        // expression there instead), so this check is safe and general
        // rather than needing a per-keyword lookahead guard (the `"on" if
        // ...`-style fix below only disambiguates `On Error` specifically).
        if *self.peek_at(1) == Tok::Eq {
            let s = self.parse_ident_stmt()?;
            self.eat_eol()?;
            return Ok(Some(s));
        }

        match first.as_str() {
            "do"      => Ok(Some(self.parse_do_loop()?)),
            "select"  => Ok(Some(self.parse_select_case()?)),
            "with"    => Ok(Some(self.parse_with()?)),
            "for" if self.is_ident_at(1, "each") => Ok(Some(self.parse_for_each()?)),
            "for"     => Ok(Some(self.parse_for()?)),
            "if"      => Ok(Some(self.parse_if()?)),
            "while"   => Ok(Some(self.parse_while_wend()?)),
            "exit"    => { let s = self.parse_exit()?; self.eat_eol()?; Ok(Some(s)) }
            "on" if self.is_ident_at(1, "error") => { let s = self.parse_on_error()?; self.eat_eol()?; Ok(Some(s)) }
            "goto"    => {
                self.advance();
                let label = self.consume_ident()?;
                self.eat_eol()?;
                Ok(Some(Stmt::GoTo(label)))
            }
            "resume"  => {
                self.advance();
                let next = if self.is_ident("next") { self.advance(); true } else { false };
                self.eat_eol()?;
                Ok(Some(Stmt::Resume { next }))
            }
            "dim"     => { let s = self.parse_dim()?; self.eat_eol()?; Ok(Some(s)) }
            "redim"   => { let s = self.parse_redim()?; self.eat_eol()?; Ok(Some(s)) }
            "const"   => { let s = self.parse_const()?; self.eat_eol()?; Ok(Some(s)) }
            "msgbox"  => { let s = self.parse_msgbox()?; self.eat_eol()?; Ok(Some(s)) }
            "call"    => { let s = self.parse_call_stmt()?; self.eat_eol()?; Ok(Some(s)) }
            "range"   => { let s = self.parse_range_stmt()?; self.eat_eol()?; Ok(Some(s)) }
            "cells"   => { let s = self.parse_cell_write_stmt()?; self.eat_eol()?; Ok(Some(s)) }
            "application" => { let s = self.parse_application_stmt()?; self.eat_eol()?; Ok(Some(s)) }
            "worksheetfunction" => { let s = self.parse_wsf_call_stmt(None)?; self.eat_eol()?; Ok(Some(s)) }
            "worksheets" | "sheets" => { let s = self.parse_sheets_stmt()?; self.eat_eol()?; Ok(Some(s)) }
            "workbooks" => { let s = self.parse_workbook_qualified_stmt()?; self.eat_eol()?; Ok(Some(s)) }
            // Access/scope modifiers before Dim/Const inside a sub
            "public" | "private" | "static" | "friend" => {
                self.advance(); // consume modifier
                if self.is_ident("dim") {
                    let s = self.parse_dim()?; self.eat_eol()?; Ok(Some(s))
                } else if self.is_ident("const") {
                    let s = self.parse_const()?; self.eat_eol()?; Ok(Some(s))
                } else {
                    self.skip_to_eol(); Ok(Some(Stmt::Dim))
                }
            }
            // Debug.Print / Debug.Assert → no-op
            "debug" => {
                self.skip_to_eol();
                Ok(Some(Stmt::Unsupported {
                    reason: "Debug.Print/Debug.Assert has no effect (no-op)".to_string(),
                }))
            }
            _ => { let s = self.parse_ident_stmt()?; self.eat_eol()?; Ok(Some(s)) }
        }
    }

    // ── Control flow ───────────────────────────────────────────────────────────

    fn parse_for(&mut self) -> Result<Stmt, String> {
        self.expect_ident("for")?;
        let var = self.consume_ident()?;
        self.expect_tok(Tok::Eq)?;
        let from = self.parse_expr()?;
        self.expect_ident("to")?;
        let to = self.parse_expr()?;
        let step = if self.is_ident("step") {
            self.advance();
            Some(self.parse_expr()?)
        } else { None };
        self.eat_eol()?;
        let body = self.parse_stmts(|p| p.is_ident("next"))?;
        self.expect_ident("next")?;
        if matches!(self.peek(), Tok::Ident(_)) { self.advance(); } // optional loop var
        self.skip_nl();
        Ok(Stmt::For { var, from, to, step, body })
    }

    fn parse_for_each(&mut self) -> Result<Stmt, String> {
        self.expect_ident("for")?;
        self.expect_ident("each")?;
        let var = self.consume_ident()?;
        self.expect_ident("in")?;
        let range_addr = self.parse_for_each_source()?;
        self.eat_eol()?;
        let body = self.parse_stmts(|p| p.is_ident("next"))?;
        self.expect_ident("next")?;
        if matches!(self.peek(), Tok::Ident(_)) { self.advance(); }
        self.skip_nl();
        Ok(Stmt::ForEach { var, range_addr, body })
    }

    fn parse_for_each_source(&mut self) -> Result<String, String> {
        if self.is_ident("range") {
            self.advance();
            self.expect_tok(Tok::LParen)?;
            let addr = self.consume_str()?.to_uppercase();
            self.expect_tok(Tok::RParen)?;
            Ok(addr)
        } else {
            self.consume_ident()?;
            Ok(String::new())
        }
    }

    fn parse_if(&mut self) -> Result<Stmt, String> {
        self.expect_ident("if")?;
        let condition = self.parse_expr()?;
        self.expect_ident("then")?;
        self.eat_eol()?;
        let then_body = self.parse_stmts(|p| {
            p.is_elseif() || p.is_ident("else") || p.is_end_kw("if")
        })?;
        let else_body = if self.is_elseif() {
            self.parse_elseif_chain()?
        } else if self.is_ident("else") {
            self.advance(); // "else"
            self.eat_eol()?;
            self.parse_stmts(|p| p.is_end_kw("if"))?
        } else {
            vec![]
        };
        self.consume_end_kw("if")?;
        self.skip_nl();
        Ok(Stmt::If { condition, then_body, else_body })
    }

    fn parse_elseif_chain(&mut self) -> Result<Vec<SpannedStmt>, String> {
        let start = self.peek_span().start;
        self.consume_elseif();
        let condition = self.parse_expr()?;
        self.expect_ident("then")?;
        self.eat_eol()?;
        let then_body = self.parse_stmts(|p| {
            p.is_elseif() || p.is_ident("else") || p.is_end_kw("if")
        })?;
        let else_body = if self.is_elseif() {
            self.parse_elseif_chain()?
        } else if self.is_ident("else") {
            self.advance();
            self.eat_eol()?;
            self.parse_stmts(|p| p.is_end_kw("if"))?
        } else {
            vec![]
        };
        let end = self.peek_span().start;
        let stmt = Stmt::If { condition, then_body, else_body };
        Ok(vec![SpannedStmt { stmt, span: SourceSpan { start, end } }])
    }

    fn parse_do_loop(&mut self) -> Result<Stmt, String> {
        self.expect_ident("do")?;
        let pre_cond = if self.is_ident("while") || self.is_ident("until") {
            Some(self.parse_do_cond()?)
        } else { None };
        self.eat_eol()?;
        let body = self.parse_stmts(|p| p.is_ident("loop"))?;
        self.expect_ident("loop")?;
        let post_cond = if self.is_ident("while") || self.is_ident("until") {
            Some(self.parse_do_cond()?)
        } else { None };
        self.skip_nl();
        Ok(Stmt::DoLoop { pre_cond, post_cond, body })
    }

    fn parse_do_cond(&mut self) -> Result<(bool, Expr), String> {
        let is_until = self.is_ident("until");
        self.advance(); // while or until
        let expr = self.parse_expr()?;
        Ok((is_until, expr))
    }

    fn parse_while_wend(&mut self) -> Result<Stmt, String> {
        self.expect_ident("while")?;
        let condition = self.parse_expr()?;
        self.eat_eol()?;
        let body = self.parse_stmts(|p| p.is_ident("wend"))?;
        self.expect_ident("wend")?;
        self.skip_nl();
        Ok(Stmt::DoLoop {
            pre_cond: Some((false, condition)),
            post_cond: None,
            body,
        })
    }

    fn parse_select_case(&mut self) -> Result<Stmt, String> {
        self.expect_ident("select")?;
        self.expect_ident("case")?;
        let expr = self.parse_expr()?;
        self.eat_eol()?;
        self.skip_nl();
        let mut cases = vec![];
        let mut else_body = vec![];
        loop {
            if self.is_end_kw("select") || *self.peek() == Tok::Eof { break; }
            if !self.is_ident("case") {
                return Err(format!("expected 'Case' in Select Case, got {:?}", self.peek()));
            }
            self.advance(); // "case"
            if self.is_ident("else") {
                self.advance(); // "else"
                self.eat_eol()?;
                else_body = self.parse_stmts(|p| p.is_ident("case") || p.is_end_kw("select"))?;
            } else {
                let matches = self.parse_case_match_list()?;
                self.eat_eol()?;
                let body = self.parse_stmts(|p| p.is_ident("case") || p.is_end_kw("select"))?;
                cases.push((matches, body));
            }
        }
        self.consume_end_kw("select")?;
        self.skip_nl();
        Ok(Stmt::SelectCase { expr, cases, else_body })
    }

    fn parse_case_match_list(&mut self) -> Result<Vec<CaseMatch>, String> {
        let mut matches = vec![];
        matches.push(self.parse_case_match()?);
        while *self.peek() == Tok::Comma {
            self.advance();
            matches.push(self.parse_case_match()?);
        }
        Ok(matches)
    }

    fn parse_case_match(&mut self) -> Result<CaseMatch, String> {
        if self.is_ident("is") {
            self.advance();
            let op = self.parse_cmp_op()?;
            let expr = self.parse_expr()?;
            Ok(CaseMatch::IsOp(op, expr))
        } else {
            let lhs = self.parse_expr()?;
            if self.is_ident("to") {
                self.advance();
                let rhs = self.parse_expr()?;
                Ok(CaseMatch::Range(lhs, rhs))
            } else {
                Ok(CaseMatch::Value(lhs))
            }
        }
    }

    fn parse_cmp_op(&mut self) -> Result<VbaBinOp, String> {
        let op = match self.peek() {
            Tok::Eq    => VbaBinOp::Eq,
            Tok::Ne    => VbaBinOp::Ne,
            Tok::Lt    => VbaBinOp::Lt,
            Tok::Le    => VbaBinOp::Le,
            Tok::Gt    => VbaBinOp::Gt,
            Tok::Ge    => VbaBinOp::Ge,
            t => return Err(format!("expected comparison operator, got {:?}", t)),
        };
        self.advance();
        Ok(op)
    }

    fn parse_with(&mut self) -> Result<Stmt, String> {
        self.expect_ident("with")?;

        // ── Sheets/Worksheets("name") ─────────────────────────────────────────
        if self.is_ident("sheets") || self.is_ident("worksheets") {
            self.advance();
            if *self.peek() == Tok::LParen {
                self.advance();
                let name = self.consume_str()?.to_lowercase();
                self.expect_tok(Tok::RParen)?;
                self.eat_eol()?;
                let body = self.parse_with_body()?;
                self.consume_end_kw("with")?;
                self.skip_nl();
                return Ok(Stmt::WithSheet { sheet_name: name, body });
            } else {
                self.skip_to_eol();
            }
            let body = self.parse_with_body()?;
            self.consume_end_kw("with")?;
            self.skip_nl();
            return Ok(Stmt::With { body });
        }

        // ── With <variable> — UDT target ─────────────────────────────────────
        if let Tok::Ident(_) = self.peek().clone() {
            let var = self.consume_ident()?.to_lowercase();
            self.eat_eol()?;
            let prev = self.with_target.replace(var.clone());
            let body = self.parse_with_body()?;
            self.with_target = prev;
            self.consume_end_kw("with")?;
            self.skip_nl();
            return Ok(Stmt::WithRecord { var, body });
        }

        // ── Generic / Application etc. — no-op body ───────────────────────────
        self.skip_to_eol();
        let body = self.parse_with_body()?;
        self.consume_end_kw("with")?;
        self.skip_nl();
        Ok(Stmt::With { body })
    }

    fn parse_with_body(&mut self) -> Result<Vec<SpannedStmt>, String> {
        let mut stmts = vec![];
        loop {
            self.skip_nl();
            if self.is_end_kw("with") || *self.peek() == Tok::Eof { break; }
            let start = self.peek_span().start;
            let stmt = if *self.peek() == Tok::Dot {
                // with_cell_write, with_range_write, or with_dot_stmt
                self.parse_with_dot_stmt()?
            } else {
                self.parse_stmt()?
            };
            if let Some(stmt) = stmt {
                let end = self.peek_span().start;
                stmts.push(SpannedStmt { stmt, span: SourceSpan { start, end } });
            }
        }
        Ok(stmts)
    }

    fn parse_with_dot_stmt(&mut self) -> Result<Option<Stmt>, String> {
        self.advance(); // consume '.'
        match self.peek().clone() {
            Tok::Ident(ref s) => {
                let s = s.clone();
                // ── UDT With target: .Field = val  /  .A.B = val ──────────────
                if let Some(var) = self.with_target.clone() {
                    if s != "range" && s != "cells" {
                        let field = self.consume_ident()?.to_lowercase();
                        let mut fields = vec![field];
                        // Collect chained fields: .A.B.C
                        while *self.peek() == Tok::Dot && matches!(self.peek_at(1), Tok::Ident(_)) {
                            self.advance(); // consume '.'
                            fields.push(self.consume_ident()?.to_lowercase());
                        }
                        if *self.peek() == Tok::Eq {
                            self.advance();
                            let value = self.parse_expr()?;
                            self.eat_eol()?;
                            return if fields.len() == 1 {
                                Ok(Some(Stmt::RecordSet { var, field: fields.remove(0), value }))
                            } else {
                                Ok(Some(Stmt::RecordSetNested { var, fields, value }))
                            };
                        }
                        // No '=' → read without assignment has no effect
                        self.skip_to_eol();
                        return Ok(Some(Stmt::Unsupported {
                            reason: format!(
                                "'{}.{}' read without assignment has no effect",
                                var,
                                fields.join(".")
                            ),
                        }));
                    }
                }
                match s.as_str() {
                    "range" => {
                        // .Range("addr").Value/Formula = expr
                        self.advance();
                        self.expect_tok(Tok::LParen)?;
                        let addr = self.consume_str()?;
                        self.expect_tok(Tok::RParen)?;
                        self.expect_tok(Tok::Dot)?;
                        let prop = self.consume_ident()?;
                        let is_formula = prop == "formula";
                        self.expect_tok(Tok::Eq)?;
                        let value = self.parse_expr()?;
                        self.eat_eol()?;
                        Ok(Some(Stmt::RangeWrite { addr, is_formula, value }))
                    }
                    "cells" => {
                        // .Cells(r, c).Value = expr
                        self.advance();
                        self.expect_tok(Tok::LParen)?;
                        let row = self.parse_expr()?;
                        self.expect_tok(Tok::Comma)?;
                        let col = self.parse_expr()?;
                        self.expect_tok(Tok::RParen)?;
                        self.expect_tok(Tok::Dot)?;
                        self.expect_ident("value")?;
                        self.expect_tok(Tok::Eq)?;
                        let value = self.parse_expr()?;
                        self.eat_eol()?;
                        Ok(Some(Stmt::CellWrite { row, col, value }))
                    }
                    _ => {
                        // with_dot_stmt: unrecognized property/method
                        self.skip_to_eol();
                        Ok(Some(Stmt::Unsupported {
                            reason: format!("With-block '.{}' is not implemented", s),
                        }))
                    }
                }
            }
            _ => {
                self.skip_to_eol();
                Ok(Some(Stmt::Unsupported {
                    reason: "With-block dotted statement is not recognized and was skipped"
                        .to_string(),
                }))
            }
        }
    }

    fn parse_exit(&mut self) -> Result<Stmt, String> {
        self.expect_ident("exit")?;
        match self.consume_ident()?.as_str() {
            "for"      => Ok(Stmt::ExitFor),
            "do"       => Ok(Stmt::ExitDo),
            "sub"      => Ok(Stmt::ExitSub),
            "function" => Ok(Stmt::ExitFunction),
            other => Err(format!("unknown exit target: {}", other)),
        }
    }

    fn parse_on_error(&mut self) -> Result<Stmt, String> {
        self.expect_ident("on")?;
        self.expect_ident("error")?;
        if self.is_ident("resume") {
            self.advance();
            self.expect_ident("next")?;
            Ok(Stmt::OnError { resume_next: true })
        } else if self.is_ident("goto") {
            self.advance();
            match self.peek().clone() {
                Tok::Int(0) => { self.advance(); Ok(Stmt::OnError { resume_next: false }) }
                Tok::Ident(_) => {
                    let label = self.consume_ident()?;
                    Ok(Stmt::OnErrorGoTo(label))
                }
                _ => { self.advance(); Ok(Stmt::OnError { resume_next: false }) }
            }
        } else {
            Err(format!("unexpected On Error action: {:?}", self.peek()))
        }
    }

    // ── Simple statements ──────────────────────────────────────────────────────

    /// Known VBA built-in type names that do NOT correspond to a user-defined type.
    fn is_vba_builtin_type(name: &str) -> bool {
        matches!(name, "integer" | "long" | "longlong" | "single" | "double" | "currency"
            | "boolean" | "string" | "date" | "object" | "variant" | "byte" | "decimal")
    }

    fn parse_dim(&mut self) -> Result<Stmt, String> {
        self.expect_ident("dim")?;
        // dim_array_decl: ident (
        if matches!(self.peek(), Tok::Ident(_)) && *self.peek_at(1) == Tok::LParen {
            let name = self.consume_ident()?;
            self.advance(); // (
            let mut sizes = vec![self.parse_expr()?];
            while *self.peek() == Tok::Comma {
                self.advance();
                sizes.push(self.parse_expr()?);
            }
            self.expect_tok(Tok::RParen)?;
            if self.is_ident("as") {
                self.advance();
                let type_name = self.consume_ident()?.to_lowercase();
                if !Self::is_vba_builtin_type(&type_name) {
                    return Ok(Stmt::DimArrayRecord { name, sizes, type_name });
                }
            }
            Ok(Stmt::DimArray { name, sizes })
        } else if matches!(self.peek(), Tok::Ident(_)) {
            // Dim varName [As TypeName]
            let var = self.consume_ident()?;
            if self.is_ident("as") {
                self.advance();
                let type_name = self.consume_ident()?.to_lowercase();
                // Emit DimRecord only for non-built-in types (user-defined types).
                if !Self::is_vba_builtin_type(&type_name) {
                    return Ok(Stmt::DimRecord { var, type_name });
                }
            }
            // Built-in type or bare Dim → no-op
            while !matches!(self.peek(), Tok::Newline | Tok::Eof) { self.advance(); }
            Ok(Stmt::Dim)
        } else {
            // dim_rest: skip to EOL
            while !matches!(self.peek(), Tok::Newline | Tok::Eof) { self.advance(); }
            Ok(Stmt::Dim)
        }
    }

    fn parse_redim(&mut self) -> Result<Stmt, String> {
        self.expect_ident("redim")?;
        let preserve = if self.is_ident("preserve") { self.advance(); true } else { false };
        let name = self.consume_ident()?;
        self.expect_tok(Tok::LParen)?;
        let mut sizes = vec![self.parse_expr()?];
        while *self.peek() == Tok::Comma {
            self.advance();
            sizes.push(self.parse_expr()?);
        }
        self.expect_tok(Tok::RParen)?;
        if self.is_ident("as") { self.advance(); self.consume_ident()?; }
        Ok(Stmt::ReDim { name, sizes, preserve })
    }

    fn parse_const(&mut self) -> Result<Stmt, String> {
        self.expect_ident("const")?;
        let var = self.consume_ident()?;
        if self.is_ident("as") { self.advance(); self.consume_ident()?; }
        self.expect_tok(Tok::Eq)?;
        let value = self.parse_expr()?;
        Ok(Stmt::Assignment { var, value })
    }

    fn parse_msgbox(&mut self) -> Result<Stmt, String> {
        self.expect_ident("msgbox")?;
        let message = self.parse_expr()?;
        // optional extra args (title, buttons) — ignore
        while *self.peek() == Tok::Comma {
            self.advance();
            self.parse_expr()?;
        }
        Ok(Stmt::MsgBox { message })
    }

    fn parse_call_stmt(&mut self) -> Result<Stmt, String> {
        self.expect_ident("call")?;
        let name = self.consume_ident()?;
        self.expect_tok(Tok::LParen)?;
        let args = self.parse_arg_list()?;
        self.expect_tok(Tok::RParen)?;
        Ok(Stmt::CallSub { name, args })
    }

    // ── Range family ───────────────────────────────────────────────────────────

    fn parse_range_stmt(&mut self) -> Result<Stmt, String> {
        self.expect_ident("range")?;
        self.expect_tok(Tok::LParen)?;
        let addr = self.consume_str()?;
        self.expect_tok(Tok::RParen)?;
        self.expect_tok(Tok::Dot)?;

        let prop = self.consume_ident()?;
        match prop.as_str() {
            "value" | "formula" => {
                let is_formula = prop == "formula";
                self.expect_tok(Tok::Eq)?;
                let value = self.parse_expr()?;
                Ok(Stmt::RangeWrite { addr, is_formula, value })
            }
            "copy" => {
                // Optional: Destination:=Range("dst") — a bare `.Copy` (no
                // Destination) only populates the clipboard (Milestone B6b).
                let dst = if self.is_ident("destination") {
                    self.advance();
                    self.expect_tok(Tok::ColonEq)?;
                    self.expect_ident("range")?;
                    self.expect_tok(Tok::LParen)?;
                    let d = self.consume_str()?;
                    self.expect_tok(Tok::RParen)?;
                    Some(d)
                } else {
                    None
                };
                Ok(Stmt::RangeCopy { src: addr, dst })
            }
            "paste" => Ok(Stmt::RangePaste {
                dest_addr: addr,
                transpose: None,
            }),
            "pastespecial" => {
                // Optional kwargs; only Transpose:= is modeled (Milestone
                // B6b) — others (Paste:=, Operation:=, SkipBlanks:=, ...)
                // are evaluated and discarded, same convention as
                // `Stmt::SetAppProp` for unmodeled Application properties.
                let mut transpose = None;
                while *self.peek() != Tok::Newline && *self.peek() != Tok::Eof {
                    if !matches!(self.peek(), Tok::Ident(_)) {
                        self.advance();
                        continue;
                    }
                    let kw_name = self.consume_ident()?;
                    if *self.peek() != Tok::ColonEq {
                        continue;
                    }
                    self.advance(); // :=
                    match kw_name.as_str() {
                        "transpose" => {
                            transpose = Some(self.parse_expr()?);
                        }
                        _ => {
                            self.parse_expr()?;
                        }
                    }
                    if *self.peek() == Tok::Comma {
                        self.advance();
                    }
                }
                Ok(Stmt::RangePaste {
                    dest_addr: addr,
                    transpose,
                })
            }
            "sort" => {
                // Optional kwargs: Key1:=Range("A1"), Order1:=xlAscending/xlDescending, etc.
                let mut key_col: u32 = 1;
                let mut descending = false;
                while *self.peek() != Tok::Newline && *self.peek() != Tok::Eof {
                    if !matches!(self.peek(), Tok::Ident(_)) { self.advance(); continue; }
                    let kw_name = self.consume_ident()?;
                    if *self.peek() != Tok::ColonEq { continue; }
                    self.advance(); // :=
                    match kw_name.as_str() {
                        "key1" => {
                            if self.is_ident("range") {
                                self.advance();
                                self.expect_tok(Tok::LParen)?;
                                let key_addr = self.consume_str()?;
                                self.expect_tok(Tok::RParen)?;
                                let trimmed = key_addr.trim_matches('"');
                                if let Some((col, _)) = parse_cell_addr(trimmed) {
                                    key_col = col;
                                }
                            } else {
                                self.parse_expr()?;
                            }
                        }
                        "order1" => {
                            let val = match self.peek().clone() {
                                Tok::Ident(s) => { self.advance(); s }
                                _ => { self.parse_expr()?; String::new() }
                            };
                            descending = val.contains("descend");
                        }
                        _ => { self.parse_expr()?; }
                    }
                    if *self.peek() == Tok::Comma { self.advance(); }
                }
                Ok(Stmt::RangeSort { addr, key_col, descending })
            }
            "delete" => Ok(Stmt::RangeDelete { addr }),
            "insert" => {
                // optional kwargs
                while *self.peek() != Tok::Newline && *self.peek() != Tok::Eof { self.advance(); }
                Ok(Stmt::RangeInsert { addr })
            }
            "offset" => {
                self.expect_tok(Tok::LParen)?;
                let row_off = self.parse_expr()?;
                self.expect_tok(Tok::Comma)?;
                let col_off = self.parse_expr()?;
                self.expect_tok(Tok::RParen)?;
                self.expect_tok(Tok::Dot)?;
                self.expect_ident("value")?;
                self.expect_tok(Tok::Eq)?;
                let value = self.parse_expr()?;
                Ok(Stmt::RangeOffsetWrite { addr: addr.to_uppercase(), row_off, col_off, value })
            }
            "entirerow" | "entirecolumn" => {
                self.expect_tok(Tok::Dot)?;
                let method = self.consume_ident()?;
                match method.as_str() {
                    "delete" => Ok(Stmt::RangeDelete { addr }),
                    "clearcontents" | "clear" => Ok(Stmt::RangeClear {
                        addr,
                        contents_only: method == "clearcontents",
                    }),
                    _ => {
                        // Leave the trailing newline for the caller's own
                        // `eat_eol()` (the "range" dispatch arm) — unlike
                        // `skip_to_eol()`, which would consume it too and
                        // cause a spurious "expected newline" error when
                        // this is the last statement before End Sub.
                        while !matches!(self.peek(), Tok::Newline | Tok::Eof) {
                            self.advance();
                        }
                        Ok(Stmt::Unsupported {
                            reason: format!("EntireRow/EntireColumn.{} is not implemented", method),
                        })
                    }
                }
            }
            "clearcontents" | "clear" => Ok(Stmt::RangeClear {
                addr,
                contents_only: prop == "clearcontents",
            }),
            "name" => {
                self.expect_tok(Tok::Eq)?;
                let name = self.consume_str()?;
                Ok(Stmt::RangeName { addr, name })
            }
            _ => {
                // range_noop_stmt
                while !matches!(self.peek(), Tok::Newline | Tok::Eof) { self.advance(); }
                Ok(Stmt::Unsupported {
                    reason: format!("Range property/method '{}' is not implemented", prop),
                })
            }
        }
    }

    fn parse_cell_write_stmt(&mut self) -> Result<Stmt, String> {
        self.expect_ident("cells")?;
        self.expect_tok(Tok::LParen)?;
        let row = self.parse_expr()?;
        self.expect_tok(Tok::Comma)?;
        let col = self.parse_expr()?;
        self.expect_tok(Tok::RParen)?;
        self.expect_tok(Tok::Dot)?;
        self.expect_ident("value")?;
        self.expect_tok(Tok::Eq)?;
        let value = self.parse_expr()?;
        Ok(Stmt::CellWrite { row, col, value })
    }

    fn parse_application_stmt(&mut self) -> Result<Stmt, String> {
        self.expect_ident("application")?;
        self.expect_tok(Tok::Dot)?;
        let prop = self.consume_ident()?;
        match prop.as_str() {
            "worksheetfunction" => self.parse_wsf_call_stmt(None),
            "calculation" => {
                self.expect_tok(Tok::Eq)?;
                let val = self.consume_ident()?;
                let mode = if val.contains("automatic") {
                    CalcModeValue::Automatic
                } else {
                    CalcModeValue::Manual
                };
                Ok(Stmt::SetCalcMode(mode))
            }
            other => {
                self.expect_tok(Tok::Eq)?;
                let value = self.parse_expr()?;
                Ok(Stmt::SetAppProp { prop: other.to_string(), value })
            }
        }
    }

    fn parse_wsf_call_stmt(&mut self, _prefix: Option<()>) -> Result<Stmt, String> {
        // consume "worksheetfunction" if still present
        if self.is_ident("worksheetfunction") { self.advance(); }
        self.expect_tok(Tok::Dot)?;
        let name = self.consume_ident()?;
        self.expect_tok(Tok::LParen)?;
        let args = self.parse_arg_list()?;
        self.expect_tok(Tok::RParen)?;
        Ok(Stmt::Assignment {
            var: "_".into(),
            value: Expr::FuncCall { name: format!("wsf_{}", name), args },
        })
    }

    /// A sheet key inside `Sheets(...)`/`Worksheets(...)`: either a string
    /// literal name (the common case) or a 1-based numeric index
    /// (Milestone B6a — lets `diagnose` classify an out-of-range index).
    /// elixcee doesn't track real workbook tab order, so a numeric index
    /// resolves against `Vm::sheet_names()`'s alphabetical order at
    /// runtime, not Excel's left-to-right tab order — an honest fidelity
    /// gap, documented in `docs/agent-contract.md`.
    ///
    /// Unlike the pre-B6a `.Cells(...)` path, a string name is kept in its
    /// as-written case here (not lowercased at parse time) — resolution
    /// (`Vm::resolve_sheet_expr`) lowercases only when it needs a
    /// `self.sheets` lookup key, so `diagnose`'s evidence can still show
    /// the name the macro actually wrote.
    fn parse_sheet_key(&mut self) -> Result<Expr, String> {
        match self.peek().clone() {
            Tok::Str(_) => Ok(Expr::Str(self.consume_str()?)),
            Tok::Int(n) => {
                self.advance();
                Ok(Expr::Integer(n))
            }
            other => Err(format!("expected a sheet name or index, got {:?}", other)),
        }
    }

    /// Parses the `.Cells(r,c).Value = ...` / `.Range(addr).Value|Formula =
    /// ...` / `.Delete` suffix shared by `Sheets(...)` and
    /// `Workbooks(...).Worksheets(...)` statement forms.
    fn parse_sheet_property_write(&mut self, sheet: Expr) -> Result<Stmt, String> {
        self.expect_tok(Tok::Dot)?;
        let method = self.consume_ident()?;
        match method.as_str() {
            "delete" => Ok(Stmt::SheetsDelete { sheet }),
            "cells" => {
                self.expect_tok(Tok::LParen)?;
                let row = self.parse_expr()?;
                self.expect_tok(Tok::Comma)?;
                let col = self.parse_expr()?;
                self.expect_tok(Tok::RParen)?;
                self.expect_tok(Tok::Dot)?;
                self.expect_ident("value")?;
                self.expect_tok(Tok::Eq)?;
                let value = self.parse_expr()?;
                Ok(Stmt::SheetCellWrite {
                    sheet,
                    row,
                    col,
                    value,
                })
            }
            "range" => {
                self.expect_tok(Tok::LParen)?;
                let addr = self.consume_str()?;
                self.expect_tok(Tok::RParen)?;
                self.expect_tok(Tok::Dot)?;
                let prop = self.consume_ident()?;
                let is_formula = match prop.as_str() {
                    "value" => false,
                    "formula" => true,
                    other => {
                        return Err(format!("unexpected property after Range(...): {}", other));
                    }
                };
                self.expect_tok(Tok::Eq)?;
                let value = self.parse_expr()?;
                Ok(Stmt::SheetRangeWrite {
                    sheet,
                    addr,
                    is_formula,
                    value,
                })
            }
            "paste" => {
                // Worksheets(sheet).Paste Destination:=Range(addr) — real
                // VBA's Worksheet.Paste has no Transpose:= parameter
                // (Milestone B6b).
                self.expect_ident("destination")?;
                self.expect_tok(Tok::ColonEq)?;
                self.expect_ident("range")?;
                self.expect_tok(Tok::LParen)?;
                let dest_addr = self.consume_str()?;
                self.expect_tok(Tok::RParen)?;
                Ok(Stmt::SheetRangePaste { sheet, dest_addr })
            }
            "protect" | "unprotect" => {
                // Optional kwargs; only UserInterfaceOnly:= is modeled
                // (Milestone B6c) — others (Password:=, DrawingObjects:=,
                // Contents:=, etc.) are evaluated and discarded, same
                // convention as `Stmt::SetAppProp`/`.PasteSpecial`.
                let mut ui_only = None;
                while *self.peek() != Tok::Newline && *self.peek() != Tok::Eof {
                    if !matches!(self.peek(), Tok::Ident(_)) {
                        self.advance();
                        continue;
                    }
                    let kw_name = self.consume_ident()?;
                    if *self.peek() != Tok::ColonEq {
                        continue;
                    }
                    self.advance(); // :=
                    match kw_name.as_str() {
                        "userinterfaceonly" => {
                            ui_only = Some(self.parse_expr()?);
                        }
                        _ => {
                            self.parse_expr()?;
                        }
                    }
                    if *self.peek() == Tok::Comma {
                        self.advance();
                    }
                }
                Ok(Stmt::SheetProtection {
                    sheet,
                    protect: method == "protect",
                    ui_only,
                })
            }
            _ => {
                while !matches!(self.peek(), Tok::Newline | Tok::Eof) {
                    self.advance();
                }
                Ok(Stmt::Unsupported {
                    reason: format!("Sheets(...).{} is not implemented", method),
                })
            }
        }
    }

    fn parse_sheets_stmt(&mut self) -> Result<Stmt, String> {
        // worksheets or sheets
        self.consume_ident()?; // consume "worksheets" or "sheets"
        if *self.peek() == Tok::Dot {
            // sheets.add ...
            self.advance(); // dot
            let method = self.consume_ident()?;
            if method == "add" {
                while !matches!(self.peek(), Tok::Newline | Tok::Eof) {
                    self.advance();
                }
                return Ok(Stmt::SheetsAdd);
            }
            // Leave the trailing newline for the caller's own `eat_eol()`
            // (the "worksheets"/"sheets" dispatch arm) — see the identical
            // note on the EntireRow/EntireColumn fallback above.
            while !matches!(self.peek(), Tok::Newline | Tok::Eof) {
                self.advance();
            }
            return Ok(Stmt::Unsupported {
                reason: format!("Sheets.{} is not implemented", method),
            });
        }
        self.expect_tok(Tok::LParen)?;
        let sheet = self.parse_sheet_key()?;
        self.expect_tok(Tok::RParen)?;
        self.parse_sheet_property_write(sheet)
    }

    /// `Workbooks(workbook).Worksheets(sheet).Cells(...)`/`.Range(...)` —
    /// Milestone B6a. elixcee never has more than one workbook loaded, so
    /// this exists only so a mismatched workbook name/index can be
    /// diagnosed (`ResolutionFailureKind::WorkbookNotFound`), not to model
    /// real multi-workbook switching.
    fn parse_workbook_qualified_stmt(&mut self) -> Result<Stmt, String> {
        self.expect_ident("workbooks")?;
        self.expect_tok(Tok::LParen)?;
        let workbook = self.parse_sheet_key()?;
        self.expect_tok(Tok::RParen)?;
        self.expect_tok(Tok::Dot)?;
        if !(self.is_ident("worksheets") || self.is_ident("sheets")) {
            while !matches!(self.peek(), Tok::Newline | Tok::Eof) {
                self.advance();
            }
            return Ok(Stmt::Unsupported {
                reason:
                    "Workbooks(...) is only supported followed by .Worksheets(...)/.Sheets(...)"
                        .to_string(),
            });
        }
        self.advance();
        self.expect_tok(Tok::LParen)?;
        let sheet = self.parse_sheet_key()?;
        self.expect_tok(Tok::RParen)?;
        let qualified = Expr::WorkbookQualifiedSheet {
            workbook: Box::new(workbook),
            sheet: Box::new(sheet),
        };
        self.parse_sheet_property_write(qualified)
    }

    // ident-starting: assignment, array_write, call_stmt (without Call keyword)
    fn parse_ident_stmt(&mut self) -> Result<Stmt, String> {
        let name = self.consume_ident()?;
        // Label: "ErrHandler:"
        if *self.peek() == Tok::Colon {
            self.advance();
            return Ok(Stmt::Label(name));
        }
        if *self.peek() == Tok::LParen {
            self.advance(); // (
            let mut args: Vec<Expr> = vec![];
            if *self.peek() != Tok::RParen {
                args.push(self.parse_expr()?);
                while *self.peek() == Tok::Comma {
                    self.advance();
                    args.push(self.parse_expr()?);
                }
            }
            self.expect_tok(Tok::RParen)?;
            if *self.peek() == Tok::Eq {
                // array write: name(indices...) = value
                self.advance();
                let value = self.parse_expr()?;
                let indices: Vec<Expr> = args;
                Ok(Stmt::ArrayWrite { name, indices, value })
            } else if *self.peek() == Tok::Dot && matches!(self.peek_at(1), Tok::Ident(_)) {
                // arr(i).Field = val
                self.advance(); // consume '.'
                let field = self.consume_ident()?.to_lowercase();
                if *self.peek() == Tok::Eq {
                    self.advance();
                    let value = self.parse_expr()?;
                    Ok(Stmt::ArrayRecordSet { name, indices: args, field, value })
                } else {
                    // Leave the trailing newline for the caller's own
                    // `eat_eol()` (the ident-statement dispatch fallback) —
                    // see the identical note on the EntireRow fallback above.
                    while !matches!(self.peek(), Tok::Newline | Tok::Eof) {
                        self.advance();
                    }
                    Ok(Stmt::Unsupported {
                        reason: format!(
                            "'{}(...).{}' read without assignment has no effect",
                            name, field
                        ),
                    })
                }
            } else {
                Ok(Stmt::CallSub { name, args })
            }
        } else if *self.peek() == Tok::Eq {
            self.advance();
            let value = self.parse_expr()?;
            Ok(Stmt::Assignment { var: name, value })
        } else if *self.peek() == Tok::Dot {
            // p.field = val  /  p.a.b = val  /  p.method (noop)
            self.advance(); // consume first '.'
            let field = self.consume_ident()?.to_lowercase();
            let mut fields = vec![field];
            // Collect additional .field segments (nested access)
            while *self.peek() == Tok::Dot && matches!(self.peek_at(1), Tok::Ident(_)) {
                self.advance(); // consume '.'
                fields.push(self.consume_ident()?.to_lowercase());
            }
            if *self.peek() == Tok::Eq {
                self.advance();
                let value = self.parse_expr()?;
                if fields.len() == 1 {
                    Ok(Stmt::RecordSet { var: name, field: fields.remove(0), value })
                } else {
                    Ok(Stmt::RecordSetNested { var: name, fields, value })
                }
            } else {
                // p.Method / property access without assignment — skip to
                // EOL (noop). Leave the trailing newline for the caller's
                // own `eat_eol()` — see the identical note above.
                while !matches!(self.peek(), Tok::Newline | Tok::Eof) {
                    self.advance();
                }
                Ok(Stmt::Unsupported {
                    reason: format!(
                        "'{}.{}' read without assignment has no effect",
                        name,
                        fields.join(".")
                    ),
                })
            }
        } else {
            // Bare ident — noop
            while !matches!(self.peek(), Tok::Newline | Tok::Eof) { self.advance(); }
            Ok(Stmt::Unsupported {
                reason: format!(
                    "'{}' as a bare statement (no Call keyword or parentheses) is not supported and was skipped",
                    name
                ),
            })
        }
    }

    // ── Expression parser ──────────────────────────────────────────────────────

    fn parse_arg_list(&mut self) -> Result<Vec<Expr>, String> {
        let mut args = vec![];
        if *self.peek() == Tok::RParen { return Ok(args); }
        args.push(self.parse_expr()?);
        while *self.peek() == Tok::Comma {
            self.advance();
            args.push(self.parse_expr()?);
        }
        Ok(args)
    }

    fn parse_expr(&mut self) -> Result<Expr, String> {
        self.parse_comparison()
    }

    fn parse_comparison(&mut self) -> Result<Expr, String> {
        let mut lhs = self.parse_additive()?;
        loop {
            let op = match self.peek() {
                Tok::Eq    => VbaBinOp::Eq,
                Tok::Ne    => VbaBinOp::Ne,
                Tok::Lt    => VbaBinOp::Lt,
                Tok::Le    => VbaBinOp::Le,
                Tok::Gt    => VbaBinOp::Gt,
                Tok::Ge    => VbaBinOp::Ge,
                _ => break,
            };
            self.advance();
            let rhs = self.parse_additive()?;
            lhs = Expr::BinOp { op, lhs: Box::new(lhs), rhs: Box::new(rhs) };
        }
        Ok(lhs)
    }

    fn parse_additive(&mut self) -> Result<Expr, String> {
        let mut lhs = self.parse_term()?;
        loop {
            let op = match self.peek() {
                Tok::Plus  => VbaBinOp::Add,
                Tok::Minus => VbaBinOp::Sub,
                Tok::Amp   => VbaBinOp::Concat,
                _ => break,
            };
            self.advance();
            let rhs = self.parse_term()?;
            lhs = Expr::BinOp { op, lhs: Box::new(lhs), rhs: Box::new(rhs) };
        }
        Ok(lhs)
    }

    fn parse_term(&mut self) -> Result<Expr, String> {
        let mut lhs = self.parse_factor()?;
        loop {
            let op = match self.peek() {
                Tok::Star  => VbaBinOp::Mul,
                Tok::Slash => VbaBinOp::Div,
                _ => break,
            };
            self.advance();
            let rhs = self.parse_factor()?;
            lhs = Expr::BinOp { op, lhs: Box::new(lhs), rhs: Box::new(rhs) };
        }
        Ok(lhs)
    }

    fn parse_factor(&mut self) -> Result<Expr, String> {
        if *self.peek() == Tok::Minus {
            self.advance();
            Ok(Expr::UnaryMinus(Box::new(self.parse_primary()?)))
        } else if self.is_ident("not") {
            self.advance();
            Ok(Expr::UnaryNot(Box::new(self.parse_primary()?)))
        } else {
            self.parse_primary()
        }
    }

    fn parse_primary(&mut self) -> Result<Expr, String> {
        match self.peek().clone() {
            Tok::LParen => {
                self.advance();
                let e = self.parse_comparison()?;
                self.expect_tok(Tok::RParen)?;
                Ok(e)
            }
            Tok::Int(n)  => { self.advance(); Ok(Expr::Integer(n)) }
            Tok::Float(f) => { self.advance(); Ok(Expr::Float(f)) }
            Tok::Str(s)  => { self.advance(); Ok(Expr::Str(s)) }
            Tok::Ident(ref s) => {
                let s = s.clone();
                match s.as_str() {
                    "true"  => { self.advance(); Ok(Expr::Bool(true)) }
                    "false" => { self.advance(); Ok(Expr::Bool(false)) }
                    "rows"  => self.parse_rows_cols_count("rows", Expr::RowsCount),
                    "columns" => self.parse_rows_cols_count("columns", Expr::ColsCount),
                    "cells" => self.parse_cells_expr(),
                    "range" => self.parse_range_expr(),
                    "worksheets" | "sheets" => self.parse_sheet_cell_read(),
                    "workbooks" => self.parse_workbook_qualified_read(),
                    "application" => self.parse_application_wsf_expr(),
                    "worksheetfunction" => self.parse_wsf_expr(),
                    _ => self.parse_ident_expr(),
                }
            }
            // ── `.Field` inside a `With p` block ──────────────────────────────
            Tok::Dot => {
                if let Some(var) = self.with_target.clone() {
                    self.advance(); // consume '.'
                    let field = self.consume_ident()?.to_lowercase();
                    let mut fields = vec![field];
                    while *self.peek() == Tok::Dot && matches!(self.peek_at(1), Tok::Ident(_)) {
                        self.advance(); // consume '.'
                        fields.push(self.consume_ident()?.to_lowercase());
                    }
                    if fields.len() == 1 {
                        Ok(Expr::RecordGet { var, field: fields.remove(0) })
                    } else {
                        Ok(Expr::RecordGetNested { var, fields })
                    }
                } else {
                    Err("Unexpected '.' outside With block".into())
                }
            }
            t => Err(format!("unexpected token in expression: {:?}", t)),
        }
    }

    fn parse_rows_cols_count(&mut self, kw: &str, expr: Expr) -> Result<Expr, String> {
        self.expect_ident(kw)?;
        self.expect_tok(Tok::Dot)?;
        self.expect_ident("count")?;
        Ok(expr)
    }

    fn parse_cells_expr(&mut self) -> Result<Expr, String> {
        self.expect_ident("cells")?;
        if *self.peek() == Tok::Dot {
            // cells.Find(...)
            self.advance();
            self.expect_ident("find")?;
            self.expect_tok(Tok::LParen)?;
            let mut what_expr = Expr::Str(String::new());
            // parse kwargs: What:=expr, ...
            while *self.peek() != Tok::RParen && *self.peek() != Tok::Eof {
                let kw_name = self.consume_ident()?;
                self.expect_tok(Tok::ColonEq)?;
                let val = self.parse_expr()?;
                if kw_name == "what" { what_expr = val; }
                if *self.peek() == Tok::Comma { self.advance(); }
            }
            self.expect_tok(Tok::RParen)?;
            self.expect_tok(Tok::Dot)?;
            let prop_kw = self.consume_ident()?;
            let find_row = prop_kw == "row";
            return Ok(Expr::CellsFind { what: Box::new(what_expr), find_row });
        }
        self.expect_tok(Tok::LParen)?;
        let row = self.parse_expr()?;
        self.expect_tok(Tok::Comma)?;
        let col = self.parse_expr()?;
        self.expect_tok(Tok::RParen)?;
        self.expect_tok(Tok::Dot)?;
        let prop = self.consume_ident()?;
        match prop.as_str() {
            "value" => Ok(Expr::CellRead { row: Box::new(row), col: Box::new(col) }),
            "end" => {
                self.expect_tok(Tok::LParen)?;
                let dir_str = self.consume_ident()?;
                let dir = match dir_str.as_str() {
                    "xlup"      => XlDir::Up,
                    "xldown"    => XlDir::Down,
                    "xltoleft"  => XlDir::Left,
                    "xltoright" => XlDir::Right,
                    other => return Err(format!("unknown xl_dir: {}", other)),
                };
                self.expect_tok(Tok::RParen)?;
                self.expect_tok(Tok::Dot)?;
                let end_prop = self.consume_ident()?;
                let prop = if end_prop == "row" { XlEndProp::Row } else { XlEndProp::Column };
                Ok(Expr::CellsEndProp {
                    row: Box::new(row), col: Box::new(col), dir, prop
                })
            }
            other => Err(format!("unexpected property after Cells(...): {}", other)),
        }
    }

    fn parse_range_expr(&mut self) -> Result<Expr, String> {
        self.expect_ident("range")?;
        self.expect_tok(Tok::LParen)?;
        let addr = self.consume_str()?.to_uppercase();
        self.expect_tok(Tok::RParen)?;
        // Without '.value': used as a Range object arg to WSF (e.g. WorksheetFunction.Sum(Range("A1:A3")))
        if *self.peek() != Tok::Dot {
            return Ok(Expr::FuncCall { name: "range".into(), args: vec![Expr::Str(addr)] });
        }
        self.advance(); // consume '.'
        let prop = self.consume_ident()?;
        match prop.as_str() {
            "value" => Ok(Expr::RangeRead { addr }),
            "offset" => {
                self.expect_tok(Tok::LParen)?;
                let row_off = self.parse_expr()?;
                self.expect_tok(Tok::Comma)?;
                let col_off = self.parse_expr()?;
                self.expect_tok(Tok::RParen)?;
                self.expect_tok(Tok::Dot)?;
                self.expect_ident("value")?;
                Ok(Expr::RangeOffsetRead {
                    addr,
                    row_off: Box::new(row_off),
                    col_off: Box::new(col_off),
                })
            }
            other => Err(format!("unexpected property after Range(...): {}", other)),
        }
    }

    /// Parses the `.Cells(r,c).Value` / `.Range(addr).Value` suffix shared
    /// by `Sheets(...)`/`Worksheets(...)` and `Workbooks(...).Worksheets(...)`
    /// read expressions.
    fn parse_sheet_property_read(&mut self, sheet: Expr) -> Result<Expr, String> {
        self.expect_tok(Tok::Dot)?;
        let prop = self.consume_ident()?;
        match prop.as_str() {
            "cells" => {
                self.expect_tok(Tok::LParen)?;
                let row = self.parse_expr()?;
                self.expect_tok(Tok::Comma)?;
                let col = self.parse_expr()?;
                self.expect_tok(Tok::RParen)?;
                self.expect_tok(Tok::Dot)?;
                self.expect_ident("value")?;
                Ok(Expr::SheetCellRead {
                    sheet: Box::new(sheet),
                    row: Box::new(row),
                    col: Box::new(col),
                })
            }
            "range" => {
                self.expect_tok(Tok::LParen)?;
                let addr = self.consume_str()?.to_uppercase();
                self.expect_tok(Tok::RParen)?;
                self.expect_tok(Tok::Dot)?;
                self.expect_ident("value")?;
                Ok(Expr::SheetRangeRead {
                    sheet: Box::new(sheet),
                    addr,
                })
            }
            other => Err(format!(
                "unexpected property after sheet reference: {}",
                other
            )),
        }
    }

    fn parse_sheet_cell_read(&mut self) -> Result<Expr, String> {
        self.consume_ident()?; // "worksheets" or "sheets"
        self.expect_tok(Tok::LParen)?;
        let sheet = self.parse_sheet_key()?;
        self.expect_tok(Tok::RParen)?;
        self.parse_sheet_property_read(sheet)
    }

    /// `Workbooks(workbook).Worksheets(sheet).Cells(...)`/`.Range(...)` read
    /// form — see `parse_workbook_qualified_stmt` for the write-side twin
    /// and the same "no real multi-workbook model" caveat.
    fn parse_workbook_qualified_read(&mut self) -> Result<Expr, String> {
        self.expect_ident("workbooks")?;
        self.expect_tok(Tok::LParen)?;
        let workbook = self.parse_sheet_key()?;
        self.expect_tok(Tok::RParen)?;
        self.expect_tok(Tok::Dot)?;
        if !(self.is_ident("worksheets") || self.is_ident("sheets")) {
            return Err(format!(
                "expected Worksheets(...)/Sheets(...) after Workbooks(...), got {:?}",
                self.peek()
            ));
        }
        self.advance();
        self.expect_tok(Tok::LParen)?;
        let sheet = self.parse_sheet_key()?;
        self.expect_tok(Tok::RParen)?;
        let qualified = Expr::WorkbookQualifiedSheet {
            workbook: Box::new(workbook),
            sheet: Box::new(sheet),
        };
        self.parse_sheet_property_read(qualified)
    }

    fn parse_application_wsf_expr(&mut self) -> Result<Expr, String> {
        self.expect_ident("application")?;
        self.expect_tok(Tok::Dot)?;
        self.expect_ident("worksheetfunction")?;
        self.parse_wsf_expr()
    }

    fn parse_wsf_expr(&mut self) -> Result<Expr, String> {
        // peek: already consumed "worksheetfunction" if coming from application path;
        // or still need to consume it
        if self.is_ident("worksheetfunction") { self.advance(); }
        self.expect_tok(Tok::Dot)?;
        let name = self.consume_ident()?;
        self.expect_tok(Tok::LParen)?;
        let args = self.parse_arg_list()?;
        self.expect_tok(Tok::RParen)?;
        Ok(Expr::FuncCall { name: format!("wsf_{}", name), args })
    }

    fn parse_ident_expr(&mut self) -> Result<Expr, String> {
        let name = self.consume_ident()?;
        if *self.peek() == Tok::LParen {
            self.advance();
            let args = self.parse_arg_list()?;
            self.expect_tok(Tok::RParen)?;
            // arr(i).Field — array element field read
            if *self.peek() == Tok::Dot && matches!(self.peek_at(1), Tok::Ident(_)) {
                self.advance(); // consume '.'
                let field = self.consume_ident()?.to_lowercase();
                return Ok(Expr::ArrayRecordGet { name, indices: args, field });
            }
            Ok(Expr::FuncCall { name, args })
        } else if *self.peek() == Tok::Dot {
            // p.field  or  p.a.b.c
            self.advance(); // consume '.'
            let field = self.consume_ident()?.to_lowercase();
            let mut fields = vec![field];
            while *self.peek() == Tok::Dot && matches!(self.peek_at(1), Tok::Ident(_)) {
                self.advance(); // consume '.'
                fields.push(self.consume_ident()?.to_lowercase());
            }
            if fields.len() == 1 {
                Ok(Expr::RecordGet { var: name, field: fields.remove(0) })
            } else {
                Ok(Expr::RecordGetNested { var: name, fields })
            }
        } else {
            Ok(Expr::Var(name))
        }
    }
}

// ── Utility ───────────────────────────────────────────────────────────────────

fn parse_cell_addr(addr: &str) -> Option<(u32, u32)> {
    let addr = addr.trim().to_uppercase();
    let alpha_end = addr.find(|c: char| c.is_ascii_digit())?;
    if alpha_end == 0 { return None; }
    let col = addr[..alpha_end]
        .chars()
        .fold(0u32, |acc, c| acc * 26 + (c as u32 - 'A' as u32 + 1));
    let row: u32 = addr[alpha_end..].parse().ok()?;
    Some((col, row))
}

// ── Public entry point ────────────────────────────────────────────────────────

pub fn parse(input: &str) -> Result<Program, String> {
    parse_with_span(input).map_err(|e| e.message)
}

/// A parse failure paired with the span of the token where it was detected.
pub struct ParseErrorWithSpan {
    pub message: String,
    pub span: SourceSpan,
}

/// Like `parse`, but on failure also reports where in the source the parser
/// gave up. Existing callers should keep using `parse` — this is additive,
/// for the `--json` CLI contract's location reporting.
pub fn parse_with_span(input: &str) -> Result<Program, ParseErrorWithSpan> {
    let (tokens, spans) = tokenize(input);
    let mut parser = Parser::new(tokens, spans);
    parser.parse_program().map_err(|message| {
        let span = parser.peek_span();
        ParseErrorWithSpan { message, span }
    })
}

// ── Multi-module resolution (Milestone B2) ────────────────────────────────────
//
// Pure functions over parsed `Program`s — no VM dependency. `modules` is a
// list of (module_name, Program) pairs; module names are expected to
// already be lowercased by the caller (mirroring the tokenizer's universal
// identifier-lowercasing convention used everywhere else).

/// Result of resolving a CLI entrypoint name against a set of modules.
pub enum EntrypointResolution<'a> {
    Found(&'a SubDef),
    NotFound,
}

/// Resolve a bare (`MySub`) or qualified (`Module1.MySub`) entrypoint name
/// against `modules`. Callers are expected to have already rejected
/// cross-module bare-name collisions (see `find_cross_module_sub_collisions`)
/// before calling this — a collision-free namespace means this only ever
/// has two outcomes, no "ambiguous" case.
pub fn resolve_entrypoint<'a>(
    modules: &'a [(String, Program)],
    entrypoint: &str,
) -> EntrypointResolution<'a> {
    let entrypoint = entrypoint.to_lowercase();
    if let Some((module_part, sub_part)) = entrypoint.rsplit_once('.') {
        for (name, prog) in modules {
            if name == module_part {
                return match prog.subs.iter().find(|s| s.name == sub_part) {
                    Some(sub) => EntrypointResolution::Found(sub),
                    None => EntrypointResolution::NotFound,
                };
            }
        }
        EntrypointResolution::NotFound
    } else {
        for (_, prog) in modules {
            if let Some(sub) = prog.subs.iter().find(|s| s.name == entrypoint) {
                return EntrypointResolution::Found(sub);
            }
        }
        EntrypointResolution::NotFound
    }
}

/// Bare Sub names that appear in 2+ modules, mapped to the list of module
/// names that declare them — the flat cross-module namespace can't
/// disambiguate these (own-module-first/Private VBA scoping isn't modeled),
/// so callers should reject the run rather than pick one silently.
pub fn find_cross_module_sub_collisions(
    modules: &[(String, Program)],
) -> Vec<(String, Vec<String>)> {
    let mut by_name: std::collections::HashMap<String, Vec<String>> =
        std::collections::HashMap::new();
    for (module_name, prog) in modules {
        for sub in &prog.subs {
            by_name
                .entry(sub.name.clone())
                .or_default()
                .push(module_name.clone());
        }
    }
    by_name
        .into_iter()
        .filter(|(_, mods)| mods.len() > 1)
        .collect()
}

/// Same as `find_cross_module_sub_collisions`, for bare Function names
/// (a separate namespace from Subs, same as within a single module today).
pub fn find_cross_module_func_collisions(
    modules: &[(String, Program)],
) -> Vec<(String, Vec<String>)> {
    let mut by_name: std::collections::HashMap<String, Vec<String>> =
        std::collections::HashMap::new();
    for (module_name, prog) in modules {
        for func in &prog.funcs {
            by_name
                .entry(func.name.clone())
                .or_default()
                .push(module_name.clone());
        }
    }
    by_name
        .into_iter()
        .filter(|(_, mods)| mods.len() > 1)
        .collect()
}

// ── Tests ─────────────────────────────────────────────────────────────────────

#[cfg(test)]
mod tests {
    use super::*;

    fn parse_body(code: &str) -> Vec<Stmt> {
        parse(code).unwrap().subs.into_iter().next().unwrap().body
            .into_iter().map(|s| s.stmt).collect()
    }

    #[test] fn test_empty_sub() {
        let prog = parse("Sub MySub()\nEnd Sub\n").unwrap();
        assert_eq!(prog.subs[0].name, "mysub");
        assert!(prog.subs[0].body.is_empty());
    }
    #[test] fn test_variable_assignment_integer() {
        let body = parse_body("Sub MySub()\n    a = 10\nEnd Sub\n");
        assert_eq!(body, vec![Stmt::Assignment { var: "a".into(), value: Expr::Integer(10) }]);
    }
    #[test] fn test_variable_assignment_float() {
        let body = parse_body("Sub MySub()\n    x = 3.14\nEnd Sub\n");
        assert_eq!(body, vec![Stmt::Assignment { var: "x".into(), value: Expr::Float(3.14) }]);
    }
    #[test] fn test_variable_assignment_string() {
        let body = parse_body("Sub MySub()\n    msg = \"hello\"\nEnd Sub\n");
        assert_eq!(body, vec![Stmt::Assignment { var: "msg".into(), value: Expr::Str("hello".into()) }]);
    }
    #[test] fn test_cell_write_integer() {
        let body = parse_body("Sub MySub()\n    Cells(1, 1).Value = 42\nEnd Sub\n");
        assert_eq!(body, vec![Stmt::CellWrite { row: Expr::Integer(1), col: Expr::Integer(1), value: Expr::Integer(42) }]);
    }
    #[test] fn test_cell_write_var_ref() {
        let body = parse_body("Sub MySub()\n    a = 10\n    Cells(1, 1).Value = a\nEnd Sub\n");
        assert_eq!(body[1], Stmt::CellWrite { row: Expr::Integer(1), col: Expr::Integer(1), value: Expr::Var("a".into()) });
    }
    #[test] fn test_case_insensitive_keywords() {
        let prog = parse("SUB MYSUB()\n    A = 10\n    CELLS(1, 1).VALUE = A\nEND SUB\n").unwrap();
        assert_eq!(prog.subs[0].name, "mysub");
    }
    #[test] fn test_comment_ignored() {
        let body = parse_body("Sub MySub()\n    ' comment\n    a = 10\nEnd Sub\n");
        assert_eq!(body, vec![Stmt::Assignment { var: "a".into(), value: Expr::Integer(10) }]);
    }
    #[test] fn test_multiple_subs() {
        let prog = parse("Sub First()\n    a = 1\nEnd Sub\n\nSub Second()\n    b = 2\nEnd Sub\n").unwrap();
        assert_eq!(prog.subs.len(), 2);
    }
    #[test] fn test_arithmetic_expr() {
        let body = parse_body("Sub MySub()\n    a = 1 + 2\nEnd Sub\n");
        assert_eq!(body, vec![Stmt::Assignment {
            var: "a".into(),
            value: Expr::BinOp { op: VbaBinOp::Add, lhs: Box::new(Expr::Integer(1)), rhs: Box::new(Expr::Integer(2)) },
        }]);
    }
    #[test] fn test_precedence_mul_over_add() {
        let body = parse_body("Sub MySub()\n    a = 1 + 2 * 3\nEnd Sub\n");
        assert_eq!(body, vec![Stmt::Assignment {
            var: "a".into(),
            value: Expr::BinOp {
                op: VbaBinOp::Add,
                lhs: Box::new(Expr::Integer(1)),
                rhs: Box::new(Expr::BinOp { op: VbaBinOp::Mul, lhs: Box::new(Expr::Integer(2)), rhs: Box::new(Expr::Integer(3)) }),
            },
        }]);
    }
    #[test] fn test_for_loop() {
        let body = parse_body("Sub MySub()\n    For i = 1 To 3\n        a = i\n    Next i\nEnd Sub\n");
        assert!(matches!(body[0], Stmt::For { .. }));
    }
    #[test] fn test_for_loop_step() {
        let body = parse_body("Sub MySub()\n    For i = 0 To 10 Step 2\n        a = i\n    Next i\nEnd Sub\n");
        if let Stmt::For { step, .. } = &body[0] { assert_eq!(*step, Some(Expr::Integer(2))); }
    }
    #[test] fn test_if_no_else() {
        let body = parse_body("Sub MySub()\n    If a > 0 Then\n        b = 1\n    End If\nEnd Sub\n");
        assert!(matches!(&body[0], Stmt::If { else_body, .. } if else_body.is_empty()));
    }
    #[test] fn test_if_with_else() {
        let body = parse_body("Sub MySub()\n    If a > 0 Then\n        b = 1\n    Else\n        b = 0\n    End If\nEnd Sub\n");
        if let Stmt::If { then_body, else_body, .. } = &body[0] {
            assert_eq!(then_body.len(), 1); assert_eq!(else_body.len(), 1);
        }
    }
    #[test] fn test_comparison_expr() {
        let body = parse_body("Sub MySub()\n    x = a > 5\nEnd Sub\n");
        assert_eq!(body, vec![Stmt::Assignment {
            var: "x".into(),
            value: Expr::BinOp { op: VbaBinOp::Gt, lhs: Box::new(Expr::Var("a".into())), rhs: Box::new(Expr::Integer(5)) },
        }]);
    }
    #[test] fn test_unary_minus() {
        let body = parse_body("Sub MySub()\n    a = -1\nEnd Sub\n");
        assert_eq!(body, vec![Stmt::Assignment { var: "a".into(), value: Expr::UnaryMinus(Box::new(Expr::Integer(1))) }]);
    }
    #[test] fn test_do_while_loop() {
        let body = parse_body("Sub MySub()\n    x = 0\n    Do While x < 3\n        x = x + 1\n    Loop\nEnd Sub\n");
        assert!(matches!(&body[1], Stmt::DoLoop { pre_cond: Some((false, _)), .. }));
    }
    #[test] fn test_do_until_loop() {
        let body = parse_body("Sub MySub()\n    x = 0\n    Do Until x >= 3\n        x = x + 1\n    Loop\nEnd Sub\n");
        assert!(matches!(&body[1], Stmt::DoLoop { pre_cond: Some((true, _)), .. }));
    }
    #[test] fn test_do_loop_while() {
        let body = parse_body("Sub MySub()\n    x = 0\n    Do\n        x = x + 1\n    Loop While x < 3\nEnd Sub\n");
        assert!(matches!(&body[1], Stmt::DoLoop { pre_cond: None, post_cond: Some((false, _)), .. }));
    }
    #[test] fn test_select_case() {
        let body = parse_body("Sub MySub()\n    Select Case x\n        Case 1\n            a = 1\n        Case 2, 3\n            a = 23\n        Case Else\n            a = 0\n    End Select\nEnd Sub\n");
        assert!(matches!(&body[0], Stmt::SelectCase { .. }));
        if let Stmt::SelectCase { cases, else_body, .. } = &body[0] {
            assert_eq!(cases.len(), 2); assert_eq!(else_body.len(), 1);
        }
    }
    #[test] fn test_dim_is_noop() {
        let body = parse_body("Sub MySub()\n    Dim x As Integer\n    x = 42\nEnd Sub\n");
        assert_eq!(body[0], Stmt::Dim);
    }
    #[test] fn test_with_block() {
        let body = parse_body("Sub MySub()\n    With Sheet1\n        .Cells(1, 1).Value = 99\n    End With\nEnd Sub\n");
        // `With Sheet1` is now parsed as WithRecord (plain identifier target).
        // Both WithRecord and With execute their body identically at runtime.
        let body_len = match &body[0] {
            Stmt::WithRecord { body, .. } => body.len(),
            Stmt::With { body }           => body.len(),
            _ => panic!("expected With or WithRecord"),
        };
        assert_eq!(body_len, 1);
    }

    #[test] fn test_with_udt_field_read_without_assignment_is_unsupported() {
        let body = parse_body("Sub MySub()\n    With p\n        .Field\n    End With\nEnd Sub\n");
        let inner = match &body[0] {
            Stmt::WithRecord { body, .. } => body,
            other => panic!("expected WithRecord, got {:?}", other),
        };
        assert_eq!(
            inner[0].stmt,
            Stmt::Unsupported {
                reason: "'p.field' read without assignment has no effect".to_string()
            }
        );
    }

    #[test] fn test_with_unrecognized_dot_method_is_unsupported() {
        let body = parse_body(
            "Sub MySub()\n    With Sheets(\"Sheet1\")\n        .Foo\n    End With\nEnd Sub\n",
        );
        let inner = match &body[0] {
            Stmt::WithSheet { body, .. } => body,
            other => panic!("expected WithSheet, got {:?}", other),
        };
        assert_eq!(
            inner[0].stmt,
            Stmt::Unsupported {
                reason: "With-block '.foo' is not implemented".to_string()
            }
        );
    }

    #[test] fn test_with_non_identifier_dotted_statement_is_unsupported() {
        // `.42` tokenizes as Dot, Int(42) — a non-identifier after the dot.
        let body = parse_body("Sub MySub()\n    With p\n        .42\n    End With\nEnd Sub\n");
        let inner = match &body[0] {
            Stmt::WithRecord { body, .. } => body,
            other => panic!("expected WithRecord, got {:?}", other),
        };
        assert_eq!(
            inner[0].stmt,
            Stmt::Unsupported {
                reason: "With-block dotted statement is not recognized and was skipped".to_string()
            }
        );
    }
    #[test] fn test_func_call_in_expr() {
        let body = parse_body("Sub MySub()\n    a = Len(\"hello\")\nEnd Sub\n");
        assert!(matches!(&body[0], Stmt::Assignment { value: Expr::FuncCall { name, .. }, .. } if name == "len"));
    }
    #[test] fn test_bool_literal() {
        let body = parse_body("Sub MySub()\n    a = True\n    b = False\nEnd Sub\n");
        assert_eq!(body[0], Stmt::Assignment { var: "a".into(), value: Expr::Bool(true) });
        assert_eq!(body[1], Stmt::Assignment { var: "b".into(), value: Expr::Bool(false) });
    }
    #[test] fn test_unary_not() {
        let body = parse_body("Sub MySub()\n    a = Not True\nEnd Sub\n");
        assert!(matches!(&body[0], Stmt::Assignment { value: Expr::UnaryNot(_), .. }));
    }
    #[test] fn test_dot_function_name() {
        // Handled in formula parser; VBA parser test
        let _ = parse("Sub MySub()\n    a = 1\nEnd Sub\n").unwrap();
    }
    #[test] fn test_elseif_chain() {
        let body = parse_body("Sub MySub()\n    If x > 10 Then\n        a = 1\n    ElseIf x > 5 Then\n        a = 2\n    Else\n        a = 3\n    End If\nEnd Sub\n");
        assert!(matches!(&body[0], Stmt::If { .. }));
        if let Stmt::If { else_body, .. } = &body[0] {
            assert!(matches!(else_body[0].stmt, Stmt::If { .. }));
        }
    }
    #[test] fn test_exit_for() {
        let body = parse_body("Sub MySub()\n    For i = 1 To 10\n        Exit For\n    Next i\nEnd Sub\n");
        if let Stmt::For { body, .. } = &body[0] { assert_eq!(body[0].stmt, Stmt::ExitFor); }
    }
    #[test] fn test_on_error_resume_next() {
        let body = parse_body("Sub MySub()\n    On Error Resume Next\n    a = 1\nEnd Sub\n");
        assert_eq!(body[0], Stmt::OnError { resume_next: true });
    }
    #[test] fn test_for_each() {
        let body = parse_body("Sub MySub()\n    For Each cell In Range(\"A1:A5\")\n        x = 1\n    Next cell\nEnd Sub\n");
        assert!(matches!(&body[0], Stmt::ForEach { var, .. } if var == "cell"));
    }
    #[test] fn test_call_stmt() {
        let body = parse_body("Sub MySub()\n    Call MySub2(1, 2)\nEnd Sub\n");
        assert!(matches!(&body[0], Stmt::CallSub { name, args } if name == "mysub2" && args.len() == 2));
    }
    #[test] fn test_func_def_parsed() {
        let prog = parse("Function Add(a, b)\n    Add = a + b\nEnd Function\n").unwrap();
        assert_eq!(prog.funcs.len(), 1);
        assert_eq!(prog.funcs[0].name, "add");
        assert_eq!(prog.funcs[0].params, vec!["a", "b"]);
    }
    #[test] fn test_sub_with_params() {
        let prog = parse("Sub Fill(startRow As Long, endRow As Long)\n    a = startRow\nEnd Sub\n").unwrap();
        assert_eq!(prog.subs[0].params, vec!["startrow", "endrow"]);
    }

    // ── Module-level declarations and access modifiers ─────────────────────────

    #[test] fn test_option_explicit_ignored() {
        let prog = parse("Option Explicit\nSub MySub()\n    a = 1\nEnd Sub\n").unwrap();
        assert_eq!(prog.subs[0].name, "mysub");
    }

    #[test] fn test_option_base_ignored() {
        let prog = parse("Option Base 1\nOption Explicit\nSub MySub()\n    a = 1\nEnd Sub\n").unwrap();
        assert_eq!(prog.subs.len(), 1);
    }

    #[test] fn test_public_sub() {
        let prog = parse("Public Sub MySub()\n    a = 1\nEnd Sub\n").unwrap();
        assert_eq!(prog.subs[0].name, "mysub");
        assert_eq!(prog.subs[0].body.len(), 1);
    }

    #[test] fn test_private_sub() {
        let prog = parse("Private Sub MySub()\n    a = 1\nEnd Sub\n").unwrap();
        assert_eq!(prog.subs[0].name, "mysub");
    }

    #[test] fn test_public_function() {
        let prog = parse("Public Function Add(a, b)\n    Add = a + b\nEnd Function\n").unwrap();
        assert_eq!(prog.funcs[0].name, "add");
    }

    #[test] fn test_private_function() {
        let prog = parse("Private Function Sq(x)\n    Sq = x * x\nEnd Function\n").unwrap();
        assert_eq!(prog.funcs[0].name, "sq");
    }

    #[test] fn test_module_level_dim_ignored() {
        // Module-level Dim (outside Sub) is skipped
        let prog = parse("Option Explicit\nDim counter As Long\nSub MySub()\n    counter = 1\nEnd Sub\n").unwrap();
        assert_eq!(prog.subs[0].name, "mysub");
    }

    #[test] fn test_module_level_const_with_modifier_is_flagged() {
        // `Public Const` never gets its value evaluated anywhere — a real
        // gap, unlike a plain declaration, so it's recorded for `check`.
        let prog = parse("Public Const MAX_RETRIES = 5\nSub MySub()\n    a = 1\nEnd Sub\n").unwrap();
        assert_eq!(prog.module_diagnostics.len(), 1);
        assert_eq!(
            prog.module_diagnostics[0].0,
            "Module-level 'Const' is not evaluated (module-level constants aren't supported outside a Sub/Function) and was skipped"
        );
    }

    #[test] fn test_module_level_bare_const_is_flagged() {
        let prog = parse("Const MAX_RETRIES = 5\nSub MySub()\n    a = 1\nEnd Sub\n").unwrap();
        assert_eq!(prog.module_diagnostics.len(), 1);
        assert_eq!(
            prog.module_diagnostics[0].0,
            "Module-level 'Const' is not evaluated (module-level constants aren't supported outside a Sub/Function) and was skipped"
        );
    }

    #[test] fn test_module_level_unrecognized_line_is_flagged() {
        let prog =
            parse("Declare Function Foo Lib \"x.dll\" ()\nSub MySub()\n    a = 1\nEnd Sub\n")
                .unwrap();
        assert_eq!(prog.module_diagnostics.len(), 1);
        assert_eq!(
            prog.module_diagnostics[0].0,
            "Module-level statement starting with 'declare' is not recognized and was skipped"
        );
    }

    #[test] fn test_module_level_plain_public_declaration_is_not_flagged() {
        // Group A parity with the Sub-level case: no separate module scope
        // exists (`Vm::variables` is one flat namespace), so a plain
        // declaration with no value is a harmless no-op, not a gap.
        let prog = parse("Public x As Long\nSub MySub()\n    x = 1\nEnd Sub\n").unwrap();
        assert!(prog.module_diagnostics.is_empty());
    }

    #[test] fn test_module_level_bare_dim_is_not_flagged() {
        let prog = parse("Dim counter As Long\nSub MySub()\n    counter = 1\nEnd Sub\n").unwrap();
        assert!(prog.module_diagnostics.is_empty());
    }

    #[test] fn test_attribute_ignored() {
        let prog = parse("Attribute VB_Name = \"Module1\"\nSub MySub()\n    a = 1\nEnd Sub\n").unwrap();
        assert_eq!(prog.subs.len(), 1);
    }

    #[test] fn test_vb_name_attribute_is_captured_as_module_name() {
        let prog = parse("Attribute VB_Name = \"Module1\"\nSub MySub()\n    a = 1\nEnd Sub\n").unwrap();
        assert_eq!(prog.module_name, Some("Module1".to_string()));
    }

    #[test] fn test_module_name_is_none_without_vb_name_attribute() {
        let prog = parse("Sub MySub()\n    a = 1\nEnd Sub\n").unwrap();
        assert_eq!(prog.module_name, None);
    }

    #[test] fn test_other_attribute_lines_still_ignored_alongside_vb_name() {
        let prog = parse(
            "Attribute VB_Name = \"Module1\"\nAttribute VB_GlobalNameSpace = False\nSub MySub()\n    a = 1\nEnd Sub\n",
        )
        .unwrap();
        assert_eq!(prog.module_name, Some("Module1".to_string()));
        assert_eq!(prog.subs.len(), 1);
    }

    // ── Debug.Print and statement-level modifiers ─────────────────────────────

    #[test] fn test_debug_print_noop() {
        let body = parse_body("Sub MySub()\n    Debug.Print \"hello\"\n    a = 1\nEnd Sub\n");
        // Debug.Print is a no-op; only the assignment remains
        assert_eq!(body.len(), 2); // Stmt::Unsupported (noop) + Assignment
        assert_eq!(body[1], Stmt::Assignment { var: "a".into(), value: Expr::Integer(1) });
    }

    #[test] fn test_debug_assert_noop() {
        let body = parse_body("Sub MySub()\n    Debug.Assert x > 0\n    a = 1\nEnd Sub\n");
        assert_eq!(body[1], Stmt::Assignment { var: "a".into(), value: Expr::Integer(1) });
    }

    // ── Stmt::Unsupported: unrecognized constructs preserve *why*, distinct
    // from Stmt::Dim's intentional no-op (see test_static_dim_inside_sub and
    // test_dim_is_noop below, which are untouched by this) ──────────────────

    #[test] fn test_debug_print_reason_is_specific() {
        let body = parse_body("Sub MySub()\n    Debug.Print \"hello\"\nEnd Sub\n");
        assert_eq!(
            body[0],
            Stmt::Unsupported {
                reason: "Debug.Print/Debug.Assert has no effect (no-op)".into()
            }
        );
    }

    #[test] fn test_entirerow_unknown_method_reason() {
        let body = parse_body("Sub MySub()\n    Range(\"A1\").EntireRow.Foo\nEnd Sub\n");
        assert_eq!(
            body[0],
            Stmt::Unsupported {
                reason: "EntireRow/EntireColumn.foo is not implemented".into()
            }
        );
    }

    #[test] fn test_range_unknown_property_reason() {
        let body = parse_body("Sub MySub()\n    Range(\"A1\").Hidden = True\nEnd Sub\n");
        assert_eq!(
            body[0],
            Stmt::Unsupported {
                reason: "Range property/method 'hidden' is not implemented".into()
            }
        );
    }

    #[test] fn test_sheets_unknown_method_reason() {
        let body = parse_body("Sub MySub()\n    Sheets.Foo\nEnd Sub\n");
        assert_eq!(
            body[0],
            Stmt::Unsupported { reason: "Sheets.foo is not implemented".into() }
        );
    }

    #[test] fn test_sheets_indexed_unknown_method_reason() {
        let body = parse_body("Sub MySub()\n    Sheets(\"Sheet1\").Foo\nEnd Sub\n");
        assert_eq!(
            body[0],
            Stmt::Unsupported { reason: "Sheets(...).foo is not implemented".into() }
        );
    }

    #[test] fn test_array_field_read_without_assignment_reason() {
        let body = parse_body("Sub MySub()\n    arr(0).Name\nEnd Sub\n");
        assert_eq!(
            body[0],
            Stmt::Unsupported {
                reason: "'arr(...).name' read without assignment has no effect".into()
            }
        );
    }

    #[test] fn test_record_field_read_without_assignment_reason() {
        let body = parse_body("Sub MySub()\n    p.Refresh\nEnd Sub\n");
        assert_eq!(
            body[0],
            Stmt::Unsupported {
                reason: "'p.refresh' read without assignment has no effect".into()
            }
        );
    }

    #[test] fn test_bare_ident_statement_reason() {
        let body = parse_body("Sub MySub()\n    Foo\nEnd Sub\n");
        assert_eq!(
            body[0],
            Stmt::Unsupported {
                reason: "'foo' as a bare statement (no Call keyword or parentheses) is not supported and was skipped".into()
            }
        );
    }

    #[test] fn test_static_dim_inside_sub() {
        let body = parse_body("Sub MySub()\n    Static counter As Long\n    counter = 1\nEnd Sub\n");
        assert_eq!(body[0], Stmt::Dim);
        assert_eq!(body[1], Stmt::Assignment { var: "counter".into(), value: Expr::Integer(1) });
    }

    #[test] fn test_mixed_module_preamble() {
        // Real-world VBA module preamble
        let code = concat!(
            "Option Explicit\n",
            "Option Base 1\n",
            "Attribute VB_Name = \"DataModule\"\n",
            "Private counter As Long\n",
            "\n",
            "Public Sub ProcessData()\n",
            "    counter = 0\n",
            "End Sub\n",
            "\n",
            "Private Function Helper(x)\n",
            "    Debug.Print \"helper called\"\n",
            "    Helper = x * 2\n",
            "End Function\n",
        );
        let prog = parse(code).unwrap();
        assert_eq!(prog.subs.len(), 1);
        assert_eq!(prog.subs[0].name, "processdata");
        assert_eq!(prog.funcs.len(), 1);
        assert_eq!(prog.funcs[0].name, "helper");
    }

    // ── On Error GoTo / labels / GoTo ─────────────────────────────────────────

    #[test] fn test_on_error_goto_label() {
        let body = parse_body(
            "Sub MySub()\n    On Error GoTo ErrH\n    a = 1\nErrH:\n    b = 2\nEnd Sub\n",
        );
        assert_eq!(body[0], Stmt::OnErrorGoTo("errh".into()));
        assert_eq!(body[1], Stmt::Assignment { var: "a".into(), value: Expr::Integer(1) });
        assert_eq!(body[2], Stmt::Label("errh".into()));
        assert_eq!(body[3], Stmt::Assignment { var: "b".into(), value: Expr::Integer(2) });
    }

    #[test] fn test_on_error_goto_zero() {
        let body = parse_body("Sub MySub()\n    On Error GoTo 0\n    a = 1\nEnd Sub\n");
        assert_eq!(body[0], Stmt::OnError { resume_next: false });
    }

    #[test] fn test_goto_stmt() {
        let body = parse_body("Sub MySub()\n    GoTo Done\nDone:\n    a = 1\nEnd Sub\n");
        assert_eq!(body[0], Stmt::GoTo("done".into()));
        assert_eq!(body[1], Stmt::Label("done".into()));
    }

    #[test] fn test_resume_next_stmt() {
        let body = parse_body("Sub MySub()\n    Resume Next\nEnd Sub\n");
        assert_eq!(body[0], Stmt::Resume { next: true });
    }

    // ── Multi-module resolution (Milestone B2) ─────────────────────────────

    fn module(name: &str, src: &str) -> (String, Program) {
        (name.to_string(), parse(src).unwrap())
    }

    #[test] fn resolve_entrypoint_bare_name_found() {
        let modules = vec![module("module1", "Sub Foo()\n    a = 1\nEnd Sub\n")];
        assert!(matches!(
            resolve_entrypoint(&modules, "Foo"),
            EntrypointResolution::Found(sub) if sub.name == "foo"
        ));
    }

    #[test] fn resolve_entrypoint_bare_name_not_found() {
        let modules = vec![module("module1", "Sub Foo()\n    a = 1\nEnd Sub\n")];
        assert!(matches!(
            resolve_entrypoint(&modules, "Bar"),
            EntrypointResolution::NotFound
        ));
    }

    #[test] fn resolve_entrypoint_bare_name_across_modules() {
        let modules = vec![
            module("module1", "Sub Foo()\n    a = 1\nEnd Sub\n"),
            module("module2", "Sub Bar()\n    a = 1\nEnd Sub\n"),
        ];
        assert!(matches!(
            resolve_entrypoint(&modules, "Bar"),
            EntrypointResolution::Found(sub) if sub.name == "bar"
        ));
    }

    #[test] fn resolve_entrypoint_qualified_found() {
        let modules = vec![
            module("module1", "Sub Foo()\n    a = 1\nEnd Sub\n"),
            module("module2", "Sub Foo()\n    a = 2\nEnd Sub\n"),
        ];
        assert!(matches!(
            resolve_entrypoint(&modules, "Module2.Foo"),
            EntrypointResolution::Found(sub) if sub.name == "foo"
        ));
    }

    #[test] fn resolve_entrypoint_qualified_unknown_module() {
        let modules = vec![module("module1", "Sub Foo()\n    a = 1\nEnd Sub\n")];
        assert!(matches!(
            resolve_entrypoint(&modules, "NoSuchModule.Foo"),
            EntrypointResolution::NotFound
        ));
    }

    #[test] fn resolve_entrypoint_qualified_unknown_sub_in_known_module() {
        let modules = vec![module("module1", "Sub Foo()\n    a = 1\nEnd Sub\n")];
        assert!(matches!(
            resolve_entrypoint(&modules, "Module1.Bar"),
            EntrypointResolution::NotFound
        ));
    }

    #[test] fn no_sub_collisions_across_disjoint_modules() {
        let modules = vec![
            module("module1", "Sub Foo()\n    a = 1\nEnd Sub\n"),
            module("module2", "Sub Bar()\n    a = 1\nEnd Sub\n"),
        ];
        assert!(find_cross_module_sub_collisions(&modules).is_empty());
    }

    #[test] fn one_sub_collision_across_two_modules() {
        let modules = vec![
            module("module1", "Sub Foo()\n    a = 1\nEnd Sub\n"),
            module("module2", "Sub Foo()\n    a = 2\nEnd Sub\n"),
        ];
        let collisions = find_cross_module_sub_collisions(&modules);
        assert_eq!(collisions.len(), 1);
        assert_eq!(collisions[0].0, "foo");
        let mut mods = collisions[0].1.clone();
        mods.sort();
        assert_eq!(mods, vec!["module1".to_string(), "module2".to_string()]);
    }

    #[test] fn sub_collision_spanning_three_modules() {
        let modules = vec![
            module("module1", "Sub Foo()\n    a = 1\nEnd Sub\n"),
            module("module2", "Sub Foo()\n    a = 2\nEnd Sub\n"),
            module("module3", "Sub Foo()\n    a = 3\nEnd Sub\n"),
        ];
        let collisions = find_cross_module_sub_collisions(&modules);
        assert_eq!(collisions.len(), 1);
        assert_eq!(collisions[0].1.len(), 3);
    }

    #[test] fn func_collisions_are_a_separate_namespace_from_subs() {
        // A Sub and a Function sharing a name across modules is not a
        // collision — Subs and Funcs are separate namespaces, as within a
        // single module today.
        let modules = vec![
            module("module1", "Sub Foo()\n    a = 1\nEnd Sub\n"),
            module("module2", "Function Foo()\n    Foo = 1\nEnd Function\n"),
        ];
        assert!(find_cross_module_sub_collisions(&modules).is_empty());
        assert!(find_cross_module_func_collisions(&modules).is_empty());
    }

    #[test] fn one_func_collision_across_two_modules() {
        let modules = vec![
            module("module1", "Function Foo()\n    Foo = 1\nEnd Function\n"),
            module("module2", "Function Foo()\n    Foo = 2\nEnd Function\n"),
        ];
        let collisions = find_cross_module_func_collisions(&modules);
        assert_eq!(collisions.len(), 1);
        assert_eq!(collisions[0].0, "foo");
    }
}
