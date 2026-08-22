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
    Backslash, // integer division (`\`)
    Caret,     // exponentiation (`^`)
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
            '\\' => { pos += 1; toks.push(Tok::Backslash); }
            '^' => { pos += 1; toks.push(Tok::Caret); }
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
                    // A literal wider than i64 — reuse the existing Float
                    // representation (the branch just above) rather than adding a
                    // new error path. f64::parse on an all-digit string never
                    // errors, so this can't panic the way the previous
                    // unconditional Int unwrap() did (found by fuzz_vba_parser).
                    match s.parse::<i64>() {
                        Ok(n) => toks.push(Tok::Int(n)),
                        Err(_) => toks.push(Tok::Float(s.parse().unwrap())),
                    }
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
    // No With-target state lives here any more. `With` targets used to be
    // resolved by a parse-time textual rewrite (a literal `Range("...")`
    // address or a bare UDT variable name substituted into every statement
    // of the body); they are now resolved once at runtime against the VM's
    // own With stack — see `ast::WithTarget` / `ast::WithMember`.
}

impl Parser {
    fn new(tokens: Vec<Tok>, spans: Vec<(u32, u32)>) -> Self {
        Parser { tokens, spans, pos: 0 }
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

    /// Consumes whatever ends the statement just parsed: a newline, EOF, or
    /// a `:` statement separator (real VBA's multi-statement-per-line form,
    /// `a = 1: b = 2`). A `:` is consumed but *no* newline is — the caller's
    /// own statement loop simply parses the next statement from the same
    /// line, so each one keeps its own `SourceSpan` exactly the way
    /// newline-separated statements already do. `:` reaches here as its own
    /// `Tok::Colon` (the tokenizer never produces one from inside a string
    /// literal, and `:=` is a separate `Tok::ColonEq`), so this can't
    /// misfire on `MsgBox "10:30"` or a `Destination:=` named argument. A
    /// label's own trailing `:` is also consumed here — `parse_ident_stmt`
    /// deliberately leaves it in place so `label1: a = 1` works.
    fn eat_stmt_end(&mut self) -> Result<(), String> {
        match self.peek() {
            Tok::Newline => { self.advance(); Ok(()) }
            Tok::Eof     => Ok(()),
            Tok::Colon   => {
                // `a = 1:: b = 2` — an empty statement between two colons is
                // legal VBA and means nothing; collapse a run of them.
                while *self.peek() == Tok::Colon { self.advance(); }
                Ok(())
            }
            t => Err(format!("expected newline, got {:?}", t)),
        }
    }

    // Consume to end of line (inclusive of the newline token). Deliberately
    // does NOT stop at a `:` — its callers skip a whole *line* (a block
    // header, an `Option` declaration), not one statement.
    fn skip_to_eol(&mut self) {
        loop {
            match self.peek() {
                Tok::Newline => { self.advance(); return; }
                Tok::Eof => return,
                _ => { self.advance(); }
            }
        }
    }

    /// Consume up to (not including) whatever ends the current statement —
    /// newline, EOF, or a `:` separator. The `skip_to_eol` sibling for the
    /// "this line isn't recognized, skip it" paths that run *inside*
    /// `parse_simple_stmt_no_eol`, so an unrecognized statement doesn't
    /// swallow the colon-separated statements after it on the same line.
    fn skip_to_stmt_end(&mut self) {
        while !matches!(self.peek(), Tok::Newline | Tok::Eof | Tok::Colon) { self.advance(); }
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
        let mut option_base: i64 = 0;
        while *self.peek() != Tok::Eof {
            // Module-level Option declarations → no-op, except `Option
            // Base <n>` (real VBA only allows a bare 0 or 1 literal here),
            // which sets the default lower bound for array declarators
            // that don't give an explicit `lo To hi`.
            if self.is_ident("option") {
                if self.is_ident_at(1, "base") {
                    self.advance(); // option
                    self.advance(); // base
                    if let Tok::Int(n) = self.peek().clone() {
                        self.advance();
                        option_base = n;
                    }
                }
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
            option_base,
        })
    }

    /// Parse a `Type Name ... End Type` block.
    fn parse_type_def(&mut self) -> Result<TypeDef, String> {
        self.expect_ident("type")?;
        let name = self.consume_ident()?.to_lowercase();
        self.eat_stmt_end()?;
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
        self.eat_stmt_end()?;
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
        // Optional return-type annotation: `Function f(...) As Integer`.
        // Not enforced anywhere (elixcee is dynamically typed at runtime,
        // same as every parameter's own `As <Type>` — see `parse_params`),
        // just consumed so it doesn't trip `eat_stmt_end()` below. Previously
        // unhandled entirely: `Function f(x As Integer) As Integer` failed
        // with "expected newline, got Ident(\"as\")" right here.
        if self.is_ident("as") {
            self.advance();
            self.consume_ident()?;
        }
        self.eat_stmt_end()?;
        let body = self.parse_stmts(|p| p.is_end_kw("function"))?;
        self.consume_end_kw("function")?;
        self.skip_nl();
        Ok(FuncDef { name, params, body })
    }

    fn parse_params(&mut self) -> Result<Vec<String>, String> {
        let mut params = vec![];
        while !matches!(self.peek(), Tok::RParen | Tok::Eof) {
            // `ByVal`/`ByRef` are recognized and discarded — elixcee's own
            // call semantics don't distinguish them (every call is
            // effectively by-value; no ByRef write-back to the caller's
            // variable is modeled), so skipping the keyword is enough to
            // keep `params` accurate without implementing ByRef semantics.
            // Without this, `consume_ident()` below would swallow the
            // keyword itself as a bogus extra parameter name (confirmed:
            // `Sub Foo(ByVal x As Integer)` used to silently parse as two
            // params, "byval" and "x", so `Foo(5)` bound 5 to the phantom
            // "byval" param and left `x` unbound) — exactly the kind of
            // wrong `params.len()` the new check-time argument-count check
            // depends on being accurate.
            if self.is_ident("byval") || self.is_ident("byref") {
                self.advance();
            }
            // `Optional`/`ParamArray` give a parameter variable arity
            // (default values, "any number of trailing args") that this VM
            // doesn't model at all — rejecting them outright, rather than
            // silently mis-consuming the keyword as a parameter name (the
            // same bug ByVal/ByRef had), keeps `params.len()` trustworthy
            // for every Sub/Function this parser does accept.
            if self.is_ident("optional") || self.is_ident("paramarray") {
                return Err(format!(
                    "parameter modifier '{}' is not supported",
                    if self.is_ident("optional") { "Optional" } else { "ParamArray" }
                ));
            }
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
        match self.peek() {
            Tok::Eof | Tok::Newline => return Ok(None),
            Tok::Ident(_) => {}
            // A bare `.member` statement inside a With body. Handled here,
            // in the general statement dispatch, so it is valid wherever a
            // statement is — including inside an `If`/`For`/`Do`/`Select
            // Case` nested in the body, which the old With-body-only
            // special case never reached.
            Tok::Dot => {
                let s = self.parse_with_dot_stmt()?;
                self.eat_stmt_end()?;
                return Ok(Some(s));
            }
            _ => return Err(format!("unexpected token starting statement: {:?}", self.peek())),
        }

        // A bare `name = ...` is always a plain assignment, even when `name`
        // collides with a block-construct keyword below (e.g. `do = 0`,
        // `select = 1`) — no VBA statement keyword's grammar puts `=`
        // immediately after itself, so this check must run before any
        // keyword dispatch, block-construct or not (moving the block-
        // construct checks ahead of this by mistake, during the refactor
        // that added `parse_simple_stmt_no_eol`, broke exactly this case —
        // caught by `prop_vba_assignment_parses` before it shipped).
        if *self.peek_at(1) == Tok::Eq {
            let s = self.parse_ident_stmt()?;
            self.eat_stmt_end()?;
            return Ok(Some(s));
        }

        // Block constructs self-manage their own terminator (End X / Next /
        // Loop / Wend, each already followed by `skip_nl()`) — they don't
        // go through `parse_simple_stmt_no_eol`/`eat_stmt_end()` at all, and
        // (being blocks, not single statements) don't belong inside a
        // single-line `If`'s Then/Else branch either.
        if self.is_ident("do")     { return Ok(Some(self.parse_do_loop()?)); }
        if self.is_ident("select") { return Ok(Some(self.parse_select_case()?)); }
        if self.is_ident("with")   { return Ok(Some(self.parse_with()?)); }
        if self.is_ident("for") && self.is_ident_at(1, "each") { return Ok(Some(self.parse_for_each()?)); }
        if self.is_ident("for")    { return Ok(Some(self.parse_for()?)); }
        if self.is_ident("if")     { return Ok(Some(self.parse_if()?)); }
        if self.is_ident("while")  { return Ok(Some(self.parse_while_wend()?)); }

        let s = self.parse_simple_stmt_no_eol()?;
        self.eat_stmt_end()?;
        Ok(Some(s))
    }

    /// Parses one "simple" statement — every statement `parse_stmt` handles
    /// except the block constructs above (Do/Select/With/For/If/While,
    /// which self-manage their own terminator and don't make sense inside
    /// a single-line `If`'s Then/Else branch anyway) — WITHOUT consuming
    /// the trailing newline/terminator. Shared by `parse_stmt` (which calls
    /// `eat_stmt_end()` right after) and `parse_single_line_if_branch` (which
    /// checks for `Else`/newline itself instead) — extracted specifically
    /// so a single-line `If`'s branches get the *same* statement coverage
    /// block-form VBA already has. Before this, `parse_single_line_if_branch`
    /// only recognized identifier-led statements via `parse_ident_stmt`, so
    /// `If cond Then Range("A1").Value = 1` mis-parsed as an array write to
    /// a variable literally named "range" (`parse_ident_stmt`'s own
    /// `name(args) = value` branch) instead of a Range statement — found by
    /// `compat/vba-semantics/` on this exact case, not by source audit.
    fn parse_simple_stmt_no_eol(&mut self) -> Result<Stmt, String> {
        let first = match self.peek() {
            Tok::Ident(s) => s.clone(),
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
            return self.parse_ident_stmt();
        }

        match first.as_str() {
            "exit"    => self.parse_exit(),
            "on" if self.is_ident_at(1, "error") => self.parse_on_error(),
            "goto"    => {
                self.advance();
                let label = self.consume_ident()?;
                Ok(Stmt::GoTo(label))
            }
            "resume"  => {
                self.advance();
                let next = if self.is_ident("next") { self.advance(); true } else { false };
                Ok(Stmt::Resume { next })
            }
            "set"     => self.parse_set(),
            "dim"     => self.parse_dim(),
            "redim"   => self.parse_redim(),
            "erase"   => self.parse_erase(),
            "const"   => self.parse_const(),
            "msgbox"  => self.parse_msgbox(),
            "call"    => self.parse_call_stmt(),
            "range"   => self.parse_range_stmt(),
            "cells"   => self.parse_cell_write_stmt(),
            "application" => self.parse_application_stmt(),
            "worksheetfunction" => self.parse_wsf_call_stmt(None),
            "worksheets" | "sheets" => self.parse_sheets_stmt(),
            "workbooks" => self.parse_workbook_qualified_stmt(),
            // `ActiveSheet.Range(...)`/`.Cells(...)`/`.Delete`/... (Milestone
            // B7c item 6) — same suffix grammar as `Worksheets(...)`, just
            // rooted at `Expr::ActiveSheetRef` instead of a parenthesized key.
            "activesheet" => {
                self.advance();
                self.parse_sheet_property_write(Expr::ActiveSheetRef)
            }
            // `ThisWorkbook.Worksheets(...)` / `ActiveWorkbook.Worksheets(...)`
            // (Milestone B7c item 6) — elixcee only ever has one workbook
            // loaded (see `Expr::WorkbookQualifiedSheet`'s doc), so these
            // are just the bare `Worksheets(...)`/`Sheets(...)` form with an
            // always-true qualifier: skip it and re-enter the same parse.
            "thisworkbook" | "activeworkbook"
                if self.is_ident_at(2, "worksheets") || self.is_ident_at(2, "sheets") =>
            {
                self.advance(); // 'thisworkbook' | 'activeworkbook'
                self.expect_tok(Tok::Dot)?;
                self.parse_sheets_stmt()
            }
            // Access/scope modifiers before Dim/Const inside a sub
            "public" | "private" | "static" | "friend" => {
                self.advance(); // consume modifier
                if self.is_ident("dim") {
                    self.parse_dim()
                } else if self.is_ident("const") {
                    self.parse_const()
                } else {
                    self.skip_to_stmt_end();
                    Ok(Stmt::Dim)
                }
            }
            // Debug.Print / Debug.Assert → no-op
            "debug" => {
                self.skip_to_stmt_end();
                Ok(Stmt::Unsupported {
                    reason: "Debug.Print/Debug.Assert has no effect (no-op)".to_string(),
                })
            }
            // `Err.Clear` / `Err.Raise ...` — guarded on the exact member
            // name (same precedent as the `thisworkbook`/`activeworkbook`
            // arms above), so a genuine user variable named `err` with an
            // unrelated UDT field (`err.code = 1`) still parses as ordinary
            // assignment/field access, untouched.
            "err" if self.is_ident_at(2, "clear") || self.is_ident_at(2, "raise") => {
                self.parse_err_stmt()
            }
            _ => self.parse_ident_stmt(),
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
        self.eat_stmt_end()?;
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
        self.eat_stmt_end()?;
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
        // `Tok::Colon` counts as end-of-header, not the start of a
        // single-line branch: `If x Then:` ... `End If` is the block form
        // with an empty statement after `Then`, not a one-liner.
        if !matches!(self.peek(), Tok::Newline | Tok::Eof | Tok::Colon) {
            return self.parse_single_line_if(condition);
        }
        self.eat_stmt_end()?;
        let then_body = self.parse_stmts(|p| {
            p.is_elseif() || p.is_ident("else") || p.is_end_kw("if")
        })?;
        let else_body = if self.is_elseif() {
            self.parse_elseif_chain()?
        } else if self.is_ident("else") {
            self.advance(); // "else"
            self.eat_stmt_end()?;
            self.parse_stmts(|p| p.is_end_kw("if"))?
        } else {
            vec![]
        };
        self.consume_end_kw("if")?;
        self.skip_nl();
        Ok(Stmt::If { condition, then_body, else_body })
    }

    /// `If cond Then stmt [Else stmt]` all on one line, no `End If` — real
    /// VBA grammar only allows a single optional `Else` here, never
    /// `ElseIf`. Entered once `parse_if` sees a non-newline token right
    /// after `Then`.
    fn parse_single_line_if(&mut self, condition: Expr) -> Result<Stmt, String> {
        let then_body = self.parse_single_line_if_branch_list()?;
        let else_body = if self.is_ident("else") {
            self.advance();
            self.parse_single_line_if_branch_list()?
        } else {
            vec![]
        };
        self.eat_stmt_end()?;
        Ok(Stmt::If { condition, then_body, else_body })
    }

    /// One single-line-`If` branch: a `:`-separated *list* of statements, not
    /// just one. Microsoft's own If...Then...Else reference documents
    /// `statements` as "One or more statements separated by colons; executed
    /// if condition is True", with the worked example
    /// `If A > 10 Then A = A + 1 : B = B + A : C = C + B`. The same applies
    /// to `elsestatements` after `Else` — a single-line `If` ends only at
    /// end-of-line, so `If x Then a = 1 Else b = 2: c = 3` puts *both*
    /// `b = 2` and `c = 3` in the Else branch.
    fn parse_single_line_if_branch_list(&mut self) -> Result<Vec<SpannedStmt>, String> {
        let mut out = vec![self.parse_single_line_if_branch()?];
        while *self.peek() == Tok::Colon {
            while *self.peek() == Tok::Colon { self.advance(); }
            if matches!(self.peek(), Tok::Newline | Tok::Eof) || self.is_ident("else") { break; }
            out.push(self.parse_single_line_if_branch()?);
        }
        Ok(out)
    }

    /// One inline statement for a single-line `If`/`Else` branch — reuses
    /// `parse_simple_stmt_no_eol`, the exact same dispatch block-form VBA's
    /// `parse_stmt` uses for everything except the block constructs (which
    /// don't make sense on a single line anyway), so a single-line `If`'s
    /// branches get full coverage: assignment, `Exit`/`GoTo`, `Set`/`Dim`,
    /// `Range`/`Cells`/`Application`/`Worksheets`/... — not just the
    /// identifier-led subset this used to be limited to. (That subset used
    /// to be this function's entire coverage, which silently mis-parsed
    /// `If cond Then Range("A1").Value = 1` as an array write to a variable
    /// literally named "range" — found by `compat/vba-semantics/`, not by
    /// source audit.) Anything still unrecognized (a token that isn't even
    /// an identifier) degrades to `Stmt::Unsupported` rather than a hard
    /// parse error, same precedent as `parse_set`'s unmodeled-target
    /// fallback.
    fn parse_single_line_if_branch(&mut self) -> Result<SpannedStmt, String> {
        let start = self.peek_span().start;
        let stmt = if matches!(self.peek(), Tok::Ident(_)) {
            self.parse_simple_stmt_no_eol()?
        } else if *self.peek() == Tok::Dot {
            // A bare `.member` branch, e.g. `With Range("A1"): If x Then .Value = 1`.
            // parse_stmt gained an equivalent Tok::Dot arm when the runtime With stack
            // replaced the old With-body-only special case (see parse_stmt's own comment),
            // but this single-line-If path checked only Tok::Ident and never got the same
            // update -- so `.Value = .Value + 1` inside a single-line If's Then/Else branch
            // silently degraded to Stmt::Unsupported (no error, but the assignment never
            // ran). Found by manual testing while integrating the With-stack work, not by
            // either subagent's own test suite.
            self.parse_with_dot_stmt()?
        } else {
            while !matches!(self.peek(), Tok::Newline | Tok::Eof | Tok::Colon)
                && !self.is_ident("else")
            {
                self.advance();
            }
            Stmt::Unsupported {
                reason: "single-line 'If ... Then ...' branch isn't a recognized statement shape"
                    .to_string(),
            }
        };
        let end = self.peek_span().start;
        Ok(SpannedStmt { stmt, span: SourceSpan { start, end } })
    }

    fn parse_elseif_chain(&mut self) -> Result<Vec<SpannedStmt>, String> {
        let start = self.peek_span().start;
        self.consume_elseif();
        let condition = self.parse_expr()?;
        self.expect_ident("then")?;
        self.eat_stmt_end()?;
        let then_body = self.parse_stmts(|p| {
            p.is_elseif() || p.is_ident("else") || p.is_end_kw("if")
        })?;
        let else_body = if self.is_elseif() {
            self.parse_elseif_chain()?
        } else if self.is_ident("else") {
            self.advance();
            self.eat_stmt_end()?;
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
        self.eat_stmt_end()?;
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
        self.eat_stmt_end()?;
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
        self.eat_stmt_end()?;
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
                self.eat_stmt_end()?;
                else_body = self.parse_stmts(|p| p.is_ident("case") || p.is_end_kw("select"))?;
            } else {
                let matches = self.parse_case_match_list()?;
                self.eat_stmt_end()?;
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

    /// `With <target> ... End With`. The target is captured as an
    /// unevaluated `WithTarget` and resolved **once at runtime**, on block
    /// entry — this used to be a parse-time rewrite that substituted a
    /// literal `Range("...")` address or a bare UDT variable name into every
    /// statement of the body, which is why a computed target (`With
    /// Cells(r, c)`) couldn't be expressed at all.
    ///
    /// `With Sheets("name")`/`Worksheets("name")` keeps its own
    /// `Stmt::WithSheet` variant: that path was already runtime-resolved
    /// (the VM swaps `active_sheet` around the body), and it also pushes the
    /// sheet onto the runtime With stack so bare `.member` statements
    /// resolve against it.
    fn parse_with(&mut self) -> Result<Stmt, String> {
        self.expect_ident("with")?;

        // ── Sheets/Worksheets("name") ─────────────────────────────────────────
        if self.is_ident("sheets") || self.is_ident("worksheets") {
            self.advance();
            if *self.peek() == Tok::LParen {
                self.advance();
                let name = self.consume_str()?.to_lowercase();
                self.expect_tok(Tok::RParen)?;
                self.eat_stmt_end()?;
                let body = self.parse_with_body()?;
                self.consume_end_kw("with")?;
                self.skip_nl();
                return Ok(Stmt::WithSheet { sheet_name: name, body });
            }
            self.skip_to_eol();
            let body = self.parse_with_body()?;
            self.consume_end_kw("with")?;
            self.skip_nl();
            return Ok(Stmt::With { target: WithTarget::Unmodeled, body });
        }

        // ── Cells(row, col) — a computed single-cell target ──────────────────
        if self.is_ident("cells") && *self.peek_at(1) == Tok::LParen {
            self.advance(); // 'cells'
            self.advance(); // '('
            let row = self.parse_expr()?;
            self.expect_tok(Tok::Comma)?;
            let col = self.parse_expr()?;
            self.expect_tok(Tok::RParen)?;
            return self.finish_with(WithTarget::Cells(row, col));
        }

        // ── Any object expression: Range("addr"), Union(...), <var>.Areas(n) ──
        if self.is_ident("range") || self.is_ident("union") {
            if let Some(obj) = self.parse_object_expr()? {
                return self.finish_with(WithTarget::Object(obj));
            }
            // Unrecognized shape — fall through to the no-op body below,
            // same leniency `parse_set` gives an unmodeled object target.
            self.skip_to_eol();
            let body = self.parse_with_body()?;
            self.consume_end_kw("with")?;
            self.skip_nl();
            return Ok(Stmt::With { target: WithTarget::Unmodeled, body });
        }

        // ── With <identifier> ────────────────────────────────────────────────
        // A Set-assigned Range/Worksheet object variable OR a UDT variable.
        // The parser can't tell which; the VM resolves it (see
        // `WithTarget::Var`). A trailing `.something` (e.g.
        // `With ws.Range("A1:B2")`) isn't modeled — skip to the no-op body
        // rather than mis-parsing it as a bare variable target.
        if matches!(self.peek(), Tok::Ident(_)) && matches!(self.peek_at(1), Tok::Newline | Tok::Eof | Tok::Colon) {
            let var = self.consume_ident()?.to_lowercase();
            return self.finish_with(WithTarget::Var(var));
        }

        // ── Generic / Application etc. — no-op body ───────────────────────────
        self.skip_to_eol();
        let body = self.parse_with_body()?;
        self.consume_end_kw("with")?;
        self.skip_nl();
        Ok(Stmt::With { target: WithTarget::Unmodeled, body })
    }

    /// Shared tail of every recognized `With` target: eat the header's
    /// terminator, parse the body, consume `End With`.
    fn finish_with(&mut self, target: WithTarget) -> Result<Stmt, String> {
        self.eat_stmt_end()?;
        let body = self.parse_with_body()?;
        self.consume_end_kw("with")?;
        self.skip_nl();
        Ok(Stmt::With { target, body })
    }

    /// A With body is an ordinary statement list — `parse_stmt` recognizes a
    /// leading `.` on its own now, so nothing here special-cases it and a
    /// bare `.member` works at any nesting depth inside the body.
    fn parse_with_body(&mut self) -> Result<Vec<SpannedStmt>, String> {
        self.parse_stmts(|p| p.is_end_kw("with"))
    }

    /// One statement beginning with a bare `.` — `.Value = 1`,
    /// `.Cells(r, c).Value = v`, `.Range("A1").Formula = f`, `.a.b = 2`, or
    /// a read-only `.Method` with no assignment (a no-op). Reached from
    /// `parse_stmt`, so it is valid wherever a statement is, including
    /// inside an `If`/`For`/`Do`/`Select Case` nested in a With body — the
    /// gap the old parse-time rewrite had.
    fn parse_with_dot_stmt(&mut self) -> Result<Stmt, String> {
        self.advance(); // consume '.'
        let head = match self.peek() {
            Tok::Ident(s) => s.clone(),
            _ => {
                self.skip_to_stmt_end();
                return Ok(Stmt::Unsupported {
                    reason: "With-block dotted statement is not recognized and was skipped"
                        .to_string(),
                });
            }
        };

        // `.Cells(r, c)...` / `.Range("addr")...` — a qualified member of the
        // With target, not a field of it. Guarded on an immediate `(` so a
        // genuine UDT field literally named "cells"/"range" still parses as
        // a field (the same caution `parse_ident_stmt` already takes).
        if (head == "cells" || head == "range") && *self.peek_at(1) == Tok::LParen {
            self.advance(); // head
            self.advance(); // '('
            let member = if head == "cells" {
                let row = self.parse_expr()?;
                self.expect_tok(Tok::Comma)?;
                let col = self.parse_expr()?;
                self.expect_tok(Tok::RParen)?;
                let fields = self.parse_dot_field_chain()?;
                WithMember::Cells { row: Box::new(row), col: Box::new(col), fields }
            } else {
                let addr = self.consume_str()?;
                self.expect_tok(Tok::RParen)?;
                let fields = self.parse_dot_field_chain()?;
                WithMember::Range { addr, fields }
            };
            return self.finish_with_dot(member, &head);
        }

        let mut fields = vec![self.consume_ident()?.to_lowercase()];
        while *self.peek() == Tok::Dot && matches!(self.peek_at(1), Tok::Ident(_)) {
            self.advance(); // '.'
            fields.push(self.consume_ident()?.to_lowercase());
        }
        let described = fields.join(".");
        self.finish_with_dot(WithMember::Fields(fields), &described)
    }

    /// Zero or more `.field` segments after a `.Cells(...)`/`.Range(...)`
    /// qualifier — `.Cells(1, 1).Value` yields `["value"]`.
    fn parse_dot_field_chain(&mut self) -> Result<Vec<String>, String> {
        let mut fields = vec![];
        while *self.peek() == Tok::Dot && matches!(self.peek_at(1), Tok::Ident(_)) {
            self.advance(); // '.'
            fields.push(self.consume_ident()?.to_lowercase());
        }
        Ok(fields)
    }

    /// Requires the `= <expr>` that turns a bare `.member` into a statement.
    /// Without one it's a property/method *read*, which has no effect —
    /// degraded to `Stmt::Unsupported` rather than a parse error, the same
    /// no-op-on-unmodeled-construct precedent used everywhere else here.
    fn finish_with_dot(&mut self, member: WithMember, described: &str) -> Result<Stmt, String> {
        if *self.peek() == Tok::Eq {
            self.advance();
            let value = self.parse_expr()?;
            return Ok(Stmt::WithDot { member, value });
        }
        self.skip_to_stmt_end();
        Ok(Stmt::Unsupported {
            reason: format!("With-block '.{}' read without assignment has no effect", described),
        })
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

    /// `Err.Clear` / `Err.Raise Number[, Source][, Description][, HelpFile]
    /// [, HelpContext]`. Real VBA's positional slots are (Number, Source,
    /// Description, HelpFile, HelpContext); any of the four after Number
    /// may be skipped with a bare comma (`Err.Raise 513, , "custom text"`
    /// — the idiomatic form for "no custom Source"), so this can't just
    /// split on commas positionally without risking reading a supplied
    /// Description as Source.
    fn parse_err_stmt(&mut self) -> Result<Stmt, String> {
        self.expect_ident("err")?;
        self.expect_tok(Tok::Dot)?;
        if self.is_ident("clear") {
            self.advance();
            return Ok(Stmt::ErrClear);
        }
        self.expect_ident("raise")?;
        let number = self.parse_expr()?;
        // Source, Description, HelpFile, HelpContext, in that fixed order —
        // each slot only advances past its own leading comma, so a bare
        // comma correctly skips exactly one slot.
        let mut rest: [Option<Expr>; 4] = [None, None, None, None];
        for slot in rest.iter_mut() {
            if *self.peek() != Tok::Comma { break; }
            self.advance();
            if *self.peek() != Tok::Comma && !self.is_stmt_end() {
                *slot = Some(self.parse_expr()?);
            }
        }
        let [source, description, help_file, help_context] = rest;
        Ok(Stmt::ErrRaise { number, source, description, help_file, help_context })
    }

    fn is_stmt_end(&self) -> bool {
        matches!(self.peek(), Tok::Newline | Tok::Eof | Tok::Colon)
    }

    // ── Simple statements ──────────────────────────────────────────────────────

    /// Known VBA built-in type names that do NOT correspond to a user-defined type.
    fn is_vba_builtin_type(name: &str) -> bool {
        matches!(name, "integer" | "long" | "longlong" | "single" | "double" | "currency"
            | "boolean" | "string" | "date" | "object" | "variant" | "byte" | "decimal")
    }

    /// `Dim <decl> [, <decl> ...]` — each declarator is parsed by
    /// `parse_dim_declarator`; a single declarator returns exactly what it
    /// used to (no `DimMulti` wrapper, so existing single-declarator tests
    /// stay byte-for-byte unchanged), two or more wrap into `DimMulti`.
    fn parse_dim(&mut self) -> Result<Stmt, String> {
        self.expect_ident("dim")?;
        let mut decls = vec![self.parse_dim_declarator()?];
        while *self.peek() == Tok::Comma {
            self.advance();
            decls.push(self.parse_dim_declarator()?);
        }
        if decls.len() == 1 {
            Ok(decls.pop().unwrap())
        } else {
            Ok(Stmt::DimMulti(decls))
        }
    }

    /// One `Dim`/`Public`/`Private`/`Static` declarator: `name [(sizes)] [As
    /// TypeName]`. Consumes exactly the declarator's own tokens — trailing
    /// `, nextDecl` is left for the caller's comma loop, and the terminating
    /// newline is left for the statement dispatcher's `eat_stmt_end()`.
    fn parse_dim_declarator(&mut self) -> Result<Stmt, String> {
        // dim_array_decl: ident (
        if matches!(self.peek(), Tok::Ident(_)) && *self.peek_at(1) == Tok::LParen {
            let name = self.consume_ident()?;
            self.advance(); // (
            // `Dim arr()` — empty parens, a dynamic array sized later by
            // `ReDim` — has no dimensions to parse at all.
            let mut sizes = Vec::new();
            if *self.peek() != Tok::RParen {
                sizes.push(self.parse_array_dim()?);
                while *self.peek() == Tok::Comma {
                    self.advance();
                    sizes.push(self.parse_array_dim()?);
                }
            }
            self.expect_tok(Tok::RParen)?;
            if self.is_ident("as") {
                self.advance();
                let type_name = self.consume_ident()?.to_lowercase();
                if !Self::is_vba_builtin_type(&type_name) {
                    // DimArrayRecord doesn't track a lower bound (no case
                    // needs it) — only the upper-bound expression carries
                    // over, same as before this method gained `lo To hi`.
                    let upper_only = sizes.into_iter().map(|d| d.upper).collect();
                    return Ok(Stmt::DimArrayRecord { name, sizes: upper_only, type_name });
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
            // Built-in type or bare Dim → no-op. Consume any trailing
            // per-declarator syntax this grammar doesn't model (e.g. `As
            // String * 10`'s fixed-length-string suffix) up to the next
            // declarator-separating comma, so it reaches the outer comma
            // loop instead of hard-failing at `eat_stmt_end()` — the
            // single-declarator form had this same tolerance (bounded by
            // EOL instead of comma) before the comma loop existed.
            while !matches!(self.peek(), Tok::Comma | Tok::Newline | Tok::Eof | Tok::Colon) { self.advance(); }
            Ok(Stmt::DimBare { var })
        } else {
            // Not even an identifier here (malformed `Dim`) — same
            // permissive no-op the pre-comma-loop parser gave any
            // unparseable `Dim` line, just bounded by comma now so a
            // trailing `, nextDecl` still reaches the outer loop.
            while !matches!(self.peek(), Tok::Comma | Tok::Newline | Tok::Eof | Tok::Colon) { self.advance(); }
            Ok(Stmt::Dim)
        }
    }

    /// One dimension inside a `Dim`/`ReDim` size list: a bare upper-bound
    /// expression (`5`), or an explicit `lo To hi` pair (`2 To 8`).
    fn parse_array_dim(&mut self) -> Result<ArrayDim, String> {
        let first = self.parse_expr()?;
        if self.is_ident("to") {
            self.advance();
            let upper = self.parse_expr()?;
            Ok(ArrayDim { lower: Some(first), upper })
        } else {
            Ok(ArrayDim { lower: None, upper: first })
        }
    }

    // ── Object references (Milestone B7c) ──────────────────────────────────────

    /// `Set <var> = <rhs>` — dispatched from `parse_stmt` like `Dim`/`Const`.
    /// If `<rhs>` isn't a shape `parse_object_expr` recognizes (e.g. `Set d =
    /// CreateObject(...)`, `Set rng = Nothing`, `Set ws = ActiveWorkbook.
    /// Sheets(1)`), the whole statement degrades to `Stmt::Unsupported`
    /// rather than a hard parse error — same precedent as `Stmt::Dim`/the
    /// generic `.Method` no-op in `parse_ident_stmt`: an otherwise-working
    /// macro that happens to use an unmodeled `Set` target should still run.
    fn parse_set(&mut self) -> Result<Stmt, String> {
        self.expect_ident("set")?;
        let var = self.consume_ident()?;
        self.expect_tok(Tok::Eq)?;
        match self.parse_object_expr()? {
            Some(value) => Ok(Stmt::Set { var, value }),
            None => {
                // NOT `skip_to_eol()` — that also consumes the trailing
                // newline, and `parse_stmt`'s "set" dispatch arm already
                // calls `eat_stmt_end()` after this returns (same double-
                // consumption pitfall `parse_ident_stmt`'s "bare ident"
                // no-op branch avoids the same way).
                self.skip_to_stmt_end();
                Ok(Stmt::Unsupported {
                    reason: format!(
                        "'Set {} = ...' targets an unmodeled object expression and was skipped",
                        var
                    ),
                })
            }
        }
    }

    /// Parses one reference-typed expression: `Range("...")`, an existing
    /// object variable, or `Union(...)`/`.Areas(n)`/`.SpecialCells(...)`
    /// applied to either of those. Returns `Ok(None)` (not `Err`) for
    /// anything it doesn't recognize — callers must still consume/skip the
    /// remainder of the statement themselves (see `parse_set`); partial
    /// speculative token consumption before bailing is harmless since every
    /// caller falls back to `skip_to_eol` on `None`.
    fn parse_object_expr(&mut self) -> Result<Option<ObjectExpr>, String> {
        match self.peek().clone() {
            Tok::Ident(ref s) if s == "range" => {
                self.advance();
                self.expect_tok(Tok::LParen)?;
                let addr = self.consume_str()?.to_uppercase();
                self.expect_tok(Tok::RParen)?;
                self.parse_object_suffix(ObjectExpr::RangeLit(addr))
            }
            Tok::Ident(ref s) if s == "union" => {
                self.advance();
                self.expect_tok(Tok::LParen)?;
                let mut parts = vec![];
                loop {
                    match self.parse_object_expr()? {
                        Some(p) => parts.push(p),
                        None => return Ok(None),
                    }
                    if *self.peek() == Tok::Comma { self.advance(); } else { break; }
                }
                self.expect_tok(Tok::RParen)?;
                self.parse_object_suffix(ObjectExpr::Union(parts))
            }
            Tok::Ident(name) => {
                // A bare identifier in object position: an existing object
                // variable (`Set b = a`, `Set b = a.Areas(1)`). Anything
                // followed by '(' here (a function call we don't model,
                // e.g. `CreateObject(...)`) is left unrecognized.
                self.advance();
                if *self.peek() == Tok::LParen {
                    Ok(None)
                } else {
                    self.parse_object_suffix(ObjectExpr::Var(name))
                }
            }
            _ => Ok(None),
        }
    }

    /// Chains zero or more `.Areas(n)` / `.SpecialCells(xlCellTypeVisible)`
    /// suffixes onto `base`. Any other `.property` (notably `.Value`, which
    /// belongs to a different grammar entirely — see `Stmt::RecordSet`'s
    /// object-variable special case in the VM) is left unconsumed: this
    /// function only ever advances past a `.` it's about to fully parse.
    fn parse_object_suffix(&mut self, base: ObjectExpr) -> Result<Option<ObjectExpr>, String> {
        let mut cur = base;
        loop {
            if *self.peek() != Tok::Dot { break; }
            let is_areas = self.is_ident_at(1, "areas");
            let is_special = self.is_ident_at(1, "specialcells");
            if !is_areas && !is_special { break; }
            self.advance(); // '.'
            self.advance(); // 'areas' | 'specialcells'
            if is_areas {
                self.expect_tok(Tok::LParen)?;
                let index = self.parse_expr()?;
                self.expect_tok(Tok::RParen)?;
                cur = ObjectExpr::Area(Box::new(cur), Box::new(index));
            } else {
                self.expect_tok(Tok::LParen)?;
                let recognized = match self.peek().clone() {
                    Tok::Ident(ref s) if s == "xlcelltypevisible" => { self.advance(); true }
                    Tok::Int(12) => { self.advance(); true }
                    _ => false,
                };
                if !recognized {
                    // Unrecognized SpecialCells type — consume through the
                    // matching ')' so the caller's eventual `skip_to_eol`
                    // still lands cleanly, then bail.
                    while *self.peek() != Tok::RParen && *self.peek() != Tok::Eof { self.advance(); }
                    if *self.peek() == Tok::RParen { self.advance(); }
                    return Ok(None);
                }
                self.expect_tok(Tok::RParen)?;
                cur = ObjectExpr::SpecialCellsVisible(Box::new(cur));
            }
        }
        Ok(Some(cur))
    }

    /// `Erase <name>` — real VBA's comma-separated `Erase a, b` form isn't
    /// parsed (no case needs it).
    fn parse_erase(&mut self) -> Result<Stmt, String> {
        self.expect_ident("erase")?;
        let name = self.consume_ident()?;
        Ok(Stmt::Erase { name })
    }

    fn parse_redim(&mut self) -> Result<Stmt, String> {
        self.expect_ident("redim")?;
        let preserve = if self.is_ident("preserve") { self.advance(); true } else { false };
        let name = self.consume_ident()?;
        self.expect_tok(Tok::LParen)?;
        let mut sizes = vec![self.parse_array_dim()?];
        while *self.peek() == Tok::Comma {
            self.advance();
            sizes.push(self.parse_array_dim()?);
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
        // Real VBA's `Call` grammar is `Call name [(argumentlist)]` — the
        // parens are optional, required only when passing arguments. Found
        // missing (`Call Foo` with no args was a syntax error, while
        // `Call Foo()` and bare `Foo` both already worked) during 0.7.0
        // release verification; unrelated to that round's own changes —
        // this function hasn't been touched since the 2026-06-21
        // hand-written-parser rewrite.
        let args = if *self.peek() == Tok::LParen {
            self.advance();
            let args = self.parse_arg_list()?;
            self.expect_tok(Tok::RParen)?;
            args
        } else {
            Vec::new()
        };
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
                        // `eat_stmt_end()` (the "range" dispatch arm) — unlike
                        // `skip_to_eol()`, which would consume it too and
                        // cause a spurious "expected newline" error when
                        // this is the last statement before End Sub.
                        self.skip_to_stmt_end();
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
                self.skip_to_stmt_end();
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
                self.skip_to_stmt_end();
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
                self.skip_to_stmt_end();
                return Ok(Stmt::SheetsAdd);
            }
            // Leave the trailing newline for the caller's own `eat_stmt_end()`
            // (the "worksheets"/"sheets" dispatch arm) — see the identical
            // note on the EntireRow/EntireColumn fallback above.
            self.skip_to_stmt_end();
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
            self.skip_to_stmt_end();
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
        // Label: "ErrHandler:". The `:` is deliberately NOT consumed here —
        // `eat_stmt_end` takes it as the statement terminator, which is what
        // makes real VBA's `label1: a = 1` (a label and a statement on one
        // line) parse as two statements rather than a parse error.
        if *self.peek() == Tok::Colon {
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
                    // `eat_stmt_end()` (the ident-statement dispatch fallback) —
                    // see the identical note on the EntireRow fallback above.
                    self.skip_to_stmt_end();
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
        } else if *self.peek() == Tok::Dot && self.is_ident_at(1, "copy") {
            // <var>.Copy [Destination:=Range(addr)] — the object-variable
            // sibling of `Range("addr").Copy` (see `parse_range_stmt`'s
            // "copy" arm, same `Destination:=` grammar). Checked ahead of
            // the generic `.field`/`.field = value` branch below so `.Copy`
            // never gets swallowed into a bogus `RecordSet`/`Unsupported`.
            self.advance(); // '.'
            self.advance(); // 'copy'
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
            Ok(Stmt::RangeObjectCopy { var: name, dst })
        } else if *self.peek() == Tok::Dot
            && (self.is_ident_at(1, "range") || self.is_ident_at(1, "cells"))
            && *self.peek_at(2) == Tok::LParen
        {
            // <var>.Range(addr).Value/Formula = val / <var>.Cells(r,c).Value
            // = val (Phase 2C item 7) — the object-variable sibling of
            // `Sheets(...).Range(...)`/`.Cells(...)` (see
            // `parse_sheet_property_write`), for a `<var>` a `Set var =
            // ActiveSheet` assigned a Worksheet reference to.
            // `Expr::ObjectVarSheet` resolves against `Vm::object_variables`
            // at *runtime* — the parser can't know `<var>`'s type here, same
            // situation the `.Copy` branch above already accepts. Checked
            // ahead of the generic `.field = value` branch below (guarded
            // on an immediate `(` so a genuine UDT field literally named
            // "range"/"cells" — vanishingly unlikely, but same caution
            // `WithRecord`'s own `s != "range" && s != "cells"` guard
            // takes) still falls through there instead).
            self.parse_sheet_property_write(Expr::ObjectVarSheet(name))
        } else if *self.peek() == Tok::Dot
            && (self.is_ident_at(1, "worksheets") || self.is_ident_at(1, "sheets"))
            && *self.peek_at(2) == Tok::LParen
        {
            // <var>.Worksheets(...)/.Sheets(...) (Phase 2C item 8) — the
            // object-variable sibling of `ThisWorkbook.Worksheets(...)`/
            // `ActiveWorkbook.Worksheets(...)` (see `parse_stmt`'s
            // "thisworkbook"|"activeworkbook" arm), for a `<var>` a `Set
            // var = ThisWorkbook` assigned a Workbook reference to. elixcee
            // only ever has one workbook loaded, so — same as those two
            // keywords — this just skips the qualifier and re-enters the
            // plain `Worksheets(...)/Sheets(...)` grammar; nothing here (or
            // in the VM — see `ObjectRef::Workbook`) checks that `<var>`
            // actually holds a Workbook reference. Guarded on an immediate
            // `(` — same reason as the `.Range(`/`.Cells(` branch above —
            // so a paren-less `wb.Worksheets.Count` (a real, if unmodeled,
            // VBA read) or a UDT field literally named "worksheets"/"sheets"
            // still falls through to the generic path below instead of a
            // hard `expected LParen` error.
            self.advance(); // '.'
            self.parse_sheets_stmt()
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
                // own `eat_stmt_end()` — see the identical note above.
                self.skip_to_stmt_end();
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
            self.skip_to_stmt_end();
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

    // Precedence climbing, lowest (outermost/loosest-binding) to highest
    // (innermost/tightest-binding), matching real VBA's documented operator
    // precedence exactly:
    //   Xor < Or < And < Not < comparison < & < (+ -) < Mod < \ < (* /)
    //   < unary - < ^
    // Every tier below is a thin left-associative "climb one level, loop on
    // same-tier operators" wrapper, same shape as the pre-existing
    // parse_comparison/parse_additive/parse_term this replaces — Xor/Or/And
    // are just three more copies of that shape at looser precedence, and
    // Mod/\ two more copies slotted between (+ -) and (* /).
    fn parse_expr(&mut self) -> Result<Expr, String> {
        self.parse_xor()
    }

    fn parse_xor(&mut self) -> Result<Expr, String> {
        let mut lhs = self.parse_or()?;
        while self.is_ident("xor") {
            self.advance();
            let rhs = self.parse_or()?;
            lhs = Expr::BinOp { op: VbaBinOp::Xor, lhs: Box::new(lhs), rhs: Box::new(rhs) };
        }
        Ok(lhs)
    }

    fn parse_or(&mut self) -> Result<Expr, String> {
        let mut lhs = self.parse_and()?;
        while self.is_ident("or") {
            self.advance();
            let rhs = self.parse_and()?;
            lhs = Expr::BinOp { op: VbaBinOp::Or, lhs: Box::new(lhs), rhs: Box::new(rhs) };
        }
        Ok(lhs)
    }

    fn parse_and(&mut self) -> Result<Expr, String> {
        let mut lhs = self.parse_not_level()?;
        while self.is_ident("and") {
            self.advance();
            let rhs = self.parse_not_level()?;
            lhs = Expr::BinOp { op: VbaBinOp::And, lhs: Box::new(lhs), rhs: Box::new(rhs) };
        }
        Ok(lhs)
    }

    // `Not` is a prefix operator, not an infix one — real VBA has no `a Not
    // b` form — but it sits in the middle of the precedence table (looser
    // than comparison, tighter than And/Or/Xor), so `Not a And b` must parse
    // as `(Not a) And b`, and `Not a = b` as `Not (a = b)`. Recurses into
    // itself (not straight to parse_comparison) so a stacked `Not Not x`
    // still parses, same allowance the pre-existing unary-minus chain makes
    // for `- -x`.
    fn parse_not_level(&mut self) -> Result<Expr, String> {
        if self.is_ident("not") {
            self.advance();
            Ok(Expr::UnaryNot(Box::new(self.parse_not_level()?)))
        } else {
            self.parse_comparison()
        }
    }

    fn parse_comparison(&mut self) -> Result<Expr, String> {
        let mut lhs = self.parse_concat()?;
        // `<var> Is Nothing` — an object-identity comparison, at the same
        // precedence tier as the value comparisons below (Microsoft's own
        // comparison-operators reference lists `result = object1 Is object2`
        // alongside them). Only the `Is Nothing` right-hand side is modeled;
        // a general `a Is b` is left unparsed rather than guessed at, and
        // `Case Is > 5` never reaches here (parse_case_match consumes its own
        // `Is` before ever calling parse_expr).
        if self.is_ident("is") && self.is_ident_at(1, "nothing")
            && let Expr::Var(name) = &lhs
        {
            let name = name.clone();
            self.advance(); // 'is'
            self.advance(); // 'nothing'
            lhs = Expr::IsNothing(name);
        }
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
            let rhs = self.parse_concat()?;
            lhs = Expr::BinOp { op, lhs: Box::new(lhs), rhs: Box::new(rhs) };
        }
        Ok(lhs)
    }

    // `&` (string concat) binds tighter than comparison but looser than
    // `+`/`-` — e.g. `"x" & 1 + 2` is `"x" & (1 + 2)` = "x3", not `("x" & 1)
    // + 2`. Previously folded into the same tier as `+`/`-` (equal
    // precedence, left-to-right); split out here to match real VBA.
    fn parse_concat(&mut self) -> Result<Expr, String> {
        let mut lhs = self.parse_additive()?;
        while *self.peek() == Tok::Amp {
            self.advance();
            let rhs = self.parse_additive()?;
            lhs = Expr::BinOp { op: VbaBinOp::Concat, lhs: Box::new(lhs), rhs: Box::new(rhs) };
        }
        Ok(lhs)
    }

    fn parse_additive(&mut self) -> Result<Expr, String> {
        let mut lhs = self.parse_modop()?;
        loop {
            let op = match self.peek() {
                Tok::Plus  => VbaBinOp::Add,
                Tok::Minus => VbaBinOp::Sub,
                _ => break,
            };
            self.advance();
            let rhs = self.parse_modop()?;
            lhs = Expr::BinOp { op, lhs: Box::new(lhs), rhs: Box::new(rhs) };
        }
        Ok(lhs)
    }

    fn parse_modop(&mut self) -> Result<Expr, String> {
        let mut lhs = self.parse_intdiv()?;
        while self.is_ident("mod") {
            self.advance();
            let rhs = self.parse_intdiv()?;
            lhs = Expr::BinOp { op: VbaBinOp::Mod, lhs: Box::new(lhs), rhs: Box::new(rhs) };
        }
        Ok(lhs)
    }

    fn parse_intdiv(&mut self) -> Result<Expr, String> {
        let mut lhs = self.parse_term()?;
        while *self.peek() == Tok::Backslash {
            self.advance();
            let rhs = self.parse_term()?;
            lhs = Expr::BinOp { op: VbaBinOp::IntDiv, lhs: Box::new(lhs), rhs: Box::new(rhs) };
        }
        Ok(lhs)
    }

    fn parse_term(&mut self) -> Result<Expr, String> {
        let mut lhs = self.parse_unary()?;
        loop {
            let op = match self.peek() {
                Tok::Star  => VbaBinOp::Mul,
                Tok::Slash => VbaBinOp::Div,
                _ => break,
            };
            self.advance();
            let rhs = self.parse_unary()?;
            lhs = Expr::BinOp { op, lhs: Box::new(lhs), rhs: Box::new(rhs) };
        }
        Ok(lhs)
    }

    // Unary minus binds looser than `^` overall (`-2 ^ 2` is `-(2 ^ 2)` =
    // -4) but a `^`'s immediate right-hand operand may still start with its
    // own unary minus (`2 ^ -2` is 2 ^ (-2)) — see `parse_pow_operand`.
    // Recurses into itself so a stacked `- -x` still parses, matching the
    // pre-existing single-level behavior's intent but now allowing repeats.
    fn parse_unary(&mut self) -> Result<Expr, String> {
        if *self.peek() == Tok::Minus {
            self.advance();
            Ok(Expr::UnaryMinus(Box::new(self.parse_unary()?)))
        } else {
            self.parse_pow()
        }
    }

    // Exponentiation — highest precedence. Left-associative (`2 ^ 3 ^ 2` is
    // `(2 ^ 3) ^ 2` = 64), matching real VBA's documented left-to-right
    // evaluation rather than the right-associative convention some other
    // languages use for `^`/`**`.
    fn parse_pow(&mut self) -> Result<Expr, String> {
        let mut lhs = self.parse_pow_operand()?;
        while *self.peek() == Tok::Caret {
            self.advance();
            let rhs = self.parse_pow_operand()?;
            lhs = Expr::BinOp { op: VbaBinOp::Pow, lhs: Box::new(lhs), rhs: Box::new(rhs) };
        }
        Ok(lhs)
    }

    // A `^` operand, allowing one tightly-bound leading unary minus so `2 ^
    // -2` parses as `2 ^ (-2)` — without this, `-2` on the right of `^`
    // would have nowhere to bind, since plain unary minus sits at a looser
    // tier than `^` (see `parse_unary`).
    fn parse_pow_operand(&mut self) -> Result<Expr, String> {
        if *self.peek() == Tok::Minus {
            self.advance();
            Ok(Expr::UnaryMinus(Box::new(self.parse_pow_operand()?)))
        } else {
            self.parse_primary()
        }
    }

    fn parse_primary(&mut self) -> Result<Expr, String> {
        match self.peek().clone() {
            Tok::LParen => {
                self.advance();
                // Full expression grammar, not just parse_comparison — a
                // parenthesized sub-expression can contain And/Or/Xor/Not
                // too (e.g. `(a And b) Or c`).
                let e = self.parse_expr()?;
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
                    // `ActiveSheet.Range(...)`/`.Cells(...)` (Milestone B7c
                    // item 6). Only when followed by `.` — a bare
                    // `ActiveSheet` (e.g. an unmodeled `Set ws =
                    // ActiveSheet`) falls through to `parse_ident_expr`
                    // like any other unrecognized bare identifier.
                    "activesheet" if *self.peek_at(1) == Tok::Dot => {
                        self.advance();
                        self.parse_sheet_property_read(Expr::ActiveSheetRef)
                    }
                    // `ThisWorkbook.Worksheets(...)` / `ActiveWorkbook.
                    // Worksheets(...)` (Milestone B7c item 6) — see the
                    // matching statement-dispatch arm in `parse_stmt`.
                    "thisworkbook" | "activeworkbook"
                        if self.is_ident_at(2, "worksheets") || self.is_ident_at(2, "sheets") =>
                    {
                        self.advance();
                        self.expect_tok(Tok::Dot)?;
                        self.parse_sheet_cell_read()
                    }
                    // `Err.Number` / `Err.Description` / `Err.Source` /
                    // `Err.HelpFile` / `Err.HelpContext` — guarded on the
                    // exact member name, same precedent as the
                    // `thisworkbook`/`activeworkbook` arm above, so a
                    // genuine user variable named `err` with an unrelated
                    // UDT field (`x = err.code`) still parses as ordinary
                    // field access.
                    "err" if self.is_ident_at(2, "number")
                        || self.is_ident_at(2, "description")
                        || self.is_ident_at(2, "source")
                        || self.is_ident_at(2, "helpfile")
                        || self.is_ident_at(2, "helpcontext") =>
                    {
                        self.advance();
                        self.expect_tok(Tok::Dot)?;
                        if self.is_ident("number") {
                            self.advance();
                            Ok(Expr::ErrNumber)
                        } else if self.is_ident("description") {
                            self.advance();
                            Ok(Expr::ErrDescription)
                        } else if self.is_ident("source") {
                            self.advance();
                            Ok(Expr::ErrSource)
                        } else if self.is_ident("helpfile") {
                            self.advance();
                            Ok(Expr::ErrHelpFile)
                        } else {
                            self.expect_ident("helpcontext")?;
                            Ok(Expr::ErrHelpContext)
                        }
                    }
                    _ => self.parse_ident_expr(),
                }
            }
            // ── A bare `.member` read inside a With body ──────────────────────
            // e.g. the right-hand side of `.Value = .Value + 1`. Resolved
            // against the innermost active With target at runtime, so this
            // needs no parser state and works at any nesting depth.
            Tok::Dot => {
                self.advance(); // consume '.'
                let mut fields = vec![self.consume_ident()?.to_lowercase()];
                while *self.peek() == Tok::Dot && matches!(self.peek_at(1), Tok::Ident(_)) {
                    self.advance(); // consume '.'
                    fields.push(self.consume_ident()?.to_lowercase());
                }
                Ok(Expr::WithDot(fields))
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
        } else if *self.peek() == Tok::Dot
            && (self.is_ident_at(1, "range") || self.is_ident_at(1, "cells"))
            && *self.peek_at(2) == Tok::LParen
        {
            // x = <var>.Range(addr).Value / x = <var>.Cells(r,c).Value —
            // read-side twin of the statement-dispatch branch in
            // `parse_ident_stmt` (Phase 2C item 7); see its comment for the
            // full rationale.
            self.parse_sheet_property_read(Expr::ObjectVarSheet(name))
        } else if *self.peek() == Tok::Dot
            && (self.is_ident_at(1, "worksheets") || self.is_ident_at(1, "sheets"))
            && *self.peek_at(2) == Tok::LParen
        {
            // x = <var>.Worksheets(...)/.Sheets(...) — read-side twin
            // (Phase 2C item 8); same paren guard as the statement-dispatch
            // branch in `parse_ident_stmt` (a paren-less `wb.Worksheets.
            // Count` or a UDT field literally named "worksheets"/"sheets"
            // must still fall through to the generic `RecordGet` path).
            self.advance(); // '.'
            self.parse_sheet_cell_read()
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
    #[test] fn test_single_line_if_no_else() {
        let body = parse_body("Sub MySub()\n    If a > m Then m = a\n    n = 1\nEnd Sub\n");
        assert!(matches!(&body[0], Stmt::If { .. }));
        if let Stmt::If { then_body, else_body, .. } = &body[0] {
            assert_eq!(then_body.len(), 1);
            assert!(else_body.is_empty());
            assert!(matches!(&then_body[0].stmt, Stmt::Assignment { var, .. } if var == "m"));
        }
        // The statement on the line after the single-line If must still parse.
        assert_eq!(body.len(), 2);
        assert!(matches!(&body[1], Stmt::Assignment { var, .. } if var == "n"));
    }
    #[test] fn test_single_line_if_with_else() {
        let body = parse_body("Sub MySub()\n    If a > 0 Then b = 1 Else b = 0\nEnd Sub\n");
        if let Stmt::If { then_body, else_body, .. } = &body[0] {
            assert_eq!(then_body.len(), 1);
            assert_eq!(else_body.len(), 1);
            assert!(matches!(&then_body[0].stmt, Stmt::Assignment { .. }));
            assert!(matches!(&else_body[0].stmt, Stmt::Assignment { .. }));
        } else {
            panic!("expected Stmt::If");
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
        assert_eq!(body[0], Stmt::DimBare { var: "x".to_string() });
    }
    #[test] fn test_dim_multi_declarator_all_builtin() {
        let body = parse_body("Sub MySub()\n    Dim a As Integer, b As String\n    a = 1\nEnd Sub\n");
        assert_eq!(
            body[0],
            Stmt::DimMulti(vec![
                Stmt::DimBare { var: "a".to_string() },
                Stmt::DimBare { var: "b".to_string() },
            ])
        );
    }
    #[test] fn test_dim_multi_declarator_mixed_record_type() {
        // The exact shape called out in CHANGELOG.md's Known limitations as
        // unparseable before this fix: a built-in-typed declarator followed
        // by a user-defined-typed one on the same `Dim`.
        let body = parse_body("Sub MySub()\n    Dim a As Integer, b As MyType\n    a = 1\nEnd Sub\n");
        assert_eq!(
            body[0],
            Stmt::DimMulti(vec![
                Stmt::DimBare { var: "a".to_string() },
                Stmt::DimRecord { var: "b".to_string(), type_name: "mytype".to_string() },
            ])
        );
    }
    #[test] fn test_dim_multi_declarator_three_way_with_array() {
        let body = parse_body("Sub MySub()\n    Dim a As Integer, b(3) As Integer, c As MyType\n    a = 1\nEnd Sub\n");
        assert_eq!(
            body[0],
            Stmt::DimMulti(vec![
                Stmt::DimBare { var: "a".to_string() },
                Stmt::DimArray { name: "b".to_string(), sizes: vec![ArrayDim { lower: None, upper: Expr::Integer(3) }] },
                Stmt::DimRecord { var: "c".to_string(), type_name: "mytype".to_string() },
            ])
        );
    }

    #[test] fn dim_array_explicit_lo_to_hi_parses() {
        let body = parse_body("Sub MySub()\n    Dim arr(2 To 8)\n    a = 1\nEnd Sub\n");
        assert_eq!(
            body[0],
            Stmt::DimArray {
                name: "arr".to_string(),
                sizes: vec![ArrayDim { lower: Some(Expr::Integer(2)), upper: Expr::Integer(8) }],
            }
        );
    }

    #[test] fn dim_array_empty_parens_parses_as_a_zero_dimension_declarator() {
        let body = parse_body("Sub MySub()\n    Dim arr()\n    a = 1\nEnd Sub\n");
        assert_eq!(body[0], Stmt::DimArray { name: "arr".to_string(), sizes: vec![] });
    }

    #[test] fn redim_explicit_lo_to_hi_parses() {
        let body = parse_body("Sub MySub()\n    ReDim arr(2 To 8)\n    a = 1\nEnd Sub\n");
        assert_eq!(
            body[0],
            Stmt::ReDim {
                name: "arr".to_string(),
                sizes: vec![ArrayDim { lower: Some(Expr::Integer(2)), upper: Expr::Integer(8) }],
                preserve: false,
            }
        );
    }

    #[test] fn erase_parses_as_its_own_statement() {
        let body = parse_body("Sub MySub()\n    Erase arr\n    a = 1\nEnd Sub\n");
        assert_eq!(body[0], Stmt::Erase { name: "arr".to_string() });
    }
    #[test] fn test_dim_fixed_length_string_trailing_syntax_is_tolerated() {
        // `As String * 10`'s fixed-length suffix isn't modeled by
        // `parse_dim_declarator`, but it must still be tolerated (consumed,
        // not left for `eat_stmt_end()` to choke on) the same way a
        // single-declarator `Dim` always tolerated unmodeled trailing
        // syntax on its own line, before the comma loop existed.
        let body = parse_body("Sub MySub()\n    Dim s As String * 10\n    s = \"hi\"\nEnd Sub\n");
        assert_eq!(body[0], Stmt::DimBare { var: "s".to_string() });
    }
    #[test] fn test_dim_fixed_length_string_mixed_with_comma_declarator() {
        let body = parse_body("Sub MySub()\n    Dim s As String * 10, i As Integer\n    i = 1\nEnd Sub\n");
        assert_eq!(
            body[0],
            Stmt::DimMulti(vec![
                Stmt::DimBare { var: "s".to_string() },
                Stmt::DimBare { var: "i".to_string() },
            ])
        );
    }
    #[test] fn test_with_block() {
        let body = parse_body("Sub MySub()\n    With Sheet1\n        .Cells(1, 1).Value = 99\n    End With\nEnd Sub\n");
        // `With Sheet1` is a bare-identifier target — `WithTarget::Var`,
        // resolved at runtime (it may name a Set-assigned object variable
        // or a UDT record; the parser can't tell).
        let body_len = match &body[0] {
            Stmt::With { target: WithTarget::Var(v), body } if v == "sheet1" => body.len(),
            other => panic!("expected With {{ target: Var(\"sheet1\") }}, got {:?}", other),
        };
        assert_eq!(body_len, 1);
    }

    #[test] fn test_with_udt_field_read_without_assignment_is_unsupported() {
        let body = parse_body("Sub MySub()\n    With p\n        .Field\n    End With\nEnd Sub\n");
        let inner = match &body[0] {
            Stmt::With { body, .. } => body,
            other => panic!("expected With, got {:?}", other),
        };
        assert_eq!(
            inner[0].stmt,
            Stmt::Unsupported {
                reason: "With-block '.field' read without assignment has no effect".to_string()
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
                reason: "With-block '.foo' read without assignment has no effect".to_string()
            }
        );
    }

    #[test] fn test_with_non_identifier_dotted_statement_is_unsupported() {
        // `.42` tokenizes as Dot, Int(42) — a non-identifier after the dot.
        let body = parse_body("Sub MySub()\n    With p\n        .42\n    End With\nEnd Sub\n");
        let inner = match &body[0] {
            Stmt::With { body, .. } => body,
            other => panic!("expected With, got {:?}", other),
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
    #[test] fn test_call_stmt_no_parens_zero_args() {
        // `Call Name` (no parens) is valid VBA for a zero-argument call —
        // the parens in `Call name [(argumentlist)]` are optional.
        let body = parse_body("Sub MySub()\n    Call MySub2\nEnd Sub\n");
        assert!(matches!(&body[0], Stmt::CallSub { name, args } if name == "mysub2" && args.is_empty()));
    }
    #[test] fn test_call_stmt_no_parens_followed_by_another_statement() {
        // A parenless `Call` must still correctly hand off statement
        // termination to the caller — regression guard against consuming
        // too much (or too little) of the line.
        let body = parse_body("Sub MySub()\n    Call MySub2\n    x = 1\nEnd Sub\n");
        assert!(matches!(&body[0], Stmt::CallSub { name, args } if name == "mysub2" && args.is_empty()));
        assert!(matches!(&body[1], Stmt::Assignment { var, .. } if var == "x"));
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

    #[test] fn byval_and_byref_param_modifiers_are_recognized_and_discarded_not_treated_as_a_param_name() {
        // Regression: `consume_ident()` used to swallow "byval"/"byref"
        // itself as a bogus extra parameter, so `Sub Foo(ByVal x As
        // Integer)` parsed as a 2-param sub (`["byval", "x"]`) and a caller
        // passing one argument bound it to the phantom "byval" param,
        // leaving `x` unbound.
        let prog = parse("Sub Foo(ByVal x As Integer, ByRef y As String)\n    a = x\nEnd Sub\n").unwrap();
        assert_eq!(prog.subs[0].params, vec!["x", "y"]);
    }

    #[test] fn optional_parameter_modifier_is_a_clear_parse_error_not_a_silent_misparse() {
        let err = parse("Sub Foo(Optional x As Integer)\n    a = x\nEnd Sub\n").unwrap_err();
        assert!(err.contains("Optional"), "error should name the unsupported modifier: {err}");
    }

    #[test] fn paramarray_parameter_modifier_is_a_clear_parse_error_not_a_silent_misparse() {
        let err = parse("Sub Foo(ParamArray items())\n    a = 1\nEnd Sub\n").unwrap_err();
        assert!(err.contains("ParamArray"), "error should name the unsupported modifier: {err}");
    }

    // ── Module-level declarations and access modifiers ─────────────────────────

    #[test] fn test_option_explicit_ignored() {
        let prog = parse("Option Explicit\nSub MySub()\n    a = 1\nEnd Sub\n").unwrap();
        assert_eq!(prog.subs[0].name, "mysub");
    }

    #[test] fn test_option_base_is_captured_and_does_not_disrupt_parsing() {
        let prog = parse("Option Base 1\nOption Explicit\nSub MySub()\n    a = 1\nEnd Sub\n").unwrap();
        assert_eq!(prog.subs.len(), 1);
        assert_eq!(prog.option_base, 1);
    }

    #[test] fn test_option_base_defaults_to_zero_when_absent() {
        let prog = parse("Sub MySub()\n    a = 1\nEnd Sub\n").unwrap();
        assert_eq!(prog.option_base, 0);
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

    #[test] fn err_clear_parses() {
        let body = parse_body("Sub MySub()\n    Err.Clear\nEnd Sub\n");
        assert_eq!(body, vec![Stmt::ErrClear]);
    }

    #[test] fn err_raise_number_only_parses() {
        let body = parse_body("Sub MySub()\n    Err.Raise 5\nEnd Sub\n");
        assert_eq!(body, vec![Stmt::ErrRaise {
            number: Expr::Integer(5),
            source: None,
            description: None,
            help_file: None,
            help_context: None,
        }]);
    }

    #[test] fn err_raise_skips_source_with_a_bare_comma() {
        let body = parse_body("Sub MySub()\n    Err.Raise 513, , \"custom text\"\nEnd Sub\n");
        assert_eq!(body, vec![Stmt::ErrRaise {
            number: Expr::Integer(513),
            source: None,
            description: Some(Expr::Str("custom text".into())),
            help_file: None,
            help_context: None,
        }]);
    }

    #[test] fn err_raise_with_source_and_description_parses() {
        let body = parse_body("Sub MySub()\n    Err.Raise 513, \"MySource\", \"custom text\"\nEnd Sub\n");
        assert_eq!(body, vec![Stmt::ErrRaise {
            number: Expr::Integer(513),
            source: Some(Expr::Str("MySource".into())),
            description: Some(Expr::Str("custom text".into())),
            help_file: None,
            help_context: None,
        }]);
    }

    #[test] fn err_raise_with_all_five_positional_arguments_parses() {
        let body = parse_body(
            "Sub MySub()\n    Err.Raise 513, \"MySource\", \"custom text\", \"help.chm\", 100\nEnd Sub\n",
        );
        assert_eq!(body, vec![Stmt::ErrRaise {
            number: Expr::Integer(513),
            source: Some(Expr::Str("MySource".into())),
            description: Some(Expr::Str("custom text".into())),
            help_file: Some(Expr::Str("help.chm".into())),
            help_context: Some(Expr::Integer(100)),
        }]);
    }

    #[test] fn err_raise_skips_help_file_with_a_bare_comma() {
        let body = parse_body(
            "Sub MySub()\n    Err.Raise 513, \"MySource\", \"custom text\", , 100\nEnd Sub\n",
        );
        assert_eq!(body, vec![Stmt::ErrRaise {
            number: Expr::Integer(513),
            source: Some(Expr::Str("MySource".into())),
            description: Some(Expr::Str("custom text".into())),
            help_file: None,
            help_context: Some(Expr::Integer(100)),
        }]);
    }

    #[test] fn err_source_help_file_help_context_parse_as_expressions() {
        let body = parse_body(concat!(
            "Sub MySub()\n",
            "    s = Err.Source\n",
            "    h = Err.HelpFile\n",
            "    c = Err.HelpContext\n",
            "End Sub\n",
        ));
        assert_eq!(body, vec![
            Stmt::Assignment { var: "s".into(), value: Expr::ErrSource },
            Stmt::Assignment { var: "h".into(), value: Expr::ErrHelpFile },
            Stmt::Assignment { var: "c".into(), value: Expr::ErrHelpContext },
        ]);
    }

    #[test] fn err_number_and_description_parse_as_expressions() {
        let body = parse_body("Sub MySub()\n    n = Err.Number\n    d = Err.Description\nEnd Sub\n");
        assert_eq!(body, vec![
            Stmt::Assignment { var: "n".into(), value: Expr::ErrNumber },
            Stmt::Assignment { var: "d".into(), value: Expr::ErrDescription },
        ]);
    }

    #[test] fn a_bare_err_variable_is_unaffected_by_err_object_parsing() {
        // No `.Number`/`.Description`/`.Clear`/`.Raise` suffix — an
        // ordinary variable assignment, same as any other identifier.
        let body = parse_body("Sub MySub()\n    err = 1\nEnd Sub\n");
        assert_eq!(body, vec![Stmt::Assignment { var: "err".into(), value: Expr::Integer(1) }]);
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

    // ── Milestone B7c: Set / object references ───────────────────────────────

    #[test] fn set_range_literal_parses_to_stmt_set_with_a_range_lit_object_expr() {
        let body = parse_body("Sub MySub()\n    Set rng = Range(\"A1:B2\")\nEnd Sub\n");
        assert_eq!(
            body,
            vec![Stmt::Set { var: "rng".into(), value: ObjectExpr::RangeLit("A1:B2".into()) }]
        );
    }

    #[test] fn set_from_another_object_variable_parses_to_object_expr_var() {
        let body = parse_body("Sub MySub()\n    Set a = Range(\"A1\")\n    Set b = a\nEnd Sub\n");
        assert_eq!(body[1], Stmt::Set { var: "b".into(), value: ObjectExpr::Var("a".into()) });
    }

    #[test] fn set_union_of_two_range_literals_parses_to_object_expr_union() {
        let body = parse_body(
            "Sub MySub()\n    Set u = Union(Range(\"A1:A2\"), Range(\"C1:C2\"))\nEnd Sub\n",
        );
        assert_eq!(
            body,
            vec![Stmt::Set {
                var: "u".into(),
                value: ObjectExpr::Union(vec![
                    ObjectExpr::RangeLit("A1:A2".into()),
                    ObjectExpr::RangeLit("C1:C2".into()),
                ]),
            }]
        );
    }

    #[test] fn set_areas_index_parses_to_object_expr_area() {
        let body = parse_body("Sub MySub()\n    Set u = Range(\"A1,C1\")\n    Set a = u.Areas(1)\nEnd Sub\n");
        assert_eq!(
            body[1],
            Stmt::Set {
                var: "a".into(),
                value: ObjectExpr::Area(Box::new(ObjectExpr::Var("u".into())), Box::new(Expr::Integer(1))),
            }
        );
    }

    #[test] fn set_specialcells_visible_parses_to_object_expr_special_cells_visible() {
        let body = parse_body(
            "Sub MySub()\n    Set u = Range(\"A1:A3\")\n    Set v = u.SpecialCells(xlCellTypeVisible)\nEnd Sub\n",
        );
        assert_eq!(
            body[1],
            Stmt::Set {
                var: "v".into(),
                value: ObjectExpr::SpecialCellsVisible(Box::new(ObjectExpr::Var("u".into()))),
            }
        );
    }

    #[test] fn set_with_an_unrecognized_rhs_degrades_to_unsupported_not_a_parse_error() {
        // `CreateObject(...)` isn't a modeled object expression — this must
        // stay a soft no-op (see `parse_set`'s doc comment), not a hard
        // parse error that would take down an otherwise-parseable module.
        let body = parse_body(
            "Sub MySub()\n    Set d = CreateObject(\"Scripting.Dictionary\")\nEnd Sub\n",
        );
        assert!(matches!(body[0], Stmt::Unsupported { .. }), "{:?}", body[0]);
    }

    #[test] fn range_object_copy_with_destination_parses_to_range_object_copy() {
        let body = parse_body(
            "Sub MySub()\n    Set rng = Range(\"A1\")\n    rng.Copy Destination:=Range(\"B1\")\nEnd Sub\n",
        );
        assert_eq!(
            body[1],
            Stmt::RangeObjectCopy { var: "rng".into(), dst: Some("B1".into()) }
        );
    }

    #[test] fn bare_range_object_copy_has_no_destination() {
        let body = parse_body("Sub MySub()\n    Set rng = Range(\"A1\")\n    rng.Copy\nEnd Sub\n");
        assert_eq!(body[1], Stmt::RangeObjectCopy { var: "rng".into(), dst: None });
    }

    // ── Milestone B7c item 6: ThisWorkbook / ActiveWorkbook / ActiveSheet ────

    #[test] fn activesheet_cell_write_parses_to_sheet_cell_write_with_activesheetref() {
        let body = parse_body("Sub MySub()\n    ActiveSheet.Cells(1, 1).Value = 5\nEnd Sub\n");
        assert_eq!(
            body,
            vec![Stmt::SheetCellWrite {
                sheet: Expr::ActiveSheetRef,
                row: Expr::Integer(1),
                col: Expr::Integer(1),
                value: Expr::Integer(5),
            }]
        );
    }

    #[test] fn thisworkbook_worksheets_cell_write_parses_identically_to_bare_worksheets() {
        let with_prefix = parse_body(
            "Sub MySub()\n    ThisWorkbook.Worksheets(\"Data\").Cells(1, 1).Value = 5\nEnd Sub\n",
        );
        let bare = parse_body("Sub MySub()\n    Worksheets(\"Data\").Cells(1, 1).Value = 5\nEnd Sub\n");
        assert_eq!(with_prefix, bare);
    }

    #[test] fn activeworkbook_worksheets_range_read_parses_identically_to_bare_worksheets() {
        let with_prefix = parse_body(
            "Sub MySub()\n    x = ActiveWorkbook.Worksheets(\"Data\").Range(\"A1\").Value\nEnd Sub\n",
        );
        let bare = parse_body("Sub MySub()\n    x = Worksheets(\"Data\").Range(\"A1\").Value\nEnd Sub\n");
        assert_eq!(with_prefix, bare);
    }

    #[test] fn a_bare_activesheet_without_a_dot_is_not_captured_by_the_activesheet_grammar() {
        // `Set ws = ActiveSheet` — a bare `ActiveSheet` with no `.property`
        // suffix falls through to a plain identifier (harmless; the VM
        // treats an unresolved bare object-variable reference as a no-op —
        // see `Stmt::Set`'s doc comment). This just confirms the parser
        // doesn't error out on it.
        let body = parse_body("Sub MySub()\n    Set ws = ActiveSheet\nEnd Sub\n");
        assert_eq!(body, vec![Stmt::Set { var: "ws".into(), value: ObjectExpr::Var("activesheet".into()) }]);
    }

    // ── Phase 2C items 7/8: object-variable sheet/workbook qualifiers ────────

    #[test] fn object_var_range_write_parses_to_range_write_with_objectvarsheet() {
        let body = parse_body("Sub MySub()\n    ws.Range(\"A1\").Value = 5\nEnd Sub\n");
        assert_eq!(
            body,
            vec![Stmt::SheetRangeWrite {
                sheet: Expr::ObjectVarSheet("ws".into()),
                addr: "A1".into(),
                is_formula: false,
                value: Expr::Integer(5),
            }]
        );
    }

    #[test] fn object_var_cells_read_parses_to_sheet_cell_read_with_objectvarsheet() {
        let body = parse_body("Sub MySub()\n    x = ws.Cells(1, 1).Value\nEnd Sub\n");
        assert_eq!(
            body,
            vec![Stmt::Assignment {
                var: "x".into(),
                value: Expr::SheetCellRead {
                    sheet: Box::new(Expr::ObjectVarSheet("ws".into())),
                    row: Box::new(Expr::Integer(1)),
                    col: Box::new(Expr::Integer(1)),
                },
            }]
        );
    }

    #[test] fn object_var_worksheets_write_parses_identically_to_bare_worksheets() {
        let with_prefix = parse_body(
            "Sub MySub()\n    wb.Worksheets(\"Data\").Cells(1, 1).Value = 5\nEnd Sub\n",
        );
        let bare = parse_body("Sub MySub()\n    Worksheets(\"Data\").Cells(1, 1).Value = 5\nEnd Sub\n");
        assert_eq!(with_prefix, bare);
    }

    // A prior review round found the `.Worksheets(`/`.Sheets(` branches
    // above missing the immediate-`(` guard their `.Range(`/`.Cells(`
    // siblings already had — without it, any paren-less `<var>.Worksheets`/
    // `<var>.Sheets` (a real, if unmodeled, VBA read — `wb.Worksheets.
    // Count`) or a UDT field literally named "worksheets"/"sheets" would
    // hit a hard `expected LParen` parse error instead of falling through
    // to the pre-existing generic `RecordGet`/`RecordSet` no-op path. These
    // pin the fix; they don't need to assert *what* the fallback parses to,
    // only that it doesn't error, matching this file's existing "confirms
    // the parser doesn't error out on it" precedent just above.
    #[test] fn paren_less_worksheets_property_falls_back_instead_of_erroring() {
        let _ = parse("Sub MySub()\n    x = wb.Worksheets.Count\nEnd Sub\n").unwrap();
        let _ = parse("Sub MySub()\n    wb.Sheets.Add\nEnd Sub\n").unwrap();
    }

    #[test] fn udt_field_literally_named_sheets_still_round_trips() {
        let body = parse_body("Sub MySub()\n    p.sheets = 5\nEnd Sub\n");
        assert_eq!(
            body,
            vec![Stmt::RecordSet { var: "p".into(), field: "sheets".into(), value: Expr::Integer(5) }]
        );
    }

    // ── `:` statement separator ──────────────────────────────────────────
    // Real VBA's multi-statement-per-line form. Handled in the parser (the
    // tokenizer's own `Tok::Colon`), never as a pre-tokenize string rewrite
    // of `:` to a newline — which would corrupt `MsgBox "10:30"`, break the
    // `label:` declaration form, and mangle a single-line `If`'s own
    // Then/Else boundary. Each of those three is pinned below.

    #[test] fn colon_separates_two_statements_on_one_line() {
        let body = parse_body("Sub MySub()\n    a = 1: b = 2\nEnd Sub\n");
        assert_eq!(body, vec![
            Stmt::Assignment { var: "a".into(), value: Expr::Integer(1) },
            Stmt::Assignment { var: "b".into(), value: Expr::Integer(2) },
        ]);
    }

    #[test] fn colon_separates_three_statements_on_one_line() {
        let body = parse_body("Sub MySub()\n    a = 1: b = 2: c = 3\nEnd Sub\n");
        assert_eq!(body.len(), 3);
        assert_eq!(body[2], Stmt::Assignment { var: "c".into(), value: Expr::Integer(3) });
    }

    #[test] fn colon_inside_a_string_literal_is_not_a_separator() {
        // The load-bearing case against a naive `:`→newline pre-rewrite.
        let body = parse_body("Sub MySub()\n    MsgBox \"10:30\": a = 1\nEnd Sub\n");
        assert_eq!(body, vec![
            Stmt::MsgBox { message: Expr::Str("10:30".into()) },
            Stmt::Assignment { var: "a".into(), value: Expr::Integer(1) },
        ]);
    }

    #[test] fn label_followed_by_a_statement_on_the_same_line() {
        let body = parse_body("Sub MySub()\n    label1: a = 1\nEnd Sub\n");
        assert_eq!(body, vec![
            Stmt::Label("label1".into()),
            Stmt::Assignment { var: "a".into(), value: Expr::Integer(1) },
        ]);
    }

    #[test] fn a_bare_label_on_its_own_line_still_parses_as_just_a_label() {
        let body = parse_body("Sub MySub()\n    errh:\n    a = 1\nEnd Sub\n");
        assert_eq!(body, vec![
            Stmt::Label("errh".into()),
            Stmt::Assignment { var: "a".into(), value: Expr::Integer(1) },
        ]);
    }

    #[test] fn single_line_if_then_takes_a_colon_separated_statement_list() {
        // Microsoft's own worked example shape: every colon-separated
        // statement after `Then` belongs to the Then branch.
        let body = parse_body("Sub MySub()\n    If x > 10 Then a = 1: b = 2: c = 3\nEnd Sub\n");
        match &body[0] {
            Stmt::If { then_body, else_body, .. } => {
                assert_eq!(then_body.len(), 3);
                assert!(else_body.is_empty());
            }
            other => panic!("expected If, got {:?}", other),
        }
    }

    #[test] fn single_line_if_else_takes_everything_after_else_on_the_line() {
        let body = parse_body("Sub MySub()\n    If x Then a = 1 Else b = 2: c = 3\nEnd Sub\n");
        match &body[0] {
            Stmt::If { then_body, else_body, .. } => {
                assert_eq!(then_body.len(), 1);
                assert_eq!(else_body.len(), 2);
            }
            other => panic!("expected If, got {:?}", other),
        }
    }

    #[test] fn colon_separated_statements_keep_their_own_distinct_spans() {
        // Not collapsed into one span: an error on the second statement must
        // point at the second statement, exactly as a newline-separated one
        // would (`SourceSpan` accuracy is what `--json`'s `location` reports).
        let sub = parse("Sub MySub()\n    a = 1: b = 2\nEnd Sub\n").unwrap()
            .subs.into_iter().next().unwrap();
        assert_ne!(sub.body[0].span.start, sub.body[1].span.start);
        // `b` starts after `a = 1: ` on the same line.
        assert!(sub.body[1].span.start > sub.body[0].span.start);
    }

    #[test] fn colon_terminates_a_block_construct_header_and_body() {
        let body = parse_body("Sub MySub()\n    For i = 1 To 3: a = i: Next i\nEnd Sub\n");
        match &body[0] {
            Stmt::For { body, .. } => assert_eq!(body.len(), 1),
            other => panic!("expected For, got {:?}", other),
        }
    }

    #[test] fn trailing_colon_after_then_still_parses_as_a_block_if() {
        let body = parse_body("Sub MySub()\n    If x Then:\n    a = 1\n    End If\nEnd Sub\n");
        match &body[0] {
            Stmt::If { then_body, .. } => assert_eq!(then_body.len(), 1),
            other => panic!("expected block If, got {:?}", other),
        }
    }

    #[test] fn named_argument_colon_equals_is_not_a_statement_separator() {
        // `:=` tokenizes as `Tok::ColonEq`, never `Tok::Colon`.
        let body = parse_body(
            "Sub MySub()\n    Range(\"A1\").Copy Destination:=Range(\"B1\")\nEnd Sub\n");
        assert_eq!(body.len(), 1);
    }

    #[test] fn integer_literal_wider_than_i64_falls_back_to_float_instead_of_panicking() {
        // Found by fuzz_vba_parser: a decimal literal too large for i64 used to
        // panic in tokenize() via `s.parse().unwrap()` on a `PosOverflow` error —
        // a crash on ordinary (if unusual) source text, not just adversarial
        // bytes.
        let body = parse_body("Sub MySub()\n    a = 99999999999999999999\nEnd Sub\n");
        assert_eq!(body, vec![Stmt::Assignment {
            var: "a".into(),
            value: Expr::Float(99999999999999999999.0),
        }]);
    }
}
