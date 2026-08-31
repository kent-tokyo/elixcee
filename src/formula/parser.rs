use super::ast::{BinOpKind, FormulaExpr, SheetQualifier};

const MAX_FORMULA_BYTES: usize = 1 * 1024 * 1024;
const MAX_FORMULA_REFS: usize = 100_000;
const MAX_FORMULA_NODES: usize = 200_000;
const MAX_FORMULA_DEPTH: usize = 256;

/// One cell/range reference as it literally appears in formula text, with its
/// exact char-offset span (relative to the normalized input `parse_with_refs`
/// was called on -- see that function's doc comment). Used by the reference
/// rewriter (`super::rewrite`) to patch only the substring that actually
/// changed on a structural edit, leaving everything else in the formula --
/// operators, function names, number/string literals, whitespace -- byte-for-
/// byte untouched. `span`/`c1_span`/`c2_span` cover the *coordinate* text only
/// (e.g. `A1` or `A1:B10`) -- never the `Sheet2!` qualifier, which the
/// rewriter never touches and which is recorded separately in `sheet`.
#[derive(Debug, Clone, PartialEq)]
pub enum RefOccurrence {
    Cell {
        span: (usize, usize),
        col: u32,
        row: u32,
        abs_col: bool,
        abs_row: bool,
        sheet: Option<SheetQualifier>,
    },
    Range {
        /// Span of the whole reference, e.g. all of `A1:B10` including the `:`.
        span: (usize, usize),
        c1: u32,
        r1: u32,
        abs_c1: bool,
        abs_r1: bool,
        c1_span: (usize, usize),
        c2: u32,
        r2: u32,
        abs_c2: bool,
        abs_r2: bool,
        c2_span: (usize, usize),
        sheet: Option<SheetQualifier>,
    },
}

pub struct FormulaParser {
    chars: Vec<char>,
    pos: usize,
    refs: Vec<RefOccurrence>,
    depth: usize,
}

impl FormulaParser {
    fn new(input: &str) -> Self {
        FormulaParser {
            chars: input.chars().collect(),
            pos: 0,
            refs: Vec::new(),
            depth: 0,
        }
    }

    fn parse_nested_expr(&mut self) -> Result<FormulaExpr, String> {
        if self.depth >= MAX_FORMULA_DEPTH {
            return Err(format!(
                "Formula nesting exceeds the maximum depth of {}",
                MAX_FORMULA_DEPTH
            ));
        }
        self.depth += 1;
        let result = self.parse_expr();
        self.depth -= 1;
        result
    }

    fn peek(&self) -> Option<char> {
        self.chars.get(self.pos).copied()
    }

    fn advance(&mut self) -> Option<char> {
        let c = self.chars.get(self.pos).copied();
        if c.is_some() {
            self.pos += 1;
        }
        c
    }

    fn skip_ws(&mut self) {
        while matches!(self.peek(), Some(' ') | Some('\t')) {
            self.pos += 1;
        }
    }

    fn consume(&mut self, c: char) -> bool {
        if self.peek() == Some(c) {
            self.pos += 1;
            true
        } else {
            false
        }
    }

    pub fn parse_expr(&mut self) -> Result<FormulaExpr, String> {
        self.skip_ws();
        self.parse_comparison()
    }

    fn parse_comparison(&mut self) -> Result<FormulaExpr, String> {
        let mut lhs = self.parse_concat()?;
        loop {
            self.skip_ws();
            let op = match self.peek() {
                Some('<') => {
                    self.advance();
                    if self.consume('>') {
                        BinOpKind::Ne
                    } else if self.consume('=') {
                        BinOpKind::Le
                    } else {
                        BinOpKind::Lt
                    }
                }
                Some('>') => {
                    self.advance();
                    if self.consume('=') {
                        BinOpKind::Ge
                    } else {
                        BinOpKind::Gt
                    }
                }
                Some('=') => {
                    self.advance();
                    BinOpKind::Eq
                }
                _ => break,
            };
            self.skip_ws();
            let rhs = self.parse_concat()?;
            lhs = FormulaExpr::BinOp {
                op,
                lhs: Box::new(lhs),
                rhs: Box::new(rhs),
            };
        }
        Ok(lhs)
    }

    fn parse_concat(&mut self) -> Result<FormulaExpr, String> {
        let mut lhs = self.parse_additive()?;
        loop {
            self.skip_ws();
            if self.peek() == Some('&') {
                self.advance();
                self.skip_ws();
                let rhs = self.parse_additive()?;
                lhs = FormulaExpr::BinOp {
                    op: BinOpKind::Concat,
                    lhs: Box::new(lhs),
                    rhs: Box::new(rhs),
                };
            } else {
                break;
            }
        }
        Ok(lhs)
    }

    fn parse_additive(&mut self) -> Result<FormulaExpr, String> {
        let mut lhs = self.parse_multiplicative()?;
        loop {
            self.skip_ws();
            let op = match self.peek() {
                Some('+') => {
                    self.advance();
                    BinOpKind::Add
                }
                Some('-') => {
                    self.advance();
                    BinOpKind::Sub
                }
                _ => break,
            };
            self.skip_ws();
            let rhs = self.parse_multiplicative()?;
            lhs = FormulaExpr::BinOp {
                op,
                lhs: Box::new(lhs),
                rhs: Box::new(rhs),
            };
        }
        Ok(lhs)
    }

    fn parse_multiplicative(&mut self) -> Result<FormulaExpr, String> {
        let mut lhs = self.parse_unary()?;
        loop {
            self.skip_ws();
            let op = match self.peek() {
                Some('*') => {
                    self.advance();
                    BinOpKind::Mul
                }
                Some('/') => {
                    self.advance();
                    BinOpKind::Div
                }
                _ => break,
            };
            self.skip_ws();
            let rhs = self.parse_unary()?;
            lhs = FormulaExpr::BinOp {
                op,
                lhs: Box::new(lhs),
                rhs: Box::new(rhs),
            };
        }
        Ok(lhs)
    }

    fn parse_unary(&mut self) -> Result<FormulaExpr, String> {
        self.skip_ws();
        if self.peek() == Some('-') {
            self.advance();
            Ok(FormulaExpr::UnaryMinus(Box::new(self.parse_primary()?)))
        } else {
            if self.peek() == Some('+') {
                self.advance();
            }
            self.parse_primary()
        }
    }

    fn parse_primary(&mut self) -> Result<FormulaExpr, String> {
        self.skip_ws();
        if let Some(qualifier) = self.try_parse_sheet_qualifier()? {
            return self.parse_qualified_ref(qualifier);
        }
        match self.peek() {
            Some('(') => {
                self.advance();
                let expr = self.parse_nested_expr()?;
                self.skip_ws();
                if !self.consume(')') {
                    return Err("Expected ')'".into());
                }
                Ok(expr)
            }
            Some('"') => self.parse_string(),
            Some(c) if c.is_ascii_digit() => self.parse_number(),
            Some(c) if c.is_ascii_alphabetic() => self.parse_ident_or_ref(),
            Some('$') => self.parse_ident_or_ref(),
            Some(c) => Err(format!("Unexpected character: '{}'", c)),
            None => Err("Unexpected end of formula".into()),
        }
    }

    fn parse_string(&mut self) -> Result<FormulaExpr, String> {
        self.advance(); // opening "
        let mut s = String::new();
        loop {
            match self.peek() {
                Some('"') => {
                    self.advance();
                    if self.peek() == Some('"') {
                        self.advance();
                        s.push('"');
                    } else {
                        break;
                    }
                }
                Some(c) => {
                    s.push(c);
                    self.advance();
                }
                None => return Err("Unterminated string literal".into()),
            }
        }
        Ok(FormulaExpr::Str(s))
    }

    fn parse_number(&mut self) -> Result<FormulaExpr, String> {
        let mut s = String::new();
        while matches!(self.peek(), Some(c) if c.is_ascii_digit()) {
            s.push(self.advance().unwrap());
        }
        if self.peek() == Some('.') {
            s.push(self.advance().unwrap());
            while matches!(self.peek(), Some(c) if c.is_ascii_digit()) {
                s.push(self.advance().unwrap());
            }
        }
        s.parse::<f64>()
            .map(FormulaExpr::Number)
            .map_err(|e| e.to_string())
    }

    /// Parse one side of a range (`[$]COL[$]ROW`), e.g. the `B10` in `A1:B10`
    /// or the `$B$10` in `A1:$B$10`. Always a cell reference — a range corner
    /// is never a function name. Returns the corner's own span alongside its
    /// coordinates so the caller can record it in a `RefOccurrence::Range`.
    fn parse_ref_corner(&mut self) -> Result<(u32, u32, bool, bool, usize, usize), String> {
        let start = self.pos;
        let abs_col = self.consume('$');
        let mut name = String::new();
        while matches!(self.peek(), Some(c) if c.is_ascii_alphabetic()) {
            name.push(self.advance().unwrap().to_ascii_uppercase());
        }
        if name.is_empty() {
            return Err("Expected column letters in range".into());
        }
        let abs_row = self.consume('$');
        let mut row_s = String::new();
        while matches!(self.peek(), Some(c) if c.is_ascii_digit()) {
            row_s.push(self.advance().unwrap());
        }
        if row_s.is_empty() {
            return Err(format!("Expected row number after '{}' in range", name));
        }
        let end = self.pos;
        let col = col_letters_to_num(&name);
        let row: u32 = row_s
            .parse()
            .map_err(|e: std::num::ParseIntError| e.to_string())?;
        Ok((col, row, abs_col, abs_row, start, end))
    }

    /// Shared tail for the `$`-flagged, bare-reference, and sheet-qualified
    /// parse paths: given the already-parsed first corner (and an optional
    /// sheet qualifier already consumed by the caller), check for a trailing
    /// `:SECOND` range and record a `RefOccurrence` (`Cell` or `Range`) for
    /// whichever this turns out to be, alongside building the `FormulaExpr`.
    /// `corner1` is `(col, row, abs_col, abs_row, start, end)` -- the same
    /// shape `parse_ref_corner` returns -- where `start`/`end` cover the
    /// coordinate only, never `sheet`'s own span (see `RefOccurrence`'s doc
    /// comment). Bundled into one tuple rather than 6 separate parameters to
    /// keep this under clippy's too-many-arguments threshold.
    fn finish_ref(
        &mut self,
        corner1: (u32, u32, bool, bool, usize, usize),
        sheet: Option<SheetQualifier>,
    ) -> Result<FormulaExpr, String> {
        let (col, row, abs_col, abs_row, corner1_start, corner1_end) = corner1;
        self.skip_ws();
        if self.peek() == Some(':') {
            self.advance();
            self.skip_ws();
            let (c2, r2, abs_c2, abs_r2, c2_start, c2_end) = self.parse_ref_corner()?;
            if self.refs.len() >= MAX_FORMULA_REFS {
                return Err(format!(
                    "Formula has too many references (maximum is {})",
                    MAX_FORMULA_REFS
                ));
            }
            self.refs.push(RefOccurrence::Range {
                span: (corner1_start, c2_end),
                c1: col,
                r1: row,
                abs_c1: abs_col,
                abs_r1: abs_row,
                c1_span: (corner1_start, corner1_end),
                c2,
                r2,
                abs_c2,
                abs_r2,
                c2_span: (c2_start, c2_end),
                sheet: sheet.clone(),
            });
            return Ok(FormulaExpr::Range {
                c1: col,
                r1: row,
                c2,
                r2,
                abs_c1: abs_col,
                abs_r1: abs_row,
                abs_c2,
                abs_r2,
                sheet,
            });
        }
        if self.refs.len() >= MAX_FORMULA_REFS {
            return Err(format!(
                "Formula has too many references (maximum is {})",
                MAX_FORMULA_REFS
            ));
        }
        self.refs.push(RefOccurrence::Cell {
            span: (corner1_start, corner1_end),
            col,
            row,
            abs_col,
            abs_row,
            sheet: sheet.clone(),
        });
        Ok(FormulaExpr::CellRef {
            col,
            row,
            abs_col,
            abs_row,
            sheet,
        })
    }

    /// Tentatively parses a `Sheet2!`/`'Sales 2026'!`/`'Bob''s Data'!` prefix.
    /// `!` never appears anywhere else in this grammar, so seeing it right
    /// after a candidate name is an unambiguous signal -- no other identifier
    /// (function name, boolean, bare name-reference) is ever legally followed
    /// by `!`. If no `!` follows, this was NOT a qualifier (just an identifier
    /// that happens to start where one could have): position is fully restored
    /// so the normal dispatch in `parse_primary` sees exactly what it would
    /// have without this check ever running.
    fn try_parse_sheet_qualifier(&mut self) -> Result<Option<SheetQualifier>, String> {
        let start = self.pos;
        if self.peek() == Some('\'') {
            self.advance();
            let mut raw = String::from("'");
            let mut normalized = String::new();
            loop {
                match self.peek() {
                    Some('\'') => {
                        self.advance();
                        raw.push('\'');
                        if self.peek() == Some('\'') {
                            // '' inside a quoted name -> a literal single quote.
                            self.advance();
                            raw.push('\'');
                            normalized.push('\'');
                        } else {
                            break; // the closing quote
                        }
                    }
                    Some(c) => {
                        self.advance();
                        raw.push(c);
                        normalized.push(c);
                    }
                    None => {
                        // An opened quote that's never closed isn't valid
                        // formula syntax under any interpretation -- report it
                        // rather than silently backtracking into more confusion.
                        return Err("Unterminated sheet name (missing closing ')".into());
                    }
                }
            }
            if self.peek() == Some('!') {
                let qualifier_end = self.pos;
                self.advance(); // consume '!'
                return Ok(Some(SheetQualifier {
                    raw_span: (start, qualifier_end),
                    raw_text: raw,
                    normalized_name: normalized,
                }));
            }
            // A quoted name not followed by '!' isn't valid formula syntax
            // either way (Excel formulas have no single-quoted string type) --
            // restore position and let normal dispatch produce its own error.
            self.pos = start;
            return Ok(None);
        }

        // Unicode-aware, unlike the rest of this parser's identifier grammar
        // (function/cell-ref letters are deliberately ASCII-only, since Excel
        // columns and function names always are) -- a plain non-ASCII sheet
        // name (e.g. Japanese) doesn't require quoting in real Excel, only
        // names with spaces or other special characters do.
        if matches!(self.peek(), Some(c) if c.is_alphabetic() || c == '_') {
            while matches!(self.peek(), Some(c) if c.is_alphanumeric() || c == '_') {
                self.advance();
            }
            if self.peek() == Some('!') {
                let qualifier_end = self.pos;
                self.advance(); // consume '!'
                let raw: String = self.chars[start..qualifier_end].iter().collect();
                return Ok(Some(SheetQualifier {
                    raw_span: (start, qualifier_end),
                    raw_text: raw.clone(),
                    normalized_name: raw,
                }));
            }
            self.pos = start;
            return Ok(None);
        }

        Ok(None)
    }

    /// Parses the `A1`/`A1:B10`/`$A$1:B$10` reference that must follow a
    /// sheet qualifier `try_parse_sheet_qualifier` already consumed -- always
    /// a cell/range reference, same as a range's second corner, so this reuses
    /// `parse_ref_corner` rather than the full identifier/function dispatch.
    fn parse_qualified_ref(&mut self, qualifier: SheetQualifier) -> Result<FormulaExpr, String> {
        let corner1 = self.parse_ref_corner()?;
        self.finish_ref(corner1, Some(qualifier))
    }

    fn parse_ident_or_ref(&mut self) -> Result<FormulaExpr, String> {
        let tok_start = self.pos;
        let abs_col = self.consume('$');
        let mut name = String::new();
        while matches!(self.peek(), Some(c) if c.is_ascii_alphabetic()) {
            name.push(self.advance().unwrap().to_ascii_uppercase());
        }
        if abs_col && name.is_empty() {
            return Err("Expected column letters after '$'".into());
        }

        // A leading '$' can only start a cell reference, never a function name,
        // so dot-separated function names (MODE.MULT) don't apply here.
        if !abs_col {
            // Support dot-separated function names (e.g. MODE.MULT, NETWORKDAYS.INTL)
            while self.peek() == Some('.')
                && matches!(self.chars.get(self.pos + 1), Some(c) if c.is_ascii_alphabetic())
            {
                name.push(self.advance().unwrap()); // consume '.'
                while matches!(self.peek(), Some(c) if c.is_ascii_alphabetic()) {
                    name.push(self.advance().unwrap().to_ascii_uppercase());
                }
            }
        }

        let abs_row = self.consume('$');

        // Collect trailing digits: could be cell-ref row (A1) or part of function name (LOG10, ATAN2)
        let mut trailing_digits = String::new();
        while matches!(self.peek(), Some(c) if c.is_ascii_digit()) {
            trailing_digits.push(self.advance().unwrap());
        }

        if abs_col || abs_row {
            // A '$' was seen anywhere in this token: it can only be a cell
            // reference, never a function name (Excel identifiers never
            // contain '$').
            if trailing_digits.is_empty() {
                return Err(format!("Expected row number after '{}'", name));
            }
            let col = col_letters_to_num(&name);
            let row: u32 = trailing_digits
                .parse()
                .map_err(|e: std::num::ParseIntError| e.to_string())?;
            let corner1_end = self.pos;
            return self.finish_ref((col, row, abs_col, abs_row, tok_start, corner1_end), None);
        }

        if !trailing_digits.is_empty() {
            // Look ahead past whitespace: if '(' follows, the digits belong to the function name
            let mut tmp = self.pos;
            while tmp < self.chars.len() && matches!(self.chars[tmp], ' ' | '\t') {
                tmp += 1;
            }
            // A dot-qualified name (e.g. "A.A0") is always a function, never a cell ref.
            let is_dotted = name.contains('.');
            if is_dotted || self.chars.get(tmp) == Some(&'(') {
                // Function name includes trailing digits (e.g. LOG10, ATAN2, MODE.MULT)
                name.push_str(&trailing_digits);
                // Fall through to function-call handling below
            } else {
                // Cell reference: name = column letters, trailing_digits = row number
                let col = col_letters_to_num(&name);
                let row: u32 = trailing_digits
                    .parse()
                    .map_err(|e: std::num::ParseIntError| e.to_string())?;
                let corner1_end = self.pos;
                return self.finish_ref((col, row, false, false, tok_start, corner1_end), None);
            }
        }

        // Function call: IDENT(...)
        self.skip_ws();
        if self.peek() == Some('(') {
            self.advance();
            let mut args = vec![];
            self.skip_ws();
            if self.peek() != Some(')') {
                args.push(self.parse_nested_expr()?);
                loop {
                    self.skip_ws();
                    if self.consume(',') {
                        self.skip_ws();
                        args.push(self.parse_nested_expr()?);
                    } else {
                        break;
                    }
                }
            }
            self.skip_ws();
            if !self.consume(')') {
                return Err(format!("Expected ')' after arguments of '{}'", name));
            }
            return Ok(FormulaExpr::FuncCall { name, args });
        }

        // Boolean literals
        match name.as_str() {
            "TRUE" => return Ok(FormulaExpr::Bool(true)),
            "FALSE" => return Ok(FormulaExpr::Bool(false)),
            _ => {}
        }

        // Bare identifier not matching any known pattern → name reference for LET/LAMBDA
        Ok(FormulaExpr::FuncCall { name, args: vec![] })
    }
}

fn col_letters_to_num(s: &str) -> u32 {
    s.chars().fold(0u32, |acc, c| {
        let digit = (c as u32).saturating_sub('A' as u32).saturating_add(1);
        acc.saturating_mul(26).saturating_add(digit)
    })
}

/// Parse an Excel formula string (with or without a leading `=`).
pub fn parse(formula: &str) -> Result<FormulaExpr, String> {
    parse_with_refs(formula).map(|(expr, _)| expr)
}

/// Parse an Excel formula string, also returning every cell/range reference
/// literally present in it (`RefOccurrence`), each with its exact char-offset
/// span. Spans are relative to `formula.trim().trim_start_matches('=')` --
/// the same normalization `parse` applies -- not to `formula` as passed in.
/// `CellContent::formula` doesn't consistently carry a leading `=` (an XLSX-
/// loaded formula never does; a VBA/Python-set one often does -- see
/// `xlsx_cell_xml`'s defensive `.trim().trim_start_matches('=')` in
/// `src/lib.rs`), so a caller that needs to splice a rewritten span back into
/// the original stored string must apply the same normalization first.
pub fn parse_with_refs(formula: &str) -> Result<(FormulaExpr, Vec<RefOccurrence>), String> {
    let input = formula.trim().trim_start_matches('=');
    if input.len() > MAX_FORMULA_BYTES {
        return Err(format!(
            "Formula is too long ({} bytes; maximum is {})",
            input.len(),
            MAX_FORMULA_BYTES
        ));
    }
    let mut p = FormulaParser::new(input);
    let expr = p.parse_expr()?;
    p.skip_ws();
    if p.pos < p.chars.len() {
        Err(format!(
            "Unexpected input at position {}: '{}'",
            p.pos,
            p.chars[p.pos..].iter().collect::<String>()
        ))
    } else {
        let mut nodes = 0usize;
        validate_expr_shape(&expr, 0, &mut nodes)?;
        Ok((expr, p.refs))
    }
}

fn validate_expr_shape(expr: &FormulaExpr, depth: usize, nodes: &mut usize) -> Result<(), String> {
    *nodes = (*nodes)
        .checked_add(1)
        .ok_or_else(|| "Formula AST node count overflows usize".to_string())?;
    if *nodes > MAX_FORMULA_NODES {
        return Err(format!(
            "Formula AST is too large (maximum is {} nodes)",
            MAX_FORMULA_NODES
        ));
    }
    if depth > MAX_FORMULA_DEPTH {
        return Err(format!(
            "Formula AST exceeds the maximum depth of {}",
            MAX_FORMULA_DEPTH
        ));
    }
    match expr {
        FormulaExpr::BinOp { lhs, rhs, .. } => {
            validate_expr_shape(lhs, depth + 1, nodes)?;
            validate_expr_shape(rhs, depth + 1, nodes)?;
        }
        FormulaExpr::UnaryMinus(inner) => validate_expr_shape(inner, depth + 1, nodes)?,
        FormulaExpr::FuncCall { args, .. } => {
            for arg in args {
                validate_expr_shape(arg, depth + 1, nodes)?;
            }
        }
        FormulaExpr::Number(_)
        | FormulaExpr::Str(_)
        | FormulaExpr::Bool(_)
        | FormulaExpr::CellRef { .. }
        | FormulaExpr::Range { .. } => {}
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    // 3.14 is an arbitrary decimal test value for number-literal parsing, not π.
    #[allow(clippy::approx_constant)]
    fn test_number() {
        assert_eq!(parse("=42").unwrap(), FormulaExpr::Number(42.0));
        assert_eq!(parse("3.14").unwrap(), FormulaExpr::Number(3.14));
    }

    #[test]
    fn test_string() {
        assert_eq!(
            parse("=\"hello\"").unwrap(),
            FormulaExpr::Str("hello".into())
        );
    }

    #[test]
    fn test_bool() {
        assert_eq!(parse("=TRUE").unwrap(), FormulaExpr::Bool(true));
        assert_eq!(parse("=FALSE").unwrap(), FormulaExpr::Bool(false));
    }

    #[test]
    fn test_cell_ref() {
        assert_eq!(
            parse("=A1").unwrap(),
            FormulaExpr::CellRef {
                col: 1,
                row: 1,
                abs_col: false,
                abs_row: false,
                sheet: None,
            }
        );
        assert_eq!(
            parse("=B3").unwrap(),
            FormulaExpr::CellRef {
                col: 2,
                row: 3,
                abs_col: false,
                abs_row: false,
                sheet: None,
            }
        );
        assert_eq!(
            parse("=AA1").unwrap(),
            FormulaExpr::CellRef {
                col: 27,
                row: 1,
                abs_col: false,
                abs_row: false,
                sheet: None,
            }
        );
    }

    #[test]
    fn test_range() {
        assert_eq!(
            parse("=A1:B10").unwrap(),
            FormulaExpr::Range {
                c1: 1,
                r1: 1,
                c2: 2,
                r2: 10,
                abs_c1: false,
                abs_r1: false,
                abs_c2: false,
                abs_r2: false,
                sheet: None,
            }
        );
    }

    #[test]
    fn test_absolute_cell_ref() {
        assert_eq!(
            parse("=$A$1").unwrap(),
            FormulaExpr::CellRef {
                col: 1,
                row: 1,
                abs_col: true,
                abs_row: true,
                sheet: None,
            }
        );
        assert_eq!(
            parse("=A$1").unwrap(),
            FormulaExpr::CellRef {
                col: 1,
                row: 1,
                abs_col: false,
                abs_row: true,
                sheet: None,
            }
        );
        assert_eq!(
            parse("=$A1").unwrap(),
            FormulaExpr::CellRef {
                col: 1,
                row: 1,
                abs_col: true,
                abs_row: false,
                sheet: None,
            }
        );
    }

    #[test]
    fn test_absolute_range_mixed_corners() {
        assert_eq!(
            parse("=$A1:B$10").unwrap(),
            FormulaExpr::Range {
                c1: 1,
                r1: 1,
                c2: 2,
                r2: 10,
                abs_c1: true,
                abs_r1: false,
                abs_c2: false,
                abs_r2: true,
                sheet: None,
            }
        );
        assert_eq!(
            parse("=$A$1:$B$10").unwrap(),
            FormulaExpr::Range {
                c1: 1,
                r1: 1,
                c2: 2,
                r2: 10,
                abs_c1: true,
                abs_r1: true,
                abs_c2: true,
                abs_r2: true,
                sheet: None,
            }
        );
    }

    #[test]
    fn test_absolute_ref_errors() {
        assert!(parse("=$1").is_err());
        assert!(parse("=$A").is_err());
        assert!(parse("=A1:$B").is_err());
    }

    #[test]
    fn test_refs_single_cell_span() {
        let (_, refs) = parse_with_refs("=A1+1").unwrap();
        assert_eq!(
            refs,
            vec![RefOccurrence::Cell {
                span: (0, 2),
                col: 1,
                row: 1,
                abs_col: false,
                abs_row: false,
                sheet: None,
            }]
        );
        // The span must index exactly "A1" in the normalized (post `=`-strip) text.
        let input = "A1+1";
        let (start, end) = (0, 2);
        assert_eq!(&input[start..end], "A1");
    }

    #[test]
    fn test_refs_absolute_cell_span() {
        let (_, refs) = parse_with_refs("=1+$A$1").unwrap();
        assert_eq!(
            refs,
            vec![RefOccurrence::Cell {
                span: (2, 6),
                col: 1,
                row: 1,
                abs_col: true,
                abs_row: true,
                sheet: None,
            }]
        );
        let input = "1+$A$1";
        assert_eq!(&input[2..6], "$A$1");
    }

    #[test]
    fn test_refs_range_spans() {
        let (_, refs) = parse_with_refs("=SUM($A1:B$10)").unwrap();
        assert_eq!(
            refs,
            vec![RefOccurrence::Range {
                span: (4, 12),
                c1: 1,
                r1: 1,
                abs_c1: true,
                abs_r1: false,
                c1_span: (4, 7),
                c2: 2,
                r2: 10,
                abs_c2: false,
                abs_r2: true,
                c2_span: (8, 12),
                sheet: None,
            }]
        );
        let input = "SUM($A1:B$10)";
        assert_eq!(&input[4..12], "$A1:B$10");
        assert_eq!(&input[4..7], "$A1");
        assert_eq!(&input[8..12], "B$10");
    }

    #[test]
    fn test_refs_multiple_occurrences_in_expression() {
        let (_, refs) = parse_with_refs("=A1+B2*SUM(C1:D2)").unwrap();
        let input = "A1+B2*SUM(C1:D2)";
        let span_of = |r: &RefOccurrence| match r {
            RefOccurrence::Cell { span, .. } => *span,
            RefOccurrence::Range { span, .. } => *span,
        };
        assert_eq!(refs.len(), 3);
        let (s0, e0) = span_of(&refs[0]);
        let (s1, e1) = span_of(&refs[1]);
        let (s2, e2) = span_of(&refs[2]);
        assert_eq!(&input[s0..e0], "A1");
        assert_eq!(&input[s1..e1], "B2");
        assert_eq!(&input[s2..e2], "C1:D2");
    }

    #[test]
    fn test_arithmetic() {
        assert_eq!(
            parse("=1+2*3").unwrap(),
            FormulaExpr::BinOp {
                op: BinOpKind::Add,
                lhs: Box::new(FormulaExpr::Number(1.0)),
                rhs: Box::new(FormulaExpr::BinOp {
                    op: BinOpKind::Mul,
                    lhs: Box::new(FormulaExpr::Number(2.0)),
                    rhs: Box::new(FormulaExpr::Number(3.0)),
                }),
            }
        );
    }

    #[test]
    fn test_function_call() {
        let expr = parse("=SUM(A1:A3)").unwrap();
        assert_eq!(
            expr,
            FormulaExpr::FuncCall {
                name: "SUM".into(),
                args: vec![FormulaExpr::Range {
                    c1: 1,
                    r1: 1,
                    c2: 1,
                    r2: 3,
                    abs_c1: false,
                    abs_r1: false,
                    abs_c2: false,
                    abs_r2: false,
                    sheet: None,
                }],
            }
        );
    }

    #[test]
    fn test_if_function() {
        let expr = parse("=IF(A1>0,B1,0)").unwrap();
        assert!(matches!(expr, FormulaExpr::FuncCall { ref name, .. } if name == "IF"));
    }

    #[test]
    fn test_concat() {
        let expr = parse("=\"A\"&\"B\"").unwrap();
        assert_eq!(
            expr,
            FormulaExpr::BinOp {
                op: BinOpKind::Concat,
                lhs: Box::new(FormulaExpr::Str("A".into())),
                rhs: Box::new(FormulaExpr::Str("B".into())),
            }
        );
    }

    #[test]
    fn test_unary_minus() {
        assert_eq!(
            parse("=-1").unwrap(),
            FormulaExpr::UnaryMinus(Box::new(FormulaExpr::Number(1.0)))
        );
    }

    #[test]
    fn test_parentheses() {
        assert_eq!(
            parse("=(1+2)*3").unwrap(),
            FormulaExpr::BinOp {
                op: BinOpKind::Mul,
                lhs: Box::new(FormulaExpr::BinOp {
                    op: BinOpKind::Add,
                    lhs: Box::new(FormulaExpr::Number(1.0)),
                    rhs: Box::new(FormulaExpr::Number(2.0)),
                }),
                rhs: Box::new(FormulaExpr::Number(3.0)),
            }
        );
    }

    #[test]
    fn test_dot_function_name() {
        let expr = parse("=MODE.MULT(1,2,2)").unwrap();
        assert!(matches!(expr, FormulaExpr::FuncCall { ref name, .. } if name == "MODE.MULT"));
    }

    #[test]
    fn test_function_name_with_digits() {
        // LOG10 and ATAN2 have digits in their names; must not be mistaken for cell references
        let expr = parse("=LOG10(100)").unwrap();
        assert!(matches!(expr, FormulaExpr::FuncCall { ref name, .. } if name == "LOG10"));

        let expr = parse("=ATAN2(1,1)").unwrap();
        assert!(matches!(expr, FormulaExpr::FuncCall { ref name, .. } if name == "ATAN2"));

        // Cell references with the same letter+digit pattern must still work
        assert_eq!(
            parse("=A1").unwrap(),
            FormulaExpr::CellRef {
                col: 1,
                row: 1,
                abs_col: false,
                abs_row: false,
                sheet: None,
            }
        );
        assert_eq!(
            parse("=B10").unwrap(),
            FormulaExpr::CellRef {
                col: 2,
                row: 10,
                abs_col: false,
                abs_row: false,
                sheet: None,
            }
        );
    }

    // ── 0.14.0-A2: sheet-qualified references ────────────────────────────

    #[test]
    fn test_qualified_cell_ref() {
        let expr = parse("=Sheet2!A1").unwrap();
        match expr {
            FormulaExpr::CellRef {
                col,
                row,
                abs_col,
                abs_row,
                sheet,
            } => {
                assert_eq!((col, row, abs_col, abs_row), (1, 1, false, false));
                let q = sheet.expect("expected a sheet qualifier");
                assert_eq!(q.normalized_name, "Sheet2");
                assert_eq!(q.raw_text, "Sheet2");
            }
            other => panic!("expected qualified CellRef, got {other:?}"),
        }
    }

    #[test]
    fn test_qualified_range_ref() {
        let expr = parse("=Sheet2!A1:B10").unwrap();
        match expr {
            FormulaExpr::Range {
                c1,
                r1,
                c2,
                r2,
                sheet,
                ..
            } => {
                assert_eq!((c1, r1, c2, r2), (1, 1, 2, 10));
                assert_eq!(sheet.unwrap().normalized_name, "Sheet2");
            }
            other => panic!("expected qualified Range, got {other:?}"),
        }
    }

    #[test]
    fn test_qualified_ref_with_mixed_dollar_signs() {
        let expr = parse("=Sheet2!$A1:B$10").unwrap();
        match expr {
            FormulaExpr::Range {
                abs_c1,
                abs_r1,
                abs_c2,
                abs_r2,
                sheet,
                ..
            } => {
                assert_eq!((abs_c1, abs_r1, abs_c2, abs_r2), (true, false, false, true));
                assert!(sheet.is_some());
            }
            other => panic!("expected Range, got {other:?}"),
        }
    }

    #[test]
    fn test_quoted_sheet_qualifier() {
        let expr = parse("='Sales 2026'!A1").unwrap();
        match expr {
            FormulaExpr::CellRef { sheet, .. } => {
                let q = sheet.unwrap();
                assert_eq!(q.normalized_name, "Sales 2026");
                assert_eq!(q.raw_text, "'Sales 2026'");
            }
            other => panic!("expected CellRef, got {other:?}"),
        }
    }

    #[test]
    fn test_quoted_sheet_qualifier_with_escaped_apostrophe() {
        let expr = parse("='Bob''s Data'!A1").unwrap();
        match expr {
            FormulaExpr::CellRef { sheet, .. } => {
                let q = sheet.unwrap();
                assert_eq!(q.normalized_name, "Bob's Data");
                assert_eq!(q.raw_text, "'Bob''s Data'");
            }
            other => panic!("expected CellRef, got {other:?}"),
        }
    }

    #[test]
    fn test_multiple_qualified_and_unqualified_refs_in_one_formula() {
        let expr = parse("=Sheet1!A1+A1+Sheet2!B2").unwrap();
        assert!(matches!(expr, FormulaExpr::BinOp { .. }));
    }

    #[test]
    fn test_sheet_qualifier_inside_a_string_literal_is_not_a_reference() {
        let (expr, refs) = parse_with_refs("=\"Sheet2!A1\"").unwrap();
        assert_eq!(expr, FormulaExpr::Str("Sheet2!A1".to_string()));
        assert!(refs.is_empty());
    }

    #[test]
    fn test_qualified_ref_inside_unknown_function_call() {
        let expr = parse("=UNKNOWNFUNC(Sheet2!A1)").unwrap();
        match expr {
            FormulaExpr::FuncCall { name, args } => {
                assert_eq!(name, "UNKNOWNFUNC");
                assert_eq!(args.len(), 1);
                assert!(matches!(
                    &args[0],
                    FormulaExpr::CellRef { sheet: Some(_), .. }
                ));
            }
            other => panic!("expected FuncCall, got {other:?}"),
        }
    }

    #[test]
    fn test_external_workbook_reference_is_a_parse_error() {
        assert!(parse("=[Book2.xlsx]Sheet1!A1").is_err());
    }

    #[test]
    fn test_3d_reference_is_a_parse_error() {
        assert!(parse("=Sheet1:Sheet3!A1").is_err());
    }

    #[test]
    fn test_qualifier_and_coordinate_spans_are_accurate() {
        let (_, refs) = parse_with_refs("=Sheet2!A1").unwrap();
        match &refs[0] {
            RefOccurrence::Cell { span, sheet, .. } => {
                let q = sheet.as_ref().unwrap();
                let input = "Sheet2!A1";
                assert_eq!(&input[q.raw_span.0..q.raw_span.1], "Sheet2");
                assert_eq!(&input[span.0..span.1], "A1");
            }
            other => panic!("expected Cell occurrence, got {other:?}"),
        }
    }

    #[test]
    fn test_non_ascii_sheet_name_qualifier() {
        let formula = "=売上!A1";
        let (_, refs) = parse_with_refs(formula).unwrap();
        match &refs[0] {
            RefOccurrence::Cell { sheet, .. } => {
                let q = sheet.as_ref().unwrap();
                assert_eq!(q.normalized_name, "売上");
                // Char-based slicing, not byte slicing: a Japanese sheet name
                // is multi-byte in UTF-8, so &str[..] byte-slicing at these
                // char-offsets would panic or mis-slice.
                let chars: Vec<char> = formula.trim_start_matches('=').chars().collect();
                let sliced: String = chars[q.raw_span.0..q.raw_span.1].iter().collect();
                assert_eq!(sliced, "売上");
            }
            other => panic!("expected Cell occurrence, got {other:?}"),
        }
    }

    #[test]
    fn formula_limits_reject_excessive_input_length() {
        let input = "1".repeat(MAX_FORMULA_BYTES + 1);
        let error = parse(&input).unwrap_err();
        assert!(error.contains("too long"));
    }

    #[test]
    fn formula_limits_reject_excessive_nesting_before_stack_growth() {
        let input = format!(
            "{}1{}",
            "(".repeat(MAX_FORMULA_DEPTH + 1),
            ")".repeat(MAX_FORMULA_DEPTH + 1)
        );
        let error = parse(&input).unwrap_err();
        assert!(error.contains("nesting"));
    }
}
