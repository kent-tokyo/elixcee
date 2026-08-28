use super::ast::{BinOpKind, FormulaExpr};

/// One cell/range reference as it literally appears in formula text, with its
/// exact char-offset span (relative to the normalized input `parse_with_refs`
/// was called on -- see that function's doc comment). Used by the reference
/// rewriter (`super::rewrite`) to patch only the substring that actually
/// changed on a structural edit, leaving everything else in the formula --
/// operators, function names, number/string literals, whitespace -- byte-for-
/// byte untouched.
#[derive(Debug, Clone, PartialEq)]
pub enum RefOccurrence {
    Cell {
        span: (usize, usize),
        col: u32,
        row: u32,
        abs_col: bool,
        abs_row: bool,
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
    },
}

pub struct FormulaParser {
    chars: Vec<char>,
    pos: usize,
    refs: Vec<RefOccurrence>,
}

impl FormulaParser {
    fn new(input: &str) -> Self {
        FormulaParser {
            chars: input.chars().collect(),
            pos: 0,
            refs: Vec::new(),
        }
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
        match self.peek() {
            Some('(') => {
                self.advance();
                let expr = self.parse_expr()?;
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

    /// Shared tail for both the `$`-flagged and bare-reference parse paths:
    /// given the already-parsed first corner, check for a trailing `:SECOND`
    /// range and record a `RefOccurrence` (`Cell` or `Range`) for whichever
    /// this turns out to be, alongside building the `FormulaExpr`.
    fn finish_ref(
        &mut self,
        tok_start: usize,
        corner1_end: usize,
        col: u32,
        row: u32,
        abs_col: bool,
        abs_row: bool,
    ) -> Result<FormulaExpr, String> {
        self.skip_ws();
        if self.peek() == Some(':') {
            self.advance();
            self.skip_ws();
            let (c2, r2, abs_c2, abs_r2, c2_start, c2_end) = self.parse_ref_corner()?;
            self.refs.push(RefOccurrence::Range {
                span: (tok_start, c2_end),
                c1: col,
                r1: row,
                abs_c1: abs_col,
                abs_r1: abs_row,
                c1_span: (tok_start, corner1_end),
                c2,
                r2,
                abs_c2,
                abs_r2,
                c2_span: (c2_start, c2_end),
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
            });
        }
        self.refs.push(RefOccurrence::Cell {
            span: (tok_start, corner1_end),
            col,
            row,
            abs_col,
            abs_row,
        });
        Ok(FormulaExpr::CellRef {
            col,
            row,
            abs_col,
            abs_row,
        })
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
            return self.finish_ref(tok_start, corner1_end, col, row, abs_col, abs_row);
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
                return self.finish_ref(tok_start, corner1_end, col, row, false, false);
            }
        }

        // Function call: IDENT(...)
        self.skip_ws();
        if self.peek() == Some('(') {
            self.advance();
            let mut args = vec![];
            self.skip_ws();
            if self.peek() != Some(')') {
                args.push(self.parse_expr()?);
                loop {
                    self.skip_ws();
                    if self.consume(',') {
                        self.skip_ws();
                        args.push(self.parse_expr()?);
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
        Ok((expr, p.refs))
    }
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
                abs_row: false
            }
        );
        assert_eq!(
            parse("=B3").unwrap(),
            FormulaExpr::CellRef {
                col: 2,
                row: 3,
                abs_col: false,
                abs_row: false
            }
        );
        assert_eq!(
            parse("=AA1").unwrap(),
            FormulaExpr::CellRef {
                col: 27,
                row: 1,
                abs_col: false,
                abs_row: false
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
                abs_row: true
            }
        );
        assert_eq!(
            parse("=A$1").unwrap(),
            FormulaExpr::CellRef {
                col: 1,
                row: 1,
                abs_col: false,
                abs_row: true
            }
        );
        assert_eq!(
            parse("=$A1").unwrap(),
            FormulaExpr::CellRef {
                col: 1,
                row: 1,
                abs_col: true,
                abs_row: false
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
                abs_row: false
            }
        );
        assert_eq!(
            parse("=B10").unwrap(),
            FormulaExpr::CellRef {
                col: 2,
                row: 10,
                abs_col: false,
                abs_row: false
            }
        );
    }
}
