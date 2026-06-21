use super::ast::{BinOpKind, FormulaExpr};

pub struct FormulaParser {
    chars: Vec<char>,
    pos: usize,
}

impl FormulaParser {
    fn new(input: &str) -> Self {
        FormulaParser { chars: input.chars().collect(), pos: 0 }
    }

    fn peek(&self) -> Option<char> {
        self.chars.get(self.pos).copied()
    }

    fn advance(&mut self) -> Option<char> {
        let c = self.chars.get(self.pos).copied();
        if c.is_some() { self.pos += 1; }
        c
    }

    fn skip_ws(&mut self) {
        while matches!(self.peek(), Some(' ') | Some('\t')) {
            self.pos += 1;
        }
    }

    fn consume(&mut self, c: char) -> bool {
        if self.peek() == Some(c) { self.pos += 1; true } else { false }
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
                    if self.consume('>') { BinOpKind::Ne }
                    else if self.consume('=') { BinOpKind::Le }
                    else { BinOpKind::Lt }
                }
                Some('>') => {
                    self.advance();
                    if self.consume('=') { BinOpKind::Ge } else { BinOpKind::Gt }
                }
                Some('=') => { self.advance(); BinOpKind::Eq }
                _ => break,
            };
            self.skip_ws();
            let rhs = self.parse_concat()?;
            lhs = FormulaExpr::BinOp { op, lhs: Box::new(lhs), rhs: Box::new(rhs) };
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
                lhs = FormulaExpr::BinOp { op: BinOpKind::Concat, lhs: Box::new(lhs), rhs: Box::new(rhs) };
            } else { break; }
        }
        Ok(lhs)
    }

    fn parse_additive(&mut self) -> Result<FormulaExpr, String> {
        let mut lhs = self.parse_multiplicative()?;
        loop {
            self.skip_ws();
            let op = match self.peek() {
                Some('+') => { self.advance(); BinOpKind::Add }
                Some('-') => { self.advance(); BinOpKind::Sub }
                _ => break,
            };
            self.skip_ws();
            let rhs = self.parse_multiplicative()?;
            lhs = FormulaExpr::BinOp { op, lhs: Box::new(lhs), rhs: Box::new(rhs) };
        }
        Ok(lhs)
    }

    fn parse_multiplicative(&mut self) -> Result<FormulaExpr, String> {
        let mut lhs = self.parse_unary()?;
        loop {
            self.skip_ws();
            let op = match self.peek() {
                Some('*') => { self.advance(); BinOpKind::Mul }
                Some('/') => { self.advance(); BinOpKind::Div }
                _ => break,
            };
            self.skip_ws();
            let rhs = self.parse_unary()?;
            lhs = FormulaExpr::BinOp { op, lhs: Box::new(lhs), rhs: Box::new(rhs) };
        }
        Ok(lhs)
    }

    fn parse_unary(&mut self) -> Result<FormulaExpr, String> {
        self.skip_ws();
        if self.peek() == Some('-') {
            self.advance();
            Ok(FormulaExpr::UnaryMinus(Box::new(self.parse_primary()?)))
        } else {
            if self.peek() == Some('+') { self.advance(); }
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
                if !self.consume(')') { return Err("Expected ')'".into()); }
                Ok(expr)
            }
            Some('"') => self.parse_string(),
            Some(c) if c.is_ascii_digit() => self.parse_number(),
            Some(c) if c.is_ascii_alphabetic() => self.parse_ident_or_ref(),
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
                    if self.peek() == Some('"') { self.advance(); s.push('"'); }
                    else { break; }
                }
                Some(c) => { s.push(c); self.advance(); }
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
        s.parse::<f64>().map(FormulaExpr::Number).map_err(|e| e.to_string())
    }

    fn parse_ident_or_ref(&mut self) -> Result<FormulaExpr, String> {
        let mut name = String::new();
        while matches!(self.peek(), Some(c) if c.is_ascii_alphabetic()) {
            name.push(self.advance().unwrap().to_ascii_uppercase());
        }
        // Support dot-separated function names (e.g. MODE.MULT, NETWORKDAYS.INTL)
        while self.peek() == Some('.')
            && matches!(self.chars.get(self.pos + 1), Some(c) if c.is_ascii_alphabetic())
        {
            name.push(self.advance().unwrap()); // consume '.'
            while matches!(self.peek(), Some(c) if c.is_ascii_alphabetic()) {
                name.push(self.advance().unwrap().to_ascii_uppercase());
            }
        }

        // Cell reference: alpha letters followed by digits (e.g., A1, AA10)
        if matches!(self.peek(), Some(c) if c.is_ascii_digit()) {
            let col = col_letters_to_num(&name);
            let mut row_s = String::new();
            while matches!(self.peek(), Some(c) if c.is_ascii_digit()) {
                row_s.push(self.advance().unwrap());
            }
            let row: u32 = row_s.parse().map_err(|e: std::num::ParseIntError| e.to_string())?;
            self.skip_ws();
            // Range: A1:B10
            if self.peek() == Some(':') {
                self.advance();
                self.skip_ws();
                let mut name2 = String::new();
                while matches!(self.peek(), Some(c) if c.is_ascii_alphabetic()) {
                    name2.push(self.advance().unwrap().to_ascii_uppercase());
                }
                let col2 = col_letters_to_num(&name2);
                let mut row2_s = String::new();
                while matches!(self.peek(), Some(c) if c.is_ascii_digit()) {
                    row2_s.push(self.advance().unwrap());
                }
                let row2: u32 = row2_s.parse().map_err(|e: std::num::ParseIntError| e.to_string())?;
                return Ok(FormulaExpr::Range { c1: col, r1: row, c2: col2, r2: row2 });
            }
            return Ok(FormulaExpr::CellRef { col, row });
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
                    if self.consume(',') { self.skip_ws(); args.push(self.parse_expr()?); }
                    else { break; }
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

        Err(format!("Unknown identifier: '{}'", name))
    }
}

fn col_letters_to_num(s: &str) -> u32 {
    s.chars().fold(0u32, |acc, c| acc * 26 + (c as u32 - 'A' as u32 + 1))
}

/// Parse an Excel formula string (with or without a leading `=`).
pub fn parse(formula: &str) -> Result<FormulaExpr, String> {
    let input = formula.trim().trim_start_matches('=');
    let mut p = FormulaParser::new(input);
    let expr = p.parse_expr()?;
    p.skip_ws();
    if p.pos < p.chars.len() {
        Err(format!("Unexpected input at position {}: '{}'", p.pos, &p.chars[p.pos..].iter().collect::<String>()))
    } else {
        Ok(expr)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_number() {
        assert_eq!(parse("=42").unwrap(), FormulaExpr::Number(42.0));
        assert_eq!(parse("3.14").unwrap(), FormulaExpr::Number(3.14));
    }

    #[test]
    fn test_string() {
        assert_eq!(parse("=\"hello\"").unwrap(), FormulaExpr::Str("hello".into()));
    }

    #[test]
    fn test_bool() {
        assert_eq!(parse("=TRUE").unwrap(), FormulaExpr::Bool(true));
        assert_eq!(parse("=FALSE").unwrap(), FormulaExpr::Bool(false));
    }

    #[test]
    fn test_cell_ref() {
        assert_eq!(parse("=A1").unwrap(), FormulaExpr::CellRef { col: 1, row: 1 });
        assert_eq!(parse("=B3").unwrap(), FormulaExpr::CellRef { col: 2, row: 3 });
        assert_eq!(parse("=AA1").unwrap(), FormulaExpr::CellRef { col: 27, row: 1 });
    }

    #[test]
    fn test_range() {
        assert_eq!(
            parse("=A1:B10").unwrap(),
            FormulaExpr::Range { c1: 1, r1: 1, c2: 2, r2: 10 }
        );
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
                args: vec![FormulaExpr::Range { c1: 1, r1: 1, c2: 1, r2: 3 }],
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
}
