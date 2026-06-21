use pest::iterators::Pair;
use pest::Parser;
use pest_derive::Parser;

pub mod ast;
pub use ast::*;

#[derive(Parser)]
#[grammar = "parser/vba.pest"]
struct VbaParser;

pub fn parse(input: &str) -> Result<Program, pest::error::Error<Rule>> {
    let mut pairs = VbaParser::parse(Rule::program, input)?;
    let program_pair = pairs.next().expect("program rule must match");
    Ok(build_program(program_pair))
}

// ── Program / Sub ─────────────────────────────────────────────────────────────

fn build_program(pair: Pair<Rule>) -> Program {
    let subs = pair
        .into_inner()
        .filter_map(|p| if p.as_rule() == Rule::sub_def { Some(build_sub_def(p)) } else { None })
        .collect();
    Program { subs }
}

fn build_sub_def(pair: Pair<Rule>) -> SubDef {
    let mut name = String::new();
    let mut body = vec![];
    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::ident => name = p.as_str().to_lowercase(),
            _ => if let Some(s) = try_build_stmt(p) { body.push(s); },
        }
    }
    SubDef { name, body }
}

// ── Statement dispatch ────────────────────────────────────────────────────────

fn try_build_stmt(p: Pair<Rule>) -> Option<Stmt> {
    match p.as_rule() {
        Rule::assignment        => Some(build_assignment(p)),
        Rule::cell_write        => Some(build_cell_write(p)),
        Rule::app_calc          => Some(build_app_calc(p)),
        Rule::app_prop          => Some(build_app_prop(p)),
        Rule::range_write_stmt  => Some(build_range_write_stmt(p)),
        Rule::with_range_write  => Some(build_with_range_write(p)),
        Rule::range_copy_stmt   => Some(build_range_copy_stmt(p)),
        Rule::for_stmt          => Some(build_for_stmt(p)),
        Rule::if_stmt           => Some(build_if_stmt(p)),
        Rule::do_loop_stmt      => Some(build_do_loop_stmt(p)),
        Rule::select_case_stmt  => Some(build_select_case_stmt(p)),
        Rule::dim_stmt          => Some(Stmt::Dim),
        Rule::with_stmt         => Some(build_with_stmt(p)),
        Rule::with_cell_write   => Some(build_with_cell_write(p)),
        Rule::msgbox_stmt       => Some(build_msgbox_stmt(p)),
        _ => None,
    }
}

fn build_assignment(pair: Pair<Rule>) -> Stmt {
    let mut inner = pair.into_inner();
    let var   = inner.next().unwrap().as_str().to_lowercase();
    let value = build_expr(inner.next().unwrap());
    Stmt::Assignment { var, value }
}

fn build_cell_write(pair: Pair<Rule>) -> Stmt {
    let mut inner = pair.into_inner();
    let row   = build_expr(inner.next().unwrap());
    let col   = build_expr(inner.next().unwrap());
    let value = build_expr(inner.next().unwrap());
    Stmt::CellWrite { row, col, value }
}

fn build_app_calc(pair: Pair<Rule>) -> Stmt {
    let mode_str = pair.into_inner().next().unwrap().as_str().to_lowercase();
    let mode = if mode_str.contains("automatic") { CalcModeValue::Automatic } else { CalcModeValue::Manual };
    Stmt::SetCalcMode(mode)
}

fn build_for_stmt(pair: Pair<Rule>) -> Stmt {
    let mut children = pair.into_inner().peekable();

    children.next(); // for_kw
    let var  = children.next().unwrap().as_str().to_lowercase();
    let from = build_expr(children.next().unwrap());
    children.next(); // to_kw
    let to   = build_expr(children.next().unwrap());

    let step = if children.peek().map(|p| p.as_rule()) == Some(Rule::step_kw) {
        children.next(); // step_kw
        Some(build_expr(children.next().unwrap()))
    } else {
        None
    };

    let mut body = vec![];
    while let Some(p) = children.peek() {
        if p.as_rule() == Rule::next_kw { break; }
        let p = children.next().unwrap();
        if let Some(s) = try_build_stmt(p) { body.push(s); }
    }

    Stmt::For { var, from, to, step, body }
}

fn build_if_stmt(pair: Pair<Rule>) -> Stmt {
    let mut children = pair.into_inner().peekable();

    children.next(); // if_kw
    let condition = build_expr(children.next().unwrap());
    children.next(); // then_kw

    let mut then_body = vec![];
    while let Some(p) = children.peek() {
        match p.as_rule() {
            Rule::else_clause | Rule::end_if_kw => break,
            _ => { let p = children.next().unwrap(); if let Some(s) = try_build_stmt(p) { then_body.push(s); } }
        }
    }

    let else_body = if children.peek().map(|p| p.as_rule()) == Some(Rule::else_clause) {
        build_else_clause(children.next().unwrap())
    } else {
        vec![]
    };

    Stmt::If { condition, then_body, else_body }
}

fn build_msgbox_stmt(pair: Pair<Rule>) -> Stmt {
    let mut inner = pair.into_inner();
    inner.next(); // msgbox_kw
    let message = build_expr(inner.next().unwrap());
    Stmt::MsgBox { message }
}

fn build_app_prop(pair: Pair<Rule>) -> Stmt {
    let mut inner = pair.into_inner();
    let prop  = inner.next().unwrap().as_str().to_lowercase();
    let value = build_expr(inner.next().unwrap());
    Stmt::SetAppProp { prop, value }
}

fn build_range_write_stmt(pair: Pair<Rule>) -> Stmt {
    let mut inner = pair.into_inner();
    let addr_raw  = inner.next().unwrap().as_str(); // string_lit with quotes
    let addr      = addr_raw[1..addr_raw.len()-1].to_string();
    let prop      = inner.next().unwrap().as_str().to_lowercase();
    let is_formula = prop == "formula";
    let value     = build_expr(inner.next().unwrap());
    Stmt::RangeWrite { addr, is_formula, value }
}

fn build_with_range_write(pair: Pair<Rule>) -> Stmt {
    let mut inner = pair.into_inner();
    let addr_raw  = inner.next().unwrap().as_str();
    let addr      = addr_raw[1..addr_raw.len()-1].to_string();
    let prop      = inner.next().unwrap().as_str().to_lowercase();
    let is_formula = prop == "formula";
    let value     = build_expr(inner.next().unwrap());
    Stmt::RangeWrite { addr, is_formula, value }
}

fn build_range_copy_stmt(pair: Pair<Rule>) -> Stmt {
    let mut inner = pair.into_inner();
    let src_raw = inner.next().unwrap().as_str();
    let src = src_raw[1..src_raw.len()-1].to_string();
    let dst_raw = inner.next().unwrap().as_str();
    let dst = dst_raw[1..dst_raw.len()-1].to_string();
    Stmt::RangeCopy { src, dst }
}

fn build_else_clause(pair: Pair<Rule>) -> Vec<Stmt> {
    pair.into_inner()
        .filter_map(|p| if p.as_rule() == Rule::else_kw { None } else { try_build_stmt(p) })
        .collect()
}

// ── Do While / Do Until ──────────────────────────────────────────────────────

fn build_do_loop_stmt(pair: Pair<Rule>) -> Stmt {
    let mut children = pair.into_inner().peekable();
    children.next(); // do_kw

    let pre_cond = if children.peek().map(|p| p.as_rule()) == Some(Rule::do_condition) {
        Some(build_do_condition(children.next().unwrap()))
    } else {
        None
    };

    let mut body = vec![];
    while let Some(p) = children.peek() {
        if p.as_rule() == Rule::loop_kw { break; }
        let p = children.next().unwrap();
        if let Some(s) = try_build_stmt(p) { body.push(s); }
    }
    children.next(); // loop_kw

    let post_cond = if children.peek().map(|p| p.as_rule()) == Some(Rule::do_condition) {
        Some(build_do_condition(children.next().unwrap()))
    } else {
        None
    };

    Stmt::DoLoop { pre_cond, post_cond, body }
}

fn build_do_condition(pair: Pair<Rule>) -> (bool, Expr) {
    let mut inner = pair.into_inner();
    let kw = inner.next().unwrap();
    let is_until = kw.as_rule() == Rule::until_kw;
    let expr = build_expr(inner.next().unwrap());
    (is_until, expr)
}

// ── Select Case ───────────────────────────────────────────────────────────────

fn build_select_case_stmt(pair: Pair<Rule>) -> Stmt {
    let mut children = pair.into_inner().peekable();
    children.next(); // select_kw
    children.next(); // case_kw
    let expr = build_expr(children.next().unwrap());

    let mut cases = vec![];
    let mut else_body = vec![];

    while let Some(p) = children.peek() {
        match p.as_rule() {
            Rule::case_clause => {
                cases.push(build_case_clause(children.next().unwrap()));
            }
            Rule::case_else_clause => {
                else_body = build_case_else_clause(children.next().unwrap());
            }
            Rule::end_select_kw => break,
            _ => { children.next(); }
        }
    }

    Stmt::SelectCase { expr, cases, else_body }
}

fn build_case_clause(pair: Pair<Rule>) -> (Vec<CaseMatch>, Vec<Stmt>) {
    let mut inner = pair.into_inner().peekable();
    inner.next(); // case_kw

    let match_list = {
        let ml = inner.next().unwrap(); // case_match_list
        ml.into_inner().map(build_case_match).collect()
    };

    let body = inner.filter_map(try_build_stmt).collect();
    (match_list, body)
}

fn build_case_match(pair: Pair<Rule>) -> CaseMatch {
    let children: Vec<Pair<Rule>> = pair.into_inner().collect();
    if children.is_empty() {
        return CaseMatch::Value(Expr::Bool(false));
    }
    match children[0].as_rule() {
        Rule::is_kw => {
            let op = parse_cmp_op(children[1].as_str());
            let expr = build_expr(children[2].clone());
            CaseMatch::IsOp(op, expr)
        }
        _ if children.len() == 3 && children[1].as_rule() == Rule::to_kw => {
            CaseMatch::Range(build_expr(children[0].clone()), build_expr(children[2].clone()))
        }
        _ => CaseMatch::Value(build_expr(children[0].clone())),
    }
}

fn parse_cmp_op(s: &str) -> VbaBinOp {
    match s {
        "="  => VbaBinOp::Eq, "<>" => VbaBinOp::Ne,
        "<"  => VbaBinOp::Lt, "<=" => VbaBinOp::Le,
        ">"  => VbaBinOp::Gt, ">=" => VbaBinOp::Ge,
        other => panic!("unknown cmp_op: {}", other),
    }
}

fn build_case_else_clause(pair: Pair<Rule>) -> Vec<Stmt> {
    pair.into_inner()
        .filter_map(|p| match p.as_rule() {
            Rule::case_kw | Rule::else_kw => None,
            _ => try_build_stmt(p),
        })
        .collect()
}

// ── With ... End With ────────────────────────────────────────────────────────

fn build_with_stmt(pair: Pair<Rule>) -> Stmt {
    let mut inner = pair.into_inner().peekable();
    inner.next(); // with_kw
    inner.next(); // with_target (ignored)

    let mut body = vec![];
    while let Some(p) = inner.peek() {
        match p.as_rule() {
            Rule::end_with_kw => break,
            Rule::with_dot_stmt => { inner.next(); }
            _ => {
                let p = inner.next().unwrap();
                if let Some(s) = try_build_stmt(p) { body.push(s); }
            }
        }
    }
    Stmt::With { body }
}

fn build_with_cell_write(pair: Pair<Rule>) -> Stmt {
    let mut inner = pair.into_inner();
    let row   = build_expr(inner.next().unwrap());
    let col   = build_expr(inner.next().unwrap());
    let value = build_expr(inner.next().unwrap());
    Stmt::CellWrite { row, col, value }
}

// ── Expression builder ────────────────────────────────────────────────────────

fn build_expr(pair: Pair<Rule>) -> Expr {
    match pair.as_rule() {
        Rule::comparison => build_comparison(pair),
        Rule::additive   => build_additive(pair),
        Rule::term       => build_term(pair),
        Rule::unary_minus => {
            Expr::UnaryMinus(Box::new(build_expr(pair.into_inner().next().unwrap())))
        }
        Rule::unary_not => {
            let mut inner = pair.into_inner();
            inner.next(); // not_kw
            Expr::UnaryNot(Box::new(build_expr(inner.next().unwrap())))
        }
        Rule::integer => Expr::Integer(pair.as_str().parse().unwrap()),
        Rule::float   => Expr::Float(pair.as_str().parse().unwrap()),
        Rule::bool_lit => {
            Expr::Bool(pair.as_str().to_lowercase() == "true")
        }
        Rule::string_lit => {
            let s = pair.as_str();
            Expr::Str(s[1..s.len() - 1].to_string())
        }
        Rule::var_ref => {
            Expr::Var(pair.into_inner().next().unwrap().as_str().to_lowercase())
        }
        Rule::cell_read => {
            let mut inner = pair.into_inner();
            let row = build_expr(inner.next().unwrap());
            let col = build_expr(inner.next().unwrap());
            Expr::CellRead { row: Box::new(row), col: Box::new(col) }
        }
        Rule::func_call => {
            let mut inner = pair.into_inner();
            let name = inner.next().unwrap().as_str().to_lowercase();
            let args = inner.map(build_expr).collect();
            Expr::FuncCall { name, args }
        }
        Rule::range_read => {
            let raw = pair.into_inner().next().unwrap().as_str();
            let addr = raw[1..raw.len()-1].to_uppercase();
            Expr::RangeRead { addr }
        }
        Rule::rows_count_expr => Expr::RowsCount,
        Rule::cols_count_expr => Expr::ColsCount,
        Rule::cells_end_expr  => {
            let mut inner = pair.into_inner();
            let row = build_expr(inner.next().unwrap());
            let col = build_expr(inner.next().unwrap());
            let dir_pair = inner.next().unwrap(); // xl_dir_kw
            let dir = match dir_pair.as_str().to_lowercase().as_str() {
                "xlup"      => XlDir::Up,
                "xldown"    => XlDir::Down,
                "xltoleft"  => XlDir::Left,
                "xltoright" => XlDir::Right,
                other => panic!("unknown xl_dir: {}", other),
            };
            let prop_pair = inner.next().unwrap(); // end_prop_kw
            let prop = if prop_pair.as_str().to_lowercase() == "row" { XlEndProp::Row } else { XlEndProp::Column };
            Expr::CellsEndProp { row: Box::new(row), col: Box::new(col), dir, prop }
        }
        _ => panic!("unexpected rule in build_expr: {:?}", pair.as_rule()),
    }
}

fn build_comparison(pair: Pair<Rule>) -> Expr {
    let mut inner = pair.into_inner();
    let mut lhs = build_expr(inner.next().unwrap());
    loop {
        let Some(op_pair) = inner.next() else { break };
        let rhs = build_expr(inner.next().unwrap());
        let op = parse_cmp_op(op_pair.as_str());
        lhs = Expr::BinOp { op, lhs: Box::new(lhs), rhs: Box::new(rhs) };
    }
    lhs
}

fn build_additive(pair: Pair<Rule>) -> Expr {
    let mut inner = pair.into_inner();
    let mut lhs = build_expr(inner.next().unwrap());
    loop {
        let Some(op_pair) = inner.next() else { break };
        let rhs = build_expr(inner.next().unwrap());
        let op = match op_pair.as_str() {
            "+" => VbaBinOp::Add, "-" => VbaBinOp::Sub, "&" => VbaBinOp::Concat,
            s   => panic!("unknown add_op: {}", s),
        };
        lhs = Expr::BinOp { op, lhs: Box::new(lhs), rhs: Box::new(rhs) };
    }
    lhs
}

fn build_term(pair: Pair<Rule>) -> Expr {
    let mut inner = pair.into_inner();
    let mut lhs = build_expr(inner.next().unwrap());
    loop {
        let Some(op_pair) = inner.next() else { break };
        let rhs = build_expr(inner.next().unwrap());
        let op = match op_pair.as_str() {
            "*" => VbaBinOp::Mul, "/" => VbaBinOp::Div,
            s   => panic!("unknown mul_op: {}", s),
        };
        lhs = Expr::BinOp { op, lhs: Box::new(lhs), rhs: Box::new(rhs) };
    }
    lhs
}

// ── Tests ─────────────────────────────────────────────────────────────────────

#[cfg(test)]
mod tests {
    use super::*;

    fn parse_body(code: &str) -> Vec<Stmt> {
        parse(code).unwrap().subs.into_iter().next().unwrap().body
    }

    #[test]
    fn test_empty_sub() {
        let prog = parse("Sub MySub()\nEnd Sub\n").unwrap();
        assert_eq!(prog.subs[0].name, "mysub");
        assert!(prog.subs[0].body.is_empty());
    }

    #[test]
    fn test_variable_assignment_integer() {
        let body = parse_body("Sub MySub()\n    a = 10\nEnd Sub\n");
        assert_eq!(body, vec![Stmt::Assignment { var: "a".into(), value: Expr::Integer(10) }]);
    }

    #[test]
    fn test_variable_assignment_float() {
        let body = parse_body("Sub MySub()\n    x = 3.14\nEnd Sub\n");
        assert_eq!(body, vec![Stmt::Assignment { var: "x".into(), value: Expr::Float(3.14) }]);
    }

    #[test]
    fn test_variable_assignment_string() {
        let body = parse_body("Sub MySub()\n    msg = \"hello\"\nEnd Sub\n");
        assert_eq!(body, vec![Stmt::Assignment { var: "msg".into(), value: Expr::Str("hello".into()) }]);
    }

    #[test]
    fn test_cell_write_integer() {
        let body = parse_body("Sub MySub()\n    Cells(1, 1).Value = 42\nEnd Sub\n");
        assert_eq!(body, vec![Stmt::CellWrite {
            row: Expr::Integer(1), col: Expr::Integer(1), value: Expr::Integer(42),
        }]);
    }

    #[test]
    fn test_cell_write_var_ref() {
        let body = parse_body("Sub MySub()\n    a = 10\n    Cells(1, 1).Value = a\nEnd Sub\n");
        assert_eq!(body[1], Stmt::CellWrite {
            row: Expr::Integer(1), col: Expr::Integer(1), value: Expr::Var("a".into()),
        });
    }

    #[test]
    fn test_case_insensitive_keywords() {
        let prog = parse("SUB MYSUB()\n    A = 10\n    CELLS(1, 1).VALUE = A\nEND SUB\n").unwrap();
        assert_eq!(prog.subs[0].name, "mysub");
    }

    #[test]
    fn test_comment_ignored() {
        let body = parse_body("Sub MySub()\n    ' comment\n    a = 10\nEnd Sub\n");
        assert_eq!(body, vec![Stmt::Assignment { var: "a".into(), value: Expr::Integer(10) }]);
    }

    #[test]
    fn test_multiple_subs() {
        let prog = parse("Sub First()\n    a = 1\nEnd Sub\n\nSub Second()\n    b = 2\nEnd Sub\n").unwrap();
        assert_eq!(prog.subs.len(), 2);
    }

    #[test]
    fn test_arithmetic_expr() {
        let body = parse_body("Sub MySub()\n    a = 1 + 2\nEnd Sub\n");
        assert_eq!(body, vec![Stmt::Assignment {
            var: "a".into(),
            value: Expr::BinOp {
                op: VbaBinOp::Add,
                lhs: Box::new(Expr::Integer(1)),
                rhs: Box::new(Expr::Integer(2)),
            },
        }]);
    }

    #[test]
    fn test_precedence_mul_over_add() {
        let body = parse_body("Sub MySub()\n    a = 1 + 2 * 3\nEnd Sub\n");
        assert_eq!(body, vec![Stmt::Assignment {
            var: "a".into(),
            value: Expr::BinOp {
                op: VbaBinOp::Add,
                lhs: Box::new(Expr::Integer(1)),
                rhs: Box::new(Expr::BinOp {
                    op: VbaBinOp::Mul,
                    lhs: Box::new(Expr::Integer(2)),
                    rhs: Box::new(Expr::Integer(3)),
                }),
            },
        }]);
    }

    #[test]
    fn test_for_loop() {
        let body = parse_body("Sub MySub()\n    For i = 1 To 3\n        a = i\n    Next i\nEnd Sub\n");
        assert!(matches!(body[0], Stmt::For { .. }));
        if let Stmt::For { var, from, to, step, body } = &body[0] {
            assert_eq!(var, "i");
            assert_eq!(*from, Expr::Integer(1));
            assert_eq!(*to, Expr::Integer(3));
            assert!(step.is_none());
            assert_eq!(body.len(), 1);
        }
    }

    #[test]
    fn test_for_loop_step() {
        let body = parse_body("Sub MySub()\n    For i = 0 To 10 Step 2\n        a = i\n    Next i\nEnd Sub\n");
        if let Stmt::For { step, .. } = &body[0] {
            assert_eq!(*step, Some(Expr::Integer(2)));
        }
    }

    #[test]
    fn test_if_no_else() {
        let body = parse_body("Sub MySub()\n    If a > 0 Then\n        b = 1\n    End If\nEnd Sub\n");
        assert!(matches!(&body[0], Stmt::If { else_body, .. } if else_body.is_empty()));
    }

    #[test]
    fn test_if_with_else() {
        let body = parse_body("Sub MySub()\n    If a > 0 Then\n        b = 1\n    Else\n        b = 0\n    End If\nEnd Sub\n");
        if let Stmt::If { then_body, else_body, .. } = &body[0] {
            assert_eq!(then_body.len(), 1);
            assert_eq!(else_body.len(), 1);
        }
    }

    #[test]
    fn test_comparison_expr() {
        let body = parse_body("Sub MySub()\n    x = a > 5\nEnd Sub\n");
        assert_eq!(body, vec![Stmt::Assignment {
            var: "x".into(),
            value: Expr::BinOp {
                op: VbaBinOp::Gt,
                lhs: Box::new(Expr::Var("a".into())),
                rhs: Box::new(Expr::Integer(5)),
            },
        }]);
    }

    #[test]
    fn test_unary_minus() {
        let body = parse_body("Sub MySub()\n    a = -1\nEnd Sub\n");
        assert_eq!(body, vec![Stmt::Assignment {
            var: "a".into(),
            value: Expr::UnaryMinus(Box::new(Expr::Integer(1))),
        }]);
    }

    #[test]
    fn test_do_while_loop() {
        let body = parse_body("Sub MySub()\n    x = 0\n    Do While x < 3\n        x = x + 1\n    Loop\nEnd Sub\n");
        assert!(matches!(&body[1], Stmt::DoLoop { pre_cond: Some((false, _)), .. }));
    }

    #[test]
    fn test_do_until_loop() {
        let body = parse_body("Sub MySub()\n    x = 0\n    Do Until x >= 3\n        x = x + 1\n    Loop\nEnd Sub\n");
        assert!(matches!(&body[1], Stmt::DoLoop { pre_cond: Some((true, _)), .. }));
    }

    #[test]
    fn test_do_loop_while() {
        let body = parse_body("Sub MySub()\n    x = 0\n    Do\n        x = x + 1\n    Loop While x < 3\nEnd Sub\n");
        assert!(matches!(&body[1], Stmt::DoLoop { pre_cond: None, post_cond: Some((false, _)), .. }));
    }

    #[test]
    fn test_select_case() {
        let body = parse_body("Sub MySub()\n    Select Case x\n        Case 1\n            a = 1\n        Case 2, 3\n            a = 23\n        Case Else\n            a = 0\n    End Select\nEnd Sub\n");
        assert!(matches!(&body[0], Stmt::SelectCase { .. }));
        if let Stmt::SelectCase { cases, else_body, .. } = &body[0] {
            assert_eq!(cases.len(), 2);
            assert_eq!(else_body.len(), 1);
        }
    }

    #[test]
    fn test_dim_is_noop() {
        let body = parse_body("Sub MySub()\n    Dim x As Integer\n    x = 42\nEnd Sub\n");
        assert_eq!(body[0], Stmt::Dim);
        assert!(matches!(&body[1], Stmt::Assignment { .. }));
    }

    #[test]
    fn test_with_block() {
        let body = parse_body("Sub MySub()\n    With Sheet1\n        .Cells(1, 1).Value = 99\n    End With\nEnd Sub\n");
        assert!(matches!(&body[0], Stmt::With { body } if body.len() == 1));
    }

    #[test]
    fn test_func_call_in_expr() {
        let body = parse_body("Sub MySub()\n    a = Len(\"hello\")\nEnd Sub\n");
        assert!(matches!(&body[0], Stmt::Assignment { value: Expr::FuncCall { name, .. }, .. } if name == "len"));
    }

    #[test]
    fn test_bool_literal() {
        let body = parse_body("Sub MySub()\n    a = True\n    b = False\nEnd Sub\n");
        assert_eq!(body[0], Stmt::Assignment { var: "a".into(), value: Expr::Bool(true) });
        assert_eq!(body[1], Stmt::Assignment { var: "b".into(), value: Expr::Bool(false) });
    }

    #[test]
    fn test_unary_not() {
        let body = parse_body("Sub MySub()\n    a = Not True\nEnd Sub\n");
        assert!(matches!(&body[0], Stmt::Assignment { value: Expr::UnaryNot(_), .. }));
    }
}
