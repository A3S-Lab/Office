use super::*;
use crate::spreadsheet_formula::lexer;

fn parse_formula(source: &str) -> SpreadsheetFormula {
    let tokens = lexer::lex(source).unwrap();
    parse(source, tokens).unwrap()
}

#[test]
fn parser_applies_excel_operator_precedence_and_right_associative_power() {
    let formula = parse_formula("-2^3^4%+5*6&\"x\"");
    let SpreadsheetFormulaExpressionKind::Binary {
        operator: SpreadsheetFormulaBinaryOperator::Concatenate,
        left,
        right,
    } = formula.root.kind
    else {
        panic!("expected concatenation root");
    };
    assert!(matches!(
        right.kind,
        SpreadsheetFormulaExpressionKind::Literal(SpreadsheetFormulaLiteral::Text(_))
    ));
    let SpreadsheetFormulaExpressionKind::Binary {
        operator: SpreadsheetFormulaBinaryOperator::Add,
        left: power,
        right: multiply,
    } = left.kind
    else {
        panic!("expected additive expression");
    };
    assert!(matches!(
        multiply.kind,
        SpreadsheetFormulaExpressionKind::Binary {
            operator: SpreadsheetFormulaBinaryOperator::Multiply,
            ..
        }
    ));
    let SpreadsheetFormulaExpressionKind::Binary {
        operator: SpreadsheetFormulaBinaryOperator::Power,
        left: negative,
        right: nested_power,
    } = power.kind
    else {
        panic!("expected power expression");
    };
    assert!(matches!(
        negative.kind,
        SpreadsheetFormulaExpressionKind::Unary {
            operator: SpreadsheetFormulaUnaryOperator::Negative,
            ..
        }
    ));
    assert!(matches!(
        nested_power.kind,
        SpreadsheetFormulaExpressionKind::Binary {
            operator: SpreadsheetFormulaBinaryOperator::Power,
            ..
        }
    ));
}

#[test]
fn parser_distinguishes_function_arguments_from_reference_unions() {
    let formula = parse_formula("SUM(A1:B2,(C1:C2,D1:D2),,TRUE,#N/A)");
    let SpreadsheetFormulaExpressionKind::FunctionCall { arguments, .. } = formula.root.kind else {
        panic!("expected function call");
    };
    assert_eq!(arguments.len(), 5);
    assert!(arguments[2].is_none());
    let union = arguments[1].as_ref().unwrap();
    let SpreadsheetFormulaExpressionKind::Parenthesized(inner) = &union.kind else {
        panic!("expected parenthesized union");
    };
    assert!(matches!(
        inner.kind,
        SpreadsheetFormulaExpressionKind::Binary {
            operator: SpreadsheetFormulaBinaryOperator::Union,
            ..
        }
    ));
}

#[test]
fn parser_types_cell_column_row_and_sheet_qualified_references() {
    let formula = parse_formula("('Q1 Data'!$A$1:B2,[Book.xlsx]Data!C:C) Sheet1!1:$3");
    assert!(matches!(
        formula.root.kind,
        SpreadsheetFormulaExpressionKind::Binary {
            operator: SpreadsheetFormulaBinaryOperator::Intersection,
            ..
        }
    ));
}

#[test]
fn parser_accepts_rectangular_arrays_and_rejects_invalid_reference_operators() {
    assert!(matches!(
        parse_formula("{1,-2;TRUE,#N/A}").root.kind,
        SpreadsheetFormulaExpressionKind::Array { .. }
    ));
    for source in ["1 2", "A1::B2", "A:1", "$A+1", "{1,2;3}"] {
        let tokens = lexer::lex(source).unwrap();
        assert!(parse(source, tokens).is_err(), "{source}");
    }
}
