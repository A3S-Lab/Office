use a3s_use_core::{UseError, UseResult};

use crate::discovery::office_error;
use crate::spreadsheet_reference::{column_name, MAX_COLUMNS, MAX_ROWS};

mod ast;
mod evaluate;
mod graph;
mod lexer;
mod parser;
mod registry;
mod structured_reference;
mod value;

pub use ast::{
    SpreadsheetFormula, SpreadsheetFormulaBinaryOperator, SpreadsheetFormulaErrorLiteral,
    SpreadsheetFormulaExpression, SpreadsheetFormulaExpressionKind, SpreadsheetFormulaLiteral,
    SpreadsheetFormulaPostfixOperator, SpreadsheetFormulaQualifier, SpreadsheetFormulaReference,
    SpreadsheetFormulaReferenceKind, SpreadsheetFormulaSpan, SpreadsheetFormulaUnaryOperator,
    MAX_SPREADSHEET_FORMULA_CHARACTERS, MAX_SPREADSHEET_FORMULA_DEPTH,
    MAX_SPREADSHEET_FORMULA_NODES, MAX_SPREADSHEET_FORMULA_REFERENCE_AREAS,
};
pub use evaluate::{
    MAX_SPREADSHEET_FORMULA_CALCULATION_TEXT_BYTES, MAX_SPREADSHEET_FORMULA_SPILL_CELLS,
    MAX_SPREADSHEET_FORMULA_TEXT_BYTES,
};
pub use graph::{
    SpreadsheetFormulaCell, SpreadsheetFormulaDependencyGraph, SpreadsheetFormulaDependencyNode,
    SpreadsheetFormulaUnresolvedReference, SpreadsheetFormulaUnresolvedReferenceKind,
    MAX_SPREADSHEET_FORMULA_CELLS, MAX_SPREADSHEET_FORMULA_DEPENDENCIES,
    MAX_SPREADSHEET_FORMULA_REFERENCE_VISITS,
};
pub use registry::{
    SpreadsheetFormulaFunctionDefinition, SpreadsheetFormulaFunctionRegistry,
    SpreadsheetFormulaFunctionReturnKind, SpreadsheetFormulaFunctionVolatility,
};
pub(crate) use structured_reference::{
    LocalStructuredReferenceContext, StructuredReferenceRewritePlan,
    StructuredReferenceRewriteResult,
};
pub use value::{
    SpreadsheetFormulaCalculatedCell, SpreadsheetFormulaCalculation, SpreadsheetFormulaValue,
};

#[derive(Debug)]
struct FormulaParseFailure {
    byte_offset: usize,
    reason: String,
}

impl FormulaParseFailure {
    fn new(byte_offset: usize, reason: impl Into<String>) -> Self {
        Self {
            byte_offset,
            reason: reason.into(),
        }
    }
}

/// Parses one Spreadsheet formula into a bounded, source-spanned typed AST.
///
/// Callers may provide a formula-bar leading `=`. Spans and parse-error
/// positions address the normalized formula body after that optional marker.
pub fn parse_spreadsheet_formula(formula: &str) -> UseResult<SpreadsheetFormula> {
    let normalized = formula.strip_prefix('=').unwrap_or(formula);
    parse_normalized_formula(normalized)
}

pub(crate) fn validate_and_normalize_formula(formula: &str) -> UseResult<&str> {
    let normalized = formula.strip_prefix('=').unwrap_or(formula);
    parse_normalized_formula(normalized)?;
    Ok(normalized)
}

fn parse_normalized_formula(formula: &str) -> UseResult<SpreadsheetFormula> {
    validate_formula_bounds(formula)?;
    let tokens = lexer::lex(formula).map_err(|failure| parse_error(formula, failure))?;
    parser::parse(formula, tokens).map_err(|failure| parse_error(formula, failure))
}

fn validate_formula_bounds(formula: &str) -> UseResult<()> {
    let characters = formula.chars().count();
    if formula.is_empty() || characters > MAX_SPREADSHEET_FORMULA_CHARACTERS {
        return Err(office_error(
            "use.office.spreadsheet_formula_invalid",
            format!(
                "Spreadsheet formulas must contain 1-{MAX_SPREADSHEET_FORMULA_CHARACTERS} characters."
            ),
        )
        .with_detail("characterOffset", characters)
        .with_detail("byteOffset", formula.len())
        .with_detail("reason", "Formula length is outside supported limits."));
    }
    if let Some((byte_offset, _)) = formula
        .char_indices()
        .find(|(_, character)| character.is_control())
    {
        let character_offset = formula[..byte_offset].chars().count();
        return Err(office_error(
            "use.office.spreadsheet_formula_invalid",
            format!(
                "Spreadsheet formula is invalid at character {}: control characters are not supported.",
                character_offset + 1
            ),
        )
        .with_detail("characterOffset", character_offset)
        .with_detail("byteOffset", byte_offset)
        .with_detail(
            "reason",
            "Formula contains an unsupported control character.",
        ));
    }
    Ok(())
}

fn parse_error(formula: &str, failure: FormulaParseFailure) -> UseError {
    let byte_offset = nearest_character_boundary(formula, failure.byte_offset.min(formula.len()));
    let character_offset = formula[..byte_offset].chars().count();
    office_error(
        "use.office.spreadsheet_formula_invalid",
        format!(
            "Spreadsheet formula is invalid at character {}: {}",
            character_offset + 1,
            failure.reason
        ),
    )
    .with_detail("characterOffset", character_offset)
    .with_detail("byteOffset", byte_offset)
    .with_detail("reason", failure.reason)
}

fn nearest_character_boundary(value: &str, mut offset: usize) -> usize {
    while offset > 0 && !value.is_char_boundary(offset) {
        offset -= 1;
    }
    offset
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum ReferenceAxis {
    Row,
    Column,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum ReferenceEdit {
    Insert { at: u32, count: u32 },
    Delete { start: u32, count: u32 },
}

pub(crate) fn rewrite_formula_references(
    formula: &str,
    current_sheet: Option<&str>,
    target_sheet: &str,
    axis: ReferenceAxis,
    edit: ReferenceEdit,
) -> UseResult<String> {
    let mut output = String::with_capacity(formula.len());
    let mut cursor = 0;
    while cursor < formula.len() {
        if formula.as_bytes()[cursor] == b'"' {
            let end = quoted_string_end(formula, cursor, b'"');
            output.push_str(&formula[cursor..end]);
            cursor = end;
            continue;
        }
        if let Some(token) = parse_reference_token(formula, cursor) {
            if token
                .qualifier
                .as_ref()
                .is_some_and(|qualifier| !qualifier.external && qualifier.name.contains(':'))
            {
                return Err(formula_error(
                    "Three-dimensional sheet references are not yet safe for structural rewriting.",
                ));
            }
            let applies = token
                .qualifier
                .as_ref()
                .map(|qualifier| {
                    !qualifier.external && qualifier.name.eq_ignore_ascii_case(target_sheet)
                })
                .unwrap_or_else(|| {
                    current_sheet.is_some_and(|sheet| sheet.eq_ignore_ascii_case(target_sheet))
                });
            if applies {
                output.push_str(token.qualifier.as_ref().map_or("", |value| value.raw));
                output.push_str(&rewrite_reference(&token.reference, axis, edit)?);
            } else {
                output.push_str(&formula[cursor..token.end]);
            }
            cursor = token.end;
            continue;
        }
        let character = formula[cursor..].chars().next().unwrap_or_default();
        output.push(character);
        cursor += character.len_utf8();
    }
    Ok(output)
}

pub(crate) fn rewrite_formula_sheet_name(formula: &str, old: &str, new: &str) -> UseResult<String> {
    let mut output = String::with_capacity(formula.len());
    let mut cursor = 0;
    while cursor < formula.len() {
        if formula.as_bytes()[cursor] == b'"' {
            let end = quoted_string_end(formula, cursor, b'"');
            output.push_str(&formula[cursor..end]);
            cursor = end;
            continue;
        }
        if let Some(qualifier) = parse_qualifier(formula, cursor) {
            if !qualifier.external && qualifier.name.contains(':') {
                return Err(formula_error(
                    "Three-dimensional sheet references are not yet safe for worksheet rename.",
                ));
            }
            if !qualifier.external && qualifier.name.eq_ignore_ascii_case(old) {
                output.push_str(&quote_sheet_name(new));
                output.push('!');
            } else {
                output.push_str(qualifier.raw);
            }
            cursor = qualifier.end;
            continue;
        }
        let character = formula[cursor..].chars().next().unwrap_or_default();
        output.push(character);
        cursor += character.len_utf8();
    }
    Ok(output)
}

pub(crate) fn rewrite_formula_deleted_sheet(formula: &str, deleted: &str) -> UseResult<String> {
    let mut output = String::with_capacity(formula.len());
    let mut cursor = 0;
    while cursor < formula.len() {
        if formula.as_bytes()[cursor] == b'"' {
            let end = quoted_string_end(formula, cursor, b'"');
            output.push_str(&formula[cursor..end]);
            cursor = end;
            continue;
        }
        if let Some(qualifier) = parse_qualifier(formula, cursor) {
            if !qualifier.external && qualifier.name.contains(':') {
                return Err(formula_error(
                    "Three-dimensional sheet references are not yet safe for worksheet deletion.",
                ));
            }
            if !qualifier.external && qualifier.name.eq_ignore_ascii_case(deleted) {
                output.push_str("#REF!");
            } else {
                output.push_str(qualifier.raw);
            }
            cursor = qualifier.end;
            continue;
        }
        let character = formula[cursor..].chars().next().unwrap_or_default();
        output.push(character);
        cursor += character.len_utf8();
    }
    Ok(output)
}

#[derive(Debug, Clone, PartialEq, Eq)]
enum ParsedReference {
    Cells {
        start: A1Reference,
        end: Option<A1Reference>,
    },
    Columns {
        start: AxisReference,
        end: AxisReference,
    },
    Rows {
        start: AxisReference,
        end: AxisReference,
    },
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct A1Reference {
    column: u32,
    row: u32,
    absolute_column: bool,
    absolute_row: bool,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct AxisReference {
    value: u32,
    absolute: bool,
}

#[derive(Debug)]
struct ReferenceToken<'a> {
    qualifier: Option<SheetQualifier<'a>>,
    reference: ParsedReference,
    end: usize,
}

#[derive(Debug)]
struct SheetQualifier<'a> {
    raw: &'a str,
    name: String,
    external: bool,
    end: usize,
}

fn parse_reference_token(formula: &str, start: usize) -> Option<ReferenceToken<'_>> {
    if !is_token_boundary_before(formula, start) {
        return None;
    }
    let qualifier = parse_qualifier(formula, start);
    let reference_start = qualifier.as_ref().map_or(start, |value| value.end);
    let (reference, end) = if let Some((first, mut end)) = parse_a1(formula, reference_start) {
        let second = if formula.as_bytes().get(end) == Some(&b':') {
            let (reference, next) = parse_a1(formula, end + 1)?;
            end = next;
            Some(reference)
        } else {
            None
        };
        (
            ParsedReference::Cells {
                start: first,
                end: second,
            },
            end,
        )
    } else if let Some((first, separator)) = parse_column_axis(formula, reference_start) {
        if formula.as_bytes().get(separator) != Some(&b':') {
            return None;
        }
        let (second, end) = parse_column_axis(formula, separator + 1)?;
        (
            ParsedReference::Columns {
                start: first,
                end: second,
            },
            end,
        )
    } else {
        let (first, separator) = parse_row_axis(formula, reference_start)?;
        if formula.as_bytes().get(separator) != Some(&b':') {
            return None;
        }
        let (second, end) = parse_row_axis(formula, separator + 1)?;
        (
            ParsedReference::Rows {
                start: first,
                end: second,
            },
            end,
        )
    };
    if !is_token_boundary_after(formula, end) {
        return None;
    }
    Some(ReferenceToken {
        qualifier,
        reference,
        end,
    })
}

fn parse_qualifier(formula: &str, start: usize) -> Option<SheetQualifier<'_>> {
    let first = formula[start..].chars().next()?;
    if first == '\'' {
        let end_quote = quoted_string_end(formula, start, b'\'');
        if formula.as_bytes().get(end_quote) != Some(&b'!') {
            return None;
        }
        let raw_name = &formula[start + 1..end_quote - 1];
        let name = raw_name.replace("''", "'");
        let end = end_quote + 1;
        return Some(SheetQualifier {
            raw: &formula[start..end],
            external: name.contains(']'),
            name,
            end,
        });
    }

    let mut cursor = start;
    while cursor < formula.len() {
        let character = formula[cursor..].chars().next()?;
        if character == '!' {
            if cursor == start {
                return None;
            }
            let name = &formula[start..cursor];
            let end = cursor + 1;
            return Some(SheetQualifier {
                raw: &formula[start..end],
                name: name.to_string(),
                external: name.contains(']'),
                end,
            });
        }
        if !(character.is_alphanumeric() || matches!(character, '_' | '.' | '\\' | '[' | ']' | ':'))
        {
            return None;
        }
        cursor += character.len_utf8();
    }
    None
}

fn parse_a1(formula: &str, start: usize) -> Option<(A1Reference, usize)> {
    let bytes = formula.as_bytes();
    let mut cursor = start;
    let absolute_column = bytes.get(cursor) == Some(&b'$');
    cursor += usize::from(absolute_column);
    let column_start = cursor;
    while bytes.get(cursor).is_some_and(u8::is_ascii_alphabetic) && cursor - column_start < 3 {
        cursor += 1;
    }
    if cursor == column_start || bytes.get(cursor).is_some_and(u8::is_ascii_alphabetic) {
        return None;
    }
    let column = bytes[column_start..cursor]
        .iter()
        .try_fold(0_u32, |value, byte| {
            value.checked_mul(26).and_then(|value| {
                value.checked_add(u32::from(byte.to_ascii_uppercase() - b'A') + 1)
            })
        })
        .filter(|column| (1..=MAX_COLUMNS).contains(column))?;
    let absolute_row = bytes.get(cursor) == Some(&b'$');
    cursor += usize::from(absolute_row);
    let row_start = cursor;
    while bytes.get(cursor).is_some_and(u8::is_ascii_digit) {
        cursor += 1;
    }
    let row = formula[row_start..cursor]
        .parse::<u32>()
        .ok()
        .filter(|row| (1..=MAX_ROWS).contains(row))?;
    Some((
        A1Reference {
            column,
            row,
            absolute_column,
            absolute_row,
        },
        cursor,
    ))
}

fn parse_column_axis(formula: &str, start: usize) -> Option<(AxisReference, usize)> {
    let bytes = formula.as_bytes();
    let mut cursor = start;
    let absolute = bytes.get(cursor) == Some(&b'$');
    cursor += usize::from(absolute);
    let value_start = cursor;
    while bytes.get(cursor).is_some_and(u8::is_ascii_alphabetic) && cursor - value_start < 3 {
        cursor += 1;
    }
    if cursor == value_start || bytes.get(cursor).is_some_and(u8::is_ascii_alphabetic) {
        return None;
    }
    let value = bytes[value_start..cursor]
        .iter()
        .try_fold(0_u32, |value, byte| {
            value.checked_mul(26).and_then(|value| {
                value.checked_add(u32::from(byte.to_ascii_uppercase() - b'A') + 1)
            })
        })
        .filter(|value| (1..=MAX_COLUMNS).contains(value))?;
    Some((AxisReference { value, absolute }, cursor))
}

fn parse_row_axis(formula: &str, start: usize) -> Option<(AxisReference, usize)> {
    let bytes = formula.as_bytes();
    let mut cursor = start;
    let absolute = bytes.get(cursor) == Some(&b'$');
    cursor += usize::from(absolute);
    let value_start = cursor;
    while bytes.get(cursor).is_some_and(u8::is_ascii_digit) {
        cursor += 1;
    }
    let value = formula[value_start..cursor]
        .parse::<u32>()
        .ok()
        .filter(|value| (1..=MAX_ROWS).contains(value))?;
    Some((AxisReference { value, absolute }, cursor))
}

fn rewrite_reference(
    reference: &ParsedReference,
    axis: ReferenceAxis,
    edit: ReferenceEdit,
) -> UseResult<String> {
    let ParsedReference::Cells { start, end } = reference else {
        return match reference {
            ParsedReference::Columns { start, end } => {
                rewrite_axis_reference(*start, *end, ReferenceAxis::Column, axis, edit)
            }
            ParsedReference::Rows { start, end } => {
                rewrite_axis_reference(*start, *end, ReferenceAxis::Row, axis, edit)
            }
            ParsedReference::Cells { .. } => Err(formula_error(
                "Spreadsheet formula reference classification is inconsistent.",
            )),
        };
    };
    let Some(end) = end else {
        return transform_coordinate(*start, axis, edit)?
            .map_or_else(|| Ok("#REF!".into()), |reference| Ok(format_a1(reference)));
    };

    let first_coordinate = axis_value(*start, axis);
    let second_coordinate = axis_value(*end, axis);
    let reversed = first_coordinate > second_coordinate;
    let (low, high) = if reversed {
        (second_coordinate, first_coordinate)
    } else {
        (first_coordinate, second_coordinate)
    };
    let Some((new_low, new_high)) = transform_interval(low, high, axis, edit)? else {
        return Ok("#REF!".into());
    };
    let (first, second) = if reversed {
        (new_high, new_low)
    } else {
        (new_low, new_high)
    };
    let start = with_axis_value(*start, axis, first);
    let end = with_axis_value(*end, axis, second);
    Ok(format!("{}:{}", format_a1(start), format_a1(end)))
}

fn rewrite_axis_reference(
    start: AxisReference,
    end: AxisReference,
    reference_axis: ReferenceAxis,
    edit_axis: ReferenceAxis,
    edit: ReferenceEdit,
) -> UseResult<String> {
    if reference_axis != edit_axis {
        return Ok(format!(
            "{}:{}",
            format_axis(start, reference_axis),
            format_axis(end, reference_axis)
        ));
    }
    let reversed = start.value > end.value;
    let (low, high) = if reversed {
        (end.value, start.value)
    } else {
        (start.value, end.value)
    };
    let Some((new_low, new_high)) = transform_interval(low, high, reference_axis, edit)? else {
        return Ok("#REF!".into());
    };
    let (first, second) = if reversed {
        (new_high, new_low)
    } else {
        (new_low, new_high)
    };
    Ok(format!(
        "{}:{}",
        format_axis(
            AxisReference {
                value: first,
                absolute: start.absolute,
            },
            reference_axis
        ),
        format_axis(
            AxisReference {
                value: second,
                absolute: end.absolute,
            },
            reference_axis
        )
    ))
}

fn transform_coordinate(
    reference: A1Reference,
    axis: ReferenceAxis,
    edit: ReferenceEdit,
) -> UseResult<Option<A1Reference>> {
    let coordinate = axis_value(reference, axis);
    let transformed = match edit {
        ReferenceEdit::Insert { at, count } => {
            if coordinate >= at {
                Some(checked_coordinate(coordinate, count, axis)?)
            } else {
                Some(coordinate)
            }
        }
        ReferenceEdit::Delete { start, count } => {
            let end = start + count - 1;
            if coordinate < start {
                Some(coordinate)
            } else if coordinate <= end {
                None
            } else {
                Some(coordinate - count)
            }
        }
    };
    Ok(transformed.map(|value| with_axis_value(reference, axis, value)))
}

fn transform_interval(
    low: u32,
    high: u32,
    axis: ReferenceAxis,
    edit: ReferenceEdit,
) -> UseResult<Option<(u32, u32)>> {
    match edit {
        ReferenceEdit::Insert { at, count } => {
            if at <= low {
                Ok(Some((
                    checked_coordinate(low, count, axis)?,
                    checked_coordinate(high, count, axis)?,
                )))
            } else if at <= high {
                Ok(Some((low, checked_coordinate(high, count, axis)?)))
            } else {
                Ok(Some((low, high)))
            }
        }
        ReferenceEdit::Delete { start, count } => {
            let end = start + count - 1;
            if high < start {
                return Ok(Some((low, high)));
            }
            if low > end {
                return Ok(Some((low - count, high - count)));
            }
            let new_low = if low < start { low } else { start };
            let new_high = if high > end {
                high - count
            } else {
                start.saturating_sub(1)
            };
            Ok((new_low <= new_high && new_high > 0).then_some((new_low, new_high)))
        }
    }
}

fn checked_coordinate(value: u32, count: u32, axis: ReferenceAxis) -> UseResult<u32> {
    let limit = match axis {
        ReferenceAxis::Row => MAX_ROWS,
        ReferenceAxis::Column => MAX_COLUMNS,
    };
    value
        .checked_add(count)
        .filter(|value| *value <= limit)
        .ok_or_else(|| formula_error("Spreadsheet structural edit would move a formula reference outside worksheet limits."))
}

fn axis_value(reference: A1Reference, axis: ReferenceAxis) -> u32 {
    match axis {
        ReferenceAxis::Row => reference.row,
        ReferenceAxis::Column => reference.column,
    }
}

fn with_axis_value(mut reference: A1Reference, axis: ReferenceAxis, value: u32) -> A1Reference {
    match axis {
        ReferenceAxis::Row => reference.row = value,
        ReferenceAxis::Column => reference.column = value,
    }
    reference
}

fn format_a1(reference: A1Reference) -> String {
    format!(
        "{}{}{}{}",
        if reference.absolute_column { "$" } else { "" },
        column_name(reference.column),
        if reference.absolute_row { "$" } else { "" },
        reference.row
    )
}

fn format_axis(reference: AxisReference, axis: ReferenceAxis) -> String {
    format!(
        "{}{}",
        if reference.absolute { "$" } else { "" },
        match axis {
            ReferenceAxis::Row => reference.value.to_string(),
            ReferenceAxis::Column => column_name(reference.value),
        }
    )
}

fn is_token_boundary_before(formula: &str, start: usize) -> bool {
    formula[..start]
        .chars()
        .next_back()
        .is_none_or(|character| !(character.is_alphanumeric() || matches!(character, '_' | '.')))
}

fn is_token_boundary_after(formula: &str, end: usize) -> bool {
    formula[end..]
        .chars()
        .next()
        .is_none_or(|character| !(character.is_alphanumeric() || character == '_'))
}

fn quoted_string_end(value: &str, start: usize, quote: u8) -> usize {
    let bytes = value.as_bytes();
    let mut cursor = start + 1;
    while cursor < bytes.len() {
        if bytes[cursor] == quote {
            if bytes.get(cursor + 1) == Some(&quote) {
                cursor += 2;
                continue;
            }
            return cursor + 1;
        }
        let character = value[cursor..].chars().next().unwrap_or_default();
        cursor += character.len_utf8();
    }
    value.len()
}

fn quote_sheet_name(name: &str) -> String {
    format!("'{}'", name.replace('\'', "''"))
}

fn formula_error(message: impl Into<String>) -> UseError {
    office_error("use.office.spreadsheet_reference_rewrite_failed", message)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn public_formula_parser_is_bounded_utf8_safe_and_send_sync() {
        fn assert_send_sync<T: Send + Sync>() {}
        assert_send_sync::<SpreadsheetFormula>();

        let formula = parse_spreadsheet_formula("=SUM('销售 数据'!A1:B2,Table1[Amount])").unwrap();
        assert!(matches!(
            formula.root.kind,
            SpreadsheetFormulaExpressionKind::FunctionCall { .. }
        ));
        assert!(parse_spreadsheet_formula("==1").is_err());

        let error = parse_spreadsheet_formula("\"销售\"+").unwrap_err();
        assert_eq!(error.code, "use.office.spreadsheet_formula_invalid");
        assert_eq!(error.details["characterOffset"], 5);
        assert_eq!(error.details["byteOffset"], 9);

        let too_long = "1".repeat(MAX_SPREADSHEET_FORMULA_CHARACTERS + 1);
        assert_eq!(
            parse_spreadsheet_formula(&too_long).unwrap_err().code,
            "use.office.spreadsheet_formula_invalid"
        );
        let too_deep = format!(
            "{}1{}",
            "(".repeat(MAX_SPREADSHEET_FORMULA_DEPTH + 1),
            ")".repeat(MAX_SPREADSHEET_FORMULA_DEPTH + 1)
        );
        assert_eq!(
            parse_spreadsheet_formula(&too_deep).unwrap_err().code,
            "use.office.spreadsheet_formula_invalid"
        );
    }

    #[test]
    fn formula_parser_rejects_incomplete_and_non_excel_operator_syntax() {
        for formula in [
            "",
            " ",
            "1+",
            "+",
            "SUM(A1",
            "SUM(A1;B1)",
            "A1::B2",
            "A:1",
            "1 2",
            "A1&&B1",
            "A1!=B1",
            "\"unterminated",
            "'unterminated!A1",
            "Table1[[Column]",
            "$A+1",
        ] {
            let error = parse_spreadsheet_formula(formula).unwrap_err();
            assert_eq!(
                error.code, "use.office.spreadsheet_formula_invalid",
                "{formula}"
            );
            assert!(error.details.contains_key("characterOffset"), "{formula}");
            assert!(error.details.contains_key("byteOffset"), "{formula}");
        }
    }

    #[test]
    fn structural_rewrite_respects_sheets_strings_ranges_and_absolute_markers() {
        let formula = r#"SUM(A1,$B$2,A3:A5,'Data Set'!C4,"A1",Other!D6)"#;
        let rewritten = rewrite_formula_references(
            formula,
            Some("Data Set"),
            "Data Set",
            ReferenceAxis::Row,
            ReferenceEdit::Insert { at: 3, count: 2 },
        )
        .unwrap();
        assert_eq!(
            rewritten,
            r#"SUM(A1,$B$2,A5:A7,'Data Set'!C6,"A1",Other!D6)"#
        );

        let deleted = rewrite_formula_references(
            "A1+A2+A3+A1:A5+A2:A3",
            Some("Sheet1"),
            "Sheet1",
            ReferenceAxis::Row,
            ReferenceEdit::Delete { start: 2, count: 2 },
        )
        .unwrap();
        assert_eq!(deleted, "A1+#REF!+#REF!+A1:A3+#REF!");

        let columns = rewrite_formula_references(
            "A1+B2+C3+SUM(B:D)+SUM($2:$4)",
            Some("Sheet1"),
            "Sheet1",
            ReferenceAxis::Column,
            ReferenceEdit::Insert { at: 2, count: 1 },
        )
        .unwrap();
        assert_eq!(columns, "A1+C2+D3+SUM(C:E)+SUM($2:$4)");
    }

    #[test]
    fn sheet_rename_rewrites_only_real_qualifiers() {
        assert_eq!(
            rewrite_formula_sheet_name(r#"Old!A1+'Old'!B2+"Old!C3"+Other!D4"#, "Old", "Q1 Data")
                .unwrap(),
            r#"'Q1 Data'!A1+'Q1 Data'!B2+"Old!C3"+Other!D4"#
        );
        assert!(rewrite_formula_sheet_name("Sheet1:Sheet3!A1", "Sheet1", "Renamed").is_err());
    }

    #[test]
    fn deleted_sheets_become_ref_qualifiers_without_touching_strings_or_external_links() {
        assert_eq!(
            rewrite_formula_deleted_sheet(
                r#"Old!A1+'Old'!B2+\"Old!C3\"+'[book.xlsx]Old'!D4+Other!E5"#,
                "Old"
            )
            .unwrap(),
            r#"#REF!A1+#REF!B2+\"Old!C3\"+'[book.xlsx]Old'!D4+Other!E5"#
        );
        assert!(rewrite_formula_deleted_sheet("Sheet1:Sheet3!A1", "Sheet2").is_err());
    }
}
