//! Shared, browser-safe Spreadsheet formula grammar.
//!
//! This crate owns the bounded lexer, parser, and typed AST used by both the
//! native Office engine and the browser WebAssembly kernel. It intentionally
//! has no filesystem, network, async-runtime, OOXML, or DOM dependency.

use std::error::Error;
use std::fmt::{Display, Formatter};

mod ast;
mod lexer;
mod parser;

pub use ast::{
    SpreadsheetFormula, SpreadsheetFormulaBinaryOperator, SpreadsheetFormulaErrorLiteral,
    SpreadsheetFormulaExpression, SpreadsheetFormulaExpressionKind, SpreadsheetFormulaLiteral,
    SpreadsheetFormulaPostfixOperator, SpreadsheetFormulaQualifier, SpreadsheetFormulaReference,
    SpreadsheetFormulaReferenceKind, SpreadsheetFormulaSpan, SpreadsheetFormulaUnaryOperator,
    MAX_SPREADSHEET_FORMULA_CHARACTERS, MAX_SPREADSHEET_FORMULA_DEPTH,
    MAX_SPREADSHEET_FORMULA_NODES, MAX_SPREADSHEET_FORMULA_REFERENCE_AREAS,
};
pub use lexer::{
    FormulaToken as SpreadsheetFormulaToken, FormulaTokenKind as SpreadsheetFormulaTokenKind,
};

pub const MAX_COLUMNS: u32 = 16_384;
pub const MAX_ROWS: u32 = 1_048_576;

/// Source-positioned failure returned by the shared Spreadsheet parser.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SpreadsheetFormulaParseError {
    byte_offset: usize,
    character_offset: usize,
    reason: String,
}

impl SpreadsheetFormulaParseError {
    fn new(byte_offset: usize, reason: impl Into<String>) -> Self {
        Self {
            byte_offset,
            character_offset: 0,
            reason: reason.into(),
        }
    }

    pub fn byte_offset(&self) -> usize {
        self.byte_offset
    }

    pub fn character_offset(&self) -> usize {
        self.character_offset
    }

    pub fn reason(&self) -> &str {
        &self.reason
    }

    fn with_source(mut self, source: &str) -> Self {
        self.byte_offset = nearest_character_boundary(source, self.byte_offset.min(source.len()));
        self.character_offset = source[..self.byte_offset].chars().count();
        self
    }
}

impl Display for SpreadsheetFormulaParseError {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        write!(
            formatter,
            "Spreadsheet formula is invalid at character {}: {}",
            self.character_offset + 1,
            self.reason
        )
    }
}

impl Error for SpreadsheetFormulaParseError {}

/// Parses one formula-bar or SpreadsheetML formula into the shared typed AST.
pub fn parse_spreadsheet_formula(
    formula: &str,
) -> Result<SpreadsheetFormula, SpreadsheetFormulaParseError> {
    let normalized = formula.strip_prefix('=').unwrap_or(formula);
    validate_formula_bounds(normalized)?;
    let tokens = lexer::lex(normalized).map_err(|error| error.with_source(normalized))?;
    parser::parse(normalized, tokens).map_err(|error| error.with_source(normalized))
}

/// Tokenizes one formula-bar or SpreadsheetML formula with normalized spans.
pub fn tokenize_spreadsheet_formula(
    formula: &str,
) -> Result<Vec<SpreadsheetFormulaToken>, SpreadsheetFormulaParseError> {
    let normalized = formula.strip_prefix('=').unwrap_or(formula);
    validate_formula_bounds(normalized)?;
    lexer::lex(normalized).map_err(|error| error.with_source(normalized))
}

fn validate_formula_bounds(formula: &str) -> Result<(), SpreadsheetFormulaParseError> {
    let characters = formula.chars().count();
    if formula.is_empty() || characters > MAX_SPREADSHEET_FORMULA_CHARACTERS {
        return Err(SpreadsheetFormulaParseError::new(
            formula.len(),
            format!("Formula must contain 1-{MAX_SPREADSHEET_FORMULA_CHARACTERS} characters."),
        )
        .with_source(formula));
    }
    if let Some((byte_offset, _)) = formula
        .char_indices()
        .find(|(_, character)| character.is_control())
    {
        return Err(SpreadsheetFormulaParseError::new(
            byte_offset,
            "Formula contains an unsupported control character.",
        )
        .with_source(formula));
    }
    Ok(())
}

fn nearest_character_boundary(value: &str, mut offset: usize) -> usize {
    while offset > 0 && !value.is_char_boundary(offset) {
        offset -= 1;
    }
    offset
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn public_tokenizer_normalizes_formula_bar_input_and_unicode_errors() {
        let tokens = tokenize_spreadsheet_formula("=A1+1").unwrap();
        assert!(matches!(
            tokens.first().map(|token| &token.kind),
            Some(SpreadsheetFormulaTokenKind::Reference(_))
        ));

        let error = tokenize_spreadsheet_formula("=名+\u{0001}").unwrap_err();
        assert_eq!(error.byte_offset(), 4);
        assert_eq!(error.character_offset(), 2);
    }
}
