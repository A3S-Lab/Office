use crate::spreadsheet_reference::{MAX_COLUMNS, MAX_ROWS};

use super::{
    ast::{
        SpreadsheetFormula, SpreadsheetFormulaBinaryOperator, SpreadsheetFormulaExpression,
        SpreadsheetFormulaExpressionKind, SpreadsheetFormulaLiteral,
        SpreadsheetFormulaPostfixOperator, SpreadsheetFormulaReference,
        SpreadsheetFormulaReferenceKind, SpreadsheetFormulaSpan, SpreadsheetFormulaUnaryOperator,
        MAX_SPREADSHEET_FORMULA_DEPTH, MAX_SPREADSHEET_FORMULA_NODES,
    },
    lexer::{FormulaToken, FormulaTokenKind},
    FormulaParseFailure,
};

const MAX_FUNCTION_ARGUMENTS: usize = 255;

const BINDING_COMPARISON: u8 = 10;
const BINDING_CONCATENATE: u8 = 20;
const BINDING_ADDITIVE: u8 = 30;
const BINDING_MULTIPLICATIVE: u8 = 40;
const BINDING_POWER: u8 = 50;
const BINDING_POSTFIX: u8 = 60;
const BINDING_PREFIX: u8 = 70;
const BINDING_UNION: u8 = 80;
const BINDING_INTERSECTION: u8 = 90;
const BINDING_RANGE: u8 = 100;

pub(super) fn parse(
    source: &str,
    tokens: Vec<FormulaToken>,
) -> Result<SpreadsheetFormula, FormulaParseFailure> {
    FormulaParser::new(source, tokens).parse()
}

#[derive(Debug)]
struct ParsedExpression {
    expression: SpreadsheetFormulaExpression,
    depth: usize,
    reference_like: bool,
}

impl ParsedExpression {
    fn span(&self) -> SpreadsheetFormulaSpan {
        self.expression.span
    }
}

struct FormulaParser<'a> {
    source: &'a str,
    tokens: Vec<FormulaToken>,
    cursor: usize,
    nodes: usize,
}

impl<'a> FormulaParser<'a> {
    fn new(source: &'a str, tokens: Vec<FormulaToken>) -> Self {
        Self {
            source,
            tokens,
            cursor: 0,
            nodes: 0,
        }
    }

    fn parse(mut self) -> Result<SpreadsheetFormula, FormulaParseFailure> {
        let expression = self.parse_expression(0, true, 1)?;
        self.skip_spaces();
        let trailing = self.current()?.clone();
        if !matches!(trailing.kind, FormulaTokenKind::End) {
            return Err(FormulaParseFailure::new(
                trailing.span.start,
                format!(
                    "Unexpected {}; expected end of formula.",
                    trailing.kind.description()
                ),
            ));
        }
        validate_axis_references(&expression.expression, false)?;
        Ok(SpreadsheetFormula {
            root: expression.expression,
        })
    }

    fn parse_expression(
        &mut self,
        minimum_binding: u8,
        allow_union: bool,
        call_depth: usize,
    ) -> Result<ParsedExpression, FormulaParseFailure> {
        self.check_call_depth(call_depth)?;
        let mut left = self.parse_prefix(allow_union, call_depth)?;
        loop {
            let space = self.consume_spaces();
            let token = self.current()?.clone();

            if matches!(
                token.kind,
                FormulaTokenKind::Percent | FormulaTokenKind::Hash
            ) && BINDING_POSTFIX >= minimum_binding
            {
                self.cursor += 1;
                let operator = match token.kind {
                    FormulaTokenKind::Percent => SpreadsheetFormulaPostfixOperator::Percent,
                    FormulaTokenKind::Hash => SpreadsheetFormulaPostfixOperator::Spill,
                    _ => {
                        return Err(FormulaParseFailure::new(
                            token.span.start,
                            "Unsupported postfix operator.",
                        ));
                    }
                };
                if matches!(operator, SpreadsheetFormulaPostfixOperator::Spill)
                    && !left.reference_like
                {
                    return Err(FormulaParseFailure::new(
                        token.span.start,
                        "The spill operator requires a reference expression.",
                    ));
                }
                let span = left.span().through(token.span);
                let depth = left.depth.saturating_add(1);
                let reference_like = matches!(operator, SpreadsheetFormulaPostfixOperator::Spill);
                left = self.node(
                    SpreadsheetFormulaExpressionKind::Postfix {
                        operator,
                        operand: Box::new(left.expression),
                    },
                    span,
                    depth,
                    reference_like,
                )?;
                continue;
            }

            let intersection = space
                .filter(|_| left.reference_like && token_starts_reference_expression(&token.kind));
            let Some((operator, left_binding, right_binding, consume_token)) =
                infix_operator(&token.kind, allow_union, intersection.is_some())
            else {
                break;
            };
            if left_binding < minimum_binding {
                break;
            }
            let operator_span = if let Some(space) = intersection {
                space
            } else {
                if consume_token {
                    self.cursor += 1;
                }
                token.span
            };
            let right =
                self.parse_expression(right_binding, allow_union, call_depth.saturating_add(1))?;
            left = self.binary(left, right, operator, operator_span)?;
        }
        Ok(left)
    }

    fn parse_prefix(
        &mut self,
        allow_union: bool,
        call_depth: usize,
    ) -> Result<ParsedExpression, FormulaParseFailure> {
        self.skip_spaces();
        let token = self.current()?.clone();
        self.cursor += 1;
        match token.kind {
            FormulaTokenKind::Number(value) => self.node(
                SpreadsheetFormulaExpressionKind::Literal(SpreadsheetFormulaLiteral::Number(value)),
                token.span,
                1,
                false,
            ),
            FormulaTokenKind::Text(value) => self.node(
                SpreadsheetFormulaExpressionKind::Literal(SpreadsheetFormulaLiteral::Text(value)),
                token.span,
                1,
                false,
            ),
            FormulaTokenKind::Error(value) => self.node(
                SpreadsheetFormulaExpressionKind::Literal(SpreadsheetFormulaLiteral::Error(value)),
                token.span,
                1,
                false,
            ),
            FormulaTokenKind::Reference(reference) => self.node(
                SpreadsheetFormulaExpressionKind::Reference(reference),
                token.span,
                1,
                true,
            ),
            FormulaTokenKind::StructuredReference {
                qualifier,
                reference,
            } => self.node(
                SpreadsheetFormulaExpressionKind::StructuredReference {
                    qualifier,
                    reference,
                },
                token.span,
                1,
                true,
            ),
            FormulaTokenKind::Name { qualifier, name } => {
                if matches!(
                    self.current().map(|value| &value.kind),
                    Ok(FormulaTokenKind::LeftParen)
                ) {
                    self.parse_function(qualifier, name, token.span, call_depth)
                } else if qualifier.is_none() && name.eq_ignore_ascii_case("TRUE") {
                    self.node(
                        SpreadsheetFormulaExpressionKind::Literal(
                            SpreadsheetFormulaLiteral::Boolean(true),
                        ),
                        token.span,
                        1,
                        false,
                    )
                } else if qualifier.is_none() && name.eq_ignore_ascii_case("FALSE") {
                    self.node(
                        SpreadsheetFormulaExpressionKind::Literal(
                            SpreadsheetFormulaLiteral::Boolean(false),
                        ),
                        token.span,
                        1,
                        false,
                    )
                } else {
                    self.node(
                        SpreadsheetFormulaExpressionKind::Name { qualifier, name },
                        token.span,
                        1,
                        true,
                    )
                }
            }
            FormulaTokenKind::Plus | FormulaTokenKind::Minus | FormulaTokenKind::At => {
                let operator = match token.kind {
                    FormulaTokenKind::Plus => SpreadsheetFormulaUnaryOperator::Positive,
                    FormulaTokenKind::Minus => SpreadsheetFormulaUnaryOperator::Negative,
                    FormulaTokenKind::At => SpreadsheetFormulaUnaryOperator::ImplicitIntersection,
                    _ => {
                        return Err(FormulaParseFailure::new(
                            token.span.start,
                            "Unsupported prefix operator.",
                        ));
                    }
                };
                let operand = self.parse_expression(
                    BINDING_PREFIX,
                    allow_union,
                    call_depth.saturating_add(1),
                )?;
                if matches!(
                    operator,
                    SpreadsheetFormulaUnaryOperator::ImplicitIntersection
                ) && !operand.reference_like
                {
                    return Err(FormulaParseFailure::new(
                        token.span.start,
                        "Implicit intersection requires a reference expression.",
                    ));
                }
                let span = token.span.through(operand.span());
                let depth = operand.depth.saturating_add(1);
                let reference_like = matches!(
                    operator,
                    SpreadsheetFormulaUnaryOperator::ImplicitIntersection
                );
                self.node(
                    SpreadsheetFormulaExpressionKind::Unary {
                        operator,
                        operand: Box::new(operand.expression),
                    },
                    span,
                    depth,
                    reference_like,
                )
            }
            FormulaTokenKind::LeftParen => {
                let inner = self.parse_expression(0, true, call_depth.saturating_add(1))?;
                self.skip_spaces();
                let close = self.expect(FormulaTokenKind::RightParen, "`)`")?;
                let span = token.span.through(close.span);
                let depth = inner.depth.saturating_add(1);
                let reference_like = inner.reference_like;
                self.node(
                    SpreadsheetFormulaExpressionKind::Parenthesized(Box::new(inner.expression)),
                    span,
                    depth,
                    reference_like,
                )
            }
            FormulaTokenKind::LeftBrace => self.parse_array(token.span, call_depth),
            other => Err(FormulaParseFailure::new(
                token.span.start,
                format!("Expected an expression but found {}.", other.description()),
            )),
        }
    }

    fn parse_function(
        &mut self,
        qualifier: Option<super::ast::SpreadsheetFormulaQualifier>,
        name: String,
        name_span: SpreadsheetFormulaSpan,
        call_depth: usize,
    ) -> Result<ParsedExpression, FormulaParseFailure> {
        self.cursor += 1;
        let mut arguments = Vec::new();
        let mut last_was_separator = false;
        loop {
            self.skip_spaces();
            let token = self.current()?.clone();
            if matches!(token.kind, FormulaTokenKind::RightParen) {
                self.cursor += 1;
                if last_was_separator {
                    arguments.push(None);
                }
                if arguments.len() > MAX_FUNCTION_ARGUMENTS {
                    return Err(FormulaParseFailure::new(
                        token.span.start,
                        format!(
                            "Function calls accept at most {MAX_FUNCTION_ARGUMENTS} arguments."
                        ),
                    ));
                }
                let child_depth = arguments
                    .iter()
                    .filter_map(Option::as_ref)
                    .map(expression_depth)
                    .max()
                    .unwrap_or(0);
                return self.node(
                    SpreadsheetFormulaExpressionKind::FunctionCall {
                        qualifier,
                        name,
                        arguments,
                    },
                    name_span.through(token.span),
                    child_depth.saturating_add(1),
                    true,
                );
            }
            if matches!(token.kind, FormulaTokenKind::End) {
                return Err(FormulaParseFailure::new(
                    token.span.start,
                    "Function call is not closed with `)`.",
                ));
            }
            if matches!(token.kind, FormulaTokenKind::Comma) {
                self.cursor += 1;
                arguments.push(None);
                last_was_separator = true;
            } else {
                let argument = self.parse_expression(0, false, call_depth.saturating_add(1))?;
                arguments.push(Some(argument.expression));
                last_was_separator = false;
                self.skip_spaces();
                let separator = self.current()?.clone();
                match separator.kind {
                    FormulaTokenKind::Comma => {
                        self.cursor += 1;
                        last_was_separator = true;
                    }
                    FormulaTokenKind::RightParen => {}
                    FormulaTokenKind::Semicolon => {
                        return Err(FormulaParseFailure::new(
                            separator.span.start,
                            "SpreadsheetML function arguments use `,`, not `;`.",
                        ));
                    }
                    _ => {
                        return Err(FormulaParseFailure::new(
                            separator.span.start,
                            format!(
                                "Expected `,` or `)` after a function argument, found {}.",
                                separator.kind.description()
                            ),
                        ));
                    }
                }
            }
            if arguments.len() >= MAX_FUNCTION_ARGUMENTS && last_was_separator {
                return Err(FormulaParseFailure::new(
                    self.current()?.span.start,
                    format!("Function calls accept at most {MAX_FUNCTION_ARGUMENTS} arguments."),
                ));
            }
        }
    }

    fn parse_array(
        &mut self,
        open_span: SpreadsheetFormulaSpan,
        call_depth: usize,
    ) -> Result<ParsedExpression, FormulaParseFailure> {
        let mut rows = Vec::<Vec<SpreadsheetFormulaExpression>>::new();
        let mut row = Vec::new();
        let mut last_was_column_separator = false;
        loop {
            self.skip_spaces();
            let token = self.current()?.clone();
            if matches!(token.kind, FormulaTokenKind::RightBrace) {
                if row.is_empty() || last_was_column_separator {
                    return Err(FormulaParseFailure::new(
                        token.span.start,
                        "Array constants cannot contain an empty row.",
                    ));
                }
                rows.push(row);
                self.cursor += 1;
                let width = rows.first().map_or(0, Vec::len);
                if rows.iter().any(|value| value.len() != width) {
                    return Err(FormulaParseFailure::new(
                        token.span.start,
                        "Array constant rows must have equal widths.",
                    ));
                }
                let child_depth = rows
                    .iter()
                    .flatten()
                    .map(expression_depth)
                    .max()
                    .unwrap_or(0);
                return self.node(
                    SpreadsheetFormulaExpressionKind::Array { rows },
                    open_span.through(token.span),
                    child_depth.saturating_add(1),
                    false,
                );
            }
            if matches!(token.kind, FormulaTokenKind::End) {
                return Err(FormulaParseFailure::new(
                    token.span.start,
                    "Array constant is not closed with `}`.",
                ));
            }
            let value = self.parse_expression(0, false, call_depth.saturating_add(1))?;
            if !is_array_constant(&value.expression) {
                return Err(FormulaParseFailure::new(
                    value.span().start,
                    "Array constants may contain only scalar literals.",
                ));
            }
            row.push(value.expression);
            last_was_column_separator = false;
            self.skip_spaces();
            let separator = self.current()?.clone();
            match separator.kind {
                FormulaTokenKind::Comma => {
                    self.cursor += 1;
                    last_was_column_separator = true;
                }
                FormulaTokenKind::Semicolon => {
                    self.cursor += 1;
                    if row.is_empty() {
                        return Err(FormulaParseFailure::new(
                            separator.span.start,
                            "Array constants cannot contain an empty row.",
                        ));
                    }
                    rows.push(std::mem::take(&mut row));
                    last_was_column_separator = false;
                }
                FormulaTokenKind::RightBrace => {}
                _ => {
                    return Err(FormulaParseFailure::new(
                        separator.span.start,
                        format!(
                            "Expected `,`, `;`, or `}}` in an array constant, found {}.",
                            separator.kind.description()
                        ),
                    ));
                }
            }
        }
    }

    fn binary(
        &mut self,
        mut left: ParsedExpression,
        mut right: ParsedExpression,
        operator: SpreadsheetFormulaBinaryOperator,
        operator_span: SpreadsheetFormulaSpan,
    ) -> Result<ParsedExpression, FormulaParseFailure> {
        if matches!(operator, SpreadsheetFormulaBinaryOperator::Range) {
            left = coerce_range_endpoint(left, operator_span.start)?;
            right = coerce_range_endpoint(right, operator_span.start)?;
            if let (Some(left_axis), Some(right_axis)) = (
                reference_axis(&left.expression),
                reference_axis(&right.expression),
            ) {
                if left_axis != right_axis {
                    return Err(FormulaParseFailure::new(
                        operator_span.start,
                        "Range endpoints must both be cells, columns, or rows.",
                    ));
                }
            }
        } else if matches!(
            operator,
            SpreadsheetFormulaBinaryOperator::Intersection
                | SpreadsheetFormulaBinaryOperator::Union
        ) && !(left.reference_like && right.reference_like)
        {
            return Err(FormulaParseFailure::new(
                operator_span.start,
                "Reference operators require reference expressions on both sides.",
            ));
        }
        let span = left.span().through(right.span());
        let depth = left.depth.max(right.depth).saturating_add(1);
        let reference_like = matches!(
            operator,
            SpreadsheetFormulaBinaryOperator::Range
                | SpreadsheetFormulaBinaryOperator::Intersection
                | SpreadsheetFormulaBinaryOperator::Union
        );
        self.node(
            SpreadsheetFormulaExpressionKind::Binary {
                operator,
                left: Box::new(left.expression),
                right: Box::new(right.expression),
            },
            span,
            depth,
            reference_like,
        )
    }

    fn node(
        &mut self,
        kind: SpreadsheetFormulaExpressionKind,
        span: SpreadsheetFormulaSpan,
        depth: usize,
        reference_like: bool,
    ) -> Result<ParsedExpression, FormulaParseFailure> {
        if depth > MAX_SPREADSHEET_FORMULA_DEPTH {
            return Err(FormulaParseFailure::new(
                span.start,
                format!("Formula AST depth exceeds {MAX_SPREADSHEET_FORMULA_DEPTH}."),
            ));
        }
        self.nodes = self.nodes.saturating_add(1);
        if self.nodes > MAX_SPREADSHEET_FORMULA_NODES {
            return Err(FormulaParseFailure::new(
                span.start,
                format!("Formula AST contains more than {MAX_SPREADSHEET_FORMULA_NODES} nodes."),
            ));
        }
        Ok(ParsedExpression {
            expression: SpreadsheetFormulaExpression { span, kind },
            depth,
            reference_like,
        })
    }

    fn expect(
        &mut self,
        expected: FormulaTokenKind,
        description: &str,
    ) -> Result<FormulaToken, FormulaParseFailure> {
        let token = self.current()?.clone();
        if std::mem::discriminant(&token.kind) != std::mem::discriminant(&expected) {
            return Err(FormulaParseFailure::new(
                token.span.start,
                format!(
                    "Expected {description}, found {}.",
                    token.kind.description()
                ),
            ));
        }
        self.cursor += 1;
        Ok(token)
    }

    fn current(&self) -> Result<&FormulaToken, FormulaParseFailure> {
        self.tokens.get(self.cursor).ok_or_else(|| {
            FormulaParseFailure::new(
                self.source.len(),
                "Formula token stream ended unexpectedly.",
            )
        })
    }

    fn skip_spaces(&mut self) {
        while self
            .tokens
            .get(self.cursor)
            .is_some_and(|token| matches!(token.kind, FormulaTokenKind::Space))
        {
            self.cursor += 1;
        }
    }

    fn consume_spaces(&mut self) -> Option<SpreadsheetFormulaSpan> {
        let start = self
            .tokens
            .get(self.cursor)
            .filter(|token| matches!(token.kind, FormulaTokenKind::Space))?
            .span;
        let mut end = start;
        while let Some(token) = self
            .tokens
            .get(self.cursor)
            .filter(|token| matches!(token.kind, FormulaTokenKind::Space))
        {
            end = token.span;
            self.cursor += 1;
        }
        Some(start.through(end))
    }

    fn check_call_depth(&self, depth: usize) -> Result<(), FormulaParseFailure> {
        if depth > MAX_SPREADSHEET_FORMULA_DEPTH {
            return Err(FormulaParseFailure::new(
                self.current()
                    .map_or(self.source.len(), |token| token.span.start),
                format!("Formula parse nesting exceeds {MAX_SPREADSHEET_FORMULA_DEPTH}."),
            ));
        }
        Ok(())
    }
}

fn infix_operator(
    token: &FormulaTokenKind,
    allow_union: bool,
    intersection: bool,
) -> Option<(SpreadsheetFormulaBinaryOperator, u8, u8, bool)> {
    if intersection {
        return Some((
            SpreadsheetFormulaBinaryOperator::Intersection,
            BINDING_INTERSECTION,
            BINDING_INTERSECTION + 1,
            false,
        ));
    }
    let (operator, binding, right_associative) = match token {
        FormulaTokenKind::Colon => (
            SpreadsheetFormulaBinaryOperator::Range,
            BINDING_RANGE,
            false,
        ),
        FormulaTokenKind::Comma if allow_union => (
            SpreadsheetFormulaBinaryOperator::Union,
            BINDING_UNION,
            false,
        ),
        FormulaTokenKind::Caret => (SpreadsheetFormulaBinaryOperator::Power, BINDING_POWER, true),
        FormulaTokenKind::Star => (
            SpreadsheetFormulaBinaryOperator::Multiply,
            BINDING_MULTIPLICATIVE,
            false,
        ),
        FormulaTokenKind::Slash => (
            SpreadsheetFormulaBinaryOperator::Divide,
            BINDING_MULTIPLICATIVE,
            false,
        ),
        FormulaTokenKind::Plus => (
            SpreadsheetFormulaBinaryOperator::Add,
            BINDING_ADDITIVE,
            false,
        ),
        FormulaTokenKind::Minus => (
            SpreadsheetFormulaBinaryOperator::Subtract,
            BINDING_ADDITIVE,
            false,
        ),
        FormulaTokenKind::Ampersand => (
            SpreadsheetFormulaBinaryOperator::Concatenate,
            BINDING_CONCATENATE,
            false,
        ),
        FormulaTokenKind::Equal => (
            SpreadsheetFormulaBinaryOperator::Equal,
            BINDING_COMPARISON,
            false,
        ),
        FormulaTokenKind::NotEqual => (
            SpreadsheetFormulaBinaryOperator::NotEqual,
            BINDING_COMPARISON,
            false,
        ),
        FormulaTokenKind::LessThan => (
            SpreadsheetFormulaBinaryOperator::LessThan,
            BINDING_COMPARISON,
            false,
        ),
        FormulaTokenKind::LessThanOrEqual => (
            SpreadsheetFormulaBinaryOperator::LessThanOrEqual,
            BINDING_COMPARISON,
            false,
        ),
        FormulaTokenKind::GreaterThan => (
            SpreadsheetFormulaBinaryOperator::GreaterThan,
            BINDING_COMPARISON,
            false,
        ),
        FormulaTokenKind::GreaterThanOrEqual => (
            SpreadsheetFormulaBinaryOperator::GreaterThanOrEqual,
            BINDING_COMPARISON,
            false,
        ),
        _ => return None,
    };
    Some((
        operator,
        binding,
        if right_associative {
            binding
        } else {
            binding + 1
        },
        true,
    ))
}

fn token_starts_reference_expression(token: &FormulaTokenKind) -> bool {
    matches!(
        token,
        FormulaTokenKind::Reference(_)
            | FormulaTokenKind::Name { .. }
            | FormulaTokenKind::StructuredReference { .. }
            | FormulaTokenKind::LeftParen
            | FormulaTokenKind::At
    )
}

fn coerce_range_endpoint(
    mut value: ParsedExpression,
    position: usize,
) -> Result<ParsedExpression, FormulaParseFailure> {
    let replacement = match &value.expression.kind {
        SpreadsheetFormulaExpressionKind::Name { qualifier, name } => {
            if name.bytes().all(|byte| byte.is_ascii_alphabetic()) {
                parse_column(name).map(|column| SpreadsheetFormulaReference {
                    qualifier: qualifier.clone(),
                    kind: SpreadsheetFormulaReferenceKind::Column {
                        column,
                        absolute: false,
                    },
                })
            } else if name.bytes().all(|byte| byte.is_ascii_digit()) {
                parse_row(name).map(|row| SpreadsheetFormulaReference {
                    qualifier: qualifier.clone(),
                    kind: SpreadsheetFormulaReferenceKind::Row {
                        row,
                        absolute: false,
                    },
                })
            } else {
                None
            }
        }
        SpreadsheetFormulaExpressionKind::Literal(SpreadsheetFormulaLiteral::Number(value)) => {
            parse_row(value).map(|row| SpreadsheetFormulaReference {
                qualifier: None,
                kind: SpreadsheetFormulaReferenceKind::Row {
                    row,
                    absolute: false,
                },
            })
        }
        _ => None,
    };
    if let Some(reference) = replacement {
        value.expression.kind = SpreadsheetFormulaExpressionKind::Reference(reference);
        value.reference_like = true;
    }
    if !value.reference_like {
        return Err(FormulaParseFailure::new(
            position,
            "Range operator requires reference endpoints.",
        ));
    }
    Ok(value)
}

fn parse_column(value: &str) -> Option<u32> {
    if value.is_empty() || value.len() > 3 {
        return None;
    }
    value
        .bytes()
        .try_fold(0_u32, |column, byte| {
            column.checked_mul(26).and_then(|column| {
                column.checked_add(u32::from(byte.to_ascii_uppercase().checked_sub(b'A')?) + 1)
            })
        })
        .filter(|column| (1..=MAX_COLUMNS).contains(column))
}

fn parse_row(value: &str) -> Option<u32> {
    if value.is_empty() || !value.bytes().all(|byte| byte.is_ascii_digit()) {
        return None;
    }
    value
        .parse::<u32>()
        .ok()
        .filter(|row| (1..=MAX_ROWS).contains(row))
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum ReferenceAxisKind {
    Cell,
    Column,
    Row,
}

fn reference_axis(expression: &SpreadsheetFormulaExpression) -> Option<ReferenceAxisKind> {
    let SpreadsheetFormulaExpressionKind::Reference(reference) = &expression.kind else {
        return None;
    };
    Some(match reference.kind {
        SpreadsheetFormulaReferenceKind::Cell { .. } => ReferenceAxisKind::Cell,
        SpreadsheetFormulaReferenceKind::Column { .. } => ReferenceAxisKind::Column,
        SpreadsheetFormulaReferenceKind::Row { .. } => ReferenceAxisKind::Row,
    })
}

fn validate_axis_references(
    expression: &SpreadsheetFormulaExpression,
    range_endpoint: bool,
) -> Result<(), FormulaParseFailure> {
    match &expression.kind {
        SpreadsheetFormulaExpressionKind::Reference(SpreadsheetFormulaReference {
            kind:
                SpreadsheetFormulaReferenceKind::Column { .. }
                | SpreadsheetFormulaReferenceKind::Row { .. },
            ..
        }) if !range_endpoint => Err(FormulaParseFailure::new(
            expression.span.start,
            "Whole-row and whole-column references must be used in a range.",
        )),
        SpreadsheetFormulaExpressionKind::Name {
            qualifier: Some(_),
            name,
        } if name.bytes().all(|byte| byte.is_ascii_digit()) => Err(FormulaParseFailure::new(
            expression.span.start,
            "Qualified row references must include a range operator.",
        )),
        SpreadsheetFormulaExpressionKind::Unary { operand, .. }
        | SpreadsheetFormulaExpressionKind::Postfix { operand, .. }
        | SpreadsheetFormulaExpressionKind::Parenthesized(operand) => {
            validate_axis_references(operand, false)
        }
        SpreadsheetFormulaExpressionKind::Binary {
            operator,
            left,
            right,
        } => {
            let endpoints = matches!(operator, SpreadsheetFormulaBinaryOperator::Range);
            validate_axis_references(left, endpoints)?;
            validate_axis_references(right, endpoints)
        }
        SpreadsheetFormulaExpressionKind::FunctionCall { arguments, .. } => {
            for argument in arguments.iter().flatten() {
                validate_axis_references(argument, false)?;
            }
            Ok(())
        }
        SpreadsheetFormulaExpressionKind::Array { rows } => {
            for value in rows.iter().flatten() {
                validate_axis_references(value, false)?;
            }
            Ok(())
        }
        _ => Ok(()),
    }
}

fn is_array_constant(expression: &SpreadsheetFormulaExpression) -> bool {
    match &expression.kind {
        SpreadsheetFormulaExpressionKind::Literal(_) => true,
        SpreadsheetFormulaExpressionKind::Unary {
            operator:
                SpreadsheetFormulaUnaryOperator::Positive | SpreadsheetFormulaUnaryOperator::Negative,
            operand,
        } => matches!(
            &operand.kind,
            SpreadsheetFormulaExpressionKind::Literal(SpreadsheetFormulaLiteral::Number(_))
        ),
        _ => false,
    }
}

fn expression_depth(expression: &SpreadsheetFormulaExpression) -> usize {
    match &expression.kind {
        SpreadsheetFormulaExpressionKind::Literal(_)
        | SpreadsheetFormulaExpressionKind::Reference(_)
        | SpreadsheetFormulaExpressionKind::Name { .. }
        | SpreadsheetFormulaExpressionKind::StructuredReference { .. } => 1,
        SpreadsheetFormulaExpressionKind::Unary { operand, .. }
        | SpreadsheetFormulaExpressionKind::Postfix { operand, .. }
        | SpreadsheetFormulaExpressionKind::Parenthesized(operand) => {
            expression_depth(operand).saturating_add(1)
        }
        SpreadsheetFormulaExpressionKind::Binary { left, right, .. } => expression_depth(left)
            .max(expression_depth(right))
            .saturating_add(1),
        SpreadsheetFormulaExpressionKind::FunctionCall { arguments, .. } => arguments
            .iter()
            .filter_map(Option::as_ref)
            .map(expression_depth)
            .max()
            .unwrap_or(0)
            .saturating_add(1),
        SpreadsheetFormulaExpressionKind::Array { rows } => rows
            .iter()
            .flatten()
            .map(expression_depth)
            .max()
            .unwrap_or(0)
            .saturating_add(1),
    }
}

#[cfg(test)]
mod tests;
