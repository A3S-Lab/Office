use crate::spreadsheet_reference::{MAX_COLUMNS, MAX_ROWS};

use super::{
    ast::{
        SpreadsheetFormulaErrorLiteral, SpreadsheetFormulaQualifier, SpreadsheetFormulaReference,
        SpreadsheetFormulaReferenceKind, SpreadsheetFormulaSpan, MAX_SPREADSHEET_FORMULA_NODES,
    },
    FormulaParseFailure,
};

const MAX_LEXICAL_TOKENS: usize = MAX_SPREADSHEET_FORMULA_NODES * 2 + 1;

#[derive(Debug, Clone, PartialEq, Eq)]
pub(super) struct FormulaToken {
    pub(super) kind: FormulaTokenKind,
    pub(super) span: SpreadsheetFormulaSpan,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(super) enum FormulaTokenKind {
    Number(String),
    Text(String),
    Error(SpreadsheetFormulaErrorLiteral),
    Reference(SpreadsheetFormulaReference),
    Name {
        qualifier: Option<SpreadsheetFormulaQualifier>,
        name: String,
    },
    StructuredReference {
        qualifier: Option<SpreadsheetFormulaQualifier>,
        reference: String,
    },
    Space,
    Plus,
    Minus,
    Star,
    Slash,
    Caret,
    Ampersand,
    Equal,
    NotEqual,
    LessThan,
    LessThanOrEqual,
    GreaterThan,
    GreaterThanOrEqual,
    Percent,
    Hash,
    At,
    Colon,
    Comma,
    Semicolon,
    LeftParen,
    RightParen,
    LeftBrace,
    RightBrace,
    End,
}

impl FormulaTokenKind {
    pub(super) const fn description(&self) -> &'static str {
        match self {
            Self::Number(_) => "number",
            Self::Text(_) => "string",
            Self::Error(_) => "error literal",
            Self::Reference(_) => "cell reference",
            Self::Name { .. } => "name",
            Self::StructuredReference { .. } => "structured reference",
            Self::Space => "space",
            Self::Plus => "`+`",
            Self::Minus => "`-`",
            Self::Star => "`*`",
            Self::Slash => "`/`",
            Self::Caret => "`^`",
            Self::Ampersand => "`&`",
            Self::Equal => "`=`",
            Self::NotEqual => "`<>`",
            Self::LessThan => "`<`",
            Self::LessThanOrEqual => "`<=`",
            Self::GreaterThan => "`>`",
            Self::GreaterThanOrEqual => "`>=`",
            Self::Percent => "`%`",
            Self::Hash => "`#`",
            Self::At => "`@`",
            Self::Colon => "`:`",
            Self::Comma => "`,`",
            Self::Semicolon => "`;`",
            Self::LeftParen => "`(`",
            Self::RightParen => "`)`",
            Self::LeftBrace => "`{`",
            Self::RightBrace => "`}`",
            Self::End => "end of formula",
        }
    }
}

pub(super) fn lex(source: &str) -> Result<Vec<FormulaToken>, FormulaParseFailure> {
    FormulaLexer::new(source).lex()
}

struct FormulaLexer<'a> {
    source: &'a str,
    cursor: usize,
    tokens: Vec<FormulaToken>,
}

impl<'a> FormulaLexer<'a> {
    fn new(source: &'a str) -> Self {
        Self {
            source,
            cursor: 0,
            tokens: Vec::with_capacity(source.len().min(MAX_LEXICAL_TOKENS)),
        }
    }

    fn lex(mut self) -> Result<Vec<FormulaToken>, FormulaParseFailure> {
        while self.cursor < self.source.len() {
            let start = self.cursor;
            let character = self.current_character().ok_or_else(|| {
                FormulaParseFailure::new(self.cursor, "Formula contains invalid UTF-8 boundaries.")
            })?;
            if character.is_whitespace() {
                self.consume_while(char::is_whitespace);
                self.push(FormulaTokenKind::Space, start, self.cursor)?;
                continue;
            }
            if character == '"' {
                let token = self.lex_text()?;
                self.push_token(token)?;
                continue;
            }
            if character == '\'' {
                let token = self.lex_qualified_atom()?.ok_or_else(|| {
                    FormulaParseFailure::new(start, "Worksheet qualifier is invalid.")
                })?;
                self.push_token(token)?;
                continue;
            }
            if let Some(token) = self.lex_qualified_atom()? {
                self.push_token(token)?;
                continue;
            }
            if character == '$' || character.is_ascii_alphabetic() {
                match self.try_cell_reference(self.cursor, None, false)? {
                    Some((reference, end)) => {
                        self.cursor = end;
                        self.push(FormulaTokenKind::Reference(reference), start, self.cursor)?;
                        continue;
                    }
                    None if character == '$' => {
                        let (reference, end) = self.lex_absolute_axis(None)?;
                        self.cursor = end;
                        self.push(FormulaTokenKind::Reference(reference), start, self.cursor)?;
                        continue;
                    }
                    None => {}
                }
            }
            if character.is_ascii_digit()
                || (character == '.'
                    && self
                        .next_character()
                        .is_some_and(|value| value.is_ascii_digit()))
            {
                let token = self.lex_number()?;
                self.push_token(token)?;
                continue;
            }
            if is_name_start(character) || character == '[' {
                let token = self.lex_name(None, start)?;
                self.push_token(token)?;
                continue;
            }
            if character == '#' {
                if let Some(token) = self.lex_error_literal()? {
                    self.push_token(token)?;
                } else {
                    self.cursor += character.len_utf8();
                    self.push(FormulaTokenKind::Hash, start, self.cursor)?;
                }
                continue;
            }

            self.cursor += character.len_utf8();
            let kind = match character {
                '+' => FormulaTokenKind::Plus,
                '-' => FormulaTokenKind::Minus,
                '*' => FormulaTokenKind::Star,
                '/' => FormulaTokenKind::Slash,
                '^' => FormulaTokenKind::Caret,
                '&' => FormulaTokenKind::Ampersand,
                '=' => FormulaTokenKind::Equal,
                '%' => FormulaTokenKind::Percent,
                '@' => FormulaTokenKind::At,
                ':' => FormulaTokenKind::Colon,
                ',' => FormulaTokenKind::Comma,
                ';' => FormulaTokenKind::Semicolon,
                '(' => FormulaTokenKind::LeftParen,
                ')' => FormulaTokenKind::RightParen,
                '{' => FormulaTokenKind::LeftBrace,
                '}' => FormulaTokenKind::RightBrace,
                '<' => {
                    if self.source[self.cursor..].starts_with('=') {
                        self.cursor += 1;
                        FormulaTokenKind::LessThanOrEqual
                    } else if self.source[self.cursor..].starts_with('>') {
                        self.cursor += 1;
                        FormulaTokenKind::NotEqual
                    } else {
                        FormulaTokenKind::LessThan
                    }
                }
                '>' => {
                    if self.source[self.cursor..].starts_with('=') {
                        self.cursor += 1;
                        FormulaTokenKind::GreaterThanOrEqual
                    } else {
                        FormulaTokenKind::GreaterThan
                    }
                }
                _ => {
                    return Err(FormulaParseFailure::new(
                        start,
                        format!("Unexpected character `{character}`."),
                    ));
                }
            };
            self.push(kind, start, self.cursor)?;
        }
        self.push(FormulaTokenKind::End, self.source.len(), self.source.len())?;
        Ok(self.tokens)
    }

    fn lex_text(&mut self) -> Result<FormulaToken, FormulaParseFailure> {
        let start = self.cursor;
        self.cursor += 1;
        let mut value = String::new();
        while self.cursor < self.source.len() {
            let character = self.current_character().ok_or_else(|| {
                FormulaParseFailure::new(self.cursor, "String literal has an invalid boundary.")
            })?;
            if character == '"' {
                self.cursor += 1;
                if self.source[self.cursor..].starts_with('"') {
                    value.push('"');
                    self.cursor += 1;
                    continue;
                }
                return Ok(FormulaToken {
                    kind: FormulaTokenKind::Text(value),
                    span: SpreadsheetFormulaSpan::new(start, self.cursor),
                });
            }
            value.push(character);
            self.cursor += character.len_utf8();
        }
        Err(FormulaParseFailure::new(
            self.source.len(),
            "String literal is not closed.",
        ))
    }

    fn lex_qualified_atom(&mut self) -> Result<Option<FormulaToken>, FormulaParseFailure> {
        let start = self.cursor;
        let Some((qualifier, target_start)) = self.scan_qualifier()? else {
            return Ok(None);
        };
        self.cursor = target_start;
        if self.cursor >= self.source.len() {
            return Err(FormulaParseFailure::new(
                self.cursor,
                "Worksheet qualifier must be followed by a reference or name.",
            ));
        }
        if let Some((reference, end)) =
            self.try_cell_reference(self.cursor, Some(qualifier.clone()), true)?
        {
            self.cursor = end;
            return Ok(Some(FormulaToken {
                kind: FormulaTokenKind::Reference(reference),
                span: SpreadsheetFormulaSpan::new(start, end),
            }));
        }
        if self.source[self.cursor..].starts_with('$') {
            let (reference, end) = self.lex_absolute_axis(Some(qualifier))?;
            self.cursor = end;
            return Ok(Some(FormulaToken {
                kind: FormulaTokenKind::Reference(reference),
                span: SpreadsheetFormulaSpan::new(start, end),
            }));
        }
        let token = self.lex_name(Some(qualifier), start)?;
        Ok(Some(token))
    }

    fn scan_qualifier(
        &self,
    ) -> Result<Option<(SpreadsheetFormulaQualifier, usize)>, FormulaParseFailure> {
        let start = self.cursor;
        let character = self.source[start..].chars().next().ok_or_else(|| {
            FormulaParseFailure::new(start, "Worksheet qualifier has an invalid boundary.")
        })?;
        if character == '\'' {
            let mut cursor = start + 1;
            let mut decoded = String::new();
            while cursor < self.source.len() {
                let current = self.source[cursor..].chars().next().ok_or_else(|| {
                    FormulaParseFailure::new(cursor, "Worksheet qualifier is not valid UTF-8.")
                })?;
                if current == '\'' {
                    let after = cursor + 1;
                    if self.source[after..].starts_with('\'') {
                        decoded.push('\'');
                        cursor = after + 1;
                        continue;
                    }
                    if self.source[after..].starts_with('!') {
                        let qualifier = parse_qualifier_value(decoded, start)?;
                        return Ok(Some((qualifier, after + 1)));
                    }
                    return Err(FormulaParseFailure::new(
                        after,
                        "A quoted worksheet qualifier must be followed by `!`.",
                    ));
                }
                decoded.push(current);
                cursor += current.len_utf8();
            }
            return Err(FormulaParseFailure::new(
                self.source.len(),
                "Quoted worksheet qualifier is not closed.",
            ));
        }
        if !is_bare_qualifier_character(character) {
            return Ok(None);
        }

        let mut cursor = start;
        while cursor < self.source.len() {
            let current = self.source[cursor..].chars().next().ok_or_else(|| {
                FormulaParseFailure::new(cursor, "Worksheet qualifier has an invalid boundary.")
            })?;
            if current == '!' {
                if cursor == start {
                    return Ok(None);
                }
                let qualifier =
                    parse_qualifier_value(self.source[start..cursor].to_string(), start)?;
                return Ok(Some((qualifier, cursor + 1)));
            }
            if !is_bare_qualifier_character(current) {
                return Ok(None);
            }
            cursor += current.len_utf8();
        }
        Ok(None)
    }

    fn try_cell_reference(
        &self,
        start: usize,
        qualifier: Option<SpreadsheetFormulaQualifier>,
        qualified: bool,
    ) -> Result<Option<(SpreadsheetFormulaReference, usize)>, FormulaParseFailure> {
        let bytes = self.source.as_bytes();
        let mut cursor = start;
        let absolute_column = bytes.get(cursor) == Some(&b'$');
        cursor += usize::from(absolute_column);
        let column_start = cursor;
        while bytes.get(cursor).is_some_and(u8::is_ascii_alphabetic) {
            cursor += 1;
        }
        let column_length = cursor.saturating_sub(column_start);
        if column_length == 0 || column_length > 3 {
            return Ok(None);
        }
        let absolute_row = bytes.get(cursor) == Some(&b'$');
        cursor += usize::from(absolute_row);
        let row_start = cursor;
        while bytes.get(cursor).is_some_and(u8::is_ascii_digit) {
            cursor += 1;
        }
        if cursor == row_start {
            return Ok(None);
        }
        if !qualified && bytes.get(cursor) == Some(&b'(') {
            return Ok(None);
        }
        if self
            .source
            .get(cursor..)
            .and_then(|value| value.chars().next())
            .is_some_and(is_name_continue)
        {
            return Ok(None);
        }
        let column = parse_ascii_column(&self.source[column_start..column_start + column_length])
            .filter(|value| *value <= MAX_COLUMNS)
            .ok_or_else(|| {
                FormulaParseFailure::new(column_start, "Cell reference column is outside A:XFD.")
            })?;
        let row = self.source[row_start..cursor]
            .parse::<u32>()
            .ok()
            .filter(|value| (1..=MAX_ROWS).contains(value))
            .ok_or_else(|| {
                FormulaParseFailure::new(row_start, "Cell reference row is outside 1:1048576.")
            })?;
        Ok(Some((
            SpreadsheetFormulaReference {
                qualifier,
                kind: SpreadsheetFormulaReferenceKind::Cell {
                    column,
                    row,
                    absolute_column,
                    absolute_row,
                },
            },
            cursor,
        )))
    }

    fn lex_absolute_axis(
        &self,
        qualifier: Option<SpreadsheetFormulaQualifier>,
    ) -> Result<(SpreadsheetFormulaReference, usize), FormulaParseFailure> {
        let start = self.cursor;
        let after_dollar = start + 1;
        let Some(character) = self.source[after_dollar..].chars().next() else {
            return Err(FormulaParseFailure::new(
                start,
                "Absolute reference marker must be followed by a row or column.",
            ));
        };
        if character.is_ascii_alphabetic() {
            let mut cursor = after_dollar;
            while self
                .source
                .as_bytes()
                .get(cursor)
                .is_some_and(u8::is_ascii_alphabetic)
            {
                cursor += 1;
            }
            if self
                .source
                .get(cursor..)
                .and_then(|value| value.chars().next())
                .is_some_and(is_name_continue)
            {
                return Err(FormulaParseFailure::new(
                    cursor,
                    "Absolute column reference has invalid trailing characters.",
                ));
            }
            let column = parse_ascii_column(&self.source[after_dollar..cursor])
                .filter(|value| *value <= MAX_COLUMNS)
                .ok_or_else(|| {
                    FormulaParseFailure::new(after_dollar, "Column reference is outside A:XFD.")
                })?;
            return Ok((
                SpreadsheetFormulaReference {
                    qualifier,
                    kind: SpreadsheetFormulaReferenceKind::Column {
                        column,
                        absolute: true,
                    },
                },
                cursor,
            ));
        }
        if character.is_ascii_digit() {
            let mut cursor = after_dollar;
            while self
                .source
                .as_bytes()
                .get(cursor)
                .is_some_and(u8::is_ascii_digit)
            {
                cursor += 1;
            }
            let row = self.source[after_dollar..cursor]
                .parse::<u32>()
                .ok()
                .filter(|value| (1..=MAX_ROWS).contains(value))
                .ok_or_else(|| {
                    FormulaParseFailure::new(after_dollar, "Row reference is outside 1:1048576.")
                })?;
            return Ok((
                SpreadsheetFormulaReference {
                    qualifier,
                    kind: SpreadsheetFormulaReferenceKind::Row {
                        row,
                        absolute: true,
                    },
                },
                cursor,
            ));
        }
        Err(FormulaParseFailure::new(
            after_dollar,
            "Absolute reference marker must be followed by an ASCII row or column.",
        ))
    }

    fn lex_number(&mut self) -> Result<FormulaToken, FormulaParseFailure> {
        let start = self.cursor;
        let bytes = self.source.as_bytes();
        while bytes.get(self.cursor).is_some_and(u8::is_ascii_digit) {
            self.cursor += 1;
        }
        if bytes.get(self.cursor) == Some(&b'.') {
            self.cursor += 1;
            while bytes.get(self.cursor).is_some_and(u8::is_ascii_digit) {
                self.cursor += 1;
            }
        }
        if matches!(bytes.get(self.cursor), Some(b'e' | b'E')) {
            let exponent = self.cursor;
            self.cursor += 1;
            if matches!(bytes.get(self.cursor), Some(b'+' | b'-')) {
                self.cursor += 1;
            }
            let digits = self.cursor;
            while bytes.get(self.cursor).is_some_and(u8::is_ascii_digit) {
                self.cursor += 1;
            }
            if self.cursor == digits {
                return Err(FormulaParseFailure::new(
                    exponent,
                    "Numeric exponent has no digits.",
                ));
            }
        }
        if self
            .source
            .get(self.cursor..)
            .and_then(|value| value.chars().next())
            .is_some_and(is_name_continue)
        {
            return Err(FormulaParseFailure::new(
                self.cursor,
                "Numeric literal has invalid trailing characters.",
            ));
        }
        let raw = &self.source[start..self.cursor];
        if !raw.parse::<f64>().ok().is_some_and(f64::is_finite) {
            return Err(FormulaParseFailure::new(
                start,
                "Numeric literal must be finite.",
            ));
        }
        Ok(FormulaToken {
            kind: FormulaTokenKind::Number(raw.to_string()),
            span: SpreadsheetFormulaSpan::new(start, self.cursor),
        })
    }

    fn lex_name(
        &mut self,
        qualifier: Option<SpreadsheetFormulaQualifier>,
        span_start: usize,
    ) -> Result<FormulaToken, FormulaParseFailure> {
        let name_start = self.cursor;
        if self.source[self.cursor..].starts_with('[') {
            self.scan_bracket_groups()?;
        } else {
            let Some(first) = self.current_character() else {
                return Err(FormulaParseFailure::new(
                    self.cursor,
                    "Expected a formula name.",
                ));
            };
            if !(is_name_start(first) || (qualifier.is_some() && first.is_ascii_digit())) {
                return Err(FormulaParseFailure::new(
                    self.cursor,
                    "Worksheet qualifier must be followed by a reference or name.",
                ));
            }
            self.cursor += first.len_utf8();
            self.consume_while(|character| {
                is_name_continue(character) && !matches!(character, '[' | ']')
            });
            while self.source[self.cursor..].starts_with('[') {
                self.scan_bracket_groups()?;
            }
        }
        let name = self.source[name_start..self.cursor].to_string();
        if name.is_empty() {
            return Err(FormulaParseFailure::new(
                name_start,
                "Formula name is empty.",
            ));
        }
        let kind = if name.contains('[') {
            FormulaTokenKind::StructuredReference {
                qualifier,
                reference: name,
            }
        } else {
            FormulaTokenKind::Name { qualifier, name }
        };
        Ok(FormulaToken {
            kind,
            span: SpreadsheetFormulaSpan::new(span_start, self.cursor),
        })
    }

    fn scan_bracket_groups(&mut self) -> Result<(), FormulaParseFailure> {
        while self.source[self.cursor..].starts_with('[') {
            let start = self.cursor;
            let mut depth = 0_usize;
            while self.cursor < self.source.len() {
                let character = self.current_character().ok_or_else(|| {
                    FormulaParseFailure::new(
                        self.cursor,
                        "Structured reference has an invalid boundary.",
                    )
                })?;
                self.cursor += character.len_utf8();
                if character == '\'' {
                    let escaped = self.current_character().ok_or_else(|| {
                        FormulaParseFailure::new(
                            self.cursor,
                            "Structured reference escape has no following character.",
                        )
                    })?;
                    self.cursor += escaped.len_utf8();
                    continue;
                }
                match character {
                    '[' => depth = depth.saturating_add(1),
                    ']' => {
                        depth = depth.saturating_sub(1);
                        if depth == 0 {
                            break;
                        }
                    }
                    _ => {}
                }
            }
            if depth != 0 {
                return Err(FormulaParseFailure::new(
                    start,
                    "Structured reference bracket is not closed.",
                ));
            }
        }
        Ok(())
    }

    fn lex_error_literal(&mut self) -> Result<Option<FormulaToken>, FormulaParseFailure> {
        const ERRORS: &[SpreadsheetFormulaErrorLiteral] = &[
            SpreadsheetFormulaErrorLiteral::GettingData,
            SpreadsheetFormulaErrorLiteral::DivisionByZero,
            SpreadsheetFormulaErrorLiteral::NotAvailable,
            SpreadsheetFormulaErrorLiteral::Calculation,
            SpreadsheetFormulaErrorLiteral::Reference,
            SpreadsheetFormulaErrorLiteral::Blocked,
            SpreadsheetFormulaErrorLiteral::Unknown,
            SpreadsheetFormulaErrorLiteral::Connect,
            SpreadsheetFormulaErrorLiteral::Python,
            SpreadsheetFormulaErrorLiteral::Value,
            SpreadsheetFormulaErrorLiteral::Field,
            SpreadsheetFormulaErrorLiteral::Spill,
            SpreadsheetFormulaErrorLiteral::Number,
            SpreadsheetFormulaErrorLiteral::Name,
            SpreadsheetFormulaErrorLiteral::Busy,
            SpreadsheetFormulaErrorLiteral::Null,
        ];
        let start = self.cursor;
        for literal in ERRORS {
            let raw = literal.as_str();
            let Some(candidate) = self.source.get(start..start.saturating_add(raw.len())) else {
                continue;
            };
            if candidate.eq_ignore_ascii_case(raw) {
                self.cursor += raw.len();
                return Ok(Some(FormulaToken {
                    kind: FormulaTokenKind::Error(*literal),
                    span: SpreadsheetFormulaSpan::new(start, self.cursor),
                }));
            }
        }
        Ok(None)
    }

    fn current_character(&self) -> Option<char> {
        self.source.get(self.cursor..)?.chars().next()
    }

    fn next_character(&self) -> Option<char> {
        let current = self.current_character()?;
        self.source
            .get(self.cursor + current.len_utf8()..)?
            .chars()
            .next()
    }

    fn consume_while(&mut self, predicate: impl Fn(char) -> bool) {
        while let Some(character) = self.current_character() {
            if !predicate(character) {
                break;
            }
            self.cursor += character.len_utf8();
        }
    }

    fn push_token(&mut self, token: FormulaToken) -> Result<(), FormulaParseFailure> {
        if self.tokens.len() >= MAX_LEXICAL_TOKENS {
            return Err(FormulaParseFailure::new(
                token.span.start,
                "Formula contains too many lexical tokens.",
            ));
        }
        self.tokens.push(token);
        Ok(())
    }

    fn push(
        &mut self,
        kind: FormulaTokenKind,
        start: usize,
        end: usize,
    ) -> Result<(), FormulaParseFailure> {
        self.push_token(FormulaToken {
            kind,
            span: SpreadsheetFormulaSpan::new(start, end),
        })
    }
}

fn parse_qualifier_value(
    value: String,
    position: usize,
) -> Result<SpreadsheetFormulaQualifier, FormulaParseFailure> {
    if value.is_empty() {
        return Err(FormulaParseFailure::new(
            position,
            "Worksheet qualifier is empty.",
        ));
    }
    let (workbook, worksheets) = value.rfind(']').map_or((None, value.as_str()), |end| {
        (
            Some(value[..=end].to_string()),
            value.get(end + 1..).unwrap_or_default(),
        )
    });
    if worksheets.is_empty() {
        return Err(FormulaParseFailure::new(
            position,
            "Worksheet qualifier has no worksheet name.",
        ));
    }
    let (worksheet, worksheet_end) = worksheets
        .split_once(':')
        .map_or((worksheets, None), |(start, end)| (start, Some(end)));
    if worksheet.is_empty() || worksheet_end.is_some_and(str::is_empty) {
        return Err(FormulaParseFailure::new(
            position,
            "Three-dimensional worksheet qualifier has an empty endpoint.",
        ));
    }
    Ok(SpreadsheetFormulaQualifier {
        workbook,
        worksheet: worksheet.to_string(),
        worksheet_end: worksheet_end.map(ToOwned::to_owned),
    })
}

fn parse_ascii_column(value: &str) -> Option<u32> {
    value.bytes().try_fold(0_u32, |column, byte| {
        column.checked_mul(26).and_then(|column| {
            column.checked_add(u32::from(byte.to_ascii_uppercase().checked_sub(b'A')?) + 1)
        })
    })
}

fn is_bare_qualifier_character(character: char) -> bool {
    character.is_alphanumeric() || matches!(character, '_' | '.' | '\\' | '[' | ']' | ':' | '$')
}

fn is_name_start(character: char) -> bool {
    character.is_alphabetic() || matches!(character, '_' | '\\' | '?')
}

fn is_name_continue(character: char) -> bool {
    character.is_alphanumeric() || matches!(character, '_' | '.' | '\\' | '?')
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn lexer_decodes_literals_and_qualified_references_with_utf8_spans() {
        let source = r#""a""b"+'销售 数据'!$C$3+[Book.xlsx]Data!A1"#;
        let tokens = lex(source).unwrap();
        assert_eq!(tokens[0].kind, FormulaTokenKind::Text("a\"b".to_string()));
        assert_eq!(
            &source[tokens[0].span.start..tokens[0].span.end],
            r#""a""b""#
        );
        assert!(matches!(tokens[2].kind, FormulaTokenKind::Reference(_)));
        assert!(matches!(tokens[4].kind, FormulaTokenKind::Reference(_)));
        let FormulaTokenKind::Reference(reference) = &tokens[2].kind else {
            unreachable!();
        };
        let qualifier = reference.qualifier.as_ref().unwrap();
        assert_eq!(qualifier.worksheet, "销售 数据");
        assert!(!qualifier.is_external());
        let FormulaTokenKind::Reference(reference) = &tokens[4].kind else {
            unreachable!();
        };
        assert!(reference.qualifier.as_ref().unwrap().is_external());
    }

    #[test]
    fn lexer_rejects_unclosed_strings_qualifiers_and_structured_references() {
        for source in ["\"open", "'Sheet 1!A1", "Table1[[Column]"] {
            assert!(lex(source).is_err(), "{source}");
        }
    }
}
