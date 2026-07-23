use std::collections::BTreeSet;

use a3s_use_core::UseResult;
use regex::Regex;
use serde::{Deserialize, Serialize};

use crate::spreadsheet_reference::CellRange;

use super::{NativeSpreadsheetAutoFilter, NativeSpreadsheetFilterColumn};

const MAX_TABLE_NAME_CHARACTERS: usize = 255;
const MAX_TABLE_COLUMN_NAME_CHARACTERS: usize = 255;

/// One closed built-in Spreadsheet table-style identity.
///
/// OOXML defines 21 light, 28 medium, and 11 dark built-in styles. `None`
/// intentionally omits `tableStyleInfo` rather than accepting an arbitrary
/// style name that a host application may silently discard.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(tag = "family", rename_all = "kebab-case", deny_unknown_fields)]
pub enum NativeSpreadsheetTableStyle {
    None,
    Light { number: u8 },
    Medium { number: u8 },
    Dark { number: u8 },
}

impl Default for NativeSpreadsheetTableStyle {
    fn default() -> Self {
        Self::Medium { number: 2 }
    }
}

impl NativeSpreadsheetTableStyle {
    pub(crate) fn validate(self) -> UseResult<()> {
        let valid = match self {
            Self::None => true,
            Self::Light { number } => (1..=21).contains(&number),
            Self::Medium { number } => (1..=28).contains(&number),
            Self::Dark { number } => (1..=11).contains(&number),
        };
        if valid {
            Ok(())
        } else {
            Err(super::super::editor_error(
                "use.office.spreadsheet_table_style_invalid",
                "Spreadsheet table styles require light 1-21, medium 1-28, dark 1-11, or none.",
            ))
        }
    }

    pub(crate) fn ooxml_name(self) -> Option<String> {
        match self {
            Self::None => None,
            Self::Light { number } => Some(format!("TableStyleLight{number}")),
            Self::Medium { number } => Some(format!("TableStyleMedium{number}")),
            Self::Dark { number } => Some(format!("TableStyleDark{number}")),
        }
    }

    pub(crate) fn from_ooxml_name(value: Option<&str>) -> Option<Self> {
        let value = value?;
        let style = if let Some(number) = value
            .strip_prefix("TableStyleLight")
            .and_then(|suffix| suffix.parse::<u8>().ok())
        {
            Self::Light { number }
        } else if let Some(number) = value
            .strip_prefix("TableStyleMedium")
            .and_then(|suffix| suffix.parse::<u8>().ok())
        {
            Self::Medium { number }
        } else {
            let number = value
                .strip_prefix("TableStyleDark")
                .and_then(|suffix| suffix.parse::<u8>().ok())?;
            Self::Dark { number }
        };
        style.validate().ok().map(|()| style)
    }
}

/// One column identity owned by a native Spreadsheet ListObject table.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct NativeSpreadsheetTableColumn {
    pub name: String,
}

impl NativeSpreadsheetTableColumn {
    pub fn new(name: impl Into<String>) -> Self {
        Self { name: name.into() }
    }
}

/// A complete typed Spreadsheet ListObject table value.
///
/// `range` is the final table range. When `totalsRow` is true, its last row is
/// the totals row; when `headerRow` is true, its first row is the header row.
/// At least one data row must remain between those structural rows.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct NativeSpreadsheetTable {
    pub name: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub display_name: Option<String>,
    pub range: String,
    pub columns: Vec<NativeSpreadsheetTableColumn>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub filters: Vec<NativeSpreadsheetFilterColumn>,
    #[serde(default = "default_true")]
    pub header_row: bool,
    #[serde(default)]
    pub totals_row: bool,
    #[serde(default)]
    pub style: NativeSpreadsheetTableStyle,
    #[serde(default)]
    pub show_first_column: bool,
    #[serde(default)]
    pub show_last_column: bool,
    #[serde(default = "default_true")]
    pub show_row_stripes: bool,
    #[serde(default)]
    pub show_column_stripes: bool,
}

impl NativeSpreadsheetTable {
    pub fn new(
        name: impl Into<String>,
        range: impl Into<String>,
        columns: impl IntoIterator<Item = impl Into<String>>,
    ) -> Self {
        Self {
            name: name.into(),
            display_name: None,
            range: range.into(),
            columns: columns
                .into_iter()
                .map(|name| NativeSpreadsheetTableColumn::new(name.into()))
                .collect(),
            filters: Vec::new(),
            header_row: true,
            totals_row: false,
            style: NativeSpreadsheetTableStyle::default(),
            show_first_column: false,
            show_last_column: false,
            show_row_stripes: true,
            show_column_stripes: false,
        }
    }

    pub fn with_display_name(mut self, display_name: impl Into<String>) -> Self {
        self.display_name = Some(display_name.into());
        self
    }

    pub fn with_header_row(mut self, enabled: bool) -> Self {
        self.header_row = enabled;
        self
    }

    pub fn with_totals_row(mut self, enabled: bool) -> Self {
        self.totals_row = enabled;
        self
    }

    pub fn with_filter(
        mut self,
        column: u32,
        criteria: super::NativeSpreadsheetFilterCriteria,
    ) -> Self {
        self.filters
            .push(NativeSpreadsheetFilterColumn::new(column, criteria));
        self
    }

    pub fn with_style(mut self, style: NativeSpreadsheetTableStyle) -> Self {
        self.style = style;
        self
    }

    pub fn with_first_column(mut self, enabled: bool) -> Self {
        self.show_first_column = enabled;
        self
    }

    pub fn with_last_column(mut self, enabled: bool) -> Self {
        self.show_last_column = enabled;
        self
    }

    pub fn with_row_stripes(mut self, enabled: bool) -> Self {
        self.show_row_stripes = enabled;
        self
    }

    pub fn with_column_stripes(mut self, enabled: bool) -> Self {
        self.show_column_stripes = enabled;
        self
    }

    pub(crate) fn validate(&self) -> UseResult<CellRange> {
        validate_table_name(&self.name)?;
        if let Some(display_name) = &self.display_name {
            validate_table_name(display_name)?;
        }
        self.style.validate()?;
        if self.style == NativeSpreadsheetTableStyle::None
            && (self.show_first_column
                || self.show_last_column
                || self.show_row_stripes
                || self.show_column_stripes)
        {
            return Err(super::super::editor_error(
                "use.office.spreadsheet_table_style_flags_invalid",
                "Spreadsheet table style flags require a built-in table style.",
            ));
        }
        let range = CellRange::parse(&self.range).map_err(|error| {
            super::super::editor_error(
                "use.office.spreadsheet_table_range_invalid",
                format!(
                    "Spreadsheet table range '{}' is invalid: {error}",
                    self.range
                ),
            )
        })?;
        let width = usize::try_from(range.end.column - range.start.column + 1).map_err(|_| {
            super::super::editor_error(
                "use.office.spreadsheet_table_range_invalid",
                "Spreadsheet table width does not fit this platform.",
            )
        })?;
        if self.columns.len() != width {
            return Err(super::super::editor_error(
                "use.office.spreadsheet_table_column_count_invalid",
                format!(
                    "Spreadsheet table range '{}' spans {width} columns, but {} column names were provided.",
                    range.a1(),
                    self.columns.len()
                ),
            ));
        }
        if !self.header_row && !self.filters.is_empty() {
            return Err(super::super::editor_error(
                "use.office.spreadsheet_table_filter_header_required",
                "Spreadsheet table filters require an enabled header row.",
            ));
        }
        NativeSpreadsheetAutoFilter {
            range: range.a1(),
            columns: self.filters.clone(),
        }
        .validate_for_width(width)?;
        let height = u64::from(range.end.row - range.start.row + 1);
        let structural_rows = u64::from(self.header_row) + u64::from(self.totals_row);
        if height <= structural_rows {
            return Err(super::super::editor_error(
                "use.office.spreadsheet_table_data_rows_invalid",
                "Spreadsheet tables require at least one data row in addition to enabled header and totals rows.",
            ));
        }
        let mut names = BTreeSet::new();
        for column in &self.columns {
            validate_column_name(&column.name)?;
            if !names.insert(column.name.to_lowercase()) {
                return Err(super::super::editor_error(
                    "use.office.spreadsheet_table_column_duplicate",
                    format!(
                        "Spreadsheet table column name '{}' is duplicated case-insensitively.",
                        column.name
                    ),
                ));
            }
        }
        Ok(range)
    }
}

fn validate_table_name(name: &str) -> UseResult<()> {
    let valid_length = (1..=MAX_TABLE_NAME_CHARACTERS).contains(&name.chars().count());
    let valid_grammar = Regex::new(r"^[\p{L}_\\][\p{L}\p{N}_\\.]*$")
        .map(|pattern| pattern.is_match(name))
        .unwrap_or(false);
    let reserved_r1c1 = name.eq_ignore_ascii_case("R") || name.eq_ignore_ascii_case("C");
    let resembles_a1 = CellRange::parse(name).is_ok();
    let resembles_r1c1 = Regex::new(r"(?i)^R[1-9][0-9]*C[1-9][0-9]*$")
        .map(|pattern| pattern.is_match(name))
        .unwrap_or(false);
    if valid_length && valid_grammar && !reserved_r1c1 && !resembles_a1 && !resembles_r1c1 {
        Ok(())
    } else {
        Err(super::super::editor_error(
            "use.office.spreadsheet_table_name_invalid",
            "Spreadsheet table names must be 1-255 characters, follow Excel identifier grammar, and not resemble A1 or R1C1 references.",
        ))
    }
}

fn validate_column_name(name: &str) -> UseResult<()> {
    if name.trim() == name
        && (1..=MAX_TABLE_COLUMN_NAME_CHARACTERS).contains(&name.chars().count())
        && !name
            .chars()
            .any(|character| character.is_control() || matches!(character, '\u{fffe}' | '\u{ffff}'))
    {
        Ok(())
    } else {
        Err(super::super::editor_error(
            "use.office.spreadsheet_table_column_name_invalid",
            "Spreadsheet table column names must contain 1-255 non-control characters without surrounding whitespace.",
        ))
    }
}

const fn default_true() -> bool {
    true
}
