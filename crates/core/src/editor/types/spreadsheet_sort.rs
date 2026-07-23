use std::collections::BTreeSet;

use a3s_use_core::UseResult;
use serde::{Deserialize, Serialize};

use crate::spreadsheet_reference::{parse_column, CellRange};
use crate::{DocumentNode, OfficeNodeType};

pub(crate) const MAX_SPREADSHEET_SORT_KEYS: usize = 64;
pub(crate) const MAX_SPREADSHEET_SORT_CELLS: usize = 100_000;

/// Direction for one Spreadsheet sort key.
#[derive(Debug, Default, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "kebab-case")]
pub enum NativeSpreadsheetSortDirection {
    #[default]
    Ascending,
    Descending,
}

impl NativeSpreadsheetSortDirection {
    pub(crate) const fn ooxml_descending(self) -> bool {
        matches!(self, Self::Descending)
    }

    pub(crate) const fn semantic_name(self) -> &'static str {
        match self {
            Self::Ascending => "ascending",
            Self::Descending => "descending",
        }
    }
}

/// One absolute A:XFD column used by a stable multi-key Spreadsheet sort.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct NativeSpreadsheetSortKey {
    pub column: String,
    #[serde(default)]
    pub direction: NativeSpreadsheetSortDirection,
}

impl NativeSpreadsheetSortKey {
    pub fn ascending(column: impl Into<String>) -> Self {
        Self {
            column: column.into(),
            direction: NativeSpreadsheetSortDirection::Ascending,
        }
    }

    pub fn descending(column: impl Into<String>) -> Self {
        Self {
            column: column.into(),
            direction: NativeSpreadsheetSortDirection::Descending,
        }
    }
}

/// A complete deterministic physical row sort request.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct NativeSpreadsheetSort {
    pub keys: Vec<NativeSpreadsheetSortKey>,
    #[serde(default)]
    pub header: bool,
    #[serde(default)]
    pub case_sensitive: bool,
}

impl NativeSpreadsheetSort {
    pub fn new(keys: Vec<NativeSpreadsheetSortKey>) -> Self {
        Self {
            keys,
            header: false,
            case_sensitive: false,
        }
    }

    pub fn with_header(mut self, header: bool) -> Self {
        self.header = header;
        self
    }

    pub fn with_case_sensitive(mut self, case_sensitive: bool) -> Self {
        self.case_sensitive = case_sensitive;
        self
    }

    pub(crate) fn validate(
        &self,
        range: CellRange,
    ) -> UseResult<Vec<(u32, NativeSpreadsheetSortDirection)>> {
        if self.keys.is_empty() || self.keys.len() > MAX_SPREADSHEET_SORT_KEYS {
            return Err(sort_error(
                "use.office.spreadsheet_sort_key_limit",
                format!("Spreadsheet sorting requires 1-{MAX_SPREADSHEET_SORT_KEYS} ordered keys."),
            ));
        }
        if range.cell_count()? > MAX_SPREADSHEET_SORT_CELLS {
            return Err(sort_error(
                "use.office.spreadsheet_sort_range_limit",
                format!(
                    "Spreadsheet sorting accepts at most {MAX_SPREADSHEET_SORT_CELLS} cells per range."
                ),
            )
            .with_detail("range", range.a1()));
        }
        let mut observed = BTreeSet::new();
        let mut parsed = Vec::with_capacity(self.keys.len());
        for key in &self.keys {
            let column = parse_column(key.column.trim()).map_err(|error| {
                sort_error(
                    "use.office.spreadsheet_sort_column_invalid",
                    format!(
                        "Spreadsheet sort column '{}' is invalid: {error}",
                        key.column
                    ),
                )
            })?;
            if !(range.start.column..=range.end.column).contains(&column) {
                return Err(sort_error(
                    "use.office.spreadsheet_sort_column_outside_range",
                    format!(
                        "Spreadsheet sort column '{}' is outside range '{}'.",
                        key.column,
                        range.a1()
                    ),
                )
                .with_detail("column", key.column.clone())
                .with_detail("range", range.a1()));
            }
            if !observed.insert(column) {
                return Err(sort_error(
                    "use.office.spreadsheet_sort_column_duplicate",
                    format!(
                        "Spreadsheet sort column '{}' is defined more than once.",
                        key.column
                    ),
                ));
            }
            parsed.push((column, key.direction));
        }
        Ok(parsed)
    }

    /// Reconstructs the typed value from a supported `/Sheet/sort` semantic node.
    pub fn from_semantic_node(node: &DocumentNode) -> UseResult<Self> {
        if node.node_type != OfficeNodeType::SortState
            || node.format.get("nativeMutable").map(String::as_str) != Some("true")
        {
            return Err(sort_error(
                "use.office.spreadsheet_sort_unknown_content",
                format!(
                    "Spreadsheet sort state '{}' contains unsupported or unknown content.",
                    node.path
                ),
            ));
        }
        let header = parse_semantic_bool(node, "header")?;
        let case_sensitive = parse_semantic_bool(node, "caseSensitive")?;
        let mut keys = Vec::with_capacity(node.children.len());
        for child in &node.children {
            if child.node_type != OfficeNodeType::SortKey {
                return Err(sort_error(
                    "use.office.spreadsheet_sort_unknown_content",
                    format!(
                        "Spreadsheet sort state '{}' has a non-key child.",
                        node.path
                    ),
                ));
            }
            let column = child.format.get("column").cloned().ok_or_else(|| {
                sort_error(
                    "use.office.spreadsheet_sort_invalid",
                    format!("Spreadsheet sort key '{}' has no column.", child.path),
                )
            })?;
            let direction = match child.format.get("direction").map(String::as_str) {
                Some("ascending") => NativeSpreadsheetSortDirection::Ascending,
                Some("descending") => NativeSpreadsheetSortDirection::Descending,
                _ => {
                    return Err(sort_error(
                        "use.office.spreadsheet_sort_invalid",
                        format!(
                            "Spreadsheet sort key '{}' has an invalid direction.",
                            child.path
                        ),
                    ))
                }
            };
            keys.push(NativeSpreadsheetSortKey { column, direction });
        }
        Ok(Self {
            keys,
            header,
            case_sensitive,
        })
    }
}

fn parse_semantic_bool(node: &DocumentNode, key: &str) -> UseResult<bool> {
    match node.format.get(key).map(String::as_str) {
        Some("true") => Ok(true),
        Some("false") => Ok(false),
        _ => Err(sort_error(
            "use.office.spreadsheet_sort_invalid",
            format!(
                "Spreadsheet sort state '{}' has invalid {key} metadata.",
                node.path
            ),
        )),
    }
}

fn sort_error(code: &str, message: impl Into<String>) -> a3s_use_core::UseError {
    super::super::editor_error(code, message)
}
