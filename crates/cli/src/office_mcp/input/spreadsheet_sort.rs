use a3s_office::{NativeSpreadsheetSort, NativeSpreadsheetSortDirection, NativeSpreadsheetSortKey};
use serde::{Deserialize, Serialize};

#[derive(
    Debug, Default, Clone, Copy, PartialEq, Eq, Serialize, Deserialize, schemars::JsonSchema,
)]
#[serde(rename_all = "kebab-case")]
pub(in crate::office_mcp) enum OfficeSpreadsheetSortDirection {
    #[default]
    Ascending,
    Descending,
}

impl From<OfficeSpreadsheetSortDirection> for NativeSpreadsheetSortDirection {
    fn from(value: OfficeSpreadsheetSortDirection) -> Self {
        match value {
            OfficeSpreadsheetSortDirection::Ascending => Self::Ascending,
            OfficeSpreadsheetSortDirection::Descending => Self::Descending,
        }
    }
}

#[derive(Debug, Clone, Serialize, Deserialize, schemars::JsonSchema)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub(in crate::office_mcp) struct OfficeSpreadsheetSortKey {
    /// Absolute Spreadsheet column from A through XFD.
    column: String,
    #[serde(default)]
    direction: OfficeSpreadsheetSortDirection,
}

impl From<OfficeSpreadsheetSortKey> for NativeSpreadsheetSortKey {
    fn from(value: OfficeSpreadsheetSortKey) -> Self {
        Self {
            column: value.column,
            direction: value.direction.into(),
        }
    }
}

#[derive(Debug, Clone, Serialize, Deserialize, schemars::JsonSchema)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub(in crate::office_mcp) struct OfficeSpreadsheetSort {
    /// Ordered stable sort keys. The first key has the highest precedence.
    keys: Vec<OfficeSpreadsheetSortKey>,
    /// Keep the range's first row fixed as a header.
    #[serde(default)]
    header: bool,
    /// Compare text with case distinctions. Defaults to false.
    #[serde(default)]
    case_sensitive: bool,
}

impl From<OfficeSpreadsheetSort> for NativeSpreadsheetSort {
    fn from(value: OfficeSpreadsheetSort) -> Self {
        Self {
            keys: value.keys.into_iter().map(Into::into).collect(),
            header: value.header,
            case_sensitive: value.case_sensitive,
        }
    }
}
