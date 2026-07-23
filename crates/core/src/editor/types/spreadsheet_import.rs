use a3s_use_core::UseResult;
use serde::{Deserialize, Serialize};

use crate::spreadsheet_reference::CellReference;

/// Maximum UTF-8 source bytes accepted by one native delimited import.
pub const MAX_NATIVE_SPREADSHEET_IMPORT_BYTES: usize = 8 * 1024 * 1024;
/// Maximum rectangular target cells admitted by one native delimited import.
pub const MAX_NATIVE_SPREADSHEET_IMPORT_CELLS: usize = 100_000;

/// Closed source syntax accepted by native Spreadsheet import.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "kebab-case")]
pub enum NativeSpreadsheetDelimitedFormat {
    Csv,
    Tsv,
}

impl NativeSpreadsheetDelimitedFormat {
    pub(crate) fn delimiter(self) -> char {
        match self {
            Self::Csv => ',',
            Self::Tsv => '\t',
        }
    }
}

/// Bounded, filesystem-independent source for one Spreadsheet import.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct NativeSpreadsheetDelimitedImport {
    pub content: String,
    pub format: NativeSpreadsheetDelimitedFormat,
    #[serde(default)]
    pub header: bool,
    #[serde(default = "default_start_cell")]
    pub start_cell: String,
}

impl NativeSpreadsheetDelimitedImport {
    pub fn new(content: impl Into<String>, format: NativeSpreadsheetDelimitedFormat) -> Self {
        Self {
            content: content.into(),
            format,
            header: false,
            start_cell: default_start_cell(),
        }
    }

    pub fn with_header(mut self, header: bool) -> Self {
        self.header = header;
        self
    }

    pub fn with_start_cell(mut self, start_cell: impl Into<String>) -> Self {
        self.start_cell = start_cell.into();
        self
    }

    pub(crate) fn validate(&self) -> UseResult<CellReference> {
        if self.content.len() > MAX_NATIVE_SPREADSHEET_IMPORT_BYTES {
            return Err(import_error(
                "use.office.spreadsheet_import_input_limit",
                format!(
                    "Native Spreadsheet delimited import accepts at most {MAX_NATIVE_SPREADSHEET_IMPORT_BYTES} UTF-8 bytes."
                ),
            )
            .with_detail("bytes", self.content.len()));
        }
        CellReference::parse(&self.start_cell).map_err(|error| {
            import_error(
                "use.office.spreadsheet_import_start_cell_invalid",
                format!(
                    "Spreadsheet import startCell '{}' is invalid: {error}",
                    self.start_cell
                ),
            )
            .with_detail("startCell", self.start_cell.clone())
        })
    }
}

/// Receipt returned for one atomic native delimited import mutation.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct NativeSpreadsheetImportResult {
    pub path: String,
    pub sheet: String,
    pub start_cell: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub range: Option<String>,
    pub format: NativeSpreadsheetDelimitedFormat,
    pub row_count: usize,
    pub column_count: usize,
    pub header: bool,
    pub changed: bool,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub filter_path: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub freeze_path: Option<String>,
}

fn default_start_cell() -> String {
    "A1".to_string()
}

fn import_error(code: &str, message: impl Into<String>) -> a3s_use_core::UseError {
    super::super::editor_error(code, message)
}
