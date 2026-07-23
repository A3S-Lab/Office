use a3s_office::{NativeSpreadsheetDelimitedFormat, NativeSpreadsheetDelimitedImport};
use serde::{Deserialize, Serialize};

#[derive(Debug, Clone, Copy, Serialize, Deserialize, schemars::JsonSchema)]
#[serde(rename_all = "kebab-case")]
pub(in crate::office_mcp) enum OfficeSpreadsheetDelimitedFormat {
    Csv,
    Tsv,
}

#[derive(Debug, Clone, Serialize, Deserialize, schemars::JsonSchema)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub(in crate::office_mcp) struct OfficeSpreadsheetDelimitedImport {
    /// Bounded UTF-8 CSV or TSV content; filesystem paths belong at the CLI boundary.
    content: String,
    format: OfficeSpreadsheetDelimitedFormat,
    /// Treat the first imported row as headers, add an AutoFilter, and freeze below it.
    #[serde(default)]
    header: bool,
    /// A1 cell at which the first source field is written. Defaults to A1.
    start_cell: Option<String>,
}

impl OfficeSpreadsheetDelimitedImport {
    pub(super) fn into_native(self) -> NativeSpreadsheetDelimitedImport {
        NativeSpreadsheetDelimitedImport::new(
            self.content,
            match self.format {
                OfficeSpreadsheetDelimitedFormat::Csv => NativeSpreadsheetDelimitedFormat::Csv,
                OfficeSpreadsheetDelimitedFormat::Tsv => NativeSpreadsheetDelimitedFormat::Tsv,
            },
        )
        .with_header(self.header)
        .with_start_cell(self.start_cell.unwrap_or_else(|| "A1".into()))
    }
}
