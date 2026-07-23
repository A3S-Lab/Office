use a3s_office::NativeSpreadsheetFrozenPane;
use serde::{Deserialize, Serialize};

#[derive(Debug, Clone, Serialize, Deserialize, schemars::JsonSchema)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub(in crate::office_mcp) struct OfficeSpreadsheetFrozenPane {
    /// Number of complete rows frozen above the scrollable pane.
    frozen_rows: u32,
    /// Number of complete columns frozen to the left of the scrollable pane.
    frozen_columns: u32,
    /// First visible cell in the scrollable pane.
    top_left_cell: String,
}

impl OfficeSpreadsheetFrozenPane {
    pub(super) fn into_native(self) -> NativeSpreadsheetFrozenPane {
        NativeSpreadsheetFrozenPane::new(self.frozen_rows, self.frozen_columns, self.top_left_cell)
    }
}
