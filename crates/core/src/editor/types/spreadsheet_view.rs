use a3s_use_core::UseResult;
use serde::{Deserialize, Serialize};

use crate::semantic::{DocumentNode, OfficeNodeType};
use crate::spreadsheet_reference::CellReference;

const MAX_ROWS: u32 = 1_048_576;
const MAX_COLUMNS: u32 = 16_384;

/// One canonical frozen Spreadsheet pane.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct NativeSpreadsheetFrozenPane {
    pub frozen_rows: u32,
    pub frozen_columns: u32,
    pub top_left_cell: String,
}

impl NativeSpreadsheetFrozenPane {
    pub fn new(frozen_rows: u32, frozen_columns: u32, top_left_cell: impl Into<String>) -> Self {
        Self {
            frozen_rows,
            frozen_columns,
            top_left_cell: top_left_cell.into(),
        }
    }

    pub(crate) fn normalized(&self) -> UseResult<Self> {
        if self.frozen_rows == 0 && self.frozen_columns == 0 {
            return Err(view_error(
                "use.office.spreadsheet_freeze_empty",
                "A frozen Spreadsheet pane requires at least one frozen row or column.",
            ));
        }
        if self.frozen_rows >= MAX_ROWS {
            return Err(view_error(
                "use.office.spreadsheet_freeze_row_limit",
                format!("Frozen Spreadsheet rows must be below {MAX_ROWS}."),
            ));
        }
        if self.frozen_columns >= MAX_COLUMNS {
            return Err(view_error(
                "use.office.spreadsheet_freeze_column_limit",
                format!("Frozen Spreadsheet columns must be below {MAX_COLUMNS}."),
            ));
        }
        let top_left = CellReference::parse(&self.top_left_cell).map_err(|error| {
            view_error(
                "use.office.spreadsheet_freeze_cell_invalid",
                format!(
                    "Frozen Spreadsheet topLeftCell '{}' is invalid: {error}",
                    self.top_left_cell
                ),
            )
        })?;
        if top_left.row <= self.frozen_rows || top_left.column <= self.frozen_columns {
            return Err(view_error(
                "use.office.spreadsheet_freeze_geometry_invalid",
                "Frozen Spreadsheet topLeftCell must be below and to the right of every frozen split.",
            ));
        }
        Ok(Self {
            frozen_rows: self.frozen_rows,
            frozen_columns: self.frozen_columns,
            top_left_cell: top_left.a1(),
        })
    }

    pub fn from_semantic_node(node: &DocumentNode) -> UseResult<Self> {
        if node.node_type != OfficeNodeType::FrozenPane {
            return Err(view_error(
                "use.office.spreadsheet_freeze_node_invalid",
                format!(
                    "Office node '{}' is not a frozen Spreadsheet pane.",
                    node.path
                ),
            ));
        }
        let parse = |key: &str| -> UseResult<u32> {
            node.format
                .get(key)
                .ok_or_else(|| {
                    view_error(
                        "use.office.spreadsheet_freeze_node_invalid",
                        format!(
                            "Frozen Spreadsheet pane '{}' has no {key} value.",
                            node.path
                        ),
                    )
                })?
                .parse::<u32>()
                .map_err(|error| {
                    view_error(
                        "use.office.spreadsheet_freeze_node_invalid",
                        format!(
                            "Frozen Spreadsheet pane '{}' has invalid {key}: {error}",
                            node.path
                        ),
                    )
                })
        };
        Self::new(
            parse("frozenRows")?,
            parse("frozenColumns")?,
            node.format.get("topLeftCell").cloned().ok_or_else(|| {
                view_error(
                    "use.office.spreadsheet_freeze_node_invalid",
                    format!(
                        "Frozen Spreadsheet pane '{}' has no topLeftCell value.",
                        node.path
                    ),
                )
            })?,
        )
        .normalized()
    }

    pub(crate) fn active_pane(&self) -> &'static str {
        match (self.frozen_rows > 0, self.frozen_columns > 0) {
            (true, true) => "bottomRight",
            (true, false) => "bottomLeft",
            (false, true) => "topRight",
            (false, false) => "topLeft",
        }
    }
}

fn view_error(code: &str, message: impl Into<String>) -> a3s_use_core::UseError {
    super::super::editor_error(code, message)
}
