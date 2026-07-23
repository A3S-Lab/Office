use serde::{Deserialize, Serialize};

use super::{SpreadsheetFormulaCell, SpreadsheetFormulaErrorLiteral};

/// Scalar or rectangular dynamic-array result produced by native calculation.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(tag = "type", rename_all = "kebab-case", deny_unknown_fields)]
pub enum SpreadsheetFormulaValue {
    Blank,
    Number {
        value: String,
    },
    Text {
        value: String,
    },
    Boolean {
        value: bool,
    },
    Error {
        error: SpreadsheetFormulaErrorLiteral,
    },
    Array {
        rows: Vec<Vec<SpreadsheetFormulaValue>>,
    },
}

impl SpreadsheetFormulaValue {
    pub fn error(error: SpreadsheetFormulaErrorLiteral) -> Self {
        Self::Error { error }
    }

    pub fn error_literal(&self) -> Option<&'static str> {
        match self {
            Self::Error { error } => Some(error.as_str()),
            _ => None,
        }
    }
}

/// One calculated formula anchor and its optional spill extent.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct SpreadsheetFormulaCalculatedCell {
    pub cell: SpreadsheetFormulaCell,
    pub value: SpreadsheetFormulaValue,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub spill_range: Option<String>,
}

/// Deterministic read-only result from one native workbook calculation pass.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct SpreadsheetFormulaCalculation {
    pub formula_count: usize,
    pub spill_cell_count: usize,
    pub calculation_order: Vec<SpreadsheetFormulaCell>,
    pub cells: Vec<SpreadsheetFormulaCalculatedCell>,
}
