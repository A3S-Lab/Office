use std::collections::{BTreeMap, BTreeSet};

use a3s_office_formula_parser::{
    parse_spreadsheet_formula, SpreadsheetFormulaErrorLiteral, MAX_SPREADSHEET_FORMULA_CHARACTERS,
};
use serde::{Deserialize, Serialize};

use crate::{KernelError, OFFICE_KERNEL_PROTOCOL_VERSION};

mod evaluate;
mod functions;
mod value;

const MAX_SPREADSHEET_SHEETS: usize = 1_024;
const MAX_SPREADSHEET_CELLS: usize = 100_000;
const MAX_SPREADSHEET_DEPENDENCY_DEPTH: usize = 64;
const MAX_SPREADSHEET_ROWS: u32 = 1_048_576;
const MAX_SPREADSHEET_COLUMNS: u32 = 16_384;
const MAX_SPREADSHEET_IDENTIFIER_BYTES: usize = 256;
const MAX_SPREADSHEET_TEXT_BYTES: usize = 32_767;

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(tag = "kind", rename_all = "camelCase", deny_unknown_fields)]
pub enum SpreadsheetValue {
    Blank,
    Number { value: f64 },
    Text { value: String },
    Boolean { value: bool },
    Error { value: String },
}

impl SpreadsheetValue {
    fn error(error: SpreadsheetFormulaErrorLiteral) -> Self {
        Self::Error {
            value: error.as_str().to_owned(),
        }
    }

    fn validate(&self) -> Result<(), KernelError> {
        match self {
            Self::Number { value } if !value.is_finite() => Err(KernelError::invalid(
                "office.kernel.spreadsheet.value_invalid",
                "Spreadsheet numeric values must be finite.",
            )),
            Self::Text { value } if value.len() > MAX_SPREADSHEET_TEXT_BYTES => {
                Err(KernelError::invalid(
                    "office.kernel.spreadsheet.value_invalid",
                    format!(
                        "Spreadsheet text values may contain at most {MAX_SPREADSHEET_TEXT_BYTES} UTF-8 bytes."
                    ),
                ))
            }
            Self::Error { value }
                if SpreadsheetFormulaErrorLiteral::parse(value).is_none() =>
            {
                Err(KernelError::invalid(
                    "office.kernel.spreadsheet.value_invalid",
                    format!("Spreadsheet error value '{value}' is not recognized."),
                ))
            }
            _ => Ok(()),
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq, PartialOrd, Ord, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct SpreadsheetCoordinate {
    pub sheet_id: String,
    pub row: u32,
    pub column: u32,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct SpreadsheetInputCell {
    pub row: u32,
    pub column: u32,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub formula: Option<String>,
    pub value: SpreadsheetValue,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct SpreadsheetInputSheet {
    pub id: String,
    pub name: String,
    pub cells: Vec<SpreadsheetInputCell>,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct SpreadsheetCalculationRequest {
    pub protocol: u32,
    pub kind: String,
    pub request_id: u32,
    pub revision: u32,
    pub document_revision: u64,
    pub sheets: Vec<SpreadsheetInputSheet>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub targets: Vec<SpreadsheetCoordinate>,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct SpreadsheetCalculatedCell {
    pub sheet_id: String,
    pub row: u32,
    pub column: u32,
    pub value: SpreadsheetValue,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct SpreadsheetCalculationIssue {
    pub cell: SpreadsheetCoordinate,
    pub code: String,
    pub message: String,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct SpreadsheetCalculationResult {
    pub protocol: u32,
    pub kind: String,
    pub request_id: u32,
    pub revision: u32,
    pub document_revision: u64,
    pub engine: String,
    pub cells: Vec<SpreadsheetCalculatedCell>,
    pub calculation_order: Vec<SpreadsheetCoordinate>,
    pub issues: Vec<SpreadsheetCalculationIssue>,
}

#[derive(Debug, Clone, PartialEq, Eq, PartialOrd, Ord)]
struct CellKey {
    sheet: usize,
    row: u32,
    column: u32,
}

#[derive(Debug, Clone)]
struct IndexedCell {
    formula: Option<String>,
    value: SpreadsheetValue,
}

#[derive(Debug, Clone)]
struct CellState {
    successful: bool,
    value: SpreadsheetValue,
}

#[derive(Debug, Clone)]
struct EvaluationFailure {
    code: &'static str,
    message: String,
}

impl EvaluationFailure {
    fn unsupported(message: impl Into<String>) -> Self {
        Self {
            code: "office.kernel.spreadsheet.formula_unsupported",
            message: message.into(),
        }
    }

    fn dependency_unresolved() -> Self {
        Self {
            code: "office.kernel.spreadsheet.dependency_unresolved",
            message: "A formula dependency requires compatibility calculation.".into(),
        }
    }

    fn dependency_depth_exceeded() -> Self {
        Self {
            code: "office.kernel.spreadsheet.dependency_depth_exceeded",
            message: format!(
                "Spreadsheet formula dependency depth may not exceed {MAX_SPREADSHEET_DEPENDENCY_DEPTH}."
            ),
        }
    }
}

struct SpreadsheetEvaluator<'request> {
    request: &'request SpreadsheetCalculationRequest,
    cells: BTreeMap<CellKey, IndexedCell>,
    sheet_ids: BTreeMap<String, usize>,
    sheet_names: BTreeMap<String, usize>,
    states: BTreeMap<CellKey, CellState>,
    stack: Vec<CellKey>,
    materialized_range_cells: Vec<usize>,
    calculation_order: Vec<CellKey>,
    issues: Vec<SpreadsheetCalculationIssue>,
}

pub fn calculate_spreadsheet(
    request: &SpreadsheetCalculationRequest,
) -> Result<SpreadsheetCalculationResult, KernelError> {
    validate_request(request)?;
    Ok(SpreadsheetEvaluator::new(request).calculate())
}

impl<'request> SpreadsheetEvaluator<'request> {
    fn new(request: &'request SpreadsheetCalculationRequest) -> Self {
        let mut cells = BTreeMap::new();
        let mut sheet_ids = BTreeMap::new();
        let mut sheet_names = BTreeMap::new();
        for (sheet_index, sheet) in request.sheets.iter().enumerate() {
            sheet_ids.insert(sheet.id.clone(), sheet_index);
            sheet_names.insert(sheet.name.to_lowercase(), sheet_index);
            for cell in &sheet.cells {
                cells.insert(
                    CellKey {
                        sheet: sheet_index,
                        row: cell.row,
                        column: cell.column,
                    },
                    IndexedCell {
                        formula: cell.formula.clone(),
                        value: cell.value.clone(),
                    },
                );
            }
        }
        Self {
            request,
            cells,
            sheet_ids,
            sheet_names,
            states: BTreeMap::new(),
            stack: Vec::new(),
            materialized_range_cells: Vec::new(),
            calculation_order: Vec::new(),
            issues: Vec::new(),
        }
    }

    fn calculate(mut self) -> SpreadsheetCalculationResult {
        let targets = if self.request.targets.is_empty() {
            self.cells
                .iter()
                .filter_map(|(key, cell)| cell.formula.as_ref().map(|_| key.clone()))
                .collect::<Vec<_>>()
        } else {
            self.request
                .targets
                .iter()
                .filter_map(|coordinate| self.key_for_coordinate(coordinate))
                .collect()
        };
        for target in targets {
            self.evaluate_cell(&target);
        }
        let calculation_order = self
            .calculation_order
            .iter()
            .filter(|key| self.states.get(*key).is_some_and(|state| state.successful))
            .map(|key| self.coordinate(key))
            .collect::<Vec<_>>();
        let cells = calculation_order
            .iter()
            .filter_map(|coordinate| {
                let key = self.key_for_coordinate(coordinate)?;
                let state = self.states.get(&key)?;
                Some(SpreadsheetCalculatedCell {
                    sheet_id: coordinate.sheet_id.clone(),
                    row: coordinate.row,
                    column: coordinate.column,
                    value: state.value.clone(),
                })
            })
            .collect();
        SpreadsheetCalculationResult {
            protocol: OFFICE_KERNEL_PROTOCOL_VERSION,
            kind: "spreadsheetCalculationResult".into(),
            request_id: self.request.request_id,
            revision: self.request.revision,
            document_revision: self.request.document_revision,
            engine: "wasm".into(),
            cells,
            calculation_order,
            issues: self.issues,
        }
    }

    fn evaluate_cell(&mut self, key: &CellKey) -> SpreadsheetValue {
        if let Some(state) = self.states.get(key) {
            return state.value.clone();
        }
        let Some(cell) = self.cells.get(key).cloned() else {
            return SpreadsheetValue::Blank;
        };
        let Some(formula) = cell.formula.as_deref() else {
            return cell.value;
        };
        if let Some(position) = self.stack.iter().position(|candidate| candidate == key) {
            self.mark_cycle(position);
            return cell.value;
        }
        if self.stack.len() >= MAX_SPREADSHEET_DEPENDENCY_DEPTH {
            return self.record_state(
                key,
                cell.value,
                false,
                Some(EvaluationFailure::dependency_depth_exceeded()),
            );
        }

        self.stack.push(key.clone());
        self.materialized_range_cells.push(0);
        let evaluation = parse_spreadsheet_formula(formula)
            .map_err(|error| EvaluationFailure {
                code: "office.kernel.spreadsheet.formula_invalid",
                message: error.to_string(),
            })
            .and_then(|formula| self.evaluate_expression(&formula.root, key));
        self.materialized_range_cells.pop();
        self.stack.pop();
        if let Some(state) = self.states.get(key) {
            return state.value.clone();
        }
        match evaluation {
            Ok(value) => self.record_state(key, value, true, None),
            Err(failure) => self.record_state(key, cell.value, false, Some(failure)),
        }
    }

    fn evaluate_dependency(
        &mut self,
        key: &CellKey,
    ) -> Result<SpreadsheetValue, EvaluationFailure> {
        let value = self.evaluate_cell(key);
        if self.states.get(key).is_some_and(|state| !state.successful) {
            return Err(EvaluationFailure::dependency_unresolved());
        }
        Ok(value)
    }

    fn mark_cycle(&mut self, start: usize) {
        let cycle = self.stack[start..].to_vec();
        for key in cycle {
            let cached = self
                .cells
                .get(&key)
                .map_or(SpreadsheetValue::Blank, |cell| cell.value.clone());
            self.record_state(
                &key,
                cached,
                false,
                Some(EvaluationFailure {
                    code: "office.kernel.spreadsheet.circular_reference",
                    message: "Formula dependency cycle is not supported by this calculation pass."
                        .into(),
                }),
            );
        }
    }

    fn record_state(
        &mut self,
        key: &CellKey,
        value: SpreadsheetValue,
        successful: bool,
        failure: Option<EvaluationFailure>,
    ) -> SpreadsheetValue {
        if let Some(state) = self.states.get(key) {
            return state.value.clone();
        }
        self.states.insert(
            key.clone(),
            CellState {
                successful,
                value: value.clone(),
            },
        );
        self.calculation_order.push(key.clone());
        if let Some(failure) = failure {
            self.issues.push(SpreadsheetCalculationIssue {
                cell: self.coordinate(key),
                code: failure.code.into(),
                message: failure.message,
            });
        }
        value
    }

    fn coordinate(&self, key: &CellKey) -> SpreadsheetCoordinate {
        SpreadsheetCoordinate {
            sheet_id: self.request.sheets[key.sheet].id.clone(),
            row: key.row,
            column: key.column,
        }
    }

    fn key_for_coordinate(&self, coordinate: &SpreadsheetCoordinate) -> Option<CellKey> {
        Some(CellKey {
            sheet: *self.sheet_ids.get(&coordinate.sheet_id)?,
            row: coordinate.row,
            column: coordinate.column,
        })
    }
}

fn validate_request(request: &SpreadsheetCalculationRequest) -> Result<(), KernelError> {
    if request.protocol != OFFICE_KERNEL_PROTOCOL_VERSION {
        return Err(KernelError::invalid(
            "office.kernel.protocol_unsupported",
            format!(
                "Office kernel protocol {} is unsupported; expected {}.",
                request.protocol, OFFICE_KERNEL_PROTOCOL_VERSION
            ),
        ));
    }
    if request.kind != "spreadsheetCalculation" {
        return Err(KernelError::invalid(
            "office.kernel.request_kind_invalid",
            "The Spreadsheet kernel only accepts calculation requests.",
        ));
    }
    if request.sheets.len() > MAX_SPREADSHEET_SHEETS {
        return Err(KernelError::invalid(
            "office.kernel.spreadsheet.sheet_limit_exceeded",
            format!(
                "A Spreadsheet calculation request may contain at most {MAX_SPREADSHEET_SHEETS} sheets."
            ),
        ));
    }
    if request.targets.len() > MAX_SPREADSHEET_CELLS {
        return Err(KernelError::invalid(
            "office.kernel.spreadsheet.target_limit_exceeded",
            format!(
                "A Spreadsheet calculation request may contain at most {MAX_SPREADSHEET_CELLS} targets."
            ),
        ));
    }
    let mut ids = BTreeSet::new();
    let mut names = BTreeSet::new();
    let mut cell_count = 0_usize;
    for sheet in &request.sheets {
        validate_identifier("sheet ID", &sheet.id)?;
        validate_identifier("sheet name", &sheet.name)?;
        if !ids.insert(sheet.id.as_str()) || !names.insert(sheet.name.to_lowercase()) {
            return Err(KernelError::invalid(
                "office.kernel.spreadsheet.sheet_invalid",
                "Spreadsheet sheet IDs and names must be unique.",
            ));
        }
        let mut coordinates = BTreeSet::new();
        for cell in &sheet.cells {
            validate_coordinate(cell.row, cell.column)?;
            if !coordinates.insert((cell.row, cell.column)) {
                return Err(KernelError::invalid(
                    "office.kernel.spreadsheet.cell_invalid",
                    "Spreadsheet cells require unique row and column coordinates.",
                ));
            }
            cell.value.validate()?;
            if let Some(formula) = &cell.formula {
                let count = formula.chars().count();
                if formula.is_empty() || count > MAX_SPREADSHEET_FORMULA_CHARACTERS {
                    return Err(KernelError::invalid(
                        "office.kernel.spreadsheet.formula_invalid",
                        format!(
                            "Spreadsheet formulas must contain 1-{MAX_SPREADSHEET_FORMULA_CHARACTERS} characters."
                        ),
                    ));
                }
            }
        }
        cell_count = cell_count
            .checked_add(sheet.cells.len())
            .ok_or_else(cell_limit_error)?;
        if cell_count > MAX_SPREADSHEET_CELLS {
            return Err(cell_limit_error());
        }
    }
    for target in &request.targets {
        validate_coordinate(target.row, target.column)?;
        if !ids.contains(target.sheet_id.as_str()) {
            return Err(KernelError::invalid(
                "office.kernel.spreadsheet.target_invalid",
                format!(
                    "Spreadsheet calculation target references missing sheet '{}'.",
                    target.sheet_id
                ),
            ));
        }
    }
    Ok(())
}

fn validate_identifier(kind: &str, value: &str) -> Result<(), KernelError> {
    if value.trim().is_empty() || value.len() > MAX_SPREADSHEET_IDENTIFIER_BYTES {
        return Err(KernelError::invalid(
            "office.kernel.spreadsheet.sheet_invalid",
            format!(
                "Every Spreadsheet {kind} must contain 1-{MAX_SPREADSHEET_IDENTIFIER_BYTES} UTF-8 bytes."
            ),
        ));
    }
    Ok(())
}

fn validate_coordinate(row: u32, column: u32) -> Result<(), KernelError> {
    if row >= MAX_SPREADSHEET_ROWS || column >= MAX_SPREADSHEET_COLUMNS {
        return Err(KernelError::invalid(
            "office.kernel.spreadsheet.cell_invalid",
            "Spreadsheet cell coordinates must remain within XFD1048576.",
        ));
    }
    Ok(())
}

fn cell_limit_error() -> KernelError {
    KernelError::invalid(
        "office.kernel.spreadsheet.cell_limit_exceeded",
        format!(
            "A Spreadsheet calculation request may contain at most {MAX_SPREADSHEET_CELLS} populated cells."
        ),
    )
}
