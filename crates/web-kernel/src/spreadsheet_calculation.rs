use std::collections::{BTreeMap, BTreeSet};
use std::sync::Arc;

use a3s_office_formula_parser::{
    parse_spreadsheet_formula, SpreadsheetFormula, SpreadsheetFormulaErrorLiteral,
    MAX_SPREADSHEET_FORMULA_CHARACTERS,
};
use serde::{Deserialize, Serialize};

use crate::{KernelError, OFFICE_KERNEL_PROTOCOL_VERSION};

mod evaluate;
mod functions;
mod session;
mod tables;
mod value;

pub use session::{
    calculate_spreadsheet_session, SpreadsheetCalculationSession,
    SpreadsheetCalculationSessionCellChange, SpreadsheetCalculationSessionRequest,
    SpreadsheetCalculationSessionResult, SpreadsheetCalculationSessionScope,
    SpreadsheetCalculationSessionStats, SpreadsheetCalculationSessionUpdate,
};

const MAX_SPREADSHEET_SHEETS: usize = 1_024;
pub(super) const MAX_SPREADSHEET_CELLS: usize = 100_000;
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

/// Bounded table metadata used by the calculation kernel to resolve
/// structured references such as `Sales[Quantity]`.
///
/// Coordinates are zero-based and inclusive. Cell values and table styling
/// intentionally remain separate so the evaluator can stay sparse.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct SpreadsheetInputTable {
    pub name: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub display_name: Option<String>,
    pub start_row: u32,
    pub end_row: u32,
    pub start_column: u32,
    pub end_column: u32,
    pub columns: Vec<String>,
    pub header_row: bool,
    pub totals_row: bool,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct SpreadsheetInputSheet {
    pub id: String,
    pub name: String,
    pub cells: Vec<SpreadsheetInputCell>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub tables: Vec<SpreadsheetInputTable>,
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
pub(super) struct CellKey {
    pub(super) sheet: usize,
    pub(super) row: u32,
    pub(super) column: u32,
}

#[derive(Debug, Clone)]
pub(super) struct IndexedCell {
    formula: Option<String>,
    parsed_formula: Option<Result<Arc<SpreadsheetFormula>, EvaluationFailure>>,
    value: SpreadsheetValue,
}

impl IndexedCell {
    pub(super) fn new(cell: &SpreadsheetInputCell) -> Self {
        let parsed_formula = cell.formula.as_deref().map(|formula| {
            parse_spreadsheet_formula(formula)
                .map(Arc::new)
                .map_err(|error| EvaluationFailure {
                    code: "office.kernel.spreadsheet.formula_invalid",
                    message: error.to_string(),
                })
        });
        Self {
            formula: cell.formula.clone(),
            parsed_formula,
            value: cell.value.clone(),
        }
    }

    pub(super) fn is_formula(&self) -> bool {
        self.formula.is_some()
    }
}

#[derive(Debug, Clone)]
struct CellState {
    successful: bool,
    value: SpreadsheetValue,
}

#[derive(Debug, Clone)]
pub(super) struct EvaluationFailure {
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

    pub(super) fn dependency_limit_exceeded() -> Self {
        Self {
            code: "office.kernel.spreadsheet.dependency_limit_exceeded",
            message: "Spreadsheet formula dependency graph exceeds 1000000 edges.".into(),
        }
    }
}

pub(super) struct SpreadsheetEvaluator<'session> {
    session: &'session mut session::SpreadsheetSessionWorkbook,
    targets: Vec<CellKey>,
    dirty: BTreeSet<CellKey>,
    force_dependencies: bool,
    states: BTreeMap<CellKey, CellState>,
    stack: Vec<CellKey>,
    materialized_range_cells: Vec<usize>,
    calculation_order: Vec<CellKey>,
    issues: Vec<SpreadsheetCalculationIssue>,
    reused_formulas: BTreeSet<CellKey>,
}

pub(super) struct SpreadsheetEvaluationOutcome {
    cells: Vec<SpreadsheetCalculatedCell>,
    calculation_order: Vec<SpreadsheetCoordinate>,
    issues: Vec<SpreadsheetCalculationIssue>,
    dirty_formula_cell_count: usize,
    evaluated_formula_cell_count: usize,
    reused_formula_cell_count: usize,
}

pub fn calculate_spreadsheet(
    request: &SpreadsheetCalculationRequest,
) -> Result<SpreadsheetCalculationResult, KernelError> {
    validate_request(request)?;
    let mut session = SpreadsheetCalculationSession::default();
    let calculation = if request.targets.is_empty() {
        SpreadsheetCalculationSessionScope::Workbook
    } else {
        SpreadsheetCalculationSessionScope::Targets {
            targets: request.targets.clone(),
        }
    };
    let result = calculate_spreadsheet_session(
        &mut session,
        &SpreadsheetCalculationSessionRequest {
            protocol: request.protocol,
            kind: "spreadsheetSessionCalculation".into(),
            request_id: request.request_id,
            revision: request.revision,
            document_revision: request.document_revision,
            update: SpreadsheetCalculationSessionUpdate::Replace {
                sheets: request.sheets.clone(),
            },
            calculation,
        },
    )?;
    Ok(SpreadsheetCalculationResult {
        protocol: result.protocol,
        kind: "spreadsheetCalculationResult".into(),
        request_id: result.request_id,
        revision: result.revision,
        document_revision: result.document_revision,
        engine: result.engine,
        cells: result.cells,
        calculation_order: result.calculation_order,
        issues: result.issues,
    })
}

impl<'session> SpreadsheetEvaluator<'session> {
    fn new(
        session: &'session mut session::SpreadsheetSessionWorkbook,
        targets: Vec<CellKey>,
        dirty: BTreeSet<CellKey>,
        force_dependencies: bool,
    ) -> Self {
        Self {
            session,
            targets,
            dirty,
            force_dependencies,
            states: BTreeMap::new(),
            stack: Vec::new(),
            materialized_range_cells: Vec::new(),
            calculation_order: Vec::new(),
            issues: Vec::new(),
            reused_formulas: BTreeSet::new(),
        }
    }

    fn calculate(mut self) -> SpreadsheetEvaluationOutcome {
        let targets = self.targets.clone();
        for target in targets {
            self.evaluate_cell(&target);
        }
        let calculation_order = self
            .calculation_order
            .iter()
            .filter(|key| self.states.get(*key).is_some_and(|state| state.successful))
            .map(|key| self.session.coordinate(key))
            .collect::<Vec<_>>();
        let cells = calculation_order
            .iter()
            .filter_map(|coordinate| {
                let key = self.session.key_for_coordinate(coordinate)?;
                let state = self.states.get(&key)?;
                Some(SpreadsheetCalculatedCell {
                    sheet_id: coordinate.sheet_id.clone(),
                    row: coordinate.row,
                    column: coordinate.column,
                    value: state.value.clone(),
                })
            })
            .collect();
        SpreadsheetEvaluationOutcome {
            cells,
            calculation_order,
            issues: self.issues,
            dirty_formula_cell_count: self.dirty.len(),
            evaluated_formula_cell_count: self.states.len(),
            reused_formula_cell_count: self.reused_formulas.len(),
        }
    }

    fn evaluate_cell(&mut self, key: &CellKey) -> SpreadsheetValue {
        if let Some(state) = self.states.get(key) {
            return state.value.clone();
        }
        let Some(cell) = self.session.cells.get(key).cloned() else {
            return SpreadsheetValue::Blank;
        };
        let Some(parsed_formula) = cell.parsed_formula.clone() else {
            return cell.value;
        };
        if !self.force_dependencies && !self.dirty.contains(key) {
            self.reused_formulas.insert(key.clone());
            return cell.value;
        }
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

        self.session.clear_dependencies(key);
        self.stack.push(key.clone());
        self.materialized_range_cells.push(0);
        let evaluation =
            parsed_formula.and_then(|formula| self.evaluate_expression(&formula.root, key));
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
        if let Some(formula) = self.stack.last().cloned() {
            self.session.record_dependency(&formula, key)?;
        }
        let value = self.evaluate_cell(key);
        if self.states.get(key).is_some_and(|state| !state.successful)
            || self.session.unresolved_formulas.contains(key)
        {
            return Err(EvaluationFailure::dependency_unresolved());
        }
        Ok(value)
    }

    fn mark_cycle(&mut self, start: usize) {
        let cycle = self.stack[start..].to_vec();
        for key in cycle {
            let cached = self
                .session
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
        if successful {
            self.session.unresolved_formulas.remove(key);
            if let Some(cell) = self.session.cells.get_mut(key) {
                cell.value = value.clone();
            }
        } else {
            self.session.unresolved_formulas.insert(key.clone());
        }
        if let Some(failure) = failure {
            self.issues.push(SpreadsheetCalculationIssue {
                cell: self.session.coordinate(key),
                code: failure.code.into(),
                message: failure.message,
            });
        }
        value
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
    if request.targets.len() > MAX_SPREADSHEET_CELLS {
        return Err(KernelError::invalid(
            "office.kernel.spreadsheet.target_limit_exceeded",
            format!(
                "A Spreadsheet calculation request may contain at most {MAX_SPREADSHEET_CELLS} targets."
            ),
        ));
    }
    validate_spreadsheet_sheets(&request.sheets)?;
    let ids = request
        .sheets
        .iter()
        .map(|sheet| sheet.id.as_str())
        .collect::<BTreeSet<_>>();
    for target in &request.targets {
        validate_spreadsheet_coordinate(target.row, target.column)?;
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

pub(super) fn validate_spreadsheet_sheets(
    sheets: &[SpreadsheetInputSheet],
) -> Result<(), KernelError> {
    if sheets.len() > MAX_SPREADSHEET_SHEETS {
        return Err(KernelError::invalid(
            "office.kernel.spreadsheet.sheet_limit_exceeded",
            format!(
                "A Spreadsheet calculation request may contain at most {MAX_SPREADSHEET_SHEETS} sheets."
            ),
        ));
    }
    let mut ids = BTreeSet::new();
    let mut names = BTreeSet::new();
    let mut cell_count = 0_usize;
    for sheet in sheets {
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
            validate_spreadsheet_coordinate(cell.row, cell.column)?;
            if !coordinates.insert((cell.row, cell.column)) {
                return Err(KernelError::invalid(
                    "office.kernel.spreadsheet.cell_invalid",
                    "Spreadsheet cells require unique row and column coordinates.",
                ));
            }
            validate_spreadsheet_input_cell(cell)?;
        }
        cell_count = cell_count
            .checked_add(sheet.cells.len())
            .ok_or_else(cell_limit_error)?;
        if cell_count > MAX_SPREADSHEET_CELLS {
            return Err(cell_limit_error());
        }
    }
    tables::validate_tables(sheets)
        .map_err(|error| KernelError::invalid("office.kernel.spreadsheet.table_invalid", error))?;
    Ok(())
}

pub(super) fn validate_spreadsheet_input_cell(
    cell: &SpreadsheetInputCell,
) -> Result<(), KernelError> {
    validate_spreadsheet_coordinate(cell.row, cell.column)?;
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

pub(super) fn validate_spreadsheet_coordinate(row: u32, column: u32) -> Result<(), KernelError> {
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
