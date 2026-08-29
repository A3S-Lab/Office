use std::collections::{BTreeMap, BTreeSet, VecDeque};

use serde::{Deserialize, Serialize};

use super::tables::SpreadsheetTableCatalog;
use super::{
    validate_spreadsheet_coordinate, validate_spreadsheet_input_cell, validate_spreadsheet_sheets,
    CellKey, EvaluationFailure, IndexedCell, SpreadsheetCalculatedCell,
    SpreadsheetCalculationIssue, SpreadsheetCoordinate, SpreadsheetEvaluationOutcome,
    SpreadsheetEvaluator, SpreadsheetInputCell, SpreadsheetInputSheet, SpreadsheetValue,
    MAX_SPREADSHEET_CELLS,
};
use crate::{KernelError, OFFICE_KERNEL_PROTOCOL_VERSION};

const MAX_SPREADSHEET_DEPENDENCY_EDGES: usize = 1_000_000;

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(
    tag = "kind",
    rename_all = "camelCase",
    rename_all_fields = "camelCase",
    deny_unknown_fields
)]
pub enum SpreadsheetCalculationSessionCellChange {
    Upsert {
        sheet_id: String,
        row: u32,
        column: u32,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        formula: Option<String>,
        value: SpreadsheetValue,
    },
    Remove {
        sheet_id: String,
        row: u32,
        column: u32,
    },
}

impl SpreadsheetCalculationSessionCellChange {
    fn coordinate(&self) -> (&str, u32, u32) {
        match self {
            Self::Upsert {
                sheet_id,
                row,
                column,
                ..
            }
            | Self::Remove {
                sheet_id,
                row,
                column,
            } => (sheet_id, *row, *column),
        }
    }
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(
    tag = "kind",
    rename_all = "camelCase",
    rename_all_fields = "camelCase",
    deny_unknown_fields
)]
pub enum SpreadsheetCalculationSessionUpdate {
    Replace {
        sheets: Vec<SpreadsheetInputSheet>,
    },
    Patch {
        base_document_revision: u64,
        changes: Vec<SpreadsheetCalculationSessionCellChange>,
    },
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(
    tag = "kind",
    rename_all = "camelCase",
    rename_all_fields = "camelCase",
    deny_unknown_fields
)]
pub enum SpreadsheetCalculationSessionScope {
    Workbook,
    Dirty,
    Targets { targets: Vec<SpreadsheetCoordinate> },
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct SpreadsheetCalculationSessionRequest {
    pub protocol: u32,
    pub kind: String,
    pub request_id: u32,
    pub revision: u32,
    pub document_revision: u64,
    pub update: SpreadsheetCalculationSessionUpdate,
    pub calculation: SpreadsheetCalculationSessionScope,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct SpreadsheetCalculationSessionStats {
    pub update_kind: String,
    pub calculation_scope: String,
    pub formula_cell_count: usize,
    pub dirty_formula_cell_count: usize,
    pub evaluated_formula_cell_count: usize,
    pub reused_formula_cell_count: usize,
    pub dependency_edge_count: usize,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct SpreadsheetCalculationSessionResult {
    pub protocol: u32,
    pub kind: String,
    pub request_id: u32,
    pub revision: u32,
    pub document_revision: u64,
    pub engine: String,
    pub cells: Vec<SpreadsheetCalculatedCell>,
    pub calculation_order: Vec<SpreadsheetCoordinate>,
    pub issues: Vec<SpreadsheetCalculationIssue>,
    pub stats: SpreadsheetCalculationSessionStats,
}

#[derive(Default)]
pub struct SpreadsheetCalculationSession {
    workbook: Option<SpreadsheetSessionWorkbook>,
}

pub fn calculate_spreadsheet_session(
    session: &mut SpreadsheetCalculationSession,
    request: &SpreadsheetCalculationSessionRequest,
) -> Result<SpreadsheetCalculationSessionResult, KernelError> {
    validate_session_request(request)?;
    let update_kind = match &request.update {
        SpreadsheetCalculationSessionUpdate::Replace { sheets } => {
            let workbook = SpreadsheetSessionWorkbook::new(request.document_revision, sheets)?;
            workbook.validate_calculation_scope(&request.calculation)?;
            session.workbook = Some(workbook);
            "replace"
        }
        SpreadsheetCalculationSessionUpdate::Patch {
            base_document_revision,
            changes,
        } => {
            let workbook = session.workbook.as_mut().ok_or_else(|| {
                KernelError::invalid(
                    "office.kernel.spreadsheet.session_uninitialized",
                    "A Spreadsheet calculation session requires a replace update before patches.",
                )
            })?;
            workbook.validate_calculation_scope(&request.calculation)?;
            workbook.apply_patch(*base_document_revision, request.document_revision, changes)?;
            "patch"
        }
    };
    let workbook = session.workbook.as_mut().ok_or_else(|| {
        KernelError::invalid(
            "office.kernel.spreadsheet.session_uninitialized",
            "The Spreadsheet calculation session is not initialized.",
        )
    })?;
    let calculation_scope = scope_name(&request.calculation);
    let outcome = workbook.calculate(&request.calculation)?;
    Ok(SpreadsheetCalculationSessionResult {
        protocol: OFFICE_KERNEL_PROTOCOL_VERSION,
        kind: "spreadsheetSessionCalculationResult".into(),
        request_id: request.request_id,
        revision: request.revision,
        document_revision: request.document_revision,
        engine: "wasm".into(),
        cells: outcome.cells,
        calculation_order: outcome.calculation_order,
        issues: outcome.issues,
        stats: SpreadsheetCalculationSessionStats {
            update_kind: update_kind.into(),
            calculation_scope: calculation_scope.into(),
            formula_cell_count: workbook.formula_cell_count(),
            dirty_formula_cell_count: outcome.dirty_formula_cell_count,
            evaluated_formula_cell_count: outcome.evaluated_formula_cell_count,
            reused_formula_cell_count: outcome.reused_formula_cell_count,
            dependency_edge_count: workbook.dependency_edge_count,
        },
    })
}

pub(super) struct SpreadsheetSessionWorkbook {
    pub(super) sheets: Vec<(String, String)>,
    pub(super) sheet_ids: BTreeMap<String, usize>,
    pub(super) sheet_names: BTreeMap<String, usize>,
    pub(super) cells: BTreeMap<CellKey, IndexedCell>,
    pub(super) dependencies: BTreeMap<CellKey, BTreeSet<CellKey>>,
    pub(super) dependents: BTreeMap<CellKey, BTreeSet<CellKey>>,
    pub(super) unresolved_formulas: BTreeSet<CellKey>,
    pub(super) table_catalog: SpreadsheetTableCatalog,
    pending_dirty: BTreeSet<CellKey>,
    dependency_edge_count: usize,
    document_revision: u64,
}

impl SpreadsheetSessionWorkbook {
    pub(super) fn new(
        document_revision: u64,
        sheets: &[SpreadsheetInputSheet],
    ) -> Result<Self, KernelError> {
        validate_spreadsheet_sheets(sheets)?;
        let table_catalog = SpreadsheetTableCatalog::from_sheets(sheets).map_err(|error| {
            KernelError::invalid("office.kernel.spreadsheet.table_invalid", error)
        })?;
        let mut sheet_metadata = Vec::with_capacity(sheets.len());
        let mut sheet_ids = BTreeMap::new();
        let mut sheet_names = BTreeMap::new();
        let mut cells = BTreeMap::new();
        for (sheet_index, sheet) in sheets.iter().enumerate() {
            sheet_metadata.push((sheet.id.clone(), sheet.name.clone()));
            sheet_ids.insert(sheet.id.clone(), sheet_index);
            sheet_names.insert(sheet.name.to_lowercase(), sheet_index);
            for cell in &sheet.cells {
                cells.insert(
                    CellKey {
                        sheet: sheet_index,
                        row: cell.row,
                        column: cell.column,
                    },
                    IndexedCell::new(cell),
                );
            }
        }
        let pending_dirty = cells
            .iter()
            .filter_map(|(key, cell)| cell.is_formula().then_some(key.clone()))
            .collect();
        Ok(Self {
            sheets: sheet_metadata,
            sheet_ids,
            sheet_names,
            cells,
            dependencies: BTreeMap::new(),
            dependents: BTreeMap::new(),
            unresolved_formulas: BTreeSet::new(),
            table_catalog,
            pending_dirty,
            dependency_edge_count: 0,
            document_revision,
        })
    }

    fn apply_patch(
        &mut self,
        base_document_revision: u64,
        document_revision: u64,
        changes: &[SpreadsheetCalculationSessionCellChange],
    ) -> Result<(), KernelError> {
        if base_document_revision != self.document_revision {
            return Err(KernelError::invalid(
                "office.kernel.spreadsheet.session_revision_mismatch",
                format!(
                    "Spreadsheet patch revision {base_document_revision} does not match session revision {}.",
                    self.document_revision
                ),
            ));
        }
        if document_revision < base_document_revision
            || (!changes.is_empty() && document_revision == base_document_revision)
        {
            return Err(KernelError::invalid(
                "office.kernel.spreadsheet.session_revision_invalid",
                "A Spreadsheet patch cannot move a session backwards, and a non-empty patch requires a newer document revision.",
            ));
        }
        if changes.len() > MAX_SPREADSHEET_CELLS {
            return Err(KernelError::invalid(
                "office.kernel.spreadsheet.patch_limit_exceeded",
                format!(
                    "A Spreadsheet session patch may contain at most {MAX_SPREADSHEET_CELLS} cell changes."
                ),
            ));
        }
        let mut changed_keys = BTreeSet::new();
        let mut resulting_cell_count = self.cells.len();
        for change in changes {
            let (sheet_id, row, column) = change.coordinate();
            validate_spreadsheet_coordinate(row, column)?;
            let sheet = *self.sheet_ids.get(sheet_id).ok_or_else(|| {
                KernelError::invalid(
                    "office.kernel.spreadsheet.patch_invalid",
                    format!("Spreadsheet patch references missing sheet '{sheet_id}'."),
                )
            })?;
            let key = CellKey { sheet, row, column };
            if !changed_keys.insert(key.clone()) {
                return Err(KernelError::invalid(
                    "office.kernel.spreadsheet.patch_invalid",
                    "Spreadsheet session patches require unique cell coordinates.",
                ));
            }
            match change {
                SpreadsheetCalculationSessionCellChange::Upsert { formula, value, .. } => {
                    validate_spreadsheet_input_cell(&SpreadsheetInputCell {
                        row,
                        column,
                        formula: formula.clone(),
                        value: value.clone(),
                    })?;
                    if !self.cells.contains_key(&key) {
                        resulting_cell_count = resulting_cell_count.saturating_add(1);
                    }
                }
                SpreadsheetCalculationSessionCellChange::Remove { .. } => {
                    if self.cells.contains_key(&key) {
                        resulting_cell_count = resulting_cell_count.saturating_sub(1);
                    }
                }
            }
        }
        if resulting_cell_count > MAX_SPREADSHEET_CELLS {
            return Err(KernelError::invalid(
                "office.kernel.spreadsheet.cell_limit_exceeded",
                format!(
                    "A Spreadsheet calculation session may contain at most {MAX_SPREADSHEET_CELLS} populated cells."
                ),
            ));
        }

        let mut dirty = self.pending_dirty.clone();
        dirty.extend(self.dependent_closure(&changed_keys));
        let unresolved = self.dependent_closure(&self.unresolved_formulas.clone());
        dirty.extend(unresolved);
        for change in changes {
            let (sheet_id, row, column) = change.coordinate();
            let sheet = *self
                .sheet_ids
                .get(sheet_id)
                .expect("validated sheet exists");
            let key = CellKey { sheet, row, column };
            if self.cells.get(&key).is_some_and(IndexedCell::is_formula) {
                self.clear_dependencies(&key);
            }
            self.unresolved_formulas.remove(&key);
            match change {
                SpreadsheetCalculationSessionCellChange::Upsert { formula, value, .. } => {
                    self.cells.insert(
                        key,
                        IndexedCell::new(&SpreadsheetInputCell {
                            row,
                            column,
                            formula: formula.clone(),
                            value: value.clone(),
                        }),
                    );
                }
                SpreadsheetCalculationSessionCellChange::Remove { .. } => {
                    self.cells.remove(&key);
                }
            }
        }
        self.pending_dirty = dirty
            .into_iter()
            .filter(|key| self.cells.get(key).is_some_and(IndexedCell::is_formula))
            .collect();
        self.document_revision = document_revision;
        Ok(())
    }

    fn calculate(
        &mut self,
        scope: &SpreadsheetCalculationSessionScope,
    ) -> Result<SpreadsheetEvaluationOutcome, KernelError> {
        let (targets, force_dependencies) = match scope {
            SpreadsheetCalculationSessionScope::Workbook => {
                (self.formula_keys().into_iter().collect::<Vec<_>>(), true)
            }
            SpreadsheetCalculationSessionScope::Dirty => (
                self.pending_dirty
                    .iter()
                    .filter(|key| self.cells.get(*key).is_some_and(IndexedCell::is_formula))
                    .cloned()
                    .collect(),
                false,
            ),
            SpreadsheetCalculationSessionScope::Targets { targets } => {
                let mut keys = BTreeSet::new();
                for target in targets {
                    validate_spreadsheet_coordinate(target.row, target.column)?;
                    let key = self.key_for_coordinate(target).ok_or_else(|| {
                        KernelError::invalid(
                            "office.kernel.spreadsheet.target_invalid",
                            format!(
                                "Spreadsheet calculation target references missing sheet '{}'.",
                                target.sheet_id
                            ),
                        )
                    })?;
                    if self.cells.get(&key).is_some_and(IndexedCell::is_formula) {
                        keys.insert(key);
                    }
                }
                (keys.into_iter().collect(), true)
            }
        };
        let dirty = targets.iter().cloned().collect();
        let outcome =
            SpreadsheetEvaluator::new(self, targets, dirty, force_dependencies).calculate();
        if matches!(scope, SpreadsheetCalculationSessionScope::Targets { .. }) {
            for coordinate in &outcome.calculation_order {
                if let Some(key) = self.key_for_coordinate(coordinate) {
                    self.pending_dirty.remove(&key);
                }
            }
        } else {
            self.pending_dirty.clear();
        }
        Ok(outcome)
    }

    fn validate_calculation_scope(
        &self,
        scope: &SpreadsheetCalculationSessionScope,
    ) -> Result<(), KernelError> {
        let SpreadsheetCalculationSessionScope::Targets { targets } = scope else {
            return Ok(());
        };
        for target in targets {
            validate_spreadsheet_coordinate(target.row, target.column)?;
            if !self.sheet_ids.contains_key(&target.sheet_id) {
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

    pub(super) fn coordinate(&self, key: &CellKey) -> SpreadsheetCoordinate {
        SpreadsheetCoordinate {
            sheet_id: self.sheets[key.sheet].0.clone(),
            row: key.row,
            column: key.column,
        }
    }

    pub(super) fn key_for_coordinate(&self, coordinate: &SpreadsheetCoordinate) -> Option<CellKey> {
        Some(CellKey {
            sheet: *self.sheet_ids.get(&coordinate.sheet_id)?,
            row: coordinate.row,
            column: coordinate.column,
        })
    }

    pub(super) fn clear_dependencies(&mut self, key: &CellKey) {
        let Some(dependencies) = self.dependencies.remove(key) else {
            return;
        };
        for dependency in dependencies {
            let remove_entry = self
                .dependents
                .get_mut(&dependency)
                .is_some_and(|dependents| {
                    let removed = dependents.remove(key);
                    if removed {
                        self.dependency_edge_count = self.dependency_edge_count.saturating_sub(1);
                    }
                    dependents.is_empty()
                });
            if remove_entry {
                self.dependents.remove(&dependency);
            }
        }
    }

    pub(super) fn record_dependency(
        &mut self,
        formula: &CellKey,
        dependency: &CellKey,
    ) -> Result<(), EvaluationFailure> {
        if self
            .dependencies
            .entry(formula.clone())
            .or_default()
            .contains(dependency)
        {
            return Ok(());
        }
        if self.dependency_edge_count >= MAX_SPREADSHEET_DEPENDENCY_EDGES {
            return Err(EvaluationFailure::dependency_limit_exceeded());
        }
        self.dependencies
            .entry(formula.clone())
            .or_default()
            .insert(dependency.clone());
        self.dependents
            .entry(dependency.clone())
            .or_default()
            .insert(formula.clone());
        self.dependency_edge_count += 1;
        Ok(())
    }

    fn formula_keys(&self) -> BTreeSet<CellKey> {
        self.cells
            .iter()
            .filter_map(|(key, cell)| cell.is_formula().then_some(key.clone()))
            .collect()
    }

    fn formula_cell_count(&self) -> usize {
        self.cells.values().filter(|cell| cell.is_formula()).count()
    }

    fn dependent_closure(&self, roots: &BTreeSet<CellKey>) -> BTreeSet<CellKey> {
        let mut closure = roots.clone();
        let mut queue = roots.iter().cloned().collect::<VecDeque<_>>();
        while let Some(key) = queue.pop_front() {
            for dependent in self.dependents.get(&key).into_iter().flatten() {
                if closure.insert(dependent.clone()) {
                    queue.push_back(dependent.clone());
                }
            }
        }
        closure
    }
}

fn validate_session_request(
    request: &SpreadsheetCalculationSessionRequest,
) -> Result<(), KernelError> {
    if request.protocol != OFFICE_KERNEL_PROTOCOL_VERSION {
        return Err(KernelError::invalid(
            "office.kernel.protocol_unsupported",
            format!(
                "Office kernel protocol {} is unsupported; expected {}.",
                request.protocol, OFFICE_KERNEL_PROTOCOL_VERSION
            ),
        ));
    }
    if request.kind != "spreadsheetSessionCalculation" {
        return Err(KernelError::invalid(
            "office.kernel.request_kind_invalid",
            "The Spreadsheet session kernel only accepts session calculation requests.",
        ));
    }
    match &request.update {
        SpreadsheetCalculationSessionUpdate::Replace { sheets } => {
            validate_spreadsheet_sheets(sheets)?;
        }
        SpreadsheetCalculationSessionUpdate::Patch { changes, .. } => {
            if changes.len() > MAX_SPREADSHEET_CELLS {
                return Err(KernelError::invalid(
                    "office.kernel.spreadsheet.patch_limit_exceeded",
                    format!(
                        "A Spreadsheet session patch may contain at most {MAX_SPREADSHEET_CELLS} cell changes."
                    ),
                ));
            }
        }
    }
    if let SpreadsheetCalculationSessionScope::Targets { targets } = &request.calculation {
        if targets.len() > MAX_SPREADSHEET_CELLS {
            return Err(KernelError::invalid(
                "office.kernel.spreadsheet.target_limit_exceeded",
                format!(
                    "A Spreadsheet calculation request may contain at most {MAX_SPREADSHEET_CELLS} targets."
                ),
            ));
        }
    }
    Ok(())
}

fn scope_name(scope: &SpreadsheetCalculationSessionScope) -> &'static str {
    match scope {
        SpreadsheetCalculationSessionScope::Workbook => "workbook",
        SpreadsheetCalculationSessionScope::Dirty => "dirty",
        SpreadsheetCalculationSessionScope::Targets { .. } => "targets",
    }
}
