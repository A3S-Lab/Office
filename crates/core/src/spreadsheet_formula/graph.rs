mod reference;

use std::collections::{BTreeMap, BTreeSet};

use a3s_use_core::{UseError, UseResult};
use serde::{Deserialize, Serialize};

use crate::discovery::office_error;
use crate::semantic::{DocumentNode, NativeOfficeDocument, OfficeNodeType};
use crate::spreadsheet_reference::CellReference;
use crate::DocumentKind;

use super::{parse_spreadsheet_formula, SpreadsheetFormula};
use crate::spreadsheet_formula::structured_reference::FormulaTableCatalog;
use reference::{
    collect_references, FormulaNamedDefinition, FormulaReferenceArea, FormulaReferenceCollection,
};

/// Maximum formula cells admitted to one native dependency graph.
pub const MAX_SPREADSHEET_FORMULA_CELLS: usize = 100_000;

/// Maximum formula-to-formula edges admitted to one dependency graph.
pub const MAX_SPREADSHEET_FORMULA_DEPENDENCIES: usize = 1_000_000;

/// Maximum formula-cell candidates visited while resolving all static
/// reference areas in one dependency graph.
pub const MAX_SPREADSHEET_FORMULA_REFERENCE_VISITS: usize = 1_000_000;

/// Stable workbook coordinate for a formula cell.
#[derive(Debug, Clone, PartialEq, Eq, PartialOrd, Ord, Hash, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct SpreadsheetFormulaCell {
    pub sheet: String,
    pub column: u32,
    pub row: u32,
}

impl SpreadsheetFormulaCell {
    pub fn path(&self) -> String {
        format!(
            "/{}/{}{}",
            self.sheet,
            crate::spreadsheet_reference::column_name(self.column),
            self.row
        )
    }
}

/// Reason a static dependency could not be resolved to workbook cells.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Serialize, Deserialize)]
#[serde(rename_all = "kebab-case")]
pub enum SpreadsheetFormulaUnresolvedReferenceKind {
    MissingWorksheet,
    ExternalWorkbook,
    UndefinedName,
    NamedRangeCycle,
    NamedRangeDepth,
    StructuredReference,
    DynamicReference,
    UnsupportedReference,
}

/// One source reference that cannot participate in the static graph.
#[derive(Debug, Clone, PartialEq, Eq, PartialOrd, Ord, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct SpreadsheetFormulaUnresolvedReference {
    pub kind: SpreadsheetFormulaUnresolvedReferenceKind,
    pub reference: String,
}

/// One formula cell and its formula-to-formula graph edges.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct SpreadsheetFormulaDependencyNode {
    pub cell: SpreadsheetFormulaCell,
    pub formula: String,
    pub dependencies: Vec<SpreadsheetFormulaCell>,
    pub dependents: Vec<SpreadsheetFormulaCell>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub unresolved_references: Vec<SpreadsheetFormulaUnresolvedReference>,
}

/// Bounded static dependency graph and deterministic calculation order.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct SpreadsheetFormulaDependencyGraph {
    pub nodes: Vec<SpreadsheetFormulaDependencyNode>,
    pub calculation_order: Vec<SpreadsheetFormulaCell>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub cycles: Vec<Vec<SpreadsheetFormulaCell>>,
}

impl SpreadsheetFormulaDependencyGraph {
    pub fn is_acyclic(&self) -> bool {
        self.cycles.is_empty()
    }
}

impl NativeOfficeDocument {
    /// Builds a bounded formula-to-formula dependency graph without evaluating
    /// any formula or fetching an external workbook.
    pub fn formula_dependency_graph(&self) -> UseResult<SpreadsheetFormulaDependencyGraph> {
        build_dependency_graph(self)
    }
}

struct FormulaRecord {
    cell: SpreadsheetFormulaCell,
    formula: String,
    parsed: SpreadsheetFormula,
    sheet_index: usize,
}

fn build_dependency_graph(
    document: &NativeOfficeDocument,
) -> UseResult<SpreadsheetFormulaDependencyGraph> {
    if document.kind() != DocumentKind::Spreadsheet {
        return Err(graph_error(
            "use.office.spreadsheet_formula_graph_type_unsupported",
            "Formula dependency graphs are available only for Spreadsheet documents.",
        ));
    }
    let sheet_names = document
        .root()
        .children
        .iter()
        .filter(|node| node.node_type == OfficeNodeType::Worksheet)
        .map(|node| node.path.trim_start_matches('/').to_string())
        .collect::<Vec<_>>();
    let tables = FormulaTableCatalog::collect(document.root(), &sheet_names)?;
    let named_definitions = collect_named_definitions(document.root(), &sheet_names);
    let records = collect_formula_records(document.root(), &sheet_names)?;
    let mut rows_by_sheet = vec![BTreeMap::<u32, BTreeMap<u32, usize>>::new(); sheet_names.len()];
    for (index, record) in records.iter().enumerate() {
        rows_by_sheet
            .get_mut(record.sheet_index)
            .ok_or_else(graph_index_error)?
            .entry(record.cell.row)
            .or_default()
            .insert(record.cell.column, index);
    }

    let mut dependencies = vec![BTreeSet::<usize>::new(); records.len()];
    let mut unresolved = Vec::with_capacity(records.len());
    let mut edge_count = 0_usize;
    let mut reference_visits = 0_usize;
    for (index, record) in records.iter().enumerate() {
        let FormulaReferenceCollection {
            areas,
            mut unresolved_references,
        } = collect_references(
            &record.parsed.root,
            record.sheet_index,
            record.cell.column,
            record.cell.row,
            &sheet_names,
            &named_definitions,
            &tables,
        )?;
        unresolved_references.sort();
        unresolved_references.dedup();
        unresolved.push(unresolved_references);
        for area in areas {
            for dependency in formula_cells_in_area(&area, &rows_by_sheet, &mut reference_visits)? {
                if dependencies
                    .get_mut(index)
                    .ok_or_else(graph_index_error)?
                    .insert(dependency)
                {
                    edge_count = edge_count.saturating_add(1);
                    if edge_count > MAX_SPREADSHEET_FORMULA_DEPENDENCIES {
                        return Err(graph_error(
                            "use.office.spreadsheet_formula_dependency_limit",
                            format!(
                                "Spreadsheet formula graph exceeds {MAX_SPREADSHEET_FORMULA_DEPENDENCIES} formula dependencies."
                            ),
                        )
                        .with_detail("dependencies", edge_count));
                    }
                }
            }
        }
    }

    let mut dependents = vec![BTreeSet::<usize>::new(); records.len()];
    for (cell, cell_dependencies) in dependencies.iter().enumerate() {
        for dependency in cell_dependencies {
            dependents
                .get_mut(*dependency)
                .ok_or_else(graph_index_error)?
                .insert(cell);
        }
    }
    let (calculation_order, remaining) = topological_order(&dependencies, &dependents)?;
    let cycles = strongly_connected_cycles(&dependencies, &dependents, &remaining)?;

    let mut nodes = Vec::with_capacity(records.len());
    for (index, record) in records.iter().enumerate() {
        let cell_dependencies = dependencies.get(index).ok_or_else(graph_index_error)?;
        let cell_dependents = dependents.get(index).ok_or_else(graph_index_error)?;
        let unresolved_references = unresolved.get(index).ok_or_else(graph_index_error)?;
        nodes.push(SpreadsheetFormulaDependencyNode {
            cell: record.cell.clone(),
            formula: record.formula.clone(),
            dependencies: cell_dependencies
                .iter()
                .map(|dependency| {
                    records
                        .get(*dependency)
                        .map(|record| record.cell.clone())
                        .ok_or_else(graph_index_error)
                })
                .collect::<UseResult<Vec<_>>>()?,
            dependents: cell_dependents
                .iter()
                .map(|dependent| {
                    records
                        .get(*dependent)
                        .map(|record| record.cell.clone())
                        .ok_or_else(graph_index_error)
                })
                .collect::<UseResult<Vec<_>>>()?,
            unresolved_references: unresolved_references.clone(),
        });
    }
    Ok(SpreadsheetFormulaDependencyGraph {
        nodes,
        calculation_order: calculation_order
            .into_iter()
            .map(|index| {
                records
                    .get(index)
                    .map(|record| record.cell.clone())
                    .ok_or_else(graph_index_error)
            })
            .collect::<UseResult<Vec<_>>>()?,
        cycles: cycles
            .into_iter()
            .map(|cycle| {
                cycle
                    .into_iter()
                    .map(|index| {
                        records
                            .get(index)
                            .map(|record| record.cell.clone())
                            .ok_or_else(graph_index_error)
                    })
                    .collect::<UseResult<Vec<_>>>()
            })
            .collect::<UseResult<Vec<_>>>()?,
    })
}

fn collect_formula_records(
    root: &DocumentNode,
    sheet_names: &[String],
) -> UseResult<Vec<FormulaRecord>> {
    let mut records = Vec::new();
    for (sheet_index, sheet_name) in sheet_names.iter().enumerate() {
        let sheet = root
            .children
            .iter()
            .find(|node| {
                node.node_type == OfficeNodeType::Worksheet
                    && node
                        .path
                        .strip_prefix('/')
                        .is_some_and(|path| path.eq_ignore_ascii_case(sheet_name))
            })
            .ok_or_else(|| {
                graph_error(
                    "use.office.spreadsheet_formula_graph_invalid",
                    format!("Formula graph cannot find worksheet '{sheet_name}'."),
                )
            })?;
        for row in &sheet.children {
            if row.node_type != OfficeNodeType::Row {
                continue;
            }
            for cell in &row.children {
                if cell.node_type != OfficeNodeType::Cell {
                    continue;
                }
                let Some(formula) = cell.format.get("formula") else {
                    continue;
                };
                if records.len() >= MAX_SPREADSHEET_FORMULA_CELLS {
                    return Err(graph_error(
                        "use.office.spreadsheet_formula_cell_limit",
                        format!(
                            "Spreadsheet formula graph accepts at most {MAX_SPREADSHEET_FORMULA_CELLS} formula cells."
                        ),
                    )
                    .with_detail("formulas", records.len().saturating_add(1)));
                }
                let reference = cell
                    .path
                    .rsplit_once('/')
                    .map(|(_, reference)| reference)
                    .ok_or_else(|| invalid_formula_cell(cell))?;
                let reference =
                    CellReference::parse(reference).map_err(|_| invalid_formula_cell(cell))?;
                let parsed = parse_spreadsheet_formula(formula)
                    .map_err(|error| formula_parse_error(error, &cell.path))?;
                records.push(FormulaRecord {
                    cell: SpreadsheetFormulaCell {
                        sheet: sheet_name.clone(),
                        column: reference.column,
                        row: reference.row,
                    },
                    formula: formula.clone(),
                    parsed,
                    sheet_index,
                });
            }
        }
    }
    Ok(records)
}

fn collect_named_definitions(
    root: &DocumentNode,
    sheet_names: &[String],
) -> BTreeMap<(Option<usize>, String), FormulaNamedDefinition> {
    let mut definitions = BTreeMap::new();
    for node in root
        .children
        .iter()
        .filter(|node| node.node_type == OfficeNodeType::NamedRangeCollection)
        .flat_map(|collection| &collection.children)
        .filter(|node| node.node_type == OfficeNodeType::NamedRange)
    {
        let (Some(name), Some(reference), Some(scope)) = (
            node.format.get("name"),
            node.format.get("ref"),
            node.format.get("scope"),
        ) else {
            continue;
        };
        let scope_sheet = if scope.eq_ignore_ascii_case("workbook") {
            None
        } else {
            let sheet = scope.strip_prefix("worksheet:").unwrap_or(scope);
            sheet_names
                .iter()
                .position(|candidate| candidate.eq_ignore_ascii_case(sheet))
        };
        definitions.insert(
            (scope_sheet, name.to_lowercase()),
            FormulaNamedDefinition {
                name: name.clone(),
                formula: reference.clone(),
                scope_sheet,
            },
        );
    }
    definitions
}

fn formula_cells_in_area(
    area: &FormulaReferenceArea,
    rows_by_sheet: &[BTreeMap<u32, BTreeMap<u32, usize>>],
    reference_visits: &mut usize,
) -> UseResult<Vec<usize>> {
    let Some(rows) = rows_by_sheet.get(area.sheet_index) else {
        return Ok(Vec::new());
    };
    let mut cells = Vec::new();
    for (_, columns) in rows.range(area.start_row..=area.end_row) {
        for (_, index) in columns.range(area.start_column..=area.end_column) {
            *reference_visits = reference_visits
                .checked_add(1)
                .ok_or_else(reference_visit_limit)?;
            if *reference_visits > MAX_SPREADSHEET_FORMULA_REFERENCE_VISITS {
                return Err(reference_visit_limit().with_detail("visits", *reference_visits));
            }
            cells.push(*index);
        }
    }
    Ok(cells)
}

fn topological_order(
    dependencies: &[BTreeSet<usize>],
    dependents: &[BTreeSet<usize>],
) -> UseResult<(Vec<usize>, BTreeSet<usize>)> {
    let mut indegree = dependencies.iter().map(BTreeSet::len).collect::<Vec<_>>();
    let mut ready = indegree
        .iter()
        .enumerate()
        .filter_map(|(index, count)| (*count == 0).then_some(index))
        .collect::<BTreeSet<_>>();
    let mut order = Vec::with_capacity(dependencies.len());
    while let Some(index) = ready.first().copied() {
        ready.remove(&index);
        order.push(index);
        for dependent in dependents.get(index).ok_or_else(graph_index_error)? {
            let degree = indegree.get_mut(*dependent).ok_or_else(graph_index_error)?;
            *degree = degree.saturating_sub(1);
            if *degree == 0 {
                ready.insert(*dependent);
            }
        }
    }
    let remaining = (0..dependencies.len())
        .filter(|index| indegree.get(*index).is_some_and(|degree| *degree > 0))
        .collect();
    Ok((order, remaining))
}

fn strongly_connected_cycles(
    dependencies: &[BTreeSet<usize>],
    dependents: &[BTreeSet<usize>],
    remaining: &BTreeSet<usize>,
) -> UseResult<Vec<Vec<usize>>> {
    let mut visited = BTreeSet::new();
    let mut finish = Vec::new();
    for start in remaining {
        if visited.contains(start) {
            continue;
        }
        let mut stack = vec![(*start, false)];
        while let Some((node, expanded)) = stack.pop() {
            if expanded {
                finish.push(node);
                continue;
            }
            if !visited.insert(node) {
                continue;
            }
            stack.push((node, true));
            for dependency in dependencies
                .get(node)
                .ok_or_else(graph_index_error)?
                .iter()
                .rev()
            {
                if remaining.contains(dependency) && !visited.contains(dependency) {
                    stack.push((*dependency, false));
                }
            }
        }
    }

    visited.clear();
    let mut cycles = Vec::new();
    for start in finish.into_iter().rev() {
        if visited.contains(&start) {
            continue;
        }
        let mut component = Vec::new();
        let mut stack = vec![start];
        visited.insert(start);
        while let Some(node) = stack.pop() {
            component.push(node);
            for dependent in dependents
                .get(node)
                .ok_or_else(graph_index_error)?
                .iter()
                .rev()
            {
                if remaining.contains(dependent) && visited.insert(*dependent) {
                    stack.push(*dependent);
                }
            }
        }
        component.sort_unstable();
        let first = component.first().copied().ok_or_else(graph_index_error)?;
        if component.len() > 1
            || dependencies
                .get(first)
                .ok_or_else(graph_index_error)?
                .contains(&first)
        {
            cycles.push(component);
        }
    }
    cycles.sort_by_key(|cycle| cycle.first().copied().unwrap_or(usize::MAX));
    Ok(cycles)
}

fn invalid_formula_cell(cell: &DocumentNode) -> UseError {
    graph_error(
        "use.office.spreadsheet_formula_graph_invalid",
        format!(
            "Formula cell '{}' has an invalid semantic coordinate.",
            cell.path
        ),
    )
    .with_detail("cell", cell.path.clone())
}

fn formula_parse_error(error: UseError, cell: &str) -> UseError {
    error.with_detail("cell", cell.to_string())
}

fn graph_error(code: &str, message: impl Into<String>) -> UseError {
    office_error(code, message)
}

fn graph_index_error() -> UseError {
    graph_error(
        "use.office.spreadsheet_formula_graph_invalid",
        "Spreadsheet formula graph contains an inconsistent internal index.",
    )
}

fn reference_visit_limit() -> UseError {
    graph_error(
        "use.office.spreadsheet_formula_reference_visit_limit",
        format!(
            "Spreadsheet formula graph visits at most {MAX_SPREADSHEET_FORMULA_REFERENCE_VISITS} formula-cell reference candidates."
        ),
    )
}
