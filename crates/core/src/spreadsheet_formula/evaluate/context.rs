use std::collections::{BTreeMap, BTreeSet};

use a3s_use_core::{UseError, UseResult};

use crate::semantic::{DocumentNode, NativeOfficeDocument, OfficeNodeType};
use crate::spreadsheet_reference::{CellRange, CellReference};

use super::super::{
    parse_spreadsheet_formula, SpreadsheetFormula, SpreadsheetFormulaCell,
    SpreadsheetFormulaDependencyNode, SpreadsheetFormulaErrorLiteral,
    SpreadsheetFormulaFunctionRegistry,
};
use super::{calculation_error, spill_limit_error, MAX_SPREADSHEET_FORMULA_SPILL_CELLS};
use crate::spreadsheet_formula::structured_reference::FormulaTableCatalog;

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
pub(super) struct FormulaCellKey {
    pub(super) sheet: usize,
    pub(super) column: u32,
    pub(super) row: u32,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) struct FormulaReferenceArea {
    pub(super) sheet: usize,
    pub(super) start_column: u32,
    pub(super) start_row: u32,
    pub(super) end_column: u32,
    pub(super) end_row: u32,
}

impl FormulaReferenceArea {
    pub(super) fn cell_count(self) -> Option<usize> {
        let columns = u64::from(self.end_column - self.start_column + 1);
        let rows = u64::from(self.end_row - self.start_row + 1);
        usize::try_from(columns.checked_mul(rows)?).ok()
    }

    pub(super) fn contains(self, key: FormulaCellKey) -> bool {
        self.sheet == key.sheet
            && (self.start_column..=self.end_column).contains(&key.column)
            && (self.start_row..=self.end_row).contains(&key.row)
    }

    pub(super) fn intersect(self, other: Self) -> Option<Self> {
        if self.sheet != other.sheet {
            return None;
        }
        let start_column = self.start_column.max(other.start_column);
        let start_row = self.start_row.max(other.start_row);
        let end_column = self.end_column.min(other.end_column);
        let end_row = self.end_row.min(other.end_row);
        (start_column <= end_column && start_row <= end_row).then_some(Self {
            sheet: self.sheet,
            start_column,
            start_row,
            end_column,
            end_row,
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub(super) enum ScalarValue {
    Blank,
    Number(f64),
    Text(String),
    Boolean(bool),
    Error(SpreadsheetFormulaErrorLiteral),
}

#[derive(Debug, Clone, PartialEq)]
pub(super) struct FormulaArray {
    pub(super) rows: Vec<Vec<ScalarValue>>,
}

impl FormulaArray {
    pub(super) fn new(rows: Vec<Vec<ScalarValue>>) -> Option<Self> {
        let width = rows.first()?.len();
        (width > 0 && rows.iter().all(|row| row.len() == width)).then_some(Self { rows })
    }

    pub(super) fn height(&self) -> usize {
        self.rows.len()
    }

    pub(super) fn width(&self) -> usize {
        self.rows.first().map_or(0, Vec::len)
    }

    pub(super) fn scalar(value: ScalarValue) -> Self {
        Self {
            rows: vec![vec![value]],
        }
    }

    pub(super) fn broadcast_value(&self, row: usize, column: usize) -> Option<&ScalarValue> {
        let row = if self.height() == 1 { 0 } else { row };
        let column = if self.width() == 1 { 0 } else { column };
        self.rows.get(row)?.get(column)
    }
}

#[derive(Debug, Clone, PartialEq)]
pub(super) enum EvalValue {
    Scalar(ScalarValue),
    Array(FormulaArray),
    Reference(Vec<FormulaReferenceArea>),
}

#[derive(Debug, Clone)]
pub(super) struct NamedFormulaDefinition {
    pub(super) name: String,
    pub(super) formula: String,
    pub(super) scope_sheet: Option<usize>,
}

pub(super) struct FormulaRecord {
    pub(super) key: FormulaCellKey,
    pub(super) formula: SpreadsheetFormula,
}

pub(super) struct EvaluationContext<'a> {
    pub(super) sheet_names: Vec<String>,
    pub(super) registry: &'a SpreadsheetFormulaFunctionRegistry,
    pub(super) values: BTreeMap<FormulaCellKey, ScalarValue>,
    pub(super) occupied: BTreeSet<FormulaCellKey>,
    pub(super) formula_cells: BTreeSet<FormulaCellKey>,
    pub(super) old_spills: BTreeMap<FormulaCellKey, FormulaReferenceArea>,
    pub(super) spills: BTreeMap<FormulaCellKey, FormulaReferenceArea>,
    pub(super) spill_owners: BTreeMap<FormulaCellKey, FormulaCellKey>,
    pub(super) named_definitions: BTreeMap<(Option<usize>, String), NamedFormulaDefinition>,
    pub(super) named_stack: BTreeSet<(Option<usize>, String)>,
    pub(super) tables: FormulaTableCatalog,
}

pub(super) fn build_context<'a>(
    document: &NativeOfficeDocument,
    registry: &'a SpreadsheetFormulaFunctionRegistry,
    graph_nodes: &[SpreadsheetFormulaDependencyNode],
) -> UseResult<(EvaluationContext<'a>, Vec<FormulaRecord>)> {
    let sheet_names = document
        .root()
        .children
        .iter()
        .filter(|node| node.node_type == OfficeNodeType::Worksheet)
        .map(|node| node.path.trim_start_matches('/').to_string())
        .collect::<Vec<_>>();
    let mut values = BTreeMap::new();
    let mut occupied = BTreeSet::new();
    let mut formula_cells = BTreeSet::new();
    let mut old_spills = BTreeMap::new();
    let mut old_spill_cell_count = 0_usize;
    for (sheet, sheet_name) in sheet_names.iter().enumerate() {
        let Some(sheet_node) = document.root().children.iter().find(|node| {
            node.node_type == OfficeNodeType::Worksheet
                && node
                    .path
                    .strip_prefix('/')
                    .is_some_and(|path| path.eq_ignore_ascii_case(sheet_name))
        }) else {
            continue;
        };
        for row in &sheet_node.children {
            if row.node_type != OfficeNodeType::Row {
                continue;
            }
            for cell in &row.children {
                if cell.node_type != OfficeNodeType::Cell {
                    continue;
                }
                let reference = cell
                    .path
                    .rsplit_once('/')
                    .and_then(|(_, value)| CellReference::parse(value).ok())
                    .ok_or_else(|| invalid_semantic_cell(cell))?;
                let key = FormulaCellKey {
                    sheet,
                    column: reference.column,
                    row: reference.row,
                };
                let is_formula = cell.format.contains_key("formula");
                let has_value = cell.format.get("valuePresent").map(String::as_str) == Some("true");
                if is_formula || has_value {
                    occupied.insert(key);
                }
                if is_formula {
                    formula_cells.insert(key);
                    if let Some(area) = stored_formula_spill(cell, key)? {
                        let spill_cells = area
                            .cell_count()
                            .ok_or_else(spill_limit_error)?
                            .saturating_sub(1);
                        old_spill_cell_count = old_spill_cell_count
                            .checked_add(spill_cells)
                            .ok_or_else(spill_limit_error)?;
                        if old_spill_cell_count > MAX_SPREADSHEET_FORMULA_SPILL_CELLS {
                            return Err(
                                spill_limit_error().with_detail("cells", old_spill_cell_count)
                            );
                        }
                        old_spills.insert(key, area);
                    }
                } else {
                    values.insert(key, semantic_scalar(cell));
                }
            }
        }
    }
    for (anchor, area) in &old_spills {
        for row in area.start_row..=area.end_row {
            for column in area.start_column..=area.end_column {
                let key = FormulaCellKey {
                    sheet: area.sheet,
                    column,
                    row,
                };
                if key != *anchor {
                    values.remove(&key);
                }
            }
        }
    }
    let tables = FormulaTableCatalog::collect(document.root(), &sheet_names)?;
    let named_definitions = collect_named_definitions(document.root(), &sheet_names);
    let records = graph_nodes
        .iter()
        .map(|node| {
            let key = public_cell_key(&node.cell, &sheet_names)?;
            Ok(FormulaRecord {
                key,
                formula: parse_spreadsheet_formula(&node.formula)
                    .map_err(|error| error.with_detail("cell", node.cell.path()))?,
            })
        })
        .collect::<UseResult<Vec<_>>>()?;
    Ok((
        EvaluationContext {
            sheet_names,
            registry,
            values,
            occupied,
            formula_cells,
            old_spills,
            spills: BTreeMap::new(),
            spill_owners: BTreeMap::new(),
            named_definitions,
            named_stack: BTreeSet::new(),
            tables,
        },
        records,
    ))
}

pub(super) fn public_cell_key(
    cell: &SpreadsheetFormulaCell,
    sheet_names: &[String],
) -> UseResult<FormulaCellKey> {
    let sheet = sheet_names
        .iter()
        .position(|sheet| sheet.eq_ignore_ascii_case(&cell.sheet))
        .ok_or_else(|| {
            calculation_error(
                "use.office.spreadsheet_formula_calculation_invalid",
                format!("Calculation references missing worksheet '{}'.", cell.sheet),
            )
        })?;
    Ok(FormulaCellKey {
        sheet,
        column: cell.column,
        row: cell.row,
    })
}

fn semantic_scalar(cell: &DocumentNode) -> ScalarValue {
    match cell.format.get("valueType").map(String::as_str) {
        Some("String" | "Date") => ScalarValue::Text(cell.text.clone()),
        Some("Boolean") => ScalarValue::Boolean(cell.text.eq_ignore_ascii_case("true")),
        Some("Error") => SpreadsheetFormulaErrorLiteral::parse(&cell.text).map_or(
            ScalarValue::Error(SpreadsheetFormulaErrorLiteral::Value),
            ScalarValue::Error,
        ),
        _ if cell.text.is_empty() => ScalarValue::Blank,
        _ => cell
            .text
            .parse::<f64>()
            .ok()
            .filter(|value| value.is_finite())
            .map_or(
                ScalarValue::Error(SpreadsheetFormulaErrorLiteral::Value),
                ScalarValue::Number,
            ),
    }
}

fn collect_named_definitions(
    root: &DocumentNode,
    sheet_names: &[String],
) -> BTreeMap<(Option<usize>, String), NamedFormulaDefinition> {
    root.children
        .iter()
        .filter(|node| node.node_type == OfficeNodeType::NamedRangeCollection)
        .flat_map(|collection| &collection.children)
        .filter_map(|node| {
            let name = node.format.get("name")?.clone();
            let formula = node.format.get("ref")?.clone();
            let scope = node.format.get("scope")?;
            let scope_sheet = if scope.eq_ignore_ascii_case("workbook") {
                None
            } else {
                let sheet = scope.strip_prefix("worksheet:").unwrap_or(scope);
                sheet_names
                    .iter()
                    .position(|candidate| candidate.eq_ignore_ascii_case(sheet))
            };
            Some((
                (scope_sheet, name.to_lowercase()),
                NamedFormulaDefinition {
                    name,
                    formula,
                    scope_sheet,
                },
            ))
        })
        .collect()
}

fn invalid_semantic_cell(cell: &DocumentNode) -> UseError {
    calculation_error(
        "use.office.spreadsheet_formula_calculation_invalid",
        format!(
            "Spreadsheet cell '{}' has an invalid coordinate.",
            cell.path
        ),
    )
    .with_detail("cell", cell.path.clone())
}

fn stored_formula_spill(
    cell: &DocumentNode,
    anchor: FormulaCellKey,
) -> UseResult<Option<FormulaReferenceArea>> {
    let formula_type = cell.format.get("formulaType").map(String::as_str);
    let formula_reference = cell.format.get("formulaRef");
    match (formula_type, formula_reference) {
        (None, None) => Ok(None),
        (Some(kind), None) if kind.eq_ignore_ascii_case("normal") => Ok(None),
        (Some(kind), Some(reference)) if kind.eq_ignore_ascii_case("array") => {
            let range = CellRange::parse(reference).map_err(|error| {
                calculation_error(
                    "use.office.spreadsheet_formula_storage_invalid",
                    format!(
                        "Formula cell '{}' has invalid array range '{reference}': {error}",
                        cell.path
                    ),
                )
            })?;
            let anchor_reference = CellReference {
                column: anchor.column,
                row: anchor.row,
            };
            if !range.contains(anchor_reference) {
                return Err(calculation_error(
                    "use.office.spreadsheet_formula_storage_invalid",
                    format!(
                        "Formula array range '{}' does not contain anchor '{}'.",
                        range.a1(),
                        cell.path
                    ),
                ));
            }
            Ok(Some(FormulaReferenceArea {
                sheet: anchor.sheet,
                start_column: range.start.column,
                start_row: range.start.row,
                end_column: range.end.column,
                end_row: range.end.row,
            }))
        }
        (Some(kind), None) if kind.eq_ignore_ascii_case("array") => Err(calculation_error(
            "use.office.spreadsheet_formula_storage_invalid",
            format!("Array formula cell '{}' has no array range.", cell.path),
        )),
        (None, Some(_)) => Err(calculation_error(
            "use.office.spreadsheet_formula_storage_invalid",
            format!(
                "Formula cell '{}' has an array range without array storage type.",
                cell.path
            ),
        )),
        (Some(kind), Some(_)) if kind.eq_ignore_ascii_case("normal") => Err(calculation_error(
            "use.office.spreadsheet_formula_storage_invalid",
            format!(
                "Normal formula cell '{}' cannot own an array range.",
                cell.path
            ),
        )),
        (Some(kind), _) => Err(calculation_error(
            "use.office.spreadsheet_formula_storage_unsupported",
            format!(
                "Formula storage type '{kind}' at '{}' is not supported by native calculation.",
                cell.path
            ),
        )),
    }
}
