mod parser;
mod rewrite;

use std::collections::BTreeMap;

use a3s_use_core::{UseError, UseResult};

use crate::discovery::office_error;
use crate::semantic::{DocumentNode, OfficeNodeType};
use crate::spreadsheet_reference::{CellRange, CellReference};

use super::SpreadsheetFormulaQualifier;
use parser::{parse_reference, ParsedStructuredReference, StructuredRowSelection};
pub(crate) use rewrite::{
    LocalStructuredReferenceContext, StructuredReferenceRewritePlan,
    StructuredReferenceRewriteResult,
};

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum StructuredReferenceErrorKind {
    ExternalWorkbook,
    Unsupported,
    MissingTable,
    MissingColumn,
}

#[derive(Debug, Clone)]
pub(crate) struct StructuredReferenceError {
    pub(crate) kind: StructuredReferenceErrorKind,
    pub(crate) message: String,
}

#[derive(Debug, Clone, Copy)]
pub(crate) struct ResolvedStructuredReference {
    pub(crate) sheet: usize,
    pub(crate) start_column: u32,
    pub(crate) start_row: u32,
    pub(crate) end_column: u32,
    pub(crate) end_row: u32,
}

#[derive(Debug, Clone)]
struct FormulaTableDefinition {
    path: String,
    name: String,
    sheet: usize,
    range: CellRange,
    header_row: bool,
    totals_row: bool,
    columns: Vec<String>,
}

#[derive(Debug, Clone)]
pub(crate) struct FormulaTableCatalog {
    sheet_names: Vec<String>,
    definitions: Vec<FormulaTableDefinition>,
    by_name: BTreeMap<String, usize>,
    by_sheet: Vec<Vec<usize>>,
}

impl FormulaTableCatalog {
    pub(crate) fn collect(root: &DocumentNode, sheet_names: &[String]) -> UseResult<Self> {
        let mut definitions = Vec::<FormulaTableDefinition>::new();
        let mut by_name = BTreeMap::<String, usize>::new();
        let mut by_sheet = vec![Vec::<usize>::new(); sheet_names.len()];
        for (sheet, sheet_name) in sheet_names.iter().enumerate() {
            let worksheet = root
                .children
                .iter()
                .find(|node| {
                    node.node_type == OfficeNodeType::Worksheet
                        && node
                            .path
                            .strip_prefix('/')
                            .is_some_and(|path| path.eq_ignore_ascii_case(sheet_name))
                })
                .ok_or_else(|| table_catalog_error(format!("Missing worksheet '{sheet_name}'.")))?;
            for table in worksheet
                .children
                .iter()
                .filter(|node| node.node_type == OfficeNodeType::Table)
            {
                let name = required_table_format(table, "name")?;
                let display_name = required_table_format(table, "displayName")?;
                let range_text = required_table_format(table, "ref")?;
                let range = CellRange::parse(range_text).map_err(|error| {
                    table_catalog_error(format!(
                        "Spreadsheet table '{}' has invalid range '{range_text}': {error}",
                        table.path
                    ))
                })?;
                let header_row = table_boolean(table, "headerRow")?;
                let totals_row = table_boolean(table, "totalsRow")?;
                let columns = table
                    .children
                    .iter()
                    .filter(|node| node.node_type == OfficeNodeType::TableColumn)
                    .map(|column| {
                        column.format.get("name").cloned().ok_or_else(|| {
                            table_catalog_error(format!(
                                "Spreadsheet table column '{}' has no name.",
                                column.path
                            ))
                        })
                    })
                    .collect::<UseResult<Vec<_>>>()?;
                let width = usize::try_from(range.end.column - range.start.column + 1)
                    .map_err(|_| table_catalog_error("Spreadsheet table width is invalid."))?;
                if columns.len() != width {
                    return Err(table_catalog_error(format!(
                        "Spreadsheet table '{}' has {} columns for range '{}'.",
                        table.path,
                        columns.len(),
                        range.a1()
                    )));
                }
                let definition = FormulaTableDefinition {
                    path: table.path.clone(),
                    name: name.to_string(),
                    sheet,
                    range,
                    header_row,
                    totals_row,
                    columns,
                };
                let aliases = [name, display_name]
                    .into_iter()
                    .map(|alias| (alias.to_string(), alias.to_lowercase()))
                    .collect::<Vec<_>>();
                for (alias, normalized) in &aliases {
                    if let Some(existing) = by_name.get(normalized) {
                        let existing = definitions.get(*existing).ok_or_else(|| {
                            table_catalog_error("Spreadsheet table name index is invalid.")
                        })?;
                        if existing.path != definition.path {
                            return Err(table_catalog_error(format!(
                                "Spreadsheet table formula name '{alias}' is ambiguous."
                            )));
                        }
                    }
                }
                let definition_index = definitions.len();
                definitions.push(definition);
                by_sheet
                    .get_mut(sheet)
                    .ok_or_else(|| {
                        table_catalog_error("Spreadsheet table worksheet index is invalid.")
                    })?
                    .push(definition_index);
                for (_, normalized) in aliases {
                    by_name.entry(normalized).or_insert(definition_index);
                }
            }
        }
        Ok(Self {
            sheet_names: sheet_names.to_vec(),
            definitions,
            by_name,
            by_sheet,
        })
    }

    pub(crate) fn resolve(
        &self,
        qualifier: Option<&SpreadsheetFormulaQualifier>,
        reference: &str,
        current_sheet: usize,
        current_column: u32,
        current_row: u32,
    ) -> Result<Vec<ResolvedStructuredReference>, StructuredReferenceError> {
        let parsed = parse_reference(reference)?;
        let qualifier_sheet = self.resolve_qualifier(qualifier)?;
        let table = self.resolve_table(
            &parsed,
            reference,
            current_sheet,
            current_column,
            current_row,
        )?;
        if let Some(sheet) = qualifier_sheet {
            if sheet != table.sheet {
                return Err(structured_error(
                    StructuredReferenceErrorKind::MissingTable,
                    format!(
                        "Spreadsheet table '{}' is not on worksheet '{}'.",
                        table.name,
                        qualifier
                            .map(|value| value.worksheet.as_str())
                            .unwrap_or_default()
                    ),
                ));
            }
        }
        if parsed.rows.current && current_sheet != table.sheet {
            return Err(structured_error(
                StructuredReferenceErrorKind::Unsupported,
                format!(
                    "Structured reference #This Row requires the current formula cell to be on table '{}'.",
                    table.name
                ),
            ));
        }
        let (start_column, end_column) = resolve_columns(table, &parsed)?;
        resolve_rows(table, parsed.rows, current_row)?
            .into_iter()
            .map(|(start_row, end_row)| {
                Ok(ResolvedStructuredReference {
                    sheet: table.sheet,
                    start_column,
                    start_row,
                    end_column,
                    end_row,
                })
            })
            .collect()
    }

    fn resolve_qualifier(
        &self,
        qualifier: Option<&SpreadsheetFormulaQualifier>,
    ) -> Result<Option<usize>, StructuredReferenceError> {
        let Some(qualifier) = qualifier else {
            return Ok(None);
        };
        if qualifier.is_external() {
            return Err(structured_error(
                StructuredReferenceErrorKind::ExternalWorkbook,
                "Structured references to external workbooks are not supported.",
            ));
        }
        if qualifier.is_three_dimensional() {
            return Err(structured_error(
                StructuredReferenceErrorKind::Unsupported,
                "Three-dimensional structured-reference qualifiers are not supported.",
            ));
        }
        self.sheet_names
            .iter()
            .position(|name| name.eq_ignore_ascii_case(&qualifier.worksheet))
            .map(Some)
            .ok_or_else(|| {
                structured_error(
                    StructuredReferenceErrorKind::MissingTable,
                    format!(
                        "Structured-reference worksheet '{}' does not exist.",
                        qualifier.worksheet
                    ),
                )
            })
    }

    fn resolve_table(
        &self,
        parsed: &ParsedStructuredReference,
        reference: &str,
        current_sheet: usize,
        current_column: u32,
        current_row: u32,
    ) -> Result<&FormulaTableDefinition, StructuredReferenceError> {
        if let Some(table_name) = &parsed.table_name {
            let index = self
                .by_name
                .get(&table_name.to_lowercase())
                .ok_or_else(|| {
                    structured_error(
                        StructuredReferenceErrorKind::MissingTable,
                        format!("Spreadsheet table '{table_name}' does not exist."),
                    )
                })?;
            return self.definitions.get(*index).ok_or_else(|| {
                structured_error(
                    StructuredReferenceErrorKind::Unsupported,
                    "Structured-reference table index is invalid.",
                )
            });
        }

        let current = CellReference {
            column: current_column,
            row: current_row,
        };
        let mut matching = self
            .by_sheet
            .get(current_sheet)
            .into_iter()
            .flatten()
            .filter_map(|index| self.definitions.get(*index))
            .filter(|table| table.range.contains(current));
        let Some(table) = matching.next() else {
            return Err(structured_error(
                StructuredReferenceErrorKind::MissingTable,
                format!(
                    "Table-local structured reference '{reference}' requires its formula cell to be inside a Spreadsheet table."
                ),
            ));
        };
        if matching.next().is_some() {
            return Err(structured_error(
                StructuredReferenceErrorKind::Unsupported,
                format!(
                    "Table-local structured reference '{reference}' is ambiguous at the current formula cell."
                ),
            ));
        }
        Ok(table)
    }
}

fn resolve_columns(
    table: &FormulaTableDefinition,
    parsed: &ParsedStructuredReference,
) -> Result<(u32, u32), StructuredReferenceError> {
    let (first, last) = match (&parsed.first_column, &parsed.last_column) {
        (None, None) => (0, table.columns.len().saturating_sub(1)),
        (Some(first), Some(last)) => {
            let first_index = table
                .columns
                .iter()
                .position(|name| name.eq_ignore_ascii_case(first))
                .ok_or_else(|| missing_column(&table.name, first))?;
            let last_index = table
                .columns
                .iter()
                .position(|name| name.eq_ignore_ascii_case(last))
                .ok_or_else(|| missing_column(&table.name, last))?;
            if first_index > last_index {
                return Err(structured_error(
                    StructuredReferenceErrorKind::Unsupported,
                    format!("Structured-reference column range '{first}:{last}' is reversed."),
                ));
            }
            (first_index, last_index)
        }
        _ => {
            return Err(structured_error(
                StructuredReferenceErrorKind::Unsupported,
                "Structured-reference column selection is incomplete.",
            ))
        }
    };
    let first = u32::try_from(first).map_err(|_| invalid_column_index())?;
    let last = u32::try_from(last).map_err(|_| invalid_column_index())?;
    let start_column = table
        .range
        .start
        .column
        .checked_add(first)
        .ok_or_else(invalid_column_index)?;
    let end_column = table
        .range
        .start
        .column
        .checked_add(last)
        .ok_or_else(invalid_column_index)?;
    if start_column > table.range.end.column || end_column > table.range.end.column {
        return Err(invalid_column_index());
    }
    Ok((start_column, end_column))
}

fn resolve_rows(
    table: &FormulaTableDefinition,
    rows: StructuredRowSelection,
    current_row: u32,
) -> Result<Vec<(u32, u32)>, StructuredReferenceError> {
    let mut selected = Vec::<(u32, u32)>::new();
    if rows.all {
        selected.push((table.range.start.row, table.range.end.row));
    }
    if rows.headers {
        if !table.header_row {
            return Err(missing_table_rows(table, "#Headers"));
        }
        selected.push((table.range.start.row, table.range.start.row));
    }
    if rows.data {
        selected.push(data_rows(table)?);
    }
    if rows.totals {
        if !table.totals_row {
            return Err(missing_table_rows(table, "#Totals"));
        }
        selected.push((table.range.end.row, table.range.end.row));
    }
    if rows.current {
        let (start, end) = data_rows(table)?;
        if !(start..=end).contains(&current_row) {
            return Err(structured_error(
                StructuredReferenceErrorKind::Unsupported,
                format!(
                    "Structured reference #This Row requires the current formula row to be inside table '{}'.",
                    table.name
                ),
            ));
        }
        selected.push((current_row, current_row));
    }
    selected.sort_unstable();
    let mut merged = Vec::<(u32, u32)>::with_capacity(selected.len());
    for (start, end) in selected {
        if let Some((_, previous_end)) = merged.last_mut() {
            if start <= previous_end.saturating_add(1) {
                *previous_end = (*previous_end).max(end);
                continue;
            }
        }
        merged.push((start, end));
    }
    if merged.is_empty() {
        return Err(structured_error(
            StructuredReferenceErrorKind::Unsupported,
            "Structured reference selects no table rows.",
        ));
    }
    Ok(merged)
}

fn data_rows(table: &FormulaTableDefinition) -> Result<(u32, u32), StructuredReferenceError> {
    let start = table
        .range
        .start
        .row
        .checked_add(u32::from(table.header_row))
        .ok_or_else(|| {
            structured_error(
                StructuredReferenceErrorKind::Unsupported,
                "Structured-reference data range exceeds worksheet limits.",
            )
        })?;
    let end = table
        .range
        .end
        .row
        .checked_sub(u32::from(table.totals_row))
        .ok_or_else(|| {
            structured_error(
                StructuredReferenceErrorKind::Unsupported,
                "Structured-reference data range is invalid.",
            )
        })?;
    if start > end {
        return Err(structured_error(
            StructuredReferenceErrorKind::Unsupported,
            format!("Spreadsheet table '{}' has no data rows.", table.name),
        ));
    }
    Ok((start, end))
}

fn missing_table_rows(table: &FormulaTableDefinition, item: &str) -> StructuredReferenceError {
    structured_error(
        StructuredReferenceErrorKind::Unsupported,
        format!(
            "Structured reference {item} requires table '{}' to contain that row.",
            table.name
        ),
    )
}

fn invalid_column_index() -> StructuredReferenceError {
    structured_error(
        StructuredReferenceErrorKind::Unsupported,
        "Structured-reference column index is invalid.",
    )
}

fn required_table_format<'a>(table: &'a DocumentNode, key: &str) -> UseResult<&'a str> {
    table.format.get(key).map(String::as_str).ok_or_else(|| {
        table_catalog_error(format!(
            "Spreadsheet table '{}' has no '{key}' property.",
            table.path
        ))
    })
}

fn table_boolean(table: &DocumentNode, key: &str) -> UseResult<bool> {
    match required_table_format(table, key)? {
        "true" => Ok(true),
        "false" => Ok(false),
        value => Err(table_catalog_error(format!(
            "Spreadsheet table '{}' has invalid boolean '{key}={value}'.",
            table.path
        ))),
    }
}

fn missing_column(table: &str, column: &str) -> StructuredReferenceError {
    structured_error(
        StructuredReferenceErrorKind::MissingColumn,
        format!("Spreadsheet table '{table}' has no column '{column}'."),
    )
}

fn invalid_reference(reference: &str) -> StructuredReferenceError {
    structured_error(
        StructuredReferenceErrorKind::Unsupported,
        format!("Structured reference '{reference}' is not in a supported canonical form."),
    )
}

fn structured_error(
    kind: StructuredReferenceErrorKind,
    message: impl Into<String>,
) -> StructuredReferenceError {
    StructuredReferenceError {
        kind,
        message: message.into(),
    }
}

fn table_catalog_error(message: impl Into<String>) -> UseError {
    office_error(
        "use.office.spreadsheet_formula_table_catalog_invalid",
        message,
    )
}
