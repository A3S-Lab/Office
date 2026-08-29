use std::collections::{BTreeMap, BTreeSet};

use a3s_office_formula_parser::{
    parse_spreadsheet_structured_reference, SpreadsheetFormulaQualifier,
    SpreadsheetStructuredReference,
};

use super::{
    SpreadsheetInputSheet, SpreadsheetInputTable, MAX_SPREADSHEET_COLUMNS,
    MAX_SPREADSHEET_IDENTIFIER_BYTES, MAX_SPREADSHEET_ROWS,
};

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) struct SpreadsheetTableArea {
    pub(super) sheet: usize,
    pub(super) start_row: u32,
    pub(super) end_row: u32,
    pub(super) start_column: u32,
    pub(super) end_column: u32,
}

#[derive(Debug, Clone)]
struct SpreadsheetTableDefinition {
    name: String,
    sheet: usize,
    start_row: u32,
    end_row: u32,
    start_column: u32,
    end_column: u32,
    columns: Vec<String>,
    header_row: bool,
    totals_row: bool,
}

#[derive(Debug, Clone, Default)]
pub(super) struct SpreadsheetTableCatalog {
    sheet_names: Vec<String>,
    definitions: Vec<SpreadsheetTableDefinition>,
    by_name: BTreeMap<String, usize>,
    by_sheet: Vec<Vec<usize>>,
}

impl SpreadsheetTableCatalog {
    pub(super) fn from_sheets(sheets: &[SpreadsheetInputSheet]) -> Result<Self, String> {
        validate_tables(sheets)?;
        let sheet_names = sheets
            .iter()
            .map(|sheet| sheet.name.clone())
            .collect::<Vec<_>>();
        let mut definitions = Vec::new();
        let mut by_name = BTreeMap::new();
        let mut by_sheet = vec![Vec::new(); sheets.len()];
        for (sheet_index, sheet) in sheets.iter().enumerate() {
            for table in &sheet.tables {
                let definition = SpreadsheetTableDefinition {
                    name: table.name.clone(),
                    sheet: sheet_index,
                    start_row: table.start_row,
                    end_row: table.end_row,
                    start_column: table.start_column,
                    end_column: table.end_column,
                    columns: table.columns.clone(),
                    header_row: table.header_row,
                    totals_row: table.totals_row,
                };
                let index = definitions.len();
                definitions.push(definition);
                by_sheet[sheet_index].push(index);
                let aliases = [Some(table.name.as_str()), table.display_name.as_deref()];
                for alias in aliases.into_iter().flatten() {
                    by_name.insert(alias.to_lowercase(), index);
                }
            }
        }
        Ok(Self {
            sheet_names,
            definitions,
            by_name,
            by_sheet,
        })
    }

    pub(super) fn resolve(
        &self,
        qualifier: Option<&SpreadsheetFormulaQualifier>,
        reference: &str,
        current_sheet: usize,
        current_column: u32,
        current_row: u32,
    ) -> Result<Vec<SpreadsheetTableArea>, String> {
        let parsed =
            parse_spreadsheet_structured_reference(reference).map_err(|error| error.to_string())?;
        let qualifier_sheet = self.resolve_qualifier(qualifier)?;
        let table = self.resolve_table(
            &parsed,
            reference,
            current_sheet,
            current_column,
            current_row,
        )?;
        if qualifier_sheet.is_some_and(|sheet| sheet != table.sheet) {
            return Err(format!(
                "Spreadsheet table '{}' is not on worksheet '{}'.",
                table.name,
                qualifier
                    .map(|value| value.worksheet.as_str())
                    .unwrap_or_default()
            ));
        }
        if parsed.rows.current && current_sheet != table.sheet {
            return Err(format!(
                "Structured reference #This Row requires the current formula cell to be on table '{}'.",
                table.name
            ));
        }
        let (start_column, end_column) = resolve_columns(table, &parsed)?;
        resolve_rows(table, parsed.rows, current_row)?
            .into_iter()
            .map(|(start_row, end_row)| {
                Ok(SpreadsheetTableArea {
                    sheet: table.sheet,
                    start_row,
                    end_row,
                    start_column,
                    end_column,
                })
            })
            .collect()
    }

    fn resolve_qualifier(
        &self,
        qualifier: Option<&SpreadsheetFormulaQualifier>,
    ) -> Result<Option<usize>, String> {
        let Some(qualifier) = qualifier else {
            return Ok(None);
        };
        if qualifier.is_external() {
            return Err("Structured references to external workbooks are not supported.".into());
        }
        if qualifier.is_three_dimensional() {
            return Err(
                "Three-dimensional structured-reference qualifiers are not supported.".into(),
            );
        }
        self.sheet_names
            .iter()
            .position(|name| name.eq_ignore_ascii_case(&qualifier.worksheet))
            .map(Some)
            .ok_or_else(|| {
                format!(
                    "Structured-reference worksheet '{}' does not exist.",
                    qualifier.worksheet
                )
            })
    }

    fn resolve_table(
        &self,
        parsed: &SpreadsheetStructuredReference,
        reference: &str,
        current_sheet: usize,
        current_column: u32,
        current_row: u32,
    ) -> Result<&SpreadsheetTableDefinition, String> {
        if let Some(table_name) = &parsed.table_name {
            let index = self
                .by_name
                .get(&table_name.to_lowercase())
                .ok_or_else(|| format!("Spreadsheet table '{table_name}' does not exist."))?;
            return self
                .definitions
                .get(*index)
                .ok_or_else(|| "Structured-reference table index is invalid.".into());
        }
        let matching = self
            .by_sheet
            .get(current_sheet)
            .into_iter()
            .flatten()
            .filter_map(|index| self.definitions.get(*index))
            .filter(|table| {
                current_row >= table.start_row
                    && current_row <= table.end_row
                    && current_column >= table.start_column
                    && current_column <= table.end_column
            })
            .collect::<Vec<_>>();
        match matching.as_slice() {
            [] => Err(format!(
                "Table-local structured reference '{reference}' requires the current formula cell to be inside a Spreadsheet table."
            )),
            [table] => Ok(table),
            _ => Err(format!(
                "Table-local structured reference '{reference}' is ambiguous at the current formula cell."
            )),
        }
    }
}

fn resolve_columns(
    table: &SpreadsheetTableDefinition,
    parsed: &SpreadsheetStructuredReference,
) -> Result<(u32, u32), String> {
    let (first, last) = match (&parsed.first_column, &parsed.last_column) {
        (None, None) => (0, table.columns.len().saturating_sub(1)),
        (Some(first), Some(last)) => {
            let first_index = table
                .columns
                .iter()
                .position(|name| name.eq_ignore_ascii_case(first))
                .ok_or_else(|| {
                    format!(
                        "Spreadsheet table '{}' has no column '{first}'.",
                        table.name
                    )
                })?;
            let last_index = table
                .columns
                .iter()
                .position(|name| name.eq_ignore_ascii_case(last))
                .ok_or_else(|| {
                    format!("Spreadsheet table '{}' has no column '{last}'.", table.name)
                })?;
            if first_index > last_index {
                return Err(format!(
                    "Structured-reference column range '{first}:{last}' is reversed."
                ));
            }
            (first_index, last_index)
        }
        _ => return Err("Structured-reference column selection is incomplete.".into()),
    };
    let start_column = table
        .start_column
        .checked_add(
            u32::try_from(first).map_err(|_| "Structured-reference column index is invalid.")?,
        )
        .ok_or("Structured-reference column index is invalid.")?;
    let end_column = table
        .start_column
        .checked_add(
            u32::try_from(last).map_err(|_| "Structured-reference column index is invalid.")?,
        )
        .ok_or("Structured-reference column index is invalid.")?;
    if end_column > table.end_column {
        return Err("Structured-reference column index is invalid.".into());
    }
    Ok((start_column, end_column))
}

fn resolve_rows(
    table: &SpreadsheetTableDefinition,
    rows: a3s_office_formula_parser::SpreadsheetStructuredRowSelection,
    current_row: u32,
) -> Result<Vec<(u32, u32)>, String> {
    let mut selected = Vec::new();
    if rows.all {
        selected.push((table.start_row, table.end_row));
    }
    if rows.headers {
        if !table.header_row {
            return Err(format!(
                "Structured reference #Headers requires table '{}' to contain that row.",
                table.name
            ));
        }
        selected.push((table.start_row, table.start_row));
    }
    if rows.data {
        selected.push(data_rows(table)?);
    }
    if rows.totals {
        if !table.totals_row {
            return Err(format!(
                "Structured reference #Totals requires table '{}' to contain that row.",
                table.name
            ));
        }
        selected.push((table.end_row, table.end_row));
    }
    if rows.current {
        let (start, end) = data_rows(table)?;
        if !(start..=end).contains(&current_row) {
            return Err(format!(
                "Structured reference #This Row requires the current formula row to be inside table '{}'.",
                table.name
            ));
        }
        selected.push((current_row, current_row));
    }
    selected.sort_unstable();
    let mut merged = Vec::<(u32, u32)>::new();
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
        return Err("Structured reference selects no table rows.".into());
    }
    Ok(merged)
}

fn data_rows(table: &SpreadsheetTableDefinition) -> Result<(u32, u32), String> {
    let start = table
        .start_row
        .checked_add(u32::from(table.header_row))
        .ok_or("Structured-reference data range exceeds worksheet limits.")?;
    let end = table
        .end_row
        .checked_sub(u32::from(table.totals_row))
        .ok_or("Structured-reference data range is invalid.")?;
    if start > end {
        return Err(format!(
            "Spreadsheet table '{}' has no data rows.",
            table.name
        ));
    }
    Ok((start, end))
}

pub(super) fn validate_tables(sheets: &[SpreadsheetInputSheet]) -> Result<(), String> {
    let mut aliases = BTreeMap::<String, usize>::new();
    let mut table_count = 0_usize;
    for sheet in sheets {
        let mut ranges = Vec::<&SpreadsheetInputTable>::new();
        for table in &sheet.tables {
            let table_index = table_count;
            table_count = table_count.saturating_add(1);
            if table_count > 1_024 {
                return Err(
                    "A Spreadsheet calculation request may contain at most 1024 tables.".into(),
                );
            }
            for (name, value, maximum) in [
                ("start row", table.start_row, MAX_SPREADSHEET_ROWS),
                ("end row", table.end_row, MAX_SPREADSHEET_ROWS),
                ("start column", table.start_column, MAX_SPREADSHEET_COLUMNS),
                ("end column", table.end_column, MAX_SPREADSHEET_COLUMNS),
            ] {
                if value >= maximum {
                    return Err(format!(
                        "Spreadsheet table '{}' has an invalid {name}.",
                        table.name
                    ));
                }
            }
            if table.start_row > table.end_row
                || table.start_column > table.end_column
                || table.columns.len()
                    != usize::try_from(table.end_column - table.start_column + 1).map_err(|_| {
                        format!("Spreadsheet table '{}' has an invalid width.", table.name)
                    })?
            {
                return Err(format!(
                    "Spreadsheet table '{}' has an invalid range or column count.",
                    table.name
                ));
            }
            validate_table_identifier(&table.name, "name")?;
            if let Some(display_name) = &table.display_name {
                validate_table_identifier(display_name, "displayName")?;
            }
            for alias in [Some(table.name.as_str()), table.display_name.as_deref()]
                .into_iter()
                .flatten()
            {
                let key = alias.to_lowercase();
                if aliases
                    .insert(key, table_index)
                    .is_some_and(|previous| previous != table_index)
                {
                    return Err(format!("Spreadsheet table name '{alias}' is ambiguous."));
                }
            }
            let mut columns = BTreeSet::new();
            for column in &table.columns {
                validate_table_identifier(column, "column")?;
                if !columns.insert(column.to_lowercase()) {
                    return Err(format!(
                        "Spreadsheet table '{}' contains duplicate column names.",
                        table.name
                    ));
                }
            }
            if table.header_row && table.totals_row && table.start_row == table.end_row {
                return Err(format!(
                    "Spreadsheet table '{}' cannot use header and totals rows in one row.",
                    table.name
                ));
            }
            for previous in &ranges {
                if table.start_row <= previous.end_row
                    && table.end_row >= previous.start_row
                    && table.start_column <= previous.end_column
                    && table.end_column >= previous.start_column
                {
                    return Err(format!(
                        "Spreadsheet tables '{}' and '{}' overlap.",
                        previous.name, table.name
                    ));
                }
            }
            ranges.push(table);
        }
    }
    Ok(())
}

fn validate_table_identifier(value: &str, kind: &str) -> Result<(), String> {
    if value.trim().is_empty() || value.len() > MAX_SPREADSHEET_IDENTIFIER_BYTES {
        return Err(format!(
            "Spreadsheet table {kind} must contain 1-{MAX_SPREADSHEET_IDENTIFIER_BYTES} UTF-8 bytes."
        ));
    }
    Ok(())
}
