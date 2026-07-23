use std::collections::BTreeMap;
use std::ops::Range;

use a3s_use_core::{UseError, UseResult};

use crate::discovery::office_error;

use super::super::lexer::{self, FormulaTokenKind};
use super::super::parse_error;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum LocalStructuredReferenceContext {
    Applies,
    DoesNotApply,
    Unknown,
}

#[derive(Debug, Clone)]
pub(crate) struct StructuredReferenceRewritePlan {
    table_name: String,
    table_sheet: String,
    aliases: BTreeMap<String, String>,
    columns: BTreeMap<String, Option<String>>,
    geometry_changed: bool,
    removal: bool,
}

#[derive(Debug, Clone)]
pub(crate) struct StructuredReferenceRewriteResult {
    pub(crate) formula: String,
    pub(crate) matched: bool,
}

impl StructuredReferenceRewritePlan {
    pub(crate) fn rename(
        table_name: impl Into<String>,
        table_sheet: impl Into<String>,
        aliases: BTreeMap<String, String>,
        columns: BTreeMap<String, Option<String>>,
        geometry_changed: bool,
    ) -> Self {
        Self {
            table_name: table_name.into(),
            table_sheet: table_sheet.into(),
            aliases,
            columns,
            geometry_changed,
            removal: false,
        }
    }

    pub(crate) fn removal(
        table_name: impl Into<String>,
        table_sheet: impl Into<String>,
        aliases: impl IntoIterator<Item = String>,
    ) -> Self {
        Self {
            table_name: table_name.into(),
            table_sheet: table_sheet.into(),
            aliases: aliases
                .into_iter()
                .map(|alias| (alias.to_lowercase(), alias))
                .collect(),
            columns: BTreeMap::new(),
            geometry_changed: false,
            removal: true,
        }
    }

    pub(crate) const fn geometry_changed(&self) -> bool {
        self.geometry_changed
    }

    pub(crate) fn rewrite(
        &self,
        formula: &str,
        local_context: LocalStructuredReferenceContext,
    ) -> UseResult<StructuredReferenceRewriteResult> {
        if !formula.contains('[') {
            return Ok(StructuredReferenceRewriteResult {
                formula: formula.to_string(),
                matched: false,
            });
        }
        let body_offset = usize::from(formula.starts_with('='));
        let body = &formula[body_offset..];
        let tokens = lexer::lex(body).map_err(|failure| parse_error(body, failure))?;
        let mut replacements = Vec::<(Range<usize>, String)>::new();
        let mut matched = false;
        for token in tokens {
            let FormulaTokenKind::StructuredReference {
                qualifier,
                reference,
            } = token.kind
            else {
                continue;
            };
            let Some(open) = reference.find('[') else {
                continue;
            };
            let reference_start = body_offset
                .checked_add(token.span.end.saturating_sub(reference.len()))
                .ok_or_else(rewrite_limit_error)?;
            let table_name = &reference[..open];
            if table_name.is_empty() {
                let local_matched = self.rewrite_local(
                    &reference,
                    reference_start,
                    local_context,
                    &mut replacements,
                )?;
                matched |= local_matched;
                continue;
            }
            let Some(replacement_name) = self.aliases.get(&table_name.to_lowercase()) else {
                continue;
            };
            if qualifier
                .as_ref()
                .is_some_and(super::super::SpreadsheetFormulaQualifier::is_external)
            {
                continue;
            }
            if let Some(qualifier) = qualifier.as_ref() {
                if qualifier.is_three_dimensional() {
                    return Err(rewrite_unsupported(
                        &self.table_name,
                        "Three-dimensional structured references cannot be rewritten safely.",
                    ));
                }
                if !qualifier.worksheet.eq_ignore_ascii_case(&self.table_sheet) {
                    continue;
                }
            }
            matched = true;
            if self.removal {
                return Err(table_referenced(&self.table_name, &reference));
            }
            if replacement_name != table_name {
                replacements.push((
                    reference_start..reference_start + open,
                    replacement_name.clone(),
                ));
            }
            self.rewrite_columns(&reference, reference_start, &mut replacements)?;
        }
        Ok(StructuredReferenceRewriteResult {
            formula: apply_replacements(formula, replacements)?,
            matched,
        })
    }

    fn rewrite_local(
        &self,
        reference: &str,
        reference_start: usize,
        context: LocalStructuredReferenceContext,
        replacements: &mut Vec<(Range<usize>, String)>,
    ) -> UseResult<bool> {
        if matches!(context, LocalStructuredReferenceContext::DoesNotApply) {
            return Ok(false);
        }
        if self.removal {
            return Err(table_referenced(&self.table_name, reference));
        }
        if self.geometry_changed {
            return Err(rewrite_unsupported(
                &self.table_name,
                "Table-local structured references cannot be retained across table geometry or structural-row changes.",
            ));
        }
        let atoms = column_atoms(reference)?;
        let affected = atoms
            .iter()
            .any(|atom| self.columns.contains_key(&atom.name.to_lowercase()));
        if !affected {
            return Ok(matches!(context, LocalStructuredReferenceContext::Applies));
        }
        if matches!(context, LocalStructuredReferenceContext::Unknown) {
            return Err(rewrite_unsupported(
                &self.table_name,
                "A table-local structured reference has no provable ListObject context.",
            ));
        }
        self.rewrite_column_atoms(atoms, reference_start, replacements)?;
        Ok(true)
    }

    fn rewrite_columns(
        &self,
        reference: &str,
        reference_start: usize,
        replacements: &mut Vec<(Range<usize>, String)>,
    ) -> UseResult<()> {
        if self.columns.is_empty() {
            return Ok(());
        }
        self.rewrite_column_atoms(column_atoms(reference)?, reference_start, replacements)
    }

    fn rewrite_column_atoms(
        &self,
        atoms: Vec<ColumnAtom>,
        reference_start: usize,
        replacements: &mut Vec<(Range<usize>, String)>,
    ) -> UseResult<()> {
        for atom in atoms {
            let Some(replacement) = self.columns.get(&atom.name.to_lowercase()) else {
                continue;
            };
            let Some(replacement) = replacement else {
                return Err(rewrite_unsupported(
                    &self.table_name,
                    format!(
                        "Structured reference column '{}' would be removed from the table.",
                        atom.name
                    ),
                ));
            };
            if replacement == &atom.name {
                continue;
            }
            let encoded = encode_column(replacement, atom.plain, atom.current)?;
            replacements.push((
                reference_start + atom.range.start..reference_start + atom.range.end,
                encoded,
            ));
        }
        Ok(())
    }
}

#[derive(Debug, Clone)]
struct ColumnAtom {
    name: String,
    range: Range<usize>,
    plain: bool,
    current: bool,
}

fn column_atoms(reference: &str) -> UseResult<Vec<ColumnAtom>> {
    let Some(mut cursor) = reference.find('[') else {
        return Ok(Vec::new());
    };
    let mut atoms = Vec::new();
    while cursor < reference.len() {
        if !reference[cursor..].starts_with('[') {
            return Err(rewrite_unsupported(
                reference,
                "Structured-reference bracket groups are not canonical.",
            ));
        }
        cursor = collect_group(reference, cursor, 1, &mut atoms)?;
    }
    Ok(atoms)
}

fn collect_group(
    reference: &str,
    start: usize,
    depth: usize,
    atoms: &mut Vec<ColumnAtom>,
) -> UseResult<usize> {
    let mut cursor = start.checked_add(1).ok_or_else(rewrite_limit_error)?;
    let content_start = cursor;
    let mut has_child = false;
    while cursor < reference.len() {
        let character = reference[cursor..]
            .chars()
            .next()
            .ok_or_else(rewrite_limit_error)?;
        if character == '\'' {
            cursor = cursor
                .checked_add(character.len_utf8())
                .ok_or_else(rewrite_limit_error)?;
            let escaped = reference[cursor..].chars().next().ok_or_else(|| {
                rewrite_unsupported(
                    reference,
                    "Structured-reference escape has no following character.",
                )
            })?;
            cursor = cursor
                .checked_add(escaped.len_utf8())
                .ok_or_else(rewrite_limit_error)?;
            continue;
        }
        match character {
            '[' => {
                has_child = true;
                cursor = collect_group(reference, cursor, depth.saturating_add(1), atoms)?;
            }
            ']' => {
                if !has_child {
                    push_column_atom(reference, content_start..cursor, depth, atoms)?;
                }
                return cursor
                    .checked_add(character.len_utf8())
                    .ok_or_else(rewrite_limit_error);
            }
            _ => {
                cursor = cursor
                    .checked_add(character.len_utf8())
                    .ok_or_else(rewrite_limit_error)?;
            }
        }
    }
    Err(rewrite_unsupported(
        reference,
        "Structured-reference bracket is not closed.",
    ))
}

fn push_column_atom(
    reference: &str,
    mut range: Range<usize>,
    depth: usize,
    atoms: &mut Vec<ColumnAtom>,
) -> UseResult<()> {
    let raw = reference
        .get(range.clone())
        .ok_or_else(rewrite_limit_error)?;
    if table_item(raw) {
        return Ok(());
    }
    let current = depth == 1 && raw.starts_with('@');
    if current {
        range.start = range.start.checked_add(1).ok_or_else(rewrite_limit_error)?;
    }
    let raw = reference
        .get(range.clone())
        .ok_or_else(rewrite_limit_error)?;
    if raw.is_empty() {
        return Err(rewrite_unsupported(
            reference,
            "Structured-reference column is empty.",
        ));
    }
    atoms.push(ColumnAtom {
        name: decode_atom(raw)?,
        range,
        plain: depth == 1,
        current,
    });
    Ok(())
}

fn decode_atom(raw: &str) -> UseResult<String> {
    let mut output = String::with_capacity(raw.len());
    let mut cursor = 0_usize;
    while cursor < raw.len() {
        let character = raw[cursor..]
            .chars()
            .next()
            .ok_or_else(rewrite_limit_error)?;
        cursor = cursor
            .checked_add(character.len_utf8())
            .ok_or_else(rewrite_limit_error)?;
        if character == '\'' {
            let escaped = raw[cursor..].chars().next().ok_or_else(|| {
                rewrite_unsupported(
                    raw,
                    "Structured-reference escape has no following character.",
                )
            })?;
            cursor = cursor
                .checked_add(escaped.len_utf8())
                .ok_or_else(rewrite_limit_error)?;
            output.push(escaped);
        } else {
            output.push(character);
        }
    }
    Ok(output)
}

fn encode_column(value: &str, plain: bool, current: bool) -> UseResult<String> {
    if plain && value.contains(['[', ']', ',', ':']) {
        return Err(rewrite_unsupported(
            value,
            "The replacement column requires a nested structured-reference form.",
        ));
    }
    let mut output = String::with_capacity(value.len());
    for (index, character) in value.chars().enumerate() {
        let escape_leading = index == 0
            && ((plain && !current && matches!(character, '#' | '@'))
                || (!plain && character == '#'));
        if escape_leading || matches!(character, '\'' | '[' | ']') {
            output.push('\'');
        }
        output.push(character);
    }
    Ok(output)
}

fn table_item(value: &str) -> bool {
    ["#all", "#headers", "#data", "#totals", "#this row"]
        .into_iter()
        .any(|item| value.eq_ignore_ascii_case(item))
}

fn apply_replacements(
    formula: &str,
    mut replacements: Vec<(Range<usize>, String)>,
) -> UseResult<String> {
    if replacements.is_empty() {
        return Ok(formula.to_string());
    }
    replacements.sort_by_key(|(range, _)| (range.start, range.end));
    if replacements
        .windows(2)
        .any(|pair| pair[0].0.end > pair[1].0.start)
    {
        return Err(rewrite_limit_error());
    }
    let mut output = formula.to_string();
    for (range, replacement) in replacements.into_iter().rev() {
        if !output.is_char_boundary(range.start) || !output.is_char_boundary(range.end) {
            return Err(rewrite_limit_error());
        }
        output.replace_range(range, &replacement);
    }
    Ok(output)
}

fn table_referenced(table: &str, reference: &str) -> UseError {
    office_error(
        "use.office.spreadsheet_table_referenced",
        format!(
            "Spreadsheet table '{table}' cannot be removed while structured reference '{reference}' still targets it."
        ),
    )
    .with_detail("table", table)
    .with_detail("reference", reference)
}

fn rewrite_unsupported(table: &str, reason: impl Into<String>) -> UseError {
    office_error(
        "use.office.spreadsheet_table_formula_rewrite_unsupported",
        format!(
            "Spreadsheet table '{table}' cannot be changed without an unsafe structured-reference rewrite: {}",
            reason.into()
        ),
    )
    .with_detail("table", table)
}

fn rewrite_limit_error() -> UseError {
    office_error(
        "use.office.spreadsheet_table_formula_rewrite_unsupported",
        "Spreadsheet structured-reference rewrite exceeded safe text boundaries.",
    )
}

#[cfg(test)]
mod tests {
    use super::*;

    fn rename_plan(geometry_changed: bool) -> StructuredReferenceRewritePlan {
        StructuredReferenceRewritePlan::rename(
            "Sales",
            "Sheet1",
            BTreeMap::from([("sales".into(), "Orders".into())]),
            BTreeMap::from([("qty".into(), Some("Units".into()))]),
            geometry_changed,
        )
    }

    #[test]
    fn rewrite_preserves_strings_and_external_workbooks() {
        let rewritten = rename_plan(false)
            .rewrite(
                r#"=CONCAT("Sales[Qty]",SUM(Sales[[#Data],[Qty]]),'Sheet1'!Sales[@Qty],'[Book.xlsx]Sheet1'!Sales[Qty])"#,
                LocalStructuredReferenceContext::Unknown,
            )
            .unwrap();
        assert_eq!(
            rewritten.formula,
            r#"=CONCAT("Sales[Qty]",SUM(Orders[[#Data],[Units]]),'Sheet1'!Orders[@Units],'[Book.xlsx]Sheet1'!Sales[Qty])"#
        );
        assert!(rewritten.matched);
    }

    #[test]
    fn rewrite_requires_provable_local_context() {
        let plan = rename_plan(false);
        let applied = plan
            .rewrite("[@Qty]", LocalStructuredReferenceContext::Applies)
            .unwrap();
        assert_eq!(applied.formula, "[@Units]");
        assert!(applied.matched);

        let unrelated = plan
            .rewrite("[@Qty]", LocalStructuredReferenceContext::DoesNotApply)
            .unwrap();
        assert_eq!(unrelated.formula, "[@Qty]");
        assert!(!unrelated.matched);

        assert_eq!(
            plan.rewrite("[@Qty]", LocalStructuredReferenceContext::Unknown,)
                .unwrap_err()
                .code,
            "use.office.spreadsheet_table_formula_rewrite_unsupported"
        );
    }

    #[test]
    fn removal_and_geometry_changes_fail_closed_for_matching_references() {
        let removal =
            StructuredReferenceRewritePlan::removal("Sales", "Sheet1", ["Sales".to_string()]);
        assert_eq!(
            removal
                .rewrite(
                    "SUM(Sales[Qty])",
                    LocalStructuredReferenceContext::DoesNotApply,
                )
                .unwrap_err()
                .code,
            "use.office.spreadsheet_table_referenced"
        );
        let external = removal
            .rewrite(
                "'[Book.xlsx]Sheet1'!Sales[Qty]",
                LocalStructuredReferenceContext::Unknown,
            )
            .unwrap();
        assert_eq!(external.formula, "'[Book.xlsx]Sheet1'!Sales[Qty]");
        assert!(!external.matched);

        assert_eq!(
            rename_plan(true)
                .rewrite("[@Qty]", LocalStructuredReferenceContext::Applies,)
                .unwrap_err()
                .code,
            "use.office.spreadsheet_table_formula_rewrite_unsupported"
        );
    }
}
