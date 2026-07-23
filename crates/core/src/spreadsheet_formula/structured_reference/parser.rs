use super::{
    invalid_reference, structured_error, StructuredReferenceError, StructuredReferenceErrorKind,
};

#[derive(Debug, Clone, Copy, Default)]
pub(super) struct StructuredRowSelection {
    pub(super) all: bool,
    pub(super) headers: bool,
    pub(super) data: bool,
    pub(super) totals: bool,
    pub(super) current: bool,
}

#[derive(Debug, Clone)]
pub(super) struct ParsedStructuredReference {
    pub(super) table_name: Option<String>,
    pub(super) first_column: Option<String>,
    pub(super) last_column: Option<String>,
    pub(super) rows: StructuredRowSelection,
}

pub(super) fn parse_reference(
    reference: &str,
) -> Result<ParsedStructuredReference, StructuredReferenceError> {
    let Some(open) = reference.find('[') else {
        return Err(invalid_reference(reference));
    };
    let table_name = (!reference[..open].is_empty()).then(|| reference[..open].to_string());
    let content = outer_group(&reference[open..]).ok_or_else(|| invalid_reference(reference))?;
    let mut rows = StructuredRowSelection::default();
    let (first_column, last_column) = if let Some(current) = content.strip_prefix('@') {
        rows.current = true;
        let column = parse_current_column(current, reference)?;
        (Some(column.clone()), Some(column))
    } else if content.starts_with('[') {
        parse_nested_selection(content, reference, &mut rows)?
    } else if let Some(item) = table_item(content) {
        apply_table_item(&mut rows, item, reference)?;
        (None, None)
    } else {
        let column = parse_plain_column(content, reference)?;
        (Some(column.clone()), Some(column))
    };
    if !rows.all && !rows.headers && !rows.data && !rows.totals && !rows.current {
        rows.data = true;
    }
    Ok(ParsedStructuredReference {
        table_name,
        first_column,
        last_column,
        rows,
    })
}

fn parse_current_column(value: &str, reference: &str) -> Result<String, StructuredReferenceError> {
    if value.starts_with('[') {
        let (atom, consumed) = bracket_atom(value).ok_or_else(|| invalid_reference(reference))?;
        if consumed != value.len() {
            return Err(invalid_reference(reference));
        }
        let column = decode_atom(atom)?;
        if column.is_empty() {
            return Err(invalid_reference(reference));
        }
        return Ok(column);
    }
    parse_plain_column(value, reference)
}

fn parse_nested_selection(
    content: &str,
    reference: &str,
    rows: &mut StructuredRowSelection,
) -> Result<(Option<String>, Option<String>), StructuredReferenceError> {
    let (atoms, separators) = nested_components(content, reference)?;
    let mut columns = Vec::<(usize, String)>::new();
    for (index, atom) in atoms.iter().enumerate() {
        if let Some(item) = table_item(atom) {
            apply_table_item(rows, item, reference)?;
        } else {
            let column = decode_atom(atom)?;
            if column.is_empty() {
                return Err(invalid_reference(reference));
            }
            columns.push((index, column));
        }
    }
    match columns.as_slice() {
        [] => {
            if separators.contains(&':') {
                return Err(invalid_reference(reference));
            }
            Ok((None, None))
        }
        [(_, column)] => {
            if separators.contains(&':') {
                return Err(invalid_reference(reference));
            }
            Ok((Some(column.clone()), Some(column.clone())))
        }
        [(first_index, first), (last_index, last)]
            if *last_index == first_index.saturating_add(1)
                && separators.get(*first_index) == Some(&':')
                && separators
                    .iter()
                    .enumerate()
                    .all(|(index, separator)| index == *first_index || *separator == ',') =>
        {
            Ok((Some(first.clone()), Some(last.clone())))
        }
        _ => Err(structured_error(
            StructuredReferenceErrorKind::Unsupported,
            "Disjoint structured-reference columns are not supported.",
        )),
    }
}

fn nested_components<'a>(
    content: &'a str,
    reference: &str,
) -> Result<(Vec<&'a str>, Vec<char>), StructuredReferenceError> {
    let mut atoms = Vec::new();
    let mut separators = Vec::new();
    let mut cursor = 0_usize;
    loop {
        let (atom, consumed) =
            bracket_atom(&content[cursor..]).ok_or_else(|| invalid_reference(reference))?;
        atoms.push(atom);
        cursor = cursor
            .checked_add(consumed)
            .ok_or_else(|| invalid_reference(reference))?;
        if cursor == content.len() {
            break;
        }
        let separator = content[cursor..]
            .chars()
            .next()
            .ok_or_else(|| invalid_reference(reference))?;
        if !matches!(separator, ',' | ':') {
            return Err(invalid_reference(reference));
        }
        separators.push(separator);
        cursor = cursor
            .checked_add(separator.len_utf8())
            .ok_or_else(|| invalid_reference(reference))?;
        if cursor >= content.len() {
            return Err(invalid_reference(reference));
        }
    }
    Ok((atoms, separators))
}

fn parse_plain_column(value: &str, reference: &str) -> Result<String, StructuredReferenceError> {
    if value.is_empty() || value.contains(['[', ']', ',', ':']) {
        return Err(invalid_reference(reference));
    }
    let column = decode_atom(value)?;
    if column.is_empty() {
        return Err(invalid_reference(reference));
    }
    Ok(column)
}

#[derive(Debug, Clone, Copy)]
enum TableItem {
    All,
    Headers,
    Data,
    Totals,
    Current,
}

fn table_item(value: &str) -> Option<TableItem> {
    [
        ("#all", TableItem::All),
        ("#headers", TableItem::Headers),
        ("#data", TableItem::Data),
        ("#totals", TableItem::Totals),
        ("#this row", TableItem::Current),
    ]
    .into_iter()
    .find_map(|(name, item)| value.eq_ignore_ascii_case(name).then_some(item))
}

fn apply_table_item(
    rows: &mut StructuredRowSelection,
    item: TableItem,
    reference: &str,
) -> Result<(), StructuredReferenceError> {
    match item {
        TableItem::All => rows.all = true,
        TableItem::Headers => rows.headers = true,
        TableItem::Data => rows.data = true,
        TableItem::Totals => rows.totals = true,
        TableItem::Current => rows.current = true,
    }
    if rows.current && (rows.all || rows.headers || rows.data || rows.totals) {
        return Err(structured_error(
            StructuredReferenceErrorKind::Unsupported,
            format!(
                "Structured reference '{reference}' cannot combine #This Row with another item."
            ),
        ));
    }
    Ok(())
}

fn outer_group(value: &str) -> Option<&str> {
    if !value.starts_with('[') {
        return None;
    }
    let end = matching_bracket(value, 0)?;
    (end == value.len()).then_some(&value[1..value.len() - 1])
}

fn bracket_atom(value: &str) -> Option<(&str, usize)> {
    if !value.starts_with('[') {
        return None;
    }
    let end = matching_bracket(value, 0)?;
    Some((&value[1..end - 1], end))
}

fn matching_bracket(value: &str, start: usize) -> Option<usize> {
    let mut depth = 0_usize;
    let mut cursor = start;
    while cursor < value.len() {
        let character = value[cursor..].chars().next()?;
        cursor += character.len_utf8();
        if character == '\'' {
            let escaped = value[cursor..].chars().next()?;
            cursor += escaped.len_utf8();
            continue;
        }
        match character {
            '[' => depth = depth.checked_add(1)?,
            ']' => {
                depth = depth.checked_sub(1)?;
                if depth == 0 {
                    return Some(cursor);
                }
            }
            _ => {}
        }
    }
    None
}

fn decode_atom(value: &str) -> Result<String, StructuredReferenceError> {
    let mut output = String::with_capacity(value.len());
    let mut cursor = 0_usize;
    while cursor < value.len() {
        let character = value[cursor..]
            .chars()
            .next()
            .ok_or_else(|| invalid_reference(value))?;
        cursor += character.len_utf8();
        if character == '\'' {
            let escaped = value[cursor..]
                .chars()
                .next()
                .ok_or_else(|| invalid_reference(value))?;
            cursor += escaped.len_utf8();
            output.push(escaped);
        } else {
            output.push(character);
        }
    }
    Ok(output)
}
