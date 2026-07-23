use std::collections::BTreeSet;

pub(super) fn formula_references_missing_sheet(
    formula: &str,
    sheet_names: &BTreeSet<String>,
) -> bool {
    let bytes = formula.as_bytes();
    let mut index = 0;
    let mut in_string = false;
    while index < bytes.len() {
        match bytes[index] {
            b'"' => {
                if in_string && bytes.get(index + 1) == Some(&b'"') {
                    index += 2;
                    continue;
                }
                in_string = !in_string;
                index += 1;
                continue;
            }
            _ if in_string => {
                index += 1;
                continue;
            }
            b'\'' => {
                let start = index + 1;
                let mut cursor = start;
                while cursor < bytes.len() {
                    if bytes[cursor] != b'\'' {
                        cursor += 1;
                        continue;
                    }
                    if bytes.get(cursor + 1) == Some(&b'\'') {
                        cursor += 2;
                        continue;
                    }
                    if bytes.get(cursor + 1) == Some(&b'!') {
                        let name = formula[start..cursor].replace("''", "'");
                        if !is_external_reference(&name)
                            && !sheet_names.contains(&name.to_lowercase())
                        {
                            return true;
                        }
                        index = cursor + 2;
                    } else {
                        index = cursor + 1;
                    }
                    break;
                }
                if cursor == bytes.len() {
                    return false;
                }
                continue;
            }
            byte if is_bare_sheet_start(byte)
                && (index == 0 || !is_bare_sheet_boundary_blocker(bytes[index - 1])) =>
            {
                let start = index;
                let external = index > 0 && bytes[index - 1] == b']';
                index += 1;
                while index < bytes.len() && is_bare_sheet_byte(bytes[index]) {
                    index += 1;
                }
                if bytes.get(index) == Some(&b'!') {
                    let name = &formula[start..index];
                    if !external && !sheet_names.contains(&name.to_lowercase()) {
                        return true;
                    }
                    index += 1;
                }
                continue;
            }
            _ => index += 1,
        }
    }
    false
}

fn is_external_reference(name: &str) -> bool {
    name.starts_with('[') && name.contains(']')
}

fn is_bare_sheet_start(byte: u8) -> bool {
    byte.is_ascii_alphabetic() || byte == b'_'
}

fn is_bare_sheet_byte(byte: u8) -> bool {
    byte.is_ascii_alphanumeric() || matches!(byte, b'_' | b'.')
}

fn is_bare_sheet_boundary_blocker(byte: u8) -> bool {
    is_bare_sheet_byte(byte) || matches!(byte, b'\'' | b'#')
}

const ERROR_LITERALS: &[&[u8]] = &[
    b"#NULL!",
    b"#DIV/0!",
    b"#VALUE!",
    b"#REF!",
    b"#NAME?",
    b"#NUM!",
    b"#N/A",
    b"#GETTING_DATA",
    b"#SPILL!",
    b"#CALC!",
    b"#FIELD!",
    b"#BLOCKED!",
    b"#UNKNOWN!",
];

pub(super) fn formula_contains_error_literal(formula: &str) -> bool {
    let bytes = formula.as_bytes();
    let mut index = 0;
    let mut in_string = false;
    while index < bytes.len() {
        if bytes[index] == b'"' {
            if in_string && bytes.get(index + 1) == Some(&b'"') {
                index += 2;
                continue;
            }
            in_string = !in_string;
            index += 1;
            continue;
        }
        if !in_string
            && ERROR_LITERALS.iter().any(|token| {
                bytes
                    .get(index..index.saturating_add(token.len()))
                    .is_some_and(|candidate| candidate.eq_ignore_ascii_case(token))
            })
        {
            return true;
        }
        index += 1;
    }
    false
}

pub(super) fn is_formula_error(value: &str) -> bool {
    let value = value.trim().as_bytes();
    ERROR_LITERALS
        .iter()
        .any(|literal| value.eq_ignore_ascii_case(literal))
}
