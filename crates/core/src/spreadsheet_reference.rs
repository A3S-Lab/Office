use std::{
    cmp::Reverse,
    collections::{BTreeMap, BinaryHeap},
};

use a3s_use_core::{UseError, UseResult};

use crate::discovery::office_error;

pub(crate) const MAX_COLUMNS: u32 = 16_384;
pub(crate) const MAX_ROWS: u32 = 1_048_576;
pub(crate) const MAX_DATA_VALIDATIONS: usize = 65_534;
pub(crate) const MAX_DATA_VALIDATION_RANGES: usize = 1_024;

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
pub(crate) struct CellReference {
    pub column: u32,
    pub row: u32,
}

impl CellReference {
    pub fn parse(reference: &str) -> UseResult<Self> {
        let normalized = reference.to_ascii_uppercase();
        let column_length = normalized
            .chars()
            .take_while(|character| character.is_ascii_alphabetic())
            .count();
        if column_length == 0
            || column_length == normalized.len()
            || !normalized[column_length..]
                .chars()
                .all(|character| character.is_ascii_digit())
        {
            return Err(reference_error(reference, "is invalid"));
        }
        let column = parse_column_name(&normalized[..column_length])?;
        let row = normalized[column_length..]
            .parse::<u32>()
            .ok()
            .filter(|row| (1..=MAX_ROWS).contains(row))
            .ok_or_else(|| reference_error(reference, "is outside row limits"))?;
        Ok(Self { column, row })
    }

    pub fn a1(self) -> String {
        format!("{}{}", column_name(self.column), self.row)
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct CellRange {
    pub start: CellReference,
    pub end: CellReference,
}

impl CellRange {
    pub fn parse(reference: &str) -> UseResult<Self> {
        let (start, end) = reference
            .split_once(':')
            .map_or((reference, reference), |(start, end)| (start, end));
        if start.is_empty() || end.is_empty() || end.contains(':') {
            return Err(reference_error(
                reference,
                "is not a cell or rectangular range",
            ));
        }
        let start = CellReference::parse(start)?;
        let end = CellReference::parse(end)?;
        Ok(Self {
            start: CellReference {
                column: start.column.min(end.column),
                row: start.row.min(end.row),
            },
            end: CellReference {
                column: start.column.max(end.column),
                row: start.row.max(end.row),
            },
        })
    }

    pub fn is_single_cell(self) -> bool {
        self.start == self.end
    }

    pub fn contains(self, cell: CellReference) -> bool {
        (self.start.column..=self.end.column).contains(&cell.column)
            && (self.start.row..=self.end.row).contains(&cell.row)
    }

    pub fn intersects(self, other: Self) -> bool {
        self.start.column <= other.end.column
            && other.start.column <= self.end.column
            && self.start.row <= other.end.row
            && other.start.row <= self.end.row
    }

    pub fn cell_count(self) -> UseResult<usize> {
        let columns = u64::from(self.end.column - self.start.column + 1);
        let rows = u64::from(self.end.row - self.start.row + 1);
        usize::try_from(columns * rows).map_err(|_| {
            office_error(
                "use.office.spreadsheet_range_too_large",
                "Spreadsheet range size does not fit this platform.",
            )
        })
    }

    pub fn a1(self) -> String {
        if self.is_single_cell() {
            self.start.a1()
        } else {
            format!("{}:{}", self.start.a1(), self.end.a1())
        }
    }
}

pub(crate) fn first_intersecting_ranges(ranges: &[CellRange]) -> Option<(usize, usize)> {
    let mut ordered = (0..ranges.len()).collect::<Vec<_>>();
    ordered.sort_unstable_by_key(|index| {
        let range = ranges[*index];
        (
            range.start.row,
            range.start.column,
            range.end.row,
            range.end.column,
        )
    });

    let mut active_by_column = BTreeMap::<u32, usize>::new();
    let mut expiration = BinaryHeap::<Reverse<(u32, usize)>>::new();
    for index in ordered {
        let current = ranges[index];
        while let Some(Reverse((end_row, expired))) = expiration.peek().copied() {
            if end_row >= current.start.row {
                break;
            }
            expiration.pop();
            let start_column = ranges[expired].start.column;
            if active_by_column.get(&start_column).copied() == Some(expired) {
                active_by_column.remove(&start_column);
            }
        }

        if let Some((_, existing)) = active_by_column.range(..=current.end.column).next_back() {
            let existing = *existing;
            if ranges[existing].end.column >= current.start.column {
                return Some((existing, index));
            }
        }

        active_by_column.insert(current.start.column, index);
        expiration.push(Reverse((current.end.row, index)));
    }
    None
}

pub(crate) fn parse_column(value: &str) -> UseResult<u32> {
    if value.is_empty() {
        return Err(column_error(value));
    }
    if value.chars().all(|character| character.is_ascii_digit()) {
        return value
            .parse::<u32>()
            .ok()
            .filter(|column| (1..=MAX_COLUMNS).contains(column))
            .ok_or_else(|| column_error(value));
    }
    parse_column_name(&value.to_ascii_uppercase())
}

pub(crate) fn column_name(mut column: u32) -> String {
    debug_assert!((1..=MAX_COLUMNS).contains(&column));
    let mut bytes = Vec::new();
    while column > 0 {
        column -= 1;
        bytes.push(b'A' + (column % 26) as u8);
        column /= 26;
    }
    bytes.reverse();
    String::from_utf8(bytes).unwrap_or_default()
}

fn parse_column_name(value: &str) -> UseResult<u32> {
    if value.is_empty() || !value.bytes().all(|byte| byte.is_ascii_uppercase()) {
        return Err(column_error(value));
    }
    value
        .bytes()
        .try_fold(0_u32, |column, byte| {
            column
                .checked_mul(26)
                .and_then(|value| value.checked_add(u32::from(byte - b'A') + 1))
        })
        .filter(|column| (1..=MAX_COLUMNS).contains(column))
        .ok_or_else(|| column_error(value))
}

fn reference_error(reference: &str, reason: &str) -> UseError {
    office_error(
        "use.office.spreadsheet_cell_reference_invalid",
        format!("Spreadsheet cell reference '{reference}' {reason}."),
    )
    .with_detail("reference", reference)
}

fn column_error(column: &str) -> UseError {
    office_error(
        "use.office.spreadsheet_column_invalid",
        format!("Spreadsheet column '{column}' is outside A:XFD (1-16384)."),
    )
    .with_detail("column", column)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parses_and_normalizes_cells_ranges_and_columns() {
        assert_eq!(
            CellReference::parse("xfd1048576").unwrap().a1(),
            "XFD1048576"
        );
        assert_eq!(CellRange::parse("B3:A1").unwrap().a1(), "A1:B3");
        assert_eq!(CellRange::parse("C7").unwrap().cell_count().unwrap(), 1);
        assert!(CellRange::parse("A1:B2")
            .unwrap()
            .intersects(CellRange::parse("B2:C3").unwrap()));
        assert!(!CellRange::parse("A1:B2")
            .unwrap()
            .intersects(CellRange::parse("C1:D2").unwrap()));
        assert_eq!(parse_column("XFD").unwrap(), MAX_COLUMNS);
        assert_eq!(parse_column("16384").unwrap(), MAX_COLUMNS);

        for invalid in ["", "A", "1", "A0", "XFE1", "A1048577", "A1:B2:C3"] {
            assert!(CellRange::parse(invalid).is_err(), "{invalid}");
        }
        for invalid in ["0", "16385", "XFE", "A1"] {
            assert!(parse_column(invalid).is_err(), "{invalid}");
        }
    }

    #[test]
    fn finds_rectangle_intersections_with_a_bounded_sweep() {
        let disjoint = (1..=20_000)
            .map(|row| CellRange {
                start: CellReference { column: 1, row },
                end: CellReference { column: 2, row },
            })
            .collect::<Vec<_>>();
        assert_eq!(first_intersecting_ranges(&disjoint), None);

        let overlapping = [
            CellRange::parse("A1:C2").unwrap(),
            CellRange::parse("D1:E2").unwrap(),
            CellRange::parse("B2:B3").unwrap(),
        ];
        let (left, right) = first_intersecting_ranges(&overlapping).unwrap();
        assert!(overlapping[left].intersects(overlapping[right]));

        let separated = [
            CellRange::parse("A1:C2").unwrap(),
            CellRange::parse("A3:C4").unwrap(),
            CellRange::parse("D1:E4").unwrap(),
        ];
        assert_eq!(first_intersecting_ranges(&separated), None);
    }
}
