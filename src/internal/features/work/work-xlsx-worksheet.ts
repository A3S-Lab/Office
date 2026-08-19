import type { CellObject, WorkSheet } from 'xlsx';

export interface XlsxWorksheetCellEntry {
  address: string;
  cell: CellObject;
  column: number;
  row: number;
}

/**
 * Iterates both SheetJS sparse address maps and dense row arrays without
 * materializing a second list of every worksheet cell.
 */
export function* xlsxWorksheetCellEntries(
  worksheet: WorkSheet,
): Generator<XlsxWorksheetCellEntry> {
  if (Array.isArray(worksheet)) {
    for (let row = 0; row < worksheet.length; row += 1) {
      const cells = worksheet[row];
      if (!Array.isArray(cells)) continue;
      for (let column = 0; column < cells.length; column += 1) {
        const cell = cells[column];
        if (!cell || typeof cell !== 'object' || Array.isArray(cell)) continue;
        yield {
          address: xlsxCellAddress(row, column),
          cell: cell as CellObject,
          column,
          row,
        };
      }
    }
    return;
  }

  for (const [address, cell] of Object.entries(worksheet)) {
    if (address.startsWith('!') || !cell || typeof cell !== 'object') continue;
    const position = decodeXlsxCellAddress(address);
    if (!position) continue;
    yield { address, cell: cell as CellObject, ...position };
  }
}

function xlsxCellAddress(row: number, column: number): string {
  let value = column + 1;
  let label = '';
  while (value > 0) {
    value -= 1;
    label = String.fromCharCode(65 + (value % 26)) + label;
    value = Math.floor(value / 26);
  }
  return `${label}${row + 1}`;
}

function decodeXlsxCellAddress(
  address: string,
): { column: number; row: number } | null {
  const match = /^([A-Za-z]+)([1-9][0-9]*)$/.exec(address);
  if (!match) return null;
  let column = 0;
  for (const character of match[1].toUpperCase()) {
    column = column * 26 + character.charCodeAt(0) - 64;
  }
  const row = Number(match[2]);
  if (!Number.isSafeInteger(row) || row <= 0) return null;
  return { column: column - 1, row: row - 1 };
}
