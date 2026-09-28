import { spreadsheetCellAddress } from '../work-spreadsheet-formulas';

/**
 * Screen-reader announcement for the active spreadsheet cell.
 * Kept separate from toolbar chrome so selection changes stay audible even
 * when format-painter / filter / freeze status is empty.
 */
export function spreadsheetActiveCellAnnouncement(options: {
  row: number | undefined;
  column: number | undefined;
  displayValue?: string | null;
  formula?: string | null;
}): string {
  if (
    options.row === undefined ||
    options.column === undefined ||
    !Number.isFinite(options.row) ||
    !Number.isFinite(options.column)
  ) {
    return '';
  }
  const address = spreadsheetCellAddress(options.row, options.column);
  const formula = options.formula?.trim();
  if (formula) {
    return `${address}, formula ${formula}`;
  }
  const value = options.displayValue?.trim();
  if (value) {
    return `${address}, ${value}`;
  }
  return `${address}, blank`;
}
