import type { SpreadsheetKeyboardSelection } from './spreadsheet-keyboard-navigation';

/**
 * Fortune clears its live selection after deleting rows or columns. Keep the
 * user's selection on the structure that moved into the deleted range while
 * clamping terminal deletions to the new grid boundary.
 */
export function spreadsheetSelectionAfterStructureDeletion(
  selection: SpreadsheetKeyboardSelection,
  axis: 'column' | 'row',
  extent: number,
  start: number,
  end: number,
): SpreadsheetKeyboardSelection {
  const deletedCount = Math.max(1, end - start + 1);
  const remainingExtent = Math.max(1, extent - deletedCount);
  const focus = Math.min(Math.max(0, start), remainingExtent - 1);

  if (axis === 'row') {
    return {
      ...selection,
      row: [focus, focus],
      column: [...selection.column],
      ...(selection.row_focus === undefined ? {} : { row_focus: focus }),
    };
  }
  return {
    ...selection,
    row: [...selection.row],
    column: [focus, focus],
    ...(selection.column_focus === undefined ? {} : { column_focus: focus }),
  };
}
