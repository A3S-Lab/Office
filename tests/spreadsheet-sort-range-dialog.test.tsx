import { expect, test } from '@rstest/core';
import { fireEvent, render, screen } from '@testing-library/react';
import type { SpreadsheetSortRangeDialogSource } from '../src/internal/features/work/editors/spreadsheet-sort';
import { SpreadsheetSortRangeDialog } from '../src/internal/features/work/editors/spreadsheet-sort-range-dialog';

test('chooses between the expanded WPS current region and the exact selection', () => {
  const choices: string[] = [];
  render(
    <SpreadsheetSortRangeDialog
      source={rangeSource()}
      restoreFocusTarget={() => null}
      onApply={(choice) => {
        choices.push(choice);
        return true;
      }}
      onClose={() => undefined}
    />,
  );

  expect(screen.getByRole('dialog', { name: '排序提醒' })).toHaveTextContent(
    '执行看板',
  );
  expect(screen.getByRole('radio', { name: /扩展选定区域/ })).toBeChecked();
  expect(screen.getByText('A3:G7')).toBeVisible();
  expect(screen.getByText('B4:B7')).toBeVisible();

  fireEvent.click(screen.getByRole('radio', { name: /以当前选定区域排序/ }));
  fireEvent.click(screen.getByRole('button', { name: '排序' }));
  expect(choices).toEqual(['selection']);
});

test('disables a candidate that cannot be sorted safely', () => {
  render(
    <SpreadsheetSortRangeDialog
      source={{
        ...rangeSource(),
        canSortSelection: false,
      }}
      restoreFocusTarget={() => null}
      onApply={() => true}
      onClose={() => undefined}
    />,
  );

  expect(
    screen.getByRole('radio', { name: /以当前选定区域排序/ }),
  ).toBeDisabled();
  expect(screen.getByRole('radio', { name: /扩展选定区域/ })).toBeChecked();
});

test('explains the locked table body instead of generic adjacent data', () => {
  render(
    <SpreadsheetSortRangeDialog
      source={{
        ...rangeSource(),
        canSortSelection: false,
        ownedScope: { kind: 'table', tableId: 'table-1', hasHeader: true },
      }}
      restoreFocusTarget={() => null}
      onApply={() => true}
      onClose={() => undefined}
    />,
  );

  expect(screen.getByRole('dialog', { name: '排序提醒' })).toHaveTextContent(
    '选定单元格位于表格中',
  );
  expect(
    screen.getByRole('radio', { name: /对整个表格数据区域排序/ }),
  ).toBeChecked();
  expect(screen.getByText(/表头与汇总行保持固定/)).toBeVisible();
});

function rangeSource(): SpreadsheetSortRangeDialogSource {
  return {
    sheetName: '执行看板',
    selectedRangeReference: 'B4:B7',
    expandedRangeReference: 'A3:G7',
    canSortSelection: true,
    canSortExpandedRange: true,
  };
}
