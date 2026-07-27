import { expect, test } from '@rstest/core';
import { render, screen } from '@testing-library/react';
import { SpreadsheetConditionalFormatPanel } from '../src/internal/features/work/editors/spreadsheet-conditional-format-panel';

test('uses product language in the empty conditional-format state', () => {
  render(
    <SpreadsheetConditionalFormatPanel
      content={{
        type: 'spreadsheet',
        sheets: [{ id: 'sheet-1', name: '工作表 1', data: [] }],
      }}
      onChange={() => undefined}
    />,
  );

  expect(screen.getByLabelText('条件格式规则')).toBeInTheDocument();
  expect(screen.getByRole('status')).toHaveTextContent('还没有条件格式规则。');
  expect(screen.queryByText(/Work/)).toBeNull();
});
