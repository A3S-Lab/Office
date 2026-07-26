import { expect, test } from '@rstest/core';
import { fireEvent, render, screen } from '@testing-library/react';
import type {
  SpreadsheetEditorCanCommands,
  SpreadsheetEditorCommands,
} from '../src/internal/features/work/editors/spreadsheet-command-controller';
import { SpreadsheetEditorRibbon } from '../src/internal/features/work/editors/spreadsheet-editor-ribbon';

test('routes number-format controls through typed spreadsheet commands', () => {
  const formats: Array<{ attribute: string; value: unknown }> = [];
  const commands = spreadsheetCommands((attribute, value) => {
    formats.push({ attribute, value });
    return true;
  });

  render(
    <SpreadsheetEditorRibbon
      activeTab="home"
      can={spreadsheetCan()}
      commands={commands}
      content={{ type: 'spreadsheet', sheets: [] }}
      gridLinesVisible
      multipleCellsSelected={false}
      panel={null}
      toolbarCell={{ v: 0.5, m: '50.0%', ct: { fa: '0.0%', t: 'n' } }}
      onTabChange={() => undefined}
      onTogglePanel={() => undefined}
    />,
  );

  const formatSelect = screen.getByRole('combobox', { name: '数字格式' });
  expect(formatSelect).toHaveTextContent('百分比');
  fireEvent.click(formatSelect);
  fireEvent.click(screen.getByRole('option', { name: '数字' }));
  fireEvent.click(screen.getByRole('button', { name: '百分比格式' }));
  fireEvent.click(screen.getByRole('button', { name: '减少小数位' }));
  fireEvent.click(screen.getByRole('button', { name: '增加小数位' }));

  expect(formats).toEqual([
    { attribute: 'ct', value: { fa: '#,##0.00', t: 'n' } },
    { attribute: 'ct', value: { fa: '0.00%', t: 'n' } },
    { attribute: 'ct', value: { fa: '0%', t: 'n' } },
    { attribute: 'ct', value: { fa: '0.00%', t: 'n' } },
  ]);
});

test('omits empty resource counts from spreadsheet ribbon actions', () => {
  render(
    <SpreadsheetEditorRibbon
      activeTab="insert"
      can={spreadsheetCan()}
      commands={spreadsheetCommands(() => true)}
      content={{ type: 'spreadsheet', sheets: [] }}
      gridLinesVisible
      multipleCellsSelected={false}
      panel={null}
      toolbarCell={null}
      onTabChange={() => undefined}
      onTogglePanel={() => undefined}
    />,
  );

  expect(screen.getByRole('button', { name: '插入图表' })).toBeInTheDocument();
  expect(screen.getByRole('button', { name: '条件格式' })).toBeInTheDocument();
  expect(screen.queryByText('0')).not.toBeInTheDocument();
});

function spreadsheetCan(): SpreadsheetEditorCanCommands {
  return {
    recalculateFormula: () => true,
    redo: () => false,
    setCellFormat: () => true,
    setGridLines: () => true,
    setSpreadsheetContent: () => true,
    setZoom: () => true,
    toggleCellMerge: () => true,
    undo: () => false,
  };
}

function spreadsheetCommands(
  setCellFormat: SpreadsheetEditorCommands['setCellFormat'],
): SpreadsheetEditorCommands {
  return {
    recalculateFormula: () => true,
    redo: () => false,
    setCellFormat,
    setGridLines: () => true,
    setSpreadsheetContent: () => true,
    setZoom: () => true,
    toggleCellMerge: () => true,
    undo: () => false,
  };
}
