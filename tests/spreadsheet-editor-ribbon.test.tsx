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

test('routes font, vertical alignment, and wrapping through cell formats', () => {
  const formats: Array<{ attribute: string; value: unknown }> = [];
  render(
    <SpreadsheetEditorRibbon
      activeTab="home"
      can={spreadsheetCan()}
      commands={spreadsheetCommands((attribute, value) => {
        formats.push({ attribute, value });
        return true;
      })}
      content={{ type: 'spreadsheet', sheets: [] }}
      gridLinesVisible
      multipleCellsSelected={false}
      panel={null}
      toolbarCell={{ ff: 'Arial', vt: 0, tb: '2' }}
      onTabChange={() => undefined}
      onTogglePanel={() => undefined}
    />,
  );

  const font = screen.getByRole('combobox', { name: '字体' });
  expect(font).toHaveTextContent('Arial');
  fireEvent.click(font);
  const simSun = screen.getByRole('option', { name: '宋体' });
  expect(simSun.querySelector('span')).toHaveAttribute(
    'style',
    'font-family: SimSun, "Songti SC", serif;',
  );
  fireEvent.click(simSun);
  fireEvent.click(screen.getByRole('button', { name: '底端对齐' }));
  fireEvent.click(screen.getByRole('button', { name: '自动换行' }));

  expect(formats).toEqual([
    { attribute: 'ff', value: 'SimSun' },
    { attribute: 'vt', value: 2 },
    { attribute: 'tb', value: '1' },
  ]);
  expect(screen.getByRole('button', { name: '加粗' })).toHaveAttribute(
    'aria-keyshortcuts',
    'Control+B Meta+B',
  );
  expect(screen.getByRole('button', { name: '斜体' })).toHaveAttribute(
    'aria-keyshortcuts',
    'Control+I Meta+I',
  );
  expect(screen.getByRole('button', { name: '下划线' })).toHaveAttribute(
    'aria-keyshortcuts',
    'Control+U Meta+U',
  );
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
    activateSheet: () => true,
    addSheet: () => true,
    clearSelectedCells: () => true,
    deleteSheet: () => true,
    duplicateSheet: () => true,
    hideSheet: () => true,
    moveSheet: () => true,
    pasteCells: () => true,
    recalculateFormula: () => true,
    renameSheet: () => true,
    redo: () => false,
    setCellFormat: () => true,
    setGridLines: () => true,
    setSheetColor: () => true,
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
    activateSheet: () => true,
    addSheet: () => true,
    clearSelectedCells: () => true,
    deleteSheet: () => true,
    duplicateSheet: () => true,
    hideSheet: () => true,
    moveSheet: () => true,
    pasteCells: () => true,
    recalculateFormula: () => true,
    renameSheet: () => true,
    redo: () => false,
    setCellFormat,
    setGridLines: () => true,
    setSheetColor: () => true,
    setSpreadsheetContent: () => true,
    setZoom: () => true,
    toggleCellMerge: () => true,
    undo: () => false,
  };
}
