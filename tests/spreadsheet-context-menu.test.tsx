import { expect, test } from '@rstest/core';
import { waitFor } from '@testing-library/react';
import {
  parseSpreadsheetClipboardText,
  readSpreadsheetClipboardText,
  spreadsheetCoreContextMenuItems,
  writeSpreadsheetClipboardText,
} from '../src/internal/features/work/editors/spreadsheet-context-menu';

test('parses tabular clipboard text without adding a trailing empty row', () => {
  expect(
    parseSpreadsheetClipboardText('项目\t金额\r\n研发\t120\r\n市场\t\r\n'),
  ).toEqual([
    ['项目', '金额'],
    ['研发', '120'],
    ['市场', ''],
  ]);
});

test('falls back when the system clipboard is blocked or does not settle', async () => {
  const localWrites: string[] = [];
  const systemReads: string[] = [];

  await expect(
    readSpreadsheetClipboardText(
      () => new Promise<string>(() => undefined),
      () => 'A3S\tOffice',
      1,
    ),
  ).resolves.toBe('A3S\tOffice');
  await expect(
    readSpreadsheetClipboardText(
      async () => {
        systemReads.push('read');
        return 'system';
      },
      () => 'local',
      10,
      () => false,
    ),
  ).resolves.toBe('local');
  await expect(
    writeSpreadsheetClipboardText(
      'A3S\tOffice',
      () => new Promise<void>(() => undefined),
      (value) => localWrites.push(value),
      1,
    ),
  ).resolves.toBeUndefined();
  expect(localWrites).toEqual(['A3S\tOffice']);
  expect(systemReads).toEqual([]);
});

test('puts standard spreadsheet editing actions before optional AI actions', async () => {
  const calls: string[] = [];
  const clipboard = {
    readText: async () => '项目\t金额\n研发\t120',
    writeText: async (value: string) => {
      calls.push(`copy:${value}`);
    },
  };
  const items = spreadsheetCoreContextMenuItems({
    can: {
      clearSelectedCells: () => true,
      pasteCells: () => true,
    },
    clipboard,
    commands: {
      clearSelectedCells: () => {
        calls.push('clear');
        return true;
      },
      pasteCells: (values) => {
        calls.push(`paste:${JSON.stringify(values)}`);
        return true;
      },
    },
    selection: {
      clipboard: 'A3S\tOffice',
      reference: 'A1:B1',
    },
  });

  expect(items.map(({ label }) => label)).toEqual([
    '剪切',
    '复制',
    '粘贴',
    '清除内容',
  ]);

  items[1].onSelect();
  await waitFor(() => expect(calls).toEqual(['copy:A3S\tOffice']));

  items[0].onSelect();
  await waitFor(() =>
    expect(calls).toEqual(['copy:A3S\tOffice', 'copy:A3S\tOffice', 'clear']),
  );

  items[2].onSelect();
  await waitFor(() =>
    expect(calls).toEqual([
      'copy:A3S\tOffice',
      'copy:A3S\tOffice',
      'clear',
      'paste:[["项目","金额"],["研发","120"]]',
    ]),
  );

  items[3].onSelect();
  expect(calls).toEqual([
    'copy:A3S\tOffice',
    'copy:A3S\tOffice',
    'clear',
    'paste:[["项目","金额"],["研发","120"]]',
    'clear',
  ]);
});
