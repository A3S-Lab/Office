import { expect, test } from '@rstest/core';
import {
  applyMarkdownSourceCommand,
  type MarkdownSourceSelection,
} from '../src/internal/features/work/editors/markdown-source-commands';

const selection = (start: number, end = start): MarkdownSourceSelection => ({
  start,
  end,
  direction: 'none',
});

test('wraps and unwraps inline Markdown source selections', () => {
  expect(
    applyMarkdownSourceCommand('Write clearly', selection(0, 5), 'bold'),
  ).toEqual({
    markdown: '**Write** clearly',
    selection: selection(2, 7),
  });

  expect(
    applyMarkdownSourceCommand('**Write** clearly', selection(2, 7), 'bold'),
  ).toEqual({
    markdown: 'Write clearly',
    selection: selection(0, 5),
  });

  expect(
    applyMarkdownSourceCommand('Ship today', selection(5, 5), 'italic'),
  ).toEqual({
    markdown: 'Ship **today',
    selection: selection(6, 6),
  });
});

test('formats every selected source line without losing the selection', () => {
  expect(
    applyMarkdownSourceCommand(
      'First item\nSecond item',
      selection(0, 22),
      'bullet-list',
    ),
  ).toEqual({
    markdown: '- First item\n- Second item',
    selection: selection(0, 26),
  });

  expect(
    applyMarkdownSourceCommand(
      '- First item\n- Second item',
      selection(0, 26),
      'bullet-list',
    ),
  ).toEqual({
    markdown: 'First item\nSecond item',
    selection: selection(0, 22),
  });

  expect(
    applyMarkdownSourceCommand(
      '- Existing item',
      selection(2, 15),
      'heading-2',
    ),
  ).toEqual({
    markdown: '## Existing item',
    selection: selection(0, 16),
  });

  expect(
    applyMarkdownSourceCommand('\nSecond line', selection(0), 'heading-1'),
  ).toEqual({
    markdown: '# \nSecond line',
    selection: selection(0, 2),
  });
});

test('inserts block content at a source caret with a useful selection', () => {
  expect(
    applyMarkdownSourceCommand('Before', selection(6), 'horizontal-rule'),
  ).toEqual({
    markdown: 'Before\n\n---\n\n',
    selection: selection(13),
  });

  const table = applyMarkdownSourceCommand('', selection(0), 'table');
  expect(table.markdown).toBe(
    '| Column 1 | Column 2 | Column 3 |\n| --- | --- | --- |\n|  |  |  |\n|  |  |  |',
  );
  expect(table.selection.start).toBe(2);
  expect(table.selection.end).toBe(10);
});
