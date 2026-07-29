import { expect, test } from '@rstest/core';
import {
  applyMarkdownSourceCommand,
  createMarkdownImageSourceInsert,
  createMarkdownLinkSourceInsert,
  createMarkdownSourceLinkRemoval,
  findMarkdownSourceLink,
  isMarkdownSourceCommandActive,
  replaceMarkdownSourceSelection,
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

test('escapes Markdown link and image inserts without changing their meaning', () => {
  expect(
    createMarkdownLinkSourceInsert(
      'A3S [Office]',
      'https://a3s.dev/docs_(next)',
    ),
  ).toEqual({
    replacement: '[A3S \\[Office\\]](https://a3s.dev/docs_\\(next\\))',
    selectedRange: { start: 1, end: 15 },
  });
  expect(
    createMarkdownImageSourceInsert(
      '架构图 [新版]',
      '../assets/office diagram.png',
    ),
  ).toEqual({
    replacement: '![架构图 \\[新版\\]](<../assets/office diagram.png>)',
    selectedRange: { start: 2, end: 12 },
  });
});

test('finds, edits, and removes the inline link targeted by a source selection', () => {
  const markdown =
    'Before [A3S \\[Office\\]](https://a3s.dev/docs_\\(next\\) "Docs") after';
  const linkStart = markdown.indexOf('[A3S');
  const linkEnd = markdown.indexOf(') after') + 1;
  const link = findMarkdownSourceLink(
    markdown,
    selection(markdown.indexOf('Office') + 2),
  );

  expect(link).toEqual({
    label: 'A3S [Office]',
    rawLabel: 'A3S \\[Office\\]',
    range: { start: linkStart, end: linkEnd },
    source: 'https://a3s.dev/docs_(next)',
  });
  if (!link) throw new Error('Expected the selected source link.');

  const edited = createMarkdownLinkSourceInsert(
    'Office docs',
    'https://a3s.dev/office',
  );
  expect(
    replaceMarkdownSourceSelection(
      markdown,
      selection(link.range.start, link.range.end),
      edited.replacement,
      edited.selectedRange,
    ),
  ).toEqual({
    markdown: 'Before [Office docs](https://a3s.dev/office) after',
    selection: selection(linkStart + 1, linkStart + 12),
  });

  const removal = createMarkdownSourceLinkRemoval(link);
  expect(
    replaceMarkdownSourceSelection(
      markdown,
      selection(link.range.start, link.range.end),
      removal.replacement,
      removal.selectedRange,
    ),
  ).toEqual({
    markdown: 'Before A3S \\[Office\\] after',
    selection: selection(linkStart, linkStart + link.rawLabel.length),
  });
});

test('does not treat images, escaped syntax, or adjacent text as an active source link', () => {
  const image = '![Diagram](https://a3s.dev/diagram.png)';
  const escaped = '\\[Literal](https://a3s.dev/literal)';
  const link = '[Office](https://a3s.dev) after';

  expect(findMarkdownSourceLink(image, selection(4))).toBeNull();
  expect(findMarkdownSourceLink(escaped, selection(3))).toBeNull();
  expect(
    findMarkdownSourceLink(link, selection(link.indexOf('Office') + 2)),
  ).not.toBeNull();
  expect(
    findMarkdownSourceLink(link, selection(link.indexOf(' after') + 1)),
  ).toBeNull();
});

test('derives source toolbar state from the selected Markdown structure', () => {
  const markdown = '## Plan\n\n**Bold** and *italic*\n\n- [ ] Ship';
  const boldStart = markdown.indexOf('Bold');

  expect(
    isMarkdownSourceCommandActive(markdown, selection(1), 'heading-2'),
  ).toBe(true);
  expect(
    isMarkdownSourceCommandActive(
      markdown,
      selection(boldStart, boldStart + 'Bold'.length),
      'bold',
    ),
  ).toBe(true);
  expect(
    isMarkdownSourceCommandActive(
      markdown,
      selection(boldStart, boldStart + 'Bold'.length),
      'italic',
    ),
  ).toBe(false);
  expect(
    isMarkdownSourceCommandActive(
      markdown,
      selection(markdown.indexOf('Ship')),
      'task-list',
    ),
  ).toBe(true);
});

test('recognizes and removes an inline wrapper around a source caret or partial selection', () => {
  const markdown = '**Write clearly** today';
  const caret = selection(markdown.indexOf('clearly') + 2);
  const partial = selection(
    markdown.indexOf('clearly'),
    markdown.indexOf('clearly') + 'clear'.length,
  );

  expect(isMarkdownSourceCommandActive(markdown, caret, 'bold')).toBe(true);
  expect(isMarkdownSourceCommandActive(markdown, partial, 'bold')).toBe(true);
  expect(applyMarkdownSourceCommand(markdown, caret, 'bold')).toEqual({
    markdown: 'Write clearly today',
    selection: selection(caret.start - 2),
  });
  expect(applyMarkdownSourceCommand(markdown, partial, 'bold')).toEqual({
    markdown: 'Write clearly today',
    selection: selection(partial.start - 2, partial.end - 2),
  });
});
