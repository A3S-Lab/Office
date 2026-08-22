import type { Cell } from '@fortune-sheet/core';
import { describe, expect, test } from '@rstest/core';
import {
  captureSpreadsheetRichTextPaste,
  parseSpreadsheetRichTextClipboard,
  stageSpreadsheetRichTextPaste,
  takeSpreadsheetRichTextPaste,
} from '../src/internal/features/work/editors/spreadsheet-rich-text-paste';
import type { WorkSpreadsheetSheet } from '../src/internal/features/work/work-types';

describe('Spreadsheet formatted rich-text paste', () => {
  test('parses bounded inline Office formatting without retaining clipboard markup', () => {
    expect(
      parseSpreadsheetRichTextClipboard(
        '<span style="font-family: Georgia; font-size: 16pt; color: rgb(192, 0, 0); font-weight: 700">Red</span><em><u> ink</u></em>',
        'Red ink',
      ),
    ).toEqual({
      runs: [
        {
          bl: 1,
          fc: '#c00000',
          ff: 'Georgia',
          fs: 16,
          v: 'Red',
        },
        { it: 1, un: 1, v: ' ink' },
      ],
      text: 'Red ink',
    });
  });

  test('supports explicit resets and Fortune underline metadata', () => {
    expect(
      parseSpreadsheetRichTextClipboard(
        '<strong><span style="font-weight: normal; lucky-underline: 2">Plain</span><span style="text-decoration: line-through"> cut</span></strong>',
        'Plain cut',
      ),
    ).toEqual({
      runs: [
        { bl: 0, un: 2, v: 'Plain' },
        { bl: 1, cl: 1, v: ' cut' },
      ],
      text: 'Plain cut',
    });
  });

  test('rejects unformatted, text-mismatched, malformed, and oversized clipboard payloads', () => {
    expect(
      parseSpreadsheetRichTextClipboard('<span>Plain</span>', 'Plain'),
    ).toBeNull();
    expect(
      parseSpreadsheetRichTextClipboard('<strong>Rich</strong>', 'Other'),
    ).toBeNull();
    expect(
      parseSpreadsheetRichTextClipboard(
        `<strong>${'x'.repeat(32_768)}</strong>`,
        'x'.repeat(32_768),
      ),
    ).toBeNull();
    expect(
      parseSpreadsheetRichTextClipboard(
        `<strong>${'x'.repeat(260_000)}</strong>`,
        'x',
      ),
    ).toBeNull();
  });

  test('keeps a staged intent through text-stable callbacks and consumes it once on change', () => {
    const sheet = worksheet();
    const intent = {
      end: 6,
      runs: [{ bl: 1 as const, v: 'styled' }],
      start: 0,
      text: 'styled',
    };

    stageSpreadsheetRichTextPaste(sheet, 0, 0, 'Source', intent);
    expect(
      takeSpreadsheetRichTextPaste(sheet, 0, 0, {
        ct: { t: 'inlineStr', s: [{ v: 'Source' }] },
      }),
    ).toBeUndefined();
    expect(takeSpreadsheetRichTextPaste(sheet, 0, 0, { v: 'styled' })).toEqual(
      intent,
    );
    expect(
      takeSpreadsheetRichTextPaste(sheet, 0, 0, { v: 'styled' }),
    ).toBeUndefined();
  });

  test('captures the exact formula-bar selection as a pending formatted paste', () => {
    const sheet = worksheet();
    const content = { sheets: [sheet], type: 'spreadsheet' as const };
    const root = document.createElement('section');
    const editor = document.createElement('div');
    editor.className = 'fortune-fx-input';
    editor.contentEditable = 'true';
    editor.textContent = 'Source';
    root.append(editor);
    document.body.append(root);
    const text = editor.firstChild;
    if (!text) throw new Error('Expected formula-bar text.');
    const range = document.createRange();
    range.setStart(text, 0);
    range.setEnd(text, 6);
    const browserSelection = window.getSelection();
    browserSelection?.removeAllRanges();
    browserSelection?.addRange(range);

    expect(
      captureSpreadsheetRichTextPaste({
        clipboardData: {
          getData: (type) =>
            type === 'text/html' ? '<strong>Styled</strong>' : 'Styled',
        },
        content,
        root,
        selection: {
          column: [0, 0],
          column_focus: 0,
          row: [0, 0],
          row_focus: 0,
        },
        sheetId: 'sheet-1',
        target: editor,
      }),
    ).toBe(true);
    expect(takeSpreadsheetRichTextPaste(sheet, 0, 0, { v: 'Styled' })).toEqual({
      end: 6,
      runs: [{ bl: 1, v: 'Styled' }],
      start: 0,
      text: 'Styled',
    });

    browserSelection?.removeAllRanges();
    root.remove();
  });

  test('isolates staged paste intents by source worksheet and coordinate', () => {
    const first = worksheet();
    const second = worksheet();
    const intent = {
      end: 0,
      runs: [{ it: 1 as const, v: 'X' }],
      start: 0,
      text: 'X',
    };
    stageSpreadsheetRichTextPaste(first, 2, 3, '', intent);

    expect(
      takeSpreadsheetRichTextPaste(first, 2, 4, { v: 'X' }),
    ).toBeUndefined();
    expect(
      takeSpreadsheetRichTextPaste(second, 2, 3, { v: 'X' }),
    ).toBeUndefined();
    expect(takeSpreadsheetRichTextPaste(first, 2, 3, { v: 'X' })).toEqual(
      intent,
    );
  });
});

function worksheet(): WorkSpreadsheetSheet {
  return {
    column: 8,
    data: [[{ v: 'Source' } satisfies Cell]],
    id: 'sheet-1',
    name: 'Sheet 1',
    row: 8,
  };
}
