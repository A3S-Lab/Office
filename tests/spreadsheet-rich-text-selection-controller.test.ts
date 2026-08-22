import { describe, expect, test } from '@rstest/core';
import {
  isSpreadsheetRichTextFormatPointerTarget,
  SpreadsheetRichTextSelectionController,
} from '../src/internal/features/work/editors/spreadsheet-rich-text-selection-controller';
import type { WorkSpreadsheetContent } from '../src/internal/features/work/work-types';

describe('spreadsheet rich-text selection controller', () => {
  test('captures a cross-run cell-editor selection, applies it, and restores it', () => {
    const fixture = editorFixture(
      '<span style="font-weight: bold">Native </span><span>text</span>',
    );
    const content = richTextContent('Native text');
    const first = fixture.editor.querySelectorAll('span')[0]?.firstChild;
    const second = fixture.editor.querySelectorAll('span')[1]?.firstChild;
    if (!first || !second) throw new Error('Expected editor text nodes.');
    select(first, 1, second, 2);
    const controller = new SpreadsheetRichTextSelectionController();

    expect(
      controller.capture({
        content,
        root: fixture.root,
        selection: {
          row: [0, 0],
          column: [0, 0],
          row_focus: 0,
          column_focus: 0,
        },
        sheetId: 'sheet-1',
      }),
    ).toBe(true);
    expect(controller.canApply(content, 'fc', '#abc')).toBe(true);

    let next: WorkSpreadsheetContent | null = null;
    expect(
      controller.apply(content, (value) => (next = value), 'fc', '#abc'),
    ).toBe(true);
    expect(next?.sheets[0]?.data?.[0]?.[0]?.ct?.s).toEqual([
      { bl: 1, v: 'N' },
      { bl: 1, fc: '#aabbcc', v: 'ative ' },
      { fc: '#aabbcc', it: 1, v: 'te' },
      { it: 1, v: 'xt' },
    ]);
    expect(controller.restore()).toBe(true);
    expect(window.getSelection()?.toString()).toBe('ative te');
    expect(document.activeElement).toBe(fixture.editor);
    fixture.root.remove();
  });

  test('reconstructs live typed spans before formatting instead of restoring stale text', () => {
    const fixture = editorFixture(
      '<span style="font-weight: bold; color: rgb(47, 111, 237)">Live </span><span>text</span>',
    );
    const content = richTextContent('Old text');
    const first = fixture.editor.querySelector('span')?.firstChild;
    if (!first) throw new Error('Expected editor text node.');
    select(first, 0, first, 4);
    const controller = new SpreadsheetRichTextSelectionController();
    expect(
      controller.capture({
        content,
        root: fixture.root,
        selection: { row: [0, 0], column: [0, 0] },
        sheetId: 'sheet-1',
      }),
    ).toBe(true);

    let next: WorkSpreadsheetContent | null = null;
    expect(controller.apply(content, (value) => (next = value), 'it', 1)).toBe(
      true,
    );
    const cell = next?.sheets[0]?.data?.[0]?.[0];
    expect(cell?.v).toBe('Live text');
    expect(cell?.ct?.s).toEqual([
      {
        bl: 1,
        cl: 0,
        fc: '#2f6fed',
        fs: 10,
        it: 1,
        un: 0,
        v: 'Live',
      },
      {
        bl: 1,
        cl: 0,
        fc: '#2f6fed',
        fs: 10,
        it: 0,
        un: 0,
        v: ' ',
      },
      { v: 'text' },
    ]);
    fixture.root.remove();
  });

  test('restores the selection onto a remounted formula bar after a controlled update', async () => {
    const root = document.createElement('div');
    root.innerHTML =
      '<div class="fortune-fx-input" contenteditable="true">Native text</div>';
    document.body.append(root);
    const editor = root.querySelector<HTMLElement>('.fortune-fx-input');
    const text = editor?.firstChild;
    if (!editor || !text) throw new Error('Expected a formula bar.');
    select(text, 7, text, 11);
    const content = richTextContent('Native text');
    const controller = new SpreadsheetRichTextSelectionController();
    expect(
      controller.capture({
        content,
        root,
        selection: { row: [0, 0], column: [0, 0] },
        sheetId: 'sheet-1',
      }),
    ).toBe(true);
    expect(controller.apply(content, () => undefined, 'bl', 1)).toBe(true);

    expect(controller.restore()).toBe(true);
    const replacement = editor.cloneNode(true) as HTMLElement;
    editor.replaceWith(replacement);
    await new Promise((resolve) => setTimeout(resolve, 0));

    expect(document.activeElement).toBe(replacement);
    expect(window.getSelection()?.toString()).toBe('text');
    root.remove();
  });

  test('recognizes only font-format controls as snapshot triggers', () => {
    const host = document.createElement('div');
    host.innerHTML = `
      <button data-spreadsheet-rich-text-format="true"><span id="bold">B</span></button>
      <button class="work-spreadsheet-fill-color"><span id="fill">Fill</span></button>
    `;
    expect(
      isSpreadsheetRichTextFormatPointerTarget(host.querySelector('#bold')),
    ).toBe(true);
    expect(
      isSpreadsheetRichTextFormatPointerTarget(host.querySelector('#fill')),
    ).toBe(false);
  });
});

function editorFixture(html: string): {
  editor: HTMLDivElement;
  root: HTMLDivElement;
} {
  const root = document.createElement('div');
  root.innerHTML = `
    <div class="luckysheet-input-box" style="z-index: 20">
      <div class="luckysheet-cell-input" contenteditable="true">${html}</div>
    </div>
  `;
  document.body.append(root);
  const editor = root.querySelector<HTMLDivElement>('.luckysheet-cell-input');
  if (!editor) throw new Error('Expected cell editor.');
  return { editor, root };
}

function richTextContent(text: string): WorkSpreadsheetContent {
  return {
    type: 'spreadsheet',
    sheets: [
      {
        id: 'sheet-1',
        name: 'Sheet 1',
        data: [
          [
            {
              v: text,
              ct: {
                t: 'inlineStr',
                s:
                  text === 'Native text'
                    ? [
                        { bl: 1, v: 'Native ' },
                        { it: 1, v: 'text' },
                      ]
                    : [{ v: text }],
              },
            },
          ],
        ],
      },
    ],
  };
}

function select(
  startNode: Node,
  startOffset: number,
  endNode: Node,
  endOffset: number,
): void {
  const range = document.createRange();
  range.setStart(startNode, startOffset);
  range.setEnd(endNode, endOffset);
  const selection = window.getSelection();
  selection?.removeAllRanges();
  selection?.addRange(range);
}
