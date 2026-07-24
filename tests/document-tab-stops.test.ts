import { Editor } from '@tiptap/core';
import { describe, expect, test } from '@rstest/core';
import { createWorkDocumentExtensions } from '../src/internal/features/work/work-document-extensions';
import {
  documentParagraphTabStops,
  normalizeDocumentTabStops,
  setDocumentParagraphTabStops,
} from '../src/internal/features/work/work-document-tab-stops';

function createEditor(content = '<p>A3S Office</p>'): Editor {
  return new Editor({
    extensions: createWorkDocumentExtensions(),
    content,
  });
}

function pressTab(editor: Editor): boolean {
  return editor.view.dom.dispatchEvent(
    new KeyboardEvent('keydown', {
      key: 'Tab',
      code: 'Tab',
      bubbles: true,
      cancelable: true,
    }),
  );
}

describe('document tab stops', () => {
  test('normalizes, sorts, and de-duplicates typed paragraph tab stops', () => {
    expect(
      normalizeDocumentTabStops([
        { position: 144, alignment: 'right', leader: 'dot' },
        { position: -12, alignment: 'left', leader: 'none' },
        { position: 48, alignment: 'invalid', leader: 'hyphen' },
        { position: 144, alignment: 'decimal', leader: 'underscore' },
        { position: Number.POSITIVE_INFINITY, alignment: 'left' },
      ]),
    ).toEqual([
      { position: 48, alignment: 'left', leader: 'hyphen' },
      { position: 144, alignment: 'decimal', leader: 'underscore' },
    ]);
  });

  test('keeps tab stops and inline tabs in the structured TipTap document', () => {
    const editor = createEditor();

    expect(
      setDocumentParagraphTabStops(editor, [
        { position: 48, alignment: 'left', leader: 'none' },
        { position: 144, alignment: 'right', leader: 'dot' },
      ]),
    ).toBe(true);
    editor.commands.setTextSelection(4);
    expect(pressTab(editor)).toBe(false);

    expect(documentParagraphTabStops(editor)).toEqual([
      { position: 48, alignment: 'left', leader: 'none' },
      { position: 144, alignment: 'right', leader: 'dot' },
    ]);
    expect(editor.getHTML()).toContain('data-office-tab-stops=');
    expect(editor.getHTML()).toContain('data-document-tab="true"');
    expect(JSON.stringify(editor.getJSON())).toContain('"type":"documentTab"');

    editor.destroy();
  });

  test('leaves Tab navigation inside table cells to the table extension', () => {
    const editor = createEditor(
      [
        '<table><tbody><tr>',
        '<td><p>First</p></td><td><p>Second</p></td>',
        '</tr></tbody></table>',
      ].join(''),
    );
    editor.commands.setTextSelection(4);

    pressTab(editor);

    expect(JSON.stringify(editor.getJSON())).not.toContain(
      '"type":"documentTab"',
    );
    editor.destroy();
  });

  test('leaves Tab indentation inside lists to the list extension', () => {
    const editor = createEditor(
      '<ol><li><p>First</p></li><li><p>Second</p></li></ol>',
    );
    editor.commands.setTextSelection(textPosition(editor, 'Second'));

    pressTab(editor);

    expect(editor.getHTML()).toContain(
      '<li><p>First</p><ol><li><p>Second</p></li></ol></li>',
    );
    expect(JSON.stringify(editor.getJSON())).not.toContain(
      '"type":"documentTab"',
    );
    editor.destroy();
  });
});

function textPosition(editor: Editor, value: string): number {
  let result = -1;
  editor.state.doc.descendants((node, position) => {
    if (result < 0 && node.isText && node.text === value) {
      result = position + 1;
    }
  });
  if (result < 0) throw new Error(`Expected document text: ${value}`);
  return result;
}
