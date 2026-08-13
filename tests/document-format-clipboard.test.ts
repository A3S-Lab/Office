import { afterEach, expect, test } from '@rstest/core';
import { Editor } from '@tiptap/core';
import {
  clearDocumentFormatClipboard,
  copyDocumentFormatting,
  hasDocumentFormatClipboard,
  pasteDocumentFormatting,
} from '../src/internal/features/work/editors/document-format-clipboard';
import { createDocumentPageChromeEditorExtensions } from '../src/internal/features/work/editors/document-page-chrome-editor';
import { createWorkDocumentExtensions } from '../src/internal/features/work/work-document-extensions';

let editor: Editor | null = null;

afterEach(() => {
  editor?.destroy();
  editor = null;
  clearDocumentFormatClipboard();
});

test('copies character and paragraph formatting without replacing semantic marks', () => {
  editor = createEditor();
  editor.commands.setTextSelection(textRange(editor, 'Source'));

  expect(copyDocumentFormatting(editor)).toBe(true);
  expect(hasDocumentFormatClipboard()).toBe(true);

  editor.commands.setTextSelection(textRange(editor, 'Target'));
  expect(pasteDocumentFormatting(editor)).toBe(true);

  expect(editor.isActive('heading', { level: 2 })).toBe(true);
  expect(editor.getAttributes('heading')).toMatchObject({
    lineHeight: '1.5',
    paragraphShading: {
      pattern: 'pct20',
      color: { value: '#112233' },
      fill: { value: '#ddeeff' },
    },
    textAlign: 'center',
  });
  expect(editor.isActive('bold')).toBe(true);
  expect(editor.isActive('italic')).toBe(false);
  expect(editor.isActive('link', { href: 'https://a3s.dev' })).toBe(true);
  expect(editor.getAttributes('textStyle')).toMatchObject({
    color: '#0070c0',
    fontFamily: 'Arial, sans-serif',
    fontSize: '14pt',
  });

  expect(editor.commands.undo()).toBe(true);
  editor.commands.setTextSelection(textRange(editor, 'Target'));
  expect(editor.isActive('paragraph')).toBe(true);
  expect(editor.isActive('italic')).toBe(true);
  expect(editor.isActive('bold')).toBe(false);
  expect(editor.isActive('link', { href: 'https://a3s.dev' })).toBe(true);
});

test('applies copied character formatting to text typed at a collapsed cursor', () => {
  editor = createEditor();
  editor.commands.setTextSelection(textRange(editor, 'Source'));
  expect(copyDocumentFormatting(editor)).toBe(true);

  const target = textRange(editor, 'Target');
  editor.commands.setTextSelection(target.to);
  expect(pasteDocumentFormatting(editor)).toBe(true);
  expect(
    new Set(editor.state.storedMarks?.map((mark) => mark.type.name)),
  ).toEqual(new Set(['link', 'bold', 'textStyle']));
  editor.commands.insertContent(' typed');

  editor.commands.setTextSelection(textRange(editor, ' typed'));
  expect(editor.isActive('bold')).toBe(true);
  expect(editor.isActive('italic')).toBe(false);
  expect(editor.getAttributes('textStyle')).toMatchObject({
    color: '#0070c0',
    fontFamily: 'Arial, sans-serif',
    fontSize: '14pt',
  });
});

test('keeps a required list paragraph when the copied block type is a heading', () => {
  editor = new Editor({
    extensions: createWorkDocumentExtensions(),
    content: [
      '<h2><strong>Source</strong></h2>',
      '<ul><li><p><em>List target</em></p></li></ul>',
    ].join(''),
  });
  editor.commands.setTextSelection(textRange(editor, 'Source'));
  expect(copyDocumentFormatting(editor)).toBe(true);

  editor.commands.setTextSelection(textRange(editor, 'List target'));
  expect(pasteDocumentFormatting(editor)).toBe(true);

  expect(editor.isActive('bulletList')).toBe(true);
  expect(editor.isActive('listItem')).toBe(true);
  expect(editor.isActive('paragraph')).toBe(true);
  expect(editor.isActive('heading')).toBe(false);
  expect(editor.isActive('bold')).toBe(true);
  expect(editor.isActive('italic')).toBe(false);
});

test('applies compatible body formatting when the target omits heading nodes', () => {
  editor = createEditor();
  editor.commands.setTextSelection(textRange(editor, 'Source'));
  expect(copyDocumentFormatting(editor)).toBe(true);
  const pageChromeEditor = new Editor({
    extensions: createDocumentPageChromeEditorExtensions(),
    content: '<p><em>Header target</em></p>',
  });
  pageChromeEditor.commands.setTextSelection(
    textRange(pageChromeEditor, 'Header target'),
  );

  expect(pasteDocumentFormatting(pageChromeEditor)).toBe(true);
  expect(pageChromeEditor.isActive('paragraph')).toBe(true);
  expect(pageChromeEditor.isActive('bold')).toBe(true);
  expect(pageChromeEditor.isActive('italic')).toBe(false);
  expect(pageChromeEditor.getAttributes('paragraph').textAlign).toBe('center');
  expect(pageChromeEditor.getAttributes('paragraph').paragraphShading).toEqual({
    pattern: 'pct20',
    color: { value: '#112233' },
    fill: { value: '#ddeeff' },
  });

  pageChromeEditor.destroy();
});

test('filters unsupported body marks at a page-chrome cursor', () => {
  editor = new Editor({
    extensions: createWorkDocumentExtensions(),
    content: '<p><mark data-color="#fff2cc"><strong>Source</strong></mark></p>',
  });
  editor.commands.setTextSelection(textRange(editor, 'Source'));
  expect(copyDocumentFormatting(editor)).toBe(true);
  const pageChromeEditor = new Editor({
    extensions: createDocumentPageChromeEditorExtensions(),
    content: '<p>Header</p>',
  });
  pageChromeEditor.commands.setTextSelection(
    textRange(pageChromeEditor, 'Header').to,
  );

  expect(pasteDocumentFormatting(pageChromeEditor)).toBe(true);
  pageChromeEditor.commands.insertContent(' typed');
  pageChromeEditor.commands.setTextSelection(
    textRange(pageChromeEditor, ' typed'),
  );
  expect(pageChromeEditor.isActive('bold')).toBe(true);
  expect(pageChromeEditor.schema.marks.highlight).toBeUndefined();

  pageChromeEditor.destroy();
});

function createEditor(): Editor {
  return new Editor({
    extensions: createWorkDocumentExtensions(),
    content: [
      `<h2 data-office-paragraph-shading='{"pattern":"pct20","color":{"value":"#112233"},"fill":{"value":"#ddeeff"}}' style="text-align: center; line-height: 1.5; background-color: #ddeeff">`,
      '<strong><span style="color: #0070c0; font-family: Arial, sans-serif; font-size: 14pt">Source</span></strong>',
      '</h2>',
      '<p><a href="https://a3s.dev"><em>Target</em></a></p>',
    ].join(''),
  });
}

function textRange(
  currentEditor: Editor,
  text: string,
): { from: number; to: number } {
  let range: { from: number; to: number } | null = null;
  currentEditor.state.doc.descendants((node, position) => {
    if (range || !node.isText || !node.text) return;
    const offset = node.text.indexOf(text);
    if (offset < 0) return;
    range = {
      from: position + offset,
      to: position + offset + text.length,
    };
  });
  if (!range) throw new Error(`Text "${text}" was not found.`);
  return range;
}
