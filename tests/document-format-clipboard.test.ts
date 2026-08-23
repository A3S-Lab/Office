import { afterEach, expect, test } from '@rstest/core';
import { Editor } from '@tiptap/core';
import {
  clearDocumentFormatClipboard,
  copyDocumentFormatting,
  hasDocumentFormatClipboard,
  pasteDocumentFormatting,
} from '../src/internal/features/work/editors/document-format-clipboard';
import { createDocumentPageChromeEditorExtensions } from '../src/internal/features/work/editors/document-page-chrome-editor';
import {
  documentScriptFontsDomAttributes,
  normalizeDocumentScriptFonts,
  parseDocumentScriptFonts,
} from '../src/internal/features/work/work-document-script-fonts';
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
    paragraphBorders: {
      bottom: {
        style: 'double',
        color: { value: '#4472c4' },
        size: 12,
      },
    },
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
    characterScalePercent: 80,
    characterPositionHalfPoints: 4,
    characterSpacingTwips: 30,
    emphasisMark: 'underDot',
    kerningThresholdHalfPoints: 24,
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
    characterScalePercent: 80,
    characterPositionHalfPoints: 4,
    characterSpacingTwips: 30,
    emphasisMark: 'underDot',
    kerningThresholdHalfPoints: 24,
    color: '#0070c0',
    fontFamily: 'Arial, sans-serif',
    fontSize: '14pt',
  });
});

test('copies native underline style and color as one formatting mark', () => {
  editor = new Editor({
    extensions: createWorkDocumentExtensions(),
    content:
      '<p><u data-office-underline-style="wavyDouble" data-office-underline-color="#4472c4">Source</u> and <em>Target</em></p>',
  });
  editor.commands.setTextSelection(textRange(editor, 'Source'));
  expect(copyDocumentFormatting(editor)).toBe(true);

  editor.commands.setTextSelection(textRange(editor, 'Target'));
  expect(pasteDocumentFormatting(editor)).toBe(true);
  expect(editor.getAttributes('underline')).toMatchObject({
    underlineColor: '#4472c4',
    underlineStyle: 'wavyDouble',
  });
  expect(editor.isActive('italic')).toBe(false);
  expect(editor.getHTML()).toContain(
    'data-office-underline-style="wavyDouble"',
  );
  expect(editor.getHTML()).toContain('data-office-underline-color="#4472c4"');

  expect(editor.commands.undo()).toBe(true);
  editor.commands.setTextSelection(textRange(editor, 'Target'));
  expect(editor.isActive('italic')).toBe(true);
  expect(editor.isActive('underline')).toBe(false);
});

test('copies native double strike as one formatting mark across body and page chrome', () => {
  editor = new Editor({
    extensions: createWorkDocumentExtensions(),
    content:
      '<p><s data-office-strike-style="double">Source</s> and <em>Target</em></p>',
  });
  editor.commands.setTextSelection(textRange(editor, 'Source'));
  expect(copyDocumentFormatting(editor)).toBe(true);

  editor.commands.setTextSelection(textRange(editor, 'Target'));
  expect(pasteDocumentFormatting(editor)).toBe(true);
  expect(editor.getAttributes('strike')).toMatchObject({
    strikeStyle: 'double',
  });
  expect(editor.isActive('italic')).toBe(false);

  const pageChromeEditor = new Editor({
    extensions: createDocumentPageChromeEditorExtensions(),
    content: '<p>Header target</p>',
  });
  pageChromeEditor.commands.setTextSelection(
    textRange(pageChromeEditor, 'Header target'),
  );
  expect(pasteDocumentFormatting(pageChromeEditor)).toBe(true);
  expect(pageChromeEditor.getAttributes('strike')).toMatchObject({
    strikeStyle: 'double',
  });
  expect(pageChromeEditor.getHTML()).toContain(
    'data-office-strike-style="double"',
  );
  pageChromeEditor.destroy();

  expect(editor.commands.undo()).toBe(true);
  editor.commands.setTextSelection(textRange(editor, 'Target'));
  expect(editor.isActive('italic')).toBe(true);
  expect(editor.isActive('strike')).toBe(false);
});

test('copies native script fonts and resegments the target in one undo step', () => {
  const scriptFonts = normalizeDocumentScriptFonts({
    ascii: { name: 'Project Latin', resolved: 'Project Latin' },
    highAnsi: { name: 'Project ANSI', resolved: 'Project ANSI' },
    eastAsia: { name: 'Project East', resolved: 'Project East' },
    complexScript: { name: 'Project Complex', resolved: 'Project Complex' },
  });
  if (!scriptFonts) throw new Error('Expected valid script fonts.');
  const source = document.createElement('span');
  for (const [name, value] of Object.entries(
    documentScriptFontsDomAttributes(scriptFonts, 'ascii'),
  )) {
    source.setAttribute(name, value);
  }
  source.textContent = 'Source';
  editor = new Editor({
    extensions: createWorkDocumentExtensions(),
    content: `<p>${source.outerHTML} and <em>Aé中文ع</em></p>`,
  });
  editor.commands.setTextSelection(textRange(editor, 'Source'));
  expect(copyDocumentFormatting(editor)).toBe(true);

  const targetRange = textRange(editor, 'Aé中文ع');
  editor.commands.setTextSelection(targetRange);
  expect(pasteDocumentFormatting(editor)).toBe(true);

  const html = new DOMParser().parseFromString(editor.getHTML(), 'text/html');
  const target = Array.from(
    html.querySelectorAll<HTMLElement>('[data-office-script-font-slot]'),
  ).filter((element) => element.textContent !== 'Source');
  expect(
    target.map((element) => ({
      family: element.style.fontFamily.split(',')[0]?.replaceAll('"', ''),
      slot: element.dataset.officeScriptFontSlot,
      text: element.textContent,
    })),
  ).toEqual([
    { family: 'Project Latin', slot: 'ascii', text: 'A' },
    { family: 'Project ANSI', slot: 'highAnsi', text: 'é' },
    { family: 'Project East', slot: 'eastAsia', text: '中文' },
    { family: 'Project Complex', slot: 'complexScript', text: 'ع' },
  ]);
  for (const element of target) {
    expect(parseDocumentScriptFonts(element.dataset.officeScriptFonts)).toEqual(
      scriptFonts,
    );
  }

  expect(editor.commands.undo()).toBe(true);
  editor.commands.setTextSelection(targetRange);
  expect(editor.isActive('italic')).toBe(true);
  expect(editor.getHTML().match(/data-office-script-fonts/g)).toHaveLength(1);
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
  expect(pageChromeEditor.getAttributes('paragraph').paragraphBorders).toEqual({
    bottom: {
      style: 'double',
      color: { value: '#4472c4' },
      size: 12,
    },
  });
  expect(pageChromeEditor.getAttributes('paragraph').paragraphShading).toEqual({
    pattern: 'pct20',
    color: { value: '#112233' },
    fill: { value: '#ddeeff' },
  });
  expect(pageChromeEditor.getAttributes('textStyle')).toMatchObject({
    characterScalePercent: 80,
    characterPositionHalfPoints: 4,
    characterSpacingTwips: 30,
    emphasisMark: 'underDot',
    kerningThresholdHalfPoints: 24,
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
      `<h2 data-office-paragraph-borders='{"bottom":{"style":"double","color":{"value":"#4472c4"},"size":12}}' data-office-paragraph-shading='{"pattern":"pct20","color":{"value":"#112233"},"fill":{"value":"#ddeeff"}}' style="text-align: center; line-height: 1.5; border-bottom: 2px double #4472c4; background-color: #ddeeff">`,
      '<strong><span data-office-character-scale-percent="80" data-office-character-position-half-points="4" data-office-character-spacing-twips="30" data-office-kerning-threshold-half-points="24" data-office-emphasis-mark="underDot" style="font-stretch: 80%; --work-document-character-position: 2pt; color: #0070c0; font-family: Arial, sans-serif; font-size: 14pt; letter-spacing: 1.5pt; font-kerning: normal; text-emphasis-style: filled dot; text-emphasis-position: under right">Source</span></strong>',
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
