import { afterEach, expect, test } from '@rstest/core';
import { Editor } from '@tiptap/core';
import { fireEvent, render, screen, waitFor } from '@testing-library/react';
import { useState } from 'react';
import { DocumentFontDialog } from '../src/internal/features/work/editors/document-font-dialog';
import {
  applyDocumentFontDialogPatch,
  documentFontDialogSource,
} from '../src/internal/features/work/editors/document-font-dialog-model';
import type { WorkDocumentLayoutFont } from '../src/internal/features/work/work-document-fonts';
import { parseDocumentOpenTypeFeatures } from '../src/internal/features/work/work-document-opentype';
import {
  documentScriptFontsDomAttributes,
  normalizeDocumentScriptFonts,
  parseDocumentScriptFonts,
  type WorkDocumentScriptFonts,
} from '../src/internal/features/work/work-document-script-fonts';
import { createWorkDocumentExtensions } from '../src/internal/features/work/work-document-extensions';

const layoutFonts: readonly WorkDocumentLayoutFont[] = [
  { id: 'project-latin', family: 'Project Latin', url: '/project-latin.woff2' },
  { id: 'project-east', family: 'Project East', url: '/project-east.woff2' },
  {
    id: 'project-complex',
    family: 'Project Complex',
    url: '/project-complex.woff2',
  },
];

let editor: Editor | null = null;

afterEach(() => {
  editor?.destroy();
  editor = null;
});

test('reports independently mixed Latin, East Asian, and complex-script fonts', () => {
  const first = scriptFonts('Latin One', 'East Shared', 'Complex Shared');
  const second = scriptFonts('Latin Two', 'East Shared', 'Complex Shared');
  editor = createEditor(
    `<p>${fontSpan(first, 'A中ع')}${fontSpan(second, 'B文ب')}</p>`,
  );
  editor.commands.selectAll();

  const source = documentFontDialogSource(editor);

  expect(source.latinFont).toEqual({ mixed: true, value: null });
  expect(source.eastAsiaFont).toEqual({
    mixed: false,
    value: 'East Shared',
  });
  expect(source.complexScriptFont).toEqual({
    mixed: false,
    value: 'Complex Shared',
  });
});

test('keeps untouched mixed font groups while changing only one script font', async () => {
  const first = scriptFonts('Latin One', 'East Shared', 'Complex Shared');
  const second = scriptFonts('Latin Two', 'East Shared', 'Complex Shared');
  editor = createEditor(
    `<p>${fontSpan(first, 'A中ع')}${fontSpan(second, 'B文ب')}</p>`,
  );
  editor.commands.selectAll();
  const source = documentFontDialogSource(editor);
  const patches: unknown[] = [];

  render(
    <DocumentFontDialog
      source={source}
      layoutFonts={layoutFonts}
      restoreFocusTarget={() => editor?.view.dom ?? null}
      onApply={(patch) => {
        patches.push(patch);
        return true;
      }}
      onClose={() => undefined}
    />,
  );

  expect(screen.getByRole('button', { name: '应用' })).toBeDisabled();
  expect(screen.getByText(/不同的拉丁文字字体/)).toBeVisible();
  expect(screen.queryByText(/不同的东亚文字字体/)).toBeNull();
  expect(screen.queryByText(/不同的复杂文字字体/)).toBeNull();

  await selectFont('东亚文字字体', 'Project East');
  fireEvent.click(screen.getByRole('button', { name: '应用' }));

  expect(patches).toEqual([{ eastAsiaFont: 'Project East' }]);
});

test('applies three script fonts and scalar formatting in one undo step', async () => {
  const original = scriptFonts('Latin Old', 'East Old', 'Complex Old');
  const targetText = 'Aé中文اختبار';
  editor = createEditor(`<p>${fontSpan(original, targetText)}</p>`);
  document.body.append(editor.view.dom);
  const selection = textRange(editor, targetText);
  editor.commands.setTextSelection(selection);
  const source = documentFontDialogSource(editor);

  render(
    <FontDialogHarness editor={editor} selection={selection} source={source} />,
  );

  await selectFont('拉丁文字字体', 'Project Latin');
  await selectFont('东亚文字字体', 'Project East');
  await selectFont('复杂文字字体', 'Project Complex');
  await selectFont('OpenType 连字', '全部');
  fireEvent.change(screen.getByRole('textbox', { name: '字符缩放比例（%）' }), {
    target: { value: '80' },
  });
  fireEvent.click(screen.getByRole('button', { name: '应用' }));

  await waitFor(() =>
    expect(screen.queryByRole('dialog', { name: '字体高级设置' })).toBeNull(),
  );

  const formatted = scriptFontSpans(editor);
  expect(formatted.map(({ slot }) => slot)).toEqual([
    'ascii',
    'highAnsi',
    'eastAsia',
    'complexScript',
  ]);
  expect(formatted.map(({ firstFamily }) => firstFamily)).toEqual([
    'Project Latin',
    'Project Latin',
    'Project East',
    'Project Complex',
  ]);
  for (const { element, fonts } of formatted) {
    expect(fonts).toEqual({
      ascii: { name: 'Project Latin', resolved: 'Project Latin' },
      highAnsi: { name: 'Project Latin', resolved: 'Project Latin' },
      eastAsia: { name: 'Project East', resolved: 'Project East' },
      complexScript: {
        name: 'Project Complex',
        resolved: 'Project Complex',
      },
    });
    expect(element.dataset.officeCharacterScalePercent).toBe('80');
    expect(
      parseDocumentOpenTypeFeatures(element.dataset.officeOpentypeFeatures),
    ).toEqual({ ligatures: 'all' });
  }

  expect(editor.commands.undo()).toBe(true);
  editor.commands.setTextSelection(selection);
  const restored = documentFontDialogSource(editor);
  expect(restored.latinFont).toEqual({ mixed: false, value: 'Latin Old' });
  expect(restored.eastAsiaFont).toEqual({ mixed: false, value: 'East Old' });
  expect(restored.complexScriptFont).toEqual({
    mixed: false,
    value: 'Complex Old',
  });
  expect(restored.characterScale).toEqual({ mixed: false, value: null });
  expect(restored.openTypeLigatures).toEqual({ mixed: false, value: null });
  expect(editor.commands.undo()).toBe(false);
  await waitFor(() => expect(editor?.view.dom).toHaveFocus());
});

test('clears one direct font group through Follow Style and restores it with undo', async () => {
  const original = scriptFonts('Latin Old', 'East Old', 'Complex Old');
  editor = createEditor(`<p>${fontSpan(original, '中文')}</p>`);
  document.body.append(editor.view.dom);
  const selection = textRange(editor, '中文');
  editor.commands.setTextSelection(selection);

  render(
    <FontDialogHarness
      editor={editor}
      selection={selection}
      source={documentFontDialogSource(editor)}
    />,
  );

  await selectFont('东亚文字字体', '跟随样式');
  fireEvent.click(screen.getByRole('button', { name: '应用' }));
  await waitFor(() =>
    expect(screen.queryByRole('dialog', { name: '字体高级设置' })).toBeNull(),
  );

  editor.commands.setTextSelection(selection);
  expect(documentFontDialogSource(editor).eastAsiaFont).toEqual({
    mixed: false,
    value: null,
  });
  for (const { fonts } of scriptFontSpans(editor)) {
    expect(fonts?.eastAsia).toBeUndefined();
    expect(fonts?.ascii?.name).toBe('Latin Old');
    expect(fonts?.highAnsi?.name).toBe('Latin Old');
    expect(fonts?.complexScript?.name).toBe('Complex Old');
  }

  expect(editor.commands.undo()).toBe(true);
  editor.commands.setTextSelection(selection);
  expect(documentFontDialogSource(editor).eastAsiaFont).toEqual({
    mixed: false,
    value: 'East Old',
  });
});

test('applies script fonts to stored marks for the next typed character', () => {
  const original = scriptFonts('Latin Old', 'East Old', 'Complex Old');
  editor = createEditor(`<p>${fontSpan(original, 'A')}</p>`);
  editor.commands.setTextSelection(2);

  expect(
    applyDocumentFontDialogPatch(
      editor,
      { from: 2, to: 2 },
      {
        eastAsiaFont: 'Project East',
      },
    ),
  ).toBe(true);
  expect(editor.commands.insertContent('中')).toBe(true);

  const inserted = scriptFontSpans(editor).find(({ text }) => text === '中');
  expect(inserted?.slot).toBe('eastAsia');
  expect(inserted?.fonts?.eastAsia).toEqual({
    name: 'Project East',
    resolved: 'Project East',
  });
  expect(inserted?.firstFamily).toBe('Project East');

  expect(editor.commands.undo()).toBe(true);
  expect(editor.getText()).toBe('A');
});

function FontDialogHarness({
  editor,
  selection,
  source,
}: {
  editor: Editor;
  selection: { from: number; to: number };
  source: ReturnType<typeof documentFontDialogSource>;
}) {
  const [open, setOpen] = useState(true);
  return open ? (
    <DocumentFontDialog
      source={source}
      layoutFonts={layoutFonts}
      restoreFocusTarget={() => editor.view.dom}
      onApply={(patch) =>
        applyDocumentFontDialogPatch(editor, selection, patch)
      }
      onClose={() => setOpen(false)}
    />
  ) : null;
}

async function selectFont(ariaLabel: string, option: string): Promise<void> {
  fireEvent.click(screen.getByRole('combobox', { name: ariaLabel }));
  fireEvent.click(await screen.findByRole('option', { name: option }));
}

function createEditor(content: string): Editor {
  return new Editor({
    extensions: createWorkDocumentExtensions(),
    content,
  });
}

function scriptFonts(
  latin: string,
  eastAsia: string,
  complexScript: string,
): WorkDocumentScriptFonts {
  const fonts = normalizeDocumentScriptFonts({
    ascii: { name: latin, resolved: latin },
    highAnsi: { name: latin, resolved: latin },
    eastAsia: { name: eastAsia, resolved: eastAsia },
    complexScript: { name: complexScript, resolved: complexScript },
  });
  if (!fonts) throw new Error('Expected valid script fonts.');
  return fonts;
}

function fontSpan(fonts: WorkDocumentScriptFonts, text: string): string {
  const span = document.createElement('span');
  for (const [name, value] of Object.entries(
    documentScriptFontsDomAttributes(fonts, 'ascii'),
  )) {
    span.setAttribute(name, value);
  }
  span.textContent = text;
  return span.outerHTML;
}

function scriptFontSpans(currentEditor: Editor): Array<{
  element: HTMLElement;
  firstFamily: string | null;
  fonts: WorkDocumentScriptFonts | null;
  slot: string | undefined;
  text: string;
}> {
  const document = new DOMParser().parseFromString(
    currentEditor.getHTML(),
    'text/html',
  );
  return Array.from(
    document.querySelectorAll<HTMLElement>('[data-office-script-font-slot]'),
  ).map((element) => ({
    element,
    firstFamily: firstCssFamily(element.style.fontFamily),
    fonts: parseDocumentScriptFonts(element.dataset.officeScriptFonts),
    slot: element.dataset.officeScriptFontSlot,
    text: element.textContent ?? '',
  }));
}

function firstCssFamily(value: string): string | null {
  const family = value.split(',')[0]?.trim();
  return family ? family.replace(/^['"]|['"]$/g, '') : null;
}

function textRange(
  currentEditor: Editor,
  text: string,
): { from: number; to: number } {
  const chunks: Array<{ position: number; text: string }> = [];
  currentEditor.state.doc.descendants((node, position) => {
    if (node.isText && node.text) chunks.push({ position, text: node.text });
  });
  const source = chunks.map((chunk) => chunk.text).join('');
  const start = source.indexOf(text);
  if (start < 0) throw new Error(`Text "${text}" was not found.`);
  const end = start + text.length;
  let cursor = 0;
  let from: number | null = null;
  let to: number | null = null;
  for (const chunk of chunks) {
    const next = cursor + chunk.text.length;
    if (from === null && start >= cursor && start < next) {
      from = chunk.position + start - cursor;
    }
    if (end > cursor && end <= next) {
      to = chunk.position + end - cursor;
      break;
    }
    cursor = next;
  }
  if (from === null || to === null) {
    throw new Error(`Text "${text}" could not be mapped to the document.`);
  }
  return { from, to };
}
