import { Editor } from '@tiptap/core';
import { expect, test } from '@rstest/core';
import {
  fireEvent,
  render,
  screen,
  waitFor,
  within,
} from '@testing-library/react';
import {
  createDocumentPageChromeEditorExtensions,
  DocumentPageChromeRichTextEditor,
  documentPageChromeEditorState,
  normalizeDocumentPageChromeHref,
} from '../src/internal/features/work/editors/document-page-chrome-editor';
import { clearDocumentFormatClipboard } from '../src/internal/features/work/editors/document-format-clipboard';
import { sanitizeDocumentPageChromeHtml } from '../src/internal/features/work/work-document-page-chrome';

const pixelPng =
  'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwC' +
  'AAAAC0lEQVR42mP8/x8AAusB9Y9Z9WQAAAAASUVORK5CYII=';

test('retains only normalized native image identities in page chrome', () => {
  const valid = sanitizeDocumentPageChromeHtml(
    `<p data-office-paragraph-id="2a2b3c4d" data-office-paragraph-text-id="2a2b3c4e"><img src="${pixelPng}" alt="Header" data-office-image-object-id="1a2b3c4d" data-office-image-doc-properties-id="42" data-office-image-anchor-id="1a2b3c4d" data-office-image-edit-id="0a0b0c0d" data-untrusted="drop" onerror="drop()"></p>`,
  );
  expect(valid).toContain('data-office-paragraph-id="2A2B3C4D"');
  expect(valid).toContain('data-office-paragraph-text-id="2A2B3C4E"');
  expect(valid).toContain('data-office-image-object-id="1A2B3C4D"');
  expect(valid).toContain('data-office-image-doc-properties-id="42"');
  expect(valid).toContain('data-office-image-anchor-id="1A2B3C4D"');
  expect(valid).toContain('data-office-image-edit-id="0A0B0C0D"');
  expect(valid).not.toContain('data-untrusted');
  expect(valid).not.toContain('onerror');

  const invalid = sanitizeDocumentPageChromeHtml(
    `<p data-office-paragraph-id="invalid" data-office-paragraph-text-id="invalid"><img src="${pixelPng}" data-office-image-object-id="invalid" data-office-image-doc-properties-id="-1" data-office-image-anchor-id="invalid" data-office-image-edit-id="invalid"></p>`,
  );
  expect(invalid).not.toContain('data-office-paragraph-');
  expect(invalid).not.toContain('data-office-image-');
});

test('retains only normalized native strikethrough metadata in page chrome', () => {
  const normalized = sanitizeDocumentPageChromeHtml(
    '<p><s data-office-strike-style="DOUBLE" data-untrusted="drop">Header</s></p>',
  );
  expect(normalized).toContain('data-office-strike-style="double"');
  expect(normalized).toContain('text-decoration-style: double');
  expect(normalized).not.toContain('data-untrusted');

  const fallback = sanitizeDocumentPageChromeHtml(
    '<p><strike data-office-strike-style="invalid">Header</strike></p>',
  );
  expect(fallback).toContain('data-office-strike-style="single"');
});

test('retains only normalized character position in page chrome', () => {
  const normalized = sanitizeDocumentPageChromeHtml(
    '<p><span data-office-character-position-half-points="-3" style="--work-document-character-position: 100pt" data-untrusted="drop">Header</span></p>',
  );
  expect(normalized).toContain(
    'data-office-character-position-half-points="-3"',
  );
  expect(normalized).toContain('--work-document-character-position: -1.5pt');
  expect(normalized).not.toContain('100pt');
  expect(normalized).not.toContain('data-untrusted');

  const invalid = sanitizeDocumentPageChromeHtml(
    '<p><span data-office-character-position-half-points="3169" style="--work-document-character-position: 2pt">Header</span></p>',
  );
  expect(invalid).not.toContain('data-office-character-position');
  expect(invalid).not.toContain('--work-document-character-position');
});

test('retains only normalized character scale in page chrome', () => {
  const normalized = sanitizeDocumentPageChromeHtml(
    '<p><span data-office-character-scale-percent="80" style="font-stretch: 125%" data-untrusted="drop">Header</span></p>',
  );
  expect(normalized).toContain('data-office-character-scale-percent="80"');
  expect(normalized).toContain('font-stretch: 80%');
  expect(normalized).not.toContain('125%');
  expect(normalized).not.toContain('data-untrusted');

  const explicitDefault = sanitizeDocumentPageChromeHtml(
    '<p><span data-office-character-scale-percent="100">Header</span></p>',
  );
  expect(explicitDefault).toContain(
    'data-office-character-scale-percent="100"',
  );
  expect(explicitDefault).toContain('font-stretch: 100%');

  const invalid = sanitizeDocumentPageChromeHtml(
    '<p><span data-office-character-scale-percent="601" style="font-stretch: 80%">Header</span></p>',
  );
  expect(invalid).not.toContain('data-office-character-scale');
  expect(invalid).not.toContain('font-stretch');
});

test('retains canonical font-size-aware kerning thresholds in page chrome', () => {
  const effective = sanitizeDocumentPageChromeHtml(
    '<p><span data-office-kerning-threshold-half-points="24" style="font-size: 12pt; font-kerning: none" data-untrusted="drop">Header</span></p>',
  );
  expect(effective).toContain('data-office-kerning-threshold-half-points="24"');
  expect(effective).toContain('font-size: 12pt');
  expect(effective).toContain('font-kerning: normal');
  expect(effective).not.toContain('data-untrusted');

  const belowThreshold = sanitizeDocumentPageChromeHtml(
    '<p><span data-office-kerning-threshold-half-points="25" style="font-size: 12pt; font-kerning: normal">Header</span></p>',
  );
  expect(belowThreshold).toContain('font-kerning: none');

  const explicitZero = sanitizeDocumentPageChromeHtml(
    '<p><span data-office-kerning-threshold-half-points="0" style="font-size: 8pt; font-kerning: none">Header</span></p>',
  );
  expect(explicitZero).toContain('font-kerning: normal');

  const invalid = sanitizeDocumentPageChromeHtml(
    '<p><span data-office-kerning-threshold-half-points="3278" style="font-size: 12pt; font-kerning: normal">Header</span></p>',
  );
  expect(invalid).not.toContain('data-office-kerning-threshold');
  expect(invalid).not.toContain('font-kerning');
  expect(invalid).toContain('font-size: 12pt');
});

test('retains only canonical native emphasis marks in page chrome', () => {
  const normalized = sanitizeDocumentPageChromeHtml(
    '<p><span data-office-emphasis-mark="underDot" style="text-emphasis-style: open circle; text-emphasis-position: over right" data-untrusted="drop">Header</span></p>',
  );
  expect(normalized).toContain('data-office-emphasis-mark="underDot"');
  expect(normalized).toContain('text-emphasis-style:filled dot');
  expect(normalized).toContain('text-emphasis-position:under right');
  expect(normalized).not.toContain('open circle');
  expect(normalized).not.toContain('data-untrusted');

  const explicitNone = sanitizeDocumentPageChromeHtml(
    '<p><span data-office-emphasis-mark="none" style="text-emphasis-style: filled dot">Header</span></p>',
  );
  expect(explicitNone).toContain('data-office-emphasis-mark="none"');
  expect(explicitNone).toContain('text-emphasis-style:none');
  expect(explicitNone).not.toContain('filled dot');

  const invalid = sanitizeDocumentPageChromeHtml(
    '<p><span data-office-emphasis-mark="Dot" style="text-emphasis-style: filled dot">Header</span></p>',
  );
  expect(invalid).not.toContain('data-office-emphasis-mark');
  expect(invalid).not.toContain('text-emphasis');
});

test('keeps native page-chrome identities through edits', () => {
  const editor = new Editor({
    extensions: createDocumentPageChromeEditorExtensions(),
    content: `<p data-office-paragraph-id="2A2B3C4D" data-office-paragraph-text-id="2A2B3C4E"><img src="${pixelPng}" data-office-image-object-id="1A2B3C4D" data-office-image-doc-properties-id="42" data-office-image-anchor-id="1A2B3C4D" data-office-image-edit-id="0A0B0C0D">Header</p>`,
  });
  expect(editor.getHTML()).toContain('data-office-paragraph-id="2A2B3C4D"');
  expect(editor.getHTML()).toContain('data-office-image-object-id="1A2B3C4D"');
  expect(editor.commands.insertContentAt(3, '!')).toBe(true);
  expect(editor.getHTML()).toContain('data-office-paragraph-id="2A2B3C4D"');
  expect(editor.getHTML()).not.toContain(
    'data-office-paragraph-text-id="2A2B3C4E"',
  );
  expect(editor.getHTML()).toContain('data-office-image-object-id="1A2B3C4D"');
  editor.destroy();
});

test('executes WPS alignment and format-copy shortcuts in page chrome', () => {
  clearDocumentFormatClipboard();
  const editor = new Editor({
    extensions: createDocumentPageChromeEditorExtensions(),
    content: '<p><strong>Source</strong> and <em>Target</em></p>',
  });
  editor.commands.setTextSelection(textRange(editor, 'Source'));

  fireEvent.keyDown(editor.view.dom, { key: 'e', ctrlKey: true });
  expect(editor.getAttributes('paragraph').textAlign).toBe('center');
  fireEvent.keyDown(editor.view.dom, {
    key: 'a',
    ctrlKey: true,
    shiftKey: true,
  });
  expect(editor.getAttributes('textStyle').textCase).toBe('all-caps');
  fireEvent.keyDown(editor.view.dom, {
    key: 'd',
    ctrlKey: true,
    shiftKey: true,
  });
  expect(editor.getAttributes('underline').underlineStyle).toBe('double');
  expect(editor.commands.setDocumentStrike('double')).toBe(true);
  fireEvent.keyDown(editor.view.dom, {
    key: 's',
    ctrlKey: true,
    shiftKey: true,
  });
  expect(editor.getAttributes('strike').strikeStyle).toBe('double');
  fireEvent.keyDown(editor.view.dom, {
    key: 'c',
    ctrlKey: true,
    shiftKey: true,
  });

  editor.commands.setTextSelection(textRange(editor, 'Target'));
  fireEvent.keyDown(editor.view.dom, {
    key: 'v',
    ctrlKey: true,
    shiftKey: true,
  });
  expect(editor.isActive('bold')).toBe(true);
  expect(editor.isActive('italic')).toBe(false);
  expect(editor.getHTML()).toContain('data-office-text-case="all-caps"');
  expect(editor.getAttributes('underline')).toMatchObject({
    underlineStyle: 'double',
  });
  expect(sanitizeDocumentPageChromeHtml(editor.getHTML())).toContain(
    'data-office-text-case="all-caps"',
  );
  expect(sanitizeDocumentPageChromeHtml(editor.getHTML())).toContain(
    'data-office-underline-style="double"',
  );

  editor.destroy();
  clearDocumentFormatClipboard();
});

test('applies typed page-chrome commands to a TipTap document', () => {
  const editor = new Editor({
    extensions: createDocumentPageChromeEditorExtensions(),
    content: '<p>Quarterly header</p>',
  });
  editor.commands.setTextSelection({ from: 1, to: 10 });

  expect(editor.chain().focus().toggleBold().run()).toBe(true);
  expect(editor.chain().focus().toggleUnderline().run()).toBe(true);
  expect(editor.chain().focus().setDocumentStrike('double').run()).toBe(true);
  expect(editor.chain().focus().setColor('#175cd3').run()).toBe(true);
  expect(
    editor
      .chain()
      .focus()
      .setDocumentPageChromeLink('https://a3s.dev/office')
      .run(),
  ).toBe(true);
  expect(editor.chain().focus().setTextAlign('right').run()).toBe(true);

  const state = documentPageChromeEditorState(editor);
  const html = sanitizeDocumentPageChromeHtml(editor.getHTML());
  expect(state).toMatchObject({
    alignment: 'right',
    bold: true,
    color: '#175cd3',
    link: 'https://a3s.dev/office',
    strike: true,
    strikeStyle: 'double',
    underline: true,
  });
  expect(html).toContain('text-align: right');
  expect(html).toContain('color: #175cd3');
  expect(html).toContain('href="https://a3s.dev/office"');
  expect(html).toContain('<strong>');
  expect(html).toContain('data-office-underline-style="single"');
  expect(html).toContain('data-office-strike-style="double"');

  expect(editor.commands.setDocumentPageChromeLink('javascript:alert(1)')).toBe(
    false,
  );
  editor.commands.setTextSelection(editor.state.doc.content.size);
  expect(
    editor.commands.insertDocumentPageChromeImage({
      alt: 'A3S mark',
      source: pixelPng,
    }),
  ).toBe(true);
  expect(editor.getHTML()).toContain('alt="A3S mark"');
  expect(editor.getHTML()).toContain('data:image/png;base64');
  editor.destroy();
});

test('preserves and switches page-chrome vertical-position marks', () => {
  const editor = new Editor({
    extensions: createDocumentPageChromeEditorExtensions(),
    content: '<p>H<sub>water</sub>O and x<sup>power</sup></p>',
  });

  expect(editor.getHTML()).toContain('<sub>water</sub>');
  expect(editor.getHTML()).toContain('<sup>power</sup>');

  editor.commands.setTextSelection(textRange(editor, 'water'));
  expect(documentPageChromeEditorState(editor)).toMatchObject({
    subscript: true,
    superscript: false,
  });
  expect(editor.commands.toggleDocumentSuperscript()).toBe(true);
  expect(editor.getHTML()).toContain('<sup>water</sup>');
  expect(editor.getHTML()).not.toContain('<sub>water</sub>');

  editor.commands.setTextSelection(textRange(editor, 'power'));
  expect(editor.commands.toggleDocumentSubscript()).toBe(true);
  expect(editor.getHTML()).toContain('<sub>power</sub>');
  expect(editor.getHTML()).not.toContain('<sup>power</sup>');
  editor.destroy();
});

test('keeps the page-chrome surface controlled and exposes active formatting', async () => {
  const changes: string[] = [];
  let editor: Editor | null = null;
  const view = render(
    <DocumentPageChromeRichTextEditor
      label="默认页眉"
      value="<p>Page title</p>"
      onChange={(html) => changes.push(html)}
      onEditorChange={(current) => {
        editor = current;
      }}
    />,
  );

  const textbox = await screen.findByRole('textbox', { name: '默认页眉' });
  expect(textbox).toHaveAttribute('data-document-page-chrome-engine', 'tiptap');
  expect(textbox).toHaveClass('work-document-page-chrome-content');
  const toolbar = screen.getByRole('toolbar', { name: '默认页眉格式' });
  const textFormatting = within(toolbar).getByRole('group', {
    name: '默认页眉文字格式',
  });
  const alignmentAndInsert = within(toolbar).getByRole('group', {
    name: '默认页眉对齐与插入',
  });
  expect(within(textFormatting).getAllByRole('button')).toHaveLength(8);
  expect(within(alignmentAndInsert).getAllByRole('button')).toHaveLength(7);
  expect(editor).not.toBeNull();
  const current = editor as Editor;
  current.commands.setTextSelection({ from: 1, to: 5 });
  fireEvent.click(screen.getByRole('button', { name: '默认页眉加粗' }));
  fireEvent.click(screen.getByRole('button', { name: '默认页眉上标' }));
  fireEvent.click(screen.getByRole('button', { name: '默认页眉删除线' }));

  await waitFor(() => {
    expect(changes.at(-1)).toContain('<strong>');
    expect(current.getHTML()).toContain('<sup>');
    expect(current.getHTML()).toContain('data-office-strike-style="single"');
  });
  expect(screen.getByRole('button', { name: '默认页眉加粗' })).toHaveAttribute(
    'aria-pressed',
    'true',
  );
  expect(screen.getByRole('button', { name: '默认页眉上标' })).toHaveAttribute(
    'aria-pressed',
    'true',
  );

  fireEvent.click(screen.getByRole('button', { name: '默认页眉下标' }));
  await waitFor(() => {
    expect(current.getHTML()).toContain('<sub>');
    expect(current.getHTML()).not.toContain('<sup>');
  });
  expect(screen.getByRole('button', { name: '默认页眉下标' })).toHaveAttribute(
    'aria-pressed',
    'true',
  );
  expect(screen.getByRole('button', { name: '默认页眉上标' })).toHaveAttribute(
    'aria-pressed',
    'false',
  );

  view.rerender(
    <DocumentPageChromeRichTextEditor
      label="默认页眉"
      value="<p>Externally updated</p>"
      onChange={(html) => changes.push(html)}
      onEditorChange={(next) => {
        editor = next;
      }}
    />,
  );
  await waitFor(() => {
    expect(screen.getByRole('textbox', { name: '默认页眉' })).toHaveTextContent(
      'Externally updated',
    );
  });
  view.unmount();
});

test('supports a toolbar-free page surface and exits with Escape', async () => {
  let exits = 0;
  render(
    <DocumentPageChromeRichTextEditor
      autoFocus
      label="页内页眉"
      value="<p>Direct page editing</p>"
      showToolbar={false}
      onChange={() => undefined}
      onExit={() => {
        exits += 1;
      }}
    />,
  );

  const textbox = await screen.findByRole('textbox', { name: '页内页眉' });
  expect(
    screen.queryByRole('toolbar', { name: '页内页眉格式' }),
  ).not.toBeInTheDocument();
  await waitFor(() => expect(document.activeElement).toBe(textbox));

  fireEvent.keyDown(textbox, { key: 'Escape' });
  expect(exits).toBe(1);
});

test('rejects unsafe links and unsupported header images without native UI', async () => {
  expect(normalizeDocumentPageChromeHref(' javascript:alert(1) ')).toBeNull();
  expect(normalizeDocumentPageChromeHref(' #summary ')).toBe('#summary');
  expect(normalizeDocumentPageChromeHref(' mailto:office@a3s.dev ')).toBe(
    'mailto:office@a3s.dev',
  );

  render(
    <DocumentPageChromeRichTextEditor
      label="默认页脚"
      value=""
      onChange={() => undefined}
    />,
  );
  const input = await screen.findByLabelText('默认页脚图片文件');
  fireEvent.change(input, {
    target: {
      files: [new File(['not an image'], 'notes.txt', { type: 'text/plain' })],
    },
  });

  expect(
    await screen.findByRole('heading', { name: '无法插入图片' }),
  ).toBeInTheDocument();
  expect(screen.getByRole('button', { name: '知道了' })).toBeInTheDocument();
});

function textRange(editor: Editor, text: string): { from: number; to: number } {
  let range: { from: number; to: number } | null = null;
  editor.state.doc.descendants((node, position) => {
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
