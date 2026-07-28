import { Editor } from '@tiptap/core';
import { closeHistory } from '@tiptap/pm/history';
import { describe, expect, test } from '@rstest/core';
import { createWorkDocumentExtensions } from '../src/internal/features/work/work-document-extensions';
import { createWorkDocumentModelFromContent } from '../src/internal/features/work/work-document-model-codec';
import {
  DocumentPagination,
  measureDocumentLayoutBlocks,
} from '../src/internal/features/work/work-document-pagination';
import {
  applyImportedDocxListMarkers,
  markDocxLists,
} from '../src/internal/features/work/work-docx-list-import';
import { parseXml } from '../src/internal/features/work/work-ooxml-package';
import type { WorkDocumentNode } from '../src/internal/features/work/work-types';

describe('document lists', () => {
  test('applies a bullet style idempotently and keeps it in undo history once', () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: '<p>Alpha</p><p>Beta</p>',
    });
    editor.commands.setTextSelection({
      from: textPosition(editor, 'Alpha'),
      to: textPosition(editor, 'Beta') + 'Beta'.length,
    });

    expect(editor.commands.applyDocumentBulletList('square')).toBe(true);
    expect(editor.getJSON().content?.[0]).toMatchObject({
      type: 'bulletList',
      attrs: { bulletStyle: 'square' },
    });
    expect(editor.getHTML()).toContain('data-office-bullet-style="square"');
    const styled = editor.getHTML();

    expect(editor.commands.applyDocumentBulletList('square')).toBe(true);
    expect(editor.getHTML()).toBe(styled);
    expect(editor.commands.undo()).toBe(true);
    expect(editor.getHTML()).toBe('<p>Alpha</p><p>Beta</p>');

    editor.destroy();
  });

  test('preserves the start value while changing an ordered-list style', () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: '<p>Alpha</p><p>Beta</p>',
    });
    editor.commands.setTextSelection({
      from: textPosition(editor, 'Alpha'),
      to: textPosition(editor, 'Beta') + 'Beta'.length,
    });

    expect(editor.commands.applyDocumentOrderedList('upper-roman')).toBe(true);
    expect(editor.commands.setDocumentNumberingStart(4)).toBe(true);
    expect(editor.commands.applyDocumentOrderedList('lower-alpha')).toBe(true);
    expect(editor.getJSON().content?.[0]).toMatchObject({
      type: 'orderedList',
      attrs: { start: 4, type: 'a' },
    });
    expect(editor.getHTML()).toContain('<ol start="4" type="a">');

    editor.view.dispatch(closeHistory(editor.state.tr));
    expect(editor.commands.restartDocumentNumbering()).toBe(true);
    expect(editor.getJSON().content?.[0]).toMatchObject({
      type: 'orderedList',
      attrs: { start: 1, type: 'a' },
    });
    expect(editor.commands.undo()).toBe(true);
    expect(editor.getJSON().content?.[0]).toMatchObject({
      type: 'orderedList',
      attrs: { start: 4, type: 'a' },
    });

    editor.destroy();
  });

  test('continues the preceding list with its numbering style', () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: [
        '<ol start="3" type="A">',
        '<li><p>First</p></li><li><p>Second</p></li>',
        '</ol>',
        '<p>Interruption</p>',
        '<ol><li><p>Third</p></li></ol>',
      ].join(''),
    });
    editor.commands.setTextSelection(textPosition(editor, 'Third'));

    expect(editor.can().continueDocumentNumbering()).toBe(true);
    expect(editor.commands.continueDocumentNumbering()).toBe(true);
    expect(editor.getJSON().content?.[2]).toMatchObject({
      type: 'orderedList',
      attrs: { start: 5, type: 'A' },
    });
    expect(editor.getHTML()).toContain('<ol start="5" type="A">');

    expect(editor.commands.undo()).toBe(true);
    expect(editor.getJSON().content?.[2]).toMatchObject({
      type: 'orderedList',
      attrs: { start: 1, type: null },
    });

    editor.destroy();
  });

  test('rejects invalid numbering values and unavailable continuation', () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: '<ol><li><p>Only list</p></li></ol>',
    });
    editor.commands.setTextSelection(textPosition(editor, 'Only list'));
    const initialHtml = editor.getHTML();

    expect(editor.commands.setDocumentNumberingStart(0)).toBe(false);
    expect(editor.commands.setDocumentNumberingStart(1.5)).toBe(false);
    expect(editor.can().continueDocumentNumbering()).toBe(false);
    expect(editor.commands.continueDocumentNumbering()).toBe(false);
    expect(editor.getHTML()).toBe(initialHtml);

    editor.destroy();
  });

  test('measures list items as independent page-layout blocks', () => {
    const editor = createListEditor();
    mockListGeometry(editor);

    const snapshot = measureDocumentLayoutBlocks(editor);

    expect(snapshot.blocks).toHaveLength(3);
    expect(snapshot.blocks.map(({ block }) => block.height)).toEqual([
      20, 30, 40,
    ]);
    expect(snapshot.blocks.map(({ block }) => block.id)).toEqual([
      expect.stringContaining('-item-0'),
      expect.stringContaining('-item-1'),
      expect.stringContaining('-item-2'),
    ]);
    expect(snapshot.blocks.map(({ from }) => from)).toEqual(
      [...snapshot.blocks.map(({ from }) => from)].sort(
        (left, right) => left - right,
      ),
    );

    editor.destroy();
  });

  test('places an automatic page gap inside the target list item', () => {
    const editor = createListEditor();
    mockListGeometry(editor);
    const snapshot = measureDocumentLayoutBlocks(editor);
    const second = snapshot.blocks[1];
    if (!second) throw new Error('Expected the second list item block.');

    editor.commands.applyDocumentPagination(1, [
      {
        beforeBlockId: second.block.id,
        pageIndex: 1,
        spacerHeight: 120,
        remainingBodyHeight: 30,
        page: {
          width: 300,
          height: 200,
          marginTop: 20,
          marginRight: 20,
          marginBottom: 20,
          marginLeft: 20,
          headerHeight: 10,
          footerHeight: 10,
          pageGap: 30,
        },
        position: second.from,
        inlineOffsetLeft: second.inlineOffsetLeft,
        inlineOffsetRight: second.inlineOffsetRight,
        previousPageChrome: {
          variant: 'first',
          headerHtml: '<p>First page header</p>',
          footerHtml: '<p>First page footer</p>',
          showPageNumber: true,
          pageNumber: 8,
        },
        nextPageChrome: {
          variant: 'even',
          headerHtml: '<p>Second page header</p>',
          footerHtml: '<p>Second page footer</p>',
          showPageNumber: true,
          pageNumber: 9,
        },
      },
    ]);

    const pageBreak = editor.view.dom.querySelector<HTMLElement>(
      '.work-document-auto-page-break',
    );
    expect(pageBreak).not.toBeNull();
    expect(pageBreak?.closest('li')?.textContent).toContain('Second item');
    expect(
      pageBreak?.querySelector('.work-document-page-break-chrome.footer'),
    ).toHaveTextContent('First page footer8');
    expect(
      pageBreak?.querySelector('.work-document-page-break-chrome.header'),
    ).toHaveTextContent('Second page header');
    expect(
      pageBreak?.querySelector('.work-document-page-break-chrome.footer'),
    ).toHaveAttribute('data-page-chrome-variant', 'first');
    expect(
      pageBreak?.querySelector('.work-document-page-break-chrome.header'),
    ).toHaveAttribute('data-page-chrome-variant', 'even');
    expect(
      pageBreak?.style.getPropertyValue(
        '--work-document-previous-page-footer-top',
      ),
    ).toBe('30px');
    expect(
      pageBreak?.style.getPropertyValue('--work-document-next-page-header-top'),
    ).toBe('80px');

    editor.destroy();
  });

  test('restores common Word numbering formats in HTML and the structured model', () => {
    const source = parseXml(
      [
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
        '<w:body>',
        ...Array.from(
          { length: 5 },
          (_, level) =>
            `<w:p><w:pPr><w:numPr><w:ilvl w:val="${level}"/>` +
            '<w:numId w:val="42"/></w:numPr></w:pPr>' +
            `<w:r><w:t>Level ${level}</w:t></w:r></w:p>`,
        ),
        '</w:body></w:document>',
      ].join(''),
      'document.xml',
    );
    const numbering = parseXml(
      [
        '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
        '<w:abstractNum w:abstractNumId="7">',
        listLevelXml(0, 'decimal', 3),
        listLevelXml(1, 'lowerLetter'),
        listLevelXml(2, 'upperLetter'),
        listLevelXml(3, 'lowerRoman'),
        listLevelXml(4, 'upperRoman'),
        '</w:abstractNum>',
        '<w:num w:numId="42"><w:abstractNumId w:val="7"/></w:num>',
        '</w:numbering>',
      ].join(''),
      'numbering.xml',
    );

    const markers = markDocxLists(source, numbering);

    expect(
      markers.lists.map(({ start, type }) => ({ start, type: type ?? null })),
    ).toEqual([
      { start: 3, type: null },
      { start: 1, type: 'a' },
      { start: 1, type: 'A' },
      { start: 1, type: 'i' },
      { start: 1, type: 'I' },
    ]);

    const imported = new DOMParser().parseFromString(
      markers.lists
        .map(
          ({ marker }, index) =>
            `<ol><li><p>${marker}Level ${index}</p></li></ol>`,
        )
        .join(''),
      'text/html',
    );
    applyImportedDocxListMarkers(imported, markers);
    const lists = Array.from(imported.body.querySelectorAll('ol'));

    expect(
      lists.map((list) => ({
        start: list.getAttribute('start'),
        type: list.getAttribute('type'),
      })),
    ).toEqual([
      { start: '3', type: null },
      { start: null, type: 'a' },
      { start: null, type: 'A' },
      { start: null, type: 'i' },
      { start: null, type: 'I' },
    ]);

    const content = createWorkDocumentModelFromContent({
      type: 'document',
      pageSize: 'a4',
      html: imported.body.innerHTML,
    });
    const orderedLists = collectNodes(content.model?.root, 'orderedList');

    expect(
      orderedLists.map((list) => ({
        start: list.attrs?.start,
        type: list.attrs?.type,
      })),
    ).toEqual([
      { start: 3, type: null },
      { start: 1, type: 'a' },
      { start: 1, type: 'A' },
      { start: 1, type: 'i' },
      { start: 1, type: 'I' },
    ]);
  });

  test('restores common Word bullet shapes in HTML and the structured model', () => {
    const source = parseXml(
      [
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
        '<w:body>',
        ...Array.from(
          { length: 3 },
          (_, level) =>
            `<w:p><w:pPr><w:numPr><w:ilvl w:val="${level}"/>` +
            '<w:numId w:val="51"/></w:numPr></w:pPr>' +
            `<w:r><w:t>Bullet ${level}</w:t></w:r></w:p>`,
        ),
        '</w:body></w:document>',
      ].join(''),
      'document.xml',
    );
    const numbering = parseXml(
      [
        '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
        '<w:abstractNum w:abstractNumId="9">',
        bulletLevelXml(0, '•'),
        bulletLevelXml(1, '◦'),
        bulletLevelXml(2, '▪'),
        '</w:abstractNum>',
        '<w:num w:numId="51"><w:abstractNumId w:val="9"/></w:num>',
        '</w:numbering>',
      ].join(''),
      'numbering.xml',
    );

    const markers = markDocxLists(source, numbering);
    expect(markers.lists.map(({ bulletStyle }) => bulletStyle)).toEqual([
      'disc',
      'circle',
      'square',
    ]);

    const imported = new DOMParser().parseFromString(
      markers.lists
        .map(
          ({ marker }, index) =>
            `<ul><li><p>${marker}Bullet ${index}</p></li></ul>`,
        )
        .join(''),
      'text/html',
    );
    applyImportedDocxListMarkers(imported, markers);
    const lists = Array.from(imported.body.querySelectorAll('ul'));
    expect(lists.map((list) => list.dataset.officeBulletStyle)).toEqual([
      'disc',
      'circle',
      'square',
    ]);

    const content = createWorkDocumentModelFromContent({
      type: 'document',
      pageSize: 'a4',
      html: imported.body.innerHTML,
    });
    expect(
      collectNodes(content.model?.root, 'bulletList').map(
        (list) => list.attrs?.bulletStyle,
      ),
    ).toEqual(['disc', 'circle', 'square']);
  });
});

function createListEditor(): Editor {
  return new Editor({
    extensions: [...createWorkDocumentExtensions(), DocumentPagination],
    content: [
      '<section data-document-section="true">',
      '<ol>',
      '<li><p>First item</p></li>',
      '<li><p>Second item</p></li>',
      '<li><p>Third item</p></li>',
      '</ol>',
      '</section>',
    ].join(''),
  });
}

function mockListGeometry(editor: Editor): void {
  const list = editor.view.dom.querySelector<HTMLElement>('ol');
  const items = Array.from(
    list?.querySelectorAll<HTMLElement>(':scope > li') ?? [],
  );
  const paragraphs = items.map((item) =>
    item.querySelector<HTMLElement>(':scope > p'),
  );
  if (!list || items.length !== 3 || paragraphs.some((item) => !item)) {
    throw new Error('Expected the mounted TipTap list.');
  }
  setElementHeight(list, 90);
  [20, 30, 40].forEach((height, index) => {
    setElementHeight(items[index], height);
    setElementHeight(paragraphs[index], height);
  });
}

function setElementHeight(
  element: HTMLElement | null | undefined,
  height: number,
): void {
  if (!element) throw new Error('Expected a measured list element.');
  Object.defineProperty(element, 'offsetHeight', {
    configurable: true,
    value: height,
  });
}

function listLevelXml(level: number, format: string, start = 1): string {
  return (
    `<w:lvl w:ilvl="${level}"><w:start w:val="${start}"/>` +
    `<w:numFmt w:val="${format}"/><w:lvlText w:val="%${level + 1}."/></w:lvl>`
  );
}

function bulletLevelXml(level: number, marker: string): string {
  return (
    `<w:lvl w:ilvl="${level}"><w:start w:val="1"/>` +
    `<w:numFmt w:val="bullet"/><w:lvlText w:val="${marker}"/></w:lvl>`
  );
}

function collectNodes(
  root: WorkDocumentNode | undefined,
  type: string,
): WorkDocumentNode[] {
  if (!root) return [];
  const matches: WorkDocumentNode[] = [];
  const pending = [root];
  while (pending.length) {
    const node = pending.shift();
    if (!node) continue;
    if (node.type === type) matches.push(node);
    pending.unshift(...(node.content ?? []));
  }
  return matches;
}

function textPosition(editor: Editor, text: string): number {
  let position: number | null = null;
  editor.state.doc.descendants((node, offset) => {
    if (position !== null || !node.isText || !node.text) return;
    const index = node.text.indexOf(text);
    if (index >= 0) position = offset + index;
  });
  if (position === null) throw new Error(`Text "${text}" was not found.`);
  return position;
}
