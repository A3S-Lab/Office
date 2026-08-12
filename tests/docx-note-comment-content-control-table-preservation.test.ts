import { describe, expect, test } from '@rstest/core';
import JSZip from 'jszip';
import { createArtifactBlob, importOfficeFile } from '../src/core';
import { createDocxBlob } from '../src/internal/features/work/work-docx-export';
import {
  attribute,
  descendants,
  directChild,
  directChildren,
  parseXml,
} from '../src/internal/features/work/work-ooxml-package';
import {
  xmlAttributeLocalName,
  xmlAttributeNamespace,
} from '../src/internal/features/work/work-docx-settings-xml';
import type { WorkDocumentContent } from '../src/internal/features/work/work-types';

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const STRICT_WORD_NAMESPACE =
  'http://purl.oclc.org/ooxml/wordprocessingml/main';
const WORD_2010_NAMESPACE =
  'http://schemas.microsoft.com/office/word/2010/wordml';
const MARKUP_COMPATIBILITY_NAMESPACE =
  'http://schemas.openxmlformats.org/markup-compatibility/2006';
const RELATIONSHIP_NAMESPACE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const XMLNS_NAMESPACE = 'http://www.w3.org/2000/xmlns/';
const VENDOR_NAMESPACE = 'urn:a3s:test:note-content-control-table';

describe('DOCX note content-control tables', () => {
  test('exports editable footnote and endnote tables as native OOXML blocks', async () => {
    for (const kind of ['footnote', 'endnote'] as const) {
      const output = await archiveFromContent(
        noteContent(tableNoteHtml('Alpha', false), kind),
      );
      const path = `word/${kind}s.xml`;
      const note = elementById(
        await xmlEntry(output, path),
        kind,
        kind === 'footnote' ? '27' : '41',
      );
      const table = directChildren(note, 'tbl')[0];
      expect(table).toBeDefined();
      expect(directChildren(table, 'tr')).toHaveLength(1);
      expect(descendants(table, 'tc').map(textOf)).toEqual(['Alpha', 'Beta']);
      expect(textOf(note)).toContain('BeforeAlphaBetaAfter');
      expect(textOf(note)).not.toContain('__A3S_PARAGRAPH_IDENTITY_');
      expect(
        word2010Attribute(directChildren(table, 'tr')[0], 'paraId'),
      ).toMatch(/^[0-9A-F]{8}$/u);
    }
  });

  test('restores a static control around a stable paragraph and table', async () => {
    const content = noteContent(tableNoteHtml('Alpha', false));
    const source = await archiveFromContent(content);
    const footnotes = await xmlEntry(source, 'word/footnotes.xml');
    const note = elementById(footnotes, 'footnote', '27');
    const before = directChildren(note, 'p').find(
      (paragraph) => textOf(paragraph) === 'Before',
    );
    const table = directChildren(note, 'tbl')[0];
    if (!before || !table) throw new Error('Missing source note blocks.');
    wrapBlockControl(footnotes, note, [before, table], '8421');
    table.setAttributeNS(VENDOR_NAMESPACE, 'vendor:token', 'stable-table');
    const width = directChild(directChild(table, 'tblPr') ?? table, 'tblW');
    width?.setAttributeNS(WORD_NAMESPACE, 'w:w', '42');
    const alphaRun = descendants(table, 'r').find(
      (run) => textOf(run) === 'Alpha',
    );
    if (!alphaRun) throw new Error('Missing source table run.');
    const runProperties = wordElement(footnotes, 'rPr');
    runProperties.append(wordElement(footnotes, 'smallCaps'));
    alphaRun.insertBefore(runProperties, alphaRun.firstChild);
    source.file('word/footnotes.xml', serializeXml(footnotes));

    const output = await exportWithSource(content, source);
    const outputNote = elementById(
      await xmlEntry(output, 'word/footnotes.xml'),
      'footnote',
      '27',
    );
    const control = directChildren(outputNote, 'sdt')[0];
    const controlled = directChild(control, 'sdtContent');
    const outputTable = directChildren(controlled ?? control, 'tbl')[0];
    expect(control).toBeDefined();
    expect(wordValue(directChild(directChild(control, 'sdtPr'), 'id'))).toBe(
      '8421',
    );
    expect(
      directChildren(controlled ?? control).map((item) => item.localName),
    ).toEqual(['p', 'tbl']);
    expect(textOf(controlled ?? control)).toContain('BeforeAlphaBeta');
    expect(vendorAttribute(outputTable, 'token')).toBe('stable-table');
    expect(descendants(outputTable, 'smallCaps')).toHaveLength(1);
    expect(
      attribute(directChild(directChild(outputTable, 'tblPr'), 'tblW'), 'w'),
    ).not.toBe('42');
    expect(
      directChildren(outputNote, 'p').some(
        (paragraph) => textOf(paragraph) === 'After',
      ),
    ).toBe(true);
  });

  test('preserves a control around a structurally stable nested endnote table', async () => {
    const content = noteContent(tableNoteHtml('Outer', true), 'endnote');
    const source = await archiveFromContent(content);
    const endnotes = await xmlEntry(source, 'word/endnotes.xml');
    const note = elementById(endnotes, 'endnote', '41');
    const table = directChildren(note, 'tbl')[0];
    if (!table) throw new Error('Missing nested source table.');
    wrapBlockControl(endnotes, note, [table], '8422');
    source.file('word/endnotes.xml', strictUtf16(endnotes));

    const output = await exportWithSource(content, source);
    const outputNote = elementById(
      await xmlEntry(output, 'word/endnotes.xml'),
      'endnote',
      '41',
    );
    const control = directChildren(outputNote, 'sdt')[0];
    expect(control).toBeDefined();
    expect(
      directChildren(directChild(control, 'sdtContent') ?? control, 'tbl'),
    ).toHaveLength(1);
    expect(descendants(control, 'tbl')).toHaveLength(2);
    expect(textOf(control)).toContain('OuterNestedBeta');
  });

  test('keeps generated merged-cell geometry inside a restored control', async () => {
    const content = noteContent(
      [
        '<table><tbody>',
        '<tr><td colspan="2"><p>Merged</p></td></tr>',
        '<tr><td><p>Left</p></td><td><p>Right</p></td></tr>',
        '</tbody></table>',
      ].join(''),
    );
    const source = await archiveFromContent(content);
    const footnotes = await xmlEntry(source, 'word/footnotes.xml');
    const note = elementById(footnotes, 'footnote', '27');
    const table = directChildren(note, 'tbl')[0];
    if (!table) throw new Error('Missing merged source table.');
    wrapBlockControl(footnotes, note, [table], '8425');
    source.file('word/footnotes.xml', serializeXml(footnotes));

    const output = await exportWithSource(content, source);
    const outputNote = elementById(
      await xmlEntry(output, 'word/footnotes.xml'),
      'footnote',
      '27',
    );
    const control = directChildren(outputNote, 'sdt')[0];
    expect(control).toBeDefined();
    if (!control) throw new Error('Missing restored merged-cell control.');
    const outputTable = directChildren(
      directChild(control, 'sdtContent') ?? control,
      'tbl',
    )[0];
    expect(directChildren(outputTable, 'tr')).toHaveLength(2);
    expect(wordValue(descendants(outputTable, 'gridSpan')[0])).toBe('2');
    expect(textOf(outputTable)).toBe('MergedLeftRight');
  });

  test('round-trips an imported controlled table through the public artifact boundary', async () => {
    const content = noteContent(tableNoteHtml('Imported', true));
    const source = await archiveFromContent(content);
    const footnotes = await xmlEntry(source, 'word/footnotes.xml');
    const note = elementById(footnotes, 'footnote', '27');
    const table = directChildren(note, 'tbl')[0];
    if (!table) throw new Error('Missing imported source table.');
    wrapBlockControl(footnotes, note, [table], '8424');
    source.file('word/footnotes.xml', serializeXml(footnotes));
    const sourceBuffer = await source.generateAsync({ type: 'arraybuffer' });

    const imported = await importOfficeFile(
      new File([sourceBuffer], 'controlled-note-table.docx', {
        type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      }),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const html = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    const importedNote = html.body.querySelector<HTMLElement>(
      'aside[data-document-note]',
    );
    expect(importedNote?.querySelectorAll('table')).toHaveLength(2);
    expect(importedNote?.textContent).toContain('ImportedNestedBeta');

    const output = await JSZip.loadAsync(
      await (await createArtifactBlob(imported)).arrayBuffer(),
    );
    const outputNote = elementById(
      await xmlEntry(output, 'word/footnotes.xml'),
      'footnote',
      '27',
    );
    const control = directChildren(outputNote, 'sdt')[0];
    expect(wordValue(directChild(directChild(control, 'sdtPr'), 'id'))).toBe(
      '8424',
    );
    expect(descendants(control, 'tbl')).toHaveLength(2);
  });

  test('drops the whole wrapper when a controlled table cell is edited', async () => {
    const sourceContent = noteContent(tableNoteHtml('Original', false));
    const source = await archiveFromContent(sourceContent);
    const footnotes = await xmlEntry(source, 'word/footnotes.xml');
    const note = elementById(footnotes, 'footnote', '27');
    const before = directChildren(note, 'p').find(
      (paragraph) => textOf(paragraph) === 'Before',
    );
    const table = directChildren(note, 'tbl')[0];
    if (!before || !table) throw new Error('Missing source note blocks.');
    wrapBlockControl(footnotes, note, [before, table], '8423');
    source.file('word/footnotes.xml', serializeXml(footnotes));

    const output = await exportWithSource(
      noteContent(tableNoteHtml('Edited', false)),
      source,
    );
    const outputNote = elementById(
      await xmlEntry(output, 'word/footnotes.xml'),
      'footnote',
      '27',
    );
    expect(descendants(outputNote, 'sdt')).toHaveLength(0);
    expect(directChildren(outputNote, 'tbl')).toHaveLength(1);
    expect(textOf(outputNote)).toContain('BeforeEditedBetaAfter');
  });

  test('fails closed for ambiguous, malformed, plain-text, changed, or relationship-bound tables', async () => {
    for (const [boundary, id] of [
      ['ambiguous', '8501'],
      ['malformed', '8502'],
      ['plain-text', '8503'],
      ['relationship', '8504'],
      ['structure', '8505'],
    ] as const) {
      const html =
        boundary === 'ambiguous'
          ? '<table><tbody><tr><td><p>Repeat</p></td></tr></tbody></table><table><tbody><tr><td><p>Repeat</p></td></tr></tbody></table>'
          : tableNoteHtml('Alpha', false);
      const content = noteContent(html);
      const source = await archiveFromContent(content);
      const footnotes = await xmlEntry(source, 'word/footnotes.xml');
      const note = elementById(footnotes, 'footnote', '27');
      const table = directChildren(note, 'tbl')[0];
      if (!table) throw new Error(`Missing ${boundary} source table.`);
      wrapBlockControl(footnotes, note, [table], id);
      if (boundary === 'plain-text') {
        const properties = directChild(directChildren(note, 'sdt')[0], 'sdtPr');
        directChild(properties ?? note, 'richText')?.remove();
        properties?.append(wordElement(footnotes, 'text'));
      } else if (boundary === 'malformed' || boundary === 'structure') {
        const span = wordElement(footnotes, 'gridSpan');
        span.setAttributeNS(
          WORD_NAMESPACE,
          'w:val',
          boundary === 'malformed' ? '0' : '2',
        );
        const cell = descendants(table, 'tc')[0];
        if (!cell) throw new Error(`Missing ${boundary} source cell.`);
        let cellProperties = directChild(cell, 'tcPr');
        if (!cellProperties) {
          cellProperties = wordElement(footnotes, 'tcPr');
          cell.insertBefore(cellProperties, cell.firstChild);
        }
        cellProperties.append(span);
      } else if (boundary === 'relationship') {
        const unsafe = footnotes.createElementNS(
          VENDOR_NAMESPACE,
          'vendor:unsafe',
        );
        unsafe.setAttributeNS(RELATIONSHIP_NAMESPACE, 'r:id', 'rIdUnsafe');
        table.append(unsafe);
      }
      source.file('word/footnotes.xml', serializeXml(footnotes));

      const output = await exportWithSource(content, source);
      const outputNote = elementById(
        await xmlEntry(output, 'word/footnotes.xml'),
        'footnote',
        '27',
      );
      expect(descendants(outputNote, 'sdt'), boundary).toHaveLength(0);
      expect(directChildren(outputNote, 'tbl').length, boundary).toBe(
        boundary === 'ambiguous' ? 2 : 1,
      );
    }
  });
});

function tableNoteHtml(firstCell: string, nested: boolean): string {
  const nestedTable = nested
    ? '<table><tbody><tr><td><p>Nested</p></td></tr></tbody></table>'
    : '';
  return [
    '<p>Before</p>',
    '<table data-office-table-width-type="pixels" data-office-table-width="300"><tbody>',
    `<tr><td><p>${firstCell}</p>${nestedTable}</td><td><p>Beta</p></td></tr>`,
    '</tbody></table>',
    '<p>After</p>',
  ].join('');
}

function noteContent(
  noteHtml: string,
  kind: 'endnote' | 'footnote' = 'footnote',
): WorkDocumentContent {
  const nativeId = kind === 'footnote' ? '27' : '41';
  return {
    type: 'document',
    pageSize: 'a4',
    html: [
      '<section data-document-section="true">',
      `<p>Body<sup data-document-note-reference="true" data-note-kind="${kind}" data-note-id="docx-${kind}-${nativeId}" data-note-number="1">1</sup></p>`,
      `<aside data-document-note="true" data-note-kind="${kind}" data-note-id="docx-${kind}-${nativeId}" data-note-number="1">${noteHtml}</aside>`,
      '</section>',
    ].join(''),
  };
}

async function archiveFromContent(
  content: WorkDocumentContent,
): Promise<JSZip> {
  const blob = await createDocxBlob(content);
  return JSZip.loadAsync(await blob.arrayBuffer());
}

async function exportWithSource(
  content: WorkDocumentContent,
  source: JSZip,
): Promise<JSZip> {
  const blob = await createDocxBlob(
    content,
    await source.generateAsync({ type: 'arraybuffer' }),
  );
  return JSZip.loadAsync(await blob.arrayBuffer());
}

function wrapBlockControl(
  document: Document,
  owner: Element,
  blocks: readonly Element[],
  id: string,
): void {
  const first = blocks[0];
  if (!first || blocks.some((block) => block.parentElement !== owner)) {
    throw new Error('Missing direct block-control content.');
  }
  declareVendor(document.documentElement);
  const control = wordElement(document, 'sdt');
  const properties = wordElement(document, 'sdtPr');
  properties.append(
    wordValueElement(document, 'alias', 'Table field'),
    wordValueElement(document, 'tag', 'table-field'),
    wordValueElement(document, 'id', id),
    wordElement(document, 'richText'),
  );
  const content = wordElement(document, 'sdtContent');
  control.setAttributeNS(VENDOR_NAMESPACE, 'vendor:token', 'wrapper');
  properties.setAttributeNS(VENDOR_NAMESPACE, 'vendor:token', 'properties');
  content.setAttributeNS(VENDOR_NAMESPACE, 'vendor:token', 'content');
  control.append(properties, content);
  owner.insertBefore(control, first);
  for (const block of blocks) content.append(block);
}

function declareVendor(root: Element): void {
  root.setAttributeNS(XMLNS_NAMESPACE, 'xmlns:vendor', VENDOR_NAMESPACE);
  if (
    !Array.from(root.attributes).some(
      (item) =>
        item.value === MARKUP_COMPATIBILITY_NAMESPACE &&
        item.name.startsWith('xmlns:'),
    )
  ) {
    root.setAttributeNS(
      XMLNS_NAMESPACE,
      'xmlns:mc',
      MARKUP_COMPATIBILITY_NAMESPACE,
    );
  }
  const current = (attribute(root, 'mc:Ignorable') ?? '').trim();
  const ignorable = Array.from(root.attributes).find(
    (item) =>
      xmlAttributeLocalName(item) === 'Ignorable' &&
      xmlAttributeNamespace(root, item) === MARKUP_COMPATIBILITY_NAMESPACE,
  );
  const value = [current, 'vendor'].filter(Boolean).join(' ');
  if (ignorable) ignorable.value = value;
  else {
    root.setAttributeNS(MARKUP_COMPATIBILITY_NAMESPACE, 'mc:Ignorable', value);
  }
}

function wordElement(document: Document, localName: string): Element {
  return document.createElementNS(WORD_NAMESPACE, `w:${localName}`);
}

function wordValueElement(
  document: Document,
  localName: string,
  value: string,
): Element {
  const element = wordElement(document, localName);
  element.setAttributeNS(WORD_NAMESPACE, 'w:val', value);
  return element;
}

function elementById(
  document: Document,
  localName: string,
  id: string,
): Element {
  const element = descendants(document, localName).find(
    (candidate) => wordValue(candidate, 'id') === id,
  );
  if (!element) throw new Error(`Missing ${localName} ${id}.`);
  return element;
}

function wordValue(
  element: Element | null | undefined,
  name = 'val',
): string | null {
  if (!element) return null;
  return (
    Array.from(element.attributes).find(
      (item) =>
        xmlAttributeLocalName(item) === name &&
        xmlAttributeNamespace(element, item) === WORD_NAMESPACE,
    )?.value ?? null
  );
}

function word2010Attribute(element: Element, name: string): string | null {
  return (
    Array.from(element.attributes).find(
      (item) =>
        xmlAttributeLocalName(item) === name &&
        xmlAttributeNamespace(element, item) === WORD_2010_NAMESPACE,
    )?.value ?? null
  );
}

function vendorAttribute(
  element: Element | null | undefined,
  name: string,
): string | null {
  if (!element) return null;
  return (
    Array.from(element.attributes).find(
      (item) =>
        xmlAttributeLocalName(item) === name &&
        xmlAttributeNamespace(element, item) === VENDOR_NAMESPACE,
    )?.value ?? null
  );
}

function textOf(element: Element): string {
  return descendants(element, 't')
    .map((item) => item.textContent ?? '')
    .join('');
}

async function xmlEntry(archive: JSZip, path: string): Promise<Document> {
  const source = await archive.file(path)?.async('text');
  if (!source) throw new Error(`Missing OOXML part: ${path}`);
  return parseXml(source, path);
}

function serializeXml(document: Document): string {
  return new XMLSerializer().serializeToString(document);
}

function strictUtf16(document: Document): Uint8Array {
  return utf16LittleEndian(
    serializeXml(document)
      .replaceAll(WORD_NAMESPACE, STRICT_WORD_NAMESPACE)
      .replace(
        /^\s*<\?xml[^?]*\?>/iu,
        '<?xml version="1.0" encoding="UTF-16" standalone="yes"?>',
      ),
  );
}

function utf16LittleEndian(source: string): Uint8Array {
  const bytes = new Uint8Array(2 + source.length * 2);
  bytes[0] = 0xff;
  bytes[1] = 0xfe;
  for (let index = 0; index < source.length; index += 1) {
    const codeUnit = source.charCodeAt(index);
    bytes[2 + index * 2] = codeUnit & 0xff;
    bytes[3 + index * 2] = codeUnit >>> 8;
  }
  return bytes;
}
