import { describe, expect, test } from '@rstest/core';
import { Editor } from '@tiptap/core';
import { closeHistory } from '@tiptap/pm/history';
import JSZip from 'jszip';
import {
  createArtifact,
  createArtifactBlob,
  importOfficeFile,
} from '../src/core';
import { createWorkDocumentExtensions } from '../src/internal/features/work/work-document-extensions';
import type {
  WorkDocumentFieldContext,
  WorkDocumentFieldKind,
} from '../src/internal/features/work/work-document-fields';
import type { WorkDocumentContent } from '../src/internal/features/work/work-types';

describe('document fields', () => {
  test('refreshes every field from its live page context in one undo step', () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: [
        '<section data-document-section="true" data-section-id="section-1"><p>Alpha',
        field('page', 'page-1'),
        field('numPages', 'pages-1'),
        field('section', 'section-1'),
        field('sectionPages', 'section-pages-1'),
        field('date', 'date-1', 'DATE \\@ "yyyy-MM-dd"'),
        field('time', 'time-1', 'TIME \\@ "HH:mm:ss"'),
        '</p></section>',
        '<section data-document-section="true" data-section-id="section-2"><p>Beta',
        field('page', 'page-2'),
        field('section', 'section-2'),
        '</p></section>',
      ].join(''),
    });
    try {
      const now = new Date(2026, 7, 11, 16, 5, 9);
      const resolveContext = (position: number): WorkDocumentFieldContext => ({
        ...(sectionIdAt(editor, position) === 'section-2'
          ? {
              pageNumber: 11,
              totalPages: 12,
              sectionNumber: 2,
              sectionPages: 2,
            }
          : {
              pageNumber: 7,
              totalPages: 12,
              sectionNumber: 1,
              sectionPages: 4,
            }),
        now,
      });

      editor.view.dispatch(closeHistory(editor.state.tr));
      expect(
        editor.commands.refreshDocumentFields(documentContent(editor), {
          resolveContext,
          now,
        }),
      ).toBe(true);
      expect(fieldDisplays(editor)).toEqual({
        'page-1': '7',
        'pages-1': '12',
        'section-1': '1',
        'section-pages-1': '4',
        'date-1': '2026-08-11',
        'time-1': '16:05:09',
        'page-2': '11',
        'section-2': '2',
      });

      expect(editor.commands.undo()).toBe(true);
      expect(new Set(Object.values(fieldDisplays(editor)))).toEqual(
        new Set(['stale']),
      );
      expect(editor.commands.redo()).toBe(true);
      expect(fieldDisplays(editor)['page-2']).toBe('11');
    } finally {
      editor.destroy();
    }
  });

  test('gives a copied field a new identity while retaining its semantics', () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: [
        '<section data-document-section="true" data-section-id="section-1">',
        `<p>Alpha${field('page', 'field-original', 'PAGE', '3')}</p>`,
        '<p>Beta</p></section>',
      ].join(''),
    });
    try {
      const original = documentFields(editor)[0];
      if (!original) throw new Error('Expected an original field.');
      const copy = editor.state.doc.slice(
        original.position,
        original.position + original.nodeSize,
        false,
      );
      editor.view.dispatch(
        closeHistory(editor.state.tr).insert(
          textRange(editor, 'Alpha').from,
          copy.content,
        ),
      );

      const copied = documentFields(editor);
      expect(copied).toHaveLength(2);
      const retained = copied.find(({ id }) => id === 'field-original');
      const duplicate = copied.find(({ id }) => id !== 'field-original');
      expect(retained).toMatchObject({
        id: 'field-original',
        kind: 'page',
        instruction: 'PAGE',
        display: '3',
      });
      expect(duplicate).toMatchObject({
        kind: 'page',
        instruction: 'PAGE',
        display: '3',
      });
      expect(duplicate?.id).toBeTruthy();
      const duplicateId = duplicate?.id;

      expect(editor.commands.undo()).toBe(true);
      expect(documentFields(editor).map(({ id }) => id)).toEqual([
        'field-original',
      ]);
      expect(editor.commands.redo()).toBe(true);
      expect(documentFields(editor).map(({ id }) => id)).toEqual([
        duplicateId,
        'field-original',
      ]);
    } finally {
      editor.destroy();
    }
  });

  test('refreshes pagination without ticking clock fields or adding history', () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: [
        '<section data-document-section="true" data-section-id="section-1"><p>',
        field('page', 'page'),
        field('date', 'date', 'DATE \\@ "yyyy-MM-dd"'),
        field('time', 'time', 'TIME \\@ "HH:mm:ss"'),
        '</p></section>',
      ].join(''),
    });
    try {
      editor.view.dispatch(closeHistory(editor.state.tr));
      expect(
        editor.commands.refreshDocumentFields(documentContent(editor), {
          resolveContext: () => ({
            pageNumber: 8,
            totalPages: 9,
            sectionNumber: 1,
            sectionPages: 9,
          }),
          now: new Date(2026, 7, 11, 16, 5, 9),
          addToHistory: false,
          updateClock: false,
        }),
      ).toBe(true);
      expect(fieldDisplays(editor)).toEqual({
        page: '8',
        date: 'stale',
        time: 'stale',
      });
      expect(editor.commands.undo()).toBe(false);
    } finally {
      editor.destroy();
    }
  });

  test('round-trips all supported body fields through native DOCX twice', async () => {
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    artifact.content.html = [
      '<section data-document-section="true" data-section-id="section-1"><p>',
      field('page', 'page', 'PAGE', '7'),
      field('numPages', 'pages', 'NUMPAGES', '12'),
      field('section', 'section', 'SECTION', '2'),
      field('sectionPages', 'section-pages', 'SECTIONPAGES', '4'),
      field('date', 'date', 'DATE \\@ "yyyy-MM-dd"', '2026-08-11'),
      field('time', 'time', 'TIME \\@ "HH:mm:ss"', '16:05:09'),
      '</p></section>',
    ].join('');

    const first = await createArtifactBlob(artifact);
    await expectNativeFields(first);
    const imported = await importOfficeFile(
      new File([first], 'fields.docx', { type: first.type }),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const document = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    const importedFields = Array.from(
      document.body.querySelectorAll<HTMLElement>('[data-document-field]'),
    );
    expect(importedFields.map(({ dataset }) => dataset.fieldKind)).toEqual([
      'page',
      'numPages',
      'section',
      'sectionPages',
      'date',
      'time',
    ]);
    expect(
      new Set(importedFields.map(({ dataset }) => dataset.fieldId)).size,
    ).toBe(6);
    expect(imported.compatibility.issues).toContainEqual(
      expect.objectContaining({ code: 'docx.fields.body', severity: 'info' }),
    );

    await expectNativeFields(await createArtifactBlob(imported));
  });
});

async function expectNativeFields(blob: Blob): Promise<void> {
  const archive = await JSZip.loadAsync(await blob.arrayBuffer());
  const documentXml =
    (await archive.file('word/document.xml')?.async('string')) ?? '';
  const settingsXml =
    (await archive.file('word/settings.xml')?.async('string')) ?? '';
  const instructions = Array.from(
    documentXml.matchAll(/<w:fldSimple\b[^>]*\bw:instr="([^"]*)"/g),
    ([, instruction]) => decodeXmlAttribute(instruction ?? '').trim(),
  );
  expect(
    instructions.map((instruction) =>
      /^\s*([A-Z]+)/i.exec(instruction)?.[1]?.toUpperCase(),
    ),
  ).toEqual(['PAGE', 'NUMPAGES', 'SECTION', 'SECTIONPAGES', 'DATE', 'TIME']);
  expect(instructions).toContain('DATE \\@ "yyyy-MM-dd"');
  expect(instructions).toContain('TIME \\@ "HH:mm:ss"');
  expect(settingsXml).toMatch(/<w:updateFields(?:\s[^>]*)?\/>/);
}

function decodeXmlAttribute(value: string): string {
  const element = new DOMParser().parseFromString(
    `<span data-value="${value}"></span>`,
    'text/html',
  ).body.firstElementChild;
  return element?.getAttribute('data-value') ?? value;
}

function field(
  kind: WorkDocumentFieldKind,
  id: string,
  instruction = fieldInstruction(kind),
  display = 'stale',
): string {
  return `<span data-document-field="true" data-field-id="${id}" data-field-kind="${kind}" data-field-instruction='${instruction}' data-field-display="${display}">${display}</span>`;
}

function fieldInstruction(kind: WorkDocumentFieldKind): string {
  return {
    page: 'PAGE',
    numPages: 'NUMPAGES',
    section: 'SECTION',
    sectionPages: 'SECTIONPAGES',
    date: 'DATE',
    time: 'TIME',
  }[kind];
}

function documentContent(editor: Editor): WorkDocumentContent {
  return {
    type: 'document',
    pageSize: 'a4',
    html: editor.getHTML(),
  };
}

function fieldDisplays(editor: Editor): Record<string, string> {
  return Object.fromEntries(
    documentFields(editor).map(({ id, display }) => [id, display]),
  );
}

function documentFields(editor: Editor): Array<{
  id: string;
  kind: string;
  instruction: string;
  display: string;
  position: number;
  nodeSize: number;
}> {
  const fields: ReturnType<typeof documentFields> = [];
  editor.state.doc.descendants((node, position) => {
    if (node.type.name !== 'documentField') return;
    fields.push({
      id: String(node.attrs.id),
      kind: String(node.attrs.kind),
      instruction: String(node.attrs.instruction),
      display: String(node.attrs.display),
      position,
      nodeSize: node.nodeSize,
    });
  });
  return fields;
}

function sectionIdAt(editor: Editor, position: number): string {
  const resolved = editor.state.doc.resolve(position);
  for (let depth = resolved.depth; depth > 0; depth -= 1) {
    const node = resolved.node(depth);
    if (node.type.name === 'documentSection') return String(node.attrs.id);
  }
  return '';
}

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
  if (!range) throw new Error(`Unable to find "${text}".`);
  return range;
}
