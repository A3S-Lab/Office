import { describe, expect, test } from '@rstest/core';
import { Editor } from '@tiptap/core';
import { closeHistory } from '@tiptap/pm/history';
import { NodeSelection } from '@tiptap/pm/state';
import JSZip from 'jszip';
import {
  createArtifact,
  createArtifactBlob,
  importOfficeFile,
} from '../src/core';
import { createWorkDocumentExtensions } from '../src/internal/features/work/work-document-extensions';
import {
  DEFAULT_DOCUMENT_INDEX_OPTIONS,
  documentIndexEntryHtml,
  documentIndexHtml,
  type WorkDocumentIndexEntry,
  type WorkDocumentIndexOptions,
} from '../src/internal/features/work/work-document-index';
import {
  parseDocumentIndexEntryInstruction,
  parseDocumentIndexInstruction,
} from '../src/internal/features/work/work-document-index-fields';

describe('document index', () => {
  test('marks selected text without replacing it and keeps the action undoable', () => {
    const editor = createEditor();
    try {
      const range = textRange(editor, 'Architecture');
      editor.commands.setTextSelection(range);

      expect(
        editor.commands.markDocumentIndexEntry({
          mainEntry: 'Architecture',
          subEntry: 'Runtime',
          crossReference: '',
          pageBold: true,
          pageItalic: false,
        }),
      ).toBe(true);

      expect(editor.getText()).toContain('Architecture overview');
      expect(indexEntryNodes(editor)).toEqual([
        expect.objectContaining({
          mainEntry: 'Architecture',
          subEntry: 'Runtime',
          pageBold: true,
          pageItalic: false,
        }),
      ]);
      expect(editor.getHTML()).toContain('data-document-index-entry="true"');

      expect(editor.commands.undo()).toBe(true);
      expect(indexEntryNodes(editor)).toHaveLength(0);
      expect(editor.getText()).toContain('Architecture overview');
    } finally {
      editor.destroy();
    }
  });

  test('builds a sorted index, merges duplicate pages, and updates in one undo record', () => {
    const editor = createEditor();
    try {
      markText(editor, 'Architecture', {
        mainEntry: 'Architecture',
        subEntry: 'Runtime',
        pageBold: true,
      });
      markText(editor, 'overview', {
        mainEntry: 'Architecture',
        subEntry: 'Runtime',
      });
      markText(editor, 'Workers', {
        mainEntry: 'Architecture',
        subEntry: 'Workers',
        pageItalic: true,
      });
      markText(editor, 'Storage', {
        mainEntry: 'Storage',
        crossReference: 'Architecture',
      });
      editor.commands.setTextSelection(textRange(editor, 'Index insertion').to);

      expect(
        editor.commands.insertDocumentIndex(DEFAULT_DOCUMENT_INDEX_OPTIONS, {
          resolveContext: (position) => ({
            pageNumber: position < textRange(editor, 'Workers').from ? 2 : 5,
            totalPages: 6,
            sectionNumber: 1,
            sectionPages: 6,
          }),
        }),
      ).toBe(true);

      const first = indexBlock(editor);
      expect(first.entries).toEqual([
        {
          mainEntry: 'Architecture',
          subEntry: 'Runtime',
          crossReference: '',
          pages: [expect.objectContaining({ pageNumber: 2, pageBold: true })],
        },
        {
          mainEntry: 'Architecture',
          subEntry: 'Workers',
          crossReference: '',
          pages: [expect.objectContaining({ pageNumber: 5, pageItalic: true })],
        },
        {
          mainEntry: 'Storage',
          subEntry: '',
          crossReference: 'Architecture',
          pages: [],
        },
      ]);
      expect(first.entries[0]?.pages[0]?.targetIds).toHaveLength(2);

      editor.view.dispatch(closeHistory(editor.state.tr));
      editor.view.dispatch(
        editor.state.tr.setSelection(
          NodeSelection.create(editor.state.doc, first.position),
        ),
      );
      const nextOptions: WorkDocumentIndexOptions = {
        ...DEFAULT_DOCUMENT_INDEX_OPTIONS,
        columns: 2,
        format: 'run-in',
        leader: 'dash',
      };
      expect(editor.commands.updateDocumentIndex(nextOptions)).toBe(true);
      expect(indexBlock(editor).options).toEqual(nextOptions);

      expect(editor.commands.undo()).toBe(true);
      expect(indexBlock(editor).options).toEqual(
        DEFAULT_DOCUMENT_INDEX_OPTIONS,
      );
      expect(editor.commands.undo()).toBe(true);
      expect(() => indexBlock(editor)).toThrow('Expected a document index.');
    } finally {
      editor.destroy();
    }
  });

  test('parses the supported native XE and INDEX instruction subsets fail closed', () => {
    expect(
      parseDocumentIndexEntryInstruction(' XE "Architecture:Runtime" \\b \\i '),
    ).toEqual({
      supported: true,
      value: {
        mainEntry: 'Architecture',
        subEntry: 'Runtime',
        crossReference: '',
        pageBold: true,
        pageItalic: true,
      },
    });
    expect(
      parseDocumentIndexEntryInstruction('XE "Storage" \\t "See Architecture"'),
    ).toEqual({
      supported: true,
      value: {
        mainEntry: 'Storage',
        subEntry: '',
        crossReference: 'Architecture',
        pageBold: false,
        pageItalic: false,
      },
    });
    expect(
      parseDocumentIndexEntryInstruction('XE "Unsafe" \\r bookmark'),
    ).toMatchObject({ supported: false, reason: 'unsupported-switch' });

    expect(parseDocumentIndexInstruction('INDEX \\c "2" \\e "——"')).toEqual({
      supported: true,
      options: {
        ...DEFAULT_DOCUMENT_INDEX_OPTIONS,
        columns: 2,
        rightAlignPageNumbers: false,
        leader: 'none',
      },
    });
    expect(parseDocumentIndexInstruction('INDEX \\f "A"')).toMatchObject({
      supported: false,
      reason: 'unsupported-switch',
    });
  });

  test('round-trips native XE and INDEX fields, cached rows, and common options twice', async () => {
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    const options: WorkDocumentIndexOptions = {
      ...DEFAULT_DOCUMENT_INDEX_OPTIONS,
      columns: 2,
      leader: 'dash',
    };
    artifact.content.html = [
      '<section data-document-section="true" data-section-id="section-1">',
      documentIndexHtml({
        id: 'index-round-trip',
        options,
        entries: [
          {
            mainEntry: 'Architecture',
            subEntry: 'Runtime',
            crossReference: '',
            pages: [
              {
                pageNumber: 3,
                pageBold: true,
                pageItalic: false,
                targetIds: ['index-entry-runtime'],
              },
            ],
          },
          {
            mainEntry: 'Storage',
            subEntry: '',
            crossReference: 'Architecture',
            pages: [],
          },
        ],
      }),
      `<p>Runtime${documentIndexEntryHtml({
        id: 'index-entry-runtime',
        mainEntry: 'Architecture',
        subEntry: 'Runtime',
        crossReference: '',
        pageBold: true,
        pageItalic: false,
      })}</p>`,
      `<p>Storage${documentIndexEntryHtml({
        id: 'index-entry-storage',
        mainEntry: 'Storage',
        subEntry: '',
        crossReference: 'Architecture',
        pageBold: false,
        pageItalic: false,
      })}</p>`,
      '</section>',
    ].join('');

    const first = await createArtifactBlob(artifact);
    await expectNativeDocumentIndex(first);
    const imported = await importOfficeFile(
      new File([first], 'index.docx', { type: first.type }),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const document = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    expect(
      document.body.querySelectorAll('[data-document-index-entry]'),
    ).toHaveLength(2);
    const index = document.body.querySelector<HTMLElement>(
      '[data-document-index]',
    );
    expect(index).not.toBeNull();
    expect(index?.dataset.indexColumns).toBe('2');
    expect(index?.dataset.indexLeader).toBe('dash');
    expect(JSON.parse(index?.dataset.indexEntries ?? '[]')).toEqual([
      expect.objectContaining({
        mainEntry: 'Architecture',
        subEntry: 'Runtime',
        pages: [expect.objectContaining({ pageNumber: 3 })],
      }),
      expect.objectContaining({
        mainEntry: 'Storage',
        crossReference: 'Architecture',
      }),
    ]);
    expect(imported.compatibility.issues).toContainEqual(
      expect.objectContaining({ code: 'docx.index', severity: 'info' }),
    );
    expect(imported.compatibility.issues).not.toContainEqual(
      expect.objectContaining({ code: 'docx.fields.structure' }),
    );

    await expectNativeDocumentIndex(await createArtifactBlob(imported));
  });
});

function createEditor(): Editor {
  return new Editor({
    extensions: createWorkDocumentExtensions(),
    content: [
      '<section data-document-section="true" data-section-id="section-1">',
      '<p>Index insertion</p>',
      '<p>Architecture overview</p>',
      '<p>Workers and scheduling</p>',
      '<hr class="work-page-break" data-page-break="true">',
      '<p>Storage details</p>',
      '</section>',
    ].join(''),
  });
}

function markText(
  editor: Editor,
  text: string,
  value: Partial<Parameters<Editor['commands']['markDocumentIndexEntry']>[0]>,
): void {
  editor.commands.setTextSelection(textRange(editor, text));
  expect(
    editor.commands.markDocumentIndexEntry({
      mainEntry: text,
      subEntry: '',
      crossReference: '',
      pageBold: false,
      pageItalic: false,
      ...value,
    }),
  ).toBe(true);
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

function indexEntryNodes(editor: Editor): WorkDocumentIndexEntry[] {
  const result: WorkDocumentIndexEntry[] = [];
  editor.state.doc.descendants((node) => {
    if (node.type.name !== 'documentIndexEntry') return;
    result.push(node.attrs as WorkDocumentIndexEntry);
  });
  return result;
}

function indexBlock(editor: Editor) {
  let result:
    | {
        position: number;
        options: WorkDocumentIndexOptions;
        entries: Array<{
          mainEntry: string;
          subEntry: string;
          crossReference: string;
          pages: Array<{
            pageNumber: number;
            pageBold: boolean;
            pageItalic: boolean;
            targetIds: string[];
          }>;
        }>;
      }
    | undefined;
  editor.state.doc.descendants((node, position) => {
    if (result || node.type.name !== 'documentIndex') return;
    result = {
      position,
      options: {
        columns: Number(node.attrs.columns),
        format: node.attrs.format,
        rightAlignPageNumbers: Boolean(node.attrs.rightAlignPageNumbers),
        leader: node.attrs.leader,
      },
      entries: node.attrs.entries,
    };
  });
  if (!result) throw new Error('Expected a document index.');
  return result;
}

async function expectNativeDocumentIndex(blob: Blob): Promise<void> {
  const archive = await JSZip.loadAsync(await blob.arrayBuffer());
  const documentXml =
    (await archive.file('word/document.xml')?.async('string')) ?? '';
  const text =
    new DOMParser().parseFromString(documentXml, 'application/xml')
      .documentElement.textContent ?? '';
  const xml = new DOMParser().parseFromString(documentXml, 'application/xml');
  const simpleInstructions = Array.from(xml.querySelectorAll('*'))
    .filter((element) => element.localName === 'fldSimple')
    .map(
      (element) =>
        Array.from(element.attributes).find(
          (candidate) =>
            candidate.localName === 'instr' ||
            candidate.name.endsWith(':instr'),
        )?.value ?? '',
    );
  expect(documentXml).toMatch(/<w:fldSimple\b[^>]*\bw:instr="[^"]*XE /);
  expect(simpleInstructions).toContain('XE "Architecture:Runtime" \\b');
  expect(simpleInstructions).toContain('XE "Storage" \\t "See Architecture"');
  expect(documentXml).toMatch(/<w:sdt(?:\s[^>]*)?>/);
  expect(text).toContain('INDEX \\c "2"');
  expect(documentXml).toMatch(
    /<w:tab\b[^>]*\bw:val="right"[^>]*\bw:leader="hyphen"/,
  );
  expect(text).toContain('Architecture');
  expect(text).toContain('Runtime');
  expect(text).toContain('Storage');
}
