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
import { collectWorkDocumentOutline } from '../src/internal/features/work/work-document-outline';
import {
  documentTableOfContentsHtml,
  parseDocumentTableOfContentsInstruction,
  type WorkDocumentTableOfContentsEntry,
  type WorkDocumentTableOfContentsOptions,
} from '../src/internal/features/work/work-document-table-of-contents';
import { diagnoseDocxCaptions } from '../src/internal/features/work/work-docx-caption-diagnostics';
import {
  applyImportedDocxTableOfContentsMarkers,
  markDocxTablesOfContents,
  type ImportedDocxTableOfContentsMarkers,
} from '../src/internal/features/work/work-docx-table-of-contents-import';

const defaultOptions: WorkDocumentTableOfContentsOptions = {
  minLevel: 1,
  maxLevel: 3,
  hyperlinks: true,
  showPageNumbers: true,
  rightAlignPageNumbers: true,
  leader: 'dot',
};

describe('document table of contents', () => {
  test('inserts bounded heading entries from the shared outline and live pages', () => {
    const editor = createEditor();
    try {
      editor.commands.setTextSelection(textRange(editor, 'Introduction').from);
      expect(
        editor.commands.insertDocumentTableOfContents(defaultOptions, {
          resolveContext: (position) => ({
            pageNumber: headingAt(editor, position) === 'Execution' ? 7 : 2,
            totalPages: 8,
            sectionNumber: 1,
            sectionPages: 8,
          }),
        }),
      ).toBe(true);

      expect(tableOfContents(editor)).toMatchObject({
        options: defaultOptions,
        entries: [
          {
            targetId: 'heading-00000011',
            title: 'Introduction',
            level: 1,
            pageNumber: 2,
          },
          {
            targetId: 'heading-00000012',
            title: 'Execution',
            level: 2,
            pageNumber: 7,
          },
        ],
        truncated: false,
      });
      expect(editor.getHTML()).toContain(
        'data-document-table-of-contents="true"',
      );
      const html = new DOMParser().parseFromString(
        editor.getHTML(),
        'text/html',
      );
      expect(
        html.body.querySelector('[data-document-table-of-contents]')
          ?.textContent,
      ).not.toContain('Out of range');
    } finally {
      editor.destroy();
    }
  });

  test('updates the selected table of contents in one transaction and one undo record', () => {
    const editor = createEditor();
    try {
      editor.commands.setTextSelection(textRange(editor, 'Introduction').from);
      editor.commands.insertDocumentTableOfContents(defaultOptions);
      const original = tableOfContents(editor);
      editor.view.dispatch(closeHistory(editor.state.tr));
      editor.view.dispatch(
        editor.state.tr.setSelection(
          NodeSelection.create(editor.state.doc, original.position),
        ),
      );

      expect(
        editor.commands.updateDocumentTableOfContents({
          ...defaultOptions,
          maxLevel: 2,
          hyperlinks: false,
          rightAlignPageNumbers: false,
          leader: 'none',
        }),
      ).toBe(true);
      expect(tableOfContents(editor).options).toMatchObject({
        maxLevel: 2,
        hyperlinks: false,
        rightAlignPageNumbers: false,
        leader: 'none',
      });

      expect(editor.commands.undo()).toBe(true);
      expect(tableOfContents(editor).options).toEqual(defaultOptions);
      expect(editor.commands.undo()).toBe(true);
      expect(() => tableOfContents(editor)).toThrow(
        'Expected a table of contents.',
      );
    } finally {
      editor.destroy();
    }
  });

  test('assigns stable targets and includes native outline levels 7 through 9', () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: [
        '<section data-document-section="true" data-section-id="section-1">',
        '<p>Insertion point</p>',
        '<h1>Top heading</h1>',
        '<p data-office-outline-level="6">Level seven</p>',
        '<p data-office-outline-level="8">Level nine</p>',
        '</section>',
      ].join(''),
    });
    try {
      editor.commands.setTextSelection(textRange(editor, 'Insertion point').to);
      expect(
        editor.commands.insertDocumentTableOfContents({
          ...defaultOptions,
          maxLevel: 9,
        }),
      ).toBe(true);

      const table = tableOfContents(editor);
      expect(
        table.entries.map(({ title, level }) => ({ title, level })),
      ).toEqual([
        { title: 'Top heading', level: 1 },
        { title: 'Level seven', level: 7 },
        { title: 'Level nine', level: 9 },
      ]);
      expect(
        table.entries.every((entry) =>
          /^heading-[0-9a-f]{8}$/.test(entry.targetId),
        ),
      ).toBe(true);
      expect(
        collectWorkDocumentOutline(editor.state.doc).map((item) => item.id),
      ).toEqual(table.entries.map((entry) => entry.targetId));

      expect(editor.commands.undo()).toBe(true);
      expect(editor.getHTML()).not.toContain('data-document-table-of-contents');
      expect(editor.getHTML()).not.toContain('data-office-paragraph-id');
    } finally {
      editor.destroy();
    }
  });

  test('refreshes titles and page numbers without silently changing options', () => {
    const editor = createEditor();
    try {
      editor.commands.setTextSelection(textRange(editor, 'Introduction').from);
      editor.commands.insertDocumentTableOfContents(defaultOptions);
      const execution = textRange(editor, 'Execution');
      editor.view.dispatch(
        editor.state.tr.insertText(' plan', execution.to, execution.to),
      );
      editor.view.dispatch(closeHistory(editor.state.tr));

      expect(
        editor.commands.refreshDocumentTablesOfContents({
          resolveContext: (position) => ({
            pageNumber: headingAt(editor, position).startsWith('Execution')
              ? 8
              : 3,
            totalPages: 9,
            sectionNumber: 1,
            sectionPages: 9,
          }),
        }),
      ).toBe(true);
      expect(tableOfContents(editor)).toMatchObject({
        options: defaultOptions,
        entries: [
          expect.objectContaining({ title: 'Introduction', pageNumber: 3 }),
          expect.objectContaining({ title: 'Execution plan', pageNumber: 8 }),
        ],
      });

      expect(editor.commands.undo()).toBe(true);
      expect(tableOfContents(editor).entries[1]).toMatchObject({
        title: 'Execution',
        pageNumber: 1,
      });
      expect(editor.getText()).toContain('Execution plan');
    } finally {
      editor.destroy();
    }
  });

  test('parses the supported native instruction subset and rejects lossy switches', () => {
    expect(
      parseDocumentTableOfContentsInstruction(' TOC \\o "1-3" \\h \\z \\u '),
    ).toEqual({ supported: true, options: defaultOptions });
    expect(
      parseDocumentTableOfContentsInstruction(
        'TOC \\o "2-4" \\n "2-4" \\p " "',
      ),
    ).toEqual({
      supported: true,
      options: {
        minLevel: 2,
        maxLevel: 4,
        hyperlinks: false,
        showPageNumbers: false,
        rightAlignPageNumbers: false,
        leader: 'none',
      },
    });
    expect(
      parseDocumentTableOfContentsInstruction(
        'TOC \\o "1-3" \\t "Custom Heading,1"',
      ),
    ).toMatchObject({ supported: false, reason: 'unsupported-switch' });
    expect(
      parseDocumentTableOfContentsInstruction('TOC \\o "9-2"'),
    ).toMatchObject({ supported: false, reason: 'invalid-level-range' });
  });

  test('diagnoses a supported cross-paragraph native TOC without a field-structure warning', () => {
    const diagnostics = diagnoseDocxCaptions(
      wordDocument(nativeTableOfContentsXml('TOC \\o "1-3" \\h \\z \\u')),
    );

    expect(diagnostics.issues).toContainEqual(
      expect.objectContaining({
        code: 'docx.tableOfContents',
        severity: 'info',
      }),
    );
    expect(diagnostics.issues).not.toContainEqual(
      expect.objectContaining({ code: 'docx.fields.structure' }),
    );
    expect(diagnostics.hasUnsupportedFields).toBe(false);
  });

  test('keeps lossy and malformed TOC fields explicit in compatibility diagnostics', () => {
    const lossy = diagnoseDocxCaptions(
      wordDocument(
        nativeTableOfContentsXml('TOC \\o "1-3" \\t "Custom Heading,1"'),
      ),
    );
    expect(lossy.issues).not.toContainEqual(
      expect.objectContaining({ code: 'docx.tableOfContents' }),
    );
    expect(lossy.issues).toContainEqual(
      expect.objectContaining({ code: 'docx.fields.structure' }),
    );
    expect(lossy.hasUnsupportedFields).toBe(true);

    const malformed = diagnoseDocxCaptions(
      wordDocument(
        '<w:sdt><w:sdtContent><w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>TOC \\o &quot;1-3&quot;</w:instrText></w:r></w:p></w:sdtContent></w:sdt>',
      ),
    );
    expect(malformed.issues).toContainEqual(
      expect.objectContaining({ code: 'docx.fields.structure' }),
    );
    expect(malformed.hasUnsupportedFields).toBe(true);
    const malformedDocument = wordDocument(
      '<w:sdt><w:sdtContent><w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>TOC \\o &quot;1-3&quot;</w:instrText></w:r></w:p></w:sdtContent></w:sdt>',
    );
    expect(markDocxTablesOfContents(malformedDocument).tables).toHaveLength(0);
    expect(
      Array.from(malformedDocument.querySelectorAll('*')).filter(
        (element) => element.localName === 'sdt',
      ),
    ).toHaveLength(1);
  });

  test('restores the same deep heading target in every imported table of contents', () => {
    const document = new DOMParser().parseFromString(
      [
        '<p>__A3S_WORK_TABLE_OF_CONTENTS_1__</p>',
        '<p>__A3S_WORK_TABLE_OF_CONTENTS_2__</p>',
        '<p data-office-outline-level="6" data-office-paragraph-id="00000071">Deep heading</p>',
      ].join(''),
      'text/html',
    );
    const markers: ImportedDocxTableOfContentsMarkers = {
      tables: [1, 2].map((index) => ({
        marker: `__A3S_WORK_TABLE_OF_CONTENTS_${index}__`,
        id: `toc-${index}`,
        options: { ...defaultOptions, minLevel: 7, maxLevel: 7 },
        entries: [
          tocEntry(`heading-import-${index}-1`, 'Deep heading', 7, index),
        ],
      })),
    };

    applyImportedDocxTableOfContentsMarkers(document, markers);

    const tables = Array.from(
      document.body.querySelectorAll<HTMLElement>(
        '[data-document-table-of-contents]',
      ),
    );
    expect(tables).toHaveLength(2);
    for (const table of tables) {
      expect(JSON.parse(table.dataset.tocEntries ?? '[]')).toEqual([
        expect.objectContaining({ targetId: 'heading-00000071', level: 7 }),
      ]);
    }
  });

  test('round-trips a native TOC field, options, cached entries, and leader twice', async () => {
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    const options: WorkDocumentTableOfContentsOptions = {
      ...defaultOptions,
      leader: 'dash',
    };
    artifact.content.html = [
      '<section data-document-section="true" data-section-id="section-1">',
      documentTableOfContentsHtml({
        id: 'toc-round-trip',
        options,
        entries: [
          tocEntry('heading-00000021', 'Overview', 1, 1),
          tocEntry('heading-00000022', 'Details', 2, 4),
        ],
      }),
      '<h1 data-office-paragraph-id="00000021" data-office-paragraph-text-id="00000031">Overview</h1>',
      '<h2 data-office-paragraph-id="00000022" data-office-paragraph-text-id="00000032">Details</h2>',
      '</section>',
    ].join('');

    const first = await createArtifactBlob(artifact);
    await expectNativeTableOfContents(first, 'hyphen');
    const imported = await importOfficeFile(
      new File([first], 'toc.docx', { type: first.type }),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const document = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    const toc = document.body.querySelector<HTMLElement>(
      '[data-document-table-of-contents]',
    );
    expect(toc).not.toBeNull();
    expect(toc?.dataset.tocLeader).toBe('dash');
    expect(toc?.dataset.tocMinLevel).toBe('1');
    expect(toc?.dataset.tocMaxLevel).toBe('3');
    expect(toc?.dataset.tocHyperlinks).toBe('true');
    expect(toc?.dataset.tocShowPageNumbers).toBe('true');
    expect(toc?.dataset.tocRightAlignPageNumbers).toBe('true');
    expect(imported.compatibility.issues).toContainEqual(
      expect.objectContaining({
        code: 'docx.tableOfContents',
        severity: 'info',
      }),
    );
    expect(imported.compatibility.issues).not.toContainEqual(
      expect.objectContaining({ code: 'docx.hyperlinks.missing-target' }),
    );

    await expectNativeTableOfContents(
      await createArtifactBlob(imported),
      'hyphen',
    );
  });
});

function createEditor(): Editor {
  return new Editor({
    extensions: createWorkDocumentExtensions(),
    content: [
      '<section data-document-section="true" data-section-id="section-1">',
      '<h1 data-office-paragraph-id="00000011" data-office-paragraph-text-id="00000021">Introduction</h1>',
      '<p>Opening text</p>',
      '<h2 data-office-paragraph-id="00000012" data-office-paragraph-text-id="00000022">Execution</h2>',
      '<p>Delivery details</p>',
      '<h4 data-office-paragraph-id="00000013" data-office-paragraph-text-id="00000023">Out of range</h4>',
      '</section>',
    ].join(''),
  });
}

function tableOfContents(editor: Editor): {
  position: number;
  options: WorkDocumentTableOfContentsOptions;
  entries: WorkDocumentTableOfContentsEntry[];
  truncated: boolean;
} {
  let result: ReturnType<typeof tableOfContents> | null = null;
  editor.state.doc.descendants((node, position) => {
    if (result || node.type.name !== 'documentTableOfContents') return;
    result = {
      position,
      options: {
        minLevel: Number(node.attrs.minLevel),
        maxLevel: Number(node.attrs.maxLevel),
        hyperlinks: Boolean(node.attrs.hyperlinks),
        showPageNumbers: Boolean(node.attrs.showPageNumbers),
        rightAlignPageNumbers: Boolean(node.attrs.rightAlignPageNumbers),
        leader: node.attrs.leader,
      },
      entries: node.attrs.entries,
      truncated: Boolean(node.attrs.truncated),
    };
  });
  if (!result) throw new Error('Expected a table of contents.');
  return result;
}

function headingAt(editor: Editor, position: number): string {
  const node = editor.state.doc.nodeAt(Math.max(0, position - 1));
  return node?.type.name === 'heading' ? node.textContent : '';
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

function tocEntry(
  targetId: string,
  title: string,
  level: number,
  pageNumber: number,
): WorkDocumentTableOfContentsEntry {
  return { targetId, title, level, pageNumber };
}

async function expectNativeTableOfContents(
  blob: Blob,
  leader: string,
): Promise<void> {
  const archive = await JSZip.loadAsync(await blob.arrayBuffer());
  const documentXml =
    (await archive.file('word/document.xml')?.async('string')) ?? '';
  expect(documentXml).toMatch(/<w:sdt(?:\s[^>]*)?>/);
  expect(decodeXml(documentXml)).toMatch(
    /TOC \\h \\o "1-3" \\u \\z|TOC \\o "1-3" \\h \\z \\u/,
  );
  expect(documentXml).toMatch(
    new RegExp(`<w:tab[^>]*w:val="right"[^>]*w:leader="${leader}"`),
  );
  expect(decodeXml(documentXml)).toContain('Overview');
  expect(decodeXml(documentXml)).toContain('Details');
  expect(documentXml).toMatch(/\bparaId="00000021"/);
  expect(documentXml).toMatch(/\bparaId="00000022"/);
  expect(documentXml).toMatch(
    /<w:bookmarkStart\b[^>]*\bw:name="heading-00000021"/,
  );
  expect(documentXml).toMatch(
    /<w:bookmarkStart\b[^>]*\bw:name="heading-00000022"/,
  );
}

function decodeXml(value: string): string {
  return (
    new DOMParser().parseFromString(value, 'application/xml').documentElement
      .textContent ?? ''
  );
}

function wordDocument(body: string): Document {
  return new DOMParser().parseFromString(
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}</w:body></w:document>`,
    'application/xml',
  );
}

function nativeTableOfContentsXml(instruction: string): string {
  const escapedInstruction = instruction
    .replace(/&/g, '&amp;')
    .replace(/"/g, '&quot;');
  return [
    '<w:sdt><w:sdtContent>',
    '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>',
    `<w:r><w:instrText>${escapedInstruction}</w:instrText><w:fldChar w:fldCharType="separate"/></w:r>`,
    '<w:hyperlink w:anchor="heading-00000021"><w:r><w:t>Overview</w:t></w:r></w:hyperlink></w:p>',
    '<w:p><w:r><w:t>Details</w:t></w:r>',
    '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>',
    '</w:sdtContent></w:sdt>',
  ].join('');
}
