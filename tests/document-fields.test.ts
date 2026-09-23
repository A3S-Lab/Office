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
import {
  documentFieldCodeDisplay,
  documentFieldDisplay,
  documentFieldDraftFromAttributes,
  documentFieldInstruction,
  documentFieldOptionsFromDraft,
  documentFieldStatisticsFromHtml,
  docxDocumentFieldKind,
  numericFieldFormatSwitch,
  supportedDocxDocumentFieldInstruction,
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
      expect(editor.getHTML()).toContain('aria-label="当前页码"');

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

  test('resolves common statistics and bookmark page-reference fields', () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: [
        '<section data-document-section="true"><p>',
        '<span data-document-bookmark-boundary="true" data-bookmark-kind="start" data-bookmark-id="target" data-bookmark-name="Target"></span>',
        'Alpha beta 中文',
        '<span data-document-bookmark-boundary="true" data-bookmark-kind="end" data-bookmark-id="target" data-bookmark-name="Target"></span>',
        '</p><p>',
        field('wordCount', 'words', 'NUMWORDS', 'stale'),
        field('characterCount', 'chars', 'NUMCHARS', 'stale'),
        field('pageReference', 'target-page', 'PAGEREF Target \\h', 'stale'),
        '</p></section>',
      ].join(''),
    });
    try {
      expect(
        editor.commands.refreshDocumentFields(documentContent(editor), {
          resolveContext: (position) => ({
            pageNumber: position < 20 ? 3 : 9,
            totalPages: 9,
            sectionNumber: 1,
            sectionPages: 9,
          }),
        }),
      ).toBe(true);
      expect(fieldDisplays(editor)).toMatchObject({
        words: '4',
        chars: '13',
        'target-page': '3',
      });
    } finally {
      editor.destroy();
    }
  });

  test('keeps the bounded instruction grammar explicit', () => {
    expect(docxDocumentFieldKind('NUMWORDS \\* MERGEFORMAT')).toBe('wordCount');
    expect(docxDocumentFieldKind('FILENAME')).toBe('fileName');
    expect(docxDocumentFieldKind('FILENAME \\* MERGEFORMAT')).toBe('fileName');
    expect(docxDocumentFieldKind('AUTHOR')).toBe('author');
    expect(docxDocumentFieldKind('AUTHOR \\* MERGEFORMAT')).toBe('author');
    expect(docxDocumentFieldKind('TITLE')).toBe('title');
    expect(docxDocumentFieldKind('TITLE \\* MERGEFORMAT')).toBe('title');
    expect(docxDocumentFieldKind('SUBJECT')).toBe('subject');
    expect(docxDocumentFieldKind('SUBJECT \\* MERGEFORMAT')).toBe('subject');
    expect(docxDocumentFieldKind('KEYWORDS')).toBe('keywords');
    expect(docxDocumentFieldKind('KEYWORDS \\* MERGEFORMAT')).toBe('keywords');
    expect(docxDocumentFieldKind('LASTSAVEDBY')).toBe('lastSavedBy');
    expect(docxDocumentFieldKind('LASTSAVEDBY \\* MERGEFORMAT')).toBe(
      'lastSavedBy',
    );
    expect(docxDocumentFieldKind('COMMENTS')).toBe('comments');
    expect(docxDocumentFieldKind('COMMENTS \\* MERGEFORMAT')).toBe('comments');
    expect(docxDocumentFieldKind('CREATEDATE')).toBe('createDate');
    expect(docxDocumentFieldKind('CREATEDATE \\@ "yyyy-MM-dd"')).toBe(
      'createDate',
    );
    expect(docxDocumentFieldKind('SAVEDATE')).toBe('saveDate');
    expect(docxDocumentFieldKind('SAVEDATE \\@ "yyyy-MM-dd"')).toBe('saveDate');
    expect(docxDocumentFieldKind('PRINTDATE')).toBe('printDate');
    expect(docxDocumentFieldKind('PRINTDATE \\@ "yyyy-MM-dd"')).toBe(
      'printDate',
    );
    expect(docxDocumentFieldKind('MERGEFIELD CustomerName')).toBe('mergeField');
    expect(
      docxDocumentFieldKind('MERGEFIELD CustomerName \\* MERGEFORMAT'),
    ).toBe('mergeField');
    expect(supportedDocxDocumentFieldInstruction('MERGEFIELD CustomerName')).toBe(
      true,
    );
    expect(
      supportedDocxDocumentFieldInstruction(
        'MERGEFIELD CustomerName \\* MERGEFORMAT',
      ),
    ).toBe(true);
    expect(
      supportedDocxDocumentFieldInstruction('MERGEFIELD Customer Name'),
    ).toBe(false);
    expect(
      supportedDocxDocumentFieldInstruction('MERGEFIELD CustomerName \\b'),
    ).toBe(false);
    expect(documentFieldCodeDisplay('PAGE \\* ROMAN \\* MERGEFORMAT')).toBe(
      '{ PAGE \\* ROMAN \\* MERGEFORMAT }',
    );
    expect(documentFieldCodeDisplay('')).toBe('{ PAGE }');
    expect(docxDocumentFieldKind('NUMCHARS')).toBe('characterCount');
    expect(docxDocumentFieldKind('PAGEREF Target \\h')).toBe('pageReference');
    expect(supportedDocxDocumentFieldInstruction('PAGEREF Target \\h')).toBe(
      true,
    );
    expect(
      supportedDocxDocumentFieldInstruction(
        'PAGEREF Target \\h \\* MERGEFORMAT',
      ),
    ).toBe(true);
    expect(supportedDocxDocumentFieldInstruction('PAGEREF Target \\p')).toBe(
      false,
    );
    expect(supportedDocxDocumentFieldInstruction('NUMWORDS unexpected')).toBe(
      false,
    );
    expect(supportedDocxDocumentFieldInstruction('FILENAME')).toBe(true);
    expect(
      supportedDocxDocumentFieldInstruction('FILENAME \\* MERGEFORMAT'),
    ).toBe(true);
    expect(supportedDocxDocumentFieldInstruction('FILENAME \\p')).toBe(false);
    expect(supportedDocxDocumentFieldInstruction('AUTHOR')).toBe(true);
    expect(
      supportedDocxDocumentFieldInstruction('AUTHOR \\* MERGEFORMAT'),
    ).toBe(true);
    expect(supportedDocxDocumentFieldInstruction('AUTHOR \\* Caps')).toBe(
      false,
    );
    expect(supportedDocxDocumentFieldInstruction('TITLE')).toBe(true);
    expect(
      supportedDocxDocumentFieldInstruction('TITLE \\* MERGEFORMAT'),
    ).toBe(true);
    expect(supportedDocxDocumentFieldInstruction('TITLE \\* Caps')).toBe(
      false,
    );
    expect(supportedDocxDocumentFieldInstruction('SUBJECT')).toBe(true);
    expect(
      supportedDocxDocumentFieldInstruction('SUBJECT \\* MERGEFORMAT'),
    ).toBe(true);
    expect(supportedDocxDocumentFieldInstruction('SUBJECT \\* Caps')).toBe(
      false,
    );
    expect(supportedDocxDocumentFieldInstruction('KEYWORDS')).toBe(true);
    expect(
      supportedDocxDocumentFieldInstruction('KEYWORDS \\* MERGEFORMAT'),
    ).toBe(true);
    expect(supportedDocxDocumentFieldInstruction('KEYWORDS \\* Caps')).toBe(
      false,
    );
    expect(supportedDocxDocumentFieldInstruction('LASTSAVEDBY')).toBe(true);
    expect(
      supportedDocxDocumentFieldInstruction('LASTSAVEDBY \\* MERGEFORMAT'),
    ).toBe(true);
    expect(supportedDocxDocumentFieldInstruction('LASTSAVEDBY \\* Caps')).toBe(
      false,
    );
    expect(supportedDocxDocumentFieldInstruction('COMMENTS')).toBe(true);
    expect(
      supportedDocxDocumentFieldInstruction('COMMENTS \\* MERGEFORMAT'),
    ).toBe(true);
    expect(supportedDocxDocumentFieldInstruction('COMMENTS \\* Caps')).toBe(
      false,
    );
    expect(supportedDocxDocumentFieldInstruction('CREATEDATE')).toBe(true);
    expect(
      supportedDocxDocumentFieldInstruction(
        'CREATEDATE \\@ "yyyy-MM-dd" \\* MERGEFORMAT',
      ),
    ).toBe(true);
    expect(supportedDocxDocumentFieldInstruction('CREATEDATE \\@ "yyyy" extra')).toBe(
      false,
    );
    expect(supportedDocxDocumentFieldInstruction('SAVEDATE')).toBe(true);
    expect(
      supportedDocxDocumentFieldInstruction(
        'SAVEDATE \\@ "yyyy-MM-dd" \\* MERGEFORMAT',
      ),
    ).toBe(true);
    expect(supportedDocxDocumentFieldInstruction('SAVEDATE \\@ "yyyy" extra')).toBe(
      false,
    );
    expect(supportedDocxDocumentFieldInstruction('PRINTDATE')).toBe(true);
    expect(
      supportedDocxDocumentFieldInstruction(
        'PRINTDATE \\@ "yyyy-MM-dd" \\* MERGEFORMAT',
      ),
    ).toBe(true);
    expect(
      supportedDocxDocumentFieldInstruction('PRINTDATE \\@ "yyyy" extra'),
    ).toBe(false);
    expect(
      supportedDocxDocumentFieldInstruction(
        'DATE \\@ "yyyy-MM-dd" \\* MERGEFORMAT',
      ),
    ).toBe(true);
    expect(supportedDocxDocumentFieldInstruction('DATE \\@ "yyyy" extra')).toBe(
      false,
    );
    expect(
      documentFieldStatisticsFromHtml(
        '<p>Alpha <span data-document-field="true">99</span> 中文</p>',
      ),
    ).toEqual({
      wordCount: 3,
      characterCount: 9,
    });
    expect(
      documentFieldStatisticsFromHtml(
        '<p>Alpha<span data-document-field="true">99</span>beta</p>',
      ),
    ).toEqual({ wordCount: 2, characterCount: 9 });
    expect(
      documentFieldDisplay(
        'pageReference',
        {
          pageNumber: 1,
          totalPages: 1,
          sectionNumber: 1,
          sectionPages: 1,
          bookmarkPageNumbers: new Map([['name:target', 7]]),
        },
        'PAGEREF Target \\h',
        'stale',
      ),
    ).toBe('7');
    expect(
      documentFieldDisplay(
        'pageReference',
        {
          pageNumber: 1,
          totalPages: 1,
          sectionNumber: 1,
          sectionPages: 1,
          referencePageNumber: null,
        },
        'PAGEREF Target \\h',
        '9',
      ),
    ).toBe('引用缺失');
    expect(
      documentFieldDisplay(
        'fileName',
        {
          pageNumber: 1,
          totalPages: 1,
          sectionNumber: 1,
          sectionPages: 1,
          fileName: 'contract.docx',
        },
        'FILENAME',
        'stale.docx',
      ),
    ).toBe('contract.docx');
    expect(
      documentFieldDisplay(
        'fileName',
        {
          pageNumber: 1,
          totalPages: 1,
          sectionNumber: 1,
          sectionPages: 1,
        },
        'FILENAME',
        'cached.docx',
      ),
    ).toBe('cached.docx');
    expect(
      documentFieldDisplay(
        'author',
        {
          pageNumber: 1,
          totalPages: 1,
          sectionNumber: 1,
          sectionPages: 1,
          author: 'Morgan',
        },
        'AUTHOR',
        'stale',
      ),
    ).toBe('Morgan');
    expect(
      documentFieldDisplay(
        'title',
        {
          pageNumber: 1,
          totalPages: 1,
          sectionNumber: 1,
          sectionPages: 1,
          title: 'Quarterly Review',
        },
        'TITLE',
        'stale',
      ),
    ).toBe('Quarterly Review');
    expect(
      documentFieldDisplay(
        'title',
        {
          pageNumber: 1,
          totalPages: 1,
          sectionNumber: 1,
          sectionPages: 1,
        },
        'TITLE',
        'Cached Title',
      ),
    ).toBe('Cached Title');
    expect(
      documentFieldDisplay(
        'subject',
        {
          pageNumber: 1,
          totalPages: 1,
          sectionNumber: 1,
          sectionPages: 1,
          subject: 'Compliance',
        },
        'SUBJECT',
        'stale',
      ),
    ).toBe('Compliance');
    expect(
      documentFieldDisplay(
        'subject',
        {
          pageNumber: 1,
          totalPages: 1,
          sectionNumber: 1,
          sectionPages: 1,
        },
        'SUBJECT',
        'Cached Subject',
      ),
    ).toBe('Cached Subject');
    expect(
      documentFieldDisplay(
        'keywords',
        {
          pageNumber: 1,
          totalPages: 1,
          sectionNumber: 1,
          sectionPages: 1,
          keywords: 'contract; review',
        },
        'KEYWORDS',
        'stale',
      ),
    ).toBe('contract; review');
    expect(
      documentFieldDisplay(
        'keywords',
        {
          pageNumber: 1,
          totalPages: 1,
          sectionNumber: 1,
          sectionPages: 1,
        },
        'KEYWORDS',
        'cached; tags',
      ),
    ).toBe('cached; tags');
    expect(
      documentFieldDisplay(
        'lastSavedBy',
        {
          pageNumber: 1,
          totalPages: 1,
          sectionNumber: 1,
          sectionPages: 1,
          lastSavedBy: 'Riley',
        },
        'LASTSAVEDBY',
        'stale',
      ),
    ).toBe('Riley');
    expect(
      documentFieldDisplay(
        'lastSavedBy',
        {
          pageNumber: 1,
          totalPages: 1,
          sectionNumber: 1,
          sectionPages: 1,
        },
        'LASTSAVEDBY',
        'Cached Saver',
      ),
    ).toBe('Cached Saver');
    expect(
      documentFieldDisplay(
        'comments',
        {
          pageNumber: 1,
          totalPages: 1,
          sectionNumber: 1,
          sectionPages: 1,
          comments: 'Ready for legal review',
        },
        'COMMENTS',
        'stale',
      ),
    ).toBe('Ready for legal review');
    expect(
      documentFieldDisplay(
        'comments',
        {
          pageNumber: 1,
          totalPages: 1,
          sectionNumber: 1,
          sectionPages: 1,
        },
        'COMMENTS',
        'Cached remarks',
      ),
    ).toBe('Cached remarks');
    expect(
      documentFieldDisplay(
        'createDate',
        {
          pageNumber: 1,
          totalPages: 1,
          sectionNumber: 1,
          sectionPages: 1,
          createDate: new Date(2024, 2, 15),
        },
        'CREATEDATE \\@ "yyyy-MM-dd"',
        'stale',
      ),
    ).toBe('2024-03-15');
    expect(
      documentFieldDisplay(
        'createDate',
        {
          pageNumber: 1,
          totalPages: 1,
          sectionNumber: 1,
          sectionPages: 1,
        },
        'CREATEDATE \\@ "yyyy年M月d日"',
        '2020年1月1日',
      ),
    ).toBe('2020年1月1日');
    expect(
      documentFieldDisplay(
        'saveDate',
        {
          pageNumber: 1,
          totalPages: 1,
          sectionNumber: 1,
          sectionPages: 1,
          saveDate: new Date(2025, 5, 8),
        },
        'SAVEDATE \\@ "yyyy-MM-dd"',
        'stale',
      ),
    ).toBe('2025-06-08');
    expect(
      documentFieldDisplay(
        'saveDate',
        {
          pageNumber: 1,
          totalPages: 1,
          sectionNumber: 1,
          sectionPages: 1,
        },
        'SAVEDATE \\@ "yyyy年M月d日"',
        '2021年2月2日',
      ),
    ).toBe('2021年2月2日');
    expect(
      documentFieldDisplay(
        'printDate',
        {
          pageNumber: 1,
          totalPages: 1,
          sectionNumber: 1,
          sectionPages: 1,
          printDate: new Date(2023, 10, 20),
        },
        'PRINTDATE \\@ "yyyy-MM-dd"',
        'stale',
      ),
    ).toBe('2023-11-20');
    expect(
      documentFieldDisplay(
        'printDate',
        {
          pageNumber: 1,
          totalPages: 1,
          sectionNumber: 1,
          sectionPages: 1,
        },
        'PRINTDATE \\@ "yyyy年M月d日"',
        '2022年3月3日',
      ),
    ).toBe('2022年3月3日');
    expect(
      documentFieldDisplay(
        'mergeField',
        {
          pageNumber: 1,
          totalPages: 1,
          sectionNumber: 1,
          sectionPages: 1,
          mergeRecord: { CustomerName: 'Ada Lovelace' },
        },
        'MERGEFIELD CustomerName',
        '«CustomerName»',
      ),
    ).toBe('Ada Lovelace');
    expect(
      documentFieldDisplay(
        'mergeField',
        {
          pageNumber: 1,
          totalPages: 1,
          sectionNumber: 1,
          sectionPages: 1,
        },
        'MERGEFIELD CustomerName',
        '',
      ),
    ).toBe('«CustomerName»');
    expect(
      documentFieldDisplay(
        'mergeField',
        {
          pageNumber: 1,
          totalPages: 1,
          sectionNumber: 1,
          sectionPages: 1,
        },
        'MERGEFIELD CustomerName',
        'Cached merge',
      ),
    ).toBe('Cached merge');
  });

  test('resolves common numeric field switches without widening the grammar', () => {
    const context = {
      pageNumber: 42,
      totalPages: 42,
      sectionNumber: 4,
      sectionPages: 42,
    };
    expect(documentFieldDisplay('page', context, 'PAGE \\* ROMAN')).toBe(
      'XLII',
    );
    expect(documentFieldDisplay('page', context, 'PAGE \\* roman')).toBe(
      'xlii',
    );
    expect(
      documentFieldDisplay('numPages', context, 'NUMPAGES \\* ALPHABETIC'),
    ).toBe('AP');
    expect(
      documentFieldDisplay(
        'sectionPages',
        context,
        'SECTIONPAGES \\* alphabetic',
      ),
    ).toBe('ap');
    expect(
      documentFieldDisplay(
        'pageReference',
        { ...context, referencePageNumber: 7 },
        'PAGEREF Target \\* Ordinal',
        '7',
      ),
    ).toBe('7th');
    expect(numericFieldFormatSwitch('PAGE \\* ROMAN')).toBe('\\* ROMAN');
    expect(
      supportedDocxDocumentFieldInstruction('PAGE \\* ROMAN \\* MERGEFORMAT'),
    ).toBe(true);
    expect(
      supportedDocxDocumentFieldInstruction(
        'PAGEREF Target \\h \\* alphabetic',
      ),
    ).toBe(true);
    expect(
      supportedDocxDocumentFieldInstruction('PAGE \\* ROMAN \\* Arabic'),
    ).toBe(false);
    expect(supportedDocxDocumentFieldInstruction('PAGE \\# "000"')).toBe(false);
  });

  test('builds typed field settings without dropping native switches', () => {
    const draft = documentFieldDraftFromAttributes({
      kind: 'pageReference',
      instruction: 'PAGEREF WpsTarget \\h \\* ROMAN \\* MERGEFORMAT',
      targetId: 'bookmark-1',
      targetName: 'WpsTarget',
    });
    expect(draft).toEqual({
      kind: 'pageReference',
      format: { kind: 'numeric', value: 'roman' },
      targetId: 'bookmark-1',
      targetName: 'WpsTarget',
      hyperlink: true,
      mergeFormat: true,
    });
    if (!draft) throw new Error('Expected a page-reference draft.');
    expect(
      documentFieldInstruction(
        'pageReference',
        documentFieldOptionsFromDraft({
          ...draft,
          format: { kind: 'numeric', value: 'ordinal' },
        }),
      ),
    ).toBe('PAGEREF WpsTarget \\h \\* Ordinal \\* MERGEFORMAT');
    expect(
      documentFieldInstruction('date', {
        format: { kind: 'clock', value: 'yyyy-MM-dd' },
      }),
    ).toBe('DATE \\@ "yyyy-MM-dd"');
    expect(
      documentFieldDraftFromAttributes({
        instruction: 'DATE \\@ "HH:mm:ss"',
      }),
    ).toMatchObject({
      kind: 'date',
      format: { kind: 'clock', value: 'yyyy年M月d日', source: 'HH:mm:ss' },
    });
    expect(
      documentFieldDraftFromAttributes({
        instruction: 'TIME \\@ "yyyy-MM-dd"',
      }),
    ).toMatchObject({
      kind: 'time',
      format: { kind: 'clock', value: 'HH:mm', source: 'yyyy-MM-dd' },
    });
  });

  test('retargets a page reference when its bookmark identity is normalized', () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: [
        '<p>',
        '<span data-document-bookmark-boundary="true" data-bookmark-kind="start" data-bookmark-id="bookmark-target" data-bookmark-name="Target"></span>',
        'Target text',
        '<span data-document-bookmark-boundary="true" data-bookmark-kind="end" data-bookmark-id="bookmark-target" data-bookmark-name="Target"></span>',
        '</p><p>',
        field(
          'pageReference',
          'page-reference',
          'PAGEREF Target \\h \\* ROMAN',
          '1',
        ).replace(
          'data-field-instruction="PAGEREF Target \\h \\* ROMAN"',
          'data-field-instruction="PAGEREF Target \\h \\* ROMAN" data-field-target-id="bookmark-target" data-field-target-name="Target"',
        ),
        '</p>',
      ].join(''),
    });
    try {
      const start = documentBookmarkBoundary(editor, 'start');
      expect(start).toBeDefined();
      editor.view.dispatch(
        editor.state.tr.setNodeMarkup(start?.position ?? 0, undefined, {
          ...start?.node.attrs,
          name: 'Renamed',
        }),
      );
      const pageReference = documentFields(editor)[0];
      expect(pageReference).toMatchObject({
        instruction: 'PAGEREF Renamed \\h \\* ROMAN',
      });
      expect(editor.getHTML()).toContain('data-field-target-name="Renamed"');
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

  test('round-trips NUMWORDS, NUMCHARS, and bookmark-backed PAGEREF natively', async () => {
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    artifact.content.html = [
      '<section data-document-section="true"><p>',
      '<span data-document-bookmark-boundary="true" data-bookmark-kind="start" data-bookmark-id="bookmark-target" data-bookmark-name="Target"></span>',
      '目标内容',
      '<span data-document-bookmark-boundary="true" data-bookmark-kind="end" data-bookmark-id="bookmark-target" data-bookmark-name="Target"></span>',
      '</p><p>',
      field('wordCount', 'words', 'NUMWORDS', '2'),
      field('characterCount', 'chars', 'NUMCHARS', '4'),
      field('pageReference', 'page', 'PAGEREF Target \\h', '1'),
      '</p></section>',
    ].join('');

    const first = await createArtifactBlob(artifact);
    const archive = await JSZip.loadAsync(await first.arrayBuffer());
    const documentXml =
      (await archive.file('word/document.xml')?.async('string')) ?? '';
    expect(documentXml).toContain('NUMWORDS');
    expect(documentXml).toContain('NUMCHARS');
    expect(documentXml).toContain('PAGEREF Target \\h');

    const imported = await importOfficeFile(
      new File([first], 'common-fields.docx', { type: first.type }),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const importedDocument = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    const importedFields = Array.from(
      importedDocument.body.querySelectorAll<HTMLElement>(
        '[data-document-field]',
      ),
    );
    expect(importedFields.map(({ dataset }) => dataset.fieldKind)).toEqual([
      'wordCount',
      'characterCount',
      'pageReference',
    ]);
    expect(importedFields[2]?.dataset).toMatchObject({
      fieldTargetName: 'Target',
      fieldTargetId: expect.any(String),
    });
    expect(imported.compatibility.issues).toContainEqual(
      expect.objectContaining({ code: 'docx.fields.body', severity: 'info' }),
    );
    await expectNativeCommonFields(await createArtifactBlob(imported));
  });

  test('round-trips FILENAME as a live native field', async () => {
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    artifact.content.html = [
      '<section data-document-section="true"><p>',
      field('fileName', 'name', 'FILENAME', 'contract.docx'),
      '</p></section>',
    ].join('');

    const first = await createArtifactBlob(artifact);
    const archive = await JSZip.loadAsync(await first.arrayBuffer());
    const documentXml =
      (await archive.file('word/document.xml')?.async('string')) ?? '';
    expect(documentXml).toContain('FILENAME');
    expect(documentXml).toContain('contract.docx');

    const imported = await importOfficeFile(
      new File([first], 'filename-field.docx', { type: first.type }),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const document = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    const importedField = document.body.querySelector<HTMLElement>(
      '[data-document-field]',
    );
    expect(importedField?.dataset.fieldKind).toBe('fileName');
    expect(importedField?.dataset.fieldInstruction).toBe('FILENAME');
    expect(importedField?.textContent).toContain('contract.docx');
    expect(imported.compatibility.issues).toContainEqual(
      expect.objectContaining({ code: 'docx.fields.body', severity: 'info' }),
    );

    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: imported.content.html,
    });
    try {
      expect(
        editor.commands.refreshDocumentFields(documentContent(editor), {
          resolveContext: () => ({
            pageNumber: 1,
            totalPages: 1,
            sectionNumber: 1,
            sectionPages: 1,
            fileName: 'reviewed-contract.docx',
          }),
        }),
      ).toBe(true);
      expect(Object.values(fieldDisplays(editor))).toEqual([
        'reviewed-contract.docx',
      ]);
      expect(
        documentFields(editor).map((entry) => ({
          kind: entry.kind,
          display: entry.display,
        })),
      ).toEqual([{ kind: 'fileName', display: 'reviewed-contract.docx' }]);
    } finally {
      editor.destroy();
    }
  });

  test('round-trips AUTHOR as a live native field', async () => {
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    artifact.content.html = [
      '<section data-document-section="true"><p>',
      field('author', 'author', 'AUTHOR', 'Ada Reviewer'),
      '</p></section>',
    ].join('');

    const first = await createArtifactBlob(artifact);
    const archive = await JSZip.loadAsync(await first.arrayBuffer());
    const documentXml =
      (await archive.file('word/document.xml')?.async('string')) ?? '';
    expect(documentXml).toContain('AUTHOR');
    expect(documentXml).toContain('Ada Reviewer');

    const imported = await importOfficeFile(
      new File([first], 'author-field.docx', { type: first.type }),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const document = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    const importedField = document.body.querySelector<HTMLElement>(
      '[data-document-field]',
    );
    expect(importedField?.dataset.fieldKind).toBe('author');
    expect(importedField?.dataset.fieldInstruction).toBe('AUTHOR');
    expect(importedField?.textContent).toContain('Ada Reviewer');

    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: imported.content.html,
    });
    try {
      expect(
        editor.commands.refreshDocumentFields(documentContent(editor), {
          resolveContext: () => ({
            pageNumber: 1,
            totalPages: 1,
            sectionNumber: 1,
            sectionPages: 1,
            author: 'Morgan',
          }),
        }),
      ).toBe(true);
      expect(
        documentFields(editor).map((entry) => ({
          kind: entry.kind,
          display: entry.display,
        })),
      ).toEqual([{ kind: 'author', display: 'Morgan' }]);
    } finally {
      editor.destroy();
    }
  });

  test('round-trips TITLE as a live native field', async () => {
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    artifact.content.html = [
      '<section data-document-section="true"><p>',
      field('title', 'title', 'TITLE', 'Draft Contract'),
      '</p></section>',
    ].join('');

    const first = await createArtifactBlob(artifact);
    const archive = await JSZip.loadAsync(await first.arrayBuffer());
    const documentXml =
      (await archive.file('word/document.xml')?.async('string')) ?? '';
    expect(documentXml).toContain('TITLE');
    expect(documentXml).toContain('Draft Contract');

    const imported = await importOfficeFile(
      new File([first], 'title-field.docx', { type: first.type }),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const document = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    const importedField = document.body.querySelector<HTMLElement>(
      '[data-document-field]',
    );
    expect(importedField?.dataset.fieldKind).toBe('title');
    expect(importedField?.dataset.fieldInstruction).toBe('TITLE');
    expect(importedField?.textContent).toContain('Draft Contract');

    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: imported.content.html,
    });
    try {
      expect(
        editor.commands.refreshDocumentFields(documentContent(editor), {
          resolveContext: () => ({
            pageNumber: 1,
            totalPages: 1,
            sectionNumber: 1,
            sectionPages: 1,
            title: 'Quarterly Review',
          }),
        }),
      ).toBe(true);
      expect(
        documentFields(editor).map((entry) => ({
          kind: entry.kind,
          display: entry.display,
        })),
      ).toEqual([{ kind: 'title', display: 'Quarterly Review' }]);
    } finally {
      editor.destroy();
    }
  });

  test('round-trips SUBJECT as a live native field', async () => {
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    artifact.content.html = [
      '<section data-document-section="true"><p>',
      field('subject', 'subject', 'SUBJECT', 'Legal Review'),
      '</p></section>',
    ].join('');

    const first = await createArtifactBlob(artifact);
    const archive = await JSZip.loadAsync(await first.arrayBuffer());
    const documentXml =
      (await archive.file('word/document.xml')?.async('string')) ?? '';
    expect(documentXml).toContain('SUBJECT');
    expect(documentXml).toContain('Legal Review');

    const imported = await importOfficeFile(
      new File([first], 'subject-field.docx', { type: first.type }),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const document = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    const importedField = document.body.querySelector<HTMLElement>(
      '[data-document-field]',
    );
    expect(importedField?.dataset.fieldKind).toBe('subject');
    expect(importedField?.dataset.fieldInstruction).toBe('SUBJECT');
    expect(importedField?.textContent).toContain('Legal Review');

    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: imported.content.html,
    });
    try {
      expect(
        editor.commands.refreshDocumentFields(documentContent(editor), {
          resolveContext: () => ({
            pageNumber: 1,
            totalPages: 1,
            sectionNumber: 1,
            sectionPages: 1,
            subject: 'Compliance',
          }),
        }),
      ).toBe(true);
      expect(
        documentFields(editor).map((entry) => ({
          kind: entry.kind,
          display: entry.display,
        })),
      ).toEqual([{ kind: 'subject', display: 'Compliance' }]);
    } finally {
      editor.destroy();
    }
  });

  test('round-trips KEYWORDS as a live native field', async () => {
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    artifact.content.html = [
      '<section data-document-section="true"><p>',
      field('keywords', 'keywords', 'KEYWORDS', 'draft; legal'),
      '</p></section>',
    ].join('');

    const first = await createArtifactBlob(artifact);
    const archive = await JSZip.loadAsync(await first.arrayBuffer());
    const documentXml =
      (await archive.file('word/document.xml')?.async('string')) ?? '';
    expect(documentXml).toContain('KEYWORDS');
    expect(documentXml).toContain('draft; legal');

    const imported = await importOfficeFile(
      new File([first], 'keywords-field.docx', { type: first.type }),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const document = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    const importedField = document.body.querySelector<HTMLElement>(
      '[data-document-field]',
    );
    expect(importedField?.dataset.fieldKind).toBe('keywords');
    expect(importedField?.dataset.fieldInstruction).toBe('KEYWORDS');
    expect(importedField?.textContent).toContain('draft; legal');

    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: imported.content.html,
    });
    try {
      expect(
        editor.commands.refreshDocumentFields(documentContent(editor), {
          resolveContext: () => ({
            pageNumber: 1,
            totalPages: 1,
            sectionNumber: 1,
            sectionPages: 1,
            keywords: 'contract; review',
          }),
        }),
      ).toBe(true);
      expect(
        documentFields(editor).map((entry) => ({
          kind: entry.kind,
          display: entry.display,
        })),
      ).toEqual([{ kind: 'keywords', display: 'contract; review' }]);
    } finally {
      editor.destroy();
    }
  });

  test('round-trips LASTSAVEDBY as a live native field', async () => {
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    artifact.content.html = [
      '<section data-document-section="true"><p>',
      field('lastSavedBy', 'last-saved', 'LASTSAVEDBY', 'Ada Reviewer'),
      '</p></section>',
    ].join('');

    const first = await createArtifactBlob(artifact);
    const archive = await JSZip.loadAsync(await first.arrayBuffer());
    const documentXml =
      (await archive.file('word/document.xml')?.async('string')) ?? '';
    expect(documentXml).toContain('LASTSAVEDBY');
    expect(documentXml).toContain('Ada Reviewer');

    const imported = await importOfficeFile(
      new File([first], 'lastsavedby-field.docx', { type: first.type }),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const document = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    const importedField = document.body.querySelector<HTMLElement>(
      '[data-document-field]',
    );
    expect(importedField?.dataset.fieldKind).toBe('lastSavedBy');
    expect(importedField?.dataset.fieldInstruction).toBe('LASTSAVEDBY');
    expect(importedField?.textContent).toContain('Ada Reviewer');

    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: imported.content.html,
    });
    try {
      expect(
        editor.commands.refreshDocumentFields(documentContent(editor), {
          resolveContext: () => ({
            pageNumber: 1,
            totalPages: 1,
            sectionNumber: 1,
            sectionPages: 1,
            lastSavedBy: 'Riley',
          }),
        }),
      ).toBe(true);
      expect(
        documentFields(editor).map((entry) => ({
          kind: entry.kind,
          display: entry.display,
        })),
      ).toEqual([{ kind: 'lastSavedBy', display: 'Riley' }]);
    } finally {
      editor.destroy();
    }
  });

  test('round-trips COMMENTS as a live native field', async () => {
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    artifact.content.html = [
      '<section data-document-section="true"><p>',
      field('comments', 'comments', 'COMMENTS', 'Draft for counsel'),
      '</p></section>',
    ].join('');

    const first = await createArtifactBlob(artifact);
    const archive = await JSZip.loadAsync(await first.arrayBuffer());
    const documentXml =
      (await archive.file('word/document.xml')?.async('string')) ?? '';
    expect(documentXml).toContain('COMMENTS');
    expect(documentXml).toContain('Draft for counsel');

    const imported = await importOfficeFile(
      new File([first], 'comments-field.docx', { type: first.type }),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const document = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    const importedField = document.body.querySelector<HTMLElement>(
      '[data-document-field]',
    );
    expect(importedField?.dataset.fieldKind).toBe('comments');
    expect(importedField?.dataset.fieldInstruction).toBe('COMMENTS');
    expect(importedField?.textContent).toContain('Draft for counsel');

    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: imported.content.html,
    });
    try {
      expect(
        editor.commands.refreshDocumentFields(documentContent(editor), {
          resolveContext: () => ({
            pageNumber: 1,
            totalPages: 1,
            sectionNumber: 1,
            sectionPages: 1,
            comments: 'Ready for legal review',
          }),
        }),
      ).toBe(true);
      expect(
        documentFields(editor).map((entry) => ({
          kind: entry.kind,
          display: entry.display,
        })),
      ).toEqual([{ kind: 'comments', display: 'Ready for legal review' }]);
    } finally {
      editor.destroy();
    }
  });

  test('round-trips CREATEDATE as a live native field', async () => {
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    artifact.content.html = [
      '<section data-document-section="true"><p>',
      field(
        'createDate',
        'created',
        'CREATEDATE \\@ "yyyy-MM-dd"',
        '2020-01-01',
      ),
      '</p></section>',
    ].join('');

    const first = await createArtifactBlob(artifact);
    const archive = await JSZip.loadAsync(await first.arrayBuffer());
    const documentXml =
      (await archive.file('word/document.xml')?.async('string')) ?? '';
    expect(documentXml).toContain('CREATEDATE');
    expect(documentXml).toContain('2020-01-01');

    const imported = await importOfficeFile(
      new File([first], 'createdate-field.docx', { type: first.type }),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const document = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    const importedField = document.body.querySelector<HTMLElement>(
      '[data-document-field]',
    );
    expect(importedField?.dataset.fieldKind).toBe('createDate');
    expect(importedField?.dataset.fieldInstruction).toBe(
      'CREATEDATE \\@ "yyyy-MM-dd"',
    );
    expect(importedField?.textContent).toContain('2020-01-01');

    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: imported.content.html,
    });
    try {
      expect(
        editor.commands.refreshDocumentFields(documentContent(editor), {
          resolveContext: () => ({
            pageNumber: 1,
            totalPages: 1,
            sectionNumber: 1,
            sectionPages: 1,
            createDate: new Date(2024, 2, 15),
          }),
        }),
      ).toBe(true);
      expect(
        documentFields(editor).map((entry) => ({
          kind: entry.kind,
          display: entry.display,
        })),
      ).toEqual([{ kind: 'createDate', display: '2024-03-15' }]);
    } finally {
      editor.destroy();
    }
  });

  test('round-trips SAVEDATE as a live native field', async () => {
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    artifact.content.html = [
      '<section data-document-section="true"><p>',
      field(
        'saveDate',
        'saved',
        'SAVEDATE \\@ "yyyy-MM-dd"',
        '2021-02-02',
      ),
      '</p></section>',
    ].join('');

    const first = await createArtifactBlob(artifact);
    const archive = await JSZip.loadAsync(await first.arrayBuffer());
    const documentXml =
      (await archive.file('word/document.xml')?.async('string')) ?? '';
    expect(documentXml).toContain('SAVEDATE');
    expect(documentXml).toContain('2021-02-02');

    const imported = await importOfficeFile(
      new File([first], 'savedate-field.docx', { type: first.type }),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const document = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    const importedField = document.body.querySelector<HTMLElement>(
      '[data-document-field]',
    );
    expect(importedField?.dataset.fieldKind).toBe('saveDate');
    expect(importedField?.dataset.fieldInstruction).toBe(
      'SAVEDATE \\@ "yyyy-MM-dd"',
    );
    expect(importedField?.textContent).toContain('2021-02-02');

    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: imported.content.html,
    });
    try {
      expect(
        editor.commands.refreshDocumentFields(documentContent(editor), {
          resolveContext: () => ({
            pageNumber: 1,
            totalPages: 1,
            sectionNumber: 1,
            sectionPages: 1,
            saveDate: new Date(2025, 5, 8),
          }),
        }),
      ).toBe(true);
      expect(
        documentFields(editor).map((entry) => ({
          kind: entry.kind,
          display: entry.display,
        })),
      ).toEqual([{ kind: 'saveDate', display: '2025-06-08' }]);
    } finally {
      editor.destroy();
    }
  });

  test('round-trips PRINTDATE as a live native field', async () => {
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    artifact.content.html = [
      '<section data-document-section="true"><p>',
      field(
        'printDate',
        'printed',
        'PRINTDATE \\@ "yyyy-MM-dd"',
        '2022-03-03',
      ),
      '</p></section>',
    ].join('');

    const first = await createArtifactBlob(artifact);
    const archive = await JSZip.loadAsync(await first.arrayBuffer());
    const documentXml =
      (await archive.file('word/document.xml')?.async('string')) ?? '';
    expect(documentXml).toContain('PRINTDATE');
    expect(documentXml).toContain('2022-03-03');

    const imported = await importOfficeFile(
      new File([first], 'printdate-field.docx', { type: first.type }),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const document = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    const importedField = document.body.querySelector<HTMLElement>(
      '[data-document-field]',
    );
    expect(importedField?.dataset.fieldKind).toBe('printDate');
    expect(importedField?.dataset.fieldInstruction).toBe(
      'PRINTDATE \\@ "yyyy-MM-dd"',
    );
    expect(importedField?.textContent).toContain('2022-03-03');

    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: imported.content.html,
    });
    try {
      expect(
        editor.commands.refreshDocumentFields(documentContent(editor), {
          resolveContext: () => ({
            pageNumber: 1,
            totalPages: 1,
            sectionNumber: 1,
            sectionPages: 1,
            printDate: new Date(2023, 10, 20),
          }),
        }),
      ).toBe(true);
      expect(
        documentFields(editor).map((entry) => ({
          kind: entry.kind,
          display: entry.display,
        })),
      ).toEqual([{ kind: 'printDate', display: '2023-11-20' }]);
    } finally {
      editor.destroy();
    }
  });

  test('round-trips MERGEFIELD as a live native field', async () => {
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    artifact.content.html = [
      '<section data-document-section="true"><p>',
      '<span data-document-field="true" data-field-id="customer" data-field-kind="mergeField" data-field-instruction="MERGEFIELD CustomerName" data-field-target-name="CustomerName" data-field-display="«CustomerName»">«CustomerName»</span>',
      '</p></section>',
    ].join('');

    const first = await createArtifactBlob(artifact);
    const archive = await JSZip.loadAsync(await first.arrayBuffer());
    const documentXml =
      (await archive.file('word/document.xml')?.async('string')) ?? '';
    expect(documentXml).toContain('MERGEFIELD');
    expect(documentXml).toContain('CustomerName');

    const imported = await importOfficeFile(
      new File([first], 'mergefield.docx', { type: first.type }),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const document = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    const importedField = document.body.querySelector<HTMLElement>(
      '[data-document-field]',
    );
    expect(importedField?.dataset.fieldKind).toBe('mergeField');
    expect(importedField?.dataset.fieldInstruction).toBe(
      'MERGEFIELD CustomerName',
    );
    expect(importedField?.dataset.fieldTargetName).toBe('CustomerName');
    expect(importedField?.textContent).toContain('CustomerName');

    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: imported.content.html,
    });
    try {
      expect(
        editor.commands.refreshDocumentFields(documentContent(editor), {
          resolveContext: () => ({
            pageNumber: 1,
            totalPages: 1,
            sectionNumber: 1,
            sectionPages: 1,
            mergeRecord: { CustomerName: 'Ada Lovelace' },
          }),
        }),
      ).toBe(true);
      expect(
        documentFields(editor).map((entry) => ({
          kind: entry.kind,
          display: entry.display,
        })),
      ).toEqual([{ kind: 'mergeField', display: 'Ada Lovelace' }]);
    } finally {
      editor.destroy();
    }
  });

  test('round-trips WPS numeric page switches as live native fields', async () => {
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    artifact.content.html = [
      '<section data-document-section="true"><p>',
      field('page', 'roman-page', 'PAGE \\* ROMAN', 'XLII'),
      ' ',
      field('numPages', 'alpha-pages', 'NUMPAGES \\* ALPHABETIC', 'AP'),
      '</p><p>',
      '<span data-document-bookmark-boundary="true" data-bookmark-kind="start" data-bookmark-id="bookmark-target" data-bookmark-name="Target"></span>',
      '目标内容',
      '<span data-document-bookmark-boundary="true" data-bookmark-kind="end" data-bookmark-id="bookmark-target" data-bookmark-name="Target"></span>',
      ' ',
      field(
        'pageReference',
        'ordinal-page',
        'PAGEREF Target \\h \\* Ordinal',
        '7th',
      ),
      '</p></section>',
    ].join('');

    const first = await createArtifactBlob(artifact);
    const archive = await JSZip.loadAsync(await first.arrayBuffer());
    const documentXml =
      (await archive.file('word/document.xml')?.async('string')) ?? '';
    expect(documentXml).toContain('PAGE \\* ROMAN');
    expect(documentXml).toContain('NUMPAGES \\* ALPHABETIC');
    expect(documentXml).toContain('PAGEREF Target \\h \\* Ordinal');

    const imported = await importOfficeFile(
      new File([first], 'wps-numeric-fields.docx', { type: first.type }),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const importedDocument = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    expect(
      Array.from(
        importedDocument.body.querySelectorAll<HTMLElement>(
          '[data-document-field]',
        ),
      ).map(({ dataset }) => dataset.fieldInstruction),
    ).toEqual([
      'PAGE \\* ROMAN',
      'NUMPAGES \\* ALPHABETIC',
      'PAGEREF Target \\h \\* Ordinal',
    ]);
    await expectNativeFieldsWithSwitches(await createArtifactBlob(imported));
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

async function expectNativeCommonFields(blob: Blob): Promise<void> {
  const archive = await JSZip.loadAsync(await blob.arrayBuffer());
  const documentXml =
    (await archive.file('word/document.xml')?.async('string')) ?? '';
  expect(documentXml).toContain('NUMWORDS');
  expect(documentXml).toContain('NUMCHARS');
  expect(documentXml).toContain('PAGEREF Target \\h');
}

async function expectNativeFieldsWithSwitches(blob: Blob): Promise<void> {
  const archive = await JSZip.loadAsync(await blob.arrayBuffer());
  const documentXml =
    (await archive.file('word/document.xml')?.async('string')) ?? '';
  expect(documentXml).toContain('PAGE \\* ROMAN');
  expect(documentXml).toContain('NUMPAGES \\* ALPHABETIC');
  expect(documentXml).toContain('PAGEREF Target \\h \\* Ordinal');
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
    wordCount: 'NUMWORDS',
    characterCount: 'NUMCHARS',
    fileName: 'FILENAME',
    author: 'AUTHOR',
    title: 'TITLE',
    subject: 'SUBJECT',
    keywords: 'KEYWORDS',
    lastSavedBy: 'LASTSAVEDBY',
    comments: 'COMMENTS',
    createDate: 'CREATEDATE \\@ "yyyy年M月d日"',
    saveDate: 'SAVEDATE \\@ "yyyy年M月d日"',
    printDate: 'PRINTDATE \\@ "yyyy年M月d日"',
    mergeField: 'MERGEFIELD FieldName',
    pageReference: 'PAGEREF Target \\h',
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

function documentBookmarkBoundary(
  editor: Editor,
  kind: 'start' | 'end',
):
  | { node: ReturnType<Editor['state']['doc']['nodeAt']>; position: number }
  | undefined {
  let found:
    | { node: ReturnType<Editor['state']['doc']['nodeAt']>; position: number }
    | undefined;
  editor.state.doc.descendants((node, position) => {
    if (found || node.type.name !== 'documentBookmarkBoundary') return;
    if (node.attrs.kind === kind) found = { node, position };
  });
  return found;
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
