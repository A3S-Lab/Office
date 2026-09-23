import { afterEach, describe, expect, test } from '@rstest/core';
import { Editor } from '@tiptap/core';
import { createWorkDocumentExtensions } from '../src/internal/features/work/work-document-extensions';
import {
  activeMailMergeRecord,
  createMailMergeFieldContextResolver,
  filteredMailMergeRecords,
  generateMailMergeDocuments,
  mailMergeFieldNamesFromEditor,
  mailMergeFieldNamesFromHtml,
  mailMergeFieldNamesFromRecords,
  mailMergeRecordMatchesFilter,
  normalizeMailMergeRecipientFilter,
  normalizeMailMergeSource,
  previewMailMergeSource,
  setMailMergeRecipientFilter,
  stepMailMergeSource,
} from '../src/internal/features/work/work-document-mail-merge';
import type { WorkDocumentContent } from '../src/internal/features/work/work-types';

let editor: Editor | null = null;

afterEach(() => {
  editor?.destroy();
  editor = null;
});

function documentContent(current: Editor): WorkDocumentContent {
  return {
    type: 'document',
    html: current.getHTML(),
    comments: [],
  };
}

describe('document mail-merge host runner', () => {
  test('normalizes records, clamps active index, and steps preview rows', () => {
    expect(
      normalizeMailMergeSource({
        records: [
          { CustomerName: 'Ada', '123bad': 'x', ok_1: 'y' },
          null as unknown as Record<string, string>,
          { CustomerName: 'Grace' },
        ],
        activeIndex: 99,
      }),
    ).toEqual({
      records: [{ CustomerName: 'Ada', ok_1: 'y' }, { CustomerName: 'Grace' }],
      activeIndex: 1,
      filter: null,
    });
    expect(
      activeMailMergeRecord({
        records: [{ A: '1' }, { A: '2' }],
        activeIndex: 0,
      }),
    ).toEqual({ A: '1' });
    expect(
      stepMailMergeSource(
        { records: [{ A: '1' }, { A: '2' }, { A: '3' }], activeIndex: 0 },
        2,
      ),
    ).toEqual({
      records: [{ A: '1' }, { A: '2' }, { A: '3' }],
      activeIndex: 2,
      filter: null,
    });
    expect(
      stepMailMergeSource(
        { records: [{ A: '1' }, { A: '2' }], activeIndex: 0 },
        -1,
      ).activeIndex,
    ).toBe(0);
  });

  test('lists MERGEFIELD names and previews the active host record', () => {
    const html = [
      '<section data-document-section="true"><p>',
      '<span data-document-field="true" data-field-id="customer" data-field-kind="mergeField" data-field-instruction="MERGEFIELD CustomerName" data-field-target-name="CustomerName" data-field-display="«CustomerName»">«CustomerName»</span>',
      ' ',
      '<span data-document-field="true" data-field-id="city" data-field-kind="mergeField" data-field-instruction="MERGEFIELD City" data-field-target-name="City" data-field-display="«City»">«City»</span>',
      ' ',
      '<span data-document-field="true" data-field-id="page" data-field-kind="page" data-field-instruction="PAGE" data-field-display="1">1</span>',
      '</p></section>',
    ].join('');
    expect(mailMergeFieldNamesFromHtml(html)).toEqual([
      'CustomerName',
      'City',
    ]);

    editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: html,
    });
    expect(mailMergeFieldNamesFromEditor(editor)).toEqual([
      'CustomerName',
      'City',
    ]);

    const source = normalizeMailMergeSource({
      records: [
        { CustomerName: 'Ada Lovelace', City: 'London' },
        { CustomerName: 'Grace Hopper', City: 'New York' },
      ],
      activeIndex: 0,
    });
    expect(
      previewMailMergeSource(editor, documentContent(editor), source),
    ).toBe(true);
    expect(editor.getText()).toContain('Ada Lovelace');
    expect(editor.getText()).toContain('London');

    const next = stepMailMergeSource(source, 1);
    expect(
      previewMailMergeSource(editor, documentContent(editor), next),
    ).toBe(true);
    expect(editor.getText()).toContain('Grace Hopper');
    expect(editor.getText()).toContain('New York');
  });

  test('layers merge records onto a base field-context resolver', () => {
    const resolve = createMailMergeFieldContextResolver(
      () => ({
        pageNumber: 3,
        totalPages: 9,
        sectionNumber: 2,
        sectionPages: 4,
        fileName: 'letter.docx',
      }),
      {
        records: [{ CustomerName: 'Ada' }],
        activeIndex: 0,
      },
    );
    expect(resolve?.(10)).toEqual({
      pageNumber: 3,
      totalPages: 9,
      sectionNumber: 2,
      sectionPages: 4,
      fileName: 'letter.docx',
      mergeRecord: { CustomerName: 'Ada' },
    });
    expect(createMailMergeFieldContextResolver(null, null)).toBeNull();
    expect(
      createMailMergeFieldContextResolver(null, {
        records: [{ A: '1' }],
        activeIndex: 0,
      })?.(0),
    ).toMatchObject({
      pageNumber: 1,
      mergeRecord: { A: '1' },
    });
  });

  test('filters recipients with AND rules and clamps the active preview row', () => {
    const source = normalizeMailMergeSource({
      records: [
        { CustomerName: 'Ada', City: 'London' },
        { CustomerName: 'Grace', City: 'New York' },
        { CustomerName: 'Alan', City: 'London' },
        { CustomerName: '', City: 'Paris' },
      ],
      activeIndex: 3,
      filter: {
        rules: [
          { field: 'City', operator: 'equals', value: 'London' },
          { field: 'CustomerName', operator: 'isNotBlank' },
        ],
      },
    });
    expect(filteredMailMergeRecords(source)).toEqual([
      { CustomerName: 'Ada', City: 'London' },
      { CustomerName: 'Alan', City: 'London' },
    ]);
    expect(activeMailMergeRecord(source)).toEqual({
      CustomerName: 'Alan',
      City: 'London',
    });
    expect(source.activeIndex).toBe(1);
    expect(
      stepMailMergeSource(source, -1).activeIndex,
    ).toBe(0);
    expect(
      mailMergeRecordMatchesFilter(
        { CustomerName: 'Ada', City: 'London' },
        normalizeMailMergeRecipientFilter({
          rules: [{ field: 'City', operator: 'contains', value: 'ond' }],
        }),
      ),
    ).toBe(true);
    expect(
      mailMergeFieldNamesFromRecords(source.records),
    ).toEqual(['CustomerName', 'City']);
    const cleared = setMailMergeRecipientFilter(source, null);
    expect(cleared.filter).toBeNull();
    expect(filteredMailMergeRecords(cleared)).toHaveLength(4);
  });

  test('batch generation respects the active recipient filter', async () => {
    const html = [
      '<section data-document-section="true"><p>',
      '<span data-document-field="true" data-field-id="customer" data-field-kind="mergeField" data-field-instruction="MERGEFIELD CustomerName" data-field-target-name="CustomerName" data-field-display="«CustomerName»">«CustomerName»</span>',
      '</p></section>',
    ].join('');
    editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: html,
    });
    const template = documentContent(editor);
    const source = normalizeMailMergeSource({
      records: [
        { CustomerName: 'Ada Lovelace', City: 'London' },
        { CustomerName: 'Grace Hopper', City: 'New York' },
        { CustomerName: 'Alan Turing', City: 'London' },
      ],
      activeIndex: 0,
      filter: {
        rules: [{ field: 'City', operator: 'equals', value: 'London' }],
      },
    });
    const generated = await generateMailMergeDocuments(editor, template, source, {
      exportDocument: async (next) =>
        new Blob([next.html], {
          type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        }),
    });
    expect(generated).toHaveLength(2);
    expect(generated.map((entry) => entry.record.CustomerName)).toEqual([
      'Ada Lovelace',
      'Alan Turing',
    ]);
  });

  test('generates one DOCX blob per recipient and restores the preview row', async () => {
    const html = [
      '<section data-document-section="true"><p>',
      '<span data-document-field="true" data-field-id="customer" data-field-kind="mergeField" data-field-instruction="MERGEFIELD CustomerName" data-field-target-name="CustomerName" data-field-display="«CustomerName»">«CustomerName»</span>',
      '</p></section>',
    ].join('');
    editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: html,
    });
    const template = documentContent(editor);
    const source = normalizeMailMergeSource({
      records: [
        { CustomerName: 'Ada Lovelace' },
        { CustomerName: 'Grace Hopper' },
      ],
      activeIndex: 0,
    });
    expect(previewMailMergeSource(editor, template, source)).toBe(true);
    expect(editor.getText()).toContain('Ada Lovelace');

    const exported: string[] = [];
    const generated = await generateMailMergeDocuments(editor, template, source, {
      exportDocument: async (next) => {
        exported.push(next.html);
        return new Blob([next.html], {
          type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        });
      },
    });
    expect(generated).toHaveLength(2);
    expect(generated[0]?.record).toEqual({ CustomerName: 'Ada Lovelace' });
    expect(generated[1]?.record).toEqual({ CustomerName: 'Grace Hopper' });
    expect(exported[0]).toContain('Ada Lovelace');
    expect(exported[1]).toContain('Grace Hopper');
    expect(editor.getText()).toContain('Ada Lovelace');
    expect(editor.getText()).not.toContain('Grace Hopper');
  });
});
