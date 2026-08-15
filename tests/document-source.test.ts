import { describe, expect, test } from '@rstest/core';
import { getSchema } from '@tiptap/core';
import {
  applyDocumentSource,
  decodeDocumentSnapshot,
  DOCUMENT_SOURCE_MEDIA_TYPE,
  DOCUMENT_SOURCE_SCHEMA,
  DOCUMENT_SOURCE_VERSION,
  encodeDocumentSnapshot,
  projectDocumentSource,
  type DocumentContent,
} from '../src/core';
import { createWorkDocumentModel } from '../src/internal/features/work/work-document-model';
import { createWorkDocumentExtensions } from '../src/internal/features/work/work-document-extensions';
import {
  collectDocumentCommentAnchors,
  retainAnchoredDocumentComments,
} from '../src/internal/features/work/work-document-comments';

function documentContent(): DocumentContent {
  const html = [
    '<section data-document-section="true" data-section-id="section-1">',
    '<h1>Contract</h1><p>Original body.</p>',
    '</section>',
  ].join('');
  return {
    type: 'document',
    html,
    model: createWorkDocumentModel(html, {
      type: 'doc',
      content: [
        {
          type: 'documentSection',
          attrs: {
            id: 'section-1',
            pageSize: 'letter',
            orientation: 'landscape',
            marginTop: 18,
            marginRight: 19,
            marginBottom: 20,
            marginLeft: 21,
            columnCount: 1,
            columnSpacing: 12,
            columnSeparator: false,
            columnLayout: '{"count":1,"spacing":12,"separator":false}',
            breakAfter: 'nextPage',
            headerText: 'Agreement',
            footerText: 'Confidential',
            showPageNumbers: true,
            pageNumberStart: null,
            pageChrome:
              '{"differentFirstPage":false,"differentOddEvenPages":false,"default":{"headerHtml":"<p>Agreement</p>","footerHtml":"<p>Confidential</p>","showPageNumber":true},"first":{"headerHtml":"","footerHtml":"","showPageNumber":false},"even":{"headerHtml":"","footerHtml":"","showPageNumber":false}}',
          },
          content: [
            {
              type: 'heading',
              attrs: { level: 1 },
              content: [{ type: 'text', text: 'Contract' }],
            },
            {
              type: 'paragraph',
              content: [{ type: 'text', text: 'Original body.' }],
            },
          ],
        },
      ],
    }),
    pageSize: 'letter',
    orientation: 'landscape',
    margins: { top: 18, right: 19, bottom: 20, left: 21 },
    headerText: 'Agreement',
    footerText: 'Confidential',
    showPageNumbers: true,
    trackChanges: true,
    comments: [
      {
        id: 'comment-1',
        author: 'Reviewer',
        date: '2026-08-14T08:00:00.000Z',
        text: 'Retain this review state.',
        resolved: false,
        replies: [],
      },
    ],
    bibliography: {
      style: 'apa',
      sources: [
        {
          id: 'source-1',
          tag: 'A3S2026',
          sourceType: 'Report',
          title: 'A3S Office',
        },
      ],
    },
  };
}

describe('public structured document source contract', () => {
  test('projects readable Markdown without replacing the lossless snapshot', () => {
    const content = documentContent();
    const source = projectDocumentSource(content);

    expect(source).toEqual({
      schema: DOCUMENT_SOURCE_SCHEMA,
      version: DOCUMENT_SOURCE_VERSION,
      mediaType: DOCUMENT_SOURCE_MEDIA_TYPE,
      content: '# Contract\n\nOriginal body.',
    });
    expect(decodeDocumentSnapshot(encodeDocumentSnapshot(content))).toEqual(
      content,
    );
  });

  test('replaces the semantic body while retaining Office-only state', () => {
    const original = documentContent();
    const next = applyDocumentSource(original, {
      schema: DOCUMENT_SOURCE_SCHEMA,
      version: DOCUMENT_SOURCE_VERSION,
      mediaType: DOCUMENT_SOURCE_MEDIA_TYPE,
      content: [
        '# Revised contract',
        '',
        'The accepted **structured** body.',
        '',
        '- First obligation',
        '- Second obligation',
      ].join('\n'),
    });

    expect(projectDocumentSource(next).content).toContain('# Revised contract');
    expect(next.html).toContain(
      'The accepted <strong>structured</strong> body.',
    );
    expect(next.pageSize).toBe(original.pageSize);
    expect(next.orientation).toBe(original.orientation);
    expect(next.margins).toEqual(original.margins);
    expect(next.headerText).toBe(original.headerText);
    expect(next.footerText).toBe(original.footerText);
    expect(next.showPageNumbers).toBe(true);
    expect(next.trackChanges).toBe(true);
    expect(next.comments).toEqual(original.comments);
    expect(next.bibliography).toEqual(original.bibliography);
    expect(next.model?.root.content?.[0]?.attrs).toMatchObject(
      original.model?.root.content?.[0]?.attrs,
    );
    if (!next.model)
      throw new Error('The revised document model is unavailable.');
    expect(
      getSchema(createWorkDocumentExtensions())
        .nodeFromJSON(next.model.root)
        .toJSON(),
    ).toEqual(next.model?.root);
    expect(decodeDocumentSnapshot(encodeDocumentSnapshot(next))).toEqual(next);
  });

  test('reattaches a surviving Office comment anchor after an Agent revision', () => {
    const original = documentContent();
    const paragraph = original.model?.root.content?.[0]?.content?.[1];
    const text = paragraph?.content?.[0];
    const originalModel = original.model;
    if (!text || !originalModel) {
      throw new Error('The test document body is unavailable.');
    }
    text.marks = [{ type: 'documentComment', attrs: { id: 'comment-1' } }];
    original.model = createWorkDocumentModel(
      original.html,
      originalModel.root,
      originalModel,
    );

    const next = applyDocumentSource(original, {
      schema: DOCUMENT_SOURCE_SCHEMA,
      version: DOCUMENT_SOURCE_VERSION,
      mediaType: DOCUMENT_SOURCE_MEDIA_TYPE,
      content: [
        '# Revised contract',
        '',
        'Agent context inserted before the reviewed text.',
        '',
        'Original body.',
        '',
        'Agent context inserted after the reviewed text.',
      ].join('\n'),
    });
    if (!next.model)
      throw new Error('The revised document model is unavailable.');
    const document = getSchema(createWorkDocumentExtensions()).nodeFromJSON(
      next.model.root,
    );
    const anchors = collectDocumentCommentAnchors(document);

    expect(anchors).toHaveLength(1);
    expect(anchors[0]).toMatchObject({
      id: 'comment-1',
      anchorText: 'Original body.',
    });
    expect(
      retainAnchoredDocumentComments(next.comments ?? [], anchors),
    ).toEqual(original.comments);
  });

  test('does not guess when a surviving comment anchor is ambiguous', () => {
    const original = documentContent();
    const paragraph = original.model?.root.content?.[0]?.content?.[1];
    const text = paragraph?.content?.[0];
    const originalModel = original.model;
    if (!text || !originalModel) {
      throw new Error('The test document body is unavailable.');
    }
    text.marks = [{ type: 'documentComment', attrs: { id: 'comment-1' } }];
    original.model = createWorkDocumentModel(
      original.html,
      originalModel.root,
      originalModel,
    );

    const next = applyDocumentSource(original, {
      schema: DOCUMENT_SOURCE_SCHEMA,
      version: DOCUMENT_SOURCE_VERSION,
      mediaType: DOCUMENT_SOURCE_MEDIA_TYPE,
      content: [
        '# Revised contract',
        '',
        'Original body.',
        '',
        'Original body.',
      ].join('\n'),
    });
    if (!next.model)
      throw new Error('The revised document model is unavailable.');
    const document = getSchema(createWorkDocumentExtensions()).nodeFromJSON(
      next.model.root,
    );

    expect(collectDocumentCommentAnchors(document)).toEqual([]);
    expect(next.comments).toEqual(original.comments);
  });

  test('fails closed instead of flattening multiple section layouts', () => {
    const content = documentContent();
    const model = content.model;
    const sections = model?.root.content;
    if (!model || !sections) {
      throw new Error('The test document sections are unavailable.');
    }
    sections.push({
      type: 'documentSection',
      attrs: { id: 'section-2', pageSize: 'a4' },
      content: [{ type: 'paragraph' }],
    });
    content.model = createWorkDocumentModel(content.html, model.root, model);

    expect(() => projectDocumentSource(content)).toThrow(
      'exactly one structured section',
    );
  });
});
