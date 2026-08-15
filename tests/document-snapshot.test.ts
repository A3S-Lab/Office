import { describe, expect, test } from '@rstest/core';
import {
  DOCUMENT_SNAPSHOT_MEDIA_TYPE,
  DOCUMENT_SNAPSHOT_SCHEMA,
  DOCUMENT_SNAPSHOT_VERSION,
  decodeDocumentSnapshot,
  encodeDocumentSnapshot,
  type DocumentContent,
} from '../src/core';
import { createWorkDocumentModel } from '../src/internal/features/work/work-document-model';

function documentContent(): DocumentContent {
  const html = [
    '<section data-document-section="true" data-section-id="section-1">',
    '<h1>Contract</h1><p><strong>Exact</strong> structured content.</p>',
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
          attrs: { id: 'section-1', pageSize: 'letter' },
          content: [
            {
              type: 'heading',
              attrs: { level: 1 },
              content: [{ type: 'text', text: 'Contract' }],
            },
            {
              type: 'paragraph',
              content: [
                {
                  type: 'text',
                  text: 'Exact',
                  marks: [{ type: 'bold' }],
                },
                { type: 'text', text: ' structured content.' },
              ],
            },
          ],
        },
      ],
    }),
    pageSize: 'letter',
    pageColor: '#fdfcf8',
    orientation: 'landscape',
    margins: { top: 18, right: 21, bottom: 19, left: 22 },
    columns: {
      count: 2,
      spacing: 14,
      separator: true,
      custom: [
        { widthPercent: 42, spacing: 12 },
        { widthPercent: 58, spacing: 12 },
      ],
    },
    headerText: 'Agreement',
    footerText: 'Confidential',
    showPageNumbers: true,
    pageNumberStart: 4,
    trackChanges: true,
    comments: [
      {
        id: 'comment-1',
        author: 'Reviewer',
        date: '2026-08-14T08:00:00.000Z',
        text: 'Keep this clause.',
        resolved: false,
        replies: [
          {
            id: 'reply-1',
            author: 'Author',
            date: '2026-08-14T08:01:00.000Z',
            text: 'Accepted.',
          },
        ],
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
          year: '2026',
          contributors: {
            author: { corporate: 'A3S Lab' },
          },
          additionalFields: { edition: '1' },
        },
      ],
    },
  };
}

describe('public structured document snapshot codec', () => {
  test('round-trips every JSON-defined document field without sharing state', () => {
    const source = documentContent();
    const encoded = encodeDocumentSnapshot(source);
    const decoded = decodeDocumentSnapshot(encoded);

    expect(JSON.parse(encoded)).toMatchObject({
      schema: DOCUMENT_SNAPSHOT_SCHEMA,
      version: DOCUMENT_SNAPSHOT_VERSION,
      content: source,
    });
    expect(decoded).toEqual(source);
    expect(decoded).not.toBe(source);
    expect(decoded.model).not.toBe(source.model);
    expect(DOCUMENT_SNAPSHOT_MEDIA_TYPE).toBe(
      'application/vnd.a3s.office.document-snapshot+json;version=1',
    );
  });

  test('emits one canonical representation for equivalent key ordering', () => {
    const first = documentContent();
    const second = {
      bibliography: first.bibliography,
      comments: first.comments,
      trackChanges: first.trackChanges,
      pageNumberStart: first.pageNumberStart,
      showPageNumbers: first.showPageNumbers,
      footerText: first.footerText,
      headerText: first.headerText,
      columns: first.columns,
      margins: first.margins,
      orientation: first.orientation,
      pageColor: first.pageColor,
      pageSize: first.pageSize,
      model: first.model,
      html: first.html,
      type: first.type,
    } satisfies DocumentContent;

    expect(encodeDocumentSnapshot(second)).toBe(encodeDocumentSnapshot(first));
  });

  test('normalizes optional undefined fields at the JSON boundary', () => {
    const source = documentContent();
    source.headerText = undefined;

    const decoded = decodeDocumentSnapshot(encodeDocumentSnapshot(source));

    expect(decoded).not.toHaveProperty('headerText');
    expect(decoded.model).toEqual(source.model);
  });

  test('rejects stale models, unsupported envelopes, and unsafe values', () => {
    const stale = documentContent();
    stale.html = '<p>Changed without its structured model</p>';
    expect(() => encodeDocumentSnapshot(stale)).toThrow(
      'synchronized structured model',
    );

    const current = JSON.parse(
      encodeDocumentSnapshot(documentContent()),
    ) as Record<string, unknown>;
    expect(() =>
      decodeDocumentSnapshot(JSON.stringify({ ...current, version: 2 })),
    ).toThrow('unsupported');
    expect(() => decodeDocumentSnapshot('{')).toThrow('valid JSON');

    const cyclic = documentContent() as DocumentContent & {
      cycle?: unknown;
    };
    cyclic.cycle = cyclic;
    expect(() => encodeDocumentSnapshot(cyclic)).toThrow('JSON-compatible');
  });
});
