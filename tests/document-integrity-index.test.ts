import { describe, expect, test } from '@rstest/core';
import { createNodeFromContent } from '@tiptap/core';
import type { Node as ProseMirrorNode } from '@tiptap/pm/model';
import {
  DOCUMENT_INTEGRITY_BOOKMARK,
  DOCUMENT_INTEGRITY_FIELD,
  DOCUMENT_INTEGRITY_IMAGE,
  DOCUMENT_INTEGRITY_NOTE,
  DOCUMENT_INTEGRITY_PARAGRAPH_IDENTITY,
  DOCUMENT_INTEGRITY_TABLE_ROW_IDENTITY,
  documentHasIntegrityFeature,
} from '../src/internal/features/work/work-document-integrity-index';
import { workDocumentSchema } from '../src/internal/features/work/work-document-model-codec';
import { windowDocumentModel } from '../src/internal/features/work/work-document-windowing';
import type { WorkDocumentNode } from '../src/internal/features/work/work-types';

describe('document integrity feature index', () => {
  test('reports a plain windowed document without allocating feature indexes', () => {
    const document = prosemirrorDocument(
      windowDocumentModel(documentRoot([paragraph('plain'), table(false)]), {
        blockSize: 1,
        blockThreshold: 1,
        tableRowThreshold: 100,
      }),
    );

    for (const feature of [
      DOCUMENT_INTEGRITY_BOOKMARK,
      DOCUMENT_INTEGRITY_FIELD,
      DOCUMENT_INTEGRITY_IMAGE,
      DOCUMENT_INTEGRITY_NOTE,
      DOCUMENT_INTEGRITY_PARAGRAPH_IDENTITY,
      DOCUMENT_INTEGRITY_TABLE_ROW_IDENTITY,
    ]) {
      expect(documentHasIntegrityFeature(document, feature)).toBe(false);
    }
  });

  test('detects every normalization feature inside document chunks', () => {
    const document = prosemirrorDocument(
      windowDocumentModel(
        documentRoot([
          {
            type: 'paragraph',
            attrs: {
              paragraphId: '1A2B3C4D',
              textId: '1A2B3C4E',
            },
            content: [
              {
                type: 'documentBookmarkBoundary',
                attrs: {
                  id: 'bookmark-1',
                  kind: 'start',
                  name: 'bookmark',
                  nativeId: 1,
                },
              },
              { type: 'documentField', attrs: { id: 'field-1' } },
              {
                type: 'documentNoteReference',
                attrs: { id: 'note-1', kind: 'footnote', number: 1 },
              },
              { type: 'text', text: 'rich' },
              {
                type: 'documentBookmarkBoundary',
                attrs: {
                  id: 'bookmark-1',
                  kind: 'end',
                  name: 'bookmark',
                  nativeId: 1,
                },
              },
            ],
          },
          {
            type: 'image',
            attrs: { src: 'data:image/png;base64,AA==' },
          },
          table(true),
          {
            type: 'documentNote',
            attrs: { id: 'note-1', kind: 'footnote', number: 1 },
            content: [paragraph('note')],
          },
        ]),
        {
          blockSize: 2,
          blockThreshold: 1,
          tableRowThreshold: 100,
        },
      ),
    );

    for (const feature of [
      DOCUMENT_INTEGRITY_BOOKMARK,
      DOCUMENT_INTEGRITY_FIELD,
      DOCUMENT_INTEGRITY_IMAGE,
      DOCUMENT_INTEGRITY_NOTE,
      DOCUMENT_INTEGRITY_PARAGRAPH_IDENTITY,
      DOCUMENT_INTEGRITY_TABLE_ROW_IDENTITY,
    ]) {
      expect(documentHasIntegrityFeature(document, feature)).toBe(true);
    }
  });
});

function prosemirrorDocument(root: WorkDocumentNode): ProseMirrorNode {
  return createNodeFromContent(root, workDocumentSchema(), {
    errorOnInvalidContent: true,
    slice: false,
  }) as ProseMirrorNode;
}

function documentRoot(content: WorkDocumentNode[]): WorkDocumentNode {
  return {
    type: 'doc',
    content: [
      {
        type: 'documentSection',
        attrs: { id: 'document-section-1' },
        content,
      },
    ],
  };
}

function paragraph(text: string): WorkDocumentNode {
  return {
    type: 'paragraph',
    ...(text ? { content: [{ type: 'text', text }] } : {}),
  };
}

function table(withIdentity: boolean): WorkDocumentNode {
  return {
    type: 'table',
    content: [
      {
        type: 'tableRow',
        ...(withIdentity
          ? { attrs: { rowId: '2A2B3C4D', rowTextId: '2A2B3C4E' } }
          : {}),
        content: [
          {
            type: 'tableCell',
            content: [paragraph('cell')],
          },
        ],
      },
    ],
  };
}
