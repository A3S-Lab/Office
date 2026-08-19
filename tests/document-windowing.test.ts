import { describe, expect, test } from '@rstest/core';
import {
  documentModelUsesWindowing,
  materializeWindowedDocumentModel,
  windowDocumentModel,
} from '../src/internal/features/work/work-document-windowing';
import { serializeWorkDocumentNode } from '../src/internal/features/work/work-document-model-codec';
import type { WorkDocumentNode } from '../src/internal/features/work/work-types';

describe('document model windowing', () => {
  test('groups large section bodies without changing their canonical model', () => {
    const root = documentRoot([
      paragraph('one'),
      paragraph('two'),
      paragraph('three'),
      paragraph('four'),
      paragraph('five'),
    ]);

    const windowed = windowDocumentModel(root, {
      blockSize: 2,
      blockThreshold: 4,
      tableRowSize: 2,
      tableRowThreshold: 4,
    });

    expect(documentModelUsesWindowing(windowed)).toBe(true);
    expect(windowed.content?.[0]?.content).toHaveLength(3);
    expect(
      windowed.content?.[0]?.content?.map((node) => node.attrs?.blockCount),
    ).toEqual([2, 2, 1]);
    expect(materializeWindowedDocumentModel(windowed)).toEqual(root);
    expect(windowDocumentModel(windowed)).toBe(windowed);
  });

  test('slices giant tables and joins them again for controlled HTML and export', () => {
    const table = tableNode(5);
    const root = documentRoot([paragraph('before'), table, paragraph('after')]);

    const windowed = windowDocumentModel(root, {
      blockSize: 2,
      blockThreshold: 100,
      tableRowSize: 2,
      tableRowThreshold: 4,
    });
    const chunks = windowed.content?.[0]?.content ?? [];
    const tables = chunks.flatMap((chunk) =>
      (chunk.content ?? []).filter((node) => node.type === 'table'),
    );

    expect(chunks.every((node) => node.type === 'documentChunk')).toBe(true);
    expect(tables).toHaveLength(3);
    expect(tables.map((node) => node.content?.length)).toEqual([2, 2, 1]);
    expect(tables.map((node) => node.attrs?.virtualTableIndex)).toEqual([
      0, 1, 2,
    ]);
    expect(materializeWindowedDocumentModel(windowed)).toEqual(root);
  });

  test('groups many table slices without changing canonical rows', () => {
    const root = documentRoot([tableNode(160)]);
    const windowed = windowDocumentModel(root, {
      blockSize: 2,
      blockThreshold: 100,
      tableRowSize: 2,
      tableRowThreshold: 4,
    });
    const containers = windowed.content?.[0]?.content ?? [];
    const leaves = containers.flatMap((container) => container.content ?? []);

    expect(containers).toHaveLength(3);
    expect(
      containers.every(
        (container) => container.attrs?.windowContainer === true,
      ),
    ).toBe(true);
    expect(leaves).toHaveLength(80);
    expect(leaves.every((leaf) => leaf.type === 'documentChunk')).toBe(true);
    expect(materializeWindowedDocumentModel(windowed)).toEqual(root);
  });

  test('estimates table rows across text runs without concatenating cells', () => {
    const root = documentRoot([
      {
        type: 'table',
        attrs: { officeImported: true },
        content: [
          {
            type: 'tableRow',
            content: [
              {
                type: 'tableCell',
                content: [
                  {
                    type: 'paragraph',
                    content: [
                      { type: 'text', text: 'a'.repeat(44) },
                      { type: 'text', text: `${'b'.repeat(45)}\nnext` },
                    ],
                  },
                ],
              },
            ],
          },
        ],
      },
    ]);

    const windowed = windowDocumentModel(root, {
      blockSize: 1,
      blockThreshold: 1,
      tableRowSize: 1,
      tableRowThreshold: 1,
    });

    expect(windowed.content?.[0]?.content?.[0]?.attrs?.estimatedHeight).toBe(
      64,
    );
  });

  test('leaves ordinary documents untouched', () => {
    const root = documentRoot([paragraph('small')]);
    expect(windowDocumentModel(root)).toBe(root);
    expect(documentModelUsesWindowing(root)).toBe(false);
  });

  test('serializes the canonical document instead of internal window nodes', () => {
    const root = windowDocumentModel(documentRoot([tableNode(5)]), {
      blockSize: 2,
      blockThreshold: 100,
      tableRowSize: 2,
      tableRowThreshold: 4,
    });

    const html = serializeWorkDocumentNode(root);
    expect(html).not.toContain('data-document-chunk');
    expect(html).not.toContain('data-document-virtual-table');
    expect(html.match(/<table/g)).toHaveLength(1);
    expect(html.match(/<tr/g)).toHaveLength(5);
  });
});

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
    content: [{ type: 'text', text }],
  };
}

function tableNode(rowCount: number): WorkDocumentNode {
  return {
    type: 'table',
    attrs: { officeImported: true },
    content: Array.from({ length: rowCount }, (_, index) => ({
      type: 'tableRow',
      content: [
        {
          type: 'tableCell',
          content: [paragraph(`row ${index + 1}`)],
        },
      ],
    })),
  };
}
