import { describe, expect, test } from '@rstest/core';
import {
  createDocumentLazyHtmlProjection,
  documentLazyHtmlChunkFragment,
  documentLazyHtmlProjectionFingerprint,
  patchDocumentLazyHtmlProjection,
} from '../src/internal/features/work/work-document-lazy-html';
import { documentHtmlFingerprint } from '../src/internal/features/work/work-document-model';
import { windowDocumentModel } from '../src/internal/features/work/work-document-windowing';
import type { WorkDocumentNode } from '../src/internal/features/work/work-types';

describe('lazy document HTML projection', () => {
  test('patches one paragraph chunk and shifts following ranges', () => {
    const root = windowDocumentModel(paragraphRoot(6), {
      blockSize: 2,
      blockThreshold: 4,
    });
    const html =
      '<section data-document-section="true"><p>one</p><p>two</p><p>three</p><p>four</p><p>five</p><p>six</p></section>';
    const projection = createDocumentLazyHtmlProjection(html, root);
    if (!projection) throw new Error('Expected a paragraph projection.');
    const secondId = projection.orderedRanges[1]?.id;
    if (!secondId) throw new Error('Expected a second paragraph chunk.');

    expect(
      patchDocumentLazyHtmlProjection(
        projection,
        new Map([[secondId, '<p>THREE!</p><p>FOUR!</p>']]),
      ),
    ).toBe(
      '<section data-document-section="true"><p>one</p><p>two</p><p>THREE!</p><p>FOUR!</p><p>five</p><p>six</p></section>',
    );
    expect(projection.ranges.get(secondId)).toMatchObject({
      from: html.indexOf('<p>three</p>'),
      to: html.indexOf('<p>three</p>') + '<p>THREE!</p><p>FOUR!</p>'.length,
    });
    const third = projection.orderedRanges[2];
    expect(projection.html.slice(third?.from, third?.to)).toBe(
      '<p>five</p><p>six</p>',
    );
    expect(documentLazyHtmlProjectionFingerprint(projection)).toBe(
      documentHtmlFingerprint(projection.html),
    );
  });

  test('indexes and slices virtual table boundaries without duplicating tags', () => {
    const root = windowDocumentModel(tableRoot(5), {
      blockThreshold: 99,
      tableRowSize: 2,
      tableRowThreshold: 4,
    });
    const html =
      '<section data-document-section="true"><table data-office-table-imported="true"><tbody><tr><td><p>1</p></td></tr><tr><td><p>2</p></td></tr><tr><td><p>3</p></td></tr><tr><td><p>4</p></td></tr><tr><td><p>5</p></td></tr></tbody></table></section>';
    const projection = createDocumentLazyHtmlProjection(html, root);

    expect(projection?.orderedRanges.map(({ tablePart }) => tablePart)).toEqual(
      ['first', 'middle', 'last'],
    );
    expect(
      documentLazyHtmlChunkFragment(
        '<table data-office-table-imported="true"><tbody><tr><td><p>3 edited</p></td></tr><tr><td><p>4</p></td></tr></tbody></table>',
        'middle',
      ),
    ).toBe('<tr><td><p>3 edited</p></td></tr><tr><td><p>4</p></td></tr>');
  });
});

function paragraphRoot(count: number): WorkDocumentNode {
  const words = ['one', 'two', 'three', 'four', 'five', 'six'];
  return {
    type: 'doc',
    content: [
      {
        type: 'documentSection',
        attrs: { id: 'document-section-1' },
        content: Array.from({ length: count }, (_, index) => ({
          type: 'paragraph',
          content: [{ type: 'text', text: words[index] ?? String(index) }],
        })),
      },
    ],
  };
}

function tableRoot(rowCount: number): WorkDocumentNode {
  return {
    type: 'doc',
    content: [
      {
        type: 'documentSection',
        attrs: { id: 'document-section-1' },
        content: [
          {
            type: 'table',
            attrs: { officeImported: true },
            content: Array.from({ length: rowCount }, (_, index) => ({
              type: 'tableRow',
              content: [
                {
                  type: 'tableCell',
                  content: [
                    {
                      type: 'paragraph',
                      content: [{ type: 'text', text: String(index + 1) }],
                    },
                  ],
                },
              ],
            })),
          },
        ],
      },
    ],
  };
}
