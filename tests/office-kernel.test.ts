import { describe, expect, test } from '@rstest/core';
import { createOfficeKernelClient } from '../src/internal/kernel/office-kernel-client';
import { layoutOfficeDocumentInJavaScript } from '../src/internal/kernel/office-kernel-fallback';
import { alignOfficePresentationInJavaScript } from '../src/internal/kernel/office-kernel-presentation-fallback';
import { layoutOfficeTextInJavaScript } from '../src/internal/kernel/office-kernel-text-fallback';
import {
  isOfficeKernelResponse,
  type OfficeKernelLayoutBlock,
  type OfficeKernelLayoutRequest,
  type OfficeKernelTextLayoutRequest,
  OFFICE_KERNEL_PROTOCOL_VERSION,
} from '../src/internal/kernel/office-kernel-protocol';

describe('Office layout kernel', () => {
  test('paginates blocks and reports the visual page spacer', () => {
    const result = layoutOfficeDocumentInJavaScript(
      request([block('one', 500), block('two', 500), block('three', 120)]),
    );

    expect(result.pages).toHaveLength(2);
    expect(result.pages[0].placements.map((item) => item.blockId)).toEqual([
      'one',
    ]);
    expect(result.pages[1].placements.map((item) => item.blockId)).toEqual([
      'two',
      'three',
    ]);
    expect(result.breaks[0]).toMatchObject({
      beforeBlockId: 'two',
      pageIndex: 1,
      remainingBodyHeight: 463,
      spacerHeight: 647,
    });
    expect(result.documentRevision).toBe(4);
  });

  test('keeps a heading with its following block when they fit together', () => {
    const result = layoutOfficeDocumentInJavaScript(
      request([
        block('intro', 850),
        { ...block('heading', 40), keepWithNext: true },
        block('paragraph', 100),
      ]),
    );

    expect(result.pages).toHaveLength(2);
    expect(result.pages[1].placements.map((item) => item.blockId)).toEqual([
      'heading',
      'paragraph',
    ]);
  });

  test('aligns presentation geometry relative to the slide', () => {
    const result = alignOfficePresentationInJavaScript({
      protocol: OFFICE_KERNEL_PROTOCOL_VERSION,
      kind: 'presentationGeometry',
      requestId: 8,
      revision: 3,
      documentRevision: 2,
      operation: { type: 'alignToSlide', alignment: 'center' },
      elements: [
        {
          id: 'title',
          x: 17,
          y: 23,
          width: 40,
          height: 20,
        },
      ],
    });

    expect(result.kind).toBe('presentationGeometryResult');
    expect(result.elements[0]).toEqual({
      id: 'title',
      x: 30,
      y: 23,
      width: 40,
      height: 20,
    });
  });

  test('lays out a suffix with absolute page indices', () => {
    const input = request([
      block('three', 500),
      block('four', 500),
      block('five', 120),
    ]);
    input.startPageIndex = 2;

    const result = layoutOfficeDocumentInJavaScript(input);

    expect(result.startPageIndex).toBe(2);
    expect(result.pages.map((page) => page.index)).toEqual([2, 3]);
    expect(result.breaks[0]).toMatchObject({
      beforeBlockId: 'four',
      pageIndex: 3,
    });
  });

  test('keeps minimum line fragments on both sides of a page break', () => {
    const result = layoutOfficeDocumentInJavaScript(
      request([
        block('intro', 843),
        ...Array.from({ length: 3 }, (_, flowIndex) => ({
          ...block(`paragraph-line-${flowIndex}`, 60),
          flowId: 'paragraph',
          flowIndex,
          flowCount: 3,
          minimumFragmentsPerPage: 2,
        })),
      ]),
    );

    expect(result.pages).toHaveLength(2);
    expect(result.pages[0].placements.map((item) => item.blockId)).toEqual([
      'intro',
    ]);
    expect(result.pages[1].placements.map((item) => item.blockId)).toEqual([
      'paragraph-line-0',
      'paragraph-line-1',
      'paragraph-line-2',
    ]);
  });

  test('keeps the end of a paragraph flow with the next paragraph', () => {
    const current = Array.from({ length: 2 }, (_, flowIndex) => ({
      ...block(`current-line-${flowIndex}`, 40),
      flowId: 'current',
      flowIndex,
      flowCount: 2,
      minimumFragmentsPerPage: 1,
      keepWithNext: flowIndex === 1,
    }));
    const next = Array.from({ length: 2 }, (_, flowIndex) => ({
      ...block(`next-line-${flowIndex}`, 50),
      flowId: 'next',
      flowIndex,
      flowCount: 2,
      minimumFragmentsPerPage: 2,
    }));

    const result = layoutOfficeDocumentInJavaScript(
      request([block('intro', 800), ...current, ...next]),
    );

    expect(result.pages[0].placements.map((item) => item.blockId)).toEqual([
      'intro',
    ]);
    expect(result.pages[1].placements.map((item) => item.blockId)).toEqual([
      'current-line-0',
      'current-line-1',
      'next-line-0',
      'next-line-1',
    ]);
  });

  test('reserves repeated table headers on continuation pages', () => {
    const table = [
      { id: 'table-header', height: 20 },
      { id: 'table-row-1', height: 40 },
      { id: 'table-row-2', height: 40 },
      { id: 'table-row-3', height: 40 },
    ].map((candidate, flowIndex, rows) => ({
      ...candidate,
      flowId: 'table',
      flowIndex,
      flowCount: rows.length,
      minimumFragmentsPerPage: 1,
      repeatHeaderCount: 1,
      repeatHeaderHeight: 20,
    }));
    const input = request([block('intro', 20), ...table]);
    input.page.height = 100;
    input.page.marginTop = 0;
    input.page.marginBottom = 0;

    const result = layoutOfficeDocumentInJavaScript(input);

    expect(result.pages).toHaveLength(2);
    expect(result.pages[0].placements.map((item) => item.blockId)).toEqual([
      'intro',
      'table-header',
      'table-row-1',
    ]);
    expect(result.pages[1].placements).toEqual([
      {
        blockId: 'table-row-2',
        y: 20,
        height: 40,
        overflow: false,
      },
      {
        blockId: 'table-row-3',
        y: 60,
        height: 40,
        overflow: false,
      },
    ]);
    expect(result.pages[1].usedHeight).toBe(100);
    expect(result.breaks[0].beforeBlockId).toBe('table-row-2');
  });

  test('keeps repeated table headers with the first body row', () => {
    const table = [
      { id: 'table-header', height: 20 },
      { id: 'table-row-1', height: 30 },
      { id: 'table-row-2', height: 30 },
    ].map((candidate, flowIndex, rows) => ({
      ...candidate,
      flowId: 'table',
      flowIndex,
      flowCount: rows.length,
      minimumFragmentsPerPage: 1,
      repeatHeaderCount: 1,
      repeatHeaderHeight: 20,
    }));
    const input = request([block('intro', 60), ...table]);
    input.page.height = 100;
    input.page.marginTop = 0;
    input.page.marginBottom = 0;

    const result = layoutOfficeDocumentInJavaScript(input);

    expect(result.pages[0].placements.map((item) => item.blockId)).toEqual([
      'intro',
    ]);
    expect(result.pages[1].placements.map((item) => item.blockId)).toEqual([
      'table-header',
      'table-row-1',
      'table-row-2',
    ]);
    expect(result.pages[1].placements[0]?.y).toBe(0);
  });

  test('rejects incomplete repeated table-header metadata', () => {
    expect(() =>
      layoutOfficeDocumentInJavaScript(
        request([
          {
            ...block('table-header', 20),
            flowId: 'table',
            flowIndex: 0,
            flowCount: 2,
            minimumFragmentsPerPage: 1,
            repeatHeaderCount: 1,
          },
          {
            ...block('table-row', 30),
            flowId: 'table',
            flowIndex: 1,
            flowCount: 2,
            minimumFragmentsPerPage: 1,
            repeatHeaderCount: 1,
          },
        ]),
      ),
    ).toThrow(
      'Repeated table headers require repeatHeaderCount and repeatHeaderHeight.',
    );
  });

  test('rejects non-consecutive paragraph flow fragments', () => {
    expect(() =>
      layoutOfficeDocumentInJavaScript(
        request([
          {
            ...block('paragraph-line-0', 20),
            flowId: 'paragraph',
            flowIndex: 0,
            flowCount: 2,
            minimumFragmentsPerPage: 1,
          },
          {
            ...block('paragraph-line-1', 20),
            flowId: 'paragraph',
            flowIndex: 0,
            flowCount: 2,
            minimumFragmentsPerPage: 1,
          },
        ]),
      ),
    ).toThrow(
      'Layout flow fragments must be consecutive and consistently indexed.',
    );
  });

  test('rejects page breaks inside a paragraph flow', () => {
    expect(() =>
      layoutOfficeDocumentInJavaScript(
        request([
          {
            ...block('paragraph-line-0', 20),
            breakAfter: true,
            flowId: 'paragraph',
            flowIndex: 0,
            flowCount: 2,
            minimumFragmentsPerPage: 1,
          },
          {
            ...block('paragraph-line-1', 20),
            flowId: 'paragraph',
            flowIndex: 1,
            flowCount: 2,
            minimumFragmentsPerPage: 1,
          },
        ]),
      ),
    ).toThrow('Only the last flow fragment may request breakAfter.');
  });

  test('rejects impossible page metrics', () => {
    const invalid = request([block('one', 10)]);
    invalid.page.marginTop = invalid.page.height;

    expect(() => layoutOfficeDocumentInJavaScript(invalid)).toThrow(
      'Page height must leave a positive body area.',
    );
  });

  test('rejects duplicate block IDs', () => {
    expect(() =>
      layoutOfficeDocumentInJavaScript(
        request([block('same', 10), block('same', 20)]),
      ),
    ).toThrow("Layout block ID 'same' is duplicated.");
  });

  test('rejects malformed responses at the Worker boundary', () => {
    expect(
      isOfficeKernelResponse({
        protocol: OFFICE_KERNEL_PROTOCOL_VERSION,
        kind: 'layoutResult',
        requestId: 1,
        revision: 1,
        engine: 'wasm',
        pages: 'not-an-array',
        breaks: [],
      }),
    ).toBe(false);
  });

  test('validates shaped text lines with UTF-16 offsets', () => {
    expect(
      isOfficeKernelResponse({
        protocol: OFFICE_KERNEL_PROTOCOL_VERSION,
        kind: 'textLayoutResult',
        requestId: 13,
        revision: 8,
        documentRevision: 5,
        engine: 'wasm',
        layouts: [
          {
            id: 'paragraph',
            glyphCount: 7,
            fallbackGlyphCount: 2,
            missingGlyphCount: 0,
            lines: [
              {
                startUtf16: 0,
                endUtf16: 4,
                width: 48,
                ascent: 12,
                descent: 4,
                height: 21,
                hardBreak: false,
              },
              {
                startUtf16: 4,
                endUtf16: 8,
                width: 36,
                ascent: 12,
                descent: 4,
                height: 21,
                hardBreak: false,
              },
            ],
          },
        ],
        unsupportedParagraphIds: [],
      }),
    ).toBe(true);

    expect(
      isOfficeKernelResponse({
        protocol: OFFICE_KERNEL_PROTOCOL_VERSION,
        kind: 'textLayoutResult',
        requestId: 13,
        revision: 8,
        documentRevision: 5,
        engine: 'wasm',
        layouts: [
          {
            id: 'paragraph',
            glyphCount: 1,
            fallbackGlyphCount: 0,
            missingGlyphCount: 0,
            lines: [
              {
                startUtf16: 3,
                endUtf16: 2,
                width: 10,
                ascent: 8,
                descent: 2,
                height: 12,
                hardBreak: false,
              },
            ],
          },
        ],
        unsupportedParagraphIds: [],
      }),
    ).toBe(false);
  });

  test('marks JavaScript text layout as unsupported instead of guessing font metrics', () => {
    const result = layoutOfficeTextInJavaScript(textLayoutRequest());

    expect(result.engine).toBe('javascript');
    expect(result.layouts).toEqual([]);
    expect(result.unsupportedParagraphIds).toEqual(['paragraph']);
  });

  test('uses the JavaScript kernel when workers are unavailable', async () => {
    const workerDescriptor = Object.getOwnPropertyDescriptor(
      globalThis,
      'Worker',
    );
    Object.defineProperty(globalThis, 'Worker', {
      configurable: true,
      value: undefined,
    });
    const client = createOfficeKernelClient();

    try {
      const result = await client.layout({
        revision: 9,
        documentRevision: 4,
        page: request([]).page,
        blocks: [block('one', 20)],
      });
      expect(result.engine).toBe('javascript');
      expect(result.revision).toBe(9);
      expect(result.documentRevision).toBe(4);
      const geometry = await client.presentationGeometry({
        revision: 10,
        documentRevision: 5,
        operation: { type: 'alignToSlide', alignment: 'bottom' },
        elements: [{ id: 'shape', x: 10, y: 10, width: 30, height: 20 }],
      });
      expect(geometry.engine).toBe('javascript');
      expect(geometry.elements[0].y).toBe(80);
      expect(geometry.documentRevision).toBe(5);
      const text = await client.textLayout({
        revision: 11,
        documentRevision: 6,
        paragraphs: textLayoutRequest().paragraphs,
      });
      expect(text.engine).toBe('javascript');
      expect(text.layouts).toEqual([]);
      expect(text.unsupportedParagraphIds).toEqual(['paragraph']);
    } finally {
      client.dispose();
      if (workerDescriptor) {
        Object.defineProperty(globalThis, 'Worker', workerDescriptor);
      } else {
        Reflect.deleteProperty(globalThis, 'Worker');
      }
    }
  });
});

function textLayoutRequest(): OfficeKernelTextLayoutRequest {
  return {
    protocol: OFFICE_KERNEL_PROTOCOL_VERSION,
    kind: 'textLayout',
    requestId: 13,
    revision: 8,
    documentRevision: 5,
    paragraphs: [
      {
        id: 'paragraph',
        text: 'A3S 文档😀',
        runs: [
          {
            startUtf16: 0,
            endUtf16: 8,
            fontId: 'noto-sans-hans-regular',
            fontSize: 14,
            lineHeight: 21,
            letterSpacing: 0,
            ligatures: false,
            kerning: true,
          },
        ],
        maxWidth: 120,
        firstLineMaxWidth: 96,
        direction: 'auto',
        whiteSpace: 'breakSpaces',
      },
    ],
  };
}

function request(blocks: OfficeKernelLayoutBlock[]): OfficeKernelLayoutRequest {
  return {
    protocol: OFFICE_KERNEL_PROTOCOL_VERSION,
    kind: 'layout',
    requestId: 7,
    revision: 12,
    documentRevision: 4,
    startPageIndex: 0,
    page: {
      width: 794,
      height: 1123,
      marginTop: 80,
      marginRight: 80,
      marginBottom: 80,
      marginLeft: 80,
      headerHeight: 0,
      footerHeight: 0,
      pageGap: 24,
    },
    blocks,
  };
}

function block(id: string, height: number): OfficeKernelLayoutBlock {
  return { id, height };
}
