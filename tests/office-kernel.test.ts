import { describe, expect, test } from '@rstest/core';
import { createOfficeKernelClient } from '../src/internal/kernel/office-kernel-client';
import { layoutOfficeDocumentInJavaScript } from '../src/internal/kernel/office-kernel-fallback';
import {
  isOfficeKernelResponse,
  type OfficeKernelLayoutBlock,
  type OfficeKernelLayoutRequest,
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
        page: request([]).page,
        blocks: [block('one', 20)],
      });
      expect(result.engine).toBe('javascript');
      expect(result.revision).toBe(9);
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

function request(blocks: OfficeKernelLayoutBlock[]): OfficeKernelLayoutRequest {
  return {
    protocol: OFFICE_KERNEL_PROTOCOL_VERSION,
    kind: 'layout',
    requestId: 7,
    revision: 12,
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
