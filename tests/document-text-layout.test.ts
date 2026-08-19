import { describe, expect, test } from '@rstest/core';
import { Editor } from '@tiptap/core';
import { createWorkDocumentExtensions } from '../src/internal/features/work/work-document-extensions';
import type { WorkDocumentLayoutFont } from '../src/internal/features/work/work-document-fonts';
import {
  collectDocumentTextLayoutParagraphs,
  collectDocumentTextLayoutRuns,
  documentTextLayoutBatches,
} from '../src/internal/features/work/work-document-pagination';
import type { DocumentPaginationSnapshot } from '../src/internal/features/work/work-document-pagination';
import type { OfficeKernelTextLayoutParagraph } from '../src/internal/kernel/office-kernel-protocol';

const layoutFont: WorkDocumentLayoutFont = {
  id: 'layout-regular',
  family: 'Test Layout Sans',
  url: '/fonts/layout-regular.otf',
  weight: 400,
  style: 'normal',
};
const fallbackLayoutFont: WorkDocumentLayoutFont = {
  id: 'layout-cjk-regular',
  family: 'Test Layout CJK',
  url: '/fonts/layout-cjk-regular.otf',
  weight: 400,
  style: 'normal',
};
const boldLayoutFont: WorkDocumentLayoutFont = {
  id: 'layout-bold',
  family: layoutFont.family,
  url: '/fonts/layout-bold.otf',
  weight: 700,
  style: 'normal',
};

test('batches a large document within the kernel paragraph limit', () => {
  const paragraphs: OfficeKernelTextLayoutParagraph[] = Array.from(
    { length: 2_050 },
    (_, index) => ({
      id: `paragraph-${index}`,
      text: 'x',
      runs: [
        {
          startUtf16: 0,
          endUtf16: 1,
          fontId: layoutFont.id,
          fontSize: 14,
          lineHeight: 21,
        },
      ],
      maxWidth: 400,
    }),
  );

  const batches = documentTextLayoutBatches(paragraphs);

  expect(batches.map((batch) => batch.length)).toEqual([1_024, 1_024, 2]);
  expect(batches.flat().map(({ id }) => id)).toEqual(
    paragraphs.map(({ id }) => id),
  );
});

describe('document mixed-run text layout', () => {
  test('collects contiguous UTF-16 runs with independent line metrics', () => {
    const paragraph = document.createElement('p');
    applyTextMetrics(paragraph, 14, 21);
    paragraph.innerHTML = [
      '常规',
      '<span style="font-family: Test Layout Sans; font-size: 22px; line-height: 30px; unicode-bidi: normal">大字</span>',
      '<span style="color: red; font-family: Test Layout Sans; font-size: 14px; line-height: 21px; unicode-bidi: normal">颜色</span>',
    ].join('');
    document.body.append(paragraph);

    try {
      expect(
        collectDocumentTextLayoutRuns(
          paragraph,
          paragraph.textContent ?? '',
          [layoutFont],
          new Set([layoutFont.id]),
        ),
      ).toEqual([
        expect.objectContaining({
          startUtf16: 0,
          endUtf16: 2,
          fontId: layoutFont.id,
          fontSize: 14,
          lineHeight: 21,
        }),
        expect.objectContaining({
          startUtf16: 2,
          endUtf16: 4,
          fontId: layoutFont.id,
          fontSize: 22,
          lineHeight: 30,
        }),
        expect.objectContaining({
          startUtf16: 4,
          endUtf16: 6,
          fontId: layoutFont.id,
          fontSize: 14,
          lineHeight: 21,
        }),
      ]);
    } finally {
      paragraph.remove();
    }
  });

  test('merges adjacent runs when paint-only styles differ', () => {
    const paragraph = document.createElement('p');
    applyTextMetrics(paragraph, 14, 21);
    paragraph.innerHTML = [
      '<span style="color: red; font-family: Test Layout Sans; font-size: 14px; line-height: 21px; unicode-bidi: normal">A</span>',
      '<span style="color: blue; font-family: Test Layout Sans; font-size: 14px; line-height: 21px; unicode-bidi: normal">3S</span>',
    ].join('');
    document.body.append(paragraph);

    try {
      expect(
        collectDocumentTextLayoutRuns(
          paragraph,
          paragraph.textContent ?? '',
          [layoutFont],
          new Set([layoutFont.id]),
        ),
      ).toEqual([
        expect.objectContaining({
          startUtf16: 0,
          endUtf16: 3,
          fontId: layoutFont.id,
        }),
      ]);
    } finally {
      paragraph.remove();
    }
  });

  test('falls back when a visible run does not have an exact registered face', () => {
    const paragraph = document.createElement('p');
    applyTextMetrics(paragraph, 14, 21);
    paragraph.innerHTML = [
      'A3S ',
      '<span style="font-family: Missing Font; font-size: 14px; line-height: 21px; unicode-bidi: normal">Office</span>',
    ].join('');
    document.body.append(paragraph);

    try {
      expect(
        collectDocumentTextLayoutRuns(
          paragraph,
          paragraph.textContent ?? '',
          [layoutFont],
          new Set([layoutFont.id]),
        ),
      ).toBeNull();
    } finally {
      paragraph.remove();
    }
  });

  test('preserves an exact registered CSS font fallback chain', () => {
    const paragraph = document.createElement('p');
    applyTextMetrics(paragraph, 14, 21);
    paragraph.style.fontFamily = [
      JSON.stringify(layoutFont.family),
      JSON.stringify(fallbackLayoutFont.family),
      'sans-serif',
    ].join(', ');
    paragraph.textContent = 'A3S 文档';
    document.body.append(paragraph);

    try {
      expect(
        collectDocumentTextLayoutRuns(
          paragraph,
          paragraph.textContent,
          [layoutFont, fallbackLayoutFont],
          new Set([layoutFont.id, fallbackLayoutFont.id]),
        ),
      ).toEqual([
        expect.objectContaining({
          startUtf16: 0,
          endUtf16: 6,
          fontId: layoutFont.id,
          fallbackFontIds: [fallbackLayoutFont.id],
        }),
      ]);
    } finally {
      paragraph.remove();
    }
  });

  test('falls back to DOM when an intermediate CSS font face is not registered', () => {
    const paragraph = document.createElement('p');
    applyTextMetrics(paragraph, 14, 21);
    paragraph.style.fontFamily = [
      JSON.stringify(layoutFont.family),
      JSON.stringify('Unregistered CJK'),
      JSON.stringify(fallbackLayoutFont.family),
      'sans-serif',
    ].join(', ');
    paragraph.textContent = 'A3S 文档';
    document.body.append(paragraph);

    try {
      expect(
        collectDocumentTextLayoutRuns(
          paragraph,
          paragraph.textContent,
          [layoutFont, fallbackLayoutFont],
          new Set([layoutFont.id, fallbackLayoutFont.id]),
        ),
      ).toBeNull();
    } finally {
      paragraph.remove();
    }
  });

  test('matches WPS-style intermediate weights to the nearest registered CSS face', () => {
    const paragraph = document.createElement('p');
    applyTextMetrics(paragraph, 14, 21);
    paragraph.style.fontWeight = '680';
    paragraph.textContent = 'Weighted heading';
    document.body.append(paragraph);

    try {
      expect(
        collectDocumentTextLayoutRuns(
          paragraph,
          paragraph.textContent,
          [layoutFont, boldLayoutFont],
          new Set([layoutFont.id, boldLayoutFont.id]),
        ),
      ).toEqual([
        expect.objectContaining({
          fontId: boldLayoutFont.id,
          startUtf16: 0,
          endUtf16: paragraph.textContent.length,
        }),
      ]);
    } finally {
      paragraph.remove();
    }
  });

  test('keeps synthetic bold metrics deterministic when only a regular face exists', () => {
    const paragraph = document.createElement('p');
    applyTextMetrics(paragraph, 14, 21);
    paragraph.style.fontWeight = '700';
    paragraph.textContent = 'Synthetic bold';
    document.body.append(paragraph);

    try {
      expect(
        collectDocumentTextLayoutRuns(
          paragraph,
          paragraph.textContent,
          [layoutFont],
          new Set([layoutFont.id]),
        ),
      ).toEqual([
        expect.objectContaining({
          fontId: layoutFont.id,
          startUtf16: 0,
          endUtf16: paragraph.textContent.length,
        }),
      ]);
    } finally {
      paragraph.remove();
    }
  });

  test('does not substitute a normal face for an unregistered italic style', () => {
    const paragraph = document.createElement('p');
    applyTextMetrics(paragraph, 14, 21);
    paragraph.style.fontStyle = 'italic';
    paragraph.textContent = 'Italic';
    document.body.append(paragraph);

    try {
      expect(
        collectDocumentTextLayoutRuns(
          paragraph,
          paragraph.textContent,
          [layoutFont],
          new Set([layoutFont.id]),
        ),
      ).toBeNull();
    } finally {
      paragraph.remove();
    }
  });

  test('keeps structured document tabs in the shaped UTF-16 stream', () => {
    const paragraph = document.createElement('p');
    applyTextMetrics(paragraph, 14, 21);
    paragraph.innerHTML = [
      '项目',
      '<span class="work-document-tab" data-document-tab="true" contenteditable="false"></span>',
      '负责人',
    ].join('');
    document.body.append(paragraph);

    try {
      expect(
        collectDocumentTextLayoutRuns(
          paragraph,
          '项目\t负责人',
          [layoutFont],
          new Set([layoutFont.id]),
        ),
      ).toEqual([
        expect.objectContaining({
          startUtf16: 0,
          endUtf16: 6,
          fontId: layoutFont.id,
          fontSize: 14,
          lineHeight: 21,
        }),
      ]);
    } finally {
      paragraph.remove();
    }
  });

  test('sends paragraph tab geometry to the deterministic layout kernel', () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: {
        type: 'doc',
        content: [
          {
            type: 'documentSection',
            content: [
              {
                type: 'paragraph',
                attrs: {
                  indentLevel: 1,
                  firstLineIndent: 12,
                  tabStops: [
                    { position: 96, alignment: 'right', leader: 'dot' },
                  ],
                },
                content: [
                  { type: 'text', text: '项目' },
                  { type: 'documentTab' },
                  { type: 'text', text: '负责人' },
                ],
              },
            ],
          },
        ],
      },
    });
    document.body.append(editor.view.dom);
    const paragraph = editor.view.dom.querySelector('p');
    expect(paragraph).toBeInstanceOf(HTMLElement);
    applyTextMetrics(paragraph as HTMLElement, 14, 21);
    (paragraph as HTMLElement).style.marginLeft = '24px';
    (paragraph as HTMLElement).style.textIndent = '12px';
    Object.defineProperty(paragraph, 'clientWidth', {
      configurable: true,
      value: 400,
    });

    try {
      expect(editor.getJSON()).toEqual(
        expect.objectContaining({
          content: [
            expect.objectContaining({
              content: [
                expect.objectContaining({
                  content: [
                    expect.objectContaining({ text: '项目' }),
                    expect.objectContaining({ type: 'documentTab' }),
                    expect.objectContaining({ text: '负责人' }),
                  ],
                }),
              ],
            }),
          ],
        }),
      );
      expect(
        collectDocumentTextLayoutRuns(
          paragraph as HTMLElement,
          '项目\t负责人',
          [layoutFont],
          new Set([layoutFont.id]),
        ),
      ).not.toBeNull();
      expect(
        collectDocumentTextLayoutParagraphs(
          editor,
          [layoutFont],
          new Set([layoutFont.id]),
        ).paragraphs,
      ).toEqual([
        expect.objectContaining({
          text: '项目\t负责人',
          maxWidth: 400,
          firstLineMaxWidth: 388,
          tabLayout: {
            origin: 24,
            firstLineIndent: 12,
            defaultInterval: 48,
            stops: [{ position: 96, alignment: 'right' }],
          },
        }),
      ]);
    } finally {
      editor.view.dom.remove();
      editor.destroy();
    }
  });

  test('sends headings to the deterministic line-layout kernel', () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content:
        '<section data-document-section="true"><h1>Long heading content that can span physical pages</h1></section>',
    });
    document.body.append(editor.view.dom);
    const heading = editor.view.dom.querySelector('h1');
    expect(heading).toBeInstanceOf(HTMLElement);
    applyTextMetrics(heading as HTMLElement, 24, 32);
    Object.defineProperty(heading, 'clientWidth', {
      configurable: true,
      value: 400,
    });

    try {
      expect(
        collectDocumentTextLayoutParagraphs(
          editor,
          [layoutFont],
          new Set([layoutFont.id]),
        ).paragraphs,
      ).toEqual([
        expect.objectContaining({
          text: 'Long heading content that can span physical pages',
          maxWidth: 400,
        }),
      ]);
    } finally {
      editor.view.dom.remove();
      editor.destroy();
    }
  });

  test('skips window chunks before consulting the previous block snapshot', () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: {
        type: 'doc',
        content: [
          {
            type: 'documentSection',
            content: [
              {
                type: 'documentChunk',
                attrs: {
                  id: 'chunk-1',
                  blockCount: 1,
                  estimatedHeight: 24,
                },
                content: [
                  {
                    type: 'paragraph',
                    content: [{ type: 'text', text: 'Windowed paragraph' }],
                  },
                ],
              },
            ],
          },
        ],
      },
    });
    document.body.append(editor.view.dom);
    const blocks: DocumentPaginationSnapshot['blocks'] = [];
    let blockFilterCalls = 0;
    Object.defineProperty(blocks, 'filter', {
      configurable: true,
      value: () => {
        blockFilterCalls += 1;
        return [];
      },
    });

    try {
      expect(
        collectDocumentTextLayoutParagraphs(
          editor,
          [layoutFont],
          new Set([layoutFont.id]),
          {
            blocks,
            measuredBlockCount: 0,
            pageStyles: [],
            reusedBlockCount: 0,
            reusedPrefixBlockCount: 0,
            unsupportedLayout: false,
          },
          Number.MAX_SAFE_INTEGER,
        ).paragraphs,
      ).toEqual([]);
      expect(blockFilterCalls).toBe(0);
    } finally {
      editor.view.dom.remove();
      editor.destroy();
    }
  });

  test('collects list paragraphs for deterministic text layout', () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content:
        '<section data-document-section="true"><ol>' +
        '<li><p>First shaped item</p></li>' +
        '<li><p>Second shaped item</p></li>' +
        '</ol></section>',
    });
    document.body.append(editor.view.dom);
    const paragraphs = Array.from(
      editor.view.dom.querySelectorAll<HTMLElement>('li > p'),
    );
    for (const paragraph of paragraphs) {
      applyTextMetrics(paragraph, 14, 21);
      Object.defineProperty(paragraph, 'clientWidth', {
        configurable: true,
        value: 400,
      });
    }

    try {
      expect(
        collectDocumentTextLayoutParagraphs(
          editor,
          [layoutFont],
          new Set([layoutFont.id]),
        ).paragraphs,
      ).toEqual([
        expect.objectContaining({
          id: expect.stringContaining('-item-0-paragraph-0'),
          text: 'First shaped item',
        }),
        expect.objectContaining({
          id: expect.stringContaining('-item-1-paragraph-0'),
          text: 'Second shaped item',
        }),
      ]);
    } finally {
      editor.view.dom.remove();
      editor.destroy();
    }
  });

  test('uses browser line measurement after a floating document image', () => {
    for (const layout of ['square', 'tight', 'through']) {
      const editor = new Editor({
        extensions: createWorkDocumentExtensions(),
        content: [
          '<section data-document-section="true">',
          '<p>Text before the image</p>',
          '<img src="data:image/png;base64,AA=="',
          ` data-office-image-layout="${layout}"`,
          ' data-office-image-alignment="left">',
          '<p>Text wrapping beside the image</p>',
          '</section>',
        ].join(''),
      });
      document.body.append(editor.view.dom);
      const paragraphs = Array.from(
        editor.view.dom.querySelectorAll<HTMLElement>('p'),
      );
      for (const paragraph of paragraphs) {
        applyTextMetrics(paragraph, 14, 21);
        Object.defineProperty(paragraph, 'clientWidth', {
          configurable: true,
          value: 400,
        });
      }

      try {
        expect(
          collectDocumentTextLayoutParagraphs(
            editor,
            [layoutFont],
            new Set([layoutFont.id]),
          ).paragraphs.map(({ text }) => text),
        ).toEqual(['Text before the image']);
      } finally {
        editor.view.dom.remove();
        editor.destroy();
      }
    }
  });

  test('sends the explicit paragraph direction to the text layout kernel', () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: {
        type: 'doc',
        content: [
          {
            type: 'documentSection',
            content: [
              {
                type: 'paragraph',
                attrs: { paragraphDirection: 'rtl' },
                content: [{ type: 'text', text: 'مرحبا A3S שלום' }],
              },
            ],
          },
        ],
      },
    });
    document.body.append(editor.view.dom);
    const paragraph = editor.view.dom.querySelector('p');
    expect(paragraph).toBeInstanceOf(HTMLElement);
    applyTextMetrics(paragraph as HTMLElement, 14, 21);
    (paragraph as HTMLElement).style.direction = 'rtl';
    Object.defineProperty(paragraph, 'clientWidth', {
      configurable: true,
      value: 400,
    });

    try {
      expect(paragraph).toHaveAttribute('dir', 'rtl');
      expect(
        collectDocumentTextLayoutParagraphs(
          editor,
          [layoutFont],
          new Set([layoutFont.id]),
        ).paragraphs,
      ).toEqual([
        expect.objectContaining({
          text: 'مرحبا A3S שלום',
          direction: 'rtl',
          maxWidth: 400,
        }),
      ]);
    } finally {
      editor.view.dom.remove();
      editor.destroy();
    }
  });
});

function applyTextMetrics(
  element: HTMLElement,
  fontSize: number,
  lineHeight: number,
): void {
  element.style.fontFamily = layoutFont.family;
  element.style.fontSize = `${fontSize}px`;
  element.style.lineHeight = `${lineHeight}px`;
  element.style.direction = 'ltr';
  element.style.unicodeBidi = 'isolate';
  element.style.whiteSpace = 'normal';
}
