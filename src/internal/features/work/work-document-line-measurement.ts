import type { Editor } from '@tiptap/core';
import type { Node as ProseMirrorNode } from '@tiptap/pm/model';
import type { OfficeKernelTextLayoutParagraphResult } from '../../kernel/office-kernel-protocol';
import {
  type DocumentNodeParagraphPagination,
  documentInlineOffsets,
  nonNegativePixels,
  verticalBlockEnd,
  verticalBlockStart,
} from './work-document-pagination-dom';
import type { MeasuredDocumentLayoutBlock } from './work-document-pagination-types';
import { documentTextLayoutContent } from './work-document-text-layout';

interface DocumentLineRect {
  top: number;
  right: number;
  bottom: number;
  left: number;
}

interface DocumentTextNodeMeasurement {
  node: Text;
  rects: DocumentLineRect[];
}

interface DocumentMeasuredLine extends DocumentLineRect {
  position: number;
}

interface DocumentLineFragment {
  from: number;
  to: number;
  height: number;
}

export function createDocumentLineFragments(
  lines: readonly Pick<DocumentMeasuredLine, 'top' | 'position'>[],
  outerTop: number,
  outerBottom: number,
  scale = 1,
): DocumentLineFragment[] {
  const normalizedScale = Number.isFinite(scale) && scale > 0 ? scale : 1;
  return lines.flatMap((line, index) => {
    const next = lines[index + 1];
    const start = index === 0 ? outerTop : line.top;
    const end = next?.top ?? outerBottom;
    const from = line.position;
    const to = next?.position ?? from + 1;
    if (end <= start || to <= from) return [];
    return [
      {
        from,
        to,
        height: Math.max(1, (end - start) / normalizedScale),
      },
    ];
  });
}

export function measureParagraphLineFragments(
  editor: Editor,
  node: ProseMirrorNode,
  element: HTMLElement,
  id: string,
  from: number,
  to: number,
  pagination: DocumentNodeParagraphPagination,
  shapedLayout?: OfficeKernelTextLayoutParagraphResult,
): MeasuredDocumentLayoutBlock[] {
  const text = documentTextLayoutContent(node);
  if (
    node.type.name !== 'paragraph' ||
    text === null ||
    text.length < 2 ||
    element.querySelector(
      [
        'img',
        'table',
        '[data-document-note]',
        [
          '[contenteditable="false"]',
          ':not(.work-document-auto-page-break)',
          ':not([data-document-tab])',
        ].join(''),
      ].join(','),
    )
  ) {
    return [];
  }
  if (shapedLayout?.missingGlyphCount === 0) {
    const shaped = measureShapedParagraphLineFragments(
      editor,
      element,
      id,
      from,
      to,
      pagination,
      shapedLayout,
    );
    if (shaped.length > 1) return shaped;
  }
  const range = element.ownerDocument.createRange();
  if (typeof range.getClientRects !== 'function') return [];
  range.selectNodeContents(element);
  const rects = groupDocumentLineRects(
    Array.from(range.getClientRects())
      .filter((rect) => rect.height > 0 && rect.width > 0)
      .map((rect) => ({
        top: rect.top,
        right: rect.right,
        bottom: rect.bottom,
        left: rect.left,
      })),
  );
  range.detach();
  if (rects.length < 2) return [];

  const textNodes = measureDocumentTextNodes(element);
  const lines = rects.flatMap((rect, index) => {
    if (index === 0) return [{ ...rect, position: from }];
    const position =
      documentPositionAtMeasuredLineStart(editor, textNodes, rect) ??
      documentPositionAtPoint(editor, element, rect);
    const previous = rects[index - 1];
    if (
      position === null ||
      position <= from ||
      position >= to ||
      rect.top <= previous.top
    ) {
      return [];
    }
    return [{ ...rect, position }];
  });
  const uniqueLines = lines.filter(
    (line, index) => index === 0 || line.position > lines[index - 1].position,
  );
  if (uniqueLines.length !== rects.length) return [];

  const elementRect = element.getBoundingClientRect();
  const style = getComputedStyle(element);
  const scaleY =
    element.offsetHeight > 0 ? elementRect.height / element.offsetHeight : 1;
  const outerTop =
    elementRect.top - nonNegativePixels(style.marginTop) * scaleY;
  const outerBottom =
    elementRect.bottom + nonNegativePixels(style.marginBottom) * scaleY;
  const fragments = createDocumentLineFragments(
    uniqueLines,
    outerTop,
    outerBottom,
    scaleY,
  );
  if (fragments.length !== uniqueLines.length) return [];

  const editorRect = editor.view.dom.getBoundingClientRect();
  const scaleX =
    editor.view.dom.offsetWidth > 0
      ? editorRect.width / editor.view.dom.offsetWidth
      : 1;
  const inlineOffsetLeft = Math.max(
    0,
    (elementRect.left - editorRect.left) / Math.max(scaleX, 0.01),
  );
  const inlineOffsetRight = Math.max(
    0,
    (editorRect.right - elementRect.right) / Math.max(scaleX, 0.01),
  );
  const minimumFragmentsPerPage = pagination.widowControl
    ? Math.min(2, fragments.length)
    : 1;
  return fragments.map((fragment, flowIndex) => ({
    block: {
      id: `${id}-line-${flowIndex}`,
      height: fragment.height,
      breakBefore: flowIndex === 0 && pagination.pageBreakBefore,
      keepTogether: flowIndex === 0 && pagination.keepLines,
      keepWithNext:
        flowIndex + 1 === fragments.length && pagination.keepWithNext,
      flowId: id,
      flowIndex,
      flowCount: fragments.length,
      minimumFragmentsPerPage,
    },
    element,
    from: fragment.from,
    to: flowIndex + 1 < fragments.length ? fragment.to : to,
    inlineOffsetLeft: flowIndex === 0 ? 0 : inlineOffsetLeft,
    inlineOffsetRight: flowIndex === 0 ? 0 : inlineOffsetRight,
    observeResize: false,
  }));
}

function measureShapedParagraphLineFragments(
  editor: Editor,
  element: HTMLElement,
  id: string,
  from: number,
  to: number,
  pagination: DocumentNodeParagraphPagination,
  layout: OfficeKernelTextLayoutParagraphResult,
): MeasuredDocumentLayoutBlock[] {
  if (layout.lines.length < 2) return [];
  const style = getComputedStyle(element);
  const fragments = createShapedDocumentLineFragments(
    layout.lines,
    from,
    to,
    verticalBlockStart(style),
    verticalBlockEnd(style),
  );
  if (fragments.length !== layout.lines.length) return [];
  const { inlineOffsetLeft, inlineOffsetRight } = documentInlineOffsets(
    editor,
    element,
  );
  const minimumFragmentsPerPage = pagination.widowControl
    ? Math.min(2, fragments.length)
    : 1;
  return fragments.map((fragment, flowIndex) => ({
    block: {
      id: `${id}-line-${flowIndex}`,
      height: fragment.height,
      breakBefore: flowIndex === 0 && pagination.pageBreakBefore,
      keepTogether: flowIndex === 0 && pagination.keepLines,
      keepWithNext:
        flowIndex + 1 === fragments.length && pagination.keepWithNext,
      flowId: id,
      flowIndex,
      flowCount: fragments.length,
      minimumFragmentsPerPage,
    },
    element,
    from: fragment.from,
    to: fragment.to,
    inlineOffsetLeft: flowIndex === 0 ? 0 : inlineOffsetLeft,
    inlineOffsetRight: flowIndex === 0 ? 0 : inlineOffsetRight,
    observeResize: false,
  }));
}

export function createShapedDocumentLineFragments(
  lines: readonly Pick<
    OfficeKernelTextLayoutParagraphResult['lines'][number],
    'endUtf16' | 'height' | 'startUtf16'
  >[],
  nodeFrom: number,
  nodeTo: number,
  blockStart: number,
  blockEnd: number,
): DocumentLineFragment[] {
  const fragments: DocumentLineFragment[] = [];
  for (const [index, line] of lines.entries()) {
    const next = lines[index + 1];
    const from =
      index === 0 ? nodeFrom : nodeFrom + 1 + Math.max(0, line.startUtf16);
    const to = next ? nodeFrom + 1 + Math.max(0, next.startUtf16) : nodeTo;
    if (
      line.endUtf16 < line.startUtf16 ||
      (next && next.startUtf16 < line.endUtf16) ||
      !Number.isSafeInteger(from) ||
      !Number.isSafeInteger(to) ||
      to <= from ||
      from < nodeFrom ||
      to > nodeTo
    ) {
      return [];
    }
    fragments.push({
      from,
      to,
      height:
        line.height +
        (index === 0 ? blockStart : 0) +
        (index + 1 === lines.length ? blockEnd : 0),
    });
  }
  return fragments;
}

function groupDocumentLineRects(
  rects: readonly DocumentLineRect[],
): DocumentLineRect[] {
  const lines: DocumentLineRect[] = [];
  for (const rect of [...rects].sort(
    (left, right) => left.top - right.top || left.left - right.left,
  )) {
    const line = lines.find((candidate) => verticallyOverlaps(candidate, rect));
    if (line) {
      line.top = Math.min(line.top, rect.top);
      line.right = Math.max(line.right, rect.right);
      line.bottom = Math.max(line.bottom, rect.bottom);
      line.left = Math.min(line.left, rect.left);
    } else {
      lines.push({ ...rect });
    }
  }
  return lines.sort((left, right) => left.top - right.top);
}

function verticallyOverlaps(
  left: DocumentLineRect,
  right: DocumentLineRect,
): boolean {
  const overlap =
    Math.min(left.bottom, right.bottom) - Math.max(left.top, right.top);
  const minimumHeight = Math.min(
    left.bottom - left.top,
    right.bottom - right.top,
  );
  return overlap > Math.max(1, minimumHeight * 0.2);
}

function measureDocumentTextNodes(
  element: HTMLElement,
): DocumentTextNodeMeasurement[] {
  const measurements: DocumentTextNodeMeasurement[] = [];
  const showText = element.ownerDocument.defaultView?.NodeFilter.SHOW_TEXT ?? 4;
  const walker = element.ownerDocument.createTreeWalker(element, showText);
  let current = walker.nextNode();
  while (current) {
    const node = current as Text;
    if (
      node.data.length > 0 &&
      !node.parentElement?.closest(
        '.work-document-auto-page-break, [contenteditable="false"]',
      )
    ) {
      const range = element.ownerDocument.createRange();
      range.selectNodeContents(node);
      const rects = Array.from(range.getClientRects())
        .filter((rect) => rect.height > 0 && rect.width > 0)
        .map((rect) => ({
          top: rect.top,
          right: rect.right,
          bottom: rect.bottom,
          left: rect.left,
        }));
      range.detach();
      if (rects.length) measurements.push({ node, rects });
    }
    current = walker.nextNode();
  }
  return measurements;
}

function documentPositionAtMeasuredLineStart(
  editor: Editor,
  textNodes: readonly DocumentTextNodeMeasurement[],
  line: DocumentLineRect,
): number | null {
  for (const measurement of textNodes) {
    if (!measurement.rects.some((rect) => verticallyOverlaps(rect, line))) {
      continue;
    }
    const range = measurement.node.ownerDocument.createRange();
    const offset = findDocumentLineStartOffset(
      measurement.node.data.length,
      line,
      (candidate) => {
        range.setStart(measurement.node, candidate);
        range.setEnd(measurement.node, candidate + 1);
        const rectangle = Array.from(range.getClientRects()).find(
          (rect) => rect.height > 0 && rect.width > 0,
        );
        return rectangle
          ? {
              top: rectangle.top,
              bottom: rectangle.bottom,
            }
          : null;
      },
    );
    range.detach();
    if (offset === null) continue;
    try {
      return editor.view.posAtDOM(measurement.node, offset);
    } catch {
      // Try the next measured text node when this DOM position is stale.
    }
  }
  return null;
}

export function findDocumentLineStartOffset(
  textLength: number,
  line: Pick<DocumentLineRect, 'top' | 'bottom'>,
  rectangleAt: (
    offset: number,
  ) => Pick<DocumentLineRect, 'top' | 'bottom'> | null,
): number | null {
  if (!Number.isSafeInteger(textLength) || textLength <= 0) return null;
  let lower = 0;
  let upper = textLength;
  while (lower < upper) {
    const middle = Math.floor((lower + upper) / 2);
    const rectangle = rectangleAt(middle);
    if (!rectangle || rectangle.bottom <= line.top) {
      lower = middle + 1;
    } else {
      upper = middle;
    }
  }

  const searchStart = Math.max(0, lower - 4);
  const searchEnd = Math.min(textLength, lower + 5);
  for (let offset = searchStart; offset < searchEnd; offset += 1) {
    const rectangle = rectangleAt(offset);
    if (
      rectangle &&
      rectangle.bottom > line.top &&
      rectangle.top < line.bottom
    ) {
      return offset;
    }
  }
  return null;
}

function documentPositionAtPoint(
  editor: Editor,
  element: HTMLElement,
  line: DocumentLineRect,
): number | null {
  const ownerDocument = element.ownerDocument as Document & {
    caretPositionFromPoint?: (
      x: number,
      y: number,
    ) => { offsetNode: Node; offset: number } | null;
    caretRangeFromPoint?: (x: number, y: number) => Range | null;
  };
  const style = getComputedStyle(element);
  const x = style.direction === 'rtl' ? line.right - 0.5 : line.left + 0.5;
  const y = line.top + (line.bottom - line.top) / 2;
  const caretPosition = ownerDocument.caretPositionFromPoint?.(x, y);
  const offsetNode =
    caretPosition?.offsetNode ??
    ownerDocument.caretRangeFromPoint?.(x, y)?.startContainer;
  const offset =
    caretPosition?.offset ??
    ownerDocument.caretRangeFromPoint?.(x, y)?.startOffset;
  if (
    !offsetNode ||
    offset === undefined ||
    (offsetNode !== element && !element.contains(offsetNode))
  ) {
    return null;
  }
  try {
    return editor.view.posAtDOM(offsetNode, offset);
  } catch {
    return null;
  }
}
