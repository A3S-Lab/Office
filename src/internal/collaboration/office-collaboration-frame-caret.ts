import type { Node as ProseMirrorNode } from '@tiptap/pm/model';
import { WorkOfficeCollaborationError } from './office-collaboration';

const DOCUMENT_CARET_NODE_TYPES = new Set([
  'paragraph',
  'heading',
  'documentCaption',
]);

/**
 * Caret carried by the same collaboration frame as a document update.
 * It is not a follow-up location message.
 */
export type WorkOfficeCollaborationFrameCaret =
  | { readonly kind: 'markdown'; readonly indexUtf16: number }
  | {
      readonly kind: 'document';
      readonly paragraphId: string;
      readonly textId: string;
      readonly indexUtf16: number;
    }
  | {
      readonly kind: 'spreadsheet';
      readonly sheetId: string;
      readonly row: number;
      readonly column: number;
      /** Present only for spreadsheet-splice. Cell writes omit it. */
      readonly indexUtf16?: number;
    }
  | {
      readonly kind: 'presentation';
      readonly containerKind: 'slide' | 'master' | 'layout';
      readonly containerId: string;
      readonly elementId: string;
      readonly indexUtf16?: number;
    }
  | { readonly kind: 'pdf-field'; readonly fieldId: string }
  | { readonly kind: 'pdf-annotation'; readonly annotationId: string };

export interface FrameCaretLineBox {
  readonly from: number;
  readonly to: number;
  readonly left: number;
  readonly top: number;
  readonly right: number;
  readonly bottom: number;
}

export interface FrameCaretPaint {
  readonly head: number;
  readonly left: number;
  readonly top: number;
  readonly width: number;
  readonly height: number;
}

export function parseWorkOfficeCollaborationFrameCaret(
  value: unknown,
): WorkOfficeCollaborationFrameCaret {
  if (!isRecord(value) || typeof value.kind !== 'string') {
    throw invalidCaret(
      'An Office collaboration frame caret must name its kind.',
    );
  }
  if ('pagePixels' in value) {
    throw invalidCaret(
      'An Office collaboration frame caret does not carry page pixels.',
    );
  }
  switch (value.kind) {
    case 'markdown':
      return { kind: 'markdown', indexUtf16: utf16Offset(value.indexUtf16) };
    case 'document':
      return {
        kind: 'document',
        paragraphId: requiredText(value.paragraphId, 'paragraphId'),
        textId: requiredText(value.textId, 'textId'),
        indexUtf16: utf16Offset(value.indexUtf16),
      };
    case 'spreadsheet': {
      const caret = {
        kind: 'spreadsheet' as const,
        sheetId: requiredText(value.sheetId, 'sheetId'),
        row: wholeNumber(value.row, 'row'),
        column: wholeNumber(value.column, 'column'),
      };
      if (!('indexUtf16' in value)) return caret;
      return { ...caret, indexUtf16: utf16Offset(value.indexUtf16) };
    }
    case 'presentation': {
      const containerKind = presentationContainerKind(value.containerKind);
      const caret = {
        kind: 'presentation' as const,
        containerKind,
        containerId: requiredText(value.containerId, 'containerId'),
        elementId: requiredText(value.elementId, 'elementId'),
      };
      if (!('indexUtf16' in value)) return caret;
      return { ...caret, indexUtf16: utf16Offset(value.indexUtf16) };
    }
    case 'pdf-field':
      if ('indexUtf16' in value) {
        throw invalidCaret(
          'A PDF form frame names the field and has no text offset.',
        );
      }
      return {
        kind: 'pdf-field',
        fieldId: requiredText(value.fieldId, 'fieldId'),
      };
    case 'pdf-annotation':
      if ('indexUtf16' in value) {
        throw invalidCaret(
          'A PDF annotation frame names the annotation and has no text offset.',
        );
      }
      return {
        kind: 'pdf-annotation',
        annotationId: requiredText(value.annotationId, 'annotationId'),
      };
    default:
      throw invalidCaret(
        `The Office collaboration frame caret kind '${value.kind}' is invalid.`,
      );
  }
}

export function frameCaretFromOrigin(
  origin: unknown,
): WorkOfficeCollaborationFrameCaret | null {
  if (!isRecord(origin) || !('caret' in origin)) return null;
  try {
    return parseWorkOfficeCollaborationFrameCaret(origin.caret);
  } catch {
    return null;
  }
}

/**
 * Source-editor caret. JavaScript string indexes are UTF-16 code units, so
 * this does not insert a newline for each rendered block.
 */
export function markdownSourceCaret(
  source: string,
  indexUtf16: number,
): number | null {
  if (!isUtf16Boundary(source, indexUtf16)) return null;
  return indexUtf16;
}

/**
 * Editor caret inside the addressed paragraph. Offsets count that node's text
 * only. Blocks before it do not add a newline.
 */
export function documentFrameCaretPosition(
  document: ProseMirrorNode,
  caret: {
    readonly paragraphId: string;
    readonly textId: string;
    readonly indexUtf16: number;
  },
): number | null {
  const matches: { node: ProseMirrorNode; position: number }[] = [];
  document.descendants((node, position) => {
    if (!DOCUMENT_CARET_NODE_TYPES.has(node.type.name)) return;
    if (node.attrs.paragraphId !== caret.paragraphId) return;
    matches.push({ node, position });
  });
  if (matches.length !== 1) return null;
  const match = matches[0];
  if (!match || match.node.attrs.textId !== caret.textId) return null;
  return utf16OffsetInsideNode(match.node, match.position, caret.indexUtf16);
}

/**
 * Paint the caret the frame already carried, after that update's layout.
 * A follow-up location is not read. A stale or early measure paints nothing
 * and does not fall back to the document end.
 */
export function paintFrameCaretAfterLayout(input: {
  readonly layoutSettled: boolean;
  readonly head: number | null;
  readonly documentSize: number;
  readonly lines: readonly FrameCaretLineBox[];
  readonly followUpHead?: number | null;
}): FrameCaretPaint | null {
  void input.followUpHead;
  if (!input.layoutSettled || input.head === null) return null;
  if (input.head < 0 || input.head > input.documentSize) return null;
  const line = lineBoxForCaret(input.lines, input.head);
  if (!line) return null;
  const height = Math.max(1, line.bottom - line.top);
  const paint = {
    head: input.head,
    left: line.left,
    top: line.top,
    width: 2,
    height,
  };
  if (
    paint.top < line.top ||
    paint.top + paint.height > line.bottom + 0.01 ||
    paint.left < line.left ||
    paint.left > line.right
  ) {
    return null;
  }
  return paint;
}

export function lineBoxForCaret(
  lines: readonly FrameCaretLineBox[],
  head: number,
): FrameCaretLineBox | null {
  const starting = lines.find((line) => line.from === head);
  if (starting) return starting;
  return (
    lines.find((line) => head > line.from && head <= line.to) ?? null
  );
}

function utf16OffsetInsideNode(
  node: ProseMirrorNode,
  nodePosition: number,
  indexUtf16: number,
): number | null {
  let seen = 0;
  let resolved: number | null = null;
  node.descendants((child, childPosition) => {
    if (!child.isText || child.text == null || resolved !== null) return;
    const text = child.text;
    const next = seen + text.length;
    if (indexUtf16 < seen || indexUtf16 > next) {
      seen = next;
      return;
    }
    const local = indexUtf16 - seen;
    if (!isUtf16Boundary(text, local)) return;
    resolved = nodePosition + 1 + childPosition + local;
    seen = next;
  });
  if (resolved !== null) return resolved;
  if (indexUtf16 === seen && isUtf16Boundary(node.textContent, indexUtf16)) {
    return nodePosition + 1 + node.content.size;
  }
  return null;
}

function isUtf16Boundary(value: string, offset: number): boolean {
  if (!Number.isSafeInteger(offset) || offset < 0 || offset > value.length) {
    return false;
  }
  if (offset === 0 || offset === value.length) return true;
  const code = value.charCodeAt(offset);
  return code < 0xdc00 || code > 0xdfff;
}

function utf16Offset(value: unknown): number {
  if (!Number.isSafeInteger(value) || (value as number) < 0) {
    throw invalidCaret(
      'An Office collaboration frame caret offset must be a non-negative integer.',
    );
  }
  return value as number;
}

function wholeNumber(value: unknown, field: string): number {
  if (!Number.isSafeInteger(value) || (value as number) < 0) {
    throw invalidCaret(
      `An Office collaboration frame caret ${field} must be a non-negative integer.`,
    );
  }
  return value as number;
}

function requiredText(value: unknown, field: string): string {
  if (typeof value !== 'string' || value.length === 0) {
    throw invalidCaret(
      `An Office collaboration frame caret requires ${field}.`,
    );
  }
  return value;
}

function invalidCaret(message: string): WorkOfficeCollaborationError {
  return new WorkOfficeCollaborationError(
    'office.collaboration.frame_caret_invalid',
    message,
  );
}

function presentationContainerKind(
  value: unknown,
): 'slide' | 'master' | 'layout' {
  if (value === 'slide' || value === 'master' || value === 'layout') {
    return value;
  }
  throw invalidCaret(
    'A presentation frame caret requires a slide, master, or layout container.',
  );
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}
