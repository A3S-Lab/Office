import {
  type Cell,
  convertSpanToShareString,
  type Selection,
} from '@fortune-sheet/core';
import { normalizeCssColor } from '../work-css-color';
import type { WorkSpreadsheetContent } from '../work-types';
import { xlsxRichTextCellText } from '../work-xlsx-rich-text';
import {
  applySpreadsheetRichTextSelectionFormat,
  canApplySpreadsheetRichTextSelectionFormat,
  type SpreadsheetRichTextFormatAttribute,
  type SpreadsheetRichTextSelectionFormatRequest,
  type SpreadsheetRichTextToggleAttribute,
  spreadsheetRichTextSelectionToggleValue,
} from './spreadsheet-rich-text-selection-format';

interface SpreadsheetRichTextSelectionSnapshot {
  column: number;
  consumed: boolean;
  editor: HTMLElement;
  end: number;
  root: HTMLElement;
  row: number;
  sheetId: string;
  sourceCell: Cell;
  start: number;
}

export interface SpreadsheetRichTextSelectionCapture {
  content: WorkSpreadsheetContent;
  root: HTMLElement | null;
  selection: Selection | null | undefined;
  sheetId: string;
}

const richTextFormatTargetSelector = [
  '[data-spreadsheet-rich-text-format="true"]',
  '.work-spreadsheet-font-family',
  '.work-spreadsheet-font-size',
  '.work-spreadsheet-font-color',
  '.work-spreadsheet-underline-split-root',
].join(', ');

export class SpreadsheetRichTextSelectionController {
  private restoreTimer: ReturnType<typeof setTimeout> | null = null;
  private snapshot: SpreadsheetRichTextSelectionSnapshot | null = null;

  capture(options: SpreadsheetRichTextSelectionCapture): boolean {
    this.cancelScheduledRestore();
    const snapshot = captureSpreadsheetRichTextSelection(options);
    this.snapshot = snapshot;
    return snapshot !== null;
  }

  clear(): void {
    this.cancelScheduledRestore();
    this.snapshot = null;
  }

  canApply(
    content: WorkSpreadsheetContent,
    attribute: keyof Cell,
    value: unknown,
  ): boolean {
    const request = this.request(attribute, value);
    return Boolean(
      request && canApplySpreadsheetRichTextSelectionFormat(content, request),
    );
  }

  apply(
    content: WorkSpreadsheetContent,
    onChange: (content: WorkSpreadsheetContent) => void,
    attribute: keyof Cell,
    value: unknown,
  ): boolean {
    const request = this.request(attribute, value);
    return request ? this.applyRequest(content, onChange, request) : false;
  }

  canToggle(
    content: WorkSpreadsheetContent,
    attribute: SpreadsheetRichTextToggleAttribute,
  ): boolean {
    const request = this.toggleRequest(attribute);
    return Boolean(
      request && canApplySpreadsheetRichTextSelectionFormat(content, request),
    );
  }

  toggle(
    content: WorkSpreadsheetContent,
    onChange: (content: WorkSpreadsheetContent) => void,
    attribute: SpreadsheetRichTextToggleAttribute,
  ): boolean {
    const request = this.toggleRequest(attribute);
    return request ? this.applyRequest(content, onChange, request) : false;
  }

  private applyRequest(
    content: WorkSpreadsheetContent,
    onChange: (content: WorkSpreadsheetContent) => void,
    request: SpreadsheetRichTextSelectionFormatRequest,
  ): boolean {
    const next = applySpreadsheetRichTextSelectionFormat(content, request);
    if (!next || !this.snapshot) return false;
    const sourceCell = sheetCellAt(
      next,
      this.snapshot.sheetId,
      this.snapshot.row,
      this.snapshot.column,
    );
    if (!sourceCell) return false;
    this.snapshot = { ...this.snapshot, consumed: true, sourceCell };
    onChange(next);
    return true;
  }

  restore(): boolean {
    const snapshot = this.snapshot;
    if (!snapshot?.consumed) return false;
    const restoredEditor = restoreSpreadsheetTextSelection(snapshot);
    if (!restoredEditor && !snapshot.root.isConnected) {
      this.clear();
      return false;
    }
    this.cancelScheduledRestore();
    this.restoreTimer = setTimeout(() => {
      this.restoreTimer = null;
      if (this.snapshot !== snapshot) return;
      const editor = restoreSpreadsheetTextSelection(snapshot);
      if (!editor) {
        this.snapshot = null;
        return;
      }
      this.snapshot = { ...snapshot, consumed: false, editor };
    }, 0);
    return Boolean(restoredEditor || snapshot.root.isConnected);
  }

  private cancelScheduledRestore(): void {
    if (this.restoreTimer === null) return;
    clearTimeout(this.restoreTimer);
    this.restoreTimer = null;
  }

  private request(
    attribute: keyof Cell,
    value: unknown,
  ): SpreadsheetRichTextSelectionFormatRequest | null {
    const snapshot = this.snapshot;
    if (
      !snapshot ||
      snapshot.consumed ||
      !isSpreadsheetRichTextFormatAttribute(attribute)
    ) {
      return null;
    }
    return {
      attribute,
      column: snapshot.column,
      row: snapshot.row,
      selection: { start: snapshot.start, end: snapshot.end },
      sheetId: snapshot.sheetId,
      sourceCell: snapshot.sourceCell,
      value,
    };
  }

  private toggleRequest(
    attribute: SpreadsheetRichTextToggleAttribute,
  ): SpreadsheetRichTextSelectionFormatRequest | null {
    const snapshot = this.snapshot;
    if (!snapshot || snapshot.consumed) return null;
    const value = spreadsheetRichTextSelectionToggleValue(
      snapshot.sourceCell,
      { start: snapshot.start, end: snapshot.end },
      attribute,
    );
    return value === null ? null : this.request(attribute, value);
  }
}

export function isSpreadsheetRichTextFormatPointerTarget(
  target: EventTarget | null,
): boolean {
  return (
    target instanceof Element &&
    Boolean(target.closest(richTextFormatTargetSelector))
  );
}

export function captureSpreadsheetRichTextSelection({
  content,
  root,
  selection,
  sheetId,
}: SpreadsheetRichTextSelectionCapture): SpreadsheetRichTextSelectionSnapshot | null {
  const browserSelection = window.getSelection();
  if (
    !root ||
    !sheetId ||
    !selection ||
    !browserSelection ||
    browserSelection.rangeCount !== 1 ||
    browserSelection.isCollapsed
  ) {
    return null;
  }
  const range = browserSelection.getRangeAt(0);
  const editor = spreadsheetTextEditor(range.commonAncestorContainer);
  if (
    !editor ||
    !root.contains(editor) ||
    !editorContainsRange(editor, range)
  ) {
    return null;
  }
  const start = textOffsetAtPoint(
    editor,
    range.startContainer,
    range.startOffset,
  );
  const end = textOffsetAtPoint(editor, range.endContainer, range.endOffset);
  if (start === null || end === null || start >= end) return null;
  const row = focusedAxisIndex(selection.row_focus, selection.row);
  const column = focusedAxisIndex(selection.column_focus, selection.column);
  const source = sheetCellAt(content, sheetId, row, column);
  if (!source) return null;
  const sourceCell = liveSpreadsheetTextCell(editor, source);
  const text = xlsxRichTextCellText(sourceCell) ?? sourceCell.v;
  if (
    typeof text !== 'string' ||
    end > text.length ||
    editorText(editor).length !== text.length
  ) {
    return null;
  }
  return {
    column,
    consumed: false,
    editor,
    end,
    root,
    row,
    sheetId,
    sourceCell,
    start,
  };
}

function liveSpreadsheetTextCell(editor: HTMLElement, source: Cell): Cell {
  const text = editorText(editor);
  const sourceText = xlsxRichTextCellText(source) ?? source.v;
  if (text === sourceText) return source;
  const next = { ...source, v: text };
  delete next.f;
  delete next.m;
  const spans = editor.querySelectorAll<HTMLSpanElement>('span');
  const runs = spans.length
    ? convertSpanToShareString(spans, source).flatMap((candidate) => {
        const run = isRecord(candidate)
          ? ({ ...candidate } as Record<string, unknown>)
          : null;
        if (!run || typeof run.v !== 'string') return [];
        const color =
          typeof run.fc === 'string' ? normalizeCssColor(run.fc) : null;
        if (color && color !== 'transparent') run.fc = color;
        else delete run.fc;
        return [run];
      })
    : [];
  if (runs.length && runs.map((run) => run.v).join('') === text) {
    next.ct = { ...source.ct, s: runs, t: 'inlineStr' };
  } else {
    delete next.ct;
  }
  return next;
}

function spreadsheetTextEditor(node: Node): HTMLElement | null {
  const element = node instanceof Element ? node : node.parentElement;
  const editor = element?.closest<HTMLElement>(
    '.luckysheet-cell-input, .fortune-fx-input',
  );
  return editor?.isContentEditable ||
    editor?.getAttribute('contenteditable') === 'true'
    ? editor
    : null;
}

function editorContainsRange(editor: HTMLElement, range: Range): boolean {
  return (
    (range.startContainer === editor ||
      editor.contains(range.startContainer)) &&
    (range.endContainer === editor || editor.contains(range.endContainer))
  );
}

function textOffsetAtPoint(
  editor: HTMLElement,
  node: Node,
  offset: number,
): number | null {
  try {
    const prefix = document.createRange();
    prefix.selectNodeContents(editor);
    prefix.setEnd(node, offset);
    return prefix.toString().length;
  } catch {
    return null;
  }
}

function textPointAtOffset(
  editor: HTMLElement,
  requestedOffset: number,
): { node: Node; offset: number } | null {
  const walker = document.createTreeWalker(editor, NodeFilter.SHOW_TEXT);
  let remaining = requestedOffset;
  let node = walker.nextNode();
  while (node) {
    const length = node.textContent?.length ?? 0;
    if (remaining <= length) return { node, offset: remaining };
    remaining -= length;
    node = walker.nextNode();
  }
  return requestedOffset === 0 ? { node: editor, offset: 0 } : null;
}

function activeSpreadsheetTextEditor(editor: HTMLElement): boolean {
  if (!editor.isConnected) return false;
  if (editor.classList.contains('fortune-fx-input')) return true;
  const inputBox = editor.closest<HTMLElement>('.luckysheet-input-box');
  if (!inputBox) return true;
  const zIndex = Number.parseInt(
    inputBox.style.zIndex || getComputedStyle(inputBox).zIndex,
    10,
  );
  return !Number.isFinite(zIndex) || zIndex >= 0;
}

function restoreSpreadsheetTextSelection(
  snapshot: SpreadsheetRichTextSelectionSnapshot,
): HTMLElement | null {
  const editor = currentSpreadsheetTextEditor(snapshot);
  if (!editor) return null;
  const start = textPointAtOffset(editor, snapshot.start);
  const end = textPointAtOffset(editor, snapshot.end);
  const selection = window.getSelection();
  if (!start || !end || !selection) return null;
  const range = document.createRange();
  range.setStart(start.node, start.offset);
  range.setEnd(end.node, end.offset);
  selection.removeAllRanges();
  selection.addRange(range);
  editor.focus({ preventScroll: true });
  return editor;
}

function currentSpreadsheetTextEditor(
  snapshot: SpreadsheetRichTextSelectionSnapshot,
): HTMLElement | null {
  const expectedText =
    xlsxRichTextCellText(snapshot.sourceCell) ?? snapshot.sourceCell.v;
  const matchesSnapshot = (candidate: HTMLElement): boolean =>
    activeSpreadsheetTextEditor(candidate) &&
    typeof expectedText === 'string' &&
    editorText(candidate) === expectedText;
  if (matchesSnapshot(snapshot.editor)) return snapshot.editor;
  if (!snapshot.root.isConnected) return null;
  const selector = snapshot.editor.classList.contains('fortune-fx-input')
    ? '.fortune-fx-input'
    : '.luckysheet-cell-input';
  return (
    Array.from(snapshot.root.querySelectorAll<HTMLElement>(selector)).find(
      matchesSnapshot,
    ) ?? null
  );
}

function editorText(editor: HTMLElement): string {
  return editor.textContent ?? '';
}

function focusedAxisIndex(value: unknown, axis: readonly number[]): number {
  const minimum = finiteIndex(axis[0], 0);
  const maximum = finiteIndex(axis[1], minimum);
  const focus = finiteIndex(value, minimum);
  return Math.min(
    Math.max(minimum, maximum),
    Math.max(Math.min(minimum, maximum), focus),
  );
}

function finiteIndex(value: unknown, fallback: number): number {
  return typeof value === 'number' && Number.isFinite(value)
    ? Math.max(0, Math.trunc(value))
    : fallback;
}

function sheetCellAt(
  content: WorkSpreadsheetContent,
  sheetId: string,
  row: number,
  column: number,
): Cell | null {
  const sheet = content.sheets.find((candidate) => candidate.id === sheetId);
  return (
    sheet?.data?.[row]?.[column] ??
    sheet?.celldata?.find(
      (candidate) => candidate.r === row && candidate.c === column,
    )?.v ??
    null
  );
}

function isSpreadsheetRichTextFormatAttribute(
  value: keyof Cell,
): value is SpreadsheetRichTextFormatAttribute {
  return ['bl', 'cl', 'fc', 'ff', 'fs', 'it', 'un'].includes(value);
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}
