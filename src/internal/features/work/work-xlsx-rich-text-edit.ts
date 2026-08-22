import type { Cell } from '@fortune-sheet/core';
import {
  normalizeXlsxSemanticColorOrigin,
  type XlsxSemanticColorOrigin,
  xlsxSemanticColorMatchesValue,
} from './work-xlsx-cell-style-origin';
import {
  coalesceXlsxRichTextRuns,
  isHighSurrogate,
  isLowSurrogate,
  MAX_XLSX_RICH_TEXT_CELL_CHARACTERS,
  MAX_XLSX_RICH_TEXT_RUNS_PER_CELL,
  type NormalizedXlsxRichTextCell,
  normalizeXlsxRichTextCell,
  normalizeXlsxRichTextEditSource,
  normalizeXlsxRichTextRun,
  validXlsxRichText,
  type XlsxRichTextRun,
} from './work-xlsx-rich-text-model';

interface IndexedRichTextRun {
  end: number;
  run: XlsxRichTextRun;
  start: number;
}

interface RichTextReplacement {
  index: number;
  insert: string;
  remove: number;
}

export interface SpreadsheetRichTextPasteIntent {
  end: number;
  runs: readonly XlsxRichTextRun[];
  start: number;
  text: string;
}

/**
 * Reconciles a text edit emitted by Fortune with the previous native XLSX
 * rich-text cell. Formula-bar edits may arrive as a plain string, while the
 * in-cell editor emits visible runs without A3S-only semantic color metadata.
 * This bounded model restores only style and metadata justified by the prior
 * controlled value and the exact text replacement.
 */
export function reconcileSpreadsheetRichTextCellEdit(
  previous: Cell | null | undefined,
  current: Cell,
  formattedPaste?: SpreadsheetRichTextPasteIntent,
): Cell {
  if (current.f) return current;
  const source = formattedPaste
    ? normalizeXlsxRichTextEditSource(previous)
    : previous
      ? normalizeXlsxRichTextCell(previous)
      : null;
  if (!source) return current;
  const live = liveRichTextValue(current);
  if (!live || live.text === '') return current;

  const replacement = richTextReplacement(source.text, live.text);
  const sourceRuns = indexRichTextRuns(source);
  const formattedPasteRuns = formattedPaste
    ? reconstructFormattedPasteRuns(
        source,
        sourceRuns,
        live.text,
        formattedPaste,
      )
    : null;
  const runs =
    formattedPasteRuns ??
    (!live.rich || live.text !== source.text
      ? reconstructPlainTextRuns(sourceRuns, replacement)
      : restoreSemanticColorOrigins(live.rich.runs, sourceRuns, replacement));
  const normalizedRuns = coalesceXlsxRichTextRuns(runs);
  if (
    !normalizedRuns.length ||
    normalizedRuns.length > MAX_XLSX_RICH_TEXT_RUNS_PER_CELL ||
    normalizedRuns.map((run) => run.v).join('') !== live.text
  ) {
    return current;
  }

  return spreadsheetRichTextCellWithRuns(
    previous,
    current,
    live.text,
    normalizedRuns,
  );
}

/**
 * Restores the controlled runs when an unauthenticated Fortune callback keeps
 * the exact visible text but drops native run metadata. Callers must reserve
 * this for non-structural batches without an exact cell-operation coordinate.
 */
export function restoreSpreadsheetRichTextCellRuns(
  previous: Cell | null | undefined,
  current: Cell,
): Cell {
  if (!previous || current.f) return current;
  const source = normalizeXlsxRichTextCell(previous);
  const live = liveRichTextValue(current);
  if (!source || !live || source.text !== live.text) return current;
  return spreadsheetRichTextCellWithRuns(
    previous,
    current,
    source.text,
    source.runs,
  );
}

function spreadsheetRichTextCellWithRuns(
  previous: Cell | null | undefined,
  current: Cell,
  text: string,
  runs: readonly XlsxRichTextRun[],
): Cell {
  const next: Cell = {
    ...current,
    ct: {
      ...previous?.ct,
      ...current.ct,
      s: runs.map((run) => ({ ...run })),
      t: 'inlineStr',
    },
    v: text,
  };
  delete next.f;
  delete next.m;
  return next;
}

/**
 * Proves that Fortune retained the visible text of a controlled native
 * rich-text cell. This permits metadata-only callbacks to restore lost run
 * structure without guessing across a text or structural edit.
 */
export function sameSpreadsheetRichTextCellText(
  previous: Cell | null | undefined,
  current: Cell,
): boolean {
  if (!previous || current.f) return false;
  const source = normalizeXlsxRichTextCell(previous);
  const live = liveRichTextValue(current);
  return Boolean(source && live && source.text === live.text);
}

function liveRichTextValue(
  cell: Cell,
): { rich: NormalizedXlsxRichTextCell | null; text: string } | null {
  const rich = normalizeXlsxRichTextCell(cell);
  if (typeof cell.v === 'string' && cell.v !== rich?.text) {
    return validBoundedText(cell.v) ? { rich: null, text: cell.v } : null;
  }
  if (rich) return { rich, text: rich.text };
  return typeof cell.v === 'string' && validBoundedText(cell.v)
    ? { rich: null, text: cell.v }
    : null;
}

function validBoundedText(value: string): boolean {
  return (
    value.length <= MAX_XLSX_RICH_TEXT_CELL_CHARACTERS &&
    validXlsxRichText(value)
  );
}

function richTextReplacement(
  previous: string,
  current: string,
): RichTextReplacement {
  let left = 0;
  let right = 0;
  while (
    left < previous.length &&
    left < current.length &&
    previous[left] === current[left]
  ) {
    left += 1;
  }
  if (left > 0 && isHighSurrogate(previous.charCodeAt(left - 1))) left -= 1;
  while (
    right + left < previous.length &&
    right + left < current.length &&
    previous[previous.length - right - 1] ===
      current[current.length - right - 1]
  ) {
    right += 1;
  }
  if (
    right > 0 &&
    isLowSurrogate(previous.charCodeAt(previous.length - right))
  ) {
    right -= 1;
  }
  return {
    index: left,
    insert: current.slice(left, current.length - right),
    remove: previous.length - left - right,
  };
}

function indexRichTextRuns(
  source: NormalizedXlsxRichTextCell,
): IndexedRichTextRun[] {
  let offset = 0;
  return source.runs.map((run) => {
    const indexed = { end: offset + run.v.length, run, start: offset };
    offset = indexed.end;
    return indexed;
  });
}

function reconstructPlainTextRuns(
  source: readonly IndexedRichTextRun[],
  replacement: RichTextReplacement,
): XlsxRichTextRun[] {
  const runs: XlsxRichTextRun[] = [];
  appendSourceRange(runs, source, 0, replacement.index);
  if (replacement.insert) {
    const inherited = insertionSourceRun(source, replacement);
    if (inherited) runs.push({ ...inherited.run, v: replacement.insert });
  }
  const sourceLength = source.at(-1)?.end ?? 0;
  appendSourceRange(
    runs,
    source,
    replacement.index + replacement.remove,
    sourceLength,
  );
  return runs;
}

function reconstructFormattedPasteRuns(
  source: NormalizedXlsxRichTextCell,
  sourceRuns: readonly IndexedRichTextRun[],
  currentText: string,
  intent: SpreadsheetRichTextPasteIntent,
): XlsxRichTextRun[] | null {
  if (
    !validRichTextPasteSelection(source.text, intent.start, intent.end) ||
    typeof intent.text !== 'string' ||
    !intent.text ||
    intent.text.length > MAX_XLSX_RICH_TEXT_CELL_CHARACTERS ||
    !validXlsxRichText(intent.text) ||
    !Array.isArray(intent.runs) ||
    !intent.runs.length ||
    intent.runs.length > MAX_XLSX_RICH_TEXT_RUNS_PER_CELL
  ) {
    return null;
  }
  const pastedRuns: XlsxRichTextRun[] = [];
  for (const candidate of intent.runs) {
    const run = normalizeXlsxRichTextRun(candidate);
    if (!run) return null;
    if (run.v) pastedRuns.push(run);
  }
  if (
    !pastedRuns.length ||
    pastedRuns.map((run) => run.v).join('') !== intent.text
  ) {
    return null;
  }
  const expectedText =
    source.text.slice(0, intent.start) +
    intent.text +
    source.text.slice(intent.end);
  if (
    expectedText !== currentText ||
    expectedText.length > MAX_XLSX_RICH_TEXT_CELL_CHARACTERS ||
    !validXlsxRichText(expectedText)
  ) {
    return null;
  }

  const runs: XlsxRichTextRun[] = [];
  appendSourceRange(runs, sourceRuns, 0, intent.start);
  runs.push(...pastedRuns);
  appendSourceRange(runs, sourceRuns, intent.end, source.text.length);
  const normalized = coalesceXlsxRichTextRuns(runs);
  return normalized.length > 0 &&
    normalized.length <= MAX_XLSX_RICH_TEXT_RUNS_PER_CELL
    ? normalized
    : null;
}

function validRichTextPasteSelection(
  text: string,
  start: number,
  end: number,
): boolean {
  return Boolean(
    Number.isInteger(start) &&
      Number.isInteger(end) &&
      start >= 0 &&
      start <= end &&
      end <= text.length &&
      !splitsSurrogatePair(text, start) &&
      !splitsSurrogatePair(text, end),
  );
}

function splitsSurrogatePair(text: string, offset: number): boolean {
  if (offset <= 0 || offset >= text.length) return false;
  return (
    isHighSurrogate(text.charCodeAt(offset - 1)) &&
    isLowSurrogate(text.charCodeAt(offset))
  );
}

function appendSourceRange(
  target: XlsxRichTextRun[],
  source: readonly IndexedRichTextRun[],
  start: number,
  end: number,
): void {
  if (start >= end) return;
  for (const indexed of source) {
    const overlapStart = Math.max(start, indexed.start);
    const overlapEnd = Math.min(end, indexed.end);
    if (overlapStart >= overlapEnd) continue;
    target.push({
      ...indexed.run,
      v: indexed.run.v.slice(
        overlapStart - indexed.start,
        overlapEnd - indexed.start,
      ),
    });
  }
}

function restoreSemanticColorOrigins(
  current: readonly XlsxRichTextRun[],
  source: readonly IndexedRichTextRun[],
  replacement: RichTextReplacement,
): XlsxRichTextRun[] {
  const runs: XlsxRichTextRun[] = [];
  let absoluteOffset = 0;
  for (const currentRun of current) {
    let localOffset = 0;
    while (localOffset < currentRun.v.length) {
      const width = codePointWidth(currentRun.v, localOffset);
      const origin = semanticOriginAt(
        source,
        replacement,
        absoluteOffset + localOffset,
      );
      const next = {
        ...currentRun,
        v: currentRun.v.slice(localOffset, localOffset + width),
      };
      delete next.a3sXlsxColorOrigin;
      if (
        origin &&
        currentRun.fc &&
        xlsxSemanticColorMatchesValue(origin, currentRun.fc)
      ) {
        next.a3sXlsxColorOrigin = origin;
      }
      runs.push(next);
      localOffset += width;
    }
    absoluteOffset += currentRun.v.length;
  }
  return runs;
}

function semanticOriginAt(
  source: readonly IndexedRichTextRun[],
  replacement: RichTextReplacement,
  currentOffset: number,
): XlsxSemanticColorOrigin | undefined {
  let indexed: IndexedRichTextRun | undefined;
  if (currentOffset < replacement.index) {
    indexed = sourceRunAt(source, currentOffset);
  } else if (currentOffset >= replacement.index + replacement.insert.length) {
    indexed = sourceRunAt(
      source,
      currentOffset - replacement.insert.length + replacement.remove,
    );
  } else {
    indexed = insertionSourceRun(source, replacement);
  }
  return normalizeXlsxSemanticColorOrigin(indexed?.run.a3sXlsxColorOrigin);
}

function insertionSourceRun(
  source: readonly IndexedRichTextRun[],
  replacement: RichTextReplacement,
): IndexedRichTextRun | undefined {
  const sourceLength = source.at(-1)?.end ?? 0;
  if (!sourceLength) return undefined;
  if (replacement.remove > 0 && replacement.index < sourceLength) {
    return sourceRunAt(source, replacement.index);
  }
  if (replacement.index > 0) {
    return sourceRunAt(source, replacement.index - 1);
  }
  return sourceRunAt(source, 0);
}

function sourceRunAt(
  source: readonly IndexedRichTextRun[],
  offset: number,
): IndexedRichTextRun | undefined {
  let left = 0;
  let right = source.length - 1;
  while (left <= right) {
    const middle = left + Math.floor((right - left) / 2);
    const indexed = source[middle];
    if (!indexed) return undefined;
    if (offset < indexed.start) right = middle - 1;
    else if (offset >= indexed.end) left = middle + 1;
    else return indexed;
  }
  return undefined;
}

function codePointWidth(value: string, offset: number): number {
  return isHighSurrogate(value.charCodeAt(offset)) &&
    isLowSurrogate(value.charCodeAt(offset + 1))
    ? 2
    : 1;
}
