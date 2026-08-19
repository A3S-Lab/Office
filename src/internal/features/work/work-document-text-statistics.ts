import type { Editor } from '@tiptap/core';
import type { Node as ProseMirrorNode } from '@tiptap/pm/model';
import {
  DOCUMENT_LAZY_BLOCK_NODE,
  DOCUMENT_LAZY_POSITION_BOUNDARY,
} from './work-document-lazy-model';

export interface DocumentTextStatistics {
  characterCountWithSpaces: number;
  characterCountWithoutSpaces: number;
  paragraphCount: number;
  wordCount: number;
}

const statisticsByDocument = new WeakMap<
  ProseMirrorNode,
  DocumentTextStatistics
>();

export function documentWordCount(value: string): number {
  let asciiCount = 0;
  let inAsciiWord = false;
  for (let index = 0; index < value.length; index += 1) {
    const code = value.charCodeAt(index);
    const asciiLetterOrNumber =
      (code >= 0x30 && code <= 0x39) ||
      (code >= 0x41 && code <= 0x5a) ||
      (code >= 0x61 && code <= 0x7a);
    if (asciiLetterOrNumber) {
      if (!inAsciiWord) asciiCount += 1;
      inAsciiWord = true;
      continue;
    }
    inAsciiWord = false;
    if (code <= 0x7f || code === 0xfffc) continue;
    return unicodeDocumentWordCount(value);
  }
  return asciiCount;
}

function unicodeDocumentWordCount(value: string): number {
  let count = 0;
  for (const _match of value.matchAll(
    /[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Hangul}]|[\p{L}\p{N}]+/gu,
  )) {
    count += 1;
  }
  return count;
}

/**
 * Counts a document once per immutable ProseMirror root. Selection, pagination,
 * and viewport transactions retain the same root and therefore stay O(1).
 */
export function documentTextStatistics(
  editor: Pick<Editor, 'getText' | 'state'>,
): DocumentTextStatistics {
  const document = editor.state.doc;
  const cached = statisticsByDocument.get(document);
  if (cached) return cached;

  const source = editor.getText({ blockSeparator: '\n' });
  let characterCountWithSpaces = 0;
  let characterCountWithoutSpaces = 0;
  for (let index = 0; index < source.length; index += 1) {
    const codePoint = source.codePointAt(index);
    if (codePoint === undefined) continue;
    if (codePoint > 0xffff) index += 1;
    if (
      codePoint === DOCUMENT_LAZY_POSITION_BOUNDARY.codePointAt(0) ||
      codePoint === 0x0a ||
      codePoint === 0x0d
    ) {
      continue;
    }
    characterCountWithSpaces += 1;
    if (!isEcmaScriptWhitespace(codePoint)) {
      characterCountWithoutSpaces += 1;
    }
  }

  let paragraphCount = 0;
  document.descendants((node) => {
    if (node.type.name === DOCUMENT_LAZY_BLOCK_NODE) {
      const logicalParagraphs = Number(node.attrs.paragraphCount);
      paragraphCount +=
        Number.isSafeInteger(logicalParagraphs) && logicalParagraphs >= 0
          ? logicalParagraphs
          : 0;
      return false;
    }
    if (node.isTextblock) paragraphCount += 1;
    return true;
  });

  const statistics = {
    characterCountWithSpaces,
    characterCountWithoutSpaces,
    paragraphCount,
    wordCount: documentWordCount(source),
  };
  statisticsByDocument.set(document, statistics);
  return statistics;
}

/** Keeps logical statistics unchanged across equal-size lazy hydration. */
export function transferDocumentTextStatistics(
  previous: ProseMirrorNode,
  next: ProseMirrorNode,
): void {
  const statistics = statisticsByDocument.get(previous);
  if (statistics) statisticsByDocument.set(next, statistics);
}

/**
 * Applies text-only subtree deltas to an already cached document total. Large
 * lazy documents edit one hydrated chunk at a time, so recounting every
 * placeholder tape would make a local keystroke O(document size).
 */
export function transferChangedDocumentTextStatistics(
  previous: ProseMirrorNode,
  next: ProseMirrorNode,
  changes: readonly {
    after: ProseMirrorNode;
    before: ProseMirrorNode;
  }[],
): boolean {
  const cached = statisticsByDocument.get(previous);
  if (!cached || !changes.length) return false;
  const statistics = { ...cached };
  for (const { after, before } of changes) {
    applyDocumentTextStatisticsDelta(
      statistics,
      simpleDocumentSubtreeStatistics(before),
      -1,
    );
    applyDocumentTextStatisticsDelta(
      statistics,
      simpleDocumentSubtreeStatistics(after),
      1,
    );
  }
  if (Object.values(statistics).some((value) => value < 0)) return false;
  statisticsByDocument.set(next, statistics);
  return true;
}

function simpleDocumentSubtreeStatistics(
  root: ProseMirrorNode,
): DocumentTextStatistics {
  const statistics: DocumentTextStatistics = {
    characterCountWithSpaces: 0,
    characterCountWithoutSpaces: 0,
    paragraphCount: 0,
    wordCount: 0,
  };
  const visit = (node: ProseMirrorNode): void => {
    if (node.type.name === DOCUMENT_LAZY_BLOCK_NODE) {
      accumulateDocumentTextStatistics(statistics, node.textContent);
      const logicalParagraphs = Number(node.attrs.paragraphCount);
      statistics.paragraphCount +=
        Number.isSafeInteger(logicalParagraphs) && logicalParagraphs >= 0
          ? logicalParagraphs
          : 0;
      return;
    }
    if (node.isTextblock) {
      accumulateDocumentTextStatistics(statistics, node.textContent);
      statistics.paragraphCount += 1;
      return;
    }
    node.forEach(visit);
  };
  visit(root);
  return statistics;
}

function accumulateDocumentTextStatistics(
  statistics: DocumentTextStatistics,
  source: string,
): void {
  for (let index = 0; index < source.length; index += 1) {
    const codePoint = source.codePointAt(index);
    if (codePoint === undefined) continue;
    if (codePoint > 0xffff) index += 1;
    if (
      codePoint === DOCUMENT_LAZY_POSITION_BOUNDARY.codePointAt(0) ||
      codePoint === 0x0a ||
      codePoint === 0x0d
    ) {
      continue;
    }
    statistics.characterCountWithSpaces += 1;
    if (!isEcmaScriptWhitespace(codePoint)) {
      statistics.characterCountWithoutSpaces += 1;
    }
  }
  statistics.wordCount += documentWordCount(source);
}

function applyDocumentTextStatisticsDelta(
  target: DocumentTextStatistics,
  value: DocumentTextStatistics,
  direction: -1 | 1,
): void {
  target.characterCountWithSpaces += direction * value.characterCountWithSpaces;
  target.characterCountWithoutSpaces +=
    direction * value.characterCountWithoutSpaces;
  target.paragraphCount += direction * value.paragraphCount;
  target.wordCount += direction * value.wordCount;
}

function isEcmaScriptWhitespace(codePoint: number): boolean {
  return (
    (codePoint >= 0x09 && codePoint <= 0x0d) ||
    codePoint === 0x20 ||
    codePoint === 0xa0 ||
    codePoint === 0x1680 ||
    (codePoint >= 0x2000 && codePoint <= 0x200a) ||
    codePoint === 0x2028 ||
    codePoint === 0x2029 ||
    codePoint === 0x202f ||
    codePoint === 0x205f ||
    codePoint === 0x3000 ||
    codePoint === 0xfeff
  );
}
