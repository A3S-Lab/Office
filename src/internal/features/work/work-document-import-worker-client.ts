import {
  createLargeSimpleDocxParagraph,
  createLargeSimpleDocxParseResult,
  createLargeSimpleDocxTable,
  type LargeSimpleDocxParseOptions,
  type LargeSimpleDocxParseResult,
  type LargeSimpleDocxStreamResult,
} from './work-docx-large-document-import';
import type { DocumentImportWorkerResponse } from './work-document-import-worker-protocol';
import type { WorkDocumentNode } from './work-types';

const DOCUMENT_IMPORT_WORKER_TIMEOUT_MS = 120_000;

export type LargeSimpleDocxWorkerAttempt =
  | { result: LargeSimpleDocxParseResult; status: 'accepted' }
  | { status: 'ineligible' };

/**
 * Streams a parser-authenticated giant DOCX from a dedicated Worker. Null
 * means the Worker path failed and the caller should use the synchronous
 * parser; an explicit ineligible result must continue to the rich importer.
 */
export function parseLargeSimpleDocxInWorker(
  xmlBytes: ArrayBuffer,
  options: LargeSimpleDocxParseOptions = {},
  signal?: AbortSignal,
): Promise<LargeSimpleDocxWorkerAttempt | null> {
  if (typeof Worker === 'undefined') return Promise.resolve(null);
  if (signal?.aborted) return Promise.reject(documentImportAbortError());

  let worker: Worker;
  try {
    worker = new Worker(
      new URL('./work-document-import.worker.js', import.meta.url),
      { name: 'a3s-office-document-import' },
    );
  } catch {
    return Promise.resolve(null);
  }

  const startedAt = documentImportNow();
  return new Promise((resolve, reject) => {
    let settled = false;
    let timeout: ReturnType<typeof setTimeout> | undefined;
    const blocks: WorkDocumentNode[] = [];
    const html: string[] = [];
    let paragraphBatchCount = 0;
    let rowBatchCount = 0;
    let streamedAssemblyMs = 0;
    let tableRows: WorkDocumentNode[] | null = null;
    let tableHtml: string[] | null = null;

    const finish = (
      result: LargeSimpleDocxWorkerAttempt | null,
      error?: Error,
    ) => {
      if (settled) return;
      settled = true;
      if (timeout !== undefined) clearTimeout(timeout);
      signal?.removeEventListener('abort', handleAbort);
      worker.removeEventListener('error', handleFailure);
      worker.removeEventListener('message', handleMessage);
      worker.removeEventListener('messageerror', handleFailure);
      worker.terminate();
      if (error) reject(error);
      else resolve(result);
    };
    const handleAbort = () => finish(null, documentImportAbortError());
    const handleFailure = () => finish(null);
    const handleMessage = (
      event: MessageEvent<DocumentImportWorkerResponse>,
    ) => {
      const response = event.data;
      if (!response || typeof response !== 'object') {
        finish(null);
        return;
      }
      if (response.kind === 'failure') {
        finish(null);
        return;
      }
      if (response.kind === 'ineligible') {
        recordDocumentImportMeasure(
          'a3s-office.document.large-simple-worker-total',
          startedAt,
          documentImportNow(),
          { accepted: false, paragraphBatchCount, rowBatchCount },
        );
        finish({ status: 'ineligible' });
        return;
      }
      if (response.kind === 'paragraphs') {
        const batchStartedAt = documentImportNow();
        if (tableRows || tableHtml || !textBatchIsValid(response.texts)) {
          finish(null);
          return;
        }
        paragraphBatchCount += 1;
        for (const text of response.texts) {
          const paragraph = createLargeSimpleDocxParagraph(text);
          blocks.push(paragraph.node);
          html.push(paragraph.html);
        }
        streamedAssemblyMs += documentImportNow() - batchStartedAt;
        return;
      }
      if (response.kind === 'table-start') {
        if (tableRows || tableHtml) {
          finish(null);
          return;
        }
        tableRows = [];
        tableHtml = [];
        return;
      }
      if (response.kind === 'table-rows') {
        const batchStartedAt = documentImportNow();
        if (!tableRows || !tableHtml || !compactTableBatchIsValid(response)) {
          finish(null);
          return;
        }
        rowBatchCount += 1;
        let cellIndex = 0;
        let textIndex = 0;
        for (const cellCount of response.rowCellCounts) {
          const row = createCompactTableRow(
            response,
            cellCount,
            cellIndex,
            textIndex,
          );
          tableRows.push(row.node);
          tableHtml.push(row.html);
          cellIndex = row.cellIndex;
          textIndex = row.textIndex;
        }
        streamedAssemblyMs += documentImportNow() - batchStartedAt;
        return;
      }
      if (response.kind === 'table-end') {
        if (!tableRows?.length || !tableHtml?.length) {
          finish(null);
          return;
        }
        const table = createLargeSimpleDocxTable(tableRows, tableHtml);
        blocks.push(table.node);
        html.push(table.html);
        tableRows = null;
        tableHtml = null;
        return;
      }
      if (
        response.kind !== 'success' ||
        tableRows ||
        tableHtml ||
        !blocks.length ||
        !streamResultIsValid(response.streamed)
      ) {
        finish(null);
        return;
      }
      const finalizationStartedAt = documentImportNow();
      let result: LargeSimpleDocxParseResult;
      try {
        result = createLargeSimpleDocxParseResult(
          response.streamed,
          blocks,
          html,
          options,
        );
      } catch {
        finish(null);
        return;
      }
      const finishedAt = documentImportNow();
      recordDocumentImportDuration(
        'a3s-office.document.large-simple-worker-parse',
        response.timings.parseMs,
        {
          contentMs: response.timings.contentMs,
          eligibilityMs: response.timings.eligibilityMs,
          envelopeMs: response.timings.envelopeMs,
          xmlMs: response.timings.xmlMs,
        },
      );
      recordDocumentImportMeasure(
        'a3s-office.document.large-simple-worker-finalize',
        finalizationStartedAt,
        finishedAt,
        { paragraphBatchCount, rowBatchCount, streamedAssemblyMs },
      );
      recordDocumentImportMeasure(
        'a3s-office.document.large-simple-worker-total',
        startedAt,
        finishedAt,
        {
          accepted: true,
          paragraphBatchCount,
          rowBatchCount,
          streamedAssemblyMs,
        },
      );
      finish({ result, status: 'accepted' });
    };

    timeout = setTimeout(handleFailure, DOCUMENT_IMPORT_WORKER_TIMEOUT_MS);
    signal?.addEventListener('abort', handleAbort, { once: true });
    worker.addEventListener('error', handleFailure, { once: true });
    worker.addEventListener('message', handleMessage);
    worker.addEventListener('messageerror', handleFailure, { once: true });
    if (signal?.aborted) {
      handleAbort();
      return;
    }
    try {
      worker.postMessage(
        { kind: 'parse', options, xmlBytes },
        { transfer: [xmlBytes] },
      );
    } catch {
      handleFailure();
    }
  });
}

function textBatchIsValid(texts: unknown): texts is string[] {
  return (
    Array.isArray(texts) &&
    texts.length > 0 &&
    texts.every((value) => typeof value === 'string')
  );
}

function compactTableBatchIsValid(
  response: Extract<DocumentImportWorkerResponse, { kind: 'table-rows' }>,
): boolean {
  if (
    !(response.rowCellCounts instanceof Uint32Array) ||
    !response.rowCellCounts.length ||
    !(response.cellParagraphCounts instanceof Uint32Array) ||
    !textBatchIsValid(response.texts)
  ) {
    return false;
  }
  let cellCount = 0;
  for (const count of response.rowCellCounts) {
    if (!count) return false;
    cellCount += count;
  }
  if (cellCount !== response.cellParagraphCounts.length) return false;
  let paragraphCount = 0;
  for (const count of response.cellParagraphCounts) {
    if (!count) return false;
    paragraphCount += count;
  }
  return paragraphCount === response.texts.length;
}

function createCompactTableRow(
  response: Extract<DocumentImportWorkerResponse, { kind: 'table-rows' }>,
  cellCount: number,
  initialCellIndex: number,
  initialTextIndex: number,
): {
  cellIndex: number;
  html: string;
  node: WorkDocumentNode;
  textIndex: number;
} {
  const cellNodes: WorkDocumentNode[] = [];
  const cellHtml: string[] = [];
  let cellIndex = initialCellIndex;
  let textIndex = initialTextIndex;
  for (let cell = 0; cell < cellCount; cell += 1) {
    const paragraphCount = response.cellParagraphCounts[cellIndex] ?? 0;
    const paragraphNodes: WorkDocumentNode[] = [];
    let paragraphsHtml = '';
    for (let paragraph = 0; paragraph < paragraphCount; paragraph += 1) {
      const value = createLargeSimpleDocxParagraph(
        response.texts[textIndex] ?? '',
      );
      paragraphNodes.push(value.node);
      paragraphsHtml += value.html;
      textIndex += 1;
    }
    cellNodes.push({ type: 'tableCell', content: paragraphNodes });
    cellHtml.push(`<td>${paragraphsHtml}</td>`);
    cellIndex += 1;
  }
  return {
    cellIndex,
    html: `<tr>${cellHtml.join('')}</tr>`,
    node: { type: 'tableRow', content: cellNodes },
    textIndex,
  };
}

function streamResultIsValid(value: LargeSimpleDocxStreamResult): boolean {
  return (
    Boolean(value?.layout) &&
    Number.isSafeInteger(value.logicalBlockCount) &&
    value.logicalBlockCount > 0 &&
    Number.isSafeInteger(value.paragraphCount) &&
    value.paragraphCount > 0 &&
    Number.isSafeInteger(value.tableRowCount) &&
    value.tableRowCount >= 0
  );
}

function documentImportNow(): number {
  return globalThis.performance?.now?.() ?? Date.now();
}

function recordDocumentImportMeasure(
  name: string,
  start: number,
  end: number,
  detail?: Record<string, number | boolean>,
): void {
  try {
    globalThis.performance?.measure(name, { detail, end, start });
  } catch {
    // User Timing is diagnostic only and must never affect file import.
  }
}

function recordDocumentImportDuration(
  name: string,
  duration: number,
  detail?: Record<string, number>,
): void {
  if (!Number.isFinite(duration) || duration < 0) return;
  try {
    globalThis.performance?.measure(name, { detail, duration, start: 0 });
  } catch {
    // User Timing is diagnostic only and must never affect file import.
  }
}

function documentImportAbortError(): Error {
  if (typeof DOMException !== 'undefined') {
    return new DOMException('Document import was cancelled.', 'AbortError');
  }
  const error = new Error('Document import was cancelled.');
  error.name = 'AbortError';
  return error;
}
