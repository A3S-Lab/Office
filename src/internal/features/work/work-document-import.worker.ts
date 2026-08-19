import {
  streamLargeSimpleDocxDocumentXml,
  type LargeSimpleDocxStreamSink,
} from './work-docx-large-document-parser';
import {
  DOCUMENT_IMPORT_PARAGRAPH_BATCH_SIZE,
  DOCUMENT_IMPORT_TABLE_ROW_BATCH_SIZE,
  type DocumentImportWorkerRequest,
  type DocumentImportWorkerResponse,
} from './work-document-import-worker-protocol';
import { decodeXmlBytes } from './work-ooxml-xml';

interface DocumentImportWorkerScope {
  onmessage:
    | ((event: MessageEvent<DocumentImportWorkerRequest>) => void)
    | null;
  postMessage: (
    message: DocumentImportWorkerResponse,
    options?: StructuredSerializeOptions,
  ) => void;
}

const scope = globalThis as unknown as DocumentImportWorkerScope;

scope.onmessage = (event) => {
  if (event.data.kind !== 'parse') return;
  void parseDocument(event.data);
};

async function parseDocument(
  request: DocumentImportWorkerRequest,
): Promise<void> {
  try {
    const xmlStartedAt = workerNow();
    const source = decodeXmlBytes(
      new Uint8Array(request.xmlBytes),
      'word/document.xml',
    );
    const xmlFinishedAt = workerNow();
    const stream = createWorkerStream();
    const parseStartedAt = workerNow();
    const streamed = streamLargeSimpleDocxDocumentXml(
      source,
      request.options,
      stream.sink,
    );
    const parseFinishedAt = workerNow();
    if (!streamed || !stream.finish()) {
      scope.postMessage({ kind: 'ineligible' });
      return;
    }
    scope.postMessage({
      kind: 'success',
      streamed,
      timings: {
        contentMs: measureDuration('a3s-office.document.large-simple-content'),
        eligibilityMs: measureDuration(
          'a3s-office.document.large-simple-eligibility',
        ),
        envelopeMs: measureDuration(
          'a3s-office.document.large-simple-envelope',
        ),
        parseMs: parseFinishedAt - parseStartedAt,
        xmlMs: xmlFinishedAt - xmlStartedAt,
      },
    });
  } catch {
    scope.postMessage({ kind: 'failure' });
  }
}

function createWorkerStream(): {
  finish: () => boolean;
  sink: LargeSimpleDocxStreamSink;
} {
  let paragraphTexts: string[] = [];
  let rowCellCounts: number[] = [];
  let cellParagraphCounts: number[] = [];
  let tableTexts: string[] = [];
  let tableActive = false;

  const flushParagraphs = () => {
    if (!paragraphTexts.length) return;
    scope.postMessage({
      kind: 'paragraphs',
      texts: paragraphTexts,
    });
    paragraphTexts = [];
  };
  const flushRows = () => {
    if (!rowCellCounts.length) return;
    const transferredRowCellCounts = new Uint32Array(rowCellCounts);
    const transferredCellParagraphCounts = new Uint32Array(cellParagraphCounts);
    scope.postMessage(
      {
        cellParagraphCounts: transferredCellParagraphCounts,
        kind: 'table-rows',
        rowCellCounts: transferredRowCellCounts,
        texts: tableTexts,
      },
      {
        transfer: [
          transferredRowCellCounts.buffer,
          transferredCellParagraphCounts.buffer,
        ],
      },
    );
    rowCellCounts = [];
    cellParagraphCounts = [];
    tableTexts = [];
  };

  return {
    finish: () => {
      flushParagraphs();
      flushRows();
      return !tableActive;
    },
    sink: {
      paragraph: (text) => {
        if (tableActive) {
          throw new Error('A table stream cannot contain a paragraph batch.');
        }
        paragraphTexts.push(text);
        if (paragraphTexts.length >= DOCUMENT_IMPORT_PARAGRAPH_BATCH_SIZE) {
          flushParagraphs();
        }
      },
      tableEnd: () => {
        if (!tableActive) {
          throw new Error('A streamed table was not started.');
        }
        flushRows();
        scope.postMessage({ kind: 'table-end' });
        tableActive = false;
      },
      tableRow: ({ cellParagraphCounts: counts, texts }) => {
        if (!tableActive) {
          throw new Error('A streamed table row has no owning table.');
        }
        rowCellCounts.push(counts.length);
        cellParagraphCounts.push(...counts);
        tableTexts.push(...texts);
        if (rowCellCounts.length >= DOCUMENT_IMPORT_TABLE_ROW_BATCH_SIZE) {
          flushRows();
        }
      },
      tableStart: () => {
        if (tableActive) throw new Error('Nested tables are not supported.');
        flushParagraphs();
        tableActive = true;
        scope.postMessage({ kind: 'table-start' });
      },
    },
  };
}

function workerNow(): number {
  return globalThis.performance?.now?.() ?? Date.now();
}

function measureDuration(name: string): number {
  return globalThis.performance?.getEntriesByName?.(name).at(-1)?.duration ?? 0;
}
