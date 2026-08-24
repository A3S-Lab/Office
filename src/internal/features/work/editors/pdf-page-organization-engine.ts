import { PDFDocument, PDFName, degrees } from 'pdf-lib';
import {
  validatedPdfInsertionIndex,
  validatedPdfPageIndexes,
  validatedPdfPageOrder,
  validatedPdfPageSize,
  validatedPdfSplitBoundaries,
} from './pdf-page-organization-plan';
import {
  PDF_PAGE_ORGANIZATION_MAX_IMPORT_BYTES,
  PDF_PAGE_ORGANIZATION_MAX_PAGES,
  PDF_PAGE_ORGANIZATION_MAX_SOURCE_BYTES,
  PdfPageOrganizationError,
  type PdfPageOrganizationDiagnostic,
  type PdfPageOrganizationExportOperation,
  type PdfPageOrganizationJob,
  type PdfPageOrganizationMutation,
  type PdfPageOrganizationResult,
  type PdfPageOrganizationSource,
} from './pdf-page-organization-types';

const RISKY_STRUCTURE_KEYS = [
  'AcroForm',
  'Outlines',
  'StructTreeRoot',
] as const;
const COPY_DIAGNOSTIC: PdfPageOrganizationDiagnostic = {
  code: 'pdf.pages.catalog-not-copied',
  message:
    'Page-only exports do not copy document-level outlines, forms, tags, attachments, scripts, or signatures.',
};

export async function applyPdfPageOrganizationJob(
  job: PdfPageOrganizationJob,
): Promise<PdfPageOrganizationResult> {
  const source = await loadBoundedPdf(
    job.source,
    PDF_PAGE_ORGANIZATION_MAX_SOURCE_BYTES,
    'source',
  );
  assertPageCount(source.getPageCount());
  if (job.kind === 'export') return exportPages(source, job.operation);
  assertUnsigned(source);
  return mutatePages(source, job.mutation, job.importSource);
}

async function mutatePages(
  source: PDFDocument,
  mutation: PdfPageOrganizationMutation,
  importSource: PdfPageOrganizationSource | undefined,
): Promise<PdfPageOrganizationResult> {
  const pageCount = source.getPageCount();
  if (mutation.kind === 'rotate') {
    const indexes = validatedPdfPageIndexes(mutation.pageIndexes, pageCount);
    for (const pageIndex of indexes) {
      const page = source.getPage(pageIndex);
      const current = page.getRotation().angle;
      if (!Number.isSafeInteger(current) || current % 90 !== 0) {
        throw new PdfPageOrganizationError(
          'pdf.pages.unsupported-rotation',
          `Page ${pageIndex + 1} has unsupported rotation ${String(current)}.`,
        );
      }
      page.setRotation(degrees(normalizeDegrees(current + mutation.degrees)));
    }
    return mutationResult(source);
  }
  if (mutation.kind === 'insert-blank') {
    const index = validatedPdfInsertionIndex(mutation.index, pageCount);
    const size = validatedPdfPageSize(mutation.size);
    assertPageCount(pageCount + 1);
    source.insertPage(index, [size.width, size.height]);
    return mutationResult(source);
  }
  if (mutation.kind === 'delete') {
    assertNoRiskyStructure(source);
    const indexes = validatedPdfPageIndexes(mutation.pageIndexes, pageCount);
    if (indexes.length >= pageCount) {
      throw new PdfPageOrganizationError(
        'pdf.pages.empty-document',
        'A PDF must retain at least one page.',
      );
    }
    for (const pageIndex of [...indexes].sort((left, right) => right - left)) {
      source.removePage(pageIndex);
    }
    return mutationResult(source);
  }
  if (mutation.kind === 'reorder') {
    assertNoRiskyStructure(source);
    const pageOrder = validatedPdfPageOrder(mutation.pageOrder, pageCount);
    const reordered = await PDFDocument.create({ updateMetadata: false });
    copyMetadata(source, reordered);
    const pages = await reordered.copyPages(source, pageOrder);
    for (const page of pages) reordered.addPage(page);
    return mutationResult(reordered);
  }

  const index = validatedPdfInsertionIndex(mutation.index, pageCount);
  if (!importSource) {
    throw new PdfPageOrganizationError(
      'pdf.pages.missing-import',
      'Choose a PDF to merge before applying this operation.',
    );
  }
  assertNoRiskyStructure(source);
  const imported = await loadBoundedPdf(
    importSource,
    PDF_PAGE_ORGANIZATION_MAX_IMPORT_BYTES,
    'import',
  );
  assertSafeImport(imported);
  assertPageCount(pageCount + imported.getPageCount());
  const copied = await source.copyPages(imported, imported.getPageIndices());
  for (let offset = 0; offset < copied.length; offset += 1) {
    const page = copied[offset];
    if (page) source.insertPage(index + offset, page);
  }
  return mutationResult(source);
}

async function exportPages(
  source: PDFDocument,
  operation: PdfPageOrganizationExportOperation,
): Promise<PdfPageOrganizationResult> {
  const pageCount = source.getPageCount();
  if (operation.kind === 'extract') {
    const indexes = validatedPdfPageIndexes(operation.pageIndexes, pageCount, {
      preserveOrder: true,
    });
    return {
      diagnostics: sourceHasRiskyStructure(source) ? [COPY_DIAGNOSTIC] : [],
      files: [await copyPageRange(source, indexes)],
      kind: 'exported',
    };
  }
  const boundaries = validatedPdfSplitBoundaries(
    operation.splitAfterPageIndexes,
    pageCount,
  );
  const ranges: number[][] = [];
  let start = 0;
  for (const boundary of boundaries) {
    ranges.push(range(start, boundary + 1));
    start = boundary + 1;
  }
  ranges.push(range(start, pageCount));
  return {
    diagnostics: sourceHasRiskyStructure(source) ? [COPY_DIAGNOSTIC] : [],
    files: await Promise.all(
      ranges.map((indexes) => copyPageRange(source, indexes)),
    ),
    kind: 'exported',
  };
}

async function copyPageRange(source: PDFDocument, indexes: number[]) {
  const output = await PDFDocument.create({ updateMetadata: false });
  copyMetadata(source, output);
  const pages = await output.copyPages(source, indexes);
  for (const page of pages) output.addPage(page);
  return {
    pageCount: indexes.length,
    source: await output.save(),
  };
}

async function mutationResult(
  source: PDFDocument,
): Promise<PdfPageOrganizationResult> {
  return {
    diagnostics: [],
    kind: 'mutated',
    pageCount: source.getPageCount(),
    source: await source.save(),
  };
}

async function loadBoundedPdf(
  source: PdfPageOrganizationSource,
  byteLimit: number,
  label: 'import' | 'source',
): Promise<PDFDocument> {
  const byteLength = source.byteLength;
  if (byteLength < 5 || byteLength > byteLimit) {
    throw new PdfPageOrganizationError(
      'pdf.pages.source-limit',
      `The PDF ${label} must be between 5 bytes and ${byteLimit} bytes.`,
    );
  }
  try {
    const pdf = await PDFDocument.load(source, {
      ignoreEncryption: false,
      throwOnInvalidObject: true,
      updateMetadata: false,
    });
    if (pdf.isEncrypted) {
      throw new PdfPageOrganizationError(
        'pdf.pages.encrypted-source',
        `The encrypted PDF ${label} cannot be reorganized without an authenticated decryption provider.`,
      );
    }
    return pdf;
  } catch (error) {
    if (error instanceof PdfPageOrganizationError) throw error;
    const message = error instanceof Error ? error.message : String(error);
    if (/encrypt/i.test(message)) {
      throw new PdfPageOrganizationError(
        'pdf.pages.encrypted-source',
        `The encrypted PDF ${label} cannot be reorganized without an authenticated decryption provider.`,
      );
    }
    throw new PdfPageOrganizationError(
      'pdf.pages.invalid-source',
      `The PDF ${label} is malformed or unsupported: ${message.slice(0, 240)}`,
    );
  }
}

function assertPageCount(pageCount: number): void {
  if (
    !Number.isSafeInteger(pageCount) ||
    pageCount < 1 ||
    pageCount > PDF_PAGE_ORGANIZATION_MAX_PAGES
  ) {
    throw new PdfPageOrganizationError(
      'pdf.pages.page-limit',
      `PDF page organization supports 1 to ${PDF_PAGE_ORGANIZATION_MAX_PAGES} pages per result.`,
    );
  }
}

function assertUnsigned(source: PDFDocument): void {
  if (source.catalog.has(PDFName.of('Perms'))) {
    throw new PdfPageOrganizationError(
      'pdf.pages.signed-source',
      'This PDF declares a signature permission dictionary. Rewriting it would invalidate its trust evidence.',
    );
  }
}

function assertNoRiskyStructure(source: PDFDocument): void {
  if (!sourceHasRiskyStructure(source)) return;
  throw new PdfPageOrganizationError(
    'pdf.pages.risky-structure',
    'Delete, reorder, and merge are disabled for PDFs with forms, outlines, or tagged structure because page references cannot be rewritten safely.',
  );
}

function assertSafeImport(source: PDFDocument): void {
  if (
    !sourceHasRiskyStructure(source) &&
    !source.catalog.has(PDFName.of('Perms'))
  ) {
    return;
  }
  throw new PdfPageOrganizationError(
    'pdf.pages.risky-import',
    'The PDF selected for merge contains forms, outlines, tags, or signature permissions that cannot be copied safely.',
  );
}

function sourceHasRiskyStructure(source: PDFDocument): boolean {
  return RISKY_STRUCTURE_KEYS.some((key) =>
    source.catalog.has(PDFName.of(key)),
  );
}

function copyMetadata(source: PDFDocument, target: PDFDocument): void {
  setIfDefined(source.getTitle(), (value) => target.setTitle(value));
  setIfDefined(source.getAuthor(), (value) => target.setAuthor(value));
  setIfDefined(source.getSubject(), (value) => target.setSubject(value));
  setIfDefined(source.getCreator(), (value) => target.setCreator(value));
  setIfDefined(source.getProducer(), (value) => target.setProducer(value));
  setIfDefined(source.getCreationDate(), (value) =>
    target.setCreationDate(value),
  );
  setIfDefined(source.getModificationDate(), (value) =>
    target.setModificationDate(value),
  );
  const keywords = source.getKeywords();
  if (keywords) target.setKeywords([keywords]);
}

function setIfDefined<T>(
  value: T | undefined,
  setter: (value: T) => void,
): void {
  if (value !== undefined) setter(value);
}

function normalizeDegrees(value: number): 0 | 90 | 180 | 270 {
  const normalized = ((value % 360) + 360) % 360;
  if (
    normalized === 0 ||
    normalized === 90 ||
    normalized === 180 ||
    normalized === 270
  ) {
    return normalized;
  }
  throw new PdfPageOrganizationError(
    'pdf.pages.unsupported-rotation',
    `PDF page rotation ${value} is not a multiple of 90 degrees.`,
  );
}

function range(start: number, end: number): number[] {
  return Array.from(
    { length: Math.max(0, end - start) },
    (_, index) => start + index,
  );
}
