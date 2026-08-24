import {
  PDF_PAGE_ORGANIZATION_MAX_PAGE_POINTS,
  PDF_PAGE_ORGANIZATION_MIN_PAGE_POINTS,
  PdfPageOrganizationError,
  type PdfPageOrganizationSize,
} from './pdf-page-organization-types';

export function validatedPdfPageIndexes(
  pageIndexes: readonly number[],
  pageCount: number,
  options: { allowEmpty?: boolean; preserveOrder?: boolean } = {},
): number[] {
  if (
    !Array.isArray(pageIndexes) ||
    (!options.allowEmpty && pageIndexes.length === 0)
  ) {
    throw new PdfPageOrganizationError(
      'pdf.pages.invalid-index',
      'Select at least one PDF page.',
    );
  }
  const seen = new Set<number>();
  for (const pageIndex of pageIndexes) {
    if (
      !Number.isSafeInteger(pageIndex) ||
      pageIndex < 0 ||
      pageIndex >= pageCount
    ) {
      throw new PdfPageOrganizationError(
        'pdf.pages.invalid-index',
        `Page index ${String(pageIndex)} is outside the current ${pageCount}-page PDF.`,
      );
    }
    if (seen.has(pageIndex)) {
      throw new PdfPageOrganizationError(
        'pdf.pages.duplicate-index',
        `Page index ${pageIndex} appears more than once.`,
      );
    }
    seen.add(pageIndex);
  }
  const result = [...pageIndexes];
  return options.preserveOrder
    ? result
    : result.sort((left, right) => left - right);
}

export function validatedPdfPageOrder(
  pageOrder: readonly number[],
  pageCount: number,
): number[] {
  if (!Array.isArray(pageOrder) || pageOrder.length !== pageCount) {
    throw invalidOrder(pageCount);
  }
  try {
    const order = validatedPdfPageIndexes(pageOrder, pageCount, {
      preserveOrder: true,
    });
    if (order.length === pageCount) return order;
  } catch {
    throw invalidOrder(pageCount);
  }
  throw invalidOrder(pageCount);
}

export function reorderedPdfPageIndexes(
  pageCount: number,
  pageIndexes: readonly number[],
  targetIndex: number,
): number[] {
  const selected = validatedPdfPageIndexes(pageIndexes, pageCount);
  if (
    !Number.isSafeInteger(targetIndex) ||
    targetIndex < 0 ||
    targetIndex > pageCount
  ) {
    throw new PdfPageOrganizationError(
      'pdf.pages.invalid-index',
      `Insertion gap ${String(targetIndex)} is outside the current ${pageCount}-page PDF.`,
    );
  }
  const selectedSet = new Set(selected);
  const remaining = Array.from(
    { length: pageCount },
    (_, index) => index,
  ).filter((index) => !selectedSet.has(index));
  const removedBeforeTarget = selected.filter(
    (index) => index < targetIndex,
  ).length;
  const insertionIndex = Math.max(
    0,
    Math.min(remaining.length, targetIndex - removedBeforeTarget),
  );
  remaining.splice(insertionIndex, 0, ...selected);
  return remaining;
}

export function validatedPdfInsertionIndex(
  index: number,
  pageCount: number,
): number {
  if (!Number.isSafeInteger(index) || index < 0 || index > pageCount) {
    throw new PdfPageOrganizationError(
      'pdf.pages.invalid-index',
      `Insertion gap ${String(index)} is outside the current ${pageCount}-page PDF.`,
    );
  }
  return index;
}

export function validatedPdfPageSize(
  size: PdfPageOrganizationSize | undefined,
): PdfPageOrganizationSize {
  const candidate = size ?? { height: 841.89, width: 595.28 };
  if (
    !Number.isFinite(candidate.width) ||
    !Number.isFinite(candidate.height) ||
    candidate.width < PDF_PAGE_ORGANIZATION_MIN_PAGE_POINTS ||
    candidate.height < PDF_PAGE_ORGANIZATION_MIN_PAGE_POINTS ||
    candidate.width > PDF_PAGE_ORGANIZATION_MAX_PAGE_POINTS ||
    candidate.height > PDF_PAGE_ORGANIZATION_MAX_PAGE_POINTS
  ) {
    throw new PdfPageOrganizationError(
      'pdf.pages.invalid-page-size',
      `PDF page dimensions must be between ${PDF_PAGE_ORGANIZATION_MIN_PAGE_POINTS} and ${PDF_PAGE_ORGANIZATION_MAX_PAGE_POINTS} points.`,
    );
  }
  return { height: candidate.height, width: candidate.width };
}

export function validatedPdfSplitBoundaries(
  splitAfterPageIndexes: readonly number[],
  pageCount: number,
): number[] {
  if (
    !Array.isArray(splitAfterPageIndexes) ||
    splitAfterPageIndexes.length === 0
  ) {
    throw invalidSplit(pageCount);
  }
  let previous = -1;
  for (const pageIndex of splitAfterPageIndexes) {
    if (
      !Number.isSafeInteger(pageIndex) ||
      pageIndex < 0 ||
      pageIndex >= pageCount - 1 ||
      pageIndex <= previous
    ) {
      throw invalidSplit(pageCount);
    }
    previous = pageIndex;
  }
  return [...splitAfterPageIndexes];
}

function invalidOrder(pageCount: number): PdfPageOrganizationError {
  return new PdfPageOrganizationError(
    'pdf.pages.invalid-order',
    `A page order must contain every index from 0 through ${Math.max(0, pageCount - 1)} exactly once.`,
  );
}

function invalidSplit(pageCount: number): PdfPageOrganizationError {
  return new PdfPageOrganizationError(
    'pdf.pages.invalid-split',
    `Split boundaries must be unique, ascending, and before the last page of the ${pageCount}-page PDF.`,
  );
}
