import { DOCX_WORDPROCESSING_NAMESPACES } from './work-docx-ignorable-extension-preservation';
import {
  inspectDocxPageSize,
  inspectDocxPaperSource,
} from './work-docx-page-size-import';
import { descendants } from './work-ooxml-package';
import type { WorkCompatibilityIssue } from './work-types';

export function diagnoseDocxPageSize(
  document: Document,
): WorkCompatibilityIssue[] {
  const sections = descendants(document, 'sectPr').filter(
    (element) =>
      DOCX_WORDPROCESSING_NAMESPACES.has(element.namespaceURI ?? '') &&
      !hasAncestor(element, 'sectPrChange'),
  );
  const sizes = sections.map(inspectDocxPageSize);
  const sources = sections.map(inspectDocxPaperSource);
  const validSizes = sizes.flatMap((item) =>
    item.status === 'valid' && item.pageGeometry ? [item.pageGeometry] : [],
  );
  const validSources = sources.flatMap((item) =>
    item.status === 'valid' && item.paperSource ? [item.paperSource] : [],
  );
  const invalidCount = [...sizes, ...sources].reduce(
    (count, item) => count + item.invalidCount + item.spoofedCount,
    0,
  );
  if (!validSizes.length && !validSources.length && !invalidCount) return [];

  const issues: WorkCompatibilityIssue[] = [];
  if (validSizes.length) {
    const customCodes = validSizes.filter(
      (geometry) => geometry.code !== undefined,
    ).length;
    const explicitOrientations = validSizes.filter(
      (geometry) => geometry.orientation !== undefined,
    ).length;
    issues.push({
      code: 'docx.page-size',
      severity: 'info',
      feature: 'Page size',
      message: `${validSizes.length} explicit section page-size set(s) preserve exact native width and height in twips, ${explicitOrientations} explicit orientation value(s), and ${customCodes} Word paper code(s) through editing, preview, PDF, and DOCX output. Strict OOXML universal measures are accepted only when they convert to an exact integer twip.`,
    });
  }
  if (validSources.length) {
    issues.push({
      code: 'docx.paper-source',
      severity: 'info',
      feature: 'Printer paper source',
      message: `${validSources.length} section printer paper-source set(s) preserve their exact first-page and other-page tray codes in DOCX. The browser does not reinterpret printer-driver-specific tray numbers.`,
    });
  }
  if (invalidCount) {
    issues.push({
      code: 'docx.page-size.invalid',
      severity: 'warning',
      feature: 'Page setup',
      message: `${invalidCount} malformed, duplicated, incomplete, out-of-range, namespace-spoofed, relationship-bound, or structurally invalid pgSz or paperSrc value(s) are ignored instead of becoming trusted page or printer geometry.`,
    });
  }
  const bounded = validSizes.filter(
    (geometry) => geometry.width < 1_440 || geometry.height < 1_440,
  ).length;
  if (bounded) {
    issues.push({
      code: 'docx.page-size.browser-bounds',
      severity: 'warning',
      feature: 'Small page rendering',
      message: `${bounded} valid section page-size set(s) use an edge below one inch. Exact native values remain authoritative for DOCX output; browser editing and PDF projection bound the affected edge to one inch so the page remains inspectable.`,
    });
  }
  if (new Set(validSizes.map((geometry) => geometryKey(geometry))).size > 1) {
    issues.push({
      code: 'docx.page-size.mixed-live-layout',
      severity: 'warning',
      feature: 'Mixed section page sizes',
      message:
        'Static print preview and PDF use each section page size exactly. The continuous live paginated editor currently uses the active page metrics for its shared page stack, so mixed-size section transitions may paginate approximately while their native section geometry remains intact.',
    });
  }
  return issues;
}

function geometryKey(geometry: {
  width: number;
  height: number;
  orientation?: string;
}): string {
  return `${geometry.width}:${geometry.height}:${geometry.orientation ?? ''}`;
}

function hasAncestor(element: Element, localName: string): boolean {
  let ancestor = element.parentElement;
  while (ancestor) {
    if (ancestor.localName === localName) return true;
    ancestor = ancestor.parentElement;
  }
  return false;
}
