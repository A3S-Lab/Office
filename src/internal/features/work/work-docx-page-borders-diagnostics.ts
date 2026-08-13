import { isDocumentParagraphArtBorderStyle } from './work-document-paragraph-borders';
import { DOCX_WORDPROCESSING_NAMESPACES } from './work-docx-ignorable-extension-preservation';
import { inspectDocxPageBorders } from './work-docx-page-borders-import';
import { createDocxThemeResolver } from './work-docx-theme';
import { attribute, descendants } from './work-ooxml-package';
import type { OoxmlPackage } from './work-ooxml-package';
import type { WorkCompatibilityIssue } from './work-types';

export async function diagnoseDocxPageBorders(
  archive: OoxmlPackage,
  document: Document,
): Promise<WorkCompatibilityIssue[]> {
  const themePath = archive
    .paths('word/theme/')
    .find((path) => /\/theme\d*\.xml$/i.test(path));
  const themeDocument = themePath ? await archive.xml(themePath) : null;
  const settingsDocument = archive.has('word/settings.xml')
    ? await archive.xml('word/settings.xml')
    : null;
  const theme = createDocxThemeResolver(themeDocument, settingsDocument);
  const sections = descendants(document, 'sectPr').filter(
    (element) =>
      DOCX_WORDPROCESSING_NAMESPACES.has(element.namespaceURI ?? '') &&
      !hasAncestor(element, 'sectPrChange'),
  );
  const inspected = sections.map((section) =>
    inspectDocxPageBorders(section, theme),
  );
  const valid = inspected.flatMap((item) =>
    item.status === 'valid' && item.pageBorders ? [item.pageBorders] : [],
  );
  const invalidCount = inspected.reduce(
    (count, item) => count + item.invalidCount,
    0,
  );
  const spoofedCount = inspected.reduce(
    (count, item) => count + item.spoofedCount,
    0,
  );
  if (!valid.length && !invalidCount && !spoofedCount) return [];

  const issues: WorkCompatibilityIssue[] = [];
  if (valid.length) {
    const edgeCount = valid.reduce(
      (count, pageBorders) =>
        count + Object.values(pageBorders.edges).filter(Boolean).length,
      0,
    );
    const artCount = valid.reduce(
      (count, pageBorders) =>
        count +
        Object.values(pageBorders.edges).filter(
          (border) => border && isDocumentParagraphArtBorderStyle(border.style),
        ).length,
      0,
    );
    issues.push({
      code: 'docx.page-borders',
      severity: 'info',
      feature: 'Page borders',
      message: `${valid.length} section page-border set(s) with ${edgeCount} native edge(s) preserve page/text offsets, first/not-first-page display, front/back z-order, theme colors, widths, spacing, shadows, frames, and all Word line or art styles. The paginated editor and PDF surface honor section/page visibility and Word defaults; ${artCount} art edge(s) use a bounded browser approximation while native OOXML remains authoritative on DOCX export.`,
    });
  }
  if (invalidCount || spoofedCount) {
    issues.push({
      code: 'docx.page-borders.invalid',
      severity: 'warning',
      feature: 'Page borders',
      message: `${invalidCount + spoofedCount} malformed, out-of-order, duplicated, namespace-spoofed, relationship-bound, or unresolved theme page-border value(s) are ignored or normalized to explicit nil edges instead of applying unsafe or stale section formatting.`,
    });
  }
  const modifiers = settingsDocument
    ? [
        'alignBordersAndEdges',
        'bordersDoNotSurroundHeader',
        'bordersDoNotSurroundFooter',
      ].filter((name) => enabledSetting(settingsDocument, name))
    : [];
  if (valid.length && modifiers.length) {
    issues.push({
      code: 'docx.page-borders.settings',
      severity: 'warning',
      feature: 'Page-border compatibility settings',
      message: `${modifiers.join(', ')} modifies Word's interaction between page borders, page edges, headers, or footers. Native section borders remain editable and exact, but the browser preview uses the standardized pgBorders geometry and may not reproduce these application-specific document-wide adjustments.`,
    });
  }
  return issues;
}

function enabledSetting(document: Document, localName: string): boolean {
  const element = descendants(document, localName).find((candidate) =>
    DOCX_WORDPROCESSING_NAMESPACES.has(candidate.namespaceURI ?? ''),
  );
  if (!element) return false;
  const value = attribute(element, 'val')?.trim().toLowerCase();
  return !value || !['0', 'false', 'off'].includes(value);
}

function hasAncestor(element: Element, localName: string): boolean {
  let ancestor = element.parentElement;
  while (ancestor) {
    if (ancestor.localName === localName) return true;
    ancestor = ancestor.parentElement;
  }
  return false;
}
