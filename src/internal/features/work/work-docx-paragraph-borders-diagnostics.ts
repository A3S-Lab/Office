import { isDocumentParagraphArtBorderStyle } from './work-document-paragraph-borders';
import { DOCX_WORDPROCESSING_NAMESPACES } from './work-docx-ignorable-extension-preservation';
import { markDocxParagraphBorders } from './work-docx-paragraph-borders-import';
import { createDocxParagraphStyleResolver } from './work-docx-paragraph-styles';
import { createDocxTableStyleResolver } from './work-docx-table-styles';
import { createDocxThemeResolver } from './work-docx-theme';
import type { OoxmlPackage } from './work-ooxml-package';
import type { WorkCompatibilityIssue } from './work-types';

export async function diagnoseDocxParagraphBorders(
  archive: OoxmlPackage,
  document: Document,
): Promise<WorkCompatibilityIssue[]> {
  const stylesDocument = archive.has('word/styles.xml')
    ? await archive.xml('word/styles.xml')
    : null;
  const themePath = archive
    .paths('word/theme/')
    .find((path) => /\/theme\d*\.xml$/i.test(path));
  const themeDocument = themePath ? await archive.xml(themePath) : null;
  const settingsDocument = archive.has('word/settings.xml')
    ? await archive.xml('word/settings.xml')
    : null;
  const markers = markDocxParagraphBorders(
    document.cloneNode(true) as Document,
    createDocxParagraphStyleResolver(stylesDocument),
    createDocxThemeResolver(themeDocument, settingsDocument),
    createDocxTableStyleResolver(stylesDocument),
  );
  const nativeCount = Array.from(document.getElementsByTagName('*')).filter(
    (element) =>
      element.localName === 'pBdr' &&
      element.parentElement?.localName === 'pPr' &&
      DOCX_WORDPROCESSING_NAMESPACES.has(element.namespaceURI ?? ''),
  ).length;
  if (
    !nativeCount &&
    !markers.paragraphs.length &&
    !markers.invalidCount &&
    !markers.spoofedCount
  ) {
    return [];
  }
  const issues: WorkCompatibilityIssue[] = [];
  if (markers.paragraphs.length) {
    const artCount = markers.paragraphs.reduce(
      (count, paragraph) =>
        count +
        Object.values(paragraph.borders).filter(
          (border) => border && isDocumentParagraphArtBorderStyle(border.style),
        ).length,
      0,
    );
    const specialCount = markers.paragraphs.reduce(
      (count, paragraph) =>
        count +
        Number(Boolean(paragraph.borders.between)) +
        Number(Boolean(paragraph.borders.bar)),
      0,
    );
    issues.push({
      code: 'docx.paragraph-borders',
      severity: 'info',
      feature: 'Paragraph borders',
      message: `${markers.paragraphs.length} paragraph border set(s) resolve through document defaults, based-on paragraph styles, conditional table styles, and direct formatting. All six native edges, theme colors, widths, spacing, shadows, frames, and all Word line or art styles round-trip; ${artCount + specialCount} art, between-paragraph, or facing-page edge(s) use a bounded browser approximation while retaining native OOXML semantics.`,
    });
  }
  if (markers.invalidCount || markers.spoofedCount) {
    issues.push({
      code: 'docx.paragraph-borders.invalid',
      severity: 'warning',
      feature: 'Paragraph borders',
      message: `${markers.invalidCount + markers.spoofedCount} malformed, out-of-order, duplicated, namespace-spoofed, relationship-bound, or unresolved theme paragraph border value(s) are ignored or normalized to explicit nil resets instead of inheriting stale formatting.`,
    });
  }
  return issues;
}
