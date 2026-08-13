import { DOCX_WORDPROCESSING_NAMESPACES } from './work-docx-ignorable-extension-preservation';
import { markDocxParagraphShading } from './work-docx-paragraph-shading-import';
import { createDocxParagraphStyleResolver } from './work-docx-paragraph-styles';
import { createDocxTableStyleResolver } from './work-docx-table-styles';
import { createDocxThemeResolver } from './work-docx-theme';
import type { OoxmlPackage } from './work-ooxml-package';
import type { WorkCompatibilityIssue } from './work-types';

export async function diagnoseDocxParagraphShading(
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
  const markers = markDocxParagraphShading(
    document.cloneNode(true) as Document,
    createDocxParagraphStyleResolver(stylesDocument),
    createDocxThemeResolver(themeDocument, settingsDocument),
    createDocxTableStyleResolver(stylesDocument),
  );
  const nativeCount = Array.from(document.getElementsByTagName('*')).filter(
    (element) =>
      element.localName === 'shd' &&
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
    issues.push({
      code: 'docx.paragraph-shading',
      severity: 'info',
      feature: 'Paragraph shading',
      message: `${markers.paragraphs.length} paragraph shading value(s) resolve through document defaults, based-on paragraph styles, conditional table styles, and direct formatting. All Word pattern masks, direct or automatic foreground/background colors, and independently tinted or shaded theme references remain editable and round-trip to native w:shd markup.`,
    });
  }
  if (markers.invalidCount || markers.spoofedCount) {
    issues.push({
      code: 'docx.paragraph-shading.invalid',
      severity: 'warning',
      feature: 'Paragraph shading',
      message: `${markers.invalidCount + markers.spoofedCount} malformed, duplicated, namespace-spoofed, relationship-bound, or unresolved theme paragraph shading value(s) are ignored or normalized to an explicit nil reset instead of inheriting stale formatting.`,
    });
  }
  return issues;
}
