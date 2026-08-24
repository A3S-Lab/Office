import { documentRunBorderIsVisible } from './work-document-run-border';
import { DOCX_WORDPROCESSING_NAMESPACES } from './work-docx-ignorable-extension-preservation';
import { inspectDocxRunBorder } from './work-docx-run-border';
import { createDocxThemeResolver } from './work-docx-theme';
import {
  descendants,
  directChildren,
  type OoxmlPackage,
} from './work-ooxml-package';
import type { WorkCompatibilityIssue } from './work-types';

const RUN_BORDER_PART_PATTERN =
  /^word\/(?:document|styles|header\d*|footer\d*|footnotes|endnotes|comments)\.xml$/i;

export async function diagnoseDocxRunBorders(
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
  const parts: Document[] = [document];
  for (const path of archive.paths('word/')) {
    if (
      path.toLowerCase() !== 'word/document.xml' &&
      RUN_BORDER_PART_PATTERN.test(path)
    ) {
      parts.push(await archive.xml(path));
    }
  }

  let validCount = 0;
  let resetCount = 0;
  let invalidCount = 0;
  for (const part of parts) {
    for (const properties of descendants(part, 'rPr').filter((element) =>
      DOCX_WORDPROCESSING_NAMESPACES.has(element.namespaceURI ?? ''),
    )) {
      if (!directChildren(properties, 'bdr').length) continue;
      const inspection = inspectDocxRunBorder(properties, theme);
      if (inspection.status === 'valid') {
        validCount += 1;
        if (!documentRunBorderIsVisible(inspection.value)) resetCount += 1;
      } else {
        invalidCount += 1 + inspection.spoofedCount;
      }
    }
    invalidCount += descendants(part, 'bdr').filter((element) => {
      const parent = element.parentElement;
      return (
        !parent ||
        parent.localName !== 'rPr' ||
        !DOCX_WORDPROCESSING_NAMESPACES.has(parent.namespaceURI ?? '')
      );
    }).length;
  }

  const issues: WorkCompatibilityIssue[] = [];
  if (validCount) {
    issues.push({
      code: 'docx.run-borders',
      severity: 'info',
      feature: 'Character borders',
      message: `${validCount} native character border(s), including ${resetCount} explicit reset(s), were detected across document defaults, styles, body, page chrome, notes, comments, and formatting revisions. Writer preserves all 25 visible Word line styles plus nil and none, direct and theme colors, eighth-point widths, spacing, shadow, and frame semantics through editing, Format Painter, one-step Undo, and exact DOCX export. The browser and PDF paint is a bounded CSS approximation, and no non-standard shortcut is invented.`,
    });
  }
  if (invalidCount) {
    issues.push({
      code: 'docx.run-borders.invalid',
      severity: 'warning',
      feature: 'Character borders',
      message: `${invalidCount} malformed, duplicated, misplaced, namespace-spoofed, text-bearing, child-bearing, extra-attribute, art-style, out-of-range, or unresolved-theme character border value(s) are normalized to explicit nil resets instead of enabling untrusted paint or inheriting stale formatting.`,
    });
  }
  return issues;
}
