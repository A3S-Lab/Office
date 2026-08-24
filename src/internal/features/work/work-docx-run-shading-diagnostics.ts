import { DOCX_WORDPROCESSING_NAMESPACES } from './work-docx-ignorable-extension-preservation';
import { inspectDocxRunShading } from './work-docx-run-shading';
import { createDocxThemeResolver } from './work-docx-theme';
import {
  descendants,
  directChildren,
  type OoxmlPackage,
} from './work-ooxml-package';
import type { WorkCompatibilityIssue } from './work-types';

const RUN_SHADING_PART_PATTERN =
  /^word\/(?:document|styles|header\d*|footer\d*|footnotes|endnotes|comments)\.xml$/i;
const VALID_SHADING_PARENTS = new Set(['rPr', 'pPr', 'tcPr', 'tblPr']);

export async function diagnoseDocxRunShading(
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
      RUN_SHADING_PART_PATTERN.test(path)
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
      if (!directChildren(properties, 'shd').length) continue;
      const inspection = inspectDocxRunShading(properties, theme);
      if (inspection.status === 'valid') {
        validCount += 1;
        if (inspection.value.pattern === 'nil') resetCount += 1;
      } else {
        invalidCount += 1 + inspection.spoofedCount;
      }
    }
    invalidCount += descendants(part, 'shd').filter((element) => {
      const parent = element.parentElement;
      return (
        DOCX_WORDPROCESSING_NAMESPACES.has(element.namespaceURI ?? '') &&
        (!parent || !VALID_SHADING_PARENTS.has(parent.localName))
      );
    }).length;
  }

  const issues: WorkCompatibilityIssue[] = [];
  if (validCount) {
    issues.push({
      code: 'docx.run-shading',
      severity: 'info',
      feature: 'Character shading',
      message: `${validCount} native character shading value(s), including ${resetCount} explicit reset(s), were detected across document defaults, styles, body, page chrome, notes, comments, and formatting revisions. Writer preserves every Word pattern mask, direct or automatic foreground and background colors, independently tinted or shaded theme references, and explicit nil semantics through editing, Format Painter, one-step Undo, and exact DOCX export. Native highlight remains independent and takes display precedence.`,
    });
  }
  if (invalidCount) {
    issues.push({
      code: 'docx.run-shading.invalid',
      severity: 'warning',
      feature: 'Character shading',
      message: `${invalidCount} malformed, duplicated, misplaced, namespace-spoofed, text-bearing, child-bearing, extra-attribute, or unresolved-theme character shading value(s) are normalized to explicit nil resets instead of enabling untrusted paint or inheriting stale formatting.`,
    });
  }
  return issues;
}
