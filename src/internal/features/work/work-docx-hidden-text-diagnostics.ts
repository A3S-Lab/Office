import { DOCX_WORDPROCESSING_NAMESPACES } from './work-docx-ignorable-extension-preservation';
import { inspectDocxHiddenText } from './work-docx-hidden-text';
import {
  descendants,
  directChildren,
  type OoxmlPackage,
} from './work-ooxml-package';
import type { WorkCompatibilityIssue } from './work-types';

const HIDDEN_TEXT_DIAGNOSTIC_PART_PATTERN =
  /^word\/(?:document|styles|header\d*|footer\d*|footnotes|endnotes|comments)\.xml$/i;

export async function diagnoseDocxHiddenText(
  archive: OoxmlPackage,
  document: Document,
): Promise<WorkCompatibilityIssue[]> {
  const parts: Document[] = [document];
  for (const path of archive.paths('word/')) {
    if (
      path.toLowerCase() === 'word/document.xml' ||
      !HIDDEN_TEXT_DIAGNOSTIC_PART_PATTERN.test(path)
    ) {
      continue;
    }
    parts.push(await archive.xml(path));
  }

  let hiddenCount = 0;
  let visibleResetCount = 0;
  let invalidCount = 0;
  for (const part of parts) {
    const propertySets = descendants(part, 'rPr').filter((element) =>
      DOCX_WORDPROCESSING_NAMESPACES.has(element.namespaceURI ?? ''),
    );
    for (const properties of propertySets) {
      if (!directChildren(properties, 'vanish').length) continue;
      const inspection = inspectDocxHiddenText(properties);
      if (inspection.status === 'valid') {
        if (inspection.value) hiddenCount += 1;
        else visibleResetCount += 1;
      } else if (inspection.status === 'invalid') {
        invalidCount += 1;
      }
    }
    invalidCount += descendants(part, 'vanish').filter((element) => {
      const parent = element.parentElement;
      return (
        !parent ||
        parent.localName !== 'rPr' ||
        !DOCX_WORDPROCESSING_NAMESPACES.has(parent.namespaceURI ?? '')
      );
    }).length;
  }

  const issues: WorkCompatibilityIssue[] = [];
  const validCount = hiddenCount + visibleResetCount;
  if (validCount) {
    issues.push({
      code: 'docx.hidden-text',
      severity: 'info',
      feature: 'Hidden text',
      message: `${validCount} native hidden-text value(s), including ${hiddenCount} hidden value(s) and ${visibleResetCount} explicit visible reset(s), were detected across document defaults, character and paragraph styles, body, headers, footers, notes, comments, and formatting revisions. Editable Writer stories preserve w:vanish through the Font dialog, the show-hidden-text editing view, and exact DOCX export; unchanged comment XML remains source-preserved rather than becoming a rich comment-editing claim. Preview and PDF output keep hidden text suppressed, while affected paragraphs use browser-authoritative line measurement.`,
    });
  }
  if (invalidCount) {
    issues.push({
      code: 'docx.hidden-text.invalid',
      severity: 'warning',
      feature: 'Hidden text',
      message: `${invalidCount} malformed, duplicated, misplaced, namespace-spoofed, text-bearing, child-bearing, extra-attribute, or unknown hidden-text property set(s) are ignored instead of hiding untrusted content.`,
    });
  }
  return issues;
}
