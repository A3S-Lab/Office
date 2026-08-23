import { DOCX_WORDPROCESSING_NAMESPACES } from './work-docx-ignorable-extension-preservation';
import { inspectDocxKerningThresholdHalfPoints } from './work-docx-kerning';
import {
  descendants,
  directChildren,
  type OoxmlPackage,
} from './work-ooxml-package';
import type { WorkCompatibilityIssue } from './work-types';

const KERNING_DIAGNOSTIC_PART_PATTERN =
  /^word\/(?:document|styles|header\d*|footer\d*|footnotes|endnotes)\.xml$/i;

export async function diagnoseDocxKerning(
  archive: OoxmlPackage,
  document: Document,
): Promise<WorkCompatibilityIssue[]> {
  const parts: Array<{ path: string; document: Document }> = [
    { path: 'word/document.xml', document },
  ];
  for (const path of archive.paths('word/')) {
    if (
      path.toLowerCase() === 'word/document.xml' ||
      !KERNING_DIAGNOSTIC_PART_PATTERN.test(path)
    ) {
      continue;
    }
    parts.push({ path, document: await archive.xml(path) });
  }

  let validCount = 0;
  let explicitZeroCount = 0;
  let invalidCount = 0;
  for (const part of parts) {
    const propertySets = descendants(part.document, 'rPr').filter((element) =>
      DOCX_WORDPROCESSING_NAMESPACES.has(element.namespaceURI ?? ''),
    );
    for (const properties of propertySets) {
      if (!directChildren(properties, 'kern').length) continue;
      const inspection = inspectDocxKerningThresholdHalfPoints(properties);
      if (inspection.status === 'valid') {
        validCount += 1;
        if (inspection.value === 0) explicitZeroCount += 1;
      } else if (inspection.status === 'invalid') {
        invalidCount += 1;
      }
    }
    invalidCount += descendants(part.document, 'kern').filter((element) => {
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
      code: 'docx.kerning',
      severity: 'info',
      feature: 'Character kerning',
      message: `${validCount} native kerning threshold(s), including ${explicitZeroCount} explicit all-font-size value(s), preserve exact half-points through style inheritance, body and page-chrome editing, effective font-size rendering, formatting revisions, and DOCX export. Pair kerning applies when the effective w:sz value meets or exceeds w:kern; absence throughout the hierarchy keeps kerning disabled.`,
    });
  }
  if (invalidCount) {
    issues.push({
      code: 'docx.kerning.invalid',
      severity: 'warning',
      feature: 'Character kerning',
      message: `${invalidCount} malformed, duplicated, misplaced, namespace-spoofed, text-bearing, or out-of-range kerning property set(s) are ignored instead of enabling untrusted typography.`,
    });
  }
  return issues;
}
