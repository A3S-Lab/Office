import { inspectDocxEmphasisMark } from './work-docx-emphasis';
import { DOCX_WORDPROCESSING_NAMESPACES } from './work-docx-ignorable-extension-preservation';
import {
  descendants,
  directChildren,
  type OoxmlPackage,
} from './work-ooxml-package';
import type { WorkCompatibilityIssue } from './work-types';

const EMPHASIS_DIAGNOSTIC_PART_PATTERN =
  /^word\/(?:document|styles|header\d*|footer\d*|footnotes|endnotes)\.xml$/i;

export async function diagnoseDocxEmphasisMarks(
  archive: OoxmlPackage,
  document: Document,
): Promise<WorkCompatibilityIssue[]> {
  const parts: Array<{ path: string; document: Document }> = [
    { path: 'word/document.xml', document },
  ];
  for (const path of archive.paths('word/')) {
    if (
      path.toLowerCase() === 'word/document.xml' ||
      !EMPHASIS_DIAGNOSTIC_PART_PATTERN.test(path)
    ) {
      continue;
    }
    parts.push({ path, document: await archive.xml(path) });
  }

  let validCount = 0;
  let explicitNoneCount = 0;
  let underDotCount = 0;
  let invalidCount = 0;
  for (const part of parts) {
    const propertySets = descendants(part.document, 'rPr').filter((element) =>
      DOCX_WORDPROCESSING_NAMESPACES.has(element.namespaceURI ?? ''),
    );
    for (const properties of propertySets) {
      if (!directChildren(properties, 'em').length) continue;
      const inspection = inspectDocxEmphasisMark(properties);
      if (inspection.status === 'valid') {
        validCount += 1;
        if (inspection.value === 'none') explicitNoneCount += 1;
        if (inspection.value === 'underDot') underDotCount += 1;
      } else if (inspection.status === 'invalid') {
        invalidCount += 1;
      }
    }
    invalidCount += descendants(part.document, 'em').filter((element) => {
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
      code: 'docx.emphasis-mark',
      severity: 'info',
      feature: 'East Asian emphasis marks',
      message: `${validCount} native emphasis mark(s), including ${explicitNoneCount} explicit none reset(s) and ${underDotCount} below-text dot(s), preserve all five w:em values through style inheritance, body and page-chrome editing, formatting revisions, CSS projection, and DOCX export. Runs with visible marks use browser line measurement so their out-of-line glyph extents are not approximated by the Worker/WASM shaper.`,
    });
  }
  if (invalidCount) {
    issues.push({
      code: 'docx.emphasis-mark.invalid',
      severity: 'warning',
      feature: 'East Asian emphasis marks',
      message: `${invalidCount} malformed, duplicated, misplaced, namespace-spoofed, text-bearing, or unknown emphasis property set(s) are ignored instead of enabling untrusted typography.`,
    });
  }
  return issues;
}
