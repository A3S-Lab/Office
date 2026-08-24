import { DOCX_WORDPROCESSING_NAMESPACES } from './work-docx-ignorable-extension-preservation';
import {
  inspectDocxNoProof,
  inspectDocxProofingLanguages,
} from './work-docx-proofing';
import {
  descendants,
  directChildren,
  type OoxmlPackage,
} from './work-ooxml-package';
import type { WorkCompatibilityIssue } from './work-types';

const PROOFING_PART_PATTERN =
  /^word\/(?:document|styles|header\d*|footer\d*|footnotes|endnotes|comments)\.xml$/i;

export async function diagnoseDocxProofing(
  archive: OoxmlPackage,
  document: Document,
): Promise<WorkCompatibilityIssue[]> {
  const parts: Document[] = [document];
  for (const path of archive.paths('word/')) {
    if (
      path.toLowerCase() !== 'word/document.xml' &&
      PROOFING_PART_PATTERN.test(path)
    ) {
      parts.push(await archive.xml(path));
    }
  }

  let languageCount = 0;
  let proofingStateCount = 0;
  let excludedCount = 0;
  let checkedCount = 0;
  let invalidCount = 0;
  for (const part of parts) {
    for (const properties of descendants(part, 'rPr').filter((element) =>
      DOCX_WORDPROCESSING_NAMESPACES.has(element.namespaceURI ?? ''),
    )) {
      if (directChildren(properties, 'lang').length) {
        const inspection = inspectDocxProofingLanguages(properties);
        if (inspection.status === 'valid') languageCount += 1;
        else invalidCount += 1 + inspection.spoofedCount;
      }
      if (directChildren(properties, 'noProof').length) {
        const inspection = inspectDocxNoProof(properties);
        if (inspection.status === 'valid') {
          proofingStateCount += 1;
          if (inspection.value) excludedCount += 1;
          else checkedCount += 1;
        } else {
          invalidCount += 1 + inspection.spoofedCount;
        }
      }
    }
    for (const name of ['lang', 'noProof'] as const) {
      invalidCount += descendants(part, name).filter((element) => {
        if (!DOCX_WORDPROCESSING_NAMESPACES.has(element.namespaceURI ?? '')) {
          return false;
        }
        const parent = element.parentElement;
        return (
          !parent ||
          parent.localName !== 'rPr' ||
          !DOCX_WORDPROCESSING_NAMESPACES.has(parent.namespaceURI ?? '')
        );
      }).length;
    }
  }

  const issues: WorkCompatibilityIssue[] = [];
  if (languageCount || proofingStateCount) {
    issues.push({
      code: 'docx.proofing-language',
      severity: 'info',
      feature: 'Proofing languages',
      message: `${languageCount} native language declaration(s) and ${proofingStateCount} explicit proofing state(s), including ${excludedCount} excluded and ${checkedCount} explicitly checked state(s), were detected across document defaults, styles, body, page chrome, notes, comments, and formatting revisions. Writer preserves independent Latin, East Asian, and bidi language slots plus native w:noProof true or false semantics through mixed-selection editing, one-step Undo, tracked formatting revisions, and exact DOCX export.`,
    });
  }
  if (invalidCount) {
    issues.push({
      code: 'docx.proofing-language.invalid',
      severity: 'warning',
      feature: 'Proofing languages',
      message: `${invalidCount} malformed, duplicated, misplaced, namespace-spoofed, text-bearing, child-bearing, extra-attribute, invalid-language-tag, or invalid on/off proofing value(s) are ignored instead of enabling untrusted language metadata or inheriting stale proofing state.`,
    });
  }
  return issues;
}
