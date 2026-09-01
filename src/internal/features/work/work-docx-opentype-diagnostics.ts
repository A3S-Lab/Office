import { DOCX_WORDPROCESSING_NAMESPACES } from './work-docx-ignorable-extension-preservation';
import {
  DOCX_WORD_2010_NAMESPACE,
  resolveDocxOpenTypeFeatures,
} from './work-docx-opentype-import';
import {
  descendants,
  directChildren,
  type OoxmlPackage,
} from './work-ooxml-package';
import type { WorkCompatibilityIssue } from './work-types';

const OPEN_TYPE_DIAGNOSTIC_PART_PATTERN =
  /^word\/(?:document|styles|header\d*|footer\d*|footnotes|endnotes|comments)\.xml$/i;
const OPEN_TYPE_PROPERTY_NAMES = new Set([
  'ligatures',
  'numForm',
  'numSpacing',
  'stylisticSets',
  'cntxtAlts',
]);

export async function diagnoseDocxOpenTypeTypography(
  archive: OoxmlPackage,
  document: Document,
): Promise<WorkCompatibilityIssue[]> {
  const parts: Array<{ path: string; document: Document }> = [
    { path: 'word/document.xml', document },
  ];
  for (const path of archive.paths('word/')) {
    if (
      path.toLowerCase() === 'word/document.xml' ||
      !OPEN_TYPE_DIAGNOSTIC_PART_PATTERN.test(path)
    ) {
      continue;
    }
    parts.push({ path, document: await archive.xml(path) });
  }

  let validCount = 0;
  let invalidCount = 0;
  let spoofedCount = 0;
  for (const part of parts) {
    const propertySets = descendants(part.document, 'rPr').filter((element) =>
      DOCX_WORDPROCESSING_NAMESPACES.has(element.namespaceURI ?? ''),
    );
    for (const properties of propertySets) {
      const candidates = directChildren(properties).filter((child) =>
        OPEN_TYPE_PROPERTY_NAMES.has(child.localName),
      );
      if (!candidates.length) continue;
      const inspection = resolveDocxOpenTypeFeatures([properties]);
      if (inspection.features) validCount += 1;
      invalidCount += inspection.invalidCount;
      spoofedCount += inspection.spoofedCount;
    }
    for (const name of OPEN_TYPE_PROPERTY_NAMES) {
      for (const property of descendants(part.document, name)) {
        const parent = property.parentElement;
        if (
          parent?.localName === 'rPr' &&
          DOCX_WORDPROCESSING_NAMESPACES.has(parent.namespaceURI ?? '')
        ) {
          continue;
        }
        if (property.namespaceURI === DOCX_WORD_2010_NAMESPACE) {
          invalidCount += 1;
        } else {
          spoofedCount += 1;
        }
      }
    }
  }

  const issues: WorkCompatibilityIssue[] = [];
  if (validCount) {
    issues.push({
      code: 'docx.typography.opentype',
      severity: 'info',
      feature: 'OpenType typography',
      message: `${validCount} native OpenType run-property set(s) preserve all 16 w14:ligatures combinations, default, lining, or old-style number forms, default, proportional, or tabular number spacing, style sets 1-20, and contextual alternates through inheritance, body and static-story editing, formatting revisions, CSS projection, and DOCX export.`,
    });
  }
  if (invalidCount || spoofedCount) {
    issues.push({
      code: 'docx.typography.opentype.invalid',
      severity: 'warning',
      feature: 'OpenType typography',
      message: `${invalidCount + spoofedCount} malformed, duplicated, misplaced, namespace-spoofed, text-bearing, child-bearing, extra-attribute, unknown, or out-of-range OpenType run-property value(s) are ignored instead of enabling untrusted typography.`,
    });
  }
  return issues;
}
