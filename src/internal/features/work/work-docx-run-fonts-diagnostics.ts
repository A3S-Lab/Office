import { DOCX_WORDPROCESSING_NAMESPACES } from './work-docx-ignorable-extension-preservation';
import { inspectDocxRunFonts } from './work-docx-run-fonts';
import {
  descendants,
  directChildren,
  type OoxmlPackage,
} from './work-ooxml-package';
import type { WorkCompatibilityIssue } from './work-types';

const RUN_FONT_DIAGNOSTIC_PART_PATTERN =
  /^word\/(?:document|styles|header\d*|footer\d*|footnotes|endnotes|comments)\.xml$/i;

export async function diagnoseDocxRunFonts(
  archive: OoxmlPackage,
  document: Document,
): Promise<WorkCompatibilityIssue[]> {
  const parts: Array<{ path: string; document: Document }> = [
    { path: 'word/document.xml', document },
  ];
  for (const path of archive.paths('word/')) {
    if (
      path.toLowerCase() === 'word/document.xml' ||
      !RUN_FONT_DIAGNOSTIC_PART_PATTERN.test(path)
    ) {
      continue;
    }
    parts.push({ path, document: await archive.xml(path) });
  }

  let validCount = 0;
  let directSlotCount = 0;
  let themeSlotCount = 0;
  let hintCount = 0;
  let invalidCount = 0;
  for (const part of parts) {
    const propertySets = descendants(part.document, 'rPr').filter((element) =>
      DOCX_WORDPROCESSING_NAMESPACES.has(element.namespaceURI ?? ''),
    );
    for (const properties of propertySets) {
      if (!directChildren(properties, 'rFonts').length) continue;
      const inspection = inspectDocxRunFonts(properties);
      if (inspection.status === 'valid') {
        validCount += 1;
        if (!inspection.value) continue;
        for (const slot of [
          inspection.value.ascii,
          inspection.value.highAnsi,
          inspection.value.eastAsia,
          inspection.value.complexScript,
        ]) {
          if (slot?.name) directSlotCount += 1;
          if (slot?.theme) themeSlotCount += 1;
        }
        if (inspection.value.hint) hintCount += 1;
      } else if (inspection.status === 'invalid') {
        invalidCount += 1;
      }
    }
    invalidCount += descendants(part.document, 'rFonts').filter((element) => {
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
      code: 'docx.script-fonts',
      severity: 'info',
      feature: 'Script-specific run fonts',
      message: `${validCount} native run-font property set(s), containing ${directSlotCount} direct slot value(s), ${themeSlotCount} theme slot reference(s), and ${hintCount} script hint(s), preserve the independent ASCII, high ANSI, East Asian, and complex-script identities through style inheritance, mixed-script browser rendering, body and page-chrome editing, formatting revisions, and DOCX export. Browser-resolved families remain separate from native source identity so font substitution does not rewrite untouched theme references.`,
    });
  }
  if (invalidCount) {
    issues.push({
      code: 'docx.script-fonts.invalid',
      severity: 'warning',
      feature: 'Script-specific run fonts',
      message: `${invalidCount} malformed, duplicated, misplaced, namespace-spoofed, child-bearing, text-bearing, unknown, or oversized run-font property set(s) are ignored instead of applying untrusted font metadata.`,
    });
  }
  return issues;
}
