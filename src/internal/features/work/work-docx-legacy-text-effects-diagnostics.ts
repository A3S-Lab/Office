import { DOCUMENT_LEGACY_TEXT_EFFECT_NAMES } from './work-document-legacy-text-effects';
import { DOCX_WORDPROCESSING_NAMESPACES } from './work-docx-ignorable-extension-preservation';
import { inspectDocxLegacyTextEffects } from './work-docx-legacy-text-effects';
import {
  descendants,
  directChildren,
  type OoxmlPackage,
} from './work-ooxml-package';
import type { WorkCompatibilityIssue } from './work-types';

const LEGACY_TEXT_EFFECTS_DIAGNOSTIC_PART_PATTERN =
  /^word\/(?:document|styles|header\d*|footer\d*|footnotes|endnotes|comments)\.xml$/i;

export async function diagnoseDocxLegacyTextEffects(
  archive: OoxmlPackage,
  document: Document,
): Promise<WorkCompatibilityIssue[]> {
  const parts: Document[] = [document];
  for (const path of archive.paths('word/')) {
    if (
      path.toLowerCase() === 'word/document.xml' ||
      !LEGACY_TEXT_EFFECTS_DIAGNOSTIC_PART_PATTERN.test(path)
    ) {
      continue;
    }
    parts.push(await archive.xml(path));
  }

  const enabled = { outline: 0, shadow: 0, emboss: 0, imprint: 0 };
  let validCount = 0;
  let explicitFalseCount = 0;
  let invalidCount = 0;
  for (const part of parts) {
    const propertySets = descendants(part, 'rPr').filter((element) =>
      DOCX_WORDPROCESSING_NAMESPACES.has(element.namespaceURI ?? ''),
    );
    for (const properties of propertySets) {
      if (
        !DOCUMENT_LEGACY_TEXT_EFFECT_NAMES.some(
          (name) => directChildren(properties, name).length > 0,
        )
      ) {
        continue;
      }
      const inspection = inspectDocxLegacyTextEffects(properties);
      if (inspection.status === 'valid') {
        for (const name of DOCUMENT_LEGACY_TEXT_EFFECT_NAMES) {
          const value = inspection.value[name];
          if (value === undefined) continue;
          validCount += 1;
          if (value) enabled[name] += 1;
          else explicitFalseCount += 1;
        }
      } else if (inspection.status === 'invalid') {
        invalidCount += 1;
      }
    }
    for (const name of DOCUMENT_LEGACY_TEXT_EFFECT_NAMES) {
      invalidCount += descendants(part, name).filter((element) => {
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
  if (validCount) {
    issues.push({
      code: 'docx.legacy-text-effects',
      severity: 'info',
      feature: 'Legacy text effects',
      message: `${validCount} native legacy text-effect value(s), including ${explicitFalseCount} explicit false reset(s), ${enabled.outline} enabled outline value(s), ${enabled.shadow} enabled shadow value(s), ${enabled.emboss} enabled emboss value(s), and ${enabled.imprint} enabled imprint value(s), were detected across document defaults, styles, body, page chrome, notes, comments, and formatting revisions. Editable Writer stories preserve w:outline, w:shadow, w:emboss, and w:imprint through independently mixed Font-dialog controls, conflict-aware one-step Undo, Format Painter, bounded CSS paint projections, and exact DOCX export. Outline plus shadow remains valid; emboss and imprint stay mutually exclusive with the other enabled legacy effects. The browser and PDF projections are bounded visual approximations rather than a claim of desktop-engine pixel identity, and no non-standard shortcut is invented.`,
    });
  }
  if (invalidCount) {
    issues.push({
      code: 'docx.legacy-text-effects.invalid',
      severity: 'warning',
      feature: 'Legacy text effects',
      message: `${invalidCount} malformed, duplicated, misplaced, namespace-spoofed, text-bearing, child-bearing, extra-attribute, unknown, or conflicting legacy text-effect property set(s) are ignored instead of enabling contradictory or untrusted paint effects.`,
    });
  }
  return issues;
}
