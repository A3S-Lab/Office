import { DOCX_WORDPROCESSING_NAMESPACES } from './work-docx-ignorable-extension-preservation';
import {
  inspectDocxPageMargins,
  inspectDocxPageMarginSettings,
} from './work-docx-page-margins-import';
import { attribute, descendants, directChild } from './work-ooxml-package';
import type { OoxmlPackage } from './work-ooxml-package';
import type { WorkCompatibilityIssue } from './work-types';

export async function diagnoseDocxPageMargins(
  archive: OoxmlPackage,
  document: Document,
): Promise<WorkCompatibilityIssue[]> {
  const settingsDocument = archive.has('word/settings.xml')
    ? await archive.xml('word/settings.xml')
    : null;
  const settings = inspectDocxPageMarginSettings(settingsDocument);
  const sections = descendants(document, 'sectPr').filter(
    (element) =>
      DOCX_WORDPROCESSING_NAMESPACES.has(element.namespaceURI ?? '') &&
      !hasAncestor(element, 'sectPrChange'),
  );
  const inspected = sections.map((section) =>
    inspectDocxPageMargins(section, settings),
  );
  const valid = inspected.flatMap((item, index) =>
    item.status === 'valid' && item.pageMargins
      ? [{ section: sections[index], pageMargins: item.pageMargins }]
      : [],
  );
  const invalidCount =
    settings.invalidCount +
    inspected.reduce((count, item) => count + item.invalidCount, 0);
  const spoofedCount =
    settings.spoofedCount +
    inspected.reduce((count, item) => count + item.spoofedCount, 0);
  if (!valid.length && !invalidCount && !spoofedCount) return [];

  const issues: WorkCompatibilityIssue[] = [];
  if (valid.length) {
    const gutterSections = valid.filter(
      ({ pageMargins }) => pageMargins.gutter > 0,
    ).length;
    const overlapSections = valid.filter(
      ({ pageMargins }) => pageMargins.top < 0 || pageMargins.bottom < 0,
    ).length;
    issues.push({
      code: 'docx.page-margins',
      severity: 'info',
      feature: 'Page margins',
      message: `${valid.length} section page-margin set(s) preserve all seven native pgMar values in exact twips, including header and footer distances, ${gutterSections} binding gutter(s), ${overlapSections} signed header/footer-overlap geometry set(s), document-wide mirror/top-gutter settings, per-section right-side gutters, strict namespaces, editing controls, paginated preview, and exact DOCX output.`,
    });
  }
  if (invalidCount || spoofedCount) {
    issues.push({
      code: 'docx.page-margins.invalid',
      severity: 'warning',
      feature: 'Page margins',
      message: `${invalidCount + spoofedCount} malformed, duplicated, incomplete, out-of-range, namespace-spoofed, relationship-bound, or structurally invalid pgMar, mirrorMargins, gutterAtTop, or rtlGutter value(s) are ignored instead of becoming trusted page geometry.`,
    });
  }
  const incompatible = settings.incompatible;
  if (
    valid.some(({ pageMargins }) => pageMargins.gutterAtTop) &&
    incompatible.length
  ) {
    issues.push({
      code: 'docx.page-margins.gutter-conflict',
      severity: 'warning',
      feature: 'Page-margin compatibility settings',
      message: `${incompatible.join(', ')} conflicts with gutterAtTop under WordprocessingML. Native values remain explicit, but the browser applies the facing-page or print-layout setting before the top-gutter request.`,
    });
  }
  if (
    valid.some(
      ({ pageMargins }) =>
        pageMargins.mirrorMargins &&
        (pageMargins.left !== pageMargins.right || pageMargins.gutter > 0),
    )
  ) {
    issues.push({
      code: 'docx.page-margins.facing-pages',
      severity: 'warning',
      feature: 'Facing-page margins',
      message:
        'PDF pages and page decorations swap facing-page margins and the binding gutter per physical page. The continuous live editing surface keeps the same text width and native geometry but projects the first-page horizontal origin between page breaks.',
    });
  }
  const bounded = valid.filter(({ section, pageMargins }) => {
    const size = directChild(section, 'pgSz');
    const width = finitePositive(size ? attribute(size, 'w') : null) ?? 11_906;
    const height = finitePositive(size ? attribute(size, 'h') : null) ?? 16_838;
    const topGutter =
      pageMargins.gutterAtTop === true && pageMargins.mirrorMargins !== true;
    return (
      pageMargins.left +
        pageMargins.right +
        (topGutter ? 0 : pageMargins.gutter) >=
        width ||
      Math.abs(pageMargins.top) +
        Math.abs(pageMargins.bottom) +
        (topGutter ? pageMargins.gutter : 0) >=
        height
    );
  }).length;
  if (bounded) {
    issues.push({
      code: 'docx.page-margins.body-bounds',
      severity: 'warning',
      feature: 'Page-margin body bounds',
      message: `${bounded} section(s) leave no positive body rectangle. Exact native values remain authoritative for DOCX output; browser rendering proportionally bounds the affected margin pair to retain a one-millimetre inspectable body surface.`,
    });
  }
  return issues;
}

function finitePositive(value: string | null): number | null {
  const number = Number(value);
  return Number.isFinite(number) && number > 0 ? number : null;
}

function hasAncestor(element: Element, localName: string): boolean {
  let ancestor = element.parentElement;
  while (ancestor) {
    if (ancestor.localName === localName) return true;
    ancestor = ancestor.parentElement;
  }
  return false;
}
