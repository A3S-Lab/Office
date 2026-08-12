import {
  inspectDocxEquation,
  isSupportedDocxEquationPlacement,
} from './work-docx-equation-import';
import type { OoxmlPackage } from './work-ooxml-package';
import type { WorkCompatibilityIssue } from './work-types';

const EQUATION_PART_PATTERN =
  /^word\/(?:header\d*|footer\d*|footnotes|endnotes)\.xml$/i;
const MAX_EDITABLE_EQUATIONS = 4_096;

export async function diagnoseDocxEquations(
  archive: OoxmlPackage,
  document: Document | null,
): Promise<WorkCompatibilityIssue[]> {
  const parts: Array<{ path: string; document: Document }> = document
    ? [{ path: 'word/document.xml', document }]
    : [];
  for (const path of archive
    .paths('word/')
    .filter((candidate) => EQUATION_PART_PATTERN.test(candidate))) {
    parts.push({ path, document: await archive.xml(path) });
  }
  let supported = 0;
  let unsupported = 0;
  let spoofed = 0;
  for (const part of parts) {
    for (const element of Array.from(
      part.document.querySelectorAll('*'),
    ).filter(
      (candidate) =>
        candidate.localName === 'oMath' || candidate.localName === 'oMathPara',
    )) {
      if (closestEquationRoot(element.parentElement)) continue;
      const inspection = inspectDocxEquation(element);
      if (
        inspection.status === 'supported' &&
        isSupportedDocxEquationPlacement(element)
      ) {
        supported += 1;
      } else if (inspection.status === 'spoofed') {
        spoofed += 1;
      } else {
        unsupported += 1;
      }
    }
  }
  const total = supported + unsupported + spoofed;
  if (!total) return [];
  const issues: WorkCompatibilityIssue[] = [];
  if (supported) {
    issues.push(
      equationIssue(
        'docx.equations',
        `${supported} native OMML equation(s) use the editable structured subset. Inline and display equations, Unicode math runs, fractions, scripts, radicals, functions, supported n-ary operators, combining accents, overbars and underbars, bounded rectangular matrices, custom delimiters, and their accessible MathML previews round-trip through body text, headers, footers, footnotes, and endnotes.`,
        'info',
      ),
    );
  }
  if (unsupported || spoofed) {
    issues.push(
      equationIssue(
        'docx.equations.unsupported',
        `${unsupported + spoofed} equation-like object(s) use unsupported, malformed, misplaced, namespace-spoofed, relationship-bound, or ambiguous markup. They are flattened to bounded text instead of being treated as trusted structured OMML. Invalid or non-combining accent characters, invalid or contradictory bar properties, matrix spacing and gap rules, ragged or over-limit matrices, boxes, group characters, equation arrays, limits, phantoms, pre-scripts, and nested Office Math remain outside the structured subset.`,
      ),
    );
  }
  if (total > MAX_EDITABLE_EQUATIONS) {
    issues.push(
      equationIssue(
        'docx.equations.limit',
        `Only the first ${MAX_EDITABLE_EQUATIONS} valid equations enter the editable structured model; later equations are flattened to bounded text.`,
      ),
    );
  }
  return issues;
}

function closestEquationRoot(element: Element | null): Element | null {
  let current = element;
  while (current) {
    if (current.localName === 'oMath' || current.localName === 'oMathPara') {
      return current;
    }
    current = current.parentElement;
  }
  return null;
}

function equationIssue(
  code: string,
  message: string,
  severity: WorkCompatibilityIssue['severity'] = 'warning',
): WorkCompatibilityIssue {
  return { code, feature: 'Equations', message, severity };
}
