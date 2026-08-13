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
        `${supported} native OMML equation(s) use the editable structured subset. Inline equations and display equations with left, right, center, or center-group paragraph justification, Unicode math runs with literal/normal-text semantics, six math scripts, four styles, manual breaks, alignment points, ordered core Word run formatting for fonts/theme references, emphasis, casing, strike, relief effects, legacy text animations, bounded line borders, proofing/grid flags, visibility, colors, bounded character spacing, horizontal scaling, kerning thresholds, signed baseline positions, half-point sizes, named highlights, patterned shading, bounded manual run widths with optional grouping IDs, explicit baseline/superscript/subscript run alignment, all five Word emphasis-mark values, underlines, script direction, languages, native-only East Asian typography with signed run IDs, two-lines-in-one brackets, rotation, and compression, inert paragraph-mark visibility resets, and Office 2010 text glow, shadow, reflection, text-outline, text-fill, 3D-scene, and 3D-property effects with bounded radii, offsets, outline widths, extrusion heights, and contour widths, exact direction and reflection opacity/position/fade geometry, signed scale and skew, rectangle and pen alignment, glow/shadow/outline/fill/extrusion/contour RGB or theme colors and ordered transform chains, none/solid/gradient outline and text fills with bounded stops and linear/path shading, preset dashes, round/bevel/miter joins, all camera, light-rig, bevel, and material presets, all light directions, exact optional rotation angles, ordered top/bottom bevel plus extrusion/contour color branches, and all 16 ligature combinations and three numeral forms with exact CSS OpenType projection, with exact direct RGB and default-black text fills projected to MathML while complex effects remain native metadata, and the same bounded Word formatting in ordered object-level control properties and native-only argument-slot control properties, bounded insertion, deletion, move-source, and move-destination provenance around those math controls, ordered fractions with four native types and a bar default, ordered right-side scripts with aligned sub/superscripts, left-side pre-sub/superscripts with empty script slots, radicals with optional degrees and canonical square-root slots, functions with required, optionally empty name/argument slots, canonical empty arguments across every supported construct, ordered argument/control properties, bounded relative argument sizes with Word-effective MathML script-level projection and native-only preservation elsewhere, and strictly aligned matrix-cell, equation-array-row, and delimiter-argument metadata, ordered n-ary operators with default integrals, canonical empty limit slots, and optional operand growth, combining accents, overbars and underbars, group characters with explicit character position and baseline justification, phantoms with visible or hidden bases, zeroed width/ascent/descent, and transparent spacing, border boxes with independently visible edges and four strike directions, semantic boxes with operator-emulation, no-break, differential-spacing, manual-break, and alignment properties, bounded rectangular matrices with ordered row-spacing, column-gap, and minimum-width controls, equation arrays with bounded rows, vertical base alignment, distribution, row-spacing, and alignment/spacer markers, lower and upper limit objects, ordered delimiters with distinct omitted/empty characters, bounded empty argument slots, growing or fixed sizing, and centered or content-matched shape, and their accessible MathML previews round-trip through body text, headers, footers, footnotes, and endnotes.`,
        'info',
      ),
    );
  }
  if (unsupported || spoofed) {
    issues.push(
      equationIssue(
        'docx.equations.unsupported',
        `${unsupported + spoofed} equation-like object(s) use unsupported, malformed, misplaced, namespace-spoofed, relationship-bound, or ambiguous markup. They are flattened to bounded text instead of being treated as trusted structured OMML. Invalid or non-combining accent characters, malformed math-run, script, function, argument-property, or math-paragraph structures, unknown, duplicated, reordered, spoofed, out-of-range, or contradictory Word math-run casing, strike, relief, animation, line-border, visibility, geometry, manual-width, vertical-alignment, emphasis-mark, East Asian typography, paragraph-mark visibility, Office 2010 glow/shadow/outline/text-fill color transforms, glow/shadow/reflection geometry, outline fill/gradient/dash/join structure, text-fill wrapper/fill/gradient structure, 3D-scene camera/light structure, or 3D-property extrusion/contour/bevel/color structure, missing, malformed, or non-leaf Office 2010 ligature or numeral-form values, coordinate values, preset values, directions, or rotation angles, highlight, shading, object-control, or argument-control formatting, malformed control-revision identities, dates, children, or nesting, out-of-range or malformed argument sizes, invalid or contradictory fraction, radical, n-ary, delimiter, bar, group-character, phantom, border-box, box, or equation-array properties, malformed pre-script or lower/upper limit structures, malformed, duplicated, reordered, or out-of-range matrix spacing/gap properties, ragged or over-limit matrices, over-limit equation arrays, and nested Office Math remain outside the structured subset.`,
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
