import JSZip from 'jszip';
import {
  documentEquationFromElement,
  normalizeDocumentEquation,
  type WorkDocumentEquation,
  type WorkDocumentEquationArgumentProperties,
  type WorkDocumentEquationControlRevision,
  type WorkDocumentEquationExpression,
  type WorkDocumentEquationJustification,
  type WorkDocumentEquationRunScript,
  type WorkDocumentEquationRunStyle,
  type WorkDocumentEquationSpacingRule,
  type WorkDocumentEquationWordBevel,
  type WorkDocumentEquationWordBevelPreset,
  type WorkDocumentEquationWordColor,
  type WorkDocumentEquationWordColorTransformType,
  type WorkDocumentEquationWordEffectColor,
  type WorkDocumentEquationWordEffectFill,
  type WorkDocumentEquationWordLightRigDirection,
  type WorkDocumentEquationWordLightRigPreset,
  type WorkDocumentEquationWordPresetLineDash,
  type WorkDocumentEquationWordPresetMaterial,
  type WorkDocumentEquationWordRectangleAlignment,
  type WorkDocumentEquationWordRunProperties,
  type WorkDocumentEquationWordTextOutlineAlignment,
  type WorkDocumentEquationWordTextOutlineCap,
  type WorkDocumentEquationWordTextOutlineCompound,
} from './work-document-equations';
import { DOCX_WORDPROCESSING_NAMESPACES } from './work-docx-ignorable-extension-preservation';
import {
  XMLNS_NAMESPACE,
  xmlAttributeLocalName,
  xmlAttributeNamespace,
  xmlDeclaredPrefix,
  xmlNamespaceUri,
} from './work-docx-settings-xml';
import { descendants, directChildren, parseXml } from './work-ooxml-package';
import { decodeXmlBytes, serializeUtf8Xml } from './work-ooxml-xml';

interface DocxEquationPatch {
  marker: string;
  equation: WorkDocumentEquation;
}

const MATH_NAMESPACE =
  'http://schemas.openxmlformats.org/officeDocument/2006/math';
const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const WORD_DATE_UTC_NAMESPACE =
  'http://schemas.microsoft.com/office/word/2023/wordml/word16du';
const WORD_2010_NAMESPACE =
  'http://schemas.microsoft.com/office/word/2010/wordml';
const MARKUP_COMPATIBILITY_NAMESPACE =
  'http://schemas.openxmlformats.org/markup-compatibility/2006';
const XML_NAMESPACE = 'http://www.w3.org/XML/1998/namespace';
const EQUATION_PART_PATTERN =
  /^word\/(?:document|header\d*|footer\d*|footnotes|endnotes)\.xml$/i;
const MAX_EQUATION_PATCHES = 4_096;
const WORD_2010_SCHEME_COLORS = new Map([
  ['background1', 'bg1'],
  ['text1', 'tx1'],
  ['background2', 'bg2'],
  ['text2', 'tx2'],
  ['accent1', 'accent1'],
  ['accent2', 'accent2'],
  ['accent3', 'accent3'],
  ['accent4', 'accent4'],
  ['accent5', 'accent5'],
  ['accent6', 'accent6'],
  ['hyperlink', 'hlink'],
  ['followedHyperlink', 'folHlink'],
  ['dark1', 'dk1'],
  ['light1', 'lt1'],
  ['dark2', 'dk2'],
  ['light2', 'lt2'],
  ['placeholder', 'phClr'],
] as const);
const WORD_2010_COLOR_TRANSFORM_NAMES: Readonly<
  Record<WorkDocumentEquationWordColorTransformType, string>
> = {
  tint: 'tint',
  shade: 'shade',
  alpha: 'alpha',
  hueMod: 'hueMod',
  saturation: 'sat',
  saturationOffset: 'satOff',
  saturationModulation: 'satMod',
  luminance: 'lum',
  luminanceOffset: 'lumOff',
  luminanceModulation: 'lumMod',
};
const WORD_2010_RECTANGLE_ALIGNMENTS: Readonly<
  Record<WorkDocumentEquationWordRectangleAlignment, string>
> = {
  none: 'none',
  topLeft: 'tl',
  top: 't',
  topRight: 'tr',
  left: 'l',
  center: 'ctr',
  right: 'r',
  bottomLeft: 'bl',
  bottom: 'b',
  bottomRight: 'br',
};
const WORD_2010_TEXT_OUTLINE_CAPS: Readonly<
  Record<WorkDocumentEquationWordTextOutlineCap, string>
> = {
  round: 'rnd',
  square: 'sq',
  flat: 'flat',
};
const WORD_2010_TEXT_OUTLINE_COMPOUNDS: Readonly<
  Record<WorkDocumentEquationWordTextOutlineCompound, string>
> = {
  single: 'sng',
  double: 'dbl',
  thickThin: 'thickThin',
  thinThick: 'thinThick',
  triple: 'tri',
};
const WORD_2010_TEXT_OUTLINE_ALIGNMENTS: Readonly<
  Record<WorkDocumentEquationWordTextOutlineAlignment, string>
> = {
  center: 'ctr',
  inset: 'in',
};
const WORD_2010_PRESET_LINE_DASHES: Readonly<
  Record<WorkDocumentEquationWordPresetLineDash, string>
> = {
  solid: 'solid',
  dot: 'dot',
  systemDot: 'sysDot',
  dash: 'dash',
  systemDash: 'sysDash',
  longDash: 'lgDash',
  dashDot: 'dashDot',
  systemDashDot: 'sysDashDot',
  longDashDot: 'lgDashDot',
  longDashDotDot: 'lgDashDotDot',
  systemDashDotDot: 'sysDashDotDot',
};
const WORD_2010_SCENE_3D_LIGHT_RIG_PRESETS: Readonly<
  Record<WorkDocumentEquationWordLightRigPreset, string>
> = {
  legacyFlat1: 'legacyFlat1',
  legacyFlat2: 'legacyFlat2',
  legacyFlat3: 'legacyFlat3',
  legacyFlat4: 'legacyFlat4',
  legacyNormal1: 'legacyNormal1',
  legacyNormal2: 'legacyNormal2',
  legacyNormal3: 'legacyNormal3',
  legacyNormal4: 'legacyNormal4',
  legacyHarsh1: 'legacyHarsh1',
  legacyHarsh2: 'legacyHarsh2',
  legacyHarsh3: 'legacyHarsh3',
  legacyHarsh4: 'legacyHarsh4',
  threePoint: 'threePt',
  balanced: 'balanced',
  soft: 'soft',
  harsh: 'harsh',
  flood: 'flood',
  contrasting: 'contrasting',
  morning: 'morning',
  sunrise: 'sunrise',
  sunset: 'sunset',
  chilly: 'chilly',
  freezing: 'freezing',
  flat: 'flat',
  twoPoint: 'twoPt',
  glow: 'glow',
  brightRoom: 'brightRoom',
};
const WORD_2010_SCENE_3D_LIGHT_RIG_DIRECTIONS: Readonly<
  Record<WorkDocumentEquationWordLightRigDirection, string>
> = {
  topLeft: 'tl',
  top: 't',
  topRight: 'tr',
  left: 'l',
  right: 'r',
  bottomLeft: 'bl',
  bottom: 'b',
  bottomRight: 'br',
};
const WORD_2010_PROPERTIES_3D_BEVEL_PRESETS: Readonly<
  Record<WorkDocumentEquationWordBevelPreset, string>
> = {
  relaxedInset: 'relaxedInset',
  circle: 'circle',
  slope: 'slope',
  cross: 'cross',
  angle: 'angle',
  softRound: 'softRound',
  convex: 'convex',
  coolSlant: 'coolSlant',
  divot: 'divot',
  riblet: 'riblet',
  hardEdge: 'hardEdge',
  artDeco: 'artDeco',
};
const WORD_2010_PROPERTIES_3D_MATERIAL_PRESETS: Readonly<
  Record<WorkDocumentEquationWordPresetMaterial, string>
> = {
  legacyMatte: 'legacyMatte',
  legacyPlastic: 'legacyPlastic',
  legacyMetal: 'legacyMetal',
  legacyWireframe: 'legacyWireframe',
  matte: 'matte',
  plastic: 'plastic',
  metal: 'metal',
  warmMatte: 'warmMatte',
  translucentPowder: 'translucentPowder',
  powder: 'powder',
  darkEdge: 'dkEdge',
  softEdge: 'softEdge',
  clear: 'clear',
  flat: 'flat',
  softMetal: 'softmetal',
  none: 'none',
};
const WORD_ANGLE_UNITS_PER_DEGREE = 60_000;
const WORD_PERCENTAGE_UNITS_PER_PERCENT = 1_000;

export class DocxEquationPatchCollector {
  readonly patches: DocxEquationPatch[] = [];
  private nextMarker = 1;

  constructor(private readonly source: string) {}

  marker(element: HTMLElement): string | null {
    const equation = documentEquationFromElement(element);
    if (!equation) return null;
    if (this.patches.length >= MAX_EQUATION_PATCHES) {
      throw new Error('Document exceeds the equation limit.');
    }
    let marker = '';
    do {
      marker = `__A3S_EQUATION_${this.nextMarker}__`;
      this.nextMarker += 1;
    } while (this.source.includes(marker));
    this.patches.push({ marker, equation });
    return marker;
  }
}

export async function patchDocxEquations(
  buffer: ArrayBuffer,
  patches: readonly DocxEquationPatch[],
): Promise<ArrayBuffer> {
  if (!patches.length) return buffer;
  if (patches.length > MAX_EQUATION_PATCHES) {
    throw new Error('Document exceeds the equation limit.');
  }
  const archive = await JSZip.loadAsync(buffer);
  const byMarker = new Map(patches.map((patch) => [patch.marker, patch]));
  const applied = new Map<string, number>();
  const entries = Object.values(archive.files).filter(
    (entry) => !entry.dir && EQUATION_PART_PATTERN.test(entry.name),
  );
  for (const entry of entries) {
    const document = parseXml(
      decodeXmlBytes(
        await entry.async('uint8array'),
        `generated DOCX ${entry.name}`,
      ),
      `generated DOCX ${entry.name}`,
    );
    if (!supportedPartRoot(document.documentElement, entry.name)) continue;
    const targets = equationMarkerTargets(document, byMarker);
    if (!targets.length) continue;
    const prefix = ensureMathPrefix(document.documentElement);
    let changed = false;
    for (const target of targets) {
      const count = (applied.get(target.patch.marker) ?? 0) + 1;
      applied.set(target.patch.marker, count);
      if (count > 1) continue;
      const equation = normalizeDocumentEquation(target.patch.equation);
      if (!equation) continue;
      const math = createMathElement(document, prefix, 'oMath');
      appendExpressions(document, prefix, math, equation.children);
      if (
        equation.display === 'block' &&
        target.paragraph &&
        paragraphHasOnlyEquationRun(target.paragraph, target.run)
      ) {
        const paragraph = createMathElement(document, prefix, 'oMathPara');
        if (equation.justification) {
          const properties = createMathElement(document, prefix, 'oMathParaPr');
          properties.append(
            mathValueElement(
              document,
              prefix,
              'jc',
              justificationToOmml(equation.justification),
            ),
          );
          paragraph.append(properties);
        }
        paragraph.append(math);
        target.paragraph.replaceWith(paragraph);
      } else {
        target.run.replaceWith(math);
      }
      changed = true;
    }
    if (changed) archive.file(entry.name, serializeUtf8Xml(document));
  }
  const invalid = patches.filter(
    (patch) => (applied.get(patch.marker) ?? 0) !== 1,
  );
  if (invalid.length) {
    throw new Error(
      `DOCX equation markers were not emitted exactly once: ${invalid
        .map((patch) => patch.marker)
        .join(', ')}.`,
    );
  }
  return archive.generateAsync({ type: 'arraybuffer' });
}

function equationMarkerTargets(
  document: Document,
  patches: ReadonlyMap<string, DocxEquationPatch>,
): Array<{
  patch: DocxEquationPatch;
  run: Element;
  paragraph: Element | null;
}> {
  const targets: Array<{
    patch: DocxEquationPatch;
    run: Element;
    paragraph: Element | null;
  }> = [];
  for (const text of descendants(document, 't')) {
    if (!DOCX_WORDPROCESSING_NAMESPACES.has(text.namespaceURI ?? '')) continue;
    const patch = patches.get(text.textContent ?? '');
    const run = text.parentElement;
    if (
      !patch ||
      !run ||
      run.localName !== 'r' ||
      !DOCX_WORDPROCESSING_NAMESPACES.has(run.namespaceURI ?? '') ||
      !runHasOnlyMarkerText(run, text)
    ) {
      continue;
    }
    targets.push({
      patch,
      run,
      paragraph: closestWordAncestor(run.parentElement, 'p'),
    });
  }
  return targets;
}

function runHasOnlyMarkerText(run: Element, marker: Element): boolean {
  return directChildren(run).every(
    (child) =>
      child === marker ||
      (child.localName === 'rPr' &&
        DOCX_WORDPROCESSING_NAMESPACES.has(child.namespaceURI ?? '')),
  );
}

function paragraphHasOnlyEquationRun(
  paragraph: Element,
  equationRun: Element,
): boolean {
  return directChildren(paragraph).every(
    (child) =>
      child === equationRun ||
      (child.localName === 'pPr' &&
        DOCX_WORDPROCESSING_NAMESPACES.has(child.namespaceURI ?? '')),
  );
}

function closestWordAncestor(
  element: Element | null,
  name: string,
): Element | null {
  let current = element;
  while (current) {
    if (
      current.localName === name &&
      DOCX_WORDPROCESSING_NAMESPACES.has(current.namespaceURI ?? '')
    ) {
      return current;
    }
    current = current.parentElement;
  }
  return null;
}

function appendExpressions(
  document: Document,
  prefix: string,
  parent: Element,
  expressions: readonly WorkDocumentEquationExpression[],
): void {
  for (const expression of expressions) {
    parent.append(createExpression(document, prefix, expression));
  }
}

function createExpression(
  document: Document,
  prefix: string,
  expression: WorkDocumentEquationExpression,
): Element {
  if (expression.type === 'run') {
    const run = createMathElement(document, prefix, 'r');
    const properties = createMathElement(document, prefix, 'rPr');
    if (expression.literal) {
      properties.append(mathOnOffElement(document, prefix, 'lit', true));
    }
    if (expression.normalText) {
      properties.append(mathOnOffElement(document, prefix, 'nor', true));
    }
    if (expression.script) {
      properties.append(
        mathValueElement(
          document,
          prefix,
          'scr',
          runScriptToOmml(expression.script),
        ),
      );
    }
    if (expression.style) {
      properties.append(
        mathValueElement(
          document,
          prefix,
          'sty',
          runStyleToOmml(expression.style),
        ),
      );
    }
    if (expression.manualBreak) {
      const manualBreak = createMathElement(document, prefix, 'brk');
      if (expression.manualBreak.alignmentAt !== undefined) {
        manualBreak.setAttributeNS(
          MATH_NAMESPACE,
          `${prefix}:alnAt`,
          String(expression.manualBreak.alignmentAt),
        );
      }
      properties.append(manualBreak);
    }
    if (expression.alignment) {
      properties.append(mathOnOffElement(document, prefix, 'aln', true));
    }
    if (properties.childNodes.length) run.append(properties);
    if (expression.wordRunProperties) {
      run.append(
        createWordRunProperties(document, expression.wordRunProperties),
      );
    }
    const text = createMathElement(document, prefix, 't');
    if (/^\s|\s$/u.test(expression.text)) {
      text.setAttributeNS(XML_NAMESPACE, 'xml:space', 'preserve');
    }
    text.textContent = expression.text;
    run.append(text);
    return run;
  }
  if (expression.type === 'accent') {
    const accent = createMathElement(document, prefix, 'acc');
    const properties = createMathElement(document, prefix, 'accPr');
    properties.append(
      mathValueElement(document, prefix, 'chr', expression.character),
    );
    appendMathControlProperties(
      document,
      prefix,
      properties,
      expression.controlProperties,
      expression.controlRevision,
    );
    accent.append(
      properties,
      expressionArgument(
        document,
        prefix,
        'e',
        expression.children,
        expression.childrenProperties,
      ),
    );
    return accent;
  }
  if (expression.type === 'bar') {
    const bar = createMathElement(document, prefix, 'bar');
    const properties = createMathElement(document, prefix, 'barPr');
    properties.append(
      mathValueElement(
        document,
        prefix,
        'pos',
        expression.position === 'top' ? 'top' : 'bot',
      ),
    );
    appendMathControlProperties(
      document,
      prefix,
      properties,
      expression.controlProperties,
      expression.controlRevision,
    );
    bar.append(
      properties,
      expressionArgument(
        document,
        prefix,
        'e',
        expression.children,
        expression.childrenProperties,
      ),
    );
    return bar;
  }
  if (expression.type === 'groupCharacter') {
    const groupCharacter = createMathElement(document, prefix, 'groupChr');
    const properties = createMathElement(document, prefix, 'groupChrPr');
    properties.append(
      mathValueElement(document, prefix, 'chr', expression.character),
      mathValueElement(
        document,
        prefix,
        'pos',
        expression.position === 'top' ? 'top' : 'bot',
      ),
      mathValueElement(
        document,
        prefix,
        'vertJc',
        expression.verticalJustification === 'top' ? 'top' : 'bot',
      ),
    );
    appendMathControlProperties(
      document,
      prefix,
      properties,
      expression.controlProperties,
      expression.controlRevision,
    );
    groupCharacter.append(
      properties,
      expressionArgument(
        document,
        prefix,
        'e',
        expression.children,
        expression.childrenProperties,
      ),
    );
    return groupCharacter;
  }
  if (expression.type === 'phantom') {
    const phantom = createMathElement(document, prefix, 'phant');
    const properties = createMathElement(document, prefix, 'phantPr');
    properties.append(
      mathOnOffElement(document, prefix, 'show', expression.show),
      mathOnOffElement(document, prefix, 'zeroWid', expression.zeroWidth),
      mathOnOffElement(document, prefix, 'zeroAsc', expression.zeroAscent),
      mathOnOffElement(document, prefix, 'zeroDesc', expression.zeroDescent),
      mathOnOffElement(document, prefix, 'transp', expression.transparent),
    );
    appendMathControlProperties(
      document,
      prefix,
      properties,
      expression.controlProperties,
      expression.controlRevision,
    );
    phantom.append(
      properties,
      expressionArgument(
        document,
        prefix,
        'e',
        expression.children,
        expression.childrenProperties,
      ),
    );
    return phantom;
  }
  if (expression.type === 'borderBox') {
    const borderBox = createMathElement(document, prefix, 'borderBox');
    const properties = createMathElement(document, prefix, 'borderBoxPr');
    properties.append(
      mathOnOffElement(document, prefix, 'hideTop', expression.hideTop),
      mathOnOffElement(document, prefix, 'hideBot', expression.hideBottom),
      mathOnOffElement(document, prefix, 'hideLeft', expression.hideLeft),
      mathOnOffElement(document, prefix, 'hideRight', expression.hideRight),
      mathOnOffElement(
        document,
        prefix,
        'strikeH',
        expression.strikeHorizontal,
      ),
      mathOnOffElement(document, prefix, 'strikeV', expression.strikeVertical),
      mathOnOffElement(
        document,
        prefix,
        'strikeBLTR',
        expression.strikeBottomLeftToTopRight,
      ),
      mathOnOffElement(
        document,
        prefix,
        'strikeTLBR',
        expression.strikeTopLeftToBottomRight,
      ),
    );
    appendMathControlProperties(
      document,
      prefix,
      properties,
      expression.controlProperties,
      expression.controlRevision,
    );
    borderBox.append(
      properties,
      expressionArgument(
        document,
        prefix,
        'e',
        expression.children,
        expression.childrenProperties,
      ),
    );
    return borderBox;
  }
  if (expression.type === 'box') {
    const box = createMathElement(document, prefix, 'box');
    const properties = createMathElement(document, prefix, 'boxPr');
    properties.append(
      mathOnOffElement(document, prefix, 'opEmu', expression.operatorEmulator),
      mathOnOffElement(document, prefix, 'noBreak', expression.noBreak),
      mathOnOffElement(document, prefix, 'diff', expression.differential),
    );
    if (expression.manualBreak) {
      const manualBreak = createMathElement(document, prefix, 'brk');
      if (expression.manualBreak.alignmentAt !== undefined) {
        manualBreak.setAttributeNS(
          MATH_NAMESPACE,
          `${prefix}:alnAt`,
          String(expression.manualBreak.alignmentAt),
        );
      }
      properties.append(manualBreak);
    }
    properties.append(
      mathOnOffElement(document, prefix, 'aln', expression.alignment),
    );
    appendMathControlProperties(
      document,
      prefix,
      properties,
      expression.controlProperties,
      expression.controlRevision,
    );
    box.append(
      properties,
      expressionArgument(
        document,
        prefix,
        'e',
        expression.children,
        expression.childrenProperties,
      ),
    );
    return box;
  }
  if (expression.type === 'fraction') {
    const fraction = createMathElement(document, prefix, 'f');
    if (
      expression.fractionType !== 'bar' ||
      expression.controlProperties ||
      expression.controlRevision
    ) {
      const properties = createMathElement(document, prefix, 'fPr');
      if (expression.fractionType !== 'bar') {
        properties.append(
          mathValueElement(
            document,
            prefix,
            'type',
            expression.fractionType === 'noBar'
              ? 'noBar'
              : expression.fractionType === 'skewed'
                ? 'skw'
                : 'lin',
          ),
        );
      }
      appendMathControlProperties(
        document,
        prefix,
        properties,
        expression.controlProperties,
        expression.controlRevision,
      );
      fraction.append(properties);
    }
    fraction.append(
      expressionArgument(
        document,
        prefix,
        'num',
        expression.numerator,
        expression.numeratorProperties,
      ),
      expressionArgument(
        document,
        prefix,
        'den',
        expression.denominator,
        expression.denominatorProperties,
      ),
    );
    return fraction;
  }
  if (expression.type === 'superscript') {
    const script = createMathElement(document, prefix, 'sSup');
    if (expression.controlProperties || expression.controlRevision) {
      const properties = createMathElement(document, prefix, 'sSupPr');
      appendMathControlProperties(
        document,
        prefix,
        properties,
        expression.controlProperties,
        expression.controlRevision,
      );
      script.append(properties);
    }
    script.append(
      expressionArgument(
        document,
        prefix,
        'e',
        expression.base,
        expression.baseProperties,
      ),
      expressionArgument(
        document,
        prefix,
        'sup',
        expression.superScript,
        expression.superScriptProperties,
      ),
    );
    return script;
  }
  if (expression.type === 'subscript') {
    const script = createMathElement(document, prefix, 'sSub');
    if (expression.controlProperties || expression.controlRevision) {
      const properties = createMathElement(document, prefix, 'sSubPr');
      appendMathControlProperties(
        document,
        prefix,
        properties,
        expression.controlProperties,
        expression.controlRevision,
      );
      script.append(properties);
    }
    script.append(
      expressionArgument(
        document,
        prefix,
        'e',
        expression.base,
        expression.baseProperties,
      ),
      expressionArgument(
        document,
        prefix,
        'sub',
        expression.subScript,
        expression.subScriptProperties,
      ),
    );
    return script;
  }
  if (expression.type === 'subSuperScript') {
    const script = createMathElement(document, prefix, 'sSubSup');
    if (
      expression.alignScripts ||
      expression.controlProperties ||
      expression.controlRevision
    ) {
      const properties = createMathElement(document, prefix, 'sSubSupPr');
      if (expression.alignScripts) {
        properties.append(mathOnOffElement(document, prefix, 'alnScr', true));
      }
      appendMathControlProperties(
        document,
        prefix,
        properties,
        expression.controlProperties,
        expression.controlRevision,
      );
      script.append(properties);
    }
    script.append(
      expressionArgument(
        document,
        prefix,
        'e',
        expression.base,
        expression.baseProperties,
      ),
      expressionArgument(
        document,
        prefix,
        'sub',
        expression.subScript,
        expression.subScriptProperties,
      ),
      expressionArgument(
        document,
        prefix,
        'sup',
        expression.superScript,
        expression.superScriptProperties,
      ),
    );
    return script;
  }
  if (expression.type === 'preSubSuperScript') {
    const script = createMathElement(document, prefix, 'sPre');
    if (expression.controlProperties || expression.controlRevision) {
      const properties = createMathElement(document, prefix, 'sPrePr');
      appendMathControlProperties(
        document,
        prefix,
        properties,
        expression.controlProperties,
        expression.controlRevision,
      );
      script.append(properties);
    }
    script.append(
      expressionArgument(
        document,
        prefix,
        'sub',
        expression.subScript,
        expression.subScriptProperties,
      ),
      expressionArgument(
        document,
        prefix,
        'sup',
        expression.superScript,
        expression.superScriptProperties,
      ),
      expressionArgument(
        document,
        prefix,
        'e',
        expression.base,
        expression.baseProperties,
      ),
    );
    return script;
  }
  if (expression.type === 'lowerLimit' || expression.type === 'upperLimit') {
    const limit = createMathElement(
      document,
      prefix,
      expression.type === 'lowerLimit' ? 'limLow' : 'limUpp',
    );
    if (expression.controlProperties || expression.controlRevision) {
      const properties = createMathElement(
        document,
        prefix,
        expression.type === 'lowerLimit' ? 'limLowPr' : 'limUppPr',
      );
      appendMathControlProperties(
        document,
        prefix,
        properties,
        expression.controlProperties,
        expression.controlRevision,
      );
      limit.append(properties);
    }
    limit.append(
      expressionArgument(
        document,
        prefix,
        'e',
        expression.base,
        expression.baseProperties,
      ),
      expressionArgument(
        document,
        prefix,
        'lim',
        expression.limit,
        expression.limitProperties,
      ),
    );
    return limit;
  }
  if (expression.type === 'radical') {
    const radical = createMathElement(document, prefix, 'rad');
    const properties = createMathElement(document, prefix, 'radPr');
    if (!expression.degree) {
      properties.append(mathValueElement(document, prefix, 'degHide', '1'));
    }
    appendMathControlProperties(
      document,
      prefix,
      properties,
      expression.controlProperties,
      expression.controlRevision,
    );
    radical.append(
      properties,
      expressionArgument(
        document,
        prefix,
        'deg',
        expression.degree ?? [],
        expression.degreeProperties,
      ),
      expressionArgument(
        document,
        prefix,
        'e',
        expression.children,
        expression.childrenProperties,
      ),
    );
    return radical;
  }
  if (expression.type === 'function') {
    const function_ = createMathElement(document, prefix, 'func');
    const properties = createMathElement(document, prefix, 'funcPr');
    appendMathControlProperties(
      document,
      prefix,
      properties,
      expression.controlProperties,
      expression.controlRevision,
    );
    function_.append(
      properties,
      expressionArgument(
        document,
        prefix,
        'fName',
        expression.name,
        expression.nameProperties,
      ),
      expressionArgument(
        document,
        prefix,
        'e',
        expression.children,
        expression.childrenProperties,
      ),
    );
    return function_;
  }
  if (expression.type === 'nary') {
    const nary = createMathElement(document, prefix, 'nary');
    const properties = createMathElement(document, prefix, 'naryPr');
    properties.append(
      mathValueElement(document, prefix, 'chr', expression.operator),
      mathValueElement(
        document,
        prefix,
        'limLoc',
        expression.limitLocation === 'underOver' ? 'undOvr' : 'subSup',
      ),
    );
    if (expression.grow) {
      properties.append(mathOnOffElement(document, prefix, 'grow', true));
    }
    if (!expression.subScript) {
      properties.append(mathValueElement(document, prefix, 'subHide', '1'));
    }
    if (!expression.superScript) {
      properties.append(mathValueElement(document, prefix, 'supHide', '1'));
    }
    appendMathControlProperties(
      document,
      prefix,
      properties,
      expression.controlProperties,
      expression.controlRevision,
    );
    nary.append(
      properties,
      expressionArgument(
        document,
        prefix,
        'sub',
        expression.subScript ?? [],
        expression.subScriptProperties,
      ),
      expressionArgument(
        document,
        prefix,
        'sup',
        expression.superScript ?? [],
        expression.superScriptProperties,
      ),
      expressionArgument(
        document,
        prefix,
        'e',
        expression.children,
        expression.childrenProperties,
      ),
    );
    return nary;
  }
  if (expression.type === 'equationArray') {
    const equationArray = createMathElement(document, prefix, 'eqArr');
    const properties = createMathElement(document, prefix, 'eqArrPr');
    properties.append(
      mathValueElement(
        document,
        prefix,
        'baseJc',
        expression.baseAlignment === 'bottom'
          ? 'bot'
          : expression.baseAlignment,
      ),
      mathOnOffElement(
        document,
        prefix,
        'maxDist',
        expression.maximumDistribution,
      ),
      mathOnOffElement(
        document,
        prefix,
        'objDist',
        expression.objectDistribution,
      ),
      mathValueElement(
        document,
        prefix,
        'rSpRule',
        equationSpacingRule(expression.rowSpacingRule),
      ),
      mathValueElement(document, prefix, 'rSp', String(expression.rowSpacing)),
    );
    appendMathControlProperties(
      document,
      prefix,
      properties,
      expression.controlProperties,
      expression.controlRevision,
    );
    equationArray.append(properties);
    for (const [rowIndex, row] of expression.rows.entries()) {
      equationArray.append(
        expressionArgument(
          document,
          prefix,
          'e',
          row,
          expression.rowProperties?.[rowIndex] ?? undefined,
        ),
      );
    }
    return equationArray;
  }
  if (expression.type === 'matrix') {
    const matrix = createMathElement(document, prefix, 'm');
    const properties = createMathElement(document, prefix, 'mPr');
    properties.append(
      mathValueElement(
        document,
        prefix,
        'baseJc',
        expression.baseAlignment === 'bottom'
          ? 'bot'
          : expression.baseAlignment,
      ),
      mathValueElement(
        document,
        prefix,
        'plcHide',
        expression.placeholdersHidden ? '1' : '0',
      ),
    );
    if (expression.spacing) {
      properties.append(
        mathValueElement(
          document,
          prefix,
          'rSpRule',
          equationSpacingRule(expression.spacing.rowSpacingRule),
        ),
        mathValueElement(
          document,
          prefix,
          'cGpRule',
          equationSpacingRule(expression.spacing.columnGapRule),
        ),
        mathValueElement(
          document,
          prefix,
          'rSp',
          String(expression.spacing.rowSpacing),
        ),
        mathValueElement(
          document,
          prefix,
          'cSp',
          String(expression.spacing.minimumColumnWidthTwips),
        ),
        mathValueElement(
          document,
          prefix,
          'cGp',
          String(expression.spacing.columnGap),
        ),
      );
    }
    const columns = createMathElement(document, prefix, 'mcs');
    for (const group of matrixColumnGroups(expression.columnAlignments)) {
      const column = createMathElement(document, prefix, 'mc');
      const columnProperties = createMathElement(document, prefix, 'mcPr');
      columnProperties.append(
        mathValueElement(document, prefix, 'count', String(group.count)),
        mathValueElement(document, prefix, 'mcJc', group.alignment),
      );
      column.append(columnProperties);
      columns.append(column);
    }
    properties.append(columns);
    appendMathControlProperties(
      document,
      prefix,
      properties,
      expression.controlProperties,
      expression.controlRevision,
    );
    matrix.append(properties);
    for (const [rowIndex, row] of expression.rows.entries()) {
      const matrixRow = createMathElement(document, prefix, 'mr');
      for (const [cellIndex, cell] of row.entries()) {
        matrixRow.append(
          expressionArgument(
            document,
            prefix,
            'e',
            cell,
            expression.cellProperties?.[rowIndex]?.[cellIndex] ?? undefined,
          ),
        );
      }
      matrix.append(matrixRow);
    }
    return matrix;
  }
  const delimiter = createMathElement(document, prefix, 'd');
  const properties = createMathElement(document, prefix, 'dPr');
  properties.append(
    mathValueElement(document, prefix, 'begChr', expression.opening),
    mathValueElement(document, prefix, 'sepChr', expression.separator),
    mathValueElement(document, prefix, 'endChr', expression.closing),
  );
  if (expression.grow === false) {
    properties.append(mathOnOffElement(document, prefix, 'grow', false));
  }
  if (expression.shape === 'match') {
    properties.append(mathValueElement(document, prefix, 'shp', 'match'));
  }
  appendMathControlProperties(
    document,
    prefix,
    properties,
    expression.controlProperties,
    expression.controlRevision,
  );
  delimiter.append(properties);
  for (const [argumentIndex, argument] of expression.arguments.entries()) {
    delimiter.append(
      expressionArgument(
        document,
        prefix,
        'e',
        argument,
        expression.argumentProperties?.[argumentIndex] ?? undefined,
      ),
    );
  }
  return delimiter;
}

function equationSpacingRule(rule: WorkDocumentEquationSpacingRule): string {
  if (rule === 'oneAndHalf') return '1';
  if (rule === 'double') return '2';
  if (rule === 'exact') return '3';
  if (rule === 'multiple') return '4';
  return '0';
}

function justificationToOmml(
  justification: WorkDocumentEquationJustification,
): string {
  return justification;
}

function runScriptToOmml(script: WorkDocumentEquationRunScript): string {
  if (script === 'sansSerif') return 'sans-serif';
  if (script === 'doubleStruck') return 'double-struck';
  return script;
}

function runStyleToOmml(style: WorkDocumentEquationRunStyle): string {
  if (style === 'plain') return 'p';
  if (style === 'bold') return 'b';
  if (style === 'boldItalic') return 'bi';
  return 'i';
}

function createWordRunProperties(
  document: Document,
  properties: WorkDocumentEquationWordRunProperties,
): Element {
  const prefix = ensureWordPrefix(document.documentElement);
  const result = createWordElement(document, prefix, 'rPr');
  if (properties.fonts) {
    const fonts = createWordElement(document, prefix, 'rFonts');
    const values = [
      ['ascii', properties.fonts.ascii],
      ['hAnsi', properties.fonts.highAnsi],
      ['eastAsia', properties.fonts.eastAsia],
      ['cs', properties.fonts.complexScript],
      ['asciiTheme', properties.fonts.asciiTheme],
      ['hAnsiTheme', properties.fonts.highAnsiTheme],
      ['eastAsiaTheme', properties.fonts.eastAsiaTheme],
      ['cstheme', properties.fonts.complexScriptTheme],
      ['hint', properties.fonts.hint],
    ] as const;
    for (const [name, value] of values) {
      if (value !== undefined) setWordAttribute(fonts, prefix, name, value);
    }
    result.append(fonts);
  }
  for (const [name, value] of [
    ['b', properties.bold],
    ['bCs', properties.boldComplexScript],
    ['i', properties.italic],
    ['iCs', properties.italicComplexScript],
    ['caps', properties.allCaps],
    ['smallCaps', properties.smallCaps],
    ['strike', properties.strike],
    ['dstrike', properties.doubleStrike],
    ['outline', properties.outline],
    ['shadow', properties.shadow],
    ['emboss', properties.emboss],
    ['imprint', properties.imprint],
    ['noProof', properties.noProof],
    ['snapToGrid', properties.snapToGrid],
    ['vanish', properties.hidden],
    ['webHidden', properties.webHidden],
  ] as const) {
    if (value !== undefined) {
      result.append(createWordOnOffElement(document, prefix, name, value));
    }
  }
  if (properties.color) {
    const color = createWordElement(document, prefix, 'color');
    setWordColorAttributes(color, prefix, properties.color, 'val');
    result.append(color);
  }
  for (const [name, value] of [
    ['spacing', properties.characterSpacingTwips],
    ['w', properties.characterScalePercent],
    ['kern', properties.kerningThresholdHalfPoints],
    ['position', properties.positionHalfPoints],
  ] as const) {
    if (value !== undefined) {
      result.append(
        createWordValueElement(document, prefix, name, String(value)),
      );
    }
  }
  if (properties.fontSize !== undefined) {
    result.append(
      createWordValueElement(
        document,
        prefix,
        'sz',
        String(properties.fontSize * 2),
      ),
    );
  }
  if (properties.fontSizeComplexScript !== undefined) {
    result.append(
      createWordValueElement(
        document,
        prefix,
        'szCs',
        String(properties.fontSizeComplexScript * 2),
      ),
    );
  }
  if (properties.highlight) {
    result.append(
      createWordValueElement(
        document,
        prefix,
        'highlight',
        properties.highlight,
      ),
    );
  }
  if (properties.underline) {
    const underline = createWordValueElement(
      document,
      prefix,
      'u',
      properties.underline.style,
    );
    if (properties.underline.color) {
      setWordColorAttributes(
        underline,
        prefix,
        properties.underline.color,
        'color',
      );
    }
    result.append(underline);
  }
  if (properties.textEffect) {
    result.append(
      createWordValueElement(document, prefix, 'effect', properties.textEffect),
    );
  }
  if (properties.border) {
    const border = createWordValueElement(
      document,
      prefix,
      'bdr',
      properties.border.style,
    );
    if (properties.border.color) {
      setWordColorAttributes(border, prefix, properties.border.color, 'color');
    }
    if (properties.border.sizeEighthPoints !== undefined) {
      setWordAttribute(
        border,
        prefix,
        'sz',
        String(properties.border.sizeEighthPoints),
      );
    }
    if (properties.border.spacingPoints !== undefined) {
      setWordAttribute(
        border,
        prefix,
        'space',
        String(properties.border.spacingPoints),
      );
    }
    if (properties.border.shadow !== undefined) {
      setWordAttribute(
        border,
        prefix,
        'shadow',
        properties.border.shadow ? '1' : '0',
      );
    }
    if (properties.border.frame !== undefined) {
      setWordAttribute(
        border,
        prefix,
        'frame',
        properties.border.frame ? '1' : '0',
      );
    }
    result.append(border);
  }
  if (properties.shading) {
    const shading = createWordValueElement(
      document,
      prefix,
      'shd',
      properties.shading.pattern,
    );
    if (properties.shading.color) {
      setWordColorAttributes(
        shading,
        prefix,
        properties.shading.color,
        'color',
      );
    }
    if (properties.shading.fill) {
      setWordColorAttributes(
        shading,
        prefix,
        properties.shading.fill,
        'fill',
        'themeFill',
        'themeFillTint',
        'themeFillShade',
      );
    }
    result.append(shading);
  }
  if (properties.fitText) {
    const fitText = createWordValueElement(
      document,
      prefix,
      'fitText',
      String(properties.fitText.widthTwips),
    );
    if (properties.fitText.id !== undefined) {
      setWordAttribute(fitText, prefix, 'id', String(properties.fitText.id));
    }
    result.append(fitText);
  }
  if (properties.verticalAlignment) {
    result.append(
      createWordValueElement(
        document,
        prefix,
        'vertAlign',
        properties.verticalAlignment,
      ),
    );
  }
  for (const [name, value] of [
    ['rtl', properties.rightToLeft],
    ['cs', properties.complexScript],
  ] as const) {
    if (value !== undefined) {
      result.append(createWordOnOffElement(document, prefix, name, value));
    }
  }
  if (properties.emphasisMark) {
    result.append(
      createWordValueElement(document, prefix, 'em', properties.emphasisMark),
    );
  }
  if (properties.languages) {
    const languages = createWordElement(document, prefix, 'lang');
    for (const [name, value] of [
      ['val', properties.languages.latin],
      ['eastAsia', properties.languages.eastAsia],
      ['bidi', properties.languages.bidi],
    ] as const) {
      if (value !== undefined) {
        setWordAttribute(languages, prefix, name, value);
      }
    }
    result.append(languages);
  }
  if (properties.eastAsianLayout) {
    const layout = createWordElement(document, prefix, 'eastAsianLayout');
    if (properties.eastAsianLayout.id !== undefined) {
      setWordAttribute(
        layout,
        prefix,
        'id',
        String(properties.eastAsianLayout.id),
      );
    }
    for (const [name, value] of [
      ['combine', properties.eastAsianLayout.combine],
      ['combineBrackets', properties.eastAsianLayout.combineBrackets],
      ['vert', properties.eastAsianLayout.vertical],
      ['vertCompress', properties.eastAsianLayout.verticalCompress],
    ] as const) {
      if (value !== undefined) {
        setWordAttribute(
          layout,
          prefix,
          name,
          typeof value === 'boolean' ? (value ? '1' : '0') : value,
        );
      }
    }
    result.append(layout);
  }
  if (properties.paragraphMarkAlwaysHidden !== undefined) {
    result.append(
      createWordOnOffElement(
        document,
        prefix,
        'specVanish',
        properties.paragraphMarkAlwaysHidden,
      ),
    );
  }
  if (properties.glow) {
    result.append(createWord2010Glow(document, properties.glow));
  }
  if (properties.shadowEffect) {
    result.append(
      createWord2010ShadowEffect(document, properties.shadowEffect),
    );
  }
  if (properties.reflectionEffect) {
    result.append(
      createWord2010ReflectionEffect(document, properties.reflectionEffect),
    );
  }
  if (properties.textOutlineEffect) {
    result.append(
      createWord2010TextOutlineEffect(document, properties.textOutlineEffect),
    );
  }
  if (properties.textFillEffect) {
    result.append(
      createWord2010TextFillEffect(document, properties.textFillEffect),
    );
  }
  if (properties.scene3D) {
    result.append(createWord2010Scene3D(document, properties.scene3D));
  }
  if (properties.properties3D) {
    result.append(
      createWord2010Properties3D(document, properties.properties3D),
    );
  }
  return result;
}

function appendMathControlProperties(
  document: Document,
  prefix: string,
  parent: Element,
  properties: WorkDocumentEquationWordRunProperties | undefined,
  revision: WorkDocumentEquationControlRevision | undefined,
): void {
  if (!properties && !revision) return;
  const controlProperties = createMathElement(document, prefix, 'ctrlPr');
  if (revision) {
    const wordPrefix = ensureWordPrefix(document.documentElement);
    controlProperties.append(
      createWordMathControlRevision(document, wordPrefix, revision, properties),
    );
  } else if (properties) {
    controlProperties.append(createWordRunProperties(document, properties));
  }
  parent.append(controlProperties);
}

function createWordMathControlRevision(
  document: Document,
  prefix: string,
  revision: WorkDocumentEquationControlRevision,
  properties: WorkDocumentEquationWordRunProperties | undefined,
): Element {
  const name =
    revision.kind === 'insertion'
      ? 'ins'
      : revision.kind === 'deletion'
        ? 'del'
        : revision.kind;
  const result = createWordElement(document, prefix, name);
  setWordAttribute(result, prefix, 'id', String(revision.id));
  setWordAttribute(result, prefix, 'author', revision.author);
  if (revision.date) {
    setWordAttribute(result, prefix, 'date', revision.date);
  }
  if (revision.dateUtc) {
    const dateUtcPrefix = ensureWordDateUtcPrefix(document.documentElement);
    result.setAttributeNS(
      WORD_DATE_UTC_NAMESPACE,
      `${dateUtcPrefix}:dateUtc`,
      revision.dateUtc,
    );
  }
  if (revision.child) {
    result.append(
      createWordMathControlRevision(
        document,
        prefix,
        revision.child,
        properties,
      ),
    );
  } else if (properties) {
    result.append(createWordRunProperties(document, properties));
  }
  return result;
}

function setWordColorAttributes(
  element: Element,
  prefix: string,
  color: WorkDocumentEquationWordColor,
  valueAttribute: 'val' | 'color' | 'fill',
  themeAttribute = 'themeColor',
  tintAttribute = 'themeTint',
  shadeAttribute = 'themeShade',
): void {
  if (color.value) {
    setWordAttribute(
      element,
      prefix,
      valueAttribute,
      color.value === 'auto' ? 'auto' : color.value.slice(1).toUpperCase(),
    );
  }
  if (color.theme) {
    setWordAttribute(element, prefix, themeAttribute, color.theme);
  }
  if (color.tint) setWordAttribute(element, prefix, tintAttribute, color.tint);
  if (color.shade) {
    setWordAttribute(element, prefix, shadeAttribute, color.shade);
  }
}

function createWord2010Glow(
  document: Document,
  glow: NonNullable<WorkDocumentEquationWordRunProperties['glow']>,
): Element {
  const prefix = ensureWord2010Prefix(document.documentElement);
  const result = createWord2010Element(document, prefix, 'glow');
  if (glow.radiusEmus !== undefined) {
    setWord2010Attribute(result, prefix, 'rad', String(glow.radiusEmus));
  }
  result.append(createWord2010EffectColor(document, prefix, glow.color));
  return result;
}

function createWord2010ShadowEffect(
  document: Document,
  shadow: NonNullable<WorkDocumentEquationWordRunProperties['shadowEffect']>,
): Element {
  const prefix = ensureWord2010Prefix(document.documentElement);
  const result = createWord2010Element(document, prefix, 'shadow');
  const attributes = [
    ['blurRad', shadow.blurRadiusEmus],
    ['dist', shadow.distanceEmus],
    [
      'dir',
      shadow.directionDegrees === undefined
        ? undefined
        : shadow.directionDegrees * WORD_ANGLE_UNITS_PER_DEGREE,
    ],
    [
      'sx',
      shadow.horizontalScalePercent === undefined
        ? undefined
        : shadow.horizontalScalePercent * WORD_PERCENTAGE_UNITS_PER_PERCENT,
    ],
    [
      'sy',
      shadow.verticalScalePercent === undefined
        ? undefined
        : shadow.verticalScalePercent * WORD_PERCENTAGE_UNITS_PER_PERCENT,
    ],
    [
      'kx',
      shadow.horizontalSkewDegrees === undefined
        ? undefined
        : shadow.horizontalSkewDegrees * WORD_ANGLE_UNITS_PER_DEGREE,
    ],
    [
      'ky',
      shadow.verticalSkewDegrees === undefined
        ? undefined
        : shadow.verticalSkewDegrees * WORD_ANGLE_UNITS_PER_DEGREE,
    ],
  ] as const;
  for (const [name, value] of attributes) {
    if (value !== undefined) {
      setWord2010Attribute(result, prefix, name, String(value));
    }
  }
  if (shadow.alignment) {
    setWord2010Attribute(
      result,
      prefix,
      'algn',
      WORD_2010_RECTANGLE_ALIGNMENTS[shadow.alignment],
    );
  }
  result.append(createWord2010EffectColor(document, prefix, shadow.color));
  return result;
}

function createWord2010ReflectionEffect(
  document: Document,
  reflection: NonNullable<
    WorkDocumentEquationWordRunProperties['reflectionEffect']
  >,
): Element {
  const prefix = ensureWord2010Prefix(document.documentElement);
  const result = createWord2010Element(document, prefix, 'reflection');
  const attributes = [
    ['blurRad', reflection.blurRadiusEmus],
    [
      'stA',
      reflection.startOpacityPercent === undefined
        ? undefined
        : reflection.startOpacityPercent * WORD_PERCENTAGE_UNITS_PER_PERCENT,
    ],
    [
      'stPos',
      reflection.startPositionPercent === undefined
        ? undefined
        : reflection.startPositionPercent * WORD_PERCENTAGE_UNITS_PER_PERCENT,
    ],
    [
      'endA',
      reflection.endOpacityPercent === undefined
        ? undefined
        : reflection.endOpacityPercent * WORD_PERCENTAGE_UNITS_PER_PERCENT,
    ],
    [
      'endPos',
      reflection.endPositionPercent === undefined
        ? undefined
        : reflection.endPositionPercent * WORD_PERCENTAGE_UNITS_PER_PERCENT,
    ],
    ['dist', reflection.distanceEmus],
    [
      'dir',
      reflection.directionDegrees === undefined
        ? undefined
        : reflection.directionDegrees * WORD_ANGLE_UNITS_PER_DEGREE,
    ],
    [
      'fadeDir',
      reflection.fadeDirectionDegrees === undefined
        ? undefined
        : reflection.fadeDirectionDegrees * WORD_ANGLE_UNITS_PER_DEGREE,
    ],
    [
      'sx',
      reflection.horizontalScalePercent === undefined
        ? undefined
        : reflection.horizontalScalePercent * WORD_PERCENTAGE_UNITS_PER_PERCENT,
    ],
    [
      'sy',
      reflection.verticalScalePercent === undefined
        ? undefined
        : reflection.verticalScalePercent * WORD_PERCENTAGE_UNITS_PER_PERCENT,
    ],
    [
      'kx',
      reflection.horizontalSkewDegrees === undefined
        ? undefined
        : reflection.horizontalSkewDegrees * WORD_ANGLE_UNITS_PER_DEGREE,
    ],
    [
      'ky',
      reflection.verticalSkewDegrees === undefined
        ? undefined
        : reflection.verticalSkewDegrees * WORD_ANGLE_UNITS_PER_DEGREE,
    ],
  ] as const;
  for (const [name, value] of attributes) {
    if (value !== undefined) {
      setWord2010Attribute(result, prefix, name, String(value));
    }
  }
  if (reflection.alignment) {
    setWord2010Attribute(
      result,
      prefix,
      'algn',
      WORD_2010_RECTANGLE_ALIGNMENTS[reflection.alignment],
    );
  }
  return result;
}

function createWord2010TextOutlineEffect(
  document: Document,
  outline: NonNullable<
    WorkDocumentEquationWordRunProperties['textOutlineEffect']
  >,
): Element {
  const prefix = ensureWord2010Prefix(document.documentElement);
  const result = createWord2010Element(document, prefix, 'textOutline');
  if (outline.widthEmus !== undefined) {
    setWord2010Attribute(result, prefix, 'w', String(outline.widthEmus));
  }
  if (outline.cap) {
    setWord2010Attribute(
      result,
      prefix,
      'cap',
      WORD_2010_TEXT_OUTLINE_CAPS[outline.cap],
    );
  }
  if (outline.compound) {
    setWord2010Attribute(
      result,
      prefix,
      'cmpd',
      WORD_2010_TEXT_OUTLINE_COMPOUNDS[outline.compound],
    );
  }
  if (outline.alignment) {
    setWord2010Attribute(
      result,
      prefix,
      'algn',
      WORD_2010_TEXT_OUTLINE_ALIGNMENTS[outline.alignment],
    );
  }
  if (outline.fill) {
    result.append(createWord2010EffectFill(document, prefix, outline.fill));
  }
  if (outline.dash) {
    const dash = createWord2010Element(document, prefix, 'prstDash');
    if (outline.dash.preset) {
      setWord2010Attribute(
        dash,
        prefix,
        'val',
        WORD_2010_PRESET_LINE_DASHES[outline.dash.preset],
      );
    }
    result.append(dash);
  }
  if (outline.join) {
    const join = createWord2010Element(document, prefix, outline.join.type);
    if (
      outline.join.type === 'miter' &&
      outline.join.limitPercent !== undefined
    ) {
      setWord2010Attribute(
        join,
        prefix,
        'lim',
        String(outline.join.limitPercent * WORD_PERCENTAGE_UNITS_PER_PERCENT),
      );
    }
    result.append(join);
  }
  return result;
}

function createWord2010TextFillEffect(
  document: Document,
  effect: NonNullable<WorkDocumentEquationWordRunProperties['textFillEffect']>,
): Element {
  const prefix = ensureWord2010Prefix(document.documentElement);
  const result = createWord2010Element(document, prefix, 'textFill');
  if (effect.fill) {
    result.append(createWord2010EffectFill(document, prefix, effect.fill));
  }
  return result;
}

function createWord2010Scene3D(
  document: Document,
  scene: NonNullable<WorkDocumentEquationWordRunProperties['scene3D']>,
): Element {
  const prefix = ensureWord2010Prefix(document.documentElement);
  const result = createWord2010Element(document, prefix, 'scene3d');
  const camera = createWord2010Element(document, prefix, 'camera');
  setWord2010Attribute(camera, prefix, 'prst', scene.cameraPreset);
  result.append(camera);

  const lightRig = createWord2010Element(document, prefix, 'lightRig');
  setWord2010Attribute(
    lightRig,
    prefix,
    'rig',
    WORD_2010_SCENE_3D_LIGHT_RIG_PRESETS[scene.lightRig.preset],
  );
  setWord2010Attribute(
    lightRig,
    prefix,
    'dir',
    WORD_2010_SCENE_3D_LIGHT_RIG_DIRECTIONS[scene.lightRig.direction],
  );
  if (scene.lightRig.rotation) {
    const rotation = createWord2010Element(document, prefix, 'rot');
    for (const [name, degrees] of [
      ['lat', scene.lightRig.rotation.latitudeDegrees],
      ['lon', scene.lightRig.rotation.longitudeDegrees],
      ['rev', scene.lightRig.rotation.revolutionDegrees],
    ] as const) {
      setWord2010Attribute(
        rotation,
        prefix,
        name,
        String(degrees * WORD_ANGLE_UNITS_PER_DEGREE),
      );
    }
    lightRig.append(rotation);
  }
  result.append(lightRig);
  return result;
}

function createWord2010Properties3D(
  document: Document,
  properties: NonNullable<
    WorkDocumentEquationWordRunProperties['properties3D']
  >,
): Element {
  const prefix = ensureWord2010Prefix(document.documentElement);
  const result = createWord2010Element(document, prefix, 'props3d');
  if (properties.extrusionHeightEmus !== undefined) {
    setWord2010Attribute(
      result,
      prefix,
      'extrusionH',
      String(properties.extrusionHeightEmus),
    );
  }
  if (properties.contourWidthEmus !== undefined) {
    setWord2010Attribute(
      result,
      prefix,
      'contourW',
      String(properties.contourWidthEmus),
    );
  }
  if (properties.materialPreset !== undefined) {
    setWord2010Attribute(
      result,
      prefix,
      'prstMaterial',
      WORD_2010_PROPERTIES_3D_MATERIAL_PRESETS[properties.materialPreset],
    );
  }
  if (properties.topBevel !== undefined) {
    result.append(
      createWord2010Properties3DBevel(
        document,
        prefix,
        'bevelT',
        properties.topBevel,
      ),
    );
  }
  if (properties.bottomBevel !== undefined) {
    result.append(
      createWord2010Properties3DBevel(
        document,
        prefix,
        'bevelB',
        properties.bottomBevel,
      ),
    );
  }
  if (properties.extrusionColor !== undefined) {
    const color = createWord2010Element(document, prefix, 'extrusionClr');
    color.append(
      createWord2010EffectColor(document, prefix, properties.extrusionColor),
    );
    result.append(color);
  }
  if (properties.contourColor !== undefined) {
    const color = createWord2010Element(document, prefix, 'contourClr');
    color.append(
      createWord2010EffectColor(document, prefix, properties.contourColor),
    );
    result.append(color);
  }
  return result;
}

function createWord2010Properties3DBevel(
  document: Document,
  prefix: string,
  name: 'bevelT' | 'bevelB',
  bevel: WorkDocumentEquationWordBevel,
): Element {
  const result = createWord2010Element(document, prefix, name);
  if (bevel.widthEmus !== undefined) {
    setWord2010Attribute(result, prefix, 'w', String(bevel.widthEmus));
  }
  if (bevel.heightEmus !== undefined) {
    setWord2010Attribute(result, prefix, 'h', String(bevel.heightEmus));
  }
  if (bevel.preset !== undefined) {
    setWord2010Attribute(
      result,
      prefix,
      'prst',
      WORD_2010_PROPERTIES_3D_BEVEL_PRESETS[bevel.preset],
    );
  }
  return result;
}

function createWord2010EffectFill(
  document: Document,
  prefix: string,
  fill: WorkDocumentEquationWordEffectFill,
): Element {
  if (fill.type === 'none') {
    return createWord2010Element(document, prefix, 'noFill');
  }
  if (fill.type === 'solid') {
    const result = createWord2010Element(document, prefix, 'solidFill');
    if (fill.color) {
      result.append(createWord2010EffectColor(document, prefix, fill.color));
    }
    return result;
  }
  const result = createWord2010Element(document, prefix, 'gradFill');
  if (fill.stops) {
    const stopList = createWord2010Element(document, prefix, 'gsLst');
    for (const stop of fill.stops) {
      const element = createWord2010Element(document, prefix, 'gs');
      setWord2010Attribute(
        element,
        prefix,
        'pos',
        String(stop.positionPercent * WORD_PERCENTAGE_UNITS_PER_PERCENT),
      );
      element.append(createWord2010EffectColor(document, prefix, stop.color));
      stopList.append(element);
    }
    result.append(stopList);
  }
  if (fill.shade?.type === 'linear') {
    const linear = createWord2010Element(document, prefix, 'lin');
    if (fill.shade.angleDegrees !== undefined) {
      setWord2010Attribute(
        linear,
        prefix,
        'ang',
        String(fill.shade.angleDegrees * WORD_ANGLE_UNITS_PER_DEGREE),
      );
    }
    if (fill.shade.scaled !== undefined) {
      setWord2010Attribute(
        linear,
        prefix,
        'scaled',
        fill.shade.scaled ? '1' : '0',
      );
    }
    result.append(linear);
  } else if (fill.shade?.type === 'path') {
    const path = createWord2010Element(document, prefix, 'path');
    if (fill.shade.path) {
      setWord2010Attribute(
        path,
        prefix,
        'path',
        fill.shade.path === 'rectangle' ? 'rect' : fill.shade.path,
      );
    }
    if (fill.shade.fillToRectangle) {
      const rectangle = createWord2010Element(document, prefix, 'fillToRect');
      for (const [name, value] of [
        ['l', fill.shade.fillToRectangle.leftPercent],
        ['t', fill.shade.fillToRectangle.topPercent],
        ['r', fill.shade.fillToRectangle.rightPercent],
        ['b', fill.shade.fillToRectangle.bottomPercent],
      ] as const) {
        if (value !== undefined) {
          setWord2010Attribute(
            rectangle,
            prefix,
            name,
            String(value * WORD_PERCENTAGE_UNITS_PER_PERCENT),
          );
        }
      }
      path.append(rectangle);
    }
    result.append(path);
  }
  return result;
}

function createWord2010EffectColor(
  document: Document,
  prefix: string,
  color: WorkDocumentEquationWordEffectColor,
): Element {
  const result = createWord2010Element(
    document,
    prefix,
    color.type === 'rgb' ? 'srgbClr' : 'schemeClr',
  );
  setWord2010Attribute(
    result,
    prefix,
    'val',
    color.type === 'rgb'
      ? color.value.slice(1).toUpperCase()
      : (WORD_2010_SCHEME_COLORS.get(color.value) ?? color.value),
  );
  for (const transform of color.transforms ?? []) {
    const element = createWord2010Element(
      document,
      prefix,
      WORD_2010_COLOR_TRANSFORM_NAMES[transform.type],
    );
    setWord2010Attribute(element, prefix, 'val', String(transform.value));
    result.append(element);
  }
  return result;
}

function matrixColumnGroups(
  alignments: readonly ('left' | 'center' | 'right')[],
): Array<{ alignment: 'left' | 'center' | 'right'; count: number }> {
  const groups: Array<{
    alignment: 'left' | 'center' | 'right';
    count: number;
  }> = [];
  for (const alignment of alignments) {
    const current = groups.at(-1);
    if (current?.alignment === alignment) current.count += 1;
    else groups.push({ alignment, count: 1 });
  }
  return groups;
}

function expressionArgument(
  document: Document,
  prefix: string,
  name: string,
  children: readonly WorkDocumentEquationExpression[],
  properties?: WorkDocumentEquationArgumentProperties,
): Element {
  const argument = createMathElement(document, prefix, name);
  if (properties?.size !== undefined) {
    const argumentProperties = createMathElement(document, prefix, 'argPr');
    argumentProperties.append(
      mathValueElement(document, prefix, 'argSz', String(properties.size)),
    );
    argument.append(argumentProperties);
  }
  appendExpressions(document, prefix, argument, children);
  appendMathControlProperties(
    document,
    prefix,
    argument,
    properties?.controlProperties,
    properties?.controlRevision,
  );
  return argument;
}

function mathValueElement(
  document: Document,
  prefix: string,
  name: string,
  value: string,
): Element {
  const element = createMathElement(document, prefix, name);
  element.setAttributeNS(MATH_NAMESPACE, `${prefix}:val`, value);
  return element;
}

function mathOnOffElement(
  document: Document,
  prefix: string,
  name: string,
  value: boolean,
): Element {
  return mathValueElement(document, prefix, name, value ? '1' : '0');
}

function createMathElement(
  document: Document,
  prefix: string,
  name: string,
): Element {
  return document.createElementNS(MATH_NAMESPACE, `${prefix}:${name}`);
}

function createWordValueElement(
  document: Document,
  prefix: string,
  name: string,
  value: string,
): Element {
  const element = createWordElement(document, prefix, name);
  setWordAttribute(element, prefix, 'val', value);
  return element;
}

function createWordOnOffElement(
  document: Document,
  prefix: string,
  name: string,
  value: boolean,
): Element {
  return createWordValueElement(document, prefix, name, value ? '1' : '0');
}

function createWordElement(
  document: Document,
  prefix: string,
  name: string,
): Element {
  return document.createElementNS(WORD_NAMESPACE, `${prefix}:${name}`);
}

function setWordAttribute(
  element: Element,
  prefix: string,
  name: string,
  value: string,
): void {
  element.setAttributeNS(WORD_NAMESPACE, `${prefix}:${name}`, value);
}

function createWord2010Element(
  document: Document,
  prefix: string,
  name: string,
): Element {
  return document.createElementNS(WORD_2010_NAMESPACE, `${prefix}:${name}`);
}

function setWord2010Attribute(
  element: Element,
  prefix: string,
  name: string,
  value: string,
): void {
  element.setAttributeNS(WORD_2010_NAMESPACE, `${prefix}:${name}`, value);
}

function ensureMathPrefix(root: Element): string {
  const existing = xmlDeclaredPrefix(root, MATH_NAMESPACE);
  if (existing) return existing;
  const preferred = xmlNamespaceUri(root, 'm');
  if (!preferred || preferred === MATH_NAMESPACE) {
    root.setAttributeNS(XMLNS_NAMESPACE, 'xmlns:m', MATH_NAMESPACE);
    return 'm';
  }
  let index = 1;
  let prefix = '';
  do {
    prefix = `a3sm${index}`;
    index += 1;
  } while (xmlNamespaceUri(root, prefix));
  root.setAttributeNS(XMLNS_NAMESPACE, `xmlns:${prefix}`, MATH_NAMESPACE);
  return prefix;
}

function ensureWordPrefix(root: Element): string {
  const existing = xmlDeclaredPrefix(root, WORD_NAMESPACE);
  if (existing) return existing;
  const preferred = xmlNamespaceUri(root, 'w');
  if (!preferred || preferred === WORD_NAMESPACE) {
    root.setAttributeNS(XMLNS_NAMESPACE, 'xmlns:w', WORD_NAMESPACE);
    return 'w';
  }
  let index = 1;
  let prefix = '';
  do {
    prefix = `a3sw${index}`;
    index += 1;
  } while (xmlNamespaceUri(root, prefix));
  root.setAttributeNS(XMLNS_NAMESPACE, `xmlns:${prefix}`, WORD_NAMESPACE);
  return prefix;
}

function ensureWord2010Prefix(root: Element): string {
  const existing = xmlDeclaredPrefix(root, WORD_2010_NAMESPACE);
  const prefix = existing ?? availableNamespacePrefix(root, 'w14', 'a3sw14');
  if (!existing) {
    root.setAttributeNS(
      XMLNS_NAMESPACE,
      `xmlns:${prefix}`,
      WORD_2010_NAMESPACE,
    );
  }
  ensureIgnorableNamespace(root, prefix, WORD_2010_NAMESPACE);
  return prefix;
}

function ensureIgnorableNamespace(
  root: Element,
  prefix: string,
  namespace: string,
): void {
  const attribute = Array.from(root.attributes).find(
    (item) =>
      xmlAttributeLocalName(item) === 'Ignorable' &&
      xmlAttributeNamespace(root, item) === MARKUP_COMPATIBILITY_NAMESPACE,
  );
  const tokens = (attribute?.value ?? '').trim().split(/\s+/u).filter(Boolean);
  if (!tokens.some((token) => xmlNamespaceUri(root, token) === namespace)) {
    tokens.push(prefix);
  }
  const compatibilityPrefix =
    xmlDeclaredPrefix(root, MARKUP_COMPATIBILITY_NAMESPACE) ??
    availableNamespacePrefix(root, 'mc', 'a3smc');
  if (!xmlDeclaredPrefix(root, MARKUP_COMPATIBILITY_NAMESPACE)) {
    root.setAttributeNS(
      XMLNS_NAMESPACE,
      `xmlns:${compatibilityPrefix}`,
      MARKUP_COMPATIBILITY_NAMESPACE,
    );
  }
  attribute?.ownerElement?.removeAttributeNode(attribute);
  root.setAttributeNS(
    MARKUP_COMPATIBILITY_NAMESPACE,
    `${compatibilityPrefix}:Ignorable`,
    tokens.join(' '),
  );
}

function availableNamespacePrefix(
  root: Element,
  preferred: string,
  fallback: string,
): string {
  if (!xmlNamespaceUri(root, preferred)) return preferred;
  let index = 1;
  let prefix = '';
  do {
    prefix = `${fallback}${index}`;
    index += 1;
  } while (xmlNamespaceUri(root, prefix));
  return prefix;
}

function ensureWordDateUtcPrefix(root: Element): string {
  const existing = xmlDeclaredPrefix(root, WORD_DATE_UTC_NAMESPACE);
  if (existing) return existing;
  const preferred = xmlNamespaceUri(root, 'w16du');
  if (!preferred || preferred === WORD_DATE_UTC_NAMESPACE) {
    root.setAttributeNS(
      XMLNS_NAMESPACE,
      'xmlns:w16du',
      WORD_DATE_UTC_NAMESPACE,
    );
    return 'w16du';
  }
  let index = 1;
  let prefix = '';
  do {
    prefix = `a3sw16du${index}`;
    index += 1;
  } while (xmlNamespaceUri(root, prefix));
  root.setAttributeNS(
    XMLNS_NAMESPACE,
    `xmlns:${prefix}`,
    WORD_DATE_UTC_NAMESPACE,
  );
  return prefix;
}

function supportedPartRoot(root: Element, path: string): boolean {
  const expected = /^word\/header\d*\.xml$/i.test(path)
    ? 'hdr'
    : /^word\/footer\d*\.xml$/i.test(path)
      ? 'ftr'
      : /^word\/footnotes\.xml$/i.test(path)
        ? 'footnotes'
        : /^word\/endnotes\.xml$/i.test(path)
          ? 'endnotes'
          : 'document';
  return (
    root.localName === expected &&
    DOCX_WORDPROCESSING_NAMESPACES.has(root.namespaceURI ?? '')
  );
}
