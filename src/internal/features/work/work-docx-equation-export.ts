import JSZip from 'jszip';
import {
  documentEquationFromElement,
  normalizeDocumentEquation,
  type WorkDocumentEquation,
  type WorkDocumentEquationExpression,
} from './work-document-equations';
import { DOCX_WORDPROCESSING_NAMESPACES } from './work-docx-ignorable-extension-preservation';
import {
  XMLNS_NAMESPACE,
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
const XML_NAMESPACE = 'http://www.w3.org/XML/1998/namespace';
const EQUATION_PART_PATTERN =
  /^word\/(?:document|header\d*|footer\d*|footnotes|endnotes)\.xml$/i;
const MAX_EQUATION_PATCHES = 4_096;

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
    accent.append(
      properties,
      expressionArgument(document, prefix, 'e', expression.children),
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
    bar.append(
      properties,
      expressionArgument(document, prefix, 'e', expression.children),
    );
    return bar;
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
    borderBox.append(
      properties,
      expressionArgument(document, prefix, 'e', expression.children),
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
    box.append(
      properties,
      expressionArgument(document, prefix, 'e', expression.children),
    );
    return box;
  }
  if (expression.type === 'fraction') {
    const fraction = createMathElement(document, prefix, 'f');
    if (expression.fractionType !== 'bar') {
      const properties = createMathElement(document, prefix, 'fPr');
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
      fraction.append(properties);
    }
    fraction.append(
      expressionArgument(document, prefix, 'num', expression.numerator),
      expressionArgument(document, prefix, 'den', expression.denominator),
    );
    return fraction;
  }
  if (expression.type === 'superscript') {
    const script = createMathElement(document, prefix, 'sSup');
    script.append(
      expressionArgument(document, prefix, 'e', expression.base),
      expressionArgument(document, prefix, 'sup', expression.superScript),
    );
    return script;
  }
  if (expression.type === 'subscript') {
    const script = createMathElement(document, prefix, 'sSub');
    script.append(
      expressionArgument(document, prefix, 'e', expression.base),
      expressionArgument(document, prefix, 'sub', expression.subScript),
    );
    return script;
  }
  if (expression.type === 'subSuperScript') {
    const script = createMathElement(document, prefix, 'sSubSup');
    script.append(
      expressionArgument(document, prefix, 'e', expression.base),
      expressionArgument(document, prefix, 'sub', expression.subScript),
      expressionArgument(document, prefix, 'sup', expression.superScript),
    );
    return script;
  }
  if (expression.type === 'radical') {
    const radical = createMathElement(document, prefix, 'rad');
    const properties = createMathElement(document, prefix, 'radPr');
    if (!expression.degree) {
      properties.append(mathValueElement(document, prefix, 'degHide', '1'));
    }
    radical.append(properties);
    if (expression.degree) {
      radical.append(
        expressionArgument(document, prefix, 'deg', expression.degree),
      );
    }
    radical.append(
      expressionArgument(document, prefix, 'e', expression.children),
    );
    return radical;
  }
  if (expression.type === 'function') {
    const function_ = createMathElement(document, prefix, 'func');
    function_.append(
      createMathElement(document, prefix, 'funcPr'),
      expressionArgument(document, prefix, 'fName', expression.name),
      expressionArgument(document, prefix, 'e', expression.children),
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
    if (!expression.subScript) {
      properties.append(mathValueElement(document, prefix, 'subHide', '1'));
    }
    if (!expression.superScript) {
      properties.append(mathValueElement(document, prefix, 'supHide', '1'));
    }
    nary.append(properties);
    if (expression.subScript) {
      nary.append(
        expressionArgument(document, prefix, 'sub', expression.subScript),
      );
    }
    if (expression.superScript) {
      nary.append(
        expressionArgument(document, prefix, 'sup', expression.superScript),
      );
    }
    nary.append(expressionArgument(document, prefix, 'e', expression.children));
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
        equationArrayRowSpacingRule(expression.rowSpacingRule),
      ),
      mathValueElement(document, prefix, 'rSp', String(expression.rowSpacing)),
    );
    equationArray.append(properties);
    for (const row of expression.rows) {
      equationArray.append(expressionArgument(document, prefix, 'e', row));
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
    matrix.append(properties);
    for (const row of expression.rows) {
      const matrixRow = createMathElement(document, prefix, 'mr');
      for (const cell of row) {
        matrixRow.append(expressionArgument(document, prefix, 'e', cell));
      }
      matrix.append(matrixRow);
    }
    return matrix;
  }
  const delimiter = createMathElement(document, prefix, 'd');
  const properties = createMathElement(document, prefix, 'dPr');
  properties.append(
    mathValueElement(document, prefix, 'begChr', expression.opening),
    mathValueElement(document, prefix, 'endChr', expression.closing),
    mathValueElement(document, prefix, 'sepChr', expression.separator),
  );
  delimiter.append(properties);
  for (const argument of expression.arguments) {
    delimiter.append(expressionArgument(document, prefix, 'e', argument));
  }
  return delimiter;
}

function equationArrayRowSpacingRule(
  rule: Extract<
    WorkDocumentEquationExpression,
    { type: 'equationArray' }
  >['rowSpacingRule'],
): string {
  if (rule === 'oneAndHalf') return '1';
  if (rule === 'double') return '2';
  if (rule === 'exact') return '3';
  if (rule === 'multiple') return '4';
  return '0';
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
): Element {
  const argument = createMathElement(document, prefix, name);
  appendExpressions(document, prefix, argument, children);
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
