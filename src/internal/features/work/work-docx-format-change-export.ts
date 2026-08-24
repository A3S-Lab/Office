import JSZip from 'jszip';
import { normalizeDocumentCharacterScalePercent } from './work-document-character-scale';
import { normalizeDocumentCharacterPositionHalfPoints } from './work-document-character-position';
import { normalizeDocumentCharacterSpacingTwips } from './work-document-character-spacing';
import { normalizeDocumentEmphasisMark } from './work-document-emphasis';
import { normalizeDocumentKerningThresholdHalfPoints } from './work-document-kerning';
import { normalizeDocumentHiddenText } from './work-document-hidden-text';
import {
  DOCUMENT_LEGACY_TEXT_EFFECT_NAMES,
  documentLegacyTextEffectsConflict,
  normalizeDocumentLegacyTextEffect,
  type WorkDocumentLegacyTextEffects,
} from './work-document-legacy-text-effects';
import {
  parseDocumentCharacterFormatting,
  type DocumentCharacterFormatMark,
} from './work-document-format-changes';
import { DOCX_WORDPROCESSING_NAMESPACES } from './work-docx-ignorable-extension-preservation';
import { normalizeDocumentUnderlineStyle } from './work-document-underline';
import { normalizeDocumentStrikeStyle } from './work-document-strike';
import { parseDocumentScriptFonts } from './work-document-script-fonts';
import { parseDocxThemeReference } from './work-docx-theme-reference';
import { parseDocumentRunBorder } from './work-document-run-border';
import { parseDocumentRunShading } from './work-document-run-shading';
import {
  normalizeDocumentNoProof,
  parseDocumentProofingLanguages,
} from './work-document-proofing';
import { setDocxRunShadingAttributes } from './work-docx-run-shading-export';
import {
  documentHighlightForCssColor,
  normalizeDocumentHighlight,
} from './work-document-highlight';
import { descendants, directChildren, parseXml } from './work-ooxml-package';
import { decodeXmlBytes, serializeUtf8Xml } from './work-ooxml-xml';

interface DocxRunFormattingChangePatch {
  start: string;
  end: string;
  id: number;
  author: string;
  date: string;
  before: string;
}

const MAX_RUN_FORMATTING_CHANGE_PATCHES = 65_536;
const RUN_FORMATTING_PART_PATTERN =
  /^word\/(?:document|header\d*|footer\d*|footnotes|endnotes)\.xml$/i;

export class DocxRunFormattingChangePatchCollector {
  readonly patches: DocxRunFormattingChangePatch[] = [];

  register(
    element: HTMLElement,
    id: number,
  ): { start: string; end: string } | null {
    if (
      element.dataset.changeKind !== 'formatting' ||
      !element.hasAttribute('data-document-change')
    ) {
      return null;
    }
    const key = element.dataset.changeId?.trim() ?? '';
    const author = element.dataset.changeAuthor?.trim() ?? '';
    const date = normalizedRevisionDate(element.dataset.changeDate);
    const before = element.dataset.changeBefore ?? '';
    if (
      !key ||
      !author ||
      author.length > 255 ||
      !parseDocumentCharacterFormatting(before)
    ) {
      throw new Error(
        'Document contains an invalid character-formatting revision.',
      );
    }
    if (this.patches.length >= MAX_RUN_FORMATTING_CHANGE_PATCHES) {
      throw new Error(
        'Document exceeds the character-formatting revision limit.',
      );
    }
    const sequence = this.patches.length + 1;
    const start = `__A3S_WORK_FORMAT_CHANGE_START_${sequence}__`;
    const end = `__A3S_WORK_FORMAT_CHANGE_END_${sequence}__`;
    this.patches.push({ start, end, id, author, date, before });
    return { start, end };
  }
}

export async function patchDocxRunFormattingChanges(
  buffer: ArrayBuffer,
  patches: readonly DocxRunFormattingChangePatch[],
): Promise<ArrayBuffer> {
  if (!patches.length) return buffer;
  if (patches.length > MAX_RUN_FORMATTING_CHANGE_PATCHES) {
    throw new Error(
      'Document exceeds the character-formatting revision limit.',
    );
  }
  const archive = await JSZip.loadAsync(buffer);
  const byMarker = new Map<string, DocxRunFormattingChangePatch>();
  for (const patch of patches) {
    byMarker.set(patch.start, patch);
    byMarker.set(patch.end, patch);
  }
  const applied = new Set<DocxRunFormattingChangePatch>();
  for (const entry of Object.values(archive.files)) {
    if (entry.dir || !RUN_FORMATTING_PART_PATTERN.test(entry.name)) continue;
    const document = parseXml(
      decodeXmlBytes(
        await entry.async('uint8array'),
        `generated DOCX ${entry.name}`,
      ),
      `generated DOCX ${entry.name}`,
    );
    const runs = descendants(document, 'r').filter((run) =>
      DOCX_WORDPROCESSING_NAMESPACES.has(run.namespaceURI ?? ''),
    );
    const markerRuns = new Map<
      DocxRunFormattingChangePatch,
      { start?: Element; end?: Element }
    >();
    for (const run of runs) {
      const marker = runMarker(run, byMarker);
      if (!marker) continue;
      const current = markerRuns.get(marker.patch) ?? {};
      if (marker.kind === 'start') {
        if (current.start) throw duplicateMarker(marker.patch.start);
        current.start = run;
      } else {
        if (current.end) throw duplicateMarker(marker.patch.end);
        current.end = run;
      }
      markerRuns.set(marker.patch, current);
    }
    let changed = false;
    for (const [patch, markers] of markerRuns) {
      if (!markers.start || !markers.end) continue;
      const start = runs.indexOf(markers.start);
      const end = runs.indexOf(markers.end);
      if (start < 0 || end <= start + 1) {
        throw new Error(
          'Generated DOCX character-formatting markers are invalid.',
        );
      }
      const targets = runs
        .slice(start + 1, end)
        .filter((run) => directChildren(run, 't').length > 0);
      if (!targets.length) {
        throw new Error('A character-formatting revision must mark text.');
      }
      for (const run of targets) setRunFormattingChange(document, run, patch);
      markers.start.remove();
      markers.end.remove();
      applied.add(patch);
      changed = true;
    }
    if (changed) archive.file(entry.name, serializeUtf8Xml(document));
  }
  const missing = patches.filter((patch) => !applied.has(patch));
  if (missing.length) {
    throw new Error(
      `DOCX character-formatting revision markers were not emitted: ${missing
        .map((patch) => patch.start)
        .join(', ')}.`,
    );
  }
  return archive.generateAsync({ type: 'arraybuffer' });
}

function runMarker(
  run: Element,
  patches: ReadonlyMap<string, DocxRunFormattingChangePatch>,
): {
  patch: DocxRunFormattingChangePatch;
  kind: 'start' | 'end';
} | null {
  const texts = directChildren(run, 't').filter((text) =>
    DOCX_WORDPROCESSING_NAMESPACES.has(text.namespaceURI ?? ''),
  );
  if (texts.length !== 1 || !runHasOnlyMarkerText(run, texts[0])) return null;
  const value = texts[0].textContent ?? '';
  const patch = patches.get(value);
  if (!patch) return null;
  return { patch, kind: value === patch.start ? 'start' : 'end' };
}

function runHasOnlyMarkerText(run: Element, text: Element): boolean {
  return directChildren(run).every(
    (child) =>
      child === text ||
      (child.localName === 'rPr' &&
        DOCX_WORDPROCESSING_NAMESPACES.has(child.namespaceURI ?? '')),
  );
}

function setRunFormattingChange(
  document: Document,
  run: Element,
  patch: DocxRunFormattingChangePatch,
): void {
  const wordNamespace = run.namespaceURI ?? '';
  const prefix = run.prefix || 'w';
  const existing = directChildren(run, 'rPr').filter(
    (element) => element.namespaceURI === wordNamespace,
  );
  if (existing.length > 1) {
    throw new Error('Generated DOCX run contains duplicate properties.');
  }
  const properties =
    existing[0] ?? insertRunProperties(document, run, wordNamespace, prefix);
  if (
    directChildren(properties, 'rPrChange').some(
      (element) => element.namespaceURI === wordNamespace,
    )
  ) {
    throw new Error(
      'Generated DOCX run already contains a formatting revision.',
    );
  }
  const change = wordElement(document, wordNamespace, prefix, 'rPrChange');
  setWordAttribute(change, wordNamespace, prefix, 'id', String(patch.id));
  setWordAttribute(change, wordNamespace, prefix, 'author', patch.author);
  setWordAttribute(change, wordNamespace, prefix, 'date', patch.date);
  const before = wordElement(document, wordNamespace, prefix, 'rPr');
  appendFormattingProperties(
    document,
    before,
    wordNamespace,
    prefix,
    patch.before,
  );
  change.append(before);
  properties.append(change);
}

function insertRunProperties(
  document: Document,
  run: Element,
  namespace: string,
  prefix: string,
): Element {
  const properties = wordElement(document, namespace, prefix, 'rPr');
  run.insertBefore(properties, run.firstChild);
  return properties;
}

function appendFormattingProperties(
  document: Document,
  properties: Element,
  namespace: string,
  prefix: string,
  serialized: string,
): void {
  const formatting = parseDocumentCharacterFormatting(serialized);
  if (!formatting) {
    throw new Error(
      'Document contains an invalid character-formatting revision.',
    );
  }
  const byType = new Map(formatting.map((mark) => [mark.type, mark]));
  appendFontProperties(
    document,
    properties,
    namespace,
    prefix,
    byType.get('textStyle'),
  );
  appendBooleanProperty(
    document,
    properties,
    namespace,
    prefix,
    byType,
    'bold',
    'b',
  );
  appendBooleanProperty(
    document,
    properties,
    namespace,
    prefix,
    byType,
    'bold',
    'bCs',
  );
  appendBooleanProperty(
    document,
    properties,
    namespace,
    prefix,
    byType,
    'italic',
    'i',
  );
  appendBooleanProperty(
    document,
    properties,
    namespace,
    prefix,
    byType,
    'italic',
    'iCs',
  );
  appendTextCaseProperties(
    document,
    properties,
    namespace,
    prefix,
    byType.get('textStyle'),
  );
  appendStrikeProperties(
    document,
    properties,
    namespace,
    prefix,
    byType.get('strike'),
  );
  appendLegacyTextEffectsProperties(
    document,
    properties,
    namespace,
    prefix,
    byType.get('textStyle'),
  );
  appendNoProofProperty(
    document,
    properties,
    namespace,
    prefix,
    byType.get('textStyle'),
  );
  appendSnapToGridProperty(
    document,
    properties,
    namespace,
    prefix,
    byType.get('textStyle'),
  );
  appendHiddenTextProperty(
    document,
    properties,
    namespace,
    prefix,
    byType.get('textStyle'),
  );
  appendColorProperty(
    document,
    properties,
    namespace,
    prefix,
    byType.get('textStyle'),
  );
  appendCharacterSpacingProperty(
    document,
    properties,
    namespace,
    prefix,
    byType.get('textStyle'),
  );
  appendCharacterScaleProperty(
    document,
    properties,
    namespace,
    prefix,
    byType.get('textStyle'),
  );
  appendKerningThresholdProperty(
    document,
    properties,
    namespace,
    prefix,
    byType.get('textStyle'),
  );
  appendCharacterPositionProperty(
    document,
    properties,
    namespace,
    prefix,
    byType.get('textStyle'),
  );
  appendFontSizeProperties(
    document,
    properties,
    namespace,
    prefix,
    byType.get('textStyle'),
  );
  appendHighlightProperty(
    document,
    properties,
    namespace,
    prefix,
    byType.get('highlight'),
    !parseDocumentRunShading(byType.get('textStyle')?.attrs?.runShading),
  );
  appendUnderlineProperty(
    document,
    properties,
    namespace,
    prefix,
    byType.get('underline'),
  );
  appendRunBorderProperty(
    document,
    properties,
    namespace,
    prefix,
    byType.get('textStyle'),
  );
  appendRunShadingProperty(
    document,
    properties,
    namespace,
    prefix,
    byType.get('textStyle'),
  );
  appendVerticalAlignProperty(document, properties, namespace, prefix, byType);
  appendEmphasisMarkProperty(
    document,
    properties,
    namespace,
    prefix,
    byType.get('textStyle'),
  );
  appendProofingLanguagesProperty(
    document,
    properties,
    namespace,
    prefix,
    byType.get('textStyle'),
  );
}

function appendNoProofProperty(
  document: Document,
  properties: Element,
  namespace: string,
  prefix: string,
  mark: DocumentCharacterFormatMark | undefined,
): void {
  const noProof = normalizeDocumentNoProof(mark?.attrs?.noProof);
  if (noProof === null) return;
  if (noProof) {
    properties.append(wordElement(document, namespace, prefix, 'noProof'));
  } else {
    appendValuedProperty(
      document,
      properties,
      namespace,
      prefix,
      'noProof',
      '0',
    );
  }
}

function appendProofingLanguagesProperty(
  document: Document,
  properties: Element,
  namespace: string,
  prefix: string,
  mark: DocumentCharacterFormatMark | undefined,
): void {
  const languages = parseDocumentProofingLanguages(
    mark?.attrs?.proofingLanguages,
  );
  if (!languages) return;
  const element = wordElement(document, namespace, prefix, 'lang');
  if (languages.latin) {
    setWordAttribute(element, namespace, prefix, 'val', languages.latin);
  }
  if (languages.eastAsia) {
    setWordAttribute(
      element,
      namespace,
      prefix,
      'eastAsia',
      languages.eastAsia,
    );
  }
  if (languages.bidi) {
    setWordAttribute(element, namespace, prefix, 'bidi', languages.bidi);
  }
  properties.append(element);
}

function appendRunBorderProperty(
  document: Document,
  properties: Element,
  namespace: string,
  prefix: string,
  mark: DocumentCharacterFormatMark | undefined,
): void {
  const border = parseDocumentRunBorder(mark?.attrs?.runBorder);
  if (!border) return;
  const element = wordElement(document, namespace, prefix, 'bdr');
  setWordAttribute(element, namespace, prefix, 'val', border.style);
  if (border.color) {
    setWordAttribute(
      element,
      namespace,
      prefix,
      'color',
      border.color.value === 'auto'
        ? 'auto'
        : border.color.value.slice(1).toUpperCase(),
    );
    if (border.color.theme) {
      setWordAttribute(
        element,
        namespace,
        prefix,
        'themeColor',
        border.color.theme.theme,
      );
      if (border.color.theme.tint) {
        setWordAttribute(
          element,
          namespace,
          prefix,
          'themeTint',
          border.color.theme.tint,
        );
      }
      if (border.color.theme.shade) {
        setWordAttribute(
          element,
          namespace,
          prefix,
          'themeShade',
          border.color.theme.shade,
        );
      }
    }
  }
  if (border.size !== undefined) {
    setWordAttribute(element, namespace, prefix, 'sz', String(border.size));
  }
  if (border.space !== undefined) {
    setWordAttribute(element, namespace, prefix, 'space', String(border.space));
  }
  if (border.shadow !== undefined) {
    setWordAttribute(
      element,
      namespace,
      prefix,
      'shadow',
      border.shadow ? '1' : '0',
    );
  }
  if (border.frame !== undefined) {
    setWordAttribute(
      element,
      namespace,
      prefix,
      'frame',
      border.frame ? '1' : '0',
    );
  }
  properties.append(element);
}

function appendRunShadingProperty(
  document: Document,
  properties: Element,
  namespace: string,
  prefix: string,
  mark: DocumentCharacterFormatMark | undefined,
): void {
  const shading = parseDocumentRunShading(mark?.attrs?.runShading);
  if (!shading) return;
  const element = wordElement(document, namespace, prefix, 'shd');
  if (!setDocxRunShadingAttributes(document, element, shading)) return;
  properties.append(element);
}

function appendLegacyTextEffectsProperties(
  document: Document,
  properties: Element,
  namespace: string,
  prefix: string,
  mark: DocumentCharacterFormatMark | undefined,
): void {
  const effects: WorkDocumentLegacyTextEffects = {};
  const attributes = [
    ['outline', 'legacyTextOutline'],
    ['shadow', 'legacyTextShadow'],
    ['emboss', 'legacyTextEmboss'],
    ['imprint', 'legacyTextImprint'],
  ] as const;
  for (const [effect, attribute] of attributes) {
    const value = normalizeDocumentLegacyTextEffect(mark?.attrs?.[attribute]);
    if (value !== null) effects[effect] = value;
  }
  if (documentLegacyTextEffectsConflict(effects)) {
    throw new Error('Document contains conflicting legacy text effects.');
  }
  for (const name of DOCUMENT_LEGACY_TEXT_EFFECT_NAMES) {
    const value = effects[name];
    if (value === undefined) continue;
    if (value) {
      properties.append(wordElement(document, namespace, prefix, name));
    } else {
      appendValuedProperty(document, properties, namespace, prefix, name, '0');
    }
  }
}

function appendHiddenTextProperty(
  document: Document,
  properties: Element,
  namespace: string,
  prefix: string,
  mark: DocumentCharacterFormatMark | undefined,
): void {
  const hiddenText = normalizeDocumentHiddenText(mark?.attrs?.hiddenText);
  if (hiddenText === null) return;
  if (hiddenText) {
    properties.append(wordElement(document, namespace, prefix, 'vanish'));
  } else {
    appendValuedProperty(
      document,
      properties,
      namespace,
      prefix,
      'vanish',
      '0',
    );
  }
}

function appendEmphasisMarkProperty(
  document: Document,
  properties: Element,
  namespace: string,
  prefix: string,
  mark: DocumentCharacterFormatMark | undefined,
): void {
  const emphasisMark = normalizeDocumentEmphasisMark(mark?.attrs?.emphasisMark);
  if (emphasisMark === null) return;
  appendValuedProperty(
    document,
    properties,
    namespace,
    prefix,
    'em',
    emphasisMark,
  );
}

function appendCharacterScaleProperty(
  document: Document,
  properties: Element,
  namespace: string,
  prefix: string,
  mark: DocumentCharacterFormatMark | undefined,
): void {
  const scale = normalizeDocumentCharacterScalePercent(
    mark?.attrs?.characterScalePercent,
  );
  if (scale === null) return;
  appendValuedProperty(
    document,
    properties,
    namespace,
    prefix,
    'w',
    String(scale),
  );
}

function appendKerningThresholdProperty(
  document: Document,
  properties: Element,
  namespace: string,
  prefix: string,
  mark: DocumentCharacterFormatMark | undefined,
): void {
  const threshold = normalizeDocumentKerningThresholdHalfPoints(
    mark?.attrs?.kerningThresholdHalfPoints,
  );
  if (threshold === null) return;
  appendValuedProperty(
    document,
    properties,
    namespace,
    prefix,
    'kern',
    String(threshold),
  );
}

function appendCharacterPositionProperty(
  document: Document,
  properties: Element,
  namespace: string,
  prefix: string,
  mark: DocumentCharacterFormatMark | undefined,
): void {
  const position = normalizeDocumentCharacterPositionHalfPoints(
    mark?.attrs?.characterPositionHalfPoints,
  );
  if (position === null) return;
  appendValuedProperty(
    document,
    properties,
    namespace,
    prefix,
    'position',
    String(position),
  );
}

function appendCharacterSpacingProperty(
  document: Document,
  properties: Element,
  namespace: string,
  prefix: string,
  mark: DocumentCharacterFormatMark | undefined,
): void {
  const spacing = normalizeDocumentCharacterSpacingTwips(
    mark?.attrs?.characterSpacingTwips,
  );
  if (spacing === null) return;
  appendValuedProperty(
    document,
    properties,
    namespace,
    prefix,
    'spacing',
    String(spacing),
  );
}

function appendFontProperties(
  document: Document,
  properties: Element,
  namespace: string,
  prefix: string,
  mark: DocumentCharacterFormatMark | undefined,
): void {
  const scriptFonts = parseDocumentScriptFonts(
    typeof mark?.attrs?.scriptFonts === 'string'
      ? mark.attrs.scriptFonts
      : undefined,
  );
  if (scriptFonts) {
    const fonts = wordElement(document, namespace, prefix, 'rFonts');
    for (const [name, value] of [
      ['ascii', scriptFonts.ascii?.name],
      ['hAnsi', scriptFonts.highAnsi?.name],
      ['eastAsia', scriptFonts.eastAsia?.name],
      ['cs', scriptFonts.complexScript?.name],
      ['asciiTheme', scriptFonts.ascii?.theme],
      ['hAnsiTheme', scriptFonts.highAnsi?.theme],
      ['eastAsiaTheme', scriptFonts.eastAsia?.theme],
      ['cstheme', scriptFonts.complexScript?.theme],
      ['hint', scriptFonts.hint],
    ] as const) {
      if (value) setWordAttribute(fonts, namespace, prefix, name, value);
    }
    properties.append(fonts);
    return;
  }
  const family = firstFontFamily(mark?.attrs?.fontFamily);
  if (!family) return;
  const fonts = wordElement(document, namespace, prefix, 'rFonts');
  for (const name of ['ascii', 'hAnsi', 'eastAsia', 'cs']) {
    setWordAttribute(fonts, namespace, prefix, name, family);
  }
  properties.append(fonts);
}

function appendUnderlineProperty(
  document: Document,
  properties: Element,
  namespace: string,
  prefix: string,
  mark: DocumentCharacterFormatMark | undefined,
): void {
  if (!mark) return;
  const style =
    normalizeDocumentUnderlineStyle(mark.attrs?.underlineStyle) ?? 'single';
  const color = normalizedHex(mark.attrs?.underlineColor);
  const theme = themeReference(mark.attrs?.underlineThemeColor);
  const underline = wordElement(document, namespace, prefix, 'u');
  setWordAttribute(underline, namespace, prefix, 'val', style);
  const resolved = color ?? normalizedHex(theme?.resolved);
  if (resolved) {
    setWordAttribute(underline, namespace, prefix, 'color', resolved);
  }
  if (theme) {
    setThemeAttributes(underline, namespace, prefix, theme, 'themeColor');
  }
  properties.append(underline);
}

function appendBooleanProperty(
  document: Document,
  properties: Element,
  namespace: string,
  prefix: string,
  marks: ReadonlyMap<string, DocumentCharacterFormatMark>,
  mark: string,
  property: string,
): void {
  if (marks.has(mark)) {
    properties.append(wordElement(document, namespace, prefix, property));
  }
}

function appendStrikeProperties(
  document: Document,
  properties: Element,
  namespace: string,
  prefix: string,
  mark: DocumentCharacterFormatMark | undefined,
): void {
  if (!mark) return;
  const style =
    normalizeDocumentStrikeStyle(mark.attrs?.strikeStyle) ?? 'single';
  appendValuedProperty(
    document,
    properties,
    namespace,
    prefix,
    'strike',
    style === 'single' ? '1' : '0',
  );
  appendValuedProperty(
    document,
    properties,
    namespace,
    prefix,
    'dstrike',
    style === 'double' ? '1' : '0',
  );
}

function appendValuedProperty(
  document: Document,
  properties: Element,
  namespace: string,
  prefix: string,
  property: string,
  value: string,
): void {
  const element = wordElement(document, namespace, prefix, property);
  setWordAttribute(element, namespace, prefix, 'val', value);
  properties.append(element);
}

function appendTextCaseProperties(
  document: Document,
  properties: Element,
  namespace: string,
  prefix: string,
  mark: DocumentCharacterFormatMark | undefined,
): void {
  const textCase = mark?.attrs?.textCase;
  if (textCase === 'all-caps') {
    properties.append(wordElement(document, namespace, prefix, 'caps'));
  } else if (textCase === 'small-caps') {
    properties.append(wordElement(document, namespace, prefix, 'smallCaps'));
  } else if (textCase === 'none') {
    appendValuedProperty(document, properties, namespace, prefix, 'caps', '0');
    appendValuedProperty(
      document,
      properties,
      namespace,
      prefix,
      'smallCaps',
      '0',
    );
  }
}

function appendColorProperty(
  document: Document,
  properties: Element,
  namespace: string,
  prefix: string,
  mark: DocumentCharacterFormatMark | undefined,
): void {
  const color = normalizedHex(mark?.attrs?.color);
  const theme = themeReference(mark?.attrs?.themeColor);
  if (!color && !theme) return;
  const element = wordElement(document, namespace, prefix, 'color');
  setWordAttribute(
    element,
    namespace,
    prefix,
    'val',
    color ?? normalizedHex(theme?.resolved) ?? '000000',
  );
  if (theme)
    setThemeAttributes(element, namespace, prefix, theme, 'themeColor');
  properties.append(element);
}

function appendFontSizeProperties(
  document: Document,
  properties: Element,
  namespace: string,
  prefix: string,
  mark: DocumentCharacterFormatMark | undefined,
): void {
  const halfPoints = fontSizeHalfPoints(mark?.attrs?.fontSize);
  if (halfPoints === null) return;
  appendValuedProperty(
    document,
    properties,
    namespace,
    prefix,
    'sz',
    String(halfPoints),
  );
  appendValuedProperty(
    document,
    properties,
    namespace,
    prefix,
    'szCs',
    String(halfPoints),
  );
}

function appendHighlightProperty(
  document: Document,
  properties: Element,
  namespace: string,
  prefix: string,
  mark: DocumentCharacterFormatMark | undefined,
  allowFallbackShading: boolean,
): void {
  const nativeHighlight =
    normalizeDocumentHighlight(mark?.attrs?.nativeHighlight) ??
    documentHighlightForCssColor(mark?.attrs?.color);
  if (nativeHighlight) {
    appendValuedProperty(
      document,
      properties,
      namespace,
      prefix,
      'highlight',
      nativeHighlight,
    );
    return;
  }
  if (!allowFallbackShading) return;
  const fill = normalizedHex(mark?.attrs?.color);
  const theme = themeReference(mark?.attrs?.themeFill);
  if (!fill && !theme) return;
  const shading = wordElement(document, namespace, prefix, 'shd');
  setWordAttribute(shading, namespace, prefix, 'val', 'clear');
  setWordAttribute(shading, namespace, prefix, 'color', 'auto');
  setWordAttribute(
    shading,
    namespace,
    prefix,
    'fill',
    fill ?? normalizedHex(theme?.resolved) ?? 'FFFFFF',
  );
  if (theme) setThemeAttributes(shading, namespace, prefix, theme, 'themeFill');
  properties.append(shading);
}

function appendSnapToGridProperty(
  document: Document,
  properties: Element,
  namespace: string,
  prefix: string,
  mark: DocumentCharacterFormatMark | undefined,
): void {
  const value = mark?.attrs?.wordSnapToGrid;
  if (typeof value !== 'boolean') return;
  appendValuedProperty(
    document,
    properties,
    namespace,
    prefix,
    'snapToGrid',
    value ? '1' : '0',
  );
}

function appendVerticalAlignProperty(
  document: Document,
  properties: Element,
  namespace: string,
  prefix: string,
  marks: ReadonlyMap<string, DocumentCharacterFormatMark>,
): void {
  if (marks.has('subscript')) {
    appendValuedProperty(
      document,
      properties,
      namespace,
      prefix,
      'vertAlign',
      'subscript',
    );
  } else if (marks.has('superscript')) {
    appendValuedProperty(
      document,
      properties,
      namespace,
      prefix,
      'vertAlign',
      'superscript',
    );
  }
}

function setThemeAttributes(
  element: Element,
  namespace: string,
  prefix: string,
  theme: NonNullable<ReturnType<typeof parseDocxThemeReference>>,
  name: 'themeColor' | 'themeFill',
): void {
  setWordAttribute(element, namespace, prefix, name, theme.theme);
  if (theme.tint) {
    setWordAttribute(element, namespace, prefix, `${name}Tint`, theme.tint);
  }
  if (theme.shade) {
    setWordAttribute(element, namespace, prefix, `${name}Shade`, theme.shade);
  }
}

function themeReference(
  value: boolean | number | string | undefined,
): ReturnType<typeof parseDocxThemeReference> {
  return typeof value === 'string' ? parseDocxThemeReference(value) : null;
}

function firstFontFamily(
  value: boolean | number | string | undefined,
): string | null {
  if (typeof value !== 'string') return null;
  const source = value.trim();
  if (!source || source.length > 255) return null;
  const quote = source[0];
  if (quote === '"' || quote === "'") {
    const end = source.indexOf(quote, 1);
    if (end > 1) return source.slice(1, end).trim() || null;
  }
  return source.split(',')[0]?.trim() || null;
}

function fontSizeHalfPoints(
  value: boolean | number | string | undefined,
): number | null {
  if (typeof value !== 'string') return null;
  const match = /^(\d+(?:\.\d+)?)pt$/i.exec(value.trim());
  if (!match) return null;
  const points = Number(match[1]);
  if (!Number.isFinite(points) || points <= 0 || points > 512) return null;
  return Math.round(points * 2);
}

function normalizedHex(
  value: boolean | number | string | undefined,
): string | null {
  if (typeof value !== 'string') return null;
  const match = /^#?([0-9a-f]{6})$/i.exec(value.trim());
  return match?.[1]?.toUpperCase() ?? null;
}

function wordElement(
  document: Document,
  namespace: string,
  prefix: string,
  name: string,
): Element {
  return document.createElementNS(namespace, `${prefix}:${name}`);
}

function setWordAttribute(
  element: Element,
  namespace: string,
  prefix: string,
  name: string,
  value: string,
): void {
  element.setAttributeNS(namespace, `${prefix}:${name}`, value);
}

function normalizedRevisionDate(value: string | undefined): string {
  const time = Date.parse(value ?? '');
  return Number.isFinite(time)
    ? new Date(time).toISOString()
    : new Date().toISOString();
}

function duplicateMarker(marker: string): Error {
  return new Error(`Generated DOCX marker '${marker}' is duplicated.`);
}
