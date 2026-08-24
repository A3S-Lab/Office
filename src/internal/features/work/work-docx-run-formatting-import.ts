import {
  type DocxParagraphStyleSource,
  type DocxParagraphStyleResolver,
  docxRunPropertySources,
  resolveDocxParagraphStyleResolver,
} from './work-docx-paragraph-styles';
import {
  type DocxThemeSource,
  docxThemeColor,
  docxThemeFont,
  resolveDocxThemeResolver,
  type DocxThemeResolver,
} from './work-docx-theme';
import {
  type DocxTableStyleSource,
  docxTableRunPropertySources,
  resolveDocxTableStyleResolver,
} from './work-docx-table-styles';
import { attribute, descendants, directChild } from './work-ooxml-package';
import { documentCharacterScaleDomAttributes } from './work-document-character-scale';
import { documentCharacterPositionDomAttributes } from './work-document-character-position';
import { documentCharacterSpacingDomAttributes } from './work-document-character-spacing';
import {
  documentEmphasisMarkDomAttributes,
  type WorkDocumentEmphasisMark,
} from './work-document-emphasis';
import { documentKerningDomAttributes } from './work-document-kerning';
import { docxCharacterScalePercentFromProperties } from './work-docx-character-scale';
import { docxCharacterPositionHalfPointsFromProperties } from './work-docx-character-position';
import { docxCharacterSpacingTwipsFromProperties } from './work-docx-character-spacing';
import { resolveDocxKerningThresholdHalfPoints } from './work-docx-kerning';
import { resolveDocxEmphasisMark } from './work-docx-emphasis';
import { resolveDocxHiddenText } from './work-docx-hidden-text';
import { documentHiddenTextDomAttributes } from './work-document-hidden-text';
import {
  documentLegacyTextEffectsDomAttributes,
  type WorkDocumentLegacyTextEffects,
} from './work-document-legacy-text-effects';
import { resolveDocxLegacyTextEffects } from './work-docx-legacy-text-effects';
import { documentWordLineHeightFactor } from './work-document-word-line-metrics';
import {
  type DocxThemeColorReference,
  serializeDocxThemeReference,
} from './work-docx-theme-reference';
import { importedDocumentCharacterFormatting } from './work-document-format-changes';
import { DOCX_WORDPROCESSING_NAMESPACES } from './work-docx-ignorable-extension-preservation';
import {
  applyDocumentTextCaseStyle,
  documentTextCaseFromWordFlags,
  type WorkDocumentTextCase,
} from './work-document-text-case';
import {
  documentUnderlineDomAttributes,
  type WorkDocumentUnderlineFormatting,
} from './work-document-underline';
import { importedDocxUnderline } from './work-docx-underline';
import {
  documentStrikeDomAttributes,
  type WorkDocumentStrikeFormatting,
} from './work-document-strike';
import {
  importedDocxStrikeFlags,
  resolvedDocxStrikeFormatting,
} from './work-docx-strike';
import {
  documentScriptFontFallbackSlots,
  documentScriptFontFamily,
  documentScriptFontSegments,
  documentScriptFontsDomAttributes,
  type WorkDocumentScriptFontSlot,
  type WorkDocumentScriptFonts,
} from './work-document-script-fonts';
import { resolveDocxRunFonts } from './work-docx-run-fonts';
import {
  directDocxRunText,
  markSegmentedDocxRunFormatting,
} from './work-docx-run-script-font-import';
import {
  type DocumentRunBorder,
  documentRunBorderDomAttributes,
  serializeDocumentRunBorder,
} from './work-document-run-border';
import { resolveDocxRunBorder } from './work-docx-run-border';
import {
  type DocumentRunShading,
  documentRunShadingDomAttributes,
  serializeDocumentRunShading,
} from './work-document-run-shading';
import { resolveDocxRunShading } from './work-docx-run-shading';
import {
  documentHighlightCssColor,
  documentHighlightDomAttributes,
  documentHighlightFromDocxValue,
  type WorkDocumentHighlight,
} from './work-document-highlight';
import {
  documentProofingDomAttributes,
  type WorkDocumentProofingLanguages,
} from './work-document-proofing';
import { resolveDocxProofing } from './work-docx-proofing';

export interface ImportedDocxRunFormatting {
  bold?: boolean;
  italic?: boolean;
  underline?: WorkDocumentUnderlineFormatting;
  strike?: WorkDocumentStrikeFormatting;
  subscript?: boolean;
  superscript?: boolean;
  characterScalePercent?: number;
  characterPositionHalfPoints?: number;
  characterSpacingTwips?: number;
  kerningThresholdHalfPoints?: number;
  emphasisMark?: WorkDocumentEmphasisMark;
  hiddenText?: boolean;
  legacyTextOutline?: boolean;
  legacyTextShadow?: boolean;
  legacyTextEmboss?: boolean;
  legacyTextImprint?: boolean;
  fontFamily?: string;
  scriptFonts?: WorkDocumentScriptFonts;
  scriptFontSlot?: WorkDocumentScriptFontSlot;
  wordLineHeightFactor?: number;
  wordSnapToGrid?: boolean;
  fontSize?: number;
  color?: string;
  backgroundColor?: string;
  highlight?: WorkDocumentHighlight;
  themeColor?: DocxThemeColorReference;
  themeFill?: DocxThemeColorReference;
  textCase?: WorkDocumentTextCase;
  runBorder?: DocumentRunBorder;
  runShading?: DocumentRunShading;
  proofingLanguages?: WorkDocumentProofingLanguages;
  noProof?: boolean;
}

export interface ImportedDocxRunFormattingMarker {
  startMarker: string;
  endMarker: string;
  formatting: ImportedDocxRunFormatting;
  change?: ImportedDocxRunFormattingChange;
}

export interface ImportedDocxRunFormattingChange {
  id: string;
  author: string;
  date: string;
  before: string;
}

export interface ImportedDocxRunFormattingMarkers {
  runs: ImportedDocxRunFormattingMarker[];
}

export interface ImportedDocxRunFormattingMarkerState {
  markers: ImportedDocxRunFormattingMarkers;
  nextMarker: number;
  occupiedText: string;
}

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const XML_NAMESPACE = 'http://www.w3.org/XML/1998/namespace';
const RUN_FORMATTING_MARKER_PATTERN = /__A3S_WORK_RUN_(?:START|END)_\d+__/g;
const SUPPORTED_RUN_PROPERTY_CHANGE_CHILDREN = new Set([
  'b',
  'bCs',
  'i',
  'iCs',
  'u',
  'strike',
  'dstrike',
  'outline',
  'shadow',
  'emboss',
  'imprint',
  'noProof',
  'caps',
  'smallCaps',
  'vanish',
  'spacing',
  'w',
  'kern',
  'position',
  'rFonts',
  'sz',
  'szCs',
  'color',
  'highlight',
  'shd',
  'snapToGrid',
  'cs',
  'rtl',
  'em',
  'lang',
  'vertAlign',
  'bdr',
]);

export function markDocxRunFormatting(
  document: Document,
  styleSource?: DocxParagraphStyleSource,
  themeSource?: DocxThemeSource,
  tableStyleSource?: DocxTableStyleSource,
): ImportedDocxRunFormattingMarkers {
  const state = createImportedDocxRunFormattingMarkerState([document]);
  markDocxRunFormattingIntoState(
    document,
    state,
    styleSource,
    themeSource,
    tableStyleSource,
  );
  return state.markers;
}

export function createImportedDocxRunFormattingMarkerState(
  documents: readonly Document[],
): ImportedDocxRunFormattingMarkerState {
  return {
    markers: { runs: [] },
    nextMarker: 1,
    occupiedText: documents
      .map((document) => document.documentElement.textContent ?? '')
      .join(''),
  };
}

export function markDocxRunFormattingIntoState(
  document: Document,
  state: ImportedDocxRunFormattingMarkerState,
  styleSource?: DocxParagraphStyleSource,
  themeSource?: DocxThemeSource,
  tableStyleSource?: DocxTableStyleSource,
): void {
  const styles = resolveDocxParagraphStyleResolver(styleSource);
  const theme = resolveDocxThemeResolver(themeSource);
  const tableStyles = resolveDocxTableStyleResolver(tableStyleSource);
  for (const run of descendants(document, 'r')) {
    const runText = directDocxRunText(run);
    if (!runText) continue;
    const paragraph = closestAncestor(run, 'p');
    if (!paragraph) continue;
    const paragraphProperties = directChild(paragraph, 'pPr');
    const runProperties = directChild(run, 'rPr');
    const contextualProperties = docxTableRunPropertySources(run, tableStyles);
    const propertySources = docxRunPropertySources(
      paragraphProperties,
      runProperties,
      styles,
      contextualProperties,
    );
    const scriptFonts = resolveDocxRunFonts(propertySources, theme);
    const forceComplexScript = propertySourcesUseComplexScript(propertySources);
    const scriptSegments = documentScriptFontSegments(
      runText,
      scriptFonts?.hint,
      forceComplexScript,
    );
    const formattingFor = (
      text: string,
      scriptFontSlot: WorkDocumentScriptFontSlot | undefined,
    ) =>
      resolvedRunFormatting(
        propertySources,
        theme,
        text,
        scriptFontSlot,
        scriptFonts,
      );
    const changeFor = (
      text: string,
      scriptFontSlot: WorkDocumentScriptFontSlot | undefined,
    ) =>
      importedRunFormattingChange(
        runProperties,
        paragraphProperties,
        styles,
        contextualProperties,
        theme,
        text,
        scriptFontSlot,
      );
    if (scriptSegments.length > 1) {
      markSegmentedDocxRunFormatting(
        run,
        state,
        scriptSegments,
        formattingFor,
        changeFor,
        () => nextRunFormattingMarkers(state),
      );
      continue;
    }
    const scriptFontSlot = scriptSegments[0]?.slot;
    const formatting = formattingFor(runText, scriptFontSlot);
    const change = changeFor(runText, scriptFontSlot);
    if (!Object.keys(formatting).length && !change) continue;
    const { startMarker, endMarker } = nextRunFormattingMarkers(state);
    insertRunMarkers(document, run, runProperties, startMarker, endMarker);
    state.markers.runs.push({
      startMarker,
      endMarker,
      formatting,
      ...(change ? { change } : {}),
    });
  }
}

export function applyImportedDocxRunFormattingMarkers(
  document: Document,
  markers: ImportedDocxRunFormattingMarkers,
): void {
  const replacements = new Map<string, string>();
  for (const marker of markers.runs) {
    const markup = formattingMarkup(document, marker.formatting);
    const change = marker.change
      ? formattingChangeMarkup(document, marker.change)
      : { start: '', end: '' };
    replacements.set(marker.startMarker, `${change.start}${markup.start}`);
    replacements.set(marker.endMarker, `${markup.end}${change.end}`);
  }
  document.body.innerHTML = document.body.innerHTML.replace(
    RUN_FORMATTING_MARKER_PATTERN,
    (marker) => replacements.get(marker) ?? marker,
  );
}

function nextRunFormattingMarkers(
  state: ImportedDocxRunFormattingMarkerState,
): { startMarker: string; endMarker: string } {
  while (true) {
    const index = state.nextMarker;
    state.nextMarker += 1;
    const startMarker = `__A3S_WORK_RUN_START_${index}__`;
    const endMarker = `__A3S_WORK_RUN_END_${index}__`;
    if (
      !state.occupiedText.includes(startMarker) &&
      !state.occupiedText.includes(endMarker)
    ) {
      return { startMarker, endMarker };
    }
  }
}

export function hasImportedDocxRunFormattingMarkers(
  markers: ImportedDocxRunFormattingMarkers,
): boolean {
  return markers.runs.length > 0;
}

function resolvedRunFormatting(
  propertySources: readonly Element[],
  theme: DocxThemeResolver,
  runText: string,
  requestedFontSlot?: WorkDocumentScriptFontSlot,
  resolvedScriptFonts: WorkDocumentScriptFonts | null = resolveDocxRunFonts(
    propertySources,
    theme,
  ),
): ImportedDocxRunFormatting {
  let fontSize: number | undefined;
  let complexFontSize: number | undefined;
  let color: string | undefined;
  let backgroundColor: string | undefined;
  let highlightValue: WorkDocumentHighlight | undefined;
  let themeColor: DocxThemeColorReference | undefined;
  let themeFill: DocxThemeColorReference | undefined;
  let bold: boolean | undefined;
  let complexBold: boolean | undefined;
  let italic: boolean | undefined;
  let complexItalic: boolean | undefined;
  let underline: WorkDocumentUnderlineFormatting | undefined;
  let strike: boolean | undefined;
  let doubleStrike: boolean | undefined;
  let subscript: boolean | undefined;
  let superscript: boolean | undefined;
  let snapToGrid: boolean | undefined;
  let complexScriptFormatting: boolean | undefined;
  let rightToLeft: boolean | undefined;
  let allCaps: boolean | undefined;
  let smallCaps: boolean | undefined;
  let characterScalePercent: number | undefined;
  let characterPositionHalfPoints: number | undefined;
  let characterSpacingTwips: number | undefined;
  const kerningThresholdHalfPoints =
    resolveDocxKerningThresholdHalfPoints(propertySources);
  const emphasisMark = resolveDocxEmphasisMark(propertySources);
  const hiddenText = resolveDocxHiddenText(propertySources);
  const legacyTextEffects = resolveDocxLegacyTextEffects(propertySources);
  const runBorder = resolveDocxRunBorder(propertySources, theme).border;
  const runShading = resolveDocxRunShading(propertySources, theme).shading;
  const proofing = resolveDocxProofing(propertySources);

  for (const properties of propertySources) {
    bold = overriddenBoolean(bold, onOffProperty(properties, 'b'));
    complexBold = overriddenBoolean(
      complexBold,
      onOffProperty(properties, 'bCs'),
    );
    italic = overriddenBoolean(italic, onOffProperty(properties, 'i'));
    complexItalic = overriddenBoolean(
      complexItalic,
      onOffProperty(properties, 'iCs'),
    );
    underline = overriddenValue(
      underline,
      underlineProperty(properties, theme),
    );
    const importedStrike = importedDocxStrikeFlags(properties);
    strike = overriddenBoolean(strike, importedStrike.single);
    doubleStrike = overriddenBoolean(doubleStrike, importedStrike.double);
    allCaps = overriddenBoolean(allCaps, onOffProperty(properties, 'caps'));
    smallCaps = overriddenBoolean(
      smallCaps,
      onOffProperty(properties, 'smallCaps'),
    );
    characterSpacingTwips = overriddenValue(
      characterSpacingTwips,
      docxCharacterSpacingTwipsFromProperties(properties),
    );
    characterScalePercent = overriddenValue(
      characterScalePercent,
      docxCharacterScalePercentFromProperties(properties),
    );
    characterPositionHalfPoints = overriddenValue(
      characterPositionHalfPoints,
      docxCharacterPositionHalfPointsFromProperties(properties),
    );
    const verticalAlign = directChild(properties, 'vertAlign');
    if (verticalAlign) {
      const value = wordAttribute(verticalAlign, 'val')?.trim().toLowerCase();
      subscript = value === 'subscript';
      superscript = value === 'superscript';
      if (value === 'baseline') {
        subscript = false;
        superscript = false;
      }
    }
    snapToGrid = overriddenBoolean(
      snapToGrid,
      onOffProperty(properties, 'snapToGrid'),
    );
    complexScriptFormatting = overriddenBoolean(
      complexScriptFormatting,
      onOffProperty(properties, 'cs'),
    );
    rightToLeft = overriddenBoolean(
      rightToLeft,
      onOffProperty(properties, 'rtl'),
    );

    const size = numericAttribute(directChild(properties, 'sz'), 'val');
    if (size !== undefined && size > 0) fontSize = Math.min(512, size / 2);
    const complexSize = numericAttribute(
      directChild(properties, 'szCs'),
      'val',
    );
    if (complexSize !== undefined && complexSize > 0) {
      complexFontSize = Math.min(512, complexSize / 2);
    }

    const colorElement = directChild(properties, 'color');
    if (colorElement) {
      const value = wordAttribute(colorElement, 'val')?.trim();
      const themed = docxThemeColor(
        theme,
        wordAttribute(colorElement, 'themeColor'),
        wordAttribute(colorElement, 'themeTint'),
        wordAttribute(colorElement, 'themeShade'),
      );
      if (themed) {
        color = `#${themed}`;
        themeColor = themeReference(
          colorElement,
          'themeColor',
          'themeTint',
          'themeShade',
          color,
        );
      } else if (value?.toLowerCase() === 'auto') {
        color = 'inherit';
        themeColor = undefined;
      } else if (value && /^[0-9a-f]{6}$/i.test(value)) {
        color = `#${value.toLowerCase()}`;
        themeColor = undefined;
      }
    }

    const highlight = directChild(properties, 'highlight');
    if (highlight) {
      const value = documentHighlightFromDocxValue(
        wordAttribute(highlight, 'val'),
      );
      if (value) {
        highlightValue = value;
        backgroundColor = documentHighlightCssColor(value) ?? 'transparent';
        themeFill = undefined;
      }
    }
  }

  const fontSlot =
    complexScriptFormatting === true || rightToLeft === true
      ? 'complexScript'
      : (requestedFontSlot ??
        documentScriptFontSegments(runText, resolvedScriptFonts?.hint)[0]
          ?.slot ??
        'ascii');
  const usesComplexFormatting = fontSlot === 'complexScript';
  const resolvedBold = usesComplexFormatting ? (complexBold ?? bold) : bold;
  const resolvedItalic = usesComplexFormatting
    ? (complexItalic ?? italic)
    : italic;
  const resolvedFontSize = usesComplexFormatting
    ? (complexFontSize ?? fontSize)
    : fontSize;
  const fontFamily =
    documentScriptFontFamily(resolvedScriptFonts, fontSlot) ??
    defaultThemeFontFamily(theme, fontSlot);
  const hasRunPropertySource = propertySources.length > 0;
  const resolvedStrike = resolvedDocxStrikeFormatting(
    strike,
    doubleStrike,
    hasRunPropertySource,
  );
  const textCase = documentTextCaseFromWordFlags(allCaps, smallCaps);
  return {
    ...(resolvedBold !== undefined || hasRunPropertySource
      ? { bold: resolvedBold ?? false }
      : {}),
    ...(resolvedItalic !== undefined || hasRunPropertySource
      ? { italic: resolvedItalic ?? false }
      : {}),
    ...(underline !== undefined || hasRunPropertySource
      ? { underline: underline ?? { style: 'none' } }
      : {}),
    ...(resolvedStrike ? { strike: resolvedStrike } : {}),
    ...(subscript !== undefined ? { subscript } : {}),
    ...(superscript !== undefined ? { superscript } : {}),
    ...(characterScalePercent !== undefined ? { characterScalePercent } : {}),
    ...(characterPositionHalfPoints !== undefined
      ? { characterPositionHalfPoints }
      : {}),
    ...(characterSpacingTwips !== undefined ? { characterSpacingTwips } : {}),
    ...(kerningThresholdHalfPoints !== undefined
      ? { kerningThresholdHalfPoints }
      : {}),
    ...(emphasisMark !== undefined ? { emphasisMark } : {}),
    ...(hiddenText !== undefined ? { hiddenText } : {}),
    ...(legacyTextEffects?.outline !== undefined
      ? { legacyTextOutline: legacyTextEffects.outline }
      : {}),
    ...(legacyTextEffects?.shadow !== undefined
      ? { legacyTextShadow: legacyTextEffects.shadow }
      : {}),
    ...(legacyTextEffects?.emboss !== undefined
      ? { legacyTextEmboss: legacyTextEffects.emboss }
      : {}),
    ...(legacyTextEffects?.imprint !== undefined
      ? { legacyTextImprint: legacyTextEffects.imprint }
      : {}),
    ...(fontFamily
      ? {
          fontFamily,
          wordLineHeightFactor: documentWordLineHeightFactor(fontFamily),
        }
      : {}),
    ...(resolvedScriptFonts
      ? { scriptFonts: resolvedScriptFonts, scriptFontSlot: fontSlot }
      : {}),
    ...(resolvedFontSize !== undefined ? { fontSize: resolvedFontSize } : {}),
    ...(snapToGrid !== undefined ? { wordSnapToGrid: snapToGrid } : {}),
    ...(color ? { color } : {}),
    ...(backgroundColor ? { backgroundColor } : {}),
    ...(highlightValue ? { highlight: highlightValue } : {}),
    ...(themeColor ? { themeColor } : {}),
    ...(themeFill ? { themeFill } : {}),
    ...(textCase ? { textCase } : {}),
    ...(runBorder ? { runBorder } : {}),
    ...(runShading ? { runShading } : {}),
    ...(proofing.languages ? { proofingLanguages: proofing.languages } : {}),
    ...(proofing.noProof !== undefined ? { noProof: proofing.noProof } : {}),
  };
}

function insertRunMarkers(
  document: Document,
  run: Element,
  properties: Element | undefined,
  startMarker: string,
  endMarker: string,
): void {
  const start = markerText(document, startMarker);
  const end = markerText(document, endMarker);
  run.insertBefore(start, properties?.nextSibling ?? run.firstChild);
  run.append(end);
}

function markerText(document: Document, marker: string): Element {
  const text = document.createElementNS(WORD_NAMESPACE, 'w:t');
  text.setAttributeNS(XML_NAMESPACE, 'xml:space', 'preserve');
  text.textContent = marker;
  return text;
}

function formattingMarkup(
  document: Document,
  formatting: ImportedDocxRunFormatting,
): { start: string; end: string } {
  const span = document.createElement('span');
  if (formatting.bold === false) span.style.fontWeight = 'normal';
  if (formatting.italic === false) span.style.fontStyle = 'normal';
  const underlineEnabled =
    formatting.underline !== undefined && formatting.underline.style !== 'none';
  const strikeEnabled =
    formatting.strike !== undefined && formatting.strike.style !== 'none';
  if (
    (formatting.underline?.style === 'none' ||
      formatting.strike?.style === 'none') &&
    !underlineEnabled &&
    !strikeEnabled
  ) {
    span.style.textDecorationLine = 'none';
  }
  if (formatting.fontFamily) span.style.fontFamily = formatting.fontFamily;
  if (formatting.scriptFonts) {
    for (const [name, value] of Object.entries(
      documentScriptFontsDomAttributes(
        formatting.scriptFonts,
        formatting.scriptFontSlot,
      ),
    )) {
      if (name === 'style') span.style.cssText += `; ${value}`;
      else span.setAttribute(name, value);
    }
  }
  if (formatting.characterSpacingTwips !== undefined) {
    for (const [name, value] of Object.entries(
      documentCharacterSpacingDomAttributes(formatting.characterSpacingTwips),
    )) {
      if (name === 'style') span.style.cssText += `; ${value}`;
      else span.setAttribute(name, value);
    }
  }
  if (formatting.characterScalePercent !== undefined) {
    for (const [name, value] of Object.entries(
      documentCharacterScaleDomAttributes(formatting.characterScalePercent),
    )) {
      if (name === 'style') span.style.cssText += `; ${value}`;
      else span.setAttribute(name, value);
    }
  }
  if (formatting.kerningThresholdHalfPoints !== undefined) {
    for (const [name, value] of Object.entries(
      documentKerningDomAttributes(
        formatting.kerningThresholdHalfPoints,
        formatting.fontSize,
      ),
    )) {
      if (name === 'style') span.style.cssText += `; ${value}`;
      else span.setAttribute(name, value);
    }
  }
  if (formatting.characterPositionHalfPoints !== undefined) {
    for (const [name, value] of Object.entries(
      documentCharacterPositionDomAttributes(
        formatting.characterPositionHalfPoints,
      ),
    )) {
      if (name === 'style') span.style.cssText += `; ${value}`;
      else span.setAttribute(name, value);
    }
  }
  if (formatting.emphasisMark !== undefined) {
    for (const [name, value] of Object.entries(
      documentEmphasisMarkDomAttributes(formatting.emphasisMark),
    )) {
      if (name === 'style') span.style.cssText += `; ${value}`;
      else span.setAttribute(name, value);
    }
  }
  if (formatting.hiddenText !== undefined) {
    for (const [name, value] of Object.entries(
      documentHiddenTextDomAttributes(formatting.hiddenText),
    )) {
      span.setAttribute(name, value);
    }
  }
  if (formatting.runBorder) {
    for (const [name, value] of Object.entries(
      documentRunBorderDomAttributes(formatting.runBorder),
    )) {
      if (name === 'style') span.style.cssText += `; ${value}`;
      else span.setAttribute(name, value);
    }
  }
  if (formatting.runShading) {
    for (const [name, value] of Object.entries(
      documentRunShadingDomAttributes(formatting.runShading),
    )) {
      if (name === 'style') span.style.cssText += `; ${value}`;
      else span.setAttribute(name, value);
    }
  }
  if (
    formatting.proofingLanguages !== undefined ||
    formatting.noProof !== undefined
  ) {
    for (const [name, value] of Object.entries(
      documentProofingDomAttributes(
        formatting.proofingLanguages,
        formatting.noProof,
        formatting.scriptFontSlot,
      ),
    )) {
      span.setAttribute(name, value);
    }
  }
  const legacyTextEffects: WorkDocumentLegacyTextEffects = {
    ...(formatting.legacyTextOutline !== undefined
      ? { outline: formatting.legacyTextOutline }
      : {}),
    ...(formatting.legacyTextShadow !== undefined
      ? { shadow: formatting.legacyTextShadow }
      : {}),
    ...(formatting.legacyTextEmboss !== undefined
      ? { emboss: formatting.legacyTextEmboss }
      : {}),
    ...(formatting.legacyTextImprint !== undefined
      ? { imprint: formatting.legacyTextImprint }
      : {}),
  };
  for (const [name, value] of Object.entries(
    documentLegacyTextEffectsDomAttributes(legacyTextEffects),
  )) {
    span.setAttribute(name, value);
  }
  if (formatting.wordLineHeightFactor !== undefined) {
    const factor = formatLineHeightFactor(formatting.wordLineHeightFactor);
    span.dataset.officeWordLineHeightFactor = factor;
    span.style.setProperty('--work-document-word-line-height-factor', factor);
  }
  if (formatting.wordSnapToGrid !== undefined) {
    span.dataset.officeWordSnapToGrid = String(formatting.wordSnapToGrid);
  }
  if (formatting.fontSize !== undefined)
    span.style.fontSize = `${formatNumber(formatting.fontSize)}pt`;
  if (formatting.color) span.style.color = formatting.color;
  const themeColor = serializeDocxThemeReference(formatting.themeColor ?? null);
  if (themeColor) span.dataset.officeThemeColor = themeColor;
  if (formatting.highlight) {
    for (const [name, value] of Object.entries(
      documentHighlightDomAttributes(formatting.highlight),
    )) {
      if (name === 'style') span.style.cssText += `; ${value}`;
      else span.setAttribute(name, value);
    }
  } else if (formatting.backgroundColor) {
    span.style.backgroundColor = formatting.backgroundColor;
  }
  const themeFill = serializeDocxThemeReference(formatting.themeFill ?? null);
  if (themeFill) span.dataset.officeThemeFill = themeFill;
  if (formatting.textCase) {
    applyDocumentTextCaseStyle(span, formatting.textCase);
  }
  const html = span.outerHTML;
  const wrappers = [
    ...(formatting.bold ? [{ start: '<strong>', end: '</strong>' }] : []),
    ...(formatting.italic ? [{ start: '<em>', end: '</em>' }] : []),
    ...(formatting.underline
      ? [underlineFormattingWrapper(document, formatting.underline)]
      : []),
    ...(formatting.strike
      ? [strikeFormattingWrapper(document, formatting.strike)]
      : []),
    ...(formatting.subscript ? [{ start: '<sub>', end: '</sub>' }] : []),
    ...(formatting.superscript ? [{ start: '<sup>', end: '</sup>' }] : []),
  ];
  return {
    start: `${html.slice(0, html.indexOf('>') + 1)}${wrappers
      .map((wrapper) => wrapper.start)
      .join('')}`,
    end: `${[...wrappers]
      .reverse()
      .map((wrapper) => wrapper.end)
      .join('')}</span>`,
  };
}

function underlineFormattingWrapper(
  document: Document,
  formatting: WorkDocumentUnderlineFormatting,
): { start: string; end: string } {
  const underline = document.createElement('u');
  for (const [name, value] of Object.entries(
    documentUnderlineDomAttributes(formatting),
  )) {
    underline.setAttribute(name, value);
  }
  const html = underline.outerHTML;
  return {
    start: html.slice(0, html.indexOf('>') + 1),
    end: '</u>',
  };
}

function strikeFormattingWrapper(
  document: Document,
  formatting: WorkDocumentStrikeFormatting,
): { start: string; end: string } {
  const strike = document.createElement('s');
  for (const [name, value] of Object.entries(
    documentStrikeDomAttributes(formatting),
  )) {
    strike.setAttribute(name, value);
  }
  const html = strike.outerHTML;
  return {
    start: html.slice(0, html.indexOf('>') + 1),
    end: '</s>',
  };
}

function importedRunFormattingChange(
  runProperties: Element | undefined,
  paragraphProperties: Element | undefined,
  styles: DocxParagraphStyleResolver,
  contextualProperties: readonly Element[],
  theme: DocxThemeResolver,
  runText: string,
  requestedFontSlot?: WorkDocumentScriptFontSlot,
): ImportedDocxRunFormattingChange | undefined {
  const parsed = supportedRunFormattingChange(runProperties);
  if (!parsed) return undefined;
  const beforeFormatting = resolvedRunFormatting(
    docxRunPropertySources(
      paragraphProperties,
      parsed.properties,
      styles,
      contextualProperties,
    ),
    theme,
    runText,
    requestedFontSlot,
  );
  return {
    id: `docx-format-change-${parsed.id}`,
    author: parsed.author,
    date: normalizeRevisionDate(parsed.date),
    before: importedDocumentCharacterFormatting({
      ...beforeFormatting,
      themeColor: serializeDocxThemeReference(
        beforeFormatting.themeColor ?? null,
      ),
      themeFill: serializeDocxThemeReference(
        beforeFormatting.themeFill ?? null,
      ),
      scriptFonts: beforeFormatting.scriptFonts,
      scriptFontSlot: beforeFormatting.scriptFontSlot,
      textCase: beforeFormatting.textCase,
      legacyTextOutline: beforeFormatting.legacyTextOutline,
      legacyTextShadow: beforeFormatting.legacyTextShadow,
      legacyTextEmboss: beforeFormatting.legacyTextEmboss,
      legacyTextImprint: beforeFormatting.legacyTextImprint,
      runBorder: serializeDocumentRunBorder(beforeFormatting.runBorder),
      runShading: serializeDocumentRunShading(beforeFormatting.runShading),
      highlight: beforeFormatting.highlight,
    }),
  };
}

export function isSupportedDocxRunFormattingChange(change: Element): boolean {
  return Boolean(supportedRunFormattingChangeElement(change));
}

function supportedRunFormattingChange(runProperties: Element | undefined): {
  id: string;
  author: string;
  date: string | null;
  properties: Element;
} | null {
  if (!runProperties) return null;
  const changes = Array.from(runProperties.children).filter(
    (child) => child.localName === 'rPrChange',
  );
  return changes.length === 1
    ? supportedRunFormattingChangeElement(changes[0])
    : null;
}

function supportedRunFormattingChangeElement(change: Element): {
  id: string;
  author: string;
  date: string | null;
  properties: Element;
} | null {
  if (
    change.localName !== 'rPrChange' ||
    !DOCX_WORDPROCESSING_NAMESPACES.has(change.namespaceURI ?? '')
  ) {
    return null;
  }
  const properties = Array.from(change.children).filter(
    (child) =>
      child.localName === 'rPr' &&
      DOCX_WORDPROCESSING_NAMESPACES.has(child.namespaceURI ?? ''),
  );
  if (properties.length !== 1 || change.children.length !== 1) return null;
  const source = properties[0];
  const names = new Set<string>();
  for (const child of Array.from(source.children)) {
    if (
      !DOCX_WORDPROCESSING_NAMESPACES.has(child.namespaceURI ?? '') ||
      !SUPPORTED_RUN_PROPERTY_CHANGE_CHILDREN.has(child.localName) ||
      names.has(child.localName)
    ) {
      return null;
    }
    names.add(child.localName);
  }
  const id = attribute(change, 'id')?.trim() ?? '';
  const author = attribute(change, 'author')?.trim() ?? '';
  const date = attribute(change, 'date');
  if (!/^\+?\d{1,10}$/.test(id) || !author || author.length > 255) return null;
  if (date && !Number.isFinite(Date.parse(date))) return null;
  return { id, author, date, properties: source };
}

function formattingChangeMarkup(
  document: Document,
  change: ImportedDocxRunFormattingChange,
): { start: string; end: string } {
  const span = document.createElement('span');
  span.dataset.documentChange = 'true';
  span.dataset.changeKind = 'formatting';
  span.dataset.changeId = change.id;
  span.dataset.changeAuthor = change.author;
  span.dataset.changeDate = change.date;
  span.dataset.changeBefore = change.before;
  const html = span.outerHTML;
  return {
    start: html.slice(0, html.indexOf('>') + 1),
    end: '</span>',
  };
}

function normalizeRevisionDate(value: string | null): string {
  if (!value) return '';
  const time = Date.parse(value);
  return Number.isFinite(time) ? new Date(time).toISOString() : '';
}

function themeReference(
  element: Element,
  themeName: string,
  tintName: string,
  shadeName: string,
  resolved: string,
): DocxThemeColorReference | undefined {
  const theme = wordAttribute(element, themeName)?.trim();
  if (!theme) return undefined;
  const tint = normalizedByteHex(wordAttribute(element, tintName));
  const shade = normalizedByteHex(wordAttribute(element, shadeName));
  return {
    theme,
    resolved,
    ...(tint ? { tint } : {}),
    ...(shade ? { shade } : {}),
  };
}

function normalizedByteHex(value: string | null): string | undefined {
  const normalized = value?.trim().toUpperCase();
  return normalized && /^[0-9A-F]{2}$/.test(normalized)
    ? normalized
    : undefined;
}

function overriddenBoolean(
  current: boolean | undefined,
  next: boolean | undefined,
): boolean | undefined {
  return next === undefined ? current : next;
}

function overriddenValue<T>(
  current: T | undefined,
  next: T | undefined,
): T | undefined {
  return next === undefined ? current : next;
}

function propertySourcesUseComplexScript(
  propertySources: readonly Element[],
): boolean {
  let complexScript: boolean | undefined;
  let rightToLeft: boolean | undefined;
  for (const properties of propertySources) {
    complexScript = overriddenBoolean(
      complexScript,
      onOffProperty(properties, 'cs'),
    );
    rightToLeft = overriddenBoolean(
      rightToLeft,
      onOffProperty(properties, 'rtl'),
    );
  }
  return complexScript === true || rightToLeft === true;
}

function onOffProperty(
  properties: Element,
  propertyName: string,
): boolean | undefined {
  const element = directChild(properties, propertyName);
  if (!element) return undefined;
  const value = wordAttribute(element, 'val')?.trim().toLowerCase();
  return value !== '0' && value !== 'false' && value !== 'off';
}

function underlineProperty(
  properties: Element,
  theme: DocxThemeResolver,
): WorkDocumentUnderlineFormatting | undefined {
  const element = directChild(properties, 'u');
  if (!element) return undefined;
  return importedDocxUnderline(element, theme);
}

function closestAncestor(element: Element, localName: string): Element | null {
  let current = element.parentElement;
  while (current) {
    if (current.localName === localName) return current;
    current = current.parentElement;
  }
  return null;
}

function defaultThemeFontFamily(
  theme: DocxThemeResolver,
  preferred: WorkDocumentScriptFontSlot,
): string | undefined {
  const fonts: Record<WorkDocumentScriptFontSlot, string | undefined> = {
    ascii: docxThemeFont(theme, 'minorAscii'),
    highAnsi: docxThemeFont(theme, 'minorHAnsi'),
    eastAsia: docxThemeFont(theme, 'minorEastAsia'),
    complexScript: docxThemeFont(theme, 'minorBidi'),
  };
  return uniqueFonts(
    documentScriptFontFallbackSlots(preferred).map((slot) => fonts[slot]),
  );
}

function uniqueFonts(values: Array<string | undefined>): string | undefined {
  const fonts = [
    ...new Set(values.filter((value): value is string => !!value)),
  ];
  return fonts.length ? fonts.map(cssFontFamily).join(', ') : undefined;
}

function cssFontFamily(value: string): string {
  return /[\s"',]/.test(value)
    ? `"${value.replaceAll('\\', '\\\\').replaceAll('"', '\\"')}"`
    : value;
}

function numericAttribute(
  element: Element | undefined,
  name: string,
): number | undefined {
  if (!element) return undefined;
  const raw = wordAttribute(element, name);
  if (raw === null || !raw.trim()) return undefined;
  const value = Number(raw);
  return Number.isFinite(value) ? value : undefined;
}

function wordAttribute(element: Element, name: string): string | null {
  return attribute(element, name) ?? attribute(element, `w:${name}`);
}

function formatNumber(value: number): string {
  return Number(value.toFixed(2)).toString();
}

function formatLineHeightFactor(value: number): string {
  return Number(value.toFixed(4)).toString();
}
