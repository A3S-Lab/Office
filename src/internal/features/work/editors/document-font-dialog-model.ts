import type { Editor } from '@tiptap/core';
import {
  DOCUMENT_CHARACTER_SCALE_DEFAULT_PERCENT,
  DOCUMENT_CHARACTER_SCALE_MAX_PERCENT,
  DOCUMENT_CHARACTER_SCALE_MIN_PERCENT,
  normalizeDocumentCharacterScalePercent,
} from '../work-document-character-scale';
import {
  DOCUMENT_CHARACTER_POSITION_MAX_HALF_POINTS,
  documentCharacterPositionPoints,
  normalizeDocumentCharacterPositionHalfPoints,
} from '../work-document-character-position';
import {
  DOCUMENT_CHARACTER_SPACING_MAX_TWIPS,
  documentCharacterSpacingPoints,
  normalizeDocumentCharacterSpacingTwips,
} from '../work-document-character-spacing';
import {
  DOCUMENT_KERNING_DEFAULT_THRESHOLD_HALF_POINTS,
  DOCUMENT_KERNING_THRESHOLD_MAX_HALF_POINTS,
  documentKerningThresholdPoints,
  normalizeDocumentKerningThresholdHalfPoints,
} from '../work-document-kerning';
import {
  normalizeDocumentEmphasisMark,
  type WorkDocumentEmphasisMark,
} from '../work-document-emphasis';
import { normalizeDocumentHiddenText } from '../work-document-hidden-text';
import {
  documentLegacyTextEffectsConflict,
  documentLegacyTextEffectsFromTextStyleAttributes,
  normalizeDocumentLegacyTextEffect,
  type WorkDocumentLegacyTextEffects,
} from '../work-document-legacy-text-effects';
import type { WorkDocumentScriptFontPatch } from '../work-document-script-fonts';
import {
  addDocumentScriptFontValues,
  appendDocumentFontDialogScriptFontPatch,
  applyDocumentScriptFontPatch,
  documentFontFamilyDraftValue,
  type DocumentFontDialogScriptFontPatch,
  type DocumentFontFamilySource,
} from './document-font-dialog-script-font-model';
import {
  addDocumentFontDialogRunBorderValue,
  createDocumentFontDialogRunBorderDraft,
  documentFontDialogRunBorderDraftError,
  documentFontDialogRunBorderPatch,
  type DocumentFontDialogRunBorderDraft,
  type DocumentFontDialogRunBorderPatch,
  type DocumentFontDialogRunBorderSource,
  selectedDocumentFontDialogRunBorderValue,
} from './document-font-dialog-run-border-model';
import {
  normalizeDocumentRunBorder,
  serializeDocumentRunBorder,
} from '../work-document-run-border';
import {
  addDocumentFontDialogRunShadingValue,
  createDocumentFontDialogRunShadingDraft,
  documentFontDialogRunShadingDraftError,
  documentFontDialogRunShadingPatch,
  type DocumentFontDialogRunShadingDraft,
  type DocumentFontDialogRunShadingPatch,
  type DocumentFontDialogRunShadingSource,
  selectedDocumentFontDialogRunShadingValue,
} from './document-font-dialog-run-shading-model';
import {
  normalizeDocumentRunShading,
  serializeDocumentRunShading,
} from '../work-document-run-shading';

export type { DocumentFontFamilySource } from './document-font-dialog-script-font-model';

export type DocumentCharacterSpacingMode =
  | 'condensed'
  | 'expanded'
  | 'mixed'
  | 'normal';

export type DocumentCharacterPositionMode =
  | 'lowered'
  | 'mixed'
  | 'normal'
  | 'raised';

export type DocumentCharacterScaleMode = 'mixed' | 'value';

export type DocumentEmphasisMarkMode =
  | 'inherit'
  | 'mixed'
  | WorkDocumentEmphasisMark;

export interface DocumentFontDialogSource {
  characterScale: {
    mixed: boolean;
    value: number | null;
  };
  characterPosition: {
    mixed: boolean;
    value: number | null;
  };
  characterSpacing: {
    mixed: boolean;
    value: number | null;
  };
  kerningThreshold: {
    mixed: boolean;
    value: number | null;
  };
  emphasisMark: {
    mixed: boolean;
    value: WorkDocumentEmphasisMark | null;
  };
  hiddenText: {
    mixed: boolean;
    value: boolean | null;
  };
  legacyTextOutline: { mixed: boolean; value: boolean | null };
  legacyTextShadow: { mixed: boolean; value: boolean | null };
  legacyTextEmboss: { mixed: boolean; value: boolean | null };
  legacyTextImprint: { mixed: boolean; value: boolean | null };
  runBorder: DocumentFontDialogRunBorderSource;
  runShading: DocumentFontDialogRunShadingSource;
  latinFont: DocumentFontFamilySource;
  eastAsiaFont: DocumentFontFamilySource;
  complexScriptFont: DocumentFontFamilySource;
  fontFamily: string | null;
  fontSize: string | null;
  previewText: string;
  selectedCharacters: number;
}

export interface DocumentFontDialogDraft
  extends DocumentFontDialogRunBorderDraft,
    DocumentFontDialogRunShadingDraft {
  characterScaleMode: DocumentCharacterScaleMode;
  characterScalePercent: string;
  characterPositionMode: DocumentCharacterPositionMode;
  characterPositionPoints: string;
  characterSpacingMode: DocumentCharacterSpacingMode;
  characterSpacingPoints: string;
  kerningEnabled: boolean;
  kerningThresholdPoints: string;
  emphasisMark: DocumentEmphasisMarkMode;
  hiddenText: boolean;
  legacyTextOutline: boolean;
  legacyTextShadow: boolean;
  legacyTextEmboss: boolean;
  legacyTextImprint: boolean;
  latinFont: string;
  eastAsiaFont: string;
  complexScriptFont: string;
}

export interface DocumentFontDialogPatch
  extends DocumentFontDialogScriptFontPatch,
    DocumentFontDialogRunBorderPatch,
    DocumentFontDialogRunShadingPatch {
  characterScalePercent?: number;
  characterPositionHalfPoints?: number;
  characterSpacingTwips?: number;
  kerningThresholdHalfPoints?: number | null;
  emphasisMark?: WorkDocumentEmphasisMark | null;
  hiddenText?: boolean;
  legacyTextOutline?: boolean;
  legacyTextShadow?: boolean;
  legacyTextEmboss?: boolean;
  legacyTextImprint?: boolean;
}

export interface DocumentFontDialogTouched {
  characterPosition: boolean;
  characterScale: boolean;
  characterSpacing: boolean;
  complexScriptFont: boolean;
  eastAsiaFont: boolean;
  emphasisMark: boolean;
  hiddenText: boolean;
  legacyTextOutline: boolean;
  legacyTextShadow: boolean;
  legacyTextEmboss: boolean;
  legacyTextImprint: boolean;
  runBorder: boolean;
  runShading: boolean;
  kerning: boolean;
  latinFont: boolean;
}

export function documentFontDialogSource(
  editor: Editor,
): DocumentFontDialogSource {
  const { doc, selection } = editor.state;
  const scaleValues = new Map<string, number | null>();
  const positionValues = new Map<string, number | null>();
  const spacingValues = new Map<string, number | null>();
  const kerningValues = new Map<string, number | null>();
  const emphasisValues = new Map<string, WorkDocumentEmphasisMark | null>();
  const hiddenTextValues = new Map<string, boolean | null>();
  const legacyTextOutlineValues = new Map<string, boolean | null>();
  const legacyTextShadowValues = new Map<string, boolean | null>();
  const legacyTextEmbossValues = new Map<string, boolean | null>();
  const legacyTextImprintValues = new Map<string, boolean | null>();
  const runBorderValues = new Map<
    string,
    DocumentFontDialogRunBorderSource['value']
  >();
  const runShadingValues = new Map<
    string,
    DocumentFontDialogRunShadingSource['value']
  >();
  const latinFontValues = new Map<string, string | null>();
  const eastAsiaFontValues = new Map<string, string | null>();
  const complexScriptFontValues = new Map<string, string | null>();
  const addValues = (attributes: Record<string, unknown>) => {
    addScaleValue(scaleValues, attributes.characterScalePercent);
    addPositionValue(positionValues, attributes.characterPositionHalfPoints);
    addSpacingValue(spacingValues, attributes.characterSpacingTwips);
    addKerningValue(kerningValues, attributes.kerningThresholdHalfPoints);
    addEmphasisValue(emphasisValues, attributes.emphasisMark);
    addHiddenTextValue(hiddenTextValues, attributes.hiddenText);
    addLegacyTextEffectValue(
      legacyTextOutlineValues,
      attributes.legacyTextOutline,
    );
    addLegacyTextEffectValue(
      legacyTextShadowValues,
      attributes.legacyTextShadow,
    );
    addLegacyTextEffectValue(
      legacyTextEmbossValues,
      attributes.legacyTextEmboss,
    );
    addLegacyTextEffectValue(
      legacyTextImprintValues,
      attributes.legacyTextImprint,
    );
    addDocumentFontDialogRunBorderValue(runBorderValues, attributes.runBorder);
    addDocumentFontDialogRunShadingValue(
      runShadingValues,
      attributes.runShading,
    );
    addDocumentScriptFontValues(
      latinFontValues,
      eastAsiaFontValues,
      complexScriptFontValues,
      attributes,
    );
  };
  if (selection.empty) {
    addValues(editor.getAttributes('textStyle'));
  } else {
    doc.nodesBetween(selection.from, selection.to, (node, position) => {
      if (
        !node.isText ||
        position >= selection.to ||
        position + node.nodeSize <= selection.from
      ) {
        return;
      }
      const textStyle = node.marks.find(
        (mark) => mark.type.name === 'textStyle',
      );
      addValues(textStyle?.attrs ?? {});
    });
  }
  if (
    !scaleValues.size ||
    !spacingValues.size ||
    !positionValues.size ||
    !kerningValues.size ||
    !emphasisValues.size ||
    !hiddenTextValues.size ||
    !legacyTextOutlineValues.size ||
    !legacyTextShadowValues.size ||
    !legacyTextEmbossValues.size ||
    !legacyTextImprintValues.size ||
    !runBorderValues.size ||
    !runShadingValues.size ||
    !latinFontValues.size ||
    !eastAsiaFontValues.size ||
    !complexScriptFontValues.size
  ) {
    addValues(editor.getAttributes('textStyle'));
  }
  const characterScale = selectedValue(scaleValues);
  const characterPosition = selectedValue(positionValues);
  const characterSpacing = selectedValue(spacingValues);
  const kerningThreshold = selectedValue(kerningValues);
  const emphasisMark = selectedValue(emphasisValues);
  const hiddenText = selectedValue(hiddenTextValues);
  const legacyTextOutline = selectedValue(legacyTextOutlineValues);
  const legacyTextShadow = selectedValue(legacyTextShadowValues);
  const legacyTextEmboss = selectedValue(legacyTextEmbossValues);
  const legacyTextImprint = selectedValue(legacyTextImprintValues);
  const runBorder = selectedDocumentFontDialogRunBorderValue(runBorderValues);
  const runShading =
    selectedDocumentFontDialogRunShadingValue(runShadingValues);
  const latinFont = selectedValue(latinFontValues);
  const eastAsiaFont = selectedValue(eastAsiaFontValues);
  const complexScriptFont = selectedValue(complexScriptFontValues);
  const textStyle = editor.getAttributes('textStyle');
  const selectedText = selection.empty
    ? ''
    : doc
        .textBetween(selection.from, selection.to, ' ', ' ')
        .replace(/\s+/g, ' ')
        .trim();
  return {
    characterScale,
    characterPosition,
    characterSpacing,
    kerningThreshold,
    emphasisMark,
    hiddenText,
    legacyTextOutline,
    legacyTextShadow,
    legacyTextEmboss,
    legacyTextImprint,
    runBorder,
    runShading,
    latinFont,
    eastAsiaFont,
    complexScriptFont,
    fontFamily:
      typeof textStyle.fontFamily === 'string' && textStyle.fontFamily.trim()
        ? textStyle.fontFamily
        : null,
    fontSize:
      typeof textStyle.fontSize === 'string' && textStyle.fontSize.trim()
        ? textStyle.fontSize
        : null,
    previewText: selectedText.slice(0, 72) || 'A3S Office 字符格式',
    selectedCharacters: selectedText.length,
  };
}

export function createDocumentFontDialogDraft(
  source: DocumentFontDialogSource,
): DocumentFontDialogDraft {
  const scale = source.characterScale.value;
  const spacing = source.characterSpacing.value;
  const spacingPoints = Math.abs(documentCharacterSpacingPoints(spacing) ?? 1);
  const position = source.characterPosition.value;
  const positionPoints = Math.abs(
    documentCharacterPositionPoints(position) ?? 1,
  );
  const kerningThreshold = source.kerningThreshold.value;
  return {
    characterScaleMode: source.characterScale.mixed ? 'mixed' : 'value',
    characterScalePercent: source.characterScale.mixed
      ? ''
      : String(scale ?? DOCUMENT_CHARACTER_SCALE_DEFAULT_PERCENT),
    characterPositionMode: source.characterPosition.mixed
      ? 'mixed'
      : position !== null && position < 0
        ? 'lowered'
        : position !== null && position > 0
          ? 'raised'
          : 'normal',
    characterPositionPoints: formatPoints(positionPoints || 1),
    characterSpacingMode: source.characterSpacing.mixed
      ? 'mixed'
      : spacing !== null && spacing < 0
        ? 'condensed'
        : spacing !== null && spacing > 0
          ? 'expanded'
          : 'normal',
    characterSpacingPoints: formatPoints(spacingPoints || 1),
    kerningEnabled: !source.kerningThreshold.mixed && kerningThreshold !== null,
    kerningThresholdPoints: formatPoints(
      documentKerningThresholdPoints(kerningThreshold) ??
        DOCUMENT_KERNING_DEFAULT_THRESHOLD_HALF_POINTS / 2,
    ),
    emphasisMark: source.emphasisMark.mixed
      ? 'mixed'
      : (source.emphasisMark.value ?? 'inherit'),
    hiddenText: source.hiddenText.value ?? false,
    legacyTextOutline: source.legacyTextOutline.value ?? false,
    legacyTextShadow: source.legacyTextShadow.value ?? false,
    legacyTextEmboss: source.legacyTextEmboss.value ?? false,
    legacyTextImprint: source.legacyTextImprint.value ?? false,
    ...createDocumentFontDialogRunBorderDraft(source.runBorder),
    ...createDocumentFontDialogRunShadingDraft(source.runShading),
    latinFont: documentFontFamilyDraftValue(source.latinFont),
    eastAsiaFont: documentFontFamilyDraftValue(source.eastAsiaFont),
    complexScriptFont: documentFontFamilyDraftValue(source.complexScriptFont),
  };
}

export function documentFontDialogDraftError(
  draft: DocumentFontDialogDraft,
): string | null {
  const runBorderError = documentFontDialogRunBorderDraftError(draft);
  if (runBorderError) return runBorderError;
  const runShadingError = documentFontDialogRunShadingDraftError(draft);
  if (runShadingError) return runShadingError;
  if (
    draft.characterScaleMode !== 'mixed' &&
    characterScalePercentFromDraft(draft) === null
  ) {
    return `请输入 ${DOCUMENT_CHARACTER_SCALE_MIN_PERCENT} 至 ${DOCUMENT_CHARACTER_SCALE_MAX_PERCENT} 的整数缩放比例。`;
  }
  if (
    draft.characterSpacingMode !== 'mixed' &&
    draft.characterSpacingMode !== 'normal' &&
    characterSpacingTwipsFromDraft(draft) === null
  ) {
    return `请输入 0.05 至 ${DOCUMENT_CHARACTER_SPACING_MAX_TWIPS / 20} 磅的间距。`;
  }
  if (
    draft.characterPositionMode !== 'mixed' &&
    draft.characterPositionMode !== 'normal' &&
    characterPositionHalfPointsFromDraft(draft) === null
  ) {
    return `请输入 0.5 至 ${DOCUMENT_CHARACTER_POSITION_MAX_HALF_POINTS / 2} 磅、以 0.5 磅递增的位置值。`;
  }
  if (
    draft.kerningEnabled &&
    kerningThresholdHalfPointsFromDraft(draft) === null
  ) {
    return `请输入 0 至 ${DOCUMENT_KERNING_THRESHOLD_MAX_HALF_POINTS / 2} 磅、以 0.5 磅递增的字距调整阈值。`;
  }
  return null;
}

export function documentFontDialogPatch(
  source: DocumentFontDialogSource,
  draft: DocumentFontDialogDraft,
  touched: DocumentFontDialogTouched,
): DocumentFontDialogPatch {
  const patch: DocumentFontDialogPatch = {};
  if (touched.characterScale && draft.characterScaleMode !== 'mixed') {
    const scale = characterScalePercentFromDraft(draft);
    if (
      scale !== null &&
      (source.characterScale.mixed || source.characterScale.value !== scale)
    ) {
      patch.characterScalePercent = scale;
    }
  }
  if (touched.characterSpacing && draft.characterSpacingMode !== 'mixed') {
    const spacing = characterSpacingTwipsFromDraft(draft);
    if (
      spacing !== null &&
      (source.characterSpacing.mixed ||
        source.characterSpacing.value !== spacing)
    ) {
      patch.characterSpacingTwips = spacing;
    }
  }
  if (touched.characterPosition && draft.characterPositionMode !== 'mixed') {
    const position = characterPositionHalfPointsFromDraft(draft);
    if (
      position !== null &&
      (source.characterPosition.mixed ||
        source.characterPosition.value !== position)
    ) {
      patch.characterPositionHalfPoints = position;
    }
  }
  if (touched.kerning) {
    if (!draft.kerningEnabled) {
      if (
        source.kerningThreshold.mixed ||
        source.kerningThreshold.value !== null
      ) {
        patch.kerningThresholdHalfPoints = null;
      }
    } else {
      const threshold = kerningThresholdHalfPointsFromDraft(draft);
      if (
        threshold !== null &&
        (source.kerningThreshold.mixed ||
          source.kerningThreshold.value !== threshold)
      ) {
        patch.kerningThresholdHalfPoints = threshold;
      }
    }
  }
  if (touched.emphasisMark && draft.emphasisMark !== 'mixed') {
    const emphasisMark =
      draft.emphasisMark === 'inherit'
        ? null
        : normalizeDocumentEmphasisMark(draft.emphasisMark);
    if (emphasisMark !== null || draft.emphasisMark === 'inherit') {
      if (
        source.emphasisMark.mixed ||
        source.emphasisMark.value !== emphasisMark
      ) {
        patch.emphasisMark = emphasisMark;
      }
    }
  }
  if (
    touched.hiddenText &&
    (source.hiddenText.mixed || source.hiddenText.value !== draft.hiddenText)
  ) {
    patch.hiddenText = draft.hiddenText;
  }
  appendLegacyTextEffectPatch(
    patch,
    'legacyTextOutline',
    source.legacyTextOutline,
    draft.legacyTextOutline,
    touched.legacyTextOutline,
  );
  Object.assign(
    patch,
    documentFontDialogRunBorderPatch(
      source.runBorder,
      draft,
      touched.runBorder,
    ),
  );
  Object.assign(
    patch,
    documentFontDialogRunShadingPatch(
      source.runShading,
      draft,
      touched.runShading,
    ),
  );
  appendLegacyTextEffectPatch(
    patch,
    'legacyTextShadow',
    source.legacyTextShadow,
    draft.legacyTextShadow,
    touched.legacyTextShadow,
  );
  appendLegacyTextEffectPatch(
    patch,
    'legacyTextEmboss',
    source.legacyTextEmboss,
    draft.legacyTextEmboss,
    touched.legacyTextEmboss,
  );
  appendLegacyTextEffectPatch(
    patch,
    'legacyTextImprint',
    source.legacyTextImprint,
    draft.legacyTextImprint,
    touched.legacyTextImprint,
  );
  appendDocumentFontDialogScriptFontPatch(
    patch,
    'latinFont',
    source.latinFont,
    draft.latinFont,
    touched.latinFont,
  );
  appendDocumentFontDialogScriptFontPatch(
    patch,
    'eastAsiaFont',
    source.eastAsiaFont,
    draft.eastAsiaFont,
    touched.eastAsiaFont,
  );
  appendDocumentFontDialogScriptFontPatch(
    patch,
    'complexScriptFont',
    source.complexScriptFont,
    draft.complexScriptFont,
    touched.complexScriptFont,
  );
  return patch;
}

export function applyDocumentFontDialogPatch(
  editor: Editor,
  selection: { from: number; to: number },
  patch: DocumentFontDialogPatch,
): boolean {
  const hasScale = patch.characterScalePercent !== undefined;
  const hasPosition = patch.characterPositionHalfPoints !== undefined;
  const hasSpacing = patch.characterSpacingTwips !== undefined;
  const hasKerning = patch.kerningThresholdHalfPoints !== undefined;
  const hasEmphasis = patch.emphasisMark !== undefined;
  const hasHiddenText = patch.hiddenText !== undefined;
  const hasLegacyTextOutline = patch.legacyTextOutline !== undefined;
  const hasLegacyTextShadow = patch.legacyTextShadow !== undefined;
  const hasLegacyTextEmboss = patch.legacyTextEmboss !== undefined;
  const hasLegacyTextImprint = patch.legacyTextImprint !== undefined;
  const hasRunBorder = patch.runBorder !== undefined;
  const hasRunShading = patch.runShading !== undefined;
  const hasLegacyTextEffects =
    hasLegacyTextOutline ||
    hasLegacyTextShadow ||
    hasLegacyTextEmboss ||
    hasLegacyTextImprint;
  const scriptFontPatch: WorkDocumentScriptFontPatch = {
    ...(patch.latinFont !== undefined ? { latin: patch.latinFont } : {}),
    ...(patch.eastAsiaFont !== undefined
      ? { eastAsia: patch.eastAsiaFont }
      : {}),
    ...(patch.complexScriptFont !== undefined
      ? { complexScript: patch.complexScriptFont }
      : {}),
  };
  const hasScriptFonts = Object.keys(scriptFontPatch).length > 0;
  const scale = hasScale
    ? normalizeDocumentCharacterScalePercent(patch.characterScalePercent)
    : null;
  const position = hasPosition
    ? normalizeDocumentCharacterPositionHalfPoints(
        patch.characterPositionHalfPoints,
      )
    : null;
  const spacing = hasSpacing
    ? normalizeDocumentCharacterSpacingTwips(patch.characterSpacingTwips)
    : null;
  const kerningThreshold =
    hasKerning && patch.kerningThresholdHalfPoints !== null
      ? normalizeDocumentKerningThresholdHalfPoints(
          patch.kerningThresholdHalfPoints,
        )
      : null;
  const emphasisMark =
    hasEmphasis && patch.emphasisMark !== null
      ? normalizeDocumentEmphasisMark(patch.emphasisMark)
      : null;
  const runBorder =
    hasRunBorder && patch.runBorder !== null
      ? normalizeDocumentRunBorder(patch.runBorder)
      : null;
  const runShading =
    hasRunShading && patch.runShading !== null
      ? normalizeDocumentRunShading(patch.runShading)
      : null;
  const maximum = editor.state.doc.content.size;
  if (
    editor.isDestroyed ||
    (!hasScale &&
      !hasPosition &&
      !hasSpacing &&
      !hasKerning &&
      !hasEmphasis &&
      !hasHiddenText &&
      !hasLegacyTextEffects &&
      !hasRunBorder &&
      !hasRunShading &&
      !hasScriptFonts) ||
    (hasScale && scale === null) ||
    (hasPosition && position === null) ||
    (hasSpacing && spacing === null) ||
    (hasKerning &&
      patch.kerningThresholdHalfPoints !== null &&
      kerningThreshold === null) ||
    (hasEmphasis && patch.emphasisMark !== null && emphasisMark === null) ||
    (hasRunBorder && patch.runBorder !== null && runBorder === null) ||
    (hasRunShading && patch.runShading !== null && runShading === null) ||
    (hasHiddenText && typeof patch.hiddenText !== 'boolean') ||
    (hasLegacyTextOutline &&
      normalizeDocumentLegacyTextEffect(patch.legacyTextOutline) === null) ||
    (hasLegacyTextShadow &&
      normalizeDocumentLegacyTextEffect(patch.legacyTextShadow) === null) ||
    (hasLegacyTextEmboss &&
      normalizeDocumentLegacyTextEffect(patch.legacyTextEmboss) === null) ||
    (hasLegacyTextImprint &&
      normalizeDocumentLegacyTextEffect(patch.legacyTextImprint) === null) ||
    !Number.isSafeInteger(selection.from) ||
    !Number.isSafeInteger(selection.to) ||
    selection.from < 0 ||
    selection.to < selection.from ||
    selection.to > maximum ||
    (hasLegacyTextEffects &&
      !documentFontDialogLegacyTextEffectsAreSafe(editor, selection, patch))
  ) {
    return false;
  }
  const attributes: Record<string, boolean | number | string | null> = {};
  if (scale !== null) attributes.characterScalePercent = scale;
  if (position !== null) attributes.characterPositionHalfPoints = position;
  if (spacing !== null) attributes.characterSpacingTwips = spacing;
  if (hasKerning) {
    attributes.kerningThresholdHalfPoints = kerningThreshold;
  }
  if (hasEmphasis) attributes.emphasisMark = emphasisMark;
  if (hasHiddenText) attributes.hiddenText = patch.hiddenText ?? false;
  if (hasLegacyTextOutline) {
    attributes.legacyTextOutline = patch.legacyTextOutline ?? false;
  }
  if (hasLegacyTextShadow) {
    attributes.legacyTextShadow = patch.legacyTextShadow ?? false;
  }
  if (hasLegacyTextEmboss) {
    attributes.legacyTextEmboss = patch.legacyTextEmboss ?? false;
  }
  if (hasLegacyTextImprint) {
    attributes.legacyTextImprint = patch.legacyTextImprint ?? false;
  }
  if (hasRunBorder) {
    attributes.runBorder =
      patch.runBorder === null
        ? null
        : (serializeDocumentRunBorder(runBorder) ?? null);
  }
  if (hasRunShading) {
    attributes.runShading =
      patch.runShading === null
        ? null
        : (serializeDocumentRunShading(runShading) ?? null);
  }
  if (hasScriptFonts) {
    return applyDocumentScriptFontPatch(
      editor,
      selection,
      attributes,
      scriptFontPatch,
    );
  }
  return editor
    .chain()
    .setTextSelection(selection)
    .setMark('textStyle', attributes)
    .removeEmptyTextStyle()
    .focus()
    .scrollIntoView()
    .run();
}

function addScaleValue(
  values: Map<string, number | null>,
  value: unknown,
): void {
  const scale = normalizeDocumentCharacterScalePercent(value);
  values.set(scale === null ? 'inherited' : String(scale), scale);
}

function documentFontDialogLegacyTextEffectsAreSafe(
  editor: Editor,
  selection: { from: number; to: number },
  patch: DocumentFontDialogPatch,
): boolean {
  const patchEffects: WorkDocumentLegacyTextEffects = {
    ...(patch.legacyTextOutline !== undefined
      ? { outline: patch.legacyTextOutline }
      : {}),
    ...(patch.legacyTextShadow !== undefined
      ? { shadow: patch.legacyTextShadow }
      : {}),
    ...(patch.legacyTextEmboss !== undefined
      ? { emboss: patch.legacyTextEmboss }
      : {}),
    ...(patch.legacyTextImprint !== undefined
      ? { imprint: patch.legacyTextImprint }
      : {}),
  };
  const attributesAreSafe = (attributes: Record<string, unknown>) => {
    const current =
      documentLegacyTextEffectsFromTextStyleAttributes(attributes);
    return (
      current !== null &&
      !documentLegacyTextEffectsConflict({ ...current, ...patchEffects })
    );
  };
  if (selection.from === selection.to) {
    return attributesAreSafe(editor.getAttributes('textStyle'));
  }
  let sawText = false;
  let safe = true;
  editor.state.doc.nodesBetween(
    selection.from,
    selection.to,
    (node, position) => {
      if (
        !safe ||
        !node.isText ||
        position >= selection.to ||
        position + node.nodeSize <= selection.from
      ) {
        return;
      }
      sawText = true;
      const textStyle = node.marks.find(
        (mark) => mark.type.name === 'textStyle',
      );
      safe = attributesAreSafe(textStyle?.attrs ?? {});
    },
  );
  return (
    safe && (sawText || attributesAreSafe(editor.getAttributes('textStyle')))
  );
}

function addPositionValue(
  values: Map<string, number | null>,
  value: unknown,
): void {
  const position = normalizeDocumentCharacterPositionHalfPoints(value);
  values.set(position === null ? 'inherited' : String(position), position);
}

function addSpacingValue(
  values: Map<string, number | null>,
  value: unknown,
): void {
  const spacing = normalizeDocumentCharacterSpacingTwips(value);
  values.set(spacing === null ? 'inherited' : String(spacing), spacing);
}

function addKerningValue(
  values: Map<string, number | null>,
  value: unknown,
): void {
  const threshold = normalizeDocumentKerningThresholdHalfPoints(value);
  values.set(threshold === null ? 'inherited' : String(threshold), threshold);
}

function addEmphasisValue(
  values: Map<string, WorkDocumentEmphasisMark | null>,
  value: unknown,
): void {
  const emphasisMark = normalizeDocumentEmphasisMark(value);
  values.set(emphasisMark === null ? 'inherited' : emphasisMark, emphasisMark);
}

function addHiddenTextValue(
  values: Map<string, boolean | null>,
  value: unknown,
): void {
  const hiddenText = normalizeDocumentHiddenText(value);
  values.set(
    hiddenText === null ? 'inherited' : String(hiddenText),
    hiddenText,
  );
}

type DocumentFontDialogLegacyTextEffectPatchName =
  | 'legacyTextOutline'
  | 'legacyTextShadow'
  | 'legacyTextEmboss'
  | 'legacyTextImprint';

function appendLegacyTextEffectPatch(
  patch: DocumentFontDialogPatch,
  name: DocumentFontDialogLegacyTextEffectPatchName,
  source: { mixed: boolean; value: boolean | null },
  value: boolean,
  touched: boolean,
): void {
  if (touched && (source.mixed || source.value !== value)) {
    patch[name] = value;
  }
}

function addLegacyTextEffectValue(
  values: Map<string, boolean | null>,
  value: unknown,
): void {
  const effect = normalizeDocumentLegacyTextEffect(value);
  values.set(effect === null ? 'inherited' : String(effect), effect);
}

function characterSpacingTwipsFromDraft(
  draft: DocumentFontDialogDraft,
): number | null {
  if (draft.characterSpacingMode === 'normal') return 0;
  if (draft.characterSpacingMode === 'mixed') return null;
  const points = Number(draft.characterSpacingPoints);
  const maximumPoints = DOCUMENT_CHARACTER_SPACING_MAX_TWIPS / 20;
  if (!Number.isFinite(points) || points < 0.05 || points > maximumPoints) {
    return null;
  }
  const magnitude = Math.round(points * 20);
  if (magnitude <= 0) return null;
  return normalizeDocumentCharacterSpacingTwips(
    draft.characterSpacingMode === 'condensed' ? -magnitude : magnitude,
  );
}

function characterScalePercentFromDraft(
  draft: DocumentFontDialogDraft,
): number | null {
  if (draft.characterScaleMode === 'mixed') return null;
  return normalizeDocumentCharacterScalePercent(draft.characterScalePercent);
}

function characterPositionHalfPointsFromDraft(
  draft: DocumentFontDialogDraft,
): number | null {
  if (draft.characterPositionMode === 'normal') return 0;
  if (draft.characterPositionMode === 'mixed') return null;
  const points = Number(draft.characterPositionPoints);
  const maximumPoints = DOCUMENT_CHARACTER_POSITION_MAX_HALF_POINTS / 2;
  const magnitude = points * 2;
  if (
    !Number.isFinite(points) ||
    points < 0.5 ||
    points > maximumPoints ||
    !Number.isSafeInteger(magnitude)
  ) {
    return null;
  }
  return normalizeDocumentCharacterPositionHalfPoints(
    draft.characterPositionMode === 'lowered' ? -magnitude : magnitude,
  );
}

function kerningThresholdHalfPointsFromDraft(
  draft: DocumentFontDialogDraft,
): number | null {
  if (!draft.kerningEnabled) return null;
  const points = Number(draft.kerningThresholdPoints);
  const halfPoints = points * 2;
  if (
    !Number.isFinite(points) ||
    points < 0 ||
    points > DOCUMENT_KERNING_THRESHOLD_MAX_HALF_POINTS / 2 ||
    !Number.isSafeInteger(halfPoints)
  ) {
    return null;
  }
  return normalizeDocumentKerningThresholdHalfPoints(halfPoints);
}

function selectedValue<T>(values: Map<string, T | null>): {
  mixed: boolean;
  value: T | null;
} {
  const mixed = values.size > 1;
  return {
    mixed,
    value: mixed ? null : (values.values().next().value ?? null),
  };
}

function formatPoints(value: number): string {
  return String(Number(value.toFixed(2)));
}
