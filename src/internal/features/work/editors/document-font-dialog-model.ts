import type { Editor } from '@tiptap/core';
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

export interface DocumentFontDialogSource {
  characterPosition: {
    mixed: boolean;
    value: number | null;
  };
  characterSpacing: {
    mixed: boolean;
    value: number | null;
  };
  fontFamily: string | null;
  fontSize: string | null;
  previewText: string;
  selectedCharacters: number;
}

export interface DocumentFontDialogDraft {
  characterPositionMode: DocumentCharacterPositionMode;
  characterPositionPoints: string;
  characterSpacingMode: DocumentCharacterSpacingMode;
  characterSpacingPoints: string;
}

export interface DocumentFontDialogPatch {
  characterPositionHalfPoints?: number;
  characterSpacingTwips?: number;
}

export function documentFontDialogSource(
  editor: Editor,
): DocumentFontDialogSource {
  const { doc, selection } = editor.state;
  const positionValues = new Map<string, number | null>();
  const spacingValues = new Map<string, number | null>();
  const addValues = (attributes: Record<string, unknown>) => {
    addPositionValue(positionValues, attributes.characterPositionHalfPoints);
    addSpacingValue(spacingValues, attributes.characterSpacingTwips);
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
  if (!spacingValues.size || !positionValues.size) {
    addValues(editor.getAttributes('textStyle'));
  }
  const characterPosition = selectedValue(positionValues);
  const characterSpacing = selectedValue(spacingValues);
  const textStyle = editor.getAttributes('textStyle');
  const selectedText = selection.empty
    ? ''
    : doc
        .textBetween(selection.from, selection.to, ' ', ' ')
        .replace(/\s+/g, ' ')
        .trim();
  return {
    characterPosition,
    characterSpacing,
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
  const spacing = source.characterSpacing.value;
  const spacingPoints = Math.abs(documentCharacterSpacingPoints(spacing) ?? 1);
  const position = source.characterPosition.value;
  const positionPoints = Math.abs(
    documentCharacterPositionPoints(position) ?? 1,
  );
  return {
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
  };
}

export function documentFontDialogDraftError(
  draft: DocumentFontDialogDraft,
): string | null {
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
  return null;
}

export function documentFontDialogPatch(
  source: DocumentFontDialogSource,
  draft: DocumentFontDialogDraft,
  characterSpacingTouched: boolean,
  characterPositionTouched: boolean,
): DocumentFontDialogPatch {
  const patch: DocumentFontDialogPatch = {};
  if (characterSpacingTouched && draft.characterSpacingMode !== 'mixed') {
    const spacing = characterSpacingTwipsFromDraft(draft);
    if (
      spacing !== null &&
      (source.characterSpacing.mixed ||
        source.characterSpacing.value !== spacing)
    ) {
      patch.characterSpacingTwips = spacing;
    }
  }
  if (characterPositionTouched && draft.characterPositionMode !== 'mixed') {
    const position = characterPositionHalfPointsFromDraft(draft);
    if (
      position !== null &&
      (source.characterPosition.mixed ||
        source.characterPosition.value !== position)
    ) {
      patch.characterPositionHalfPoints = position;
    }
  }
  return patch;
}

export function applyDocumentFontDialogPatch(
  editor: Editor,
  selection: { from: number; to: number },
  patch: DocumentFontDialogPatch,
): boolean {
  const hasPosition = patch.characterPositionHalfPoints !== undefined;
  const hasSpacing = patch.characterSpacingTwips !== undefined;
  const position = hasPosition
    ? normalizeDocumentCharacterPositionHalfPoints(
        patch.characterPositionHalfPoints,
      )
    : null;
  const spacing = hasSpacing
    ? normalizeDocumentCharacterSpacingTwips(patch.characterSpacingTwips)
    : null;
  const maximum = editor.state.doc.content.size;
  if (
    editor.isDestroyed ||
    (!hasPosition && !hasSpacing) ||
    (hasPosition && position === null) ||
    (hasSpacing && spacing === null) ||
    !Number.isSafeInteger(selection.from) ||
    !Number.isSafeInteger(selection.to) ||
    selection.from < 0 ||
    selection.to < selection.from ||
    selection.to > maximum
  ) {
    return false;
  }
  const attributes: Record<string, number> = {};
  if (position !== null) attributes.characterPositionHalfPoints = position;
  if (spacing !== null) attributes.characterSpacingTwips = spacing;
  return editor
    .chain()
    .setTextSelection(selection)
    .setMark('textStyle', attributes)
    .focus()
    .scrollIntoView()
    .run();
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

function selectedValue(values: Map<string, number | null>): {
  mixed: boolean;
  value: number | null;
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
