import type { Editor } from '@tiptap/core';
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

export interface DocumentFontDialogSource {
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
  characterSpacingMode: DocumentCharacterSpacingMode;
  characterSpacingPoints: string;
}

export interface DocumentFontDialogPatch {
  characterSpacingTwips?: number;
}

export function documentFontDialogSource(
  editor: Editor,
): DocumentFontDialogSource {
  const { doc, selection } = editor.state;
  const values = new Map<string, number | null>();
  if (selection.empty) {
    addSpacingValue(
      values,
      editor.getAttributes('textStyle').characterSpacingTwips,
    );
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
      addSpacingValue(values, textStyle?.attrs.characterSpacingTwips);
    });
  }
  if (!values.size) {
    addSpacingValue(
      values,
      editor.getAttributes('textStyle').characterSpacingTwips,
    );
  }
  const mixed = values.size > 1;
  const value = mixed ? null : (values.values().next().value ?? null);
  const textStyle = editor.getAttributes('textStyle');
  const selectedText = selection.empty
    ? ''
    : doc
        .textBetween(selection.from, selection.to, ' ', ' ')
        .replace(/\s+/g, ' ')
        .trim();
  return {
    characterSpacing: { mixed, value },
    fontFamily:
      typeof textStyle.fontFamily === 'string' && textStyle.fontFamily.trim()
        ? textStyle.fontFamily
        : null,
    fontSize:
      typeof textStyle.fontSize === 'string' && textStyle.fontSize.trim()
        ? textStyle.fontSize
        : null,
    previewText: selectedText.slice(0, 72) || 'A3S Office 字符间距',
    selectedCharacters: selectedText.length,
  };
}

export function createDocumentFontDialogDraft(
  source: DocumentFontDialogSource,
): DocumentFontDialogDraft {
  const spacing = source.characterSpacing.value;
  const points = Math.abs(documentCharacterSpacingPoints(spacing) ?? 1);
  return {
    characterSpacingMode: source.characterSpacing.mixed
      ? 'mixed'
      : spacing !== null && spacing < 0
        ? 'condensed'
        : spacing !== null && spacing > 0
          ? 'expanded'
          : 'normal',
    characterSpacingPoints: formatPoints(points || 1),
  };
}

export function documentFontDialogDraftError(
  draft: DocumentFontDialogDraft,
): string | null {
  if (draft.characterSpacingMode === 'mixed') return null;
  if (draft.characterSpacingMode === 'normal') return null;
  return characterSpacingTwipsFromDraft(draft) === null
    ? `请输入 0.05 至 ${DOCUMENT_CHARACTER_SPACING_MAX_TWIPS / 20} 磅的间距。`
    : null;
}

export function documentFontDialogPatch(
  source: DocumentFontDialogSource,
  draft: DocumentFontDialogDraft,
  characterSpacingTouched: boolean,
): DocumentFontDialogPatch {
  if (!characterSpacingTouched || draft.characterSpacingMode === 'mixed') {
    return {};
  }
  const spacing = characterSpacingTwipsFromDraft(draft);
  if (spacing === null) return {};
  if (
    !source.characterSpacing.mixed &&
    source.characterSpacing.value === spacing
  ) {
    return {};
  }
  return { characterSpacingTwips: spacing };
}

export function applyDocumentFontDialogPatch(
  editor: Editor,
  selection: { from: number; to: number },
  patch: DocumentFontDialogPatch,
): boolean {
  const spacing = normalizeDocumentCharacterSpacingTwips(
    patch.characterSpacingTwips,
  );
  const maximum = editor.state.doc.content.size;
  if (
    editor.isDestroyed ||
    spacing === null ||
    !Number.isSafeInteger(selection.from) ||
    !Number.isSafeInteger(selection.to) ||
    selection.from < 0 ||
    selection.to < selection.from ||
    selection.to > maximum
  ) {
    return false;
  }
  return editor
    .chain()
    .setTextSelection(selection)
    .setDocumentCharacterSpacing(spacing)
    .focus()
    .scrollIntoView()
    .run();
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

function formatPoints(value: number): string {
  return String(Number(value.toFixed(2)));
}
