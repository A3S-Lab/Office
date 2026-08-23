import type { Editor } from '@tiptap/core';
import type { Mark } from '@tiptap/pm/model';
import { TextSelection } from '@tiptap/pm/state';
import {
  cssDocumentFontFamily,
  documentFontNameFromCssFamily,
  documentScriptFontDirectFamily,
  documentScriptFontFamilyForRendering,
  documentScriptFontSegments,
  documentScriptFontSlotFromHint,
  normalizeDocumentScriptFontSlot,
  parseDocumentScriptFonts,
  patchDocumentScriptFonts,
  serializeDocumentScriptFonts,
  type WorkDocumentScriptFontPatch,
  type WorkDocumentScriptFontSlot,
} from '../work-document-script-fonts';
import { documentWordLineHeightFactor } from '../work-document-word-line-metrics';

export interface DocumentFontFamilySource {
  mixed: boolean;
  value: string | null;
}

export interface DocumentFontDialogScriptFontPatch {
  latinFont?: string | null;
  eastAsiaFont?: string | null;
  complexScriptFont?: string | null;
}

export function documentFontFamilyDraftValue(
  source: DocumentFontFamilySource,
): string {
  if (source.mixed) return 'mixed';
  return cssDocumentFontFamily(source.value) ?? 'inherit';
}

export function appendDocumentFontDialogScriptFontPatch(
  patch: DocumentFontDialogScriptFontPatch,
  key: keyof DocumentFontDialogScriptFontPatch,
  source: DocumentFontFamilySource,
  draft: string,
  touched: boolean,
): void {
  if (!touched || draft === 'mixed') return;
  const value =
    draft === 'inherit' ? null : documentFontNameFromCssFamily(draft);
  if (draft !== 'inherit' && !value) return;
  if (
    source.mixed ||
    (source.value?.toLocaleLowerCase() ?? null) !==
      (value?.toLocaleLowerCase() ?? null)
  ) {
    patch[key] = value;
  }
}

export function addDocumentScriptFontValues(
  latin: Map<string, string | null>,
  eastAsia: Map<string, string | null>,
  complexScript: Map<string, string | null>,
  attributes: Record<string, unknown>,
): void {
  const fonts = parseDocumentScriptFonts(
    typeof attributes.scriptFonts === 'string'
      ? attributes.scriptFonts
      : undefined,
  );
  const fallback = documentFontNameFromCssFamily(attributes.fontFamily);
  if (!fonts) {
    addFontValue(latin, fallback);
    addFontValue(eastAsia, fallback);
    addFontValue(complexScript, fallback);
    return;
  }
  addFontValue(latin, documentScriptFontDirectFamily(fonts, 'ascii'));
  addFontValue(latin, documentScriptFontDirectFamily(fonts, 'highAnsi'));
  addFontValue(eastAsia, documentScriptFontDirectFamily(fonts, 'eastAsia'));
  addFontValue(
    complexScript,
    documentScriptFontDirectFamily(fonts, 'complexScript'),
  );
}

export function applyDocumentScriptFontPatch(
  editor: Editor,
  selection: { from: number; to: number },
  scalarPatch: Readonly<Record<string, number | string | null>>,
  scriptFontPatch: WorkDocumentScriptFontPatch,
): boolean {
  if (!validScriptFontPatch(scriptFontPatch)) return false;
  const { state } = editor;
  const textStyle = state.schema.marks.textStyle;
  if (!textStyle) return false;
  const transaction = state.tr;
  try {
    transaction.setSelection(
      TextSelection.create(transaction.doc, selection.from, selection.to),
    );
  } catch {
    return false;
  }
  if (selection.from === selection.to) {
    const marks = transaction.selection.$from.marks();
    const current = marks.find((mark) => mark.type === textStyle);
    const attributes = patchedTextStyleAttributes(
      current,
      scalarPatch,
      scriptFontPatch,
      normalizeDocumentScriptFontSlot(current?.attrs.scriptFontSlot) ??
        documentScriptFontSlotFromHint(
          parseDocumentScriptFonts(current?.attrs.scriptFonts)?.hint,
        ),
    );
    const retained = marks.filter((mark) => mark.type !== textStyle);
    if (textStyleAttributesHaveValue(attributes)) {
      retained.push(textStyle.create(attributes));
    }
    transaction.setStoredMarks(retained);
    transaction.scrollIntoView();
    editor.view.dispatch(transaction);
    editor.view.focus();
    return true;
  }

  const updates: Array<{
    from: number;
    to: number;
    current?: Mark;
    attributes: Record<string, unknown>;
  }> = [];
  transaction.doc.nodesBetween(
    selection.from,
    selection.to,
    (node, position) => {
      if (!node.isText || !node.text) return;
      const from = Math.max(selection.from, position);
      const to = Math.min(selection.to, position + node.nodeSize);
      if (to <= from) return;
      const current = node.marks.find((mark) => mark.type === textStyle);
      const currentFonts = parseDocumentScriptFonts(current?.attrs.scriptFonts);
      const nextFonts = patchDocumentScriptFonts(
        currentFonts,
        scriptFontPatch,
        current?.attrs.fontFamily,
      );
      const text = node.text.slice(from - position, to - position);
      const segments = documentScriptFontSegments(text, nextFonts?.hint);
      for (const segment of segments) {
        updates.push({
          from: from + segment.from,
          to: from + segment.to,
          current,
          attributes: patchedTextStyleAttributes(
            current,
            scalarPatch,
            scriptFontPatch,
            segment.slot,
            currentFonts,
            nextFonts,
          ),
        });
      }
    },
  );
  for (const update of updates) {
    if (update.current) {
      transaction.removeMark(update.from, update.to, update.current);
    }
    if (textStyleAttributesHaveValue(update.attributes)) {
      transaction.addMark(
        update.from,
        update.to,
        textStyle.create(update.attributes),
      );
    }
  }
  if (!transaction.docChanged) return false;
  editor.view.dispatch(transaction.scrollIntoView());
  editor.view.focus();
  return true;
}

function patchedTextStyleAttributes(
  current: Mark | undefined,
  scalarPatch: Readonly<Record<string, number | string | null>>,
  scriptFontPatch: WorkDocumentScriptFontPatch,
  slot: WorkDocumentScriptFontSlot,
  currentFonts = parseDocumentScriptFonts(current?.attrs.scriptFonts),
  nextFonts = patchDocumentScriptFonts(
    currentFonts,
    scriptFontPatch,
    current?.attrs.fontFamily,
  ),
): Record<string, unknown> {
  const attributes: Record<string, unknown> = {
    ...(current?.attrs ?? {}),
    ...scalarPatch,
  };
  const serialized = serializeDocumentScriptFonts(nextFonts);
  if (!nextFonts || !serialized) {
    attributes.scriptFonts = null;
    attributes.scriptFontSlot = null;
    attributes.fontFamily = null;
    attributes.wordLineHeightFactor = null;
    return attributes;
  }
  const currentOwnsSlot = Boolean(
    currentFonts && documentScriptFontDirectFamily(currentFonts, slot),
  );
  const nextOwnsSlot = Boolean(documentScriptFontDirectFamily(nextFonts, slot));
  const family =
    nextOwnsSlot || currentOwnsSlot
      ? documentScriptFontFamilyForRendering(
          nextFonts,
          slot,
          current?.attrs.fontFamily,
        )
      : typeof current?.attrs.fontFamily === 'string'
        ? current.attrs.fontFamily
        : documentScriptFontFamilyForRendering(nextFonts, slot);
  attributes.scriptFonts = serialized;
  attributes.scriptFontSlot = slot;
  attributes.fontFamily = family ?? null;
  attributes.wordLineHeightFactor = family
    ? documentWordLineHeightFactor(family)
    : null;
  return attributes;
}

function textStyleAttributesHaveValue(
  attributes: Readonly<Record<string, unknown>>,
): boolean {
  return Object.values(attributes).some(
    (value) => value !== null && value !== undefined && value !== '',
  );
}

function validScriptFontPatch(patch: WorkDocumentScriptFontPatch): boolean {
  return Object.values(patch).every(
    (value) => value === null || documentFontNameFromCssFamily(value) !== null,
  );
}

function addFontValue(
  values: Map<string, string | null>,
  value: string | null,
): void {
  const normalized = value?.trim() || null;
  values.set(normalized?.toLocaleLowerCase() ?? 'inherited', normalized);
}
