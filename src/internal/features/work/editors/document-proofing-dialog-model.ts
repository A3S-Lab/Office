import type { Editor } from '@tiptap/core';
import type { Mark } from '@tiptap/pm/model';
import { TextSelection } from '@tiptap/pm/state';
import {
  normalizeDocumentLanguageTag,
  normalizeDocumentNoProof,
  parseDocumentProofingLanguages,
  patchDocumentProofingLanguages,
  serializeDocumentProofingLanguages,
  type WorkDocumentProofingLanguageSlot,
} from '../work-document-proofing';

export interface DocumentProofingDialogValue<T> {
  mixed: boolean;
  value: T | null;
}

export interface DocumentProofingDialogSource {
  latin: DocumentProofingDialogValue<string>;
  eastAsia: DocumentProofingDialogValue<string>;
  bidi: DocumentProofingDialogValue<string>;
  noProof: DocumentProofingDialogValue<boolean>;
  selectedCharacters: number;
}

export interface DocumentProofingDialogPatch {
  languages?: Partial<Record<WorkDocumentProofingLanguageSlot, string | null>>;
  noProof?: boolean | null;
}

const PROOFING_LANGUAGE_SLOTS = [
  'latin',
  'eastAsia',
  'bidi',
] as const satisfies readonly WorkDocumentProofingLanguageSlot[];
const PROOFING_LANGUAGE_SLOT_SET = new Set<string>(PROOFING_LANGUAGE_SLOTS);

export function documentProofingDialogSource(
  editor: Editor,
): DocumentProofingDialogSource {
  const latin = new Map<string, string | null>();
  const eastAsia = new Map<string, string | null>();
  const bidi = new Map<string, string | null>();
  const noProof = new Map<string, boolean | null>();
  const { doc, selection } = editor.state;
  const addValues = (attributes: Readonly<Record<string, unknown>>) => {
    const languages = parseDocumentProofingLanguages(
      attributes.proofingLanguages,
    );
    addSelectedValue(latin, languages?.latin ?? null);
    addSelectedValue(eastAsia, languages?.eastAsia ?? null);
    addSelectedValue(bidi, languages?.bidi ?? null);
    addSelectedValue(noProof, normalizeDocumentNoProof(attributes.noProof));
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
  if (!latin.size) addValues(editor.getAttributes('textStyle'));

  const selectedText = selection.empty
    ? ''
    : doc.textBetween(selection.from, selection.to, '', '');
  return {
    latin: selectedValue(latin),
    eastAsia: selectedValue(eastAsia),
    bidi: selectedValue(bidi),
    noProof: selectedValue(noProof),
    selectedCharacters: Array.from(selectedText).length,
  };
}

export function applyDocumentProofingDialogPatch(
  editor: Editor,
  selection: { from: number; to: number },
  patch: DocumentProofingDialogPatch,
): boolean {
  if (!validDocumentProofingDialogPatch(patch) || editor.isDestroyed) {
    return false;
  }
  const maximum = editor.state.doc.content.size;
  if (
    !Number.isSafeInteger(selection.from) ||
    !Number.isSafeInteger(selection.to) ||
    selection.from < 0 ||
    selection.to < selection.from ||
    selection.to > maximum
  ) {
    return false;
  }
  const textStyle = editor.state.schema.marks.textStyle;
  if (!textStyle) return false;

  const transaction = editor.state.tr;
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
    const attributes = patchedProofingAttributes(current, patch);
    if (!attributes) return false;
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
      const attributes = patchedProofingAttributes(current, patch);
      if (attributes) updates.push({ from, to, current, attributes });
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

function patchedProofingAttributes(
  current: Mark | undefined,
  patch: DocumentProofingDialogPatch,
): Record<string, unknown> | null {
  const attributes: Record<string, unknown> = { ...(current?.attrs ?? {}) };
  let changed = false;
  if (patch.languages) {
    const currentLanguages = parseDocumentProofingLanguages(
      current?.attrs.proofingLanguages,
    );
    const nextLanguages = patchDocumentProofingLanguages(
      currentLanguages,
      patch.languages,
    );
    const currentSerialized =
      serializeDocumentProofingLanguages(currentLanguages) ?? null;
    const nextSerialized =
      serializeDocumentProofingLanguages(nextLanguages) ?? null;
    if (currentSerialized !== nextSerialized) {
      attributes.proofingLanguages = nextSerialized;
      changed = true;
    }
  }
  if (patch.noProof !== undefined) {
    const currentNoProof = normalizeDocumentNoProof(current?.attrs.noProof);
    if (currentNoProof !== patch.noProof) {
      attributes.noProof = patch.noProof;
      changed = true;
    }
  }
  return changed ? attributes : null;
}

function validDocumentProofingDialogPatch(
  patch: DocumentProofingDialogPatch,
): boolean {
  if (!isRecord(patch)) return false;
  const keys = Object.keys(patch);
  if (
    !keys.length ||
    keys.some((key) => key !== 'languages' && key !== 'noProof')
  ) {
    return false;
  }
  if (
    patch.noProof !== undefined &&
    patch.noProof !== null &&
    normalizeDocumentNoProof(patch.noProof) === null
  ) {
    return false;
  }
  if (patch.languages !== undefined) {
    if (!isRecord(patch.languages)) return false;
    const languageEntries = Object.entries(patch.languages);
    if (
      !languageEntries.length ||
      languageEntries.some(
        ([slot, language]) =>
          !PROOFING_LANGUAGE_SLOT_SET.has(slot) ||
          (language !== null &&
            normalizeDocumentLanguageTag(language) === null),
      )
    ) {
      return false;
    }
  }
  return patch.languages !== undefined || patch.noProof !== undefined;
}

function addSelectedValue<T>(values: Map<string, T | null>, value: T | null) {
  values.set(value === null ? 'inherited' : String(value), value);
}

function selectedValue<T>(values: Map<string, T | null>): {
  mixed: boolean;
  value: T | null;
} {
  if (values.size !== 1) return { mixed: true, value: null };
  return { mixed: false, value: values.values().next().value ?? null };
}

function textStyleAttributesHaveValue(
  attributes: Readonly<Record<string, unknown>>,
): boolean {
  return Object.values(attributes).some(
    (value) => value !== null && value !== undefined && value !== '',
  );
}

function isRecord(source: unknown): source is Record<string, unknown> {
  return (
    typeof source === 'object' && source !== null && !Array.isArray(source)
  );
}
