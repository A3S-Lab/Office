import type { Editor } from '@tiptap/core';
import { NodeSelection } from '@tiptap/pm/state';

/** Collect field ids covered by the current selection (WPS Shift+F9 scope). */
export function selectedDocumentFieldIds(editor: Editor): string[] {
  if (editor.isDestroyed) return [];
  const { doc, selection } = editor.state;
  if (
    selection instanceof NodeSelection &&
    selection.node.type.name === 'documentField'
  ) {
    const id = fieldId(selection.node.attrs.id);
    return id ? [id] : [];
  }

  const ids: string[] = [];
  const seen = new Set<string>();
  doc.nodesBetween(selection.from, selection.to, (node) => {
    if (node.type.name !== 'documentField') return;
    const id = fieldId(node.attrs.id);
    if (!id || seen.has(id)) return;
    seen.add(id);
    ids.push(id);
  });
  return ids;
}

/** XOR overrides relative to the document-wide Alt+F9 field-codes mode. */
export function toggleDocumentFieldCodeOverrides(
  current: ReadonlySet<string>,
  fieldIds: readonly string[],
): Set<string> {
  const next = new Set(current);
  for (const id of fieldIds) {
    if (!id) continue;
    if (next.has(id)) next.delete(id);
    else next.add(id);
  }
  return next;
}

/** Effective field-code visibility for one field (WPS Alt+F9 ⊕ Shift+F9). */
export function documentFieldShowsCode(
  fieldIdValue: string,
  showFieldCodes: boolean,
  overrides: ReadonlySet<string>,
): boolean {
  if (!fieldIdValue) return showFieldCodes;
  return overrides.has(fieldIdValue) ? !showFieldCodes : showFieldCodes;
}

export function syncDocumentFieldCodeClasses(
  root: ParentNode | null | undefined,
  showFieldCodes: boolean,
  overrides: ReadonlySet<string>,
): { applied: number; showing: number } {
  if (!root || typeof root.querySelectorAll !== 'function') {
    return { applied: 0, showing: 0 };
  }
  const fields = root.querySelectorAll('.work-document-field[data-field-id]');
  let showing = 0;
  for (const node of fields) {
    if (!(node instanceof HTMLElement)) continue;
    const id = fieldId(node.dataset.fieldId);
    const show = documentFieldShowsCode(id, showFieldCodes, overrides);
    node.classList.toggle('show-field-code', show && !showFieldCodes);
    node.classList.toggle('show-field-code-result', !show && showFieldCodes);
    if (show) showing += 1;
  }
  return { applied: fields.length, showing };
}

function fieldId(value: unknown): string {
  return typeof value === 'string' ? value.trim() : '';
}
