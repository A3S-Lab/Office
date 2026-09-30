import type { NodeViewRendererProps } from '@tiptap/core';
import type { Node as ProseMirrorNode } from '@tiptap/pm/model';
import { Fragment } from '@tiptap/pm/model';
import type { EditorView, ViewMutationRecord } from '@tiptap/pm/view';
import { officeMessage, resolveOfficeMessages } from '../../i18n/office-locale';
import {
  contentControlDomAttributes,
  contentControlLocksContent,
  DOCUMENT_CONTENT_CONTROL_MUTATION_META,
  normalizeDocumentContentControlProperties,
  type WorkDocumentContentControlProperties,
} from './work-document-content-control';

const CARET_CLASS = 'work-document-content-control-caret';
const PLACEHOLDER_CLASS = 'work-document-content-control-placeholder';
const VALUE_CLASS = 'work-document-content-control-value';

interface ContentControlDomObserver {
  flush?: () => void;
  start: () => void;
  stop: () => void;
}

/**
 * Editable inline controls need a real caret target. An empty inline-block
 * whose only label is a CSS `::after` lets WebKit paint typed characters in
 * the box without putting them in the document, so a later save writes an
 * empty `w:t` or fails outright.
 */
export function createDocumentContentControlNodeView({
  editor,
  node,
  getPos,
}: NodeViewRendererProps): {
  dom: HTMLElement;
  contentDOM: HTMLElement;
  update: (updated: ProseMirrorNode) => boolean;
  ignoreMutation: (mutation: ViewMutationRecord) => boolean;
} {
  const dom = document.createElement('span');
  const contentDOM = document.createElement('span');
  contentDOM.className = VALUE_CLASS;
  const placeholder = document.createElement('span');
  placeholder.className = PLACEHOLDER_CLASS;
  placeholder.dataset.contentControlPlaceholder = 'true';
  placeholder.contentEditable = 'false';
  placeholder.setAttribute('aria-hidden', 'true');
  placeholder.textContent = officeMessage(
    resolveOfficeMessages(),
    'document.contentControl.placeholder',
  );
  dom.append(contentDOM, placeholder);

  let current = node;
  const sync = (updated: ProseMirrorNode) => {
    current = updated;
    const properties = normalizeDocumentContentControlProperties(updated.attrs);
    applyContentControlChrome(dom, contentDOM, properties);
    const showPlaceholder = showsTextPlaceholder(properties, updated);
    placeholder.hidden = !showPlaceholder;
    if (showPlaceholder) ensureCaret(contentDOM);
    else removeCarets(contentDOM);
  };
  sync(node);

  const commitVisibleText = (visible: string) => {
    if (!visible || editor.isDestroyed) return;
    const observer = domObserver(editor.view);
    try {
      observer?.flush?.();
    } catch {
      // WebKit can leave a DOM shape this schema cannot parse. The captured
      // text is still the value the user saw in the control.
    }
    if (editor.isDestroyed) return;
    const position = resolveContentControlPosition(getPos);
    if (position === null) return;
    const live = editor.state.doc.nodeAt(position);
    if (!live || live.type.name !== 'documentContentControl') return;
    const properties = normalizeDocumentContentControlProperties(live.attrs);
    if (contentControlLocksContent(properties.lock)) return;
    if (visible === live.textContent) return;
    let text: ProseMirrorNode;
    try {
      text = editor.schema.text(visible);
    } catch {
      return;
    }
    if (!live.type.validContent(Fragment.from(text))) return;
    const replacement = live.type.create(live.attrs, text);
    const transaction = editor.state.tr.replaceWith(
      position,
      position + live.nodeSize,
      replacement,
    );
    transaction.setMeta(DOCUMENT_CONTENT_CONTROL_MUTATION_META, true);
    observer?.stop();
    try {
      editor.view.dispatch(transaction);
    } finally {
      observer?.start();
    }
  };

  const scheduleVisibleTextCommit = () => {
    const visible = readContentControlDomText(contentDOM);
    queueMicrotask(() => commitVisibleText(visible));
  };

  contentDOM.addEventListener('input', scheduleVisibleTextCommit);
  contentDOM.addEventListener('focusout', (event) => {
    const next = event.relatedTarget;
    if (next instanceof Node && contentDOM.contains(next)) return;
    scheduleVisibleTextCommit();
  });

  return {
    dom,
    contentDOM,
    update(updated) {
      if (updated.type !== current.type) return false;
      sync(updated);
      return true;
    },
    ignoreMutation(mutation) {
      if (
        mutation.target === placeholder ||
        placeholder.contains(mutation.target)
      ) {
        return true;
      }
      return mutationIsCaretOnly(mutation);
    },
  };
}

function applyContentControlChrome(
  dom: HTMLElement,
  contentDOM: HTMLElement,
  properties: WorkDocumentContentControlProperties,
): void {
  const label = properties.alias || properties.tag || '内容控件';
  const locked = contentControlLocksContent(properties.lock);
  for (const [name, value] of Object.entries(
    contentControlDomAttributes(properties),
  )) {
    if (value === undefined) dom.removeAttribute(name);
    else dom.setAttribute(name, value);
  }
  dom.classList.add('work-document-content-control');
  dom.setAttribute('role', contentControlRole(properties));
  dom.setAttribute('aria-label', label);
  if (properties.type === 'checkbox') {
    dom.setAttribute('aria-checked', properties.checked ? 'true' : 'false');
  } else {
    dom.removeAttribute('aria-checked');
  }
  contentDOM.contentEditable = locked ? 'false' : 'true';
}

function contentControlRole(
  properties: WorkDocumentContentControlProperties,
): string {
  if (properties.type === 'checkbox') return 'checkbox';
  if (properties.type === 'dropDownList') return 'listbox';
  if (properties.type === 'comboBox') return 'combobox';
  return 'textbox';
}

function showsTextPlaceholder(
  properties: WorkDocumentContentControlProperties,
  node: ProseMirrorNode,
): boolean {
  return (
    (properties.type === 'text' || properties.type === 'richText') &&
    node.textContent.length === 0
  );
}

function ensureCaret(contentDOM: HTMLElement): void {
  if (contentDOM.querySelector(`br.${CARET_CLASS}`)) return;
  const caret = document.createElement('br');
  caret.className = CARET_CLASS;
  contentDOM.append(caret);
}

function removeCarets(contentDOM: HTMLElement): void {
  for (const caret of contentDOM.querySelectorAll(`br.${CARET_CLASS}`)) {
    caret.remove();
  }
}

function readContentControlDomText(contentDOM: HTMLElement): string {
  let value = '';
  const visit = (node: Node) => {
    if (node instanceof HTMLBRElement) {
      if (!node.classList.contains(CARET_CLASS)) value += '\n';
      return;
    }
    if (node.nodeType === Node.TEXT_NODE) {
      value += node.textContent ?? '';
      return;
    }
    for (const child of node.childNodes) visit(child);
  };
  for (const child of contentDOM.childNodes) visit(child);
  return value.replaceAll('\u200b', '');
}

function mutationIsCaretOnly(mutation: ViewMutationRecord): boolean {
  if (mutation.type === 'attributes') {
    return (
      mutation.target instanceof Element &&
      mutation.target.classList.contains(CARET_CLASS)
    );
  }
  if (mutation.type !== 'childList') return false;
  const nodes = [...mutation.addedNodes, ...mutation.removedNodes];
  return (
    nodes.length > 0 &&
    nodes.every(
      (node) =>
        node instanceof HTMLBRElement && node.classList.contains(CARET_CLASS),
    )
  );
}

function domObserver(view: EditorView): ContentControlDomObserver | null {
  const observer = (
    view as EditorView & { domObserver?: ContentControlDomObserver }
  ).domObserver;
  if (
    !observer ||
    typeof observer.stop !== 'function' ||
    typeof observer.start !== 'function'
  ) {
    return null;
  }
  return observer;
}

function resolveContentControlPosition(
  getPos: () => number | undefined,
): number | null {
  try {
    const position = getPos();
    return typeof position === 'number' ? position : null;
  } catch {
    return null;
  }
}
