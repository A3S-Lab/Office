import type { WorkDocumentNode } from './work-types';

const DOCUMENT_LAZY_PREVIEW_POOL_LIMIT = 12;

export interface DocumentLazyPreviewLease {
  nodes: readonly HTMLElement[];
  shape: string;
  spacer: HTMLElement;
}

export interface DocumentLazyPreviewPool {
  available: Map<string, DocumentLazyPreviewLease[]>;
  size: number;
}

interface RenderedLazyDocumentNode {
  element: HTMLElement;
  size: number;
}

export function createDocumentLazyPreviewPool(): DocumentLazyPreviewPool {
  return { available: new Map(), size: 0 };
}

/**
 * Mounts a compact preview while recycling its element tree across chunks with
 * the same schema shape. Text and logical positions are rebound in place, so a
 * long scroll does not leave tens of thousands of detached nodes for V8 to
 * reclaim.
 */
export function mountDocumentLazyPreview(
  pool: DocumentLazyPreviewPool,
  dom: HTMLElement,
  content: readonly WorkDocumentNode[],
  firstPosition: number,
  paginationExtraHeight: number,
): DocumentLazyPreviewLease {
  const shape = lazyDocumentPreviewShape(content);
  const available = pool.available.get(shape);
  let lease = available?.pop();
  if (lease) {
    pool.size -= 1;
    if (!available?.length) pool.available.delete(shape);
  } else {
    lease = createDocumentLazyPreview(content, shape);
  }

  let position = firstPosition;
  for (let index = 0; index < content.length; index += 1) {
    const node = content[index];
    const element = lease.nodes[index];
    if (!node || !element) {
      throw new Error('The lazy document preview shape is incomplete.');
    }
    const size = bindLazyDocumentNode(node, element, position);
    if (size === null) {
      throw new Error('The lazy document preview shape changed during reuse.');
    }
    position += size;
  }

  dom.replaceChildren(...lease.nodes, lease.spacer);
  dom.dataset.documentLazyPreview = 'true';
  dom.setAttribute('contenteditable', 'false');
  updateDocumentLazyPreviewSpacer(lease, paginationExtraHeight);
  return lease;
}

export function releaseDocumentLazyPreview(
  pool: DocumentLazyPreviewPool,
  dom: HTMLElement,
  lease: DocumentLazyPreviewLease | null,
): void {
  if (!lease) return;
  dom.replaceChildren();
  if (pool.size >= DOCUMENT_LAZY_PREVIEW_POOL_LIMIT) return;
  const available = pool.available.get(lease.shape) ?? [];
  available.push(lease);
  pool.available.set(lease.shape, available);
  pool.size += 1;
}

export function updateDocumentLazyPreviewSpacer(
  lease: DocumentLazyPreviewLease | null,
  height: number,
): void {
  if (!lease) return;
  lease.spacer.style.height = `${Math.max(0, height)}px`;
}

export function documentLazyPreviewPointerPosition(
  event: Event,
  fallback: number,
): number | null {
  if (!(event instanceof MouseEvent) || event.button !== 0) return null;
  const target = event.target;
  if (!(target instanceof Element)) return null;
  const preview = target.closest<HTMLElement>(
    '[data-document-lazy-preview="true"]',
  );
  if (!preview) return null;
  const textElement = target.closest<HTMLElement>(
    '[data-document-text-position]',
  );
  const base = Number(textElement?.dataset.documentTextPosition);
  let offset = 0;
  const caret = target.ownerDocument.caretPositionFromPoint?.(
    event.clientX,
    event.clientY,
  );
  if (
    caret &&
    caret.offsetNode.nodeType === globalThis.Node.TEXT_NODE &&
    textElement?.contains(caret.offsetNode)
  ) {
    offset = Math.max(
      0,
      Math.min(caret.offset, caret.offsetNode.textContent?.length ?? 0),
    );
  }
  return Number.isSafeInteger(base) ? base + offset : fallback;
}

function createDocumentLazyPreview(
  content: readonly WorkDocumentNode[],
  shape: string,
): DocumentLazyPreviewLease {
  const nodes: HTMLElement[] = [];
  let position = 0;
  for (const node of content) {
    const rendered = createLazyDocumentNode(node, position);
    nodes.push(rendered.element);
    position += rendered.size;
  }
  const spacer = document.createElement('div');
  spacer.dataset.documentLazyPaginationSpacer = 'true';
  spacer.setAttribute('aria-hidden', 'true');
  return { nodes, shape, spacer };
}

function createLazyDocumentNode(
  node: WorkDocumentNode,
  position: number,
): RenderedLazyDocumentNode {
  if (node.type === 'text' && node.text) {
    const span = document.createElement('span');
    span.dataset.documentTextPosition = String(position);
    span.append(document.createTextNode(node.text));
    return { element: span, size: node.text.length };
  }
  if (node.type === 'paragraph') {
    const paragraph = document.createElement('p');
    const size = appendLazyDocumentChildren(node, paragraph, position + 1);
    if (!paragraph.hasChildNodes())
      paragraph.append(document.createElement('br'));
    return { element: paragraph, size: size + 2 };
  }
  if (node.type === 'table') {
    const wrapper = document.createElement('div');
    wrapper.className = 'tableWrapper';
    const table = document.createElement('table');
    bindLazyDocumentTableAttributes(table, node);
    const body = document.createElement('tbody');
    table.append(body);
    wrapper.append(table);
    const size = appendLazyDocumentChildren(node, body, position + 1);
    return { element: wrapper, size: size + 2 };
  }
  if (node.type === 'tableRow') {
    const row = document.createElement('tr');
    const size = appendLazyDocumentChildren(node, row, position + 1);
    return { element: row, size: size + 2 };
  }
  if (node.type === 'tableCell' || node.type === 'tableHeader') {
    const cell = document.createElement(
      node.type === 'tableHeader' ? 'th' : 'td',
    );
    const size = appendLazyDocumentChildren(node, cell, position + 1);
    return { element: cell, size: size + 2 };
  }
  throw new Error(`Unsupported lazy document preview node: ${node.type}`);
}

function appendLazyDocumentChildren(
  node: WorkDocumentNode,
  parent: HTMLElement,
  firstPosition: number,
): number {
  let size = 0;
  for (const child of node.content ?? []) {
    const rendered = createLazyDocumentNode(child, firstPosition + size);
    parent.append(rendered.element);
    size += rendered.size;
  }
  return size;
}

function bindLazyDocumentNode(
  node: WorkDocumentNode,
  element: HTMLElement,
  position: number,
): number | null {
  if (node.type === 'text' && node.text && element.tagName === 'SPAN') {
    const text = element.firstChild;
    if (!(text instanceof Text) || element.childNodes.length !== 1) return null;
    element.dataset.documentTextPosition = String(position);
    if (text.data !== node.text) text.data = node.text;
    return node.text.length;
  }
  if (node.type === 'paragraph' && element.tagName === 'P') {
    if (!(node.content?.length ?? 0)) {
      return element.children.length === 1 &&
        element.firstElementChild?.tagName === 'BR'
        ? 2
        : null;
    }
    const size = bindLazyDocumentChildren(node, element, position + 1);
    return size === null ? null : size + 2;
  }
  if (node.type === 'table' && element.classList.contains('tableWrapper')) {
    const table = element.firstElementChild;
    const body = table?.firstElementChild;
    if (!(table instanceof HTMLTableElement) || body?.tagName !== 'TBODY') {
      return null;
    }
    bindLazyDocumentTableAttributes(table, node);
    const size = bindLazyDocumentChildren(
      node,
      body as HTMLElement,
      position + 1,
    );
    return size === null ? null : size + 2;
  }
  if (node.type === 'tableRow' && element.tagName === 'TR') {
    const size = bindLazyDocumentChildren(node, element, position + 1);
    return size === null ? null : size + 2;
  }
  const expectedCellTag = node.type === 'tableHeader' ? 'TH' : 'TD';
  if (
    (node.type === 'tableCell' || node.type === 'tableHeader') &&
    element.tagName === expectedCellTag
  ) {
    const size = bindLazyDocumentChildren(node, element, position + 1);
    return size === null ? null : size + 2;
  }
  return null;
}

function bindLazyDocumentChildren(
  node: WorkDocumentNode,
  parent: HTMLElement,
  firstPosition: number,
): number | null {
  const children = node.content ?? [];
  if (parent.children.length !== children.length) return null;
  let size = 0;
  for (let index = 0; index < children.length; index += 1) {
    const child = children[index];
    const element = parent.children.item(index);
    if (!child || !(element instanceof HTMLElement)) return null;
    const childSize = bindLazyDocumentNode(
      child,
      element,
      firstPosition + size,
    );
    if (childSize === null) return null;
    size += childSize;
  }
  return size;
}

function bindLazyDocumentTableAttributes(
  table: HTMLTableElement,
  node: WorkDocumentNode,
): void {
  setBooleanDataset(
    table,
    'officeTableImported',
    node.attrs?.officeImported === true,
  );
  setStringDataset(table, 'documentVirtualTableId', node.attrs?.virtualTableId);
  setNumberDataset(
    table,
    'documentVirtualTableIndex',
    node.attrs?.virtualTableIndex,
  );
  setNumberDataset(
    table,
    'documentVirtualTableCount',
    node.attrs?.virtualTableCount,
  );
}

function lazyDocumentPreviewShape(
  content: readonly WorkDocumentNode[],
): string {
  const parts: string[] = [];
  for (const node of content) appendLazyDocumentPreviewShape(node, parts);
  return parts.join('');
}

function appendLazyDocumentPreviewShape(
  node: WorkDocumentNode,
  parts: string[],
): void {
  const children = node.content ?? [];
  parts.push(node.type, '[', String(children.length), '](');
  for (const child of children) appendLazyDocumentPreviewShape(child, parts);
  parts.push(')');
}

function setBooleanDataset(
  element: HTMLElement,
  name: keyof DOMStringMap,
  value: boolean,
): void {
  if (value) element.dataset[name] = 'true';
  else delete element.dataset[name];
}

function setStringDataset(
  element: HTMLElement,
  name: keyof DOMStringMap,
  value: unknown,
): void {
  if (typeof value === 'string' && value) element.dataset[name] = value;
  else delete element.dataset[name];
}

function setNumberDataset(
  element: HTMLElement,
  name: keyof DOMStringMap,
  value: unknown,
): void {
  const number = Number(value);
  if (Number.isSafeInteger(number) && number >= 0) {
    element.dataset[name] = String(number);
  } else {
    delete element.dataset[name];
  }
}
