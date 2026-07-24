import { mergeAttributes, Node as TiptapNode } from '@tiptap/core';
import { Plugin } from '@tiptap/pm/state';
import type { EditorView } from '@tiptap/pm/view';
import {
  DEFAULT_DOCUMENT_TAB_INTERVAL_PX,
  type DocumentTabAlignment,
  type DocumentTabLeader,
  type DocumentTabStop,
  normalizeDocumentTabStops,
} from './work-document-tab-stops';

const DOCUMENT_TAB_SELECTOR = 'span[data-document-tab]';
const DOCUMENT_TAB_BLOCK_SELECTOR = 'p, h1, h2, h3, h4, h5, h6';
const MINIMUM_DOCUMENT_TAB_WIDTH_PX = 1;

export const DocumentTab = TiptapNode.create({
  name: 'documentTab',
  group: 'inline',
  inline: true,
  atom: true,
  selectable: false,

  parseHTML() {
    return [{ tag: DOCUMENT_TAB_SELECTOR }];
  },

  renderHTML({ HTMLAttributes }) {
    return [
      'span',
      mergeAttributes(HTMLAttributes, {
        'data-document-tab': 'true',
        class: 'work-document-tab',
        role: 'separator',
        'aria-label': '制表符',
      }),
    ];
  },

  addNodeView() {
    return () => {
      const dom = document.createElement('span');
      dom.className = 'work-document-tab';
      dom.dataset.documentTab = 'true';
      dom.setAttribute('role', 'separator');
      dom.setAttribute('aria-label', '制表符');
      dom.contentEditable = 'false';
      return {
        dom,
        ignoreMutation: () => true,
      };
    };
  },

  addKeyboardShortcuts() {
    return {
      Tab: () => {
        const position = this.editor.state.selection.$from;
        if (
          selectionIsInsideTableCell(position) ||
          selectionIsInsideListItem(position)
        ) {
          return false;
        }
        return this.editor
          .chain()
          .focus()
          .insertContent({ type: this.name })
          .run();
      },
    };
  },

  addProseMirrorPlugins() {
    return [
      new Plugin({
        view: (view) => createDocumentTabLayoutView(view),
      }),
    ];
  },
});

export function layoutDocumentTabs(root: HTMLElement): void {
  const tabs = Array.from(
    root.querySelectorAll<HTMLElement>(DOCUMENT_TAB_SELECTOR),
  );
  for (const tab of tabs) tab.style.width = '0px';

  for (const tab of tabs) {
    const block = tab.closest<HTMLElement>(DOCUMENT_TAB_BLOCK_SELECTOR);
    if (!block || !root.contains(block)) continue;
    const tabStops = normalizeDocumentTabStops(
      block.dataset.officeTabStops ?? '',
    );
    const blockRectangle = block.getBoundingClientRect();
    const tabRectangle = tab.getBoundingClientRect();
    const layoutScale = documentElementLayoutScale(block, blockRectangle);
    const leftIndent = paragraphLeftIndent(block);
    const currentPosition =
      (tabRectangle.left - blockRectangle.left) / layoutScale + leftIndent;
    const stop = nextDocumentTabStop(tabStops, currentPosition);
    const offset = documentTabAlignmentOffset(
      tab,
      block,
      stop.alignment,
      layoutScale,
    );
    const width = Math.max(
      MINIMUM_DOCUMENT_TAB_WIDTH_PX,
      stop.position - currentPosition - offset,
    );
    tab.style.width = `${roundLayoutValue(width)}px`;
    tab.dataset.documentTabAlignment = stop.alignment;
    tab.dataset.documentTabLeader = stop.leader;
  }
}

export function nextDocumentTabStop(
  tabStops: readonly DocumentTabStop[],
  currentPosition: number,
): DocumentTabStop {
  const current = Math.max(0, currentPosition);
  const custom = normalizeDocumentTabStops(tabStops).find(
    (stop) => stop.position > current + 0.5,
  );
  const defaultPosition =
    (Math.floor(current / DEFAULT_DOCUMENT_TAB_INTERVAL_PX) + 1) *
    DEFAULT_DOCUMENT_TAB_INTERVAL_PX;
  if (custom && custom.position <= defaultPosition + 0.5) return custom;
  return {
    position: defaultPosition,
    alignment: 'left',
    leader: 'none',
  };
}

function createDocumentTabLayoutView(view: EditorView) {
  let frame: number | null = null;
  const schedule = (): void => {
    if (frame !== null) cancelAnimationFrame(frame);
    frame = requestAnimationFrame(() => {
      frame = null;
      if (view.dom.isConnected) layoutDocumentTabs(view.dom);
    });
  };
  const resizeObserver =
    typeof ResizeObserver === 'undefined'
      ? null
      : new ResizeObserver(() => schedule());
  resizeObserver?.observe(view.dom);
  window.addEventListener('resize', schedule);
  void document.fonts?.ready.then(schedule);
  schedule();
  return {
    update: schedule,
    destroy: () => {
      if (frame !== null) cancelAnimationFrame(frame);
      resizeObserver?.disconnect();
      window.removeEventListener('resize', schedule);
    },
  };
}

function selectionIsInsideTableCell(position: {
  depth: number;
  node: (depth: number) => { type: { name: string } };
}): boolean {
  for (let depth = position.depth; depth > 0; depth -= 1) {
    const name = position.node(depth).type.name;
    if (name === 'tableCell' || name === 'tableHeader') return true;
  }
  return false;
}

function selectionIsInsideListItem(position: {
  depth: number;
  node: (depth: number) => { type: { name: string } };
}): boolean {
  for (let depth = position.depth; depth > 0; depth -= 1) {
    if (position.node(depth).type.name === 'listItem') return true;
  }
  return false;
}

function paragraphLeftIndent(block: HTMLElement): number {
  const level = Number(block.dataset.officeIndentLevel);
  if (Number.isFinite(level) && level > 0) return level * 24;
  const pixels = Number.parseFloat(block.style.marginLeft);
  return Number.isFinite(pixels) ? Math.max(0, pixels) : 0;
}

function documentTabAlignmentOffset(
  tab: HTMLElement,
  block: HTMLElement,
  alignment: DocumentTabAlignment,
  layoutScale: number,
): number {
  if (alignment === 'left') return 0;
  const width =
    followingSegmentWidth(tab, block, alignment === 'decimal') / layoutScale;
  return alignment === 'center' ? width / 2 : width;
}

function documentElementLayoutScale(
  element: HTMLElement,
  rectangle: DOMRect,
): number {
  const layoutWidth = element.offsetWidth;
  if (layoutWidth <= 0 || rectangle.width <= 0) return 1;
  const scale = rectangle.width / layoutWidth;
  return Number.isFinite(scale) && scale > 0 ? scale : 1;
}

function followingSegmentWidth(
  tab: HTMLElement,
  block: HTMLElement,
  stopAtDecimal: boolean,
): number {
  const document = block.ownerDocument;
  const range = document.createRange();
  range.setStartAfter(tab);
  const nextTab = nextTabElement(tab, block);
  const decimal = stopAtDecimal
    ? firstDecimalPosition(tab, block, nextTab)
    : null;
  if (decimal) range.setEnd(decimal.node, decimal.offset);
  else if (nextTab) range.setEndBefore(nextTab);
  else range.setEnd(block, block.childNodes.length);

  const rects =
    typeof range.getClientRects === 'function'
      ? Array.from(range.getClientRects())
      : [];
  const first = rects[0];
  if (first && Number.isFinite(first.width)) return Math.max(0, first.width);
  if (typeof range.getBoundingClientRect !== 'function') return 0;
  const rectangle = range.getBoundingClientRect();
  return Number.isFinite(rectangle.width) ? Math.max(0, rectangle.width) : 0;
}

function nextTabElement(
  current: HTMLElement,
  block: HTMLElement,
): HTMLElement | null {
  const tabs = Array.from(
    block.querySelectorAll<HTMLElement>(DOCUMENT_TAB_SELECTOR),
  );
  const index = tabs.indexOf(current);
  return index >= 0 ? (tabs[index + 1] ?? null) : null;
}

function firstDecimalPosition(
  current: HTMLElement,
  block: HTMLElement,
  nextTab: HTMLElement | null,
): { node: Text; offset: number } | null {
  const walker = block.ownerDocument.createTreeWalker(
    block,
    NodeFilter.SHOW_TEXT,
  );
  let afterCurrent = false;
  while (walker.nextNode()) {
    const node = walker.currentNode as Text;
    if (!afterCurrent) {
      afterCurrent = Boolean(
        current.compareDocumentPosition(node) &
          globalThis.Node.DOCUMENT_POSITION_FOLLOWING,
      );
      if (!afterCurrent) continue;
    }
    if (
      nextTab &&
      !(
        nextTab.compareDocumentPosition(node) &
        globalThis.Node.DOCUMENT_POSITION_PRECEDING
      )
    ) {
      return null;
    }
    const offset = node.data.search(/[.,，。]/);
    if (offset >= 0) return { node, offset };
  }
  return null;
}

function roundLayoutValue(value: number): number {
  return Math.round(value * 100) / 100;
}

export function documentTabLeaderLabel(leader: DocumentTabLeader): string {
  if (leader === 'dot') return '点线前导符';
  if (leader === 'hyphen') return '短横线前导符';
  if (leader === 'underscore') return '下划线前导符';
  if (leader === 'middleDot') return '居中点前导符';
  return '无前导符';
}
