import { Extension, type Editor } from '@tiptap/core';
import type { Node as ProseMirrorNode } from '@tiptap/pm/model';
import { Plugin, PluginKey, type EditorState } from '@tiptap/pm/state';
import { Decoration, DecorationSet } from '@tiptap/pm/view';
import type {
  OfficeKernelLayoutBlock,
  OfficeKernelLayoutBreak,
  OfficeKernelPageMetrics,
} from '../../kernel/office-kernel-protocol';
import { millimetersToPixels } from './work-document-layout';
import {
  documentSectionLayoutFromNodeAttributes,
  type DocumentSectionNodeAttributes,
} from './work-document-section';
import type { WorkDocumentSectionLayout } from './work-types';

interface DocumentPaginationPluginState {
  decorations: DecorationSet;
  revision: number;
}

interface DocumentPaginationMeta {
  kind: 'apply' | 'clear';
  revision: number;
  breaks?: DocumentPaginationVisualBreak[];
}

export interface DocumentPaginationVisualBreak extends OfficeKernelLayoutBreak {
  position: number;
  page: OfficeKernelPageMetrics;
}

export interface MeasuredDocumentLayoutBlock {
  block: OfficeKernelLayoutBlock;
  element: HTMLElement;
  from: number;
  to: number;
}

export interface DocumentPaginationSnapshot {
  blocks: MeasuredDocumentLayoutBlock[];
  unsupportedLayout: boolean;
}

const documentPaginationPluginKey =
  new PluginKey<DocumentPaginationPluginState>('documentPagination');

export const DocumentPagination = Extension.create({
  name: 'documentPagination',

  addProseMirrorPlugins() {
    return [
      new Plugin<DocumentPaginationPluginState>({
        key: documentPaginationPluginKey,
        state: {
          init: () => ({
            decorations: DecorationSet.empty,
            revision: 0,
          }),
          apply: (transaction, current, _previous, next) => {
            const meta = transaction.getMeta(documentPaginationPluginKey) as
              | DocumentPaginationMeta
              | undefined;
            if (meta?.kind === 'clear') {
              return {
                decorations: DecorationSet.empty,
                revision: meta.revision,
              };
            }
            if (meta?.kind === 'apply') {
              return {
                decorations: DecorationSet.create(
                  next.doc,
                  (meta.breaks ?? []).map(pageBreakDecoration),
                ),
                revision: meta.revision,
              };
            }
            if (transaction.docChanged) {
              return {
                decorations: DecorationSet.empty,
                revision: current.revision,
              };
            }
            return {
              ...current,
              decorations: current.decorations.map(
                transaction.mapping,
                next.doc,
              ),
            };
          },
        },
        props: {
          decorations(state) {
            return documentPaginationState(state).decorations;
          },
        },
      }),
    ];
  },
});

export function clearDocumentPagination(
  editor: Editor,
  revision: number,
): void {
  if (editor.isDestroyed) return;
  const current = documentPaginationState(editor.state);
  if (!current.decorations.find().length && current.revision === revision)
    return;
  editor.view.dispatch(
    editor.state.tr
      .setMeta(documentPaginationPluginKey, {
        kind: 'clear',
        revision,
      } satisfies DocumentPaginationMeta)
      .setMeta('addToHistory', false),
  );
}

export function applyDocumentPagination(
  editor: Editor,
  revision: number,
  breaks: DocumentPaginationVisualBreak[],
): void {
  if (editor.isDestroyed) return;
  editor.view.dispatch(
    editor.state.tr
      .setMeta(documentPaginationPluginKey, {
        kind: 'apply',
        revision,
        breaks,
      } satisfies DocumentPaginationMeta)
      .setMeta('addToHistory', false),
  );
}

export function measureDocumentLayoutBlocks(
  editor: Editor,
): DocumentPaginationSnapshot {
  const blocks: MeasuredDocumentLayoutBlock[] = [];
  let unsupportedLayout = false;
  let firstLayoutKey: string | null = null;
  editor.state.doc.forEach((section, sectionPosition) => {
    if (section.type.name !== 'documentSection') return;
    const layout = documentSectionLayoutFromNodeAttributes(
      section.attrs as Partial<DocumentSectionNodeAttributes>,
    );
    const layoutKey = documentPaginationLayoutKey(layout);
    firstLayoutKey ??= layoutKey;
    if (layout.columns.count > 1 || layoutKey !== firstLayoutKey) {
      unsupportedLayout = true;
    }
    measureSectionBlocks(editor, section, sectionPosition, layout, blocks);
  });
  return { blocks, unsupportedLayout };
}

export function documentPageMetrics(
  layout: WorkDocumentSectionLayout,
): OfficeKernelPageMetrics {
  const portrait =
    layout.pageSize === 'letter'
      ? { width: 816, height: 1056 }
      : {
          width: millimetersToPixels(210),
          height: millimetersToPixels(297),
        };
  const landscape = layout.orientation === 'landscape';
  return {
    width: landscape ? portrait.height : portrait.width,
    height: landscape ? portrait.width : portrait.height,
    marginTop: millimetersToPixels(layout.margins.top),
    marginRight: millimetersToPixels(layout.margins.right),
    marginBottom: millimetersToPixels(layout.margins.bottom),
    marginLeft: millimetersToPixels(layout.margins.left),
    headerHeight: 0,
    footerHeight: 0,
    pageGap: 28,
  };
}

export function documentPageChromeHeights(
  editor: Editor,
): Pick<OfficeKernelPageMetrics, 'headerHeight' | 'footerHeight'> {
  const page = editor.view.dom.closest<HTMLElement>('.work-document-page');
  return {
    headerHeight: outerHeight(
      page?.querySelector<HTMLElement>('.work-document-page-header') ?? null,
    ),
    footerHeight: outerHeight(
      page?.querySelector<HTMLElement>('.work-document-page-footer') ?? null,
    ),
  };
}

function documentPaginationState(
  state: EditorState,
): DocumentPaginationPluginState {
  return (
    documentPaginationPluginKey.getState(state) ?? {
      decorations: DecorationSet.empty,
      revision: 0,
    }
  );
}

function measureSectionBlocks(
  editor: Editor,
  section: ProseMirrorNode,
  sectionPosition: number,
  layout: WorkDocumentSectionLayout,
  result: MeasuredDocumentLayoutBlock[],
): void {
  const sectionBlocks: MeasuredDocumentLayoutBlock[] = [];
  section.forEach((node, offset, index) => {
    const position = sectionPosition + offset + 1;
    const element = elementForNode(editor, position);
    if (!element) return;
    sectionBlocks.push({
      block: {
        id: `section-${sectionPosition}-block-${index}-${position}`,
        height: Math.max(1, outerHeight(element)),
        breakAfter: node.type.name === 'pageBreak',
        keepTogether: shouldKeepDocumentBlockTogether(node),
        keepWithNext: node.type.name === 'heading',
      },
      element,
      from: position,
      to: position + node.nodeSize,
    });
  });

  const last = sectionBlocks.at(-1);
  if (
    last &&
    layout.breakAfter !== 'continuous' &&
    layout.breakAfter !== 'nextColumn'
  ) {
    last.block.breakAfter = true;
  }
  result.push(...sectionBlocks);
}

function elementForNode(editor: Editor, position: number): HTMLElement | null {
  const node = editor.view.nodeDOM(position);
  if (!node || node.nodeType !== Node.ELEMENT_NODE) return null;
  return node as HTMLElement;
}

function shouldKeepDocumentBlockTogether(node: ProseMirrorNode): boolean {
  return (
    node.type.name === 'table' ||
    node.type.name === 'blockquote' ||
    node.type.name === 'codeBlock' ||
    node.type.name === 'image' ||
    node.type.name === 'documentNote'
  );
}

function documentPaginationLayoutKey(
  layout: WorkDocumentSectionLayout,
): string {
  return JSON.stringify({
    footerText: layout.footerText ?? '',
    headerText: layout.headerText ?? '',
    margins: layout.margins,
    orientation: layout.orientation,
    pageChrome: layout.pageChrome ?? null,
    pageSize: layout.pageSize,
    showPageNumbers: Boolean(layout.showPageNumbers),
  });
}

function outerHeight(element: HTMLElement | null): number {
  if (!element) return 0;
  const style = getComputedStyle(element);
  return (
    element.offsetHeight +
    nonNegativePixels(style.marginTop) +
    nonNegativePixels(style.marginBottom)
  );
}

function nonNegativePixels(value: string): number {
  const parsed = Number.parseFloat(value);
  return Number.isFinite(parsed) ? Math.max(0, parsed) : 0;
}

function pageBreakDecoration(
  pageBreak: DocumentPaginationVisualBreak,
): Decoration {
  return Decoration.widget(
    pageBreak.position,
    () => {
      const element = document.createElement('div');
      element.className = 'work-document-auto-page-break';
      element.contentEditable = 'false';
      element.dataset.pageIndex = String(pageBreak.pageIndex + 1);
      element.setAttribute('aria-hidden', 'true');
      element.style.setProperty(
        '--work-document-page-spacer-height',
        `${pageBreak.spacerHeight}px`,
      );
      element.style.setProperty(
        '--work-document-page-gap-offset',
        `${
          pageBreak.remainingBodyHeight +
          pageBreak.page.marginBottom +
          pageBreak.page.footerHeight
        }px`,
      );
      element.style.setProperty(
        '--work-document-page-gap-height',
        `${pageBreak.page.pageGap}px`,
      );
      element.style.setProperty(
        '--work-document-page-margin-left',
        `${pageBreak.page.marginLeft}px`,
      );
      element.style.setProperty(
        '--work-document-page-margin-right',
        `${pageBreak.page.marginRight}px`,
      );
      return element;
    },
    {
      key: `document-page-${pageBreak.pageIndex}-${pageBreak.beforeBlockId}`,
      side: -1,
    },
  );
}
