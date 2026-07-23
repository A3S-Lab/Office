import type { Editor } from '@tiptap/core';
import { useEffect, useMemo, useRef, useState } from 'react';
import {
  createOfficeKernelClient,
  type OfficeKernelClient,
} from '../../../kernel/office-kernel-client';
import type {
  OfficeKernelLayoutResult,
  OfficeKernelPageMetrics,
} from '../../../kernel/office-kernel-protocol';
import {
  applyDocumentPagination,
  clearDocumentPagination,
  documentPageChromeHeights,
  measureDocumentLayoutBlocks,
  type MeasuredDocumentLayoutBlock,
} from '../work-document-pagination';

interface DocumentPaginationResult {
  layout: OfficeKernelLayoutResult;
  blocks: MeasuredDocumentLayoutBlock[];
  pageByBlockId: ReadonlyMap<string, number>;
}

export interface UseDocumentPaginationOptions {
  editor: Editor | null;
  enabled: boolean;
  layoutKey: string;
  page: OfficeKernelPageMetrics;
  selectionVersion: number;
  wasmUrl?: string;
}

export interface UseDocumentPaginationValue {
  currentPage: number | null;
  pageCount: number | null;
  paginating: boolean;
}

export function useDocumentPagination({
  editor,
  enabled,
  layoutKey,
  page,
  selectionVersion,
  wasmUrl,
}: UseDocumentPaginationOptions): UseDocumentPaginationValue {
  const client = useOfficeKernelClient(wasmUrl);
  const editorMounted = useEditorMounted(editor);
  const revision = useRef(0);
  const activeRequest = useRef<AbortController | null>(null);
  const [pagination, setPagination] = useState<DocumentPaginationResult | null>(
    null,
  );
  const pageKey = pageMetricsKey(page);

  useEffect(() => {
    if (
      !editor ||
      !editorMounted ||
      editor.isDestroyed ||
      !enabled ||
      !client
    ) {
      const nextRevision = ++revision.current;
      activeRequest.current?.abort();
      activeRequest.current = null;
      if (editor && editorMounted && !editor.isDestroyed) {
        const editorDom = editor.view.dom;
        clearDocumentPagination(editor, nextRevision);
        editorDom.dataset.paginationState =
          enabled && !client ? 'initializing' : 'disabled';
        if (!enabled) {
          delete editorDom.dataset.paginationBlocks;
          delete editorDom.dataset.paginationEngine;
          delete editorDom.dataset.paginationError;
          delete editorDom.dataset.paginationPages;
        }
      }
      setPagination(null);
      return;
    }

    const editorDom = editor.view.dom;
    let disposed = false;
    let frame = 0;
    let observedElements: HTMLElement[] = [];
    const observedHeights = new Map<HTMLElement, number>();
    const observer =
      typeof ResizeObserver === 'undefined'
        ? null
        : new ResizeObserver((entries) => {
            const changed = entries.some((entry) => {
              const element = entry.target as HTMLElement;
              const previous = observedHeights.get(element);
              const next = element.offsetHeight;
              observedHeights.set(element, next);
              return previous !== undefined && previous !== next;
            });
            if (changed) {
              schedule();
            }
          });

    const observeBlocks = (blocks: MeasuredDocumentLayoutBlock[]) => {
      observer?.disconnect();
      observedElements = blocks.map((block) => block.element);
      observedHeights.clear();
      for (const element of observedElements) observer?.observe(element);
      for (const element of observedElements) {
        observedHeights.set(element, element.offsetHeight);
      }
    };

    const run = async () => {
      frame = 0;
      const nextRevision = ++revision.current;
      activeRequest.current?.abort();
      const controller = new AbortController();
      activeRequest.current = controller;
      clearDocumentPagination(editor, nextRevision);
      editorDom.dataset.paginationState = 'measuring';
      delete editorDom.dataset.paginationError;
      const snapshot = measureDocumentLayoutBlocks(editor);
      editorDom.dataset.paginationBlocks = String(snapshot.blocks.length);
      observeBlocks(snapshot.blocks);
      if (snapshot.unsupportedLayout || !snapshot.blocks.length) {
        editorDom.dataset.paginationState = snapshot.unsupportedLayout
          ? 'unsupported'
          : 'empty';
        setPagination(null);
        return;
      }

      const effectivePage = {
        ...page,
        ...documentPageChromeHeights(editor),
      };
      try {
        editorDom.dataset.paginationState = 'layout';
        const layout = await client.layout(
          {
            revision: nextRevision,
            page: effectivePage,
            blocks: snapshot.blocks.map((block) => block.block),
          },
          controller.signal,
        );
        if (
          disposed ||
          controller.signal.aborted ||
          nextRevision !== revision.current ||
          editor.isDestroyed
        ) {
          return;
        }
        const blockById = new Map(
          snapshot.blocks.map((block) => [block.block.id, block] as const),
        );
        applyDocumentPagination(
          editor,
          nextRevision,
          layout.breaks.flatMap((pageBreak) => {
            const block = blockById.get(pageBreak.beforeBlockId);
            return block
              ? [{ ...pageBreak, page: effectivePage, position: block.from }]
              : [];
          }),
        );
        editorDom.dataset.paginationEngine = layout.engine;
        editorDom.dataset.paginationPages = String(layout.pages.length);
        editorDom.dataset.paginationState = 'ready';
        delete editorDom.dataset.paginationError;
        setPagination({
          layout,
          blocks: snapshot.blocks,
          pageByBlockId: new Map(
            layout.pages.flatMap((page) =>
              page.placements.map(
                (placement) => [placement.blockId, page.index + 1] as const,
              ),
            ),
          ),
        });
      } catch (error) {
        if (
          disposed ||
          controller.signal.aborted ||
          (error instanceof DOMException && error.name === 'AbortError')
        ) {
          return;
        }
        editorDom.dataset.paginationState = 'error';
        editorDom.dataset.paginationError =
          error instanceof Error
            ? `${error.name}: ${error.message}`
            : 'UnknownError';
        clearDocumentPagination(editor, nextRevision);
        setPagination(null);
      }
    };

    const schedule = () => {
      if (disposed || frame) return;
      frame = requestAnimationFrame(() => {
        frame = requestAnimationFrame(() => void run());
      });
    };
    const handleDocumentUpdate = () => schedule();
    const handleLoadedAsset = () => schedule();
    const fonts = document.fonts;
    editor.on('update', handleDocumentUpdate);
    editorDom.addEventListener('load', handleLoadedAsset, true);
    fonts?.addEventListener('loadingdone', schedule);
    window.addEventListener('resize', schedule);
    schedule();

    return () => {
      disposed = true;
      if (frame) cancelAnimationFrame(frame);
      observer?.disconnect();
      editor.off('update', handleDocumentUpdate);
      editorDom.removeEventListener('load', handleLoadedAsset, true);
      fonts?.removeEventListener('loadingdone', schedule);
      window.removeEventListener('resize', schedule);
      activeRequest.current?.abort();
      activeRequest.current = null;
    };
  }, [client, editor, editorMounted, enabled, layoutKey, pageKey]);

  return useMemo(
    () => ({
      currentPage: pagination
        ? pageForPosition(pagination, editor?.state.selection.from ?? 0)
        : null,
      pageCount: pagination
        ? Math.max(1, pagination.layout.pages.length)
        : null,
      paginating: enabled && pagination === null,
    }),
    [editor, enabled, pagination, selectionVersion],
  );
}

function useEditorMounted(editor: Editor | null): boolean {
  const [mounted, setMounted] = useState(() =>
    Boolean(editor && !editor.isDestroyed),
  );

  useEffect(() => {
    if (!editor) {
      setMounted(false);
      return;
    }

    const handleMount = () => setMounted(true);
    const handleUnmount = () => setMounted(false);
    setMounted(!editor.isDestroyed);
    editor.on('mount', handleMount);
    editor.on('unmount', handleUnmount);
    return () => {
      editor.off('mount', handleMount);
      editor.off('unmount', handleUnmount);
    };
  }, [editor]);

  return mounted;
}

function useOfficeKernelClient(
  wasmUrl: string | undefined,
): OfficeKernelClient | null {
  const [client, setClient] = useState<OfficeKernelClient | null>(null);
  useEffect(() => {
    const next = createOfficeKernelClient(wasmUrl);
    setClient(next);
    return () => next.dispose();
  }, [wasmUrl]);
  return client;
}

function pageForPosition(
  pagination: DocumentPaginationResult,
  position: number,
): number {
  let lower = 0;
  let upper = pagination.blocks.length - 1;
  let containing = pagination.blocks[0];
  while (lower <= upper) {
    const middle = Math.floor((lower + upper) / 2);
    const candidate = pagination.blocks[middle];
    if (candidate.from <= position) {
      containing = candidate;
      lower = middle + 1;
    } else {
      upper = middle - 1;
    }
  }
  return containing
    ? (pagination.pageByBlockId.get(containing.block.id) ?? 1)
    : 1;
}

function pageMetricsKey(page: OfficeKernelPageMetrics): string {
  return [
    page.width,
    page.height,
    page.marginTop,
    page.marginRight,
    page.marginBottom,
    page.marginLeft,
    page.headerHeight,
    page.footerHeight,
    page.pageGap,
  ].join(':');
}
