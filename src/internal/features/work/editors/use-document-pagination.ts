import type { Editor, EditorEvents } from '@tiptap/core';
import type { Transaction } from '@tiptap/pm/state';
import { useEffect, useMemo, useRef, useState } from 'react';
import {
  createOfficeKernelClient,
  type OfficeKernelClient,
} from '../../../kernel/office-kernel-client';
import type {
  OfficeKernelLayoutResult,
  OfficeKernelPageMetrics,
  OfficeKernelTextLayoutParagraphResult,
} from '../../../kernel/office-kernel-protocol';
import {
  documentLayoutFontKey,
  type WorkDocumentLayoutFont,
} from '../work-document-fonts';
import {
  mergeIncrementalDocumentLayout,
  planIncrementalDocumentLayout,
} from '../work-document-incremental-layout';
import {
  applyDocumentPagination,
  clearDocumentPagination,
  collectDocumentTextLayoutParagraphs,
  documentPageChromeHeights,
  measureDocumentLayoutBlocks,
  type DocumentPaginationSnapshot,
  type DocumentPaginationVisualPageChrome,
  type MeasuredDocumentLayoutBlock,
} from '../work-document-pagination';
import {
  resolveDocumentPageChrome,
  type ResolvedDocumentPageChrome,
} from '../work-document-page-chrome';
import type { WorkDocumentSectionLayout } from '../work-types';

export interface DocumentPaginationResult {
  layout: OfficeKernelLayoutResult;
  blocks: MeasuredDocumentLayoutBlock[];
  pageByBlockId: ReadonlyMap<string, number>;
  pages: DocumentPaginationPageDescriptor[];
}

export interface DocumentPaginationPageDescriptor {
  pageIndex: number;
  physicalPage: number;
  pageNumber: number;
  sectionPage: number;
  sectionId: string;
  sectionIndex: number;
  layout: WorkDocumentSectionLayout;
  pageChrome: ResolvedDocumentPageChrome;
}

export interface UseDocumentPaginationOptions {
  editor: Editor | null;
  documentRevision: number;
  enabled: boolean;
  layoutKey: string;
  page: OfficeKernelPageMetrics;
  selectionVersion: number;
  wasmUrl?: string;
  layoutFonts: readonly WorkDocumentLayoutFont[];
  loadedLayoutFontIds: ReadonlySet<string>;
}

export interface UseDocumentPaginationValue {
  currentPage: number | null;
  currentPageDescriptor: DocumentPaginationPageDescriptor | null;
  pageCount: number | null;
  pages: readonly DocumentPaginationPageDescriptor[];
  paginating: boolean;
}

export function useDocumentPagination({
  editor,
  documentRevision,
  enabled,
  layoutKey,
  page,
  selectionVersion,
  wasmUrl,
  layoutFonts,
  loadedLayoutFontIds,
}: UseDocumentPaginationOptions): UseDocumentPaginationValue {
  const client = useOfficeKernelClient(wasmUrl, layoutFonts);
  const editorMounted = useEditorMounted(editor);
  const revision = useRef(0);
  const activeRequest = useRef<AbortController | null>(null);
  const measurementCache = useRef<DocumentPaginationSnapshot | null>(null);
  const paginationCache = useRef<DocumentPaginationResult | null>(null);
  const measurementEditor = useRef<Editor | null>(null);
  const measurementLayoutKey = useRef('');
  const dirtyMeasurementFrom = useRef(0);
  const observedDocumentRevision = useRef(documentRevision);
  const [pagination, setPagination] = useState<DocumentPaginationResult | null>(
    null,
  );
  const pageKey = pageMetricsKey(page);
  const layoutFontKey = documentLayoutFontKey(layoutFonts);
  const loadedLayoutFontKey = [...loadedLayoutFontIds].sort().join('\u0000');

  useEffect(() => {
    const measurementKey = `${layoutKey}:${pageKey}:${layoutFontKey}:${loadedLayoutFontKey}`;
    if (
      measurementEditor.current !== editor ||
      measurementLayoutKey.current !== measurementKey
    ) {
      measurementCache.current = null;
      paginationCache.current = null;
      measurementEditor.current = editor;
      measurementLayoutKey.current = measurementKey;
      dirtyMeasurementFrom.current = 0;
    }
    if (observedDocumentRevision.current !== documentRevision) {
      if (!Number.isFinite(dirtyMeasurementFrom.current)) {
        dirtyMeasurementFrom.current = 0;
      }
      observedDocumentRevision.current = documentRevision;
    }
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
          delete editorDom.dataset.paginationDocumentRevision;
          delete editorDom.dataset.paginationEngine;
          delete editorDom.dataset.paginationError;
          delete editorDom.dataset.paginationFallbackGlyphs;
          delete editorDom.dataset.paginationFlows;
          delete editorDom.dataset.paginationMeasuredBlocks;
          delete editorDom.dataset.paginationLaidOutBlocks;
          delete editorDom.dataset.paginationPages;
          delete editorDom.dataset.paginationReusedBlocks;
          delete editorDom.dataset.paginationReusedPages;
          delete editorDom.dataset.paginationShapedParagraphs;
          delete editorDom.dataset.paginationShapedRuns;
          delete editorDom.dataset.paginationTextCandidates;
          delete editorDom.dataset.paginationTextEngine;
          delete editorDom.dataset.paginationTextRuns;
          delete editorDom.dataset.paginationUnsupportedText;
        }
      }
      if (!enabled) {
        measurementCache.current = null;
        paginationCache.current = null;
        dirtyMeasurementFrom.current = 0;
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
            let earliestChangedBlock = Number.POSITIVE_INFINITY;
            let changed = false;
            for (const entry of entries) {
              const element = entry.target as HTMLElement;
              const previous = observedHeights.get(element);
              const next = element.offsetHeight;
              observedHeights.set(element, next);
              const resized = previous !== undefined && previous !== next;
              changed ||= resized;
              if (resized) {
                for (const block of measurementCache.current?.blocks ?? []) {
                  if (block.element === element) {
                    earliestChangedBlock = Math.min(
                      earliestChangedBlock,
                      block.from,
                    );
                  }
                }
              }
            }
            if (changed) {
              markDirty(
                Number.isFinite(earliestChangedBlock)
                  ? earliestChangedBlock
                  : 0,
              );
              schedule();
            }
          });

    const observeBlocks = (blocks: MeasuredDocumentLayoutBlock[]) => {
      observer?.disconnect();
      observedElements = Array.from(
        new Set(
          blocks
            .filter((block) => block.observeResize)
            .map((block) => block.element),
        ),
      );
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
      const measurementStart = dirtyMeasurementFrom.current;
      dirtyMeasurementFrom.current = Number.POSITIVE_INFINITY;
      const textLayoutCollection = collectDocumentTextLayoutParagraphs(
        editor,
        layoutFonts,
        loadedLayoutFontIds,
        measurementCache.current,
        measurementStart,
      );
      editorDom.dataset.paginationTextCandidates = String(
        textLayoutCollection.paragraphs.length,
      );
      editorDom.dataset.paginationTextRuns = String(
        textLayoutCollection.paragraphs.reduce(
          (count, paragraph) => count + paragraph.runs.length,
          0,
        ),
      );
      const textLayouts = new Map<
        string,
        OfficeKernelTextLayoutParagraphResult
      >();
      let fallbackGlyphCount = 0;
      if (textLayoutCollection.paragraphs.length) {
        editorDom.dataset.paginationState = 'shaping';
        try {
          const textLayout = await client.textLayout(
            {
              revision: nextRevision,
              documentRevision,
              paragraphs: textLayoutCollection.paragraphs,
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
          for (const layout of textLayout.layouts) {
            if (layout.missingGlyphCount === 0) {
              textLayouts.set(layout.id, layout);
              fallbackGlyphCount += layout.fallbackGlyphCount;
            }
          }
          editorDom.dataset.paginationTextEngine = textLayout.engine;
          editorDom.dataset.paginationUnsupportedText = String(
            textLayout.unsupportedParagraphIds.length +
              textLayout.layouts.filter(
                (layout) => layout.missingGlyphCount > 0,
              ).length,
          );
        } catch (error) {
          if (
            controller.signal.aborted ||
            (error instanceof DOMException && error.name === 'AbortError')
          ) {
            return;
          }
          editorDom.dataset.paginationTextEngine = 'dom';
          editorDom.dataset.paginationUnsupportedText = String(
            textLayoutCollection.paragraphs.length,
          );
        }
      } else {
        editorDom.dataset.paginationTextEngine = 'dom';
        editorDom.dataset.paginationUnsupportedText = '0';
      }
      editorDom.dataset.paginationShapedParagraphs = String(textLayouts.size);
      editorDom.dataset.paginationFallbackGlyphs = String(fallbackGlyphCount);
      editorDom.dataset.paginationShapedRuns = String(
        textLayoutCollection.paragraphs.reduce(
          (count, paragraph) =>
            count + (textLayouts.has(paragraph.id) ? paragraph.runs.length : 0),
          0,
        ),
      );
      const effectivePage = {
        ...page,
        ...documentPageChromeHeights(editor),
      };
      const availablePageHeight = Math.max(
        1,
        effectivePage.height -
          effectivePage.marginTop -
          effectivePage.marginBottom -
          effectivePage.headerHeight -
          effectivePage.footerHeight,
      );
      const snapshot = measureDocumentLayoutBlocks(
        editor,
        measurementCache.current,
        measurementStart,
        textLayouts,
        availablePageHeight,
      );
      measurementCache.current = snapshot;
      editorDom.dataset.paginationBlocks = String(snapshot.blocks.length);
      editorDom.dataset.paginationFlows = String(
        new Set(
          snapshot.blocks.flatMap((block) =>
            block.block.flowId ? [block.block.flowId] : [],
          ),
        ).size,
      );
      editorDom.dataset.paginationMeasuredBlocks = String(
        snapshot.measuredBlockCount,
      );
      editorDom.dataset.paginationReusedBlocks = String(
        snapshot.reusedBlockCount,
      );
      observeBlocks(snapshot.blocks);
      if (snapshot.unsupportedLayout || !snapshot.blocks.length) {
        editorDom.dataset.paginationState = snapshot.unsupportedLayout
          ? 'unsupported'
          : 'empty';
        setPagination(null);
        paginationCache.current = null;
        return;
      }

      const previousPagination = paginationCache.current;
      const layoutPlan = planIncrementalDocumentLayout(
        previousPagination,
        snapshot.blocks,
        measurementStart,
      );
      editorDom.dataset.paginationLaidOutBlocks = String(
        layoutPlan.blocks.length,
      );
      editorDom.dataset.paginationReusedPages = String(
        layoutPlan.reusedPageCount,
      );
      try {
        editorDom.dataset.paginationState = 'layout';
        const partialLayout = await client.layout(
          {
            revision: nextRevision,
            documentRevision,
            startPageIndex: layoutPlan.startPageIndex,
            page: effectivePage,
            blocks: layoutPlan.blocks,
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
        const layout =
          previousPagination && layoutPlan.startPageIndex > 0
            ? mergeIncrementalDocumentLayout(
                previousPagination.layout,
                partialLayout,
                effectivePage,
              )
            : partialLayout;
        const blockById = new Map(
          snapshot.blocks.map((block) => [block.block.id, block] as const),
        );
        const pages = documentPaginationPageDescriptors(
          layout,
          snapshot.blocks,
        );
        const pageByIndex = new Map(
          pages.map(
            (descriptor) => [descriptor.pageIndex, descriptor] as const,
          ),
        );
        applyDocumentPagination(
          editor,
          nextRevision,
          layout.breaks.flatMap((pageBreak) => {
            const block = blockById.get(pageBreak.beforeBlockId);
            return block
              ? [
                  {
                    ...pageBreak,
                    page: effectivePage,
                    position: block.from,
                    inlineOffsetLeft: block.inlineOffsetLeft,
                    inlineOffsetRight: block.inlineOffsetRight,
                    previousPageChrome: visualPageChrome(
                      pageByIndex.get(pageBreak.pageIndex - 1),
                    ),
                    nextPageChrome: visualPageChrome(
                      pageByIndex.get(pageBreak.pageIndex),
                    ),
                    tableBreak: block.tableBreak,
                  },
                ]
              : [];
          }),
        );
        editorDom.dataset.paginationEngine = layout.engine;
        editorDom.dataset.paginationDocumentRevision = String(
          layout.documentRevision,
        );
        editorDom.dataset.paginationPages = String(layout.pages.length);
        editorDom.dataset.paginationState = 'ready';
        delete editorDom.dataset.paginationError;
        const nextPagination = {
          layout,
          blocks: snapshot.blocks,
          pages,
          pageByBlockId: new Map(
            layout.pages.flatMap((page) =>
              page.placements.map(
                (placement) => [placement.blockId, page.index + 1] as const,
              ),
            ),
          ),
        };
        paginationCache.current = nextPagination;
        setPagination(nextPagination);
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
        paginationCache.current = null;
        setPagination(null);
      }
    };

    const schedule = () => {
      if (disposed || frame) return;
      frame = requestAnimationFrame(() => {
        frame = requestAnimationFrame(() => void run());
      });
    };
    const markDirty = (position: number) => {
      dirtyMeasurementFrom.current = Math.min(
        dirtyMeasurementFrom.current,
        Math.max(0, position),
      );
    };
    const handleDocumentUpdate = ({ transaction }: EditorEvents['update']) => {
      markDirty(earliestChangedPosition(transaction));
      schedule();
    };
    const handleLoadedAsset = (event: Event) => {
      const target = event.target;
      const block = measurementCache.current?.blocks.find(
        (candidate) =>
          target instanceof Node &&
          (candidate.element === target || candidate.element.contains(target)),
      );
      markDirty(block?.from ?? 0);
      schedule();
    };
    const handleFontLoading = () => {
      markDirty(0);
      schedule();
    };
    const handleWindowResize = () => {
      markDirty(0);
      schedule();
    };
    const fonts = document.fonts;
    editor.on('update', handleDocumentUpdate);
    editorDom.addEventListener('load', handleLoadedAsset, true);
    fonts?.addEventListener('loadingdone', handleFontLoading);
    window.addEventListener('resize', handleWindowResize);
    schedule();

    return () => {
      disposed = true;
      if (frame) cancelAnimationFrame(frame);
      observer?.disconnect();
      editor.off('update', handleDocumentUpdate);
      editorDom.removeEventListener('load', handleLoadedAsset, true);
      fonts?.removeEventListener('loadingdone', handleFontLoading);
      window.removeEventListener('resize', handleWindowResize);
      activeRequest.current?.abort();
      activeRequest.current = null;
    };
  }, [
    client,
    documentRevision,
    editor,
    editorMounted,
    enabled,
    layoutKey,
    layoutFontKey,
    layoutFonts,
    loadedLayoutFontKey,
    loadedLayoutFontIds,
    pageKey,
  ]);

  return useMemo(() => {
    const currentPage = pagination
      ? pageForPosition(pagination, editor?.state.selection.from ?? 0)
      : null;
    return {
      currentPage,
      currentPageDescriptor:
        currentPage === null
          ? null
          : (pagination?.pages.find(
              (page) => page.physicalPage === currentPage,
            ) ?? null),
      pageCount: pagination
        ? Math.max(1, pagination.layout.pages.length)
        : null,
      pages: pagination?.pages ?? [],
      paginating: enabled && pagination === null,
    };
  }, [editor, enabled, pagination, selectionVersion]);
}

export function documentPaginationPageDescriptors(
  layout: OfficeKernelLayoutResult,
  blocks: readonly MeasuredDocumentLayoutBlock[],
): DocumentPaginationPageDescriptor[] {
  const blockById = new Map(
    blocks.map((candidate) => [candidate.block.id, candidate] as const),
  );
  const sectionPages = new Map<string, number>();
  const descriptors: DocumentPaginationPageDescriptor[] = [];

  for (const page of layout.pages) {
    const pageSections = Array.from(
      new Map(
        page.placements.flatMap((placement) => {
          const section = blockById.get(placement.blockId)?.section;
          return section ? [[section.id, section] as const] : [];
        }),
      ).values(),
    );
    for (const section of pageSections) {
      sectionPages.set(section.id, (sectionPages.get(section.id) ?? 0) + 1);
    }
    const section = pageSections[0];
    if (!section) continue;

    const physicalPage = page.index + 1;
    const sectionPage = sectionPages.get(section.id) ?? 1;
    const previousPageNumber = descriptors.at(-1)?.pageNumber ?? 0;
    const firstSectionPage = sectionPage === 1;
    const pageNumber =
      firstSectionPage && section.layout.pageNumberStart !== undefined
        ? section.layout.pageNumberStart
        : Math.max(1, previousPageNumber + 1);
    descriptors.push({
      pageIndex: page.index,
      physicalPage,
      pageNumber,
      sectionPage,
      sectionId: section.id,
      sectionIndex: section.index,
      layout: section.layout,
      pageChrome: resolveDocumentPageChrome(
        section.layout,
        sectionPage,
        physicalPage,
      ),
    });
  }
  return descriptors;
}

function visualPageChrome(
  page: DocumentPaginationPageDescriptor | undefined,
): DocumentPaginationVisualPageChrome | undefined {
  if (!page) return undefined;
  return {
    ...page.pageChrome,
    pageNumber: page.pageNumber,
  };
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
  fonts: readonly WorkDocumentLayoutFont[],
): OfficeKernelClient | null {
  const [client, setClient] = useState<OfficeKernelClient | null>(null);
  const fontKey = documentLayoutFontKey(fonts);
  const fontSources = useMemo(
    () => fonts.map((font) => ({ id: font.id, url: font.url })),
    [fontKey],
  );
  useEffect(() => {
    const next = createOfficeKernelClient(wasmUrl, fontSources);
    setClient(next);
    return () => next.dispose();
  }, [fontSources, wasmUrl]);
  return client;
}

export function pageForPosition(
  pagination: DocumentPaginationResult,
  position: number,
): number {
  const rangedBlock = pagination.blocks.find((block) =>
    block.selectionRanges?.some(
      (range) => range.from <= position && position < range.to,
    ),
  );
  if (rangedBlock) {
    return pagination.pageByBlockId.get(rangedBlock.block.id) ?? 1;
  }

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

function earliestChangedPosition(transaction: Transaction): number {
  let earliest = Number.POSITIVE_INFINITY;
  for (const step of transaction.steps) {
    step.getMap().forEach((_oldStart, _oldEnd, newStart) => {
      earliest = Math.min(earliest, newStart);
    });
  }
  return Number.isFinite(earliest) ? earliest : 0;
}
