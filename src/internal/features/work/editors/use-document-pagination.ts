import type { Editor, EditorEvents } from '@tiptap/core';
import type { Node as ProseMirrorNode } from '@tiptap/pm/model';
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
import type {
  WorkDocumentFieldContext,
  WorkDocumentFieldContextResolver,
} from '../work-document-fields';
import {
  documentLayoutFontKey,
  type WorkDocumentLayoutFont,
} from '../work-document-fonts';
import {
  mergeIncrementalDocumentLayout,
  planIncrementalDocumentLayout,
} from '../work-document-incremental-layout';
import {
  type ResolvedDocumentPageChrome,
  resolveDocumentPageChrome,
} from '../work-document-page-chrome';
import {
  collectDocumentTextLayoutParagraphs,
  type DocumentPaginationSnapshot,
  type DocumentPaginationVisualPageChrome,
  documentPageMetrics,
  documentTextLayoutBatches,
  type MeasuredDocumentLayoutBlock,
  measureDocumentLayoutBlocksIncrementally,
} from '../work-document-pagination';
import type { WorkDocumentSectionLayout } from '../work-types';
import { createDocumentPaginationMeasurementRange } from './document-pagination-measurement-range';
import { createDocumentPaginationRunCoordinator } from './document-pagination-run-coordinator';

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
  previewText: string;
  selectionPosition: number;
  sectionPage: number;
  sectionId: string;
  sectionIndex: number;
  layout: WorkDocumentSectionLayout;
  pageChrome: ResolvedDocumentPageChrome;
  page: OfficeKernelPageMetrics;
}

export interface DocumentPaginationPageDescriptorDerivation {
  pages: DocumentPaginationPageDescriptor[];
  reusedPageCount: number;
  derivedPageCount: number;
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
  resolveFieldContext: WorkDocumentFieldContextResolver | null;
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
  const measurementCache = useRef<DocumentPaginationSnapshot | null>(null);
  const paginationCache = useRef<DocumentPaginationResult | null>(null);
  const measurementEditor = useRef<Editor | null>(null);
  const measurementLayoutKey = useRef('');
  const measurementRangeRef = useRef<ReturnType<
    typeof createDocumentPaginationMeasurementRange
  > | null>(null);
  measurementRangeRef.current ??= createDocumentPaginationMeasurementRange();
  const measurementRange = measurementRangeRef.current;
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
      measurementRange.reset();
    }
    if (observedDocumentRevision.current !== documentRevision) {
      measurementRange.ensureDirty();
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
      if (editor && editorMounted && !editor.isDestroyed) {
        const editorDom = editor.view.dom;
        editor.commands.clearDocumentPagination(nextRevision);
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
          delete editorDom.dataset.paginationReusedPageChrome;
          delete editorDom.dataset.paginationReusedPages;
          delete editorDom.dataset.paginationDerivedPageChrome;
          delete editorDom.dataset.paginationShapedParagraphs;
          delete editorDom.dataset.paginationShapedRuns;
          delete editorDom.dataset.paginationTextCandidates;
          delete editorDom.dataset.paginationTextEngine;
          delete editorDom.dataset.paginationTextRuns;
          delete editorDom.dataset.paginationUnsupportedText;
          delete editorDom.dataset.paginationAborts;
          delete editorDom.dataset.paginationActive;
          delete editorDom.dataset.paginationAssetTriggers;
          delete editorDom.dataset.paginationCoalescedRequests;
          delete editorDom.dataset.paginationDocumentTriggers;
          delete editorDom.dataset.paginationFontTriggers;
          delete editorDom.dataset.paginationResizeTriggers;
          delete editorDom.dataset.paginationRuns;
          delete editorDom.dataset.paginationTextBatch;
          delete editorDom.dataset.paginationTextBatches;
          delete editorDom.dataset.paginationViewportTriggers;
        }
      }
      if (!enabled) {
        measurementCache.current = null;
        paginationCache.current = null;
        measurementRange.reset();
      }
      setPagination(null);
      return;
    }

    const editorDom = editor.view.dom;
    let disposed = false;
    let observedElements: HTMLElement[] = [];
    const observedHeights = new Map<HTMLElement, number>();
    let observedBlockPositions: ReadonlyMap<HTMLElement, number> = new Map();
    let observingExternalLayout = false;
    const diagnosticCounts = {
      aborts: 0,
      assetTriggers: 0,
      coalescedRequests: 0,
      documentTriggers: 0,
      fontTriggers: 0,
      resizeTriggers: 0,
      runs: 0,
      viewportTriggers: 0,
    };
    const updateDiagnostic = (
      name: keyof typeof diagnosticCounts,
      datasetName: keyof DOMStringMap,
    ) => {
      diagnosticCounts[name] += 1;
      editorDom.dataset[datasetName] = String(diagnosticCounts[name]);
    };
    const observer =
      typeof ResizeObserver === 'undefined'
        ? null
        : new ResizeObserver((entries) => {
            if (!observingExternalLayout) return;
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
                const blockPosition = observedBlockPositions.get(element);
                if (blockPosition !== undefined) {
                  earliestChangedBlock = Math.min(
                    earliestChangedBlock,
                    blockPosition,
                  );
                }
              }
            }
            if (changed) {
              updateDiagnostic('resizeTriggers', 'paginationResizeTriggers');
              markDirty(
                Number.isFinite(earliestChangedBlock)
                  ? earliestChangedBlock
                  : 0,
              );
              schedule();
            }
          });

    const stopObservingBlocks = () => {
      observingExternalLayout = false;
      observer?.disconnect();
    };

    const observeBlocks = (blocks: MeasuredDocumentLayoutBlock[]) => {
      stopObservingBlocks();
      observedBlockPositions = documentResizeObservationPositions(blocks);
      observedElements = Array.from(observedBlockPositions.keys());
      observedHeights.clear();
      for (const element of observedElements) {
        observedHeights.set(element, element.offsetHeight);
      }
      observingExternalLayout = true;
      for (const element of observedElements) observer?.observe(element);
    };

    const run = async (signal: AbortSignal) => {
      stopObservingBlocks();
      const nextRevision = ++revision.current;
      editor.commands.clearDocumentPagination(nextRevision);
      editorDom.dataset.paginationState = 'measuring';
      editorDom.dataset.paginationReusedPageChrome = '0';
      editorDom.dataset.paginationDerivedPageChrome = '0';
      delete editorDom.dataset.paginationError;
      const measurementPass = measurementRange.begin();
      const measurementStart = measurementPass.from;
      try {
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
            let textLayoutEngine = '';
            let unsupportedTextLayoutCount = 0;
            const textLayoutBatches = documentTextLayoutBatches(
              textLayoutCollection.paragraphs,
            );
            editorDom.dataset.paginationTextBatches = String(
              textLayoutBatches.length,
            );
            for (const [
              batchIndex,
              paragraphs,
            ] of textLayoutBatches.entries()) {
              editorDom.dataset.paginationTextBatch = String(batchIndex + 1);
              const textLayout = await client.textLayout(
                {
                  revision: nextRevision,
                  documentRevision,
                  paragraphs,
                },
                signal,
              );
              if (
                disposed ||
                signal.aborted ||
                nextRevision !== revision.current ||
                editor.isDestroyed
              ) {
                return;
              }
              textLayoutEngine ||= textLayout.engine;
              unsupportedTextLayoutCount +=
                textLayout.unsupportedParagraphIds.length +
                textLayout.layouts.filter(
                  (layout) => layout.missingGlyphCount > 0,
                ).length;
              for (const layout of textLayout.layouts) {
                if (layout.missingGlyphCount === 0) {
                  textLayouts.set(layout.id, layout);
                  fallbackGlyphCount += layout.fallbackGlyphCount;
                }
              }
            }
            editorDom.dataset.paginationTextEngine = textLayoutEngine;
            editorDom.dataset.paginationUnsupportedText = String(
              unsupportedTextLayoutCount,
            );
          } catch (error) {
            if (
              signal.aborted ||
              (error instanceof DOMException && error.name === 'AbortError')
            ) {
              return;
            }
            textLayouts.clear();
            fallbackGlyphCount = 0;
            editorDom.dataset.paginationTextEngine = 'dom';
            editorDom.dataset.paginationUnsupportedText = String(
              textLayoutCollection.paragraphs.length,
            );
          }
        } else {
          editorDom.dataset.paginationTextEngine = 'dom';
          editorDom.dataset.paginationUnsupportedText = '0';
          editorDom.dataset.paginationTextBatch = '0';
          editorDom.dataset.paginationTextBatches = '0';
        }
        editorDom.dataset.paginationShapedParagraphs = String(textLayouts.size);
        editorDom.dataset.paginationFallbackGlyphs = String(fallbackGlyphCount);
        editorDom.dataset.paginationShapedRuns = String(
          textLayoutCollection.paragraphs.reduce(
            (count, paragraph) =>
              count +
              (textLayouts.has(paragraph.id) ? paragraph.runs.length : 0),
            0,
          ),
        );
        const snapshot = await measureDocumentLayoutBlocksIncrementally(
          editor,
          measurementCache.current,
          measurementStart,
          textLayouts,
          1_000_000,
          { signal },
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
        if (snapshot.unsupportedLayout || !snapshot.blocks.length) {
          observeBlocks(snapshot.blocks);
          editorDom.dataset.paginationState = snapshot.unsupportedLayout
            ? 'unsupported'
            : 'empty';
          setPagination(null);
          paginationCache.current = null;
          measurementRange.commit(measurementPass);
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
              page,
              pageStyles: snapshot.pageStyles,
              blocks: layoutPlan.blocks,
            },
            signal,
          );
          if (
            disposed ||
            signal.aborted ||
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
                )
              : partialLayout;
          const blockById = new Map(
            snapshot.blocks.map((block) => [block.block.id, block] as const),
          );
          const pageDescriptorDerivation =
            deriveDocumentPaginationPageDescriptors(
              layout,
              snapshot.blocks,
              editor.state.doc,
              previousPagination?.pages,
              layoutPlan.reusedPageCount,
            );
          const pages = pageDescriptorDerivation.pages;
          editorDom.dataset.paginationReusedPageChrome = String(
            pageDescriptorDerivation.reusedPageCount,
          );
          editorDom.dataset.paginationDerivedPageChrome = String(
            pageDescriptorDerivation.derivedPageCount,
          );
          const pageByIndex = new Map(
            pages.map(
              (descriptor) => [descriptor.pageIndex, descriptor] as const,
            ),
          );
          editor.commands.applyDocumentPagination(
            nextRevision,
            layout.breaks.flatMap((pageBreak) => {
              const block = blockById.get(pageBreak.beforeBlockId);
              return block
                ? [
                    {
                      ...pageBreak,
                      previousPage:
                        pageByIndex.get(pageBreak.pageIndex - 1)?.page ??
                        layout.pages[pageBreak.pageIndex - 1]?.page ??
                        page,
                      nextPage:
                        pageByIndex.get(pageBreak.pageIndex)?.page ??
                        layout.pages[pageBreak.pageIndex]?.page ??
                        page,
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
          observeBlocks(snapshot.blocks);
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
          measurementRange.commit(measurementPass);
          setPagination(nextPagination);
        } catch (error) {
          if (
            disposed ||
            signal.aborted ||
            (error instanceof DOMException && error.name === 'AbortError')
          ) {
            return;
          }
          editorDom.dataset.paginationState = 'error';
          editorDom.dataset.paginationError =
            error instanceof Error
              ? `${error.name}: ${error.message}`
              : 'UnknownError';
          editor.commands.clearDocumentPagination(nextRevision);
          paginationCache.current = null;
          setPagination(null);
        }
      } finally {
        measurementRange.restore(measurementPass);
      }
    };

    const coordinator = createDocumentPaginationRunCoordinator({
      cancelFrame: cancelAnimationFrame,
      onAbort: () => {
        measurementRange.restoreActive();
        updateDiagnostic('aborts', 'paginationAborts');
      },
      onCoalescedRequest: () =>
        updateDiagnostic('coalescedRequests', 'paginationCoalescedRequests'),
      onError: (error) => {
        measurementRange.restoreActive();
        if (disposed || editor.isDestroyed) return;
        editorDom.dataset.paginationState = 'error';
        editorDom.dataset.paginationError =
          error instanceof Error
            ? `${error.name}: ${error.message}`
            : 'UnknownError';
        paginationCache.current = null;
        setPagination(null);
      },
      onRunFinish: () => {
        if (!disposed) editorDom.dataset.paginationActive = 'false';
      },
      onRunStart: () => {
        editorDom.dataset.paginationActive = 'true';
        updateDiagnostic('runs', 'paginationRuns');
      },
      requestFrame: requestAnimationFrame,
      run,
    });
    const schedule = (invalidateActive = false) => {
      coordinator.request({ invalidateActive });
    };
    const markDirty = (position: number) => {
      measurementRange.invalidate(position);
    };
    const handleDocumentUpdate = ({ transaction }: EditorEvents['update']) => {
      updateDiagnostic('documentTriggers', 'paginationDocumentTriggers');
      markDirty(earliestChangedPosition(transaction));
      schedule(true);
    };
    const handleLoadedAsset = (event: Event) => {
      updateDiagnostic('assetTriggers', 'paginationAssetTriggers');
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
      updateDiagnostic('fontTriggers', 'paginationFontTriggers');
      markDirty(0);
      schedule();
    };
    const handleWindowResize = () => {
      updateDiagnostic('viewportTriggers', 'paginationViewportTriggers');
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
      stopObservingBlocks();
      editor.off('update', handleDocumentUpdate);
      editorDom.removeEventListener('load', handleLoadedAsset, true);
      fonts?.removeEventListener('loadingdone', handleFontLoading);
      window.removeEventListener('resize', handleWindowResize);
      coordinator.dispose();
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

  const resolveFieldContext = useMemo(
    () =>
      pagination
        ? createDocumentFieldPaginationContextResolver(pagination)
        : null,
    [pagination],
  );

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
      resolveFieldContext,
    };
  }, [editor, enabled, pagination, resolveFieldContext, selectionVersion]);
}

export function createDocumentFieldPaginationContextResolver(
  pagination: DocumentPaginationResult,
): WorkDocumentFieldContextResolver {
  const pagesByPhysicalNumber = new Map(
    pagination.pages.map((page) => [page.physicalPage, page] as const),
  );
  const sectionPhysicalPages = new Map<string, Set<number>>();
  for (const block of pagination.blocks) {
    const sectionId = block.section?.id;
    const physicalPage = pagination.pageByBlockId.get(block.block.id);
    if (!sectionId || !physicalPage) continue;
    const pages = sectionPhysicalPages.get(sectionId) ?? new Set<number>();
    pages.add(physicalPage);
    sectionPhysicalPages.set(sectionId, pages);
  }
  return (position): WorkDocumentFieldContext | null => {
    const block = blockForPosition(pagination.blocks, position);
    const section = block?.section;
    const physicalPage = block
      ? pagination.pageByBlockId.get(block.block.id)
      : undefined;
    const page = physicalPage
      ? pagesByPhysicalNumber.get(physicalPage)
      : undefined;
    if (!section || !page) return null;
    return {
      pageNumber: page.pageNumber,
      totalPages: Math.max(1, pagination.layout.pages.length),
      sectionNumber: section.index + 1,
      sectionPages: Math.max(
        1,
        sectionPhysicalPages.get(section.id)?.size ?? 0,
      ),
    };
  };
}

export function documentResizeObservationPositions(
  blocks: readonly MeasuredDocumentLayoutBlock[],
): ReadonlyMap<HTMLElement, number> {
  const positions = new Map<HTMLElement, number>();
  for (const block of blocks) {
    if (!block.observeResize) continue;
    const previous = positions.get(block.element);
    positions.set(
      block.element,
      previous === undefined ? block.from : Math.min(previous, block.from),
    );
  }
  return positions;
}

export function documentPaginationPageDescriptors(
  layout: OfficeKernelLayoutResult,
  blocks: readonly MeasuredDocumentLayoutBlock[],
  documentNode?: ProseMirrorNode,
): DocumentPaginationPageDescriptor[] {
  return deriveDocumentPaginationPageDescriptors(layout, blocks, documentNode)
    .pages;
}

export function deriveDocumentPaginationPageDescriptors(
  layout: OfficeKernelLayoutResult,
  blocks: readonly MeasuredDocumentLayoutBlock[],
  documentNode?: ProseMirrorNode,
  previousPages: readonly DocumentPaginationPageDescriptor[] = [],
  requestedReusedPageCount = 0,
): DocumentPaginationPageDescriptorDerivation {
  const blockById = new Map(
    blocks.map((candidate) => [candidate.block.id, candidate] as const),
  );
  const previousPageByIndex = new Map(
    previousPages.map((page) => [page.pageIndex, page] as const),
  );
  const reusablePhysicalPageCount = Number.isSafeInteger(
    requestedReusedPageCount,
  )
    ? Math.max(0, Math.min(requestedReusedPageCount, layout.pages.length))
    : 0;
  const sectionPages = new Map<string, number>();
  const descriptors: DocumentPaginationPageDescriptor[] = [];
  let reuseOpen = reusablePhysicalPageCount > 0;
  let reusedPageCount = 0;

  for (const page of layout.pages) {
    const pageBlocks = Array.from(
      new Map(
        page.placements.flatMap((placement) => {
          const block = blockById.get(placement.blockId);
          return block ? [[block.block.id, block] as const] : [];
        }),
      ).values(),
    );
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
    const previousPage = previousPageByIndex.get(page.index);
    if (
      reuseOpen &&
      page.index < reusablePhysicalPageCount &&
      previousPage &&
      reusableDocumentPageDescriptor(
        previousPage,
        physicalPage,
        pageNumber,
        sectionPage,
        section.id,
        section.index,
        section.layout,
        page.page,
      )
    ) {
      descriptors.push(previousPage);
      reusedPageCount += 1;
      continue;
    }
    if (page.index < reusablePhysicalPageCount) reuseOpen = false;
    descriptors.push({
      pageIndex: page.index,
      physicalPage,
      pageNumber,
      previewText: documentPagePreviewText(pageBlocks, documentNode),
      selectionPosition: Math.max(
        1,
        Math.min(
          ...pageBlocks.map(
            (block) => block.selectionRanges?.[0]?.from ?? block.from,
          ),
        ),
      ),
      sectionPage,
      sectionId: section.id,
      sectionIndex: section.index,
      layout: section.layout,
      pageChrome: resolveDocumentPageChrome(
        section.layout,
        sectionPage,
        physicalPage,
      ),
      page: documentPageMetrics(section.layout, physicalPage),
    });
  }
  return {
    pages: descriptors,
    reusedPageCount,
    derivedPageCount: descriptors.length - reusedPageCount,
  };
}

function reusableDocumentPageDescriptor(
  candidate: DocumentPaginationPageDescriptor,
  physicalPage: number,
  pageNumber: number,
  sectionPage: number,
  sectionId: string,
  sectionIndex: number,
  layout: WorkDocumentSectionLayout,
  page: OfficeKernelPageMetrics,
): boolean {
  return (
    candidate.pageIndex === physicalPage - 1 &&
    candidate.physicalPage === physicalPage &&
    candidate.pageNumber === pageNumber &&
    candidate.sectionPage === sectionPage &&
    candidate.sectionId === sectionId &&
    candidate.sectionIndex === sectionIndex &&
    pageMetricsKey(candidate.page) === pageMetricsKey(page) &&
    JSON.stringify(candidate.layout) === JSON.stringify(layout)
  );
}

function documentPagePreviewText(
  blocks: readonly MeasuredDocumentLayoutBlock[],
  documentNode: ProseMirrorNode | undefined,
): string {
  const measuredText = documentNode
    ? documentPageTextRanges(blocks)
        .reduce(
          (preview, range) => {
            const from = Math.max(preview.cursor, range.from);
            const to = Math.min(documentNode.content.size, range.to);
            if (to <= from) return preview;
            const separator = preview.text && from > preview.cursor ? ' ' : '';
            return {
              cursor: to,
              text: `${preview.text}${separator}${documentNode.textBetween(
                from,
                to,
                ' ',
                ' ',
              )}`,
            };
          },
          { cursor: 0, text: '' },
        )
        .text.replaceAll(/\s+/g, ' ')
        .trim()
    : '';
  if (measuredText) return measuredText.slice(0, 320);
  return blocks
    .map((block) => block.element.textContent?.replaceAll(/\s+/g, ' ').trim())
    .filter((text): text is string => Boolean(text))
    .join(' ')
    .slice(0, 320);
}

function documentPageTextRanges(
  blocks: readonly MeasuredDocumentLayoutBlock[],
): Array<{ from: number; to: number }> {
  return Array.from(
    new Map(
      blocks
        .flatMap((block) =>
          block.selectionRanges?.length
            ? block.selectionRanges
            : [{ from: block.from, to: block.to }],
        )
        .filter(
          (range) =>
            Number.isSafeInteger(range.from) &&
            Number.isSafeInteger(range.to) &&
            range.to > range.from,
        )
        .map((range) => [`${range.from}:${range.to}`, range] as const),
    ).values(),
  ).sort((left, right) => left.from - right.from || left.to - right.to);
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
  const block = blockForPosition(pagination.blocks, position);
  return block ? (pagination.pageByBlockId.get(block.block.id) ?? 1) : 1;
}

function blockForPosition(
  blocks: readonly MeasuredDocumentLayoutBlock[],
  position: number,
): MeasuredDocumentLayoutBlock | undefined {
  const rangedBlock = blocks.find((block) =>
    block.selectionRanges?.some(
      (range) => range.from <= position && position < range.to,
    ),
  );
  if (rangedBlock) return rangedBlock;

  let lower = 0;
  let upper = blocks.length - 1;
  let containing = blocks[0];
  while (lower <= upper) {
    const middle = Math.floor((lower + upper) / 2);
    const candidate = blocks[middle];
    if (candidate.from <= position) {
      containing = candidate;
      lower = middle + 1;
    } else {
      upper = middle - 1;
    }
  }
  return containing;
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
  ]
    .map(canonicalPageMetric)
    .join(':');
}

function canonicalPageMetric(value: number): number {
  return Math.round(value * 1_000_000) / 1_000_000;
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
