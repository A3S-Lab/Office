import { Editor } from '@tiptap/core';
import { afterEach, describe, expect, test } from '@rstest/core';
import { createWorkDocumentExtensions } from '../src/internal/features/work/work-document-extensions';
import { documentChunkViewportRange } from '../src/internal/features/work/work-document-chunk-node';
import {
  DocumentPagination,
  measureDocumentLayoutBlocks,
} from '../src/internal/features/work/work-document-pagination';
import { windowDocumentModel } from '../src/internal/features/work/work-document-windowing';
import type { WorkDocumentNode } from '../src/internal/features/work/work-types';

const originalIntersectionObserver = globalThis.IntersectionObserver;

afterEach(() => {
  Object.defineProperty(globalThis, 'IntersectionObserver', {
    configurable: true,
    value: originalIntersectionObserver,
    writable: true,
  });
});

describe('document chunk viewport', () => {
  test('mounts the exact geometry window without adjacent overscan', () => {
    const geometry = [
      { start: 0, end: 100 },
      { start: 100, end: 200 },
      { start: 200, end: 300 },
      { start: 300, end: 400 },
    ];

    expect(documentChunkViewportRange(geometry, 120, 180)).toEqual({
      start: 1,
      end: 2,
    });
    expect(documentChunkViewportRange(geometry, 120, 240)).toEqual({
      start: 1,
      end: 3,
    });
  });

  test('mounts a bounded initial window and materializes a selected chunk', () => {
    installPassiveIntersectionObserver();
    const root = windowDocumentModel(documentRoot(8), {
      blockSize: 2,
      blockThreshold: 4,
      tableRowSize: 2,
      tableRowThreshold: 4,
    });
    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: root,
    });

    expect(editor.view.dom.dataset.documentWindowed).toBe('true');
    expect(editor.view.dom.dataset.documentChunkCount).toBe('4');
    expect(
      editor.view.dom.querySelectorAll('[data-document-chunk-mounted="true"]'),
    ).toHaveLength(2);
    expect(editor.view.dom.querySelectorAll('p')).toHaveLength(4);

    const lastChunk = chunkPositions(editor.state.doc).at(-1);
    if (!lastChunk) throw new Error('Expected a final document chunk.');
    editor.commands.setTextSelection(lastChunk.position + 2);

    expect(
      editor.view.dom.querySelector(
        `[data-document-chunk-id="${lastChunk.id}"]`,
      ),
    ).toHaveAttribute('data-document-chunk-mounted', 'true');
    expect(editor.view.dom.dataset.documentMountedChunkCount).toBe('3');
    expect(editor.view.dom.querySelectorAll('p')).toHaveLength(6);
    editor.destroy();
  });

  test('navigates to model boundaries without relying on mounted DOM', () => {
    installPassiveIntersectionObserver();
    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: windowDocumentModel(documentRoot(8), {
        blockSize: 2,
        blockThreshold: 4,
      }),
    });
    const chunks = chunkPositions(editor.state.doc);
    const first = chunks[0];
    const last = chunks.at(-1);
    if (!first || !last) throw new Error('Expected document chunks.');

    editor.view.dom.dispatchEvent(
      new KeyboardEvent('keydown', {
        bubbles: true,
        ctrlKey: true,
        key: 'End',
      }),
    );

    expect(editor.state.selection.anchor).toBeGreaterThan(last.position);
    expect(
      editor.view.dom.querySelector(`[data-document-chunk-id="${last.id}"]`),
    ).toHaveAttribute('data-document-chunk-mounted', 'true');

    editor.view.dom.dispatchEvent(
      new KeyboardEvent('keydown', {
        bubbles: true,
        ctrlKey: true,
        key: 'Home',
      }),
    );

    expect(editor.state.selection.anchor).toBeGreaterThan(first.position);
    expect(editor.state.selection.anchor).toBeLessThan(last.position);
    editor.destroy();
  });

  test('falls back to the complete document when IntersectionObserver is absent', () => {
    Object.defineProperty(globalThis, 'IntersectionObserver', {
      configurable: true,
      value: undefined,
      writable: true,
    });
    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: windowDocumentModel(documentRoot(6), {
        blockSize: 2,
        blockThreshold: 4,
      }),
    });

    expect(editor.view.dom.querySelectorAll('p')).toHaveLength(6);
    expect(editor.view.dom.dataset.documentWindowingFallback).toBe('true');
    editor.destroy();
  });

  test('retains virtual table continuity metadata in the TipTap state', () => {
    installPassiveIntersectionObserver();
    const root: WorkDocumentNode = {
      type: 'doc',
      content: [
        {
          type: 'documentSection',
          attrs: { id: 'document-section-1' },
          content: [tableNode(5)],
        },
      ],
    };
    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: windowDocumentModel(root, {
        blockThreshold: 100,
        tableRowSize: 2,
        tableRowThreshold: 4,
      }),
    });
    const tableIds: unknown[] = [];
    editor.state.doc.descendants((node) => {
      if (node.type.name === 'table') tableIds.push(node.attrs.virtualTableId);
    });

    expect(tableIds).toEqual([
      'document-table-1-1',
      'document-table-1-1',
      'document-table-1-1',
    ]);
    editor.destroy();
  });

  test('paginates every logical block without mounting every paragraph', () => {
    installPassiveIntersectionObserver();
    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: windowDocumentModel(documentRoot(8), {
        blockSize: 2,
        blockThreshold: 4,
      }),
    });

    const snapshot = measureDocumentLayoutBlocks(editor);
    expect(editor.view.dom.querySelectorAll('p')).toHaveLength(4);
    expect(snapshot.blocks).toHaveLength(8);
    expect(snapshot.blocks.every((block) => !block.observeResize)).toBe(true);
    expect(snapshot.blocks.map((block) => block.from)).toEqual(
      [...snapshot.blocks.map((block) => block.from)].sort(
        (left, right) => left - right,
      ),
    );
    editor.destroy();
  });

  test('paginates unmounted virtual table slices as row flows', () => {
    installPassiveIntersectionObserver();
    const root: WorkDocumentNode = {
      type: 'doc',
      content: [
        {
          type: 'documentSection',
          attrs: { id: 'document-section-1' },
          content: [tableNode(5)],
        },
      ],
    };
    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: windowDocumentModel(root, {
        blockThreshold: 100,
        tableRowSize: 2,
        tableRowThreshold: 4,
      }),
    });

    const snapshot = measureDocumentLayoutBlocks(editor);
    expect(snapshot.blocks).toHaveLength(5);
    expect(snapshot.blocks.every((block) => Boolean(block.block.flowId))).toBe(
      true,
    );
    expect(snapshot.blocks.every((block) => Boolean(block.tableBreak))).toBe(
      true,
    );
    editor.destroy();
  });

  test('paginates nested table windows while mounting only leaf slices', () => {
    installPassiveIntersectionObserver();
    const root: WorkDocumentNode = {
      type: 'doc',
      content: [
        {
          type: 'documentSection',
          attrs: { id: 'document-section-1' },
          content: [tableNode(80)],
        },
      ],
    };
    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: windowDocumentModel(root, {
        blockThreshold: 100,
        tableRowSize: 2,
        tableRowThreshold: 4,
      }),
    });

    const snapshot = measureDocumentLayoutBlocks(editor);
    expect(snapshot.blocks).toHaveLength(80);
    expect(editor.view.dom.dataset.documentChunkCount).toBe('40');
    expect(
      editor.view.dom.querySelectorAll(
        '[data-document-chunk-window-container="true"]',
      ),
    ).toHaveLength(2);
    const containers = editor.view.dom.querySelectorAll<HTMLElement>(
      '[data-document-chunk-window-container="true"]',
    );
    expect(containers[0]).toHaveAttribute(
      'data-document-chunk-mounted',
      'true',
    );
    expect(containers[1]).toHaveAttribute(
      'data-document-chunk-mounted',
      'false',
    );
    expect(
      editor.view.dom.querySelectorAll(
        '[data-document-chunk-window-container="false"]',
      ),
    ).toHaveLength(32);
    expect(editor.view.dom.querySelectorAll('tbody > tr')).toHaveLength(4);
    editor.destroy();
  });

  test('includes automatic page spacers in an unmounted chunk placeholder', () => {
    installPassiveIntersectionObserver();
    const editor = new Editor({
      extensions: [...createWorkDocumentExtensions(), DocumentPagination],
      content: windowDocumentModel(documentRoot(8), {
        blockSize: 2,
        blockThreshold: 4,
      }),
    });
    const chunks = chunkPositions(editor.state.doc);
    const target = chunks[2];
    if (!target) throw new Error('Expected a third document chunk.');
    const placeholder = editor.view.dom.querySelector<HTMLElement>(
      `[data-document-chunk-id="${target.id}"]`,
    );
    if (!placeholder) throw new Error('Expected a chunk placeholder.');
    const estimatedHeight = Number(
      placeholder.dataset.documentChunkEstimatedHeight,
    );

    editor.commands.applyDocumentPagination(1, [
      {
        beforeBlockId: 'third-chunk-child',
        pageIndex: 1,
        spacerHeight: 120,
        remainingBodyHeight: 30,
        previousPage: testPageMetrics(),
        nextPage: testPageMetrics(),
        position: target.position + 2,
        inlineOffsetLeft: 0,
        inlineOffsetRight: 0,
      },
    ]);

    expect(placeholder).toHaveAttribute('data-document-chunk-mounted', 'false');
    expect(placeholder.dataset.documentChunkPaginationExtraHeight).toBe('120');
    expect(placeholder.style.height).toBe(`${estimatedHeight + 120}px`);
    expect(
      editor.view.dom.querySelector('.work-document-auto-page-break'),
    ).toBeNull();

    editor.commands.clearDocumentPagination(2);

    expect(placeholder.dataset.documentChunkPaginationExtraHeight).toBe('0');
    expect(placeholder.style.height).toBe(`${estimatedHeight}px`);
    editor.destroy();
  });

  test('windows pagination widgets with nested table leaf chunks', () => {
    installPassiveIntersectionObserver();
    const editor = new Editor({
      extensions: [...createWorkDocumentExtensions(), DocumentPagination],
      content: windowDocumentModel(
        {
          type: 'doc',
          content: [
            {
              type: 'documentSection',
              attrs: { id: 'document-section-1' },
              content: [tableNode(80)],
            },
          ],
        },
        {
          blockThreshold: 100,
          tableRowSize: 2,
          tableRowThreshold: 4,
        },
      ),
    });
    const leaves = leafChunkPositions(editor.state.doc);
    const first = leaves[0];
    const last = leaves.at(-1);
    if (!first || !last) throw new Error('Expected nested table leaf chunks.');
    const tableBreak = {
      tableId: 'windowed-table',
      columnCount: 1,
      colgroupHtml: '',
      repeatedHeaderRowsHtml: [],
      repeatedHeaderOverlayHtml: '',
      repeatHeaderHeight: 0,
      tableWidth: 200,
      leadingCellOffsetLeft: 0,
    };

    editor.commands.applyDocumentPagination(1, [
      {
        beforeBlockId: 'first-row',
        pageIndex: 1,
        spacerHeight: 80,
        remainingBodyHeight: 20,
        previousPage: testPageMetrics(),
        nextPage: testPageMetrics(),
        position: first.position + 2,
        inlineOffsetLeft: 0,
        inlineOffsetRight: 0,
        tableBreak,
      },
      {
        beforeBlockId: 'last-row',
        pageIndex: 2,
        spacerHeight: 90,
        remainingBodyHeight: 20,
        previousPage: testPageMetrics(),
        nextPage: testPageMetrics(),
        position: last.position + 2,
        inlineOffsetLeft: 0,
        inlineOffsetRight: 0,
        tableBreak,
      },
    ]);

    const lastContainerId = windowContainerIdAt(
      editor.state.doc,
      last.position,
    );
    const lastContainer = editor.view.dom.querySelector<HTMLElement>(
      `[data-document-chunk-id="${lastContainerId}"]`,
    );
    if (!lastContainer) throw new Error('Expected the final table window.');
    expect(
      editor.view.dom.querySelectorAll('.work-document-table-page-break'),
    ).toHaveLength(1);
    expect(lastContainer.dataset.documentChunkPaginationExtraHeight).toBe('90');
    expect(lastContainer).toHaveAttribute(
      'data-document-chunk-mounted',
      'false',
    );
    expect(
      editor.view.dom.querySelector(`[data-document-chunk-id="${last.id}"]`),
    ).toBeNull();

    editor.commands.setTextSelection(last.position + 5);

    expect(
      editor.view.dom.querySelectorAll('.work-document-table-page-break'),
    ).toHaveLength(2);
    expect(
      editor.view.dom.querySelector(`[data-document-chunk-id="${last.id}"]`),
    ).toHaveAttribute('data-document-chunk-mounted', 'true');
    editor.destroy();
  });

  test('mounts the final chunk and drops detached observer targets', async () => {
    const observer = installControlledIntersectionObserver();
    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: windowDocumentModel(documentRoot(8), {
        blockSize: 2,
        blockThreshold: 4,
      }),
    });
    const finalChunk = chunkPositions(editor.state.doc).at(-1);
    if (!finalChunk) throw new Error('Expected a final document chunk.');
    const initialElement = editor.view.dom.querySelector<HTMLElement>(
      `[data-document-chunk-id="${finalChunk.id}"]`,
    );
    if (!initialElement) throw new Error('Expected the final placeholder.');

    observer.emit(initialElement, true);
    await nextAnimationFrame();

    const mountedFinal = editor.view.dom.querySelector<HTMLElement>(
      `[data-document-chunk-id="${finalChunk.id}"]`,
    );
    expect(mountedFinal).toHaveAttribute('data-document-chunk-mounted', 'true');
    expect(mountedFinal).toHaveTextContent('Paragraph 8');
    expect(observer.targets.size).toBe(4);
    expect(
      Array.from(observer.targets).every((target) =>
        editor.view.dom.contains(target),
      ),
    ).toBe(true);
    expect(observer.targets.has(initialElement)).toBe(false);
    editor.destroy();
  });
});

function documentRoot(paragraphCount: number): WorkDocumentNode {
  return {
    type: 'doc',
    content: [
      {
        type: 'documentSection',
        attrs: { id: 'document-section-1' },
        content: Array.from({ length: paragraphCount }, (_, index) => ({
          type: 'paragraph',
          content: [{ type: 'text', text: `Paragraph ${index + 1}` }],
        })),
      },
    ],
  };
}

function tableNode(rowCount: number): WorkDocumentNode {
  return {
    type: 'table',
    attrs: { officeImported: true },
    content: Array.from({ length: rowCount }, (_, index) => ({
      type: 'tableRow',
      content: [
        {
          type: 'tableCell',
          content: [
            {
              type: 'paragraph',
              content: [{ type: 'text', text: `Row ${index + 1}` }],
            },
          ],
        },
      ],
    })),
  };
}

function chunkPositions(root: {
  descendants: (
    callback: (
      node: { attrs: Record<string, unknown>; type: { name: string } },
      position: number,
    ) => boolean | undefined,
  ) => void;
}): Array<{ id: string; position: number }> {
  const result: Array<{ id: string; position: number }> = [];
  root.descendants((node, position) => {
    if (node.type.name !== 'documentChunk') return true;
    result.push({ id: String(node.attrs.id), position });
    return false;
  });
  return result;
}

function leafChunkPositions(root: {
  descendants: (
    callback: (
      node: { attrs: Record<string, unknown>; type: { name: string } },
      position: number,
    ) => boolean | undefined,
  ) => void;
}): Array<{ id: string; position: number }> {
  const result: Array<{ id: string; position: number }> = [];
  root.descendants((node, position) => {
    if (node.type.name !== 'documentChunk') return true;
    if (node.attrs.windowContainer === true) return true;
    result.push({ id: String(node.attrs.id), position });
    return false;
  });
  return result;
}

function windowContainerIdAt(
  root: {
    resolve: (position: number) => {
      depth: number;
      node: (depth: number) => {
        attrs: Record<string, unknown>;
        type: { name: string };
      };
    };
  },
  position: number,
): string {
  const resolved = root.resolve(position);
  for (let depth = resolved.depth; depth >= 0; depth -= 1) {
    const node = resolved.node(depth);
    if (
      node.type.name === 'documentChunk' &&
      node.attrs.windowContainer === true
    ) {
      return String(node.attrs.id);
    }
  }
  throw new Error('Expected a table window container.');
}

function installPassiveIntersectionObserver(): void {
  class PassiveIntersectionObserver {
    readonly root = null;
    readonly rootMargin = '0px';
    readonly thresholds = [0];

    disconnect() {}
    observe() {}
    takeRecords(): IntersectionObserverEntry[] {
      return [];
    }
    unobserve() {}
  }
  Object.defineProperty(globalThis, 'IntersectionObserver', {
    configurable: true,
    value: PassiveIntersectionObserver,
    writable: true,
  });
}

function installControlledIntersectionObserver(): {
  emit: (target: Element, isIntersecting: boolean) => void;
  targets: Set<Element>;
} {
  const controller: {
    callback?: IntersectionObserverCallback;
    instance?: IntersectionObserver;
    targets: Set<Element>;
  } = { targets: new Set() };
  class ControlledIntersectionObserver implements IntersectionObserver {
    readonly root = null;
    readonly rootMargin = '0px';
    readonly thresholds = [0];

    constructor(callback: IntersectionObserverCallback) {
      controller.callback = callback;
      controller.instance = this;
    }

    disconnect() {
      controller.targets.clear();
    }

    observe(target: Element) {
      controller.targets.add(target);
    }

    takeRecords(): IntersectionObserverEntry[] {
      return [];
    }

    unobserve(target: Element) {
      controller.targets.delete(target);
    }
  }
  Object.defineProperty(globalThis, 'IntersectionObserver', {
    configurable: true,
    value: ControlledIntersectionObserver,
    writable: true,
  });
  return {
    emit(target, isIntersecting) {
      controller.callback?.(
        [
          {
            boundingClientRect: target.getBoundingClientRect(),
            intersectionRatio: isIntersecting ? 1 : 0,
            intersectionRect: target.getBoundingClientRect(),
            isIntersecting,
            rootBounds: null,
            target,
            time: performance.now(),
          },
        ],
        controller.instance as IntersectionObserver,
      );
    },
    targets: controller.targets,
  };
}

function nextAnimationFrame(): Promise<void> {
  return new Promise((resolve) => requestAnimationFrame(() => resolve()));
}

function testPageMetrics() {
  return {
    width: 300,
    height: 200,
    marginTop: 20,
    marginRight: 20,
    marginBottom: 20,
    marginLeft: 20,
    headerHeight: 10,
    footerHeight: 10,
    pageGap: 30,
  };
}
