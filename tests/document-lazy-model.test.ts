import { Editor } from '@tiptap/core';
import { afterEach, describe, expect, test } from '@rstest/core';
import type { Node as ProseMirrorNode } from '@tiptap/pm/model';
import { documentTextStatistics } from '../src/internal/features/work/editors/document-editor-support';
import { shouldPublishDocumentUpdate } from '../src/internal/features/work/editors/document-external-content';
import { createWorkDocumentExtensions } from '../src/internal/features/work/work-document-extensions';
import { documentLazyChunkContentForEditor } from '../src/internal/features/work/work-document-chunk-node';
import {
  documentLazyChunkContent,
  materializeLazyDocumentEditorRoot,
  prepareLazyDocumentEditorSource,
} from '../src/internal/features/work/work-document-lazy-model';
import {
  createSchemaValidatedWorkDocumentModel,
  createWorkDocumentModel,
} from '../src/internal/features/work/work-document-model';
import { workDocumentSchema } from '../src/internal/features/work/work-document-model-codec';
import { measureDocumentLayoutBlocks } from '../src/internal/features/work/work-document-pagination';
import { windowDocumentModel } from '../src/internal/features/work/work-document-windowing';
import type {
  WorkDocumentContent,
  WorkDocumentNode,
} from '../src/internal/features/work/work-types';

const originalIntersectionObserver = globalThis.IntersectionObserver;

afterEach(() => {
  Object.defineProperty(globalThis, 'IntersectionObserver', {
    configurable: true,
    value: originalIntersectionObserver,
    writable: true,
  });
});

describe('lazy document editor model', () => {
  test('keeps exact ProseMirror positions while compacting off-screen chunks', () => {
    const root = windowedParagraphRoot(12);
    const model = createSchemaValidatedWorkDocumentModel(
      '<p>trusted</p>',
      root,
      {
        initialIntegrityFeatures: 0,
      },
    );

    const prepared = prepareLazyDocumentEditorSource(model, true);

    expect(prepared).not.toBeNull();
    expect(prepared?.lazyChunkCount).toBe(4);
    expect(prepared?.payloads.size).toBe(6);
    const chunks = leafChunks(prepared?.root ?? root);
    expect(chunks.slice(0, 2).every((chunk) => !isLazyChunk(chunk))).toBe(true);
    expect(chunks.slice(2).every(isLazyChunk)).toBe(true);

    const schema = workDocumentSchema();
    const complete = schema.nodeFromJSON(root);
    const compact = schema.nodeFromJSON(prepared?.root ?? root);
    expect(compact.nodeSize).toBe(complete.nodeSize);
    expect(chunkPositions(compact)).toEqual(chunkPositions(complete));
    expect(compact.textContent).toContain('paragraph 12');
    expect(compact.childCount).toBe(complete.childCount);
  });

  test('restores exact structured payloads without rebuilding untouched chunks', () => {
    const root = windowedParagraphRoot(12);
    const model = createSchemaValidatedWorkDocumentModel(
      '<p>trusted</p>',
      root,
      {
        initialIntegrityFeatures: 0,
      },
    );
    const prepared = prepareLazyDocumentEditorSource(model, true);
    if (!prepared) throw new Error('Expected a prepared lazy source.');
    const lazyChunk = leafChunks(prepared.root)[2];
    const id = String(lazyChunk?.attrs?.id ?? '');

    expect(documentLazyChunkContent(model, id)).toBe(
      leafChunks(root)[2]?.content,
    );
    expect(materializeLazyDocumentEditorRoot(prepared.root, model)).toEqual(
      root,
    );
  });

  test('does not compact an untrusted external model', () => {
    const root = windowedParagraphRoot(12);
    const model = createWorkDocumentModel('<p>external</p>', root);

    expect(prepareLazyDocumentEditorSource(model, false)).toBeNull();
  });

  test('hydrates a selected leaf without publishing a controlled edit', async () => {
    installPassiveIntersectionObserver();
    const root = windowedParagraphRoot(12);
    const model = createSchemaValidatedWorkDocumentModel(
      '<p>trusted</p>',
      root,
      {
        initialIntegrityFeatures: 0,
      },
    );
    const content: WorkDocumentContent = {
      type: 'document',
      html: '<p>trusted</p>',
      model,
      pageSize: 'a4',
    };
    const prepared = prepareLazyDocumentEditorSource(model, true);
    if (!prepared) throw new Error('Expected a prepared lazy source.');
    let publishedUpdates = 0;
    const editor = new Editor({
      extensions: createWorkDocumentExtensions({
        getContent: () => content,
        trustInitialIntegrityFeatures: true,
      }),
      content: prepared.root,
      onUpdate: ({ appendedTransactions, transaction }) => {
        if (
          shouldPublishDocumentUpdate(transaction, appendedTransactions ?? [])
        ) {
          publishedUpdates += 1;
        }
      },
    });
    const before = chunkPositions(editor.state.doc);
    const third = before[2];
    if (!third) throw new Error('Expected a third document chunk.');
    expect(editor.state.doc.nodeAt(third.position)?.firstChild?.type.name).toBe(
      'documentLazyBlock',
    );
    expect(documentLazyChunkContentForEditor(editor, third.id)).toHaveLength(2);

    editor.commands.setTextSelection(third.position + 2);
    expect(editor.state.selection.anchor).toBe(third.position + 2);
    await flushMicrotasks();

    expect(editor.state.doc.nodeAt(third.position)?.firstChild?.type.name).toBe(
      'paragraph',
    );
    expect(chunkPositions(editor.state.doc)).toEqual(before);
    expect(editor.state.selection.anchor).toBe(third.position + 2);
    expect(publishedUpdates).toBe(0);
    editor.destroy();
  });

  test('measures and counts the complete logical document before hydration', () => {
    installPassiveIntersectionObserver();
    const root = windowedParagraphRoot(12);
    const model = createSchemaValidatedWorkDocumentModel(
      '<p>trusted</p>',
      root,
      {
        initialIntegrityFeatures: 0,
      },
    );
    const content: WorkDocumentContent = {
      type: 'document',
      html: '<p>trusted</p>',
      model,
      pageSize: 'a4',
    };
    const prepared = prepareLazyDocumentEditorSource(model, true);
    if (!prepared) throw new Error('Expected a prepared lazy source.');
    const editor = new Editor({
      extensions: createWorkDocumentExtensions({ getContent: () => content }),
      content: prepared.root,
    });

    expect(measureDocumentLayoutBlocks(editor).blocks).toHaveLength(12);
    expect(documentTextStatistics(editor)).toEqual({
      characterCountWithSpaces: 135,
      characterCountWithoutSpaces: 123,
      paragraphCount: 12,
      wordCount: 24,
    });
    editor.destroy();
  });

  test('renders a scrolling preview and hydrates only when editing starts', async () => {
    const observer = installControlledIntersectionObserver();
    const root = windowedParagraphRoot(12);
    const model = createSchemaValidatedWorkDocumentModel(
      '<p>trusted</p>',
      root,
      {
        initialIntegrityFeatures: 0,
      },
    );
    const content: WorkDocumentContent = {
      type: 'document',
      html: '<p>trusted</p>',
      model,
      pageSize: 'a4',
    };
    const prepared = prepareLazyDocumentEditorSource(model, true);
    if (!prepared) throw new Error('Expected a prepared lazy source.');
    const editor = new Editor({
      extensions: createWorkDocumentExtensions({ getContent: () => content }),
      content: prepared.root,
    });
    const third = chunkPositions(editor.state.doc)[2];
    if (!third) throw new Error('Expected a third document chunk.');
    const placeholder = editor.view.dom.querySelector<HTMLElement>(
      `[data-document-chunk-id="${third.id}"]`,
    );
    if (!placeholder) throw new Error('Expected a lazy chunk placeholder.');

    observer.emit(placeholder, true);
    await flushMicrotasks();

    const preview = editor.view.dom.querySelector<HTMLElement>(
      `[data-document-chunk-id="${third.id}"]`,
    );
    expect(preview).toHaveAttribute('data-document-lazy-preview', 'true');
    expect(preview).toHaveTextContent('paragraph 5paragraph 6');
    expect(editor.state.doc.nodeAt(third.position)?.firstChild?.type.name).toBe(
      'documentLazyBlock',
    );
    const statisticsBeforeHydration = documentTextStatistics(editor);
    const pooledParagraph = preview?.querySelector('p');
    const fourth = chunkPositions(editor.state.doc)[3];
    if (!fourth) throw new Error('Expected a fourth document chunk.');

    observer.emit(preview as HTMLElement, false);
    await flushMicrotasks();
    const fourthPlaceholder = editor.view.dom.querySelector<HTMLElement>(
      `[data-document-chunk-id="${fourth.id}"]`,
    );
    if (!fourthPlaceholder) throw new Error('Expected a fourth placeholder.');
    observer.emit(fourthPlaceholder, true);
    await flushMicrotasks();

    const fourthPreview = editor.view.dom.querySelector<HTMLElement>(
      `[data-document-chunk-id="${fourth.id}"]`,
    );
    expect(fourthPreview).toHaveTextContent('paragraph 7paragraph 8');
    expect(fourthPreview?.querySelector('p')).toBe(pooledParagraph);

    fourthPreview?.dispatchEvent(
      new PointerEvent('pointerdown', { bubbles: true, button: 0 }),
    );
    await flushMicrotasks();

    expect(
      editor.state.doc.nodeAt(fourth.position)?.firstChild?.type.name,
    ).toBe('paragraph');
    expect(documentTextStatistics(editor)).toBe(statisticsBeforeHydration);
    expect(
      editor.view.dom.querySelector(`[data-document-chunk-id="${fourth.id}"]`),
    ).not.toHaveAttribute('data-document-lazy-preview');
    editor.destroy();
  });
});

function windowedParagraphRoot(count: number): WorkDocumentNode {
  return windowDocumentModel(
    {
      type: 'doc',
      content: [
        {
          type: 'documentSection',
          attrs: { id: 'document-section-1' },
          content: Array.from({ length: count }, (_, index) => ({
            type: 'paragraph',
            content: [{ type: 'text', text: `paragraph ${index + 1}` }],
          })),
        },
      ],
    },
    {
      blockSize: 2,
      blockThreshold: 4,
      tableRowSize: 2,
      tableRowThreshold: 4,
      trustedIntegrityFeatures: 0,
    },
  );
}

function leafChunks(root: WorkDocumentNode): WorkDocumentNode[] {
  const result: WorkDocumentNode[] = [];
  const pending = [root];
  while (pending.length) {
    const node = pending.shift();
    if (!node) continue;
    if (node.type === 'documentChunk' && node.attrs?.windowContainer !== true) {
      result.push(node);
      continue;
    }
    pending.unshift(...(node.content ?? []));
  }
  return result;
}

function isLazyChunk(node: WorkDocumentNode | undefined): boolean {
  return node?.content?.[0]?.type === 'documentLazyBlock';
}

function chunkPositions(document: ProseMirrorNode) {
  const result: Array<{ id: string; nodeSize: number; position: number }> = [];
  document.descendants((node, position) => {
    if (node.type.name !== 'documentChunk') return true;
    if (node.attrs.windowContainer === true) return true;
    result.push({
      id: String(node.attrs.id),
      nodeSize: node.nodeSize,
      position,
    });
    return false;
  });
  return result;
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
  };
}

async function flushMicrotasks(): Promise<void> {
  await Promise.resolve();
  await Promise.resolve();
  await Promise.resolve();
}
