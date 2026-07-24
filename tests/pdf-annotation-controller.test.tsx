import type { PluginRegistry } from '@embedpdf/react-pdf-viewer';
import { expect, test } from '@rstest/core';
import { act, renderHook, waitFor } from '@testing-library/react';
import { usePdfAnnotationController } from '../src/internal/features/work/editors/pdf-annotation-controller';

test('selects and deletes PDF annotations through the annotation capability', async () => {
  const calls: string[] = [];
  const annotationChange = createEvent();
  const activeToolChange = createEvent();
  const documentChange = createEvent();
  const documentOpened = createEvent();
  const documentClosed = createEvent();
  const annotationState = {
    pages: { 0: ['annotation-1'], 2: ['annotation-2'] },
    byUid: {},
    selectedUids: ['annotation-2'],
    selectedUid: 'annotation-2',
    activeToolId: 'highlight',
    hasPendingChanges: true,
    locked: { type: 'none' },
  };
  const capabilities = {
    annotation: {
      forDocument: () => ({
        getState: () => annotationState,
        getSelectedAnnotationIds: () => annotationState.selectedUids,
        setActiveTool: (toolId: string | null) =>
          calls.push(`tool:${toolId ?? 'pointer'}`),
        deleteAnnotations: (
          annotations: Array<{ pageIndex: number; id: string }>,
        ) => calls.push(`delete:${JSON.stringify(annotations)}`),
      }),
      getTool: (toolId: string) =>
        ['highlight', 'underline', 'strikeout', 'ink', 'freeText'].includes(
          toolId,
        )
          ? { id: toolId }
          : undefined,
      onStateChange: annotationChange.subscribe,
      onActiveToolChange: activeToolChange.subscribe,
    },
    'document-manager': {
      getActiveDocumentId: () => 'document-1',
      onActiveDocumentChanged: documentChange.subscribe,
      onDocumentOpened: documentOpened.subscribe,
      onDocumentClosed: documentClosed.subscribe,
    },
  };
  const registry = {
    pluginsReady: () => Promise.resolve(),
    getPlugin: (id: keyof typeof capabilities) => ({
      provides: () => capabilities[id],
    }),
  } as unknown as PluginRegistry;

  const { result } = renderHook(() => usePdfAnnotationController(registry));
  await waitFor(() => expect(result.current.state.available).toBe(true));
  expect(result.current.state).toMatchObject({
    activeToolId: 'highlight',
    hasPendingChanges: true,
    selectedCount: 1,
  });

  act(() => {
    result.current.selectTool('ink');
    result.current.selectTool(null);
    result.current.deleteSelection();
  });

  expect(calls).toEqual([
    'tool:ink',
    'tool:pointer',
    'delete:[{"pageIndex":2,"id":"annotation-2"}]',
  ]);
});

function createEvent<T = unknown>() {
  const listeners = new Set<(value: T) => void>();
  return {
    emit: (value: T) => {
      for (const listener of listeners) listener(value);
    },
    subscribe: (listener: (value: T) => void) => {
      listeners.add(listener);
      return () => listeners.delete(listener);
    },
  };
}
