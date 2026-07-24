import { expect, test } from '@rstest/core';
import { act, renderHook, waitFor } from '@testing-library/react';
import type { PointerEvent as ReactPointerEvent } from 'react';
import type { WorkSlideElement } from '../src/internal/features/work/work-types';
import type {
  PresentationGeometryController,
  PresentationSnapResolution,
} from '../src/internal/features/work/editors/use-presentation-geometry';
import {
  presentationTransformCandidate,
  usePresentationTransform,
} from '../src/internal/features/work/editors/use-presentation-transform';

test('keeps pointer transforms inside the slide', () => {
  const element = presentationElement();
  expect(
    presentationTransformCandidate(
      {
        element,
        mode: 'move',
        pointerId: 1,
        startX: 100,
        startY: 100,
      },
      { clientX: 2_000, clientY: -2_000 },
      { width: 1_000, height: 500 },
    ),
  ).toMatchObject({
    x: 80,
    y: 0,
    width: 20,
    height: 20,
  });
  expect(
    presentationTransformCandidate(
      {
        element,
        mode: 'resize',
        pointerId: 1,
        startX: 100,
        startY: 100,
      },
      { clientX: -2_000, clientY: 2_000 },
      { width: 1_000, height: 500 },
    ),
  ).toMatchObject({
    x: 10,
    y: 10,
    width: 4,
    height: 90,
  });
});

test('commits one snapped controlled value when a pointer drag ends', async () => {
  const element = presentationElement();
  const commits: Array<{
    elementId: string;
    patch: { x: number; y: number; width: number; height: number };
  }> = [];
  const snapRequests: number[] = [];
  const canvas = document.createElement('div');
  canvas.getBoundingClientRect = () =>
    ({
      bottom: 500,
      height: 500,
      left: 0,
      right: 1_000,
      top: 0,
      width: 1_000,
      x: 0,
      y: 0,
      toJSON: () => ({}),
    }) as DOMRect;
  const geometry = presentationGeometry((candidate) => {
    snapRequests.push(candidate.x);
    return {
      element: {
        id: candidate.id,
        x: 30,
        y: candidate.y,
        width: candidate.width,
        height: candidate.height,
      },
      guides: [{ axis: 'x', position: 50, source: 'slide' }],
    };
  });
  const { result } = renderHook(() =>
    usePresentationTransform({
      canvasRef: { current: canvas },
      elements: [element],
      geometry,
      onCommit: (elementId, patch) => commits.push({ elementId, patch }),
      onSelect: () => undefined,
      snapTargets: [element],
    }),
  );

  act(() => {
    result.current.beginDrag(pointer(100, 100), element, 'move');
    result.current.continueDrag(pointer(200, 100));
  });
  expect(commits).toEqual([]);

  act(() => result.current.endDrag(pointer(300, 100)));
  await waitFor(() => expect(commits).toHaveLength(1));
  expect(snapRequests).toEqual([30]);
  expect(commits[0]).toEqual({
    elementId: 'element-1',
    patch: { x: 30, y: 10, width: 20, height: 20 },
  });
});

test('previews a snapped transform without mutating the controlled value', async () => {
  const element = presentationElement();
  const commits: unknown[] = [];
  const canvas = document.createElement('div');
  canvas.getBoundingClientRect = () =>
    ({
      height: 500,
      width: 1_000,
    }) as DOMRect;
  const geometry = presentationGeometry((candidate) => ({
    element: {
      id: candidate.id,
      x: 30,
      y: candidate.y,
      width: candidate.width,
      height: candidate.height,
    },
    guides: [{ axis: 'x', position: 50, source: 'slide' }],
  }));
  const { result } = renderHook(() =>
    usePresentationTransform({
      canvasRef: { current: canvas },
      elements: [element],
      geometry,
      onCommit: (...values) => commits.push(values),
      onSelect: () => undefined,
      snapTargets: [element],
    }),
  );

  act(() => {
    result.current.beginDrag(pointer(100, 100), element, 'move');
    result.current.continueDrag(pointer(200, 100));
  });

  await waitFor(() => {
    expect(result.current.displayElements[0]?.x).toBe(30);
    expect(result.current.guides).toEqual([
      { axis: 'x', position: 50, source: 'slide' },
    ]);
  });
  expect(result.current.dragging).toBe(true);
  expect(commits).toEqual([]);

  act(() => result.current.cancelDrag());
  expect(result.current.displayElements).toEqual([element]);
  expect(result.current.guides).toEqual([]);
  expect(commits).toEqual([]);
});

test('selecting without moving does not create a controlled change', async () => {
  const element = presentationElement();
  const commits: unknown[] = [];
  const snapRequests: unknown[] = [];
  const canvas = document.createElement('div');
  canvas.getBoundingClientRect = () =>
    ({
      height: 500,
      width: 1_000,
    }) as DOMRect;
  const geometry = presentationGeometry((candidate) => {
    snapRequests.push(candidate);
    return null;
  });
  const { result } = renderHook(() =>
    usePresentationTransform({
      canvasRef: { current: canvas },
      elements: [element],
      geometry,
      onCommit: (...values) => commits.push(values),
      onSelect: () => undefined,
      snapTargets: [element],
    }),
  );

  act(() => {
    result.current.beginDrag(pointer(100, 100), element, 'move');
    result.current.endDrag(pointer(100, 100));
  });
  await waitFor(() => expect(result.current.dragging).toBe(false));
  expect(snapRequests).toEqual([]);
  expect(commits).toEqual([]);
});

function presentationElement(): WorkSlideElement {
  return {
    id: 'element-1',
    type: 'shape',
    x: 10,
    y: 10,
    width: 20,
    height: 20,
    text: '',
    fontSize: 14,
    color: '#172033',
    fill: '#dce6fb',
    bold: false,
    align: 'center',
  };
}

function presentationGeometry(
  snap: (element: WorkSlideElement) => PresentationSnapResolution | null,
): PresentationGeometryController {
  return {
    alignElement: async () => null,
    cancel: () => undefined,
    engine: 'javascript',
    pending: false,
    snapElement: async (element) => snap(element),
  };
}

function pointer(
  clientX: number,
  clientY: number,
): ReactPointerEvent<HTMLElement> {
  return {
    button: 0,
    clientX,
    clientY,
    currentTarget: {
      setPointerCapture: () => undefined,
    },
    pointerId: 1,
    stopPropagation: () => undefined,
  } as unknown as ReactPointerEvent<HTMLElement>;
}
