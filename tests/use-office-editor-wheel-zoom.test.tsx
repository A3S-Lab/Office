import { expect, test } from '@rstest/core';
import { renderHook } from '@testing-library/react';
import {
  stepOfficeZoom,
  useOfficeEditorWheelZoom,
} from '../src/internal/features/work/editors/use-office-editor-wheel-zoom';

test('steps editor zoom within the same bounds used by status controls', () => {
  expect(stepOfficeZoom(100, 'in')).toBe(110);
  expect(stepOfficeZoom(100, 'out')).toBe(90);
  expect(stepOfficeZoom(200, 'in')).toBe(200);
  expect(stepOfficeZoom(50, 'out')).toBe(50);
  expect(stepOfficeZoom(175, 'in', { minimum: 60, maximum: 180 })).toBe(180);
});

test('routes modified wheel gestures to editor zoom and cancels browser zoom', () => {
  const scope = document.createElement('section');
  const child = document.createElement('div');
  scope.append(child);
  document.body.append(scope);
  const directions: string[] = [];

  const { unmount } = renderHook(() =>
    useOfficeEditorWheelZoom({
      scopeRef: { current: scope },
      onZoomIn: () => directions.push('in'),
      onZoomOut: () => directions.push('out'),
    }),
  );

  try {
    const zoomIn = wheelEvent({ ctrlKey: true, deltaY: -100 });
    child.dispatchEvent(zoomIn);
    expect(zoomIn.defaultPrevented).toBe(true);

    const zoomOut = wheelEvent({ deltaY: 100, metaKey: true });
    child.dispatchEvent(zoomOut);
    expect(zoomOut.defaultPrevented).toBe(true);
    expect(directions).toEqual(['in', 'out']);

    const ordinaryScroll = wheelEvent({ deltaY: 100 });
    child.dispatchEvent(ordinaryScroll);
    expect(ordinaryScroll.defaultPrevented).toBe(false);
    expect(directions).toEqual(['in', 'out']);

    const horizontalModifiedScroll = wheelEvent({
      ctrlKey: true,
      deltaX: 100,
    });
    child.dispatchEvent(horizontalModifiedScroll);
    expect(horizontalModifiedScroll.defaultPrevented).toBe(false);
    expect(directions).toEqual(['in', 'out']);
  } finally {
    unmount();
    scope.remove();
  }
});

test('does not intercept wheel gestures outside its scope or while disabled', () => {
  const scope = document.createElement('section');
  const outside = document.createElement('div');
  document.body.append(scope, outside);
  let zoomCount = 0;

  const { unmount } = renderHook(() =>
    useOfficeEditorWheelZoom({
      enabled: false,
      scopeRef: { current: scope },
      onZoomIn: () => {
        zoomCount += 1;
      },
      onZoomOut: () => {
        zoomCount += 1;
      },
    }),
  );

  try {
    const inside = wheelEvent({ ctrlKey: true, deltaY: -100 });
    scope.dispatchEvent(inside);
    const elsewhere = wheelEvent({ ctrlKey: true, deltaY: -100 });
    outside.dispatchEvent(elsewhere);

    expect(inside.defaultPrevented).toBe(false);
    expect(elsewhere.defaultPrevented).toBe(false);
    expect(zoomCount).toBe(0);
  } finally {
    unmount();
    scope.remove();
    outside.remove();
  }
});

function wheelEvent(
  init: Pick<WheelEventInit, 'deltaY'> & Partial<WheelEventInit>,
): WheelEvent {
  const event = new WheelEvent('wheel', {
    bubbles: true,
    cancelable: true,
    ...init,
  });
  // Happy DOM currently omits modifier fields from WheelEventInit.
  Object.defineProperties(event, {
    ctrlKey: { configurable: true, value: Boolean(init.ctrlKey) },
    metaKey: { configurable: true, value: Boolean(init.metaKey) },
  });
  return event;
}
