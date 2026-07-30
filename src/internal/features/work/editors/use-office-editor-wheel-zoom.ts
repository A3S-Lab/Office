import { type RefObject, useEffect, useRef } from 'react';

export interface OfficeEditorWheelZoomOptions {
  enabled?: boolean;
  scopeRef: RefObject<HTMLElement | null>;
  onZoomIn: () => void;
  onZoomOut: () => void;
}

export interface OfficeZoomStepOptions {
  minimum?: number;
  maximum?: number;
  step?: number;
}

export function stepOfficeZoom(
  zoom: number,
  direction: 'in' | 'out',
  { minimum = 50, maximum = 200, step = 10 }: OfficeZoomStepOptions = {},
): number {
  const delta = direction === 'in' ? step : -step;
  return Math.min(maximum, Math.max(minimum, Math.round(zoom + delta)));
}

/**
 * Routes Ctrl/Cmd + wheel to an editor's own zoom model instead of the host
 * browser. A native non-passive capture listener is required because React's
 * delegated wheel listener cannot reliably cancel browser zoom.
 */
export function useOfficeEditorWheelZoom({
  enabled = true,
  scopeRef,
  onZoomIn,
  onZoomOut,
}: OfficeEditorWheelZoomOptions): void {
  const zoomInRef = useRef(onZoomIn);
  const zoomOutRef = useRef(onZoomOut);
  zoomInRef.current = onZoomIn;
  zoomOutRef.current = onZoomOut;

  useEffect(() => {
    const scope = scopeRef.current;
    if (!enabled || !scope) return;

    const handleWheel = (event: WheelEvent) => {
      if ((!event.ctrlKey && !event.metaKey) || event.deltaY === 0) return;
      event.preventDefault();
      event.stopPropagation();
      if (event.deltaY < 0) zoomInRef.current();
      else zoomOutRef.current();
    };

    scope.addEventListener('wheel', handleWheel, {
      capture: true,
      passive: false,
    });
    return () => {
      scope.removeEventListener('wheel', handleWheel, { capture: true });
    };
  }, [enabled, scopeRef]);
}
