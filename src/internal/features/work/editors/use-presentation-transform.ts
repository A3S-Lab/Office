import {
  useCallback,
  useEffect,
  useMemo,
  useRef,
  useState,
  type PointerEvent as ReactPointerEvent,
  type RefObject,
} from 'react';
import type { OfficeKernelPresentationSnapGuide } from '../../../kernel/office-kernel-protocol';
import type { WorkSlideElement } from '../work-types';
import { clamp } from './presentation-editor-operations';
import type { PresentationGeometryController } from './use-presentation-geometry';

const PRESENTATION_SNAP_DISTANCE_PIXELS = 6;
const PRESENTATION_FRAME_FALLBACK_MILLISECONDS = 48;

interface PresentationFrame {
  animationFrame: number | null;
  cancelled: boolean;
  timeout: number;
}

export interface PresentationTransformDrag {
  element: WorkSlideElement;
  mode: 'move' | 'resize';
  pointerId: number;
  startX: number;
  startY: number;
}

export interface PresentationTransformPoint {
  clientX: number;
  clientY: number;
}

export interface PresentationTransformBounds {
  width: number;
  height: number;
}

export interface PresentationTransformPatch {
  x: number;
  y: number;
  width: number;
  height: number;
}

export interface PresentationTransformController {
  cancelDrag: () => void;
  displayElements: WorkSlideElement[];
  dragging: boolean;
  guides: OfficeKernelPresentationSnapGuide[];
  beginDrag: (
    event: ReactPointerEvent,
    element: WorkSlideElement,
    mode: 'move' | 'resize',
  ) => void;
  continueDrag: (event: ReactPointerEvent) => void;
  endDrag: (event: ReactPointerEvent) => void;
}

export function usePresentationTransform({
  canvasRef,
  elements,
  geometry,
  onCommit,
  onSelect,
  snapTargets,
}: {
  canvasRef: RefObject<HTMLElement | null>;
  elements: WorkSlideElement[];
  geometry: PresentationGeometryController;
  onCommit: (elementId: string, patch: PresentationTransformPatch) => void;
  onSelect: (elementId: string) => void;
  snapTargets: WorkSlideElement[];
}): PresentationTransformController {
  const dragRef = useRef<PresentationTransformDrag | null>(null);
  const elementsRef = useRef(elements);
  const snapTargetsRef = useRef(snapTargets);
  const onCommitRef = useRef(onCommit);
  const frameRef = useRef<PresentationFrame | null>(null);
  const pendingPointRef = useRef<PresentationTransformPoint | null>(null);
  const generationRef = useRef(0);
  const requestSequenceRef = useRef(0);
  const [preview, setPreview] = useState<WorkSlideElement | null>(null);
  const [guides, setGuides] = useState<OfficeKernelPresentationSnapGuide[]>([]);
  const cancelGeometry = geometry.cancel;
  const snapElement = geometry.snapElement;
  elementsRef.current = elements;
  snapTargetsRef.current = snapTargets;
  onCommitRef.current = onCommit;

  const clearFrame = useCallback(() => {
    if (frameRef.current !== null) cancelPresentationFrame(frameRef.current);
    frameRef.current = null;
    pendingPointRef.current = null;
  }, []);

  const cancelDrag = useCallback(() => {
    generationRef.current += 1;
    requestSequenceRef.current += 1;
    dragRef.current = null;
    clearFrame();
    cancelGeometry();
    setPreview(null);
    setGuides([]);
  }, [cancelGeometry, clearFrame]);

  const resolveCandidate = useCallback(
    (
      drag: PresentationTransformDrag,
      point: PresentationTransformPoint,
      generation: number,
      commit: boolean,
    ): void => {
      const bounds = canvasRef.current?.getBoundingClientRect();
      if (!bounds?.width || !bounds.height) return;
      const candidate = presentationTransformCandidate(drag, point, bounds);
      if (commit && samePresentationTransform(candidate, drag.element)) {
        setPreview(null);
        setGuides([]);
        return;
      }
      const requestSequence = ++requestSequenceRef.current;
      setPreview(candidate);
      if (!commit) setGuides([]);
      const threshold = {
        x: Math.min(
          10,
          (PRESENTATION_SNAP_DISTANCE_PIXELS / bounds.width) * 100,
        ),
        y: Math.min(
          10,
          (PRESENTATION_SNAP_DISTANCE_PIXELS / bounds.height) * 100,
        ),
      };
      void snapElement(
        candidate,
        snapTargetsRef.current,
        drag.mode,
        threshold,
      ).then((resolution) => {
        if (
          generation !== generationRef.current ||
          requestSequence !== requestSequenceRef.current
        ) {
          return;
        }
        const resolved = resolution
          ? {
              ...candidate,
              x: resolution.element.x,
              y: resolution.element.y,
              width: resolution.element.width,
              height: resolution.element.height,
            }
          : candidate;
        setPreview(resolved);
        setGuides(resolution?.guides ?? []);
        if (!commit) return;
        const current = elementsRef.current.find(
          (element) => element.id === drag.element.id,
        );
        if (
          !current ||
          !samePresentationTransform(current, drag.element) ||
          samePresentationTransform(resolved, drag.element)
        ) {
          setPreview(null);
          setGuides([]);
          return;
        }
        onCommitRef.current(
          drag.element.id,
          presentationTransformPatch(resolved),
        );
        setPreview(null);
        setGuides([]);
      });
    },
    [canvasRef, snapElement],
  );

  const beginDrag = useCallback(
    (
      event: ReactPointerEvent,
      element: WorkSlideElement,
      mode: 'move' | 'resize',
    ) => {
      if (event.button !== 0) return;
      event.stopPropagation();
      event.currentTarget.setPointerCapture(event.pointerId);
      clearFrame();
      cancelGeometry();
      generationRef.current += 1;
      requestSequenceRef.current += 1;
      dragRef.current = {
        element,
        mode,
        pointerId: event.pointerId,
        startX: event.clientX,
        startY: event.clientY,
      };
      setPreview(element);
      setGuides([]);
      onSelect(element.id);
    },
    [cancelGeometry, clearFrame, onSelect],
  );

  const continueDrag = useCallback(
    (event: ReactPointerEvent) => {
      const drag = dragRef.current;
      if (!drag || drag.pointerId !== event.pointerId) return;
      pendingPointRef.current = {
        clientX: event.clientX,
        clientY: event.clientY,
      };
      if (frameRef.current !== null) return;
      const generation = generationRef.current;
      frameRef.current = requestPresentationFrame(() => {
        frameRef.current = null;
        const current = dragRef.current;
        const point = pendingPointRef.current;
        pendingPointRef.current = null;
        if (
          !current ||
          !point ||
          current.pointerId !== drag.pointerId ||
          generation !== generationRef.current
        ) {
          return;
        }
        resolveCandidate(current, point, generation, false);
      });
    },
    [resolveCandidate],
  );

  const endDrag = useCallback(
    (event: ReactPointerEvent) => {
      const drag = dragRef.current;
      if (!drag || drag.pointerId !== event.pointerId) return;
      dragRef.current = null;
      clearFrame();
      resolveCandidate(
        drag,
        { clientX: event.clientX, clientY: event.clientY },
        generationRef.current,
        true,
      );
    },
    [clearFrame, resolveCandidate],
  );

  useEffect(() => {
    const drag = dragRef.current;
    if (!drag) return;
    const current = elements.find((element) => element.id === drag.element.id);
    if (!current || !samePresentationTransform(current, drag.element)) {
      cancelDrag();
    }
  }, [cancelDrag, elements]);

  useEffect(
    () => () => {
      generationRef.current += 1;
      requestSequenceRef.current += 1;
      dragRef.current = null;
      clearFrame();
    },
    [clearFrame],
  );

  const displayElements = useMemo(
    () =>
      preview
        ? elements.map((element) =>
            element.id === preview.id ? preview : element,
          )
        : elements,
    [elements, preview],
  );

  return {
    beginDrag,
    cancelDrag,
    continueDrag,
    displayElements,
    dragging: preview !== null || dragRef.current !== null,
    endDrag,
    guides,
  };
}

export function presentationTransformCandidate(
  drag: PresentationTransformDrag,
  point: PresentationTransformPoint,
  bounds: PresentationTransformBounds,
): WorkSlideElement {
  const dx = ((point.clientX - drag.startX) / bounds.width) * 100;
  const dy = ((point.clientY - drag.startY) / bounds.height) * 100;
  if (drag.mode === 'move') {
    return {
      ...drag.element,
      x: clamp(drag.element.x + dx, 0, 100 - drag.element.width),
      y: clamp(drag.element.y + dy, 0, 100 - drag.element.height),
    };
  }
  return {
    ...drag.element,
    width: clamp(drag.element.width + dx, 4, 100 - drag.element.x),
    height: clamp(drag.element.height + dy, 4, 100 - drag.element.y),
  };
}

function presentationTransformPatch(
  element: WorkSlideElement,
): PresentationTransformPatch {
  return {
    x: element.x,
    y: element.y,
    width: element.width,
    height: element.height,
  };
}

function samePresentationTransform(
  left: WorkSlideElement,
  right: WorkSlideElement,
): boolean {
  return (
    left.x === right.x &&
    left.y === right.y &&
    left.width === right.width &&
    left.height === right.height
  );
}

function requestPresentationFrame(
  callback: FrameRequestCallback,
): PresentationFrame {
  const frame: PresentationFrame = {
    animationFrame: null,
    cancelled: false,
    timeout: 0,
  };
  const invoke = (timestamp = presentationFrameTimestamp()) => {
    if (frame.cancelled) return;
    frame.cancelled = true;
    if (
      frame.animationFrame !== null &&
      typeof cancelAnimationFrame === 'function'
    ) {
      cancelAnimationFrame(frame.animationFrame);
    }
    window.clearTimeout(frame.timeout);
    callback(timestamp);
  };
  frame.timeout = window.setTimeout(
    invoke,
    PRESENTATION_FRAME_FALLBACK_MILLISECONDS,
  );
  if (typeof requestAnimationFrame === 'function') {
    frame.animationFrame = requestAnimationFrame(invoke);
  }
  return frame;
}

function cancelPresentationFrame(frame: PresentationFrame): void {
  frame.cancelled = true;
  if (
    frame.animationFrame !== null &&
    typeof cancelAnimationFrame === 'function'
  ) {
    cancelAnimationFrame(frame.animationFrame);
  }
  window.clearTimeout(frame.timeout);
}

function presentationFrameTimestamp(): number {
  return typeof performance === 'undefined' ? Date.now() : performance.now();
}
