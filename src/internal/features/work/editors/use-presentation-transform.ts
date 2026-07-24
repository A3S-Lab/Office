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
import { expandPresentationGroupSelection } from '../work-presentation-groups';
import { clamp } from './presentation-editor-operations';
import { presentationSelectionBounds } from './presentation-selection';
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
  elements?: readonly WorkSlideElement[];
  frame?: WorkSlideElement;
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
  borderWidth?: number;
  fontSize?: number;
  textRuns?: WorkSlideElement['textRuns'];
}

export interface PresentationTransformCommit {
  elementId: string;
  patch: PresentationTransformPatch;
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
  selectedElementIds,
  snapTargets,
}: {
  canvasRef: RefObject<HTMLElement | null>;
  elements: WorkSlideElement[];
  geometry: PresentationGeometryController;
  onCommit: (changes: readonly PresentationTransformCommit[]) => void;
  onSelect: (elementId: string) => void;
  selectedElementIds: readonly string[];
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
  const [preview, setPreview] = useState<WorkSlideElement[] | null>(null);
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
      const frame = drag.frame ?? drag.element;
      const candidate = presentationTransformCandidate(
        { ...drag, element: frame },
        point,
        bounds,
      );
      if (commit && samePresentationTransform(candidate, frame)) {
        setPreview(null);
        setGuides([]);
        return;
      }
      const requestSequence = ++requestSequenceRef.current;
      setPreview(presentationTransformPreview(drag, candidate));
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
        snapTargetsRef.current.filter(
          (target) =>
            !(drag.elements ?? [drag.element]).some(
              (element) => element.id === target.id,
            ),
        ),
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
        const resolvedElements = presentationTransformPreview(drag, resolved);
        setPreview(resolvedElements);
        setGuides(resolution?.guides ?? []);
        if (!commit) return;
        const originals = drag.elements ?? [drag.element];
        const current = originals.map((original) =>
          elementsRef.current.find((element) => element.id === original.id),
        );
        if (
          current.some(
            (element, index) =>
              !element || !samePresentationTransform(element, originals[index]),
          ) ||
          samePresentationTransform(resolved, frame)
        ) {
          setPreview(null);
          setGuides([]);
          return;
        }
        const changes = resolvedElements
          .filter(
            (element, index) =>
              !samePresentationTransform(element, originals[index]),
          )
          .map((element, index) => ({
            elementId: element.id,
            patch: presentationTransformPatch(element, originals[index]),
          }));
        if (changes.length) onCommitRef.current(changes);
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
      const selected = new Set(selectedElementIds);
      const clickedUnit = new Set(
        expandPresentationGroupSelection(elements, [element.id]),
      );
      const transforming = selected.has(element.id) ? selected : clickedUnit;
      const dragElements = elements.filter((candidate) =>
        transforming.has(candidate.id),
      );
      dragRef.current = {
        element,
        elements: dragElements,
        frame: presentationTransformFrame(dragElements, element),
        mode,
        pointerId: event.pointerId,
        startX: event.clientX,
        startY: event.clientY,
      };
      setPreview(dragElements);
      setGuides([]);
      if (!selected.has(element.id)) onSelect(element.id);
    },
    [cancelGeometry, clearFrame, elements, onSelect, selectedElementIds],
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
    const originals = drag.elements ?? [drag.element];
    if (
      originals.some((original) => {
        const current = elements.find((element) => element.id === original.id);
        return !current || !samePresentationTransform(current, original);
      })
    ) {
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

  const displayElements = useMemo(() => {
    if (!preview) return elements;
    const previews = new Map(preview.map((element) => [element.id, element]));
    return elements.map((element) => previews.get(element.id) ?? element);
  }, [elements, preview]);

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

function presentationTransformFrame(
  elements: readonly WorkSlideElement[],
  fallback: WorkSlideElement,
): WorkSlideElement {
  const bounds = presentationSelectionBounds(elements);
  if (!bounds) return fallback;
  return {
    ...fallback,
    x: bounds.left,
    y: bounds.top,
    width: bounds.width,
    height: bounds.height,
  };
}

function presentationTransformPreview(
  drag: PresentationTransformDrag,
  frame: WorkSlideElement,
): WorkSlideElement[] {
  const originals = drag.elements ?? [drag.element];
  const originalFrame = drag.frame ?? drag.element;
  if (drag.mode === 'resize') {
    if (originals.length === 1) return [frame];
    const scaleX =
      originalFrame.width > 0 ? frame.width / originalFrame.width : 1;
    const scaleY =
      originalFrame.height > 0 ? frame.height / originalFrame.height : 1;
    const visualScale = Math.min(scaleX, scaleY);
    return originals.map((element) => ({
      ...element,
      x: frame.x + (element.x - originalFrame.x) * scaleX,
      y: frame.y + (element.y - originalFrame.y) * scaleY,
      width: element.width * scaleX,
      height: element.height * scaleY,
      ...presentationScaledVisuals(element, visualScale),
    }));
  }
  const deltaX = frame.x - originalFrame.x;
  const deltaY = frame.y - originalFrame.y;
  return originals.map((element) => ({
    ...element,
    x: element.x + deltaX,
    y: element.y + deltaY,
  }));
}

function presentationTransformPatch(
  element: WorkSlideElement,
  original: WorkSlideElement,
): PresentationTransformPatch {
  return {
    x: element.x,
    y: element.y,
    width: element.width,
    height: element.height,
    ...(element.borderWidth !== original.borderWidth
      ? { borderWidth: element.borderWidth }
      : {}),
    ...(element.fontSize !== original.fontSize
      ? { fontSize: element.fontSize }
      : {}),
    ...(element.textRuns !== original.textRuns
      ? { textRuns: element.textRuns }
      : {}),
  };
}

function presentationScaledVisuals(
  element: WorkSlideElement,
  scale: number,
): Pick<WorkSlideElement, 'fontSize'> &
  Partial<Pick<WorkSlideElement, 'borderWidth' | 'textRuns'>> {
  const fontSize = presentationElementScalesTypography(element)
    ? scaledPresentationFontSize(element.fontSize, scale)
    : element.fontSize;
  const borderWidth =
    element.borderWidth === undefined
      ? undefined
      : scaledPresentationBorderWidth(element.borderWidth, scale);
  const textRuns = scaledPresentationTextRuns(element, scale);
  return {
    fontSize,
    ...(borderWidth !== undefined ? { borderWidth } : {}),
    ...(textRuns !== element.textRuns ? { textRuns } : {}),
  };
}

function presentationElementScalesTypography(
  element: WorkSlideElement,
): boolean {
  return (
    element.type === 'text' ||
    element.type === 'shape' ||
    element.type === 'table' ||
    Boolean(element.text || element.textRuns?.length)
  );
}

function scaledPresentationTextRuns(
  element: WorkSlideElement,
  scale: number,
): WorkSlideElement['textRuns'] {
  if (!element.textRuns?.some((run) => run.fontSize !== undefined)) {
    return element.textRuns;
  }
  return element.textRuns.map((run) =>
    run.fontSize === undefined
      ? run
      : { ...run, fontSize: scaledPresentationFontSize(run.fontSize, scale) },
  );
}

function scaledPresentationFontSize(value: number, scale: number): number {
  return roundedPresentationMetric(clamp(value * scale, 1, 400));
}

function scaledPresentationBorderWidth(value: number, scale: number): number {
  if (value === 0) return 0;
  return roundedPresentationMetric(Math.max(0.1, value * scale));
}

function roundedPresentationMetric(value: number): number {
  return Math.round(value * 1_000) / 1_000;
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
