import { useCallback, useEffect, useRef, useState } from 'react';
import {
  createOfficeKernelClient,
  type OfficeKernelClient,
} from '../../../kernel/office-kernel-client';
import type {
  OfficeKernelEngine,
  OfficeKernelPresentationAlignment,
  OfficeKernelPresentationGeometryElement,
  OfficeKernelPresentationGeometryOperation,
  OfficeKernelPresentationGeometryResult,
  OfficeKernelPresentationSnapGuide,
  OfficeKernelPresentationTransformMode,
} from '../../../kernel/office-kernel-protocol';
import type { WorkSlideElement } from '../work-types';

export interface PresentationSnapResolution {
  element: OfficeKernelPresentationGeometryElement;
  guides: OfficeKernelPresentationSnapGuide[];
}

export interface PresentationGeometryController {
  alignElement: (
    element: WorkSlideElement,
    alignment: OfficeKernelPresentationAlignment,
  ) => Promise<OfficeKernelPresentationGeometryElement | null>;
  cancel: () => void;
  engine: OfficeKernelEngine | null;
  pending: boolean;
  snapElement: (
    element: WorkSlideElement,
    elements: readonly WorkSlideElement[],
    mode: OfficeKernelPresentationTransformMode,
    threshold: { x: number; y: number },
  ) => Promise<PresentationSnapResolution | null>;
}

export function usePresentationGeometry(
  wasmUrl?: string,
  enabled = true,
): PresentationGeometryController {
  const clientRef = useRef<OfficeKernelClient | null>(null);
  const activeRequest = useRef<AbortController | null>(null);
  const revision = useRef(0);
  const [engine, setEngine] = useState<OfficeKernelEngine | null>(null);
  const [pending, setPending] = useState(false);

  useEffect(() => {
    if (!enabled) {
      clientRef.current = null;
      return;
    }
    const client = createOfficeKernelClient(wasmUrl);
    clientRef.current = client;
    return () => {
      activeRequest.current?.abort();
      activeRequest.current = null;
      clientRef.current = null;
      client.dispose();
    };
  }, [enabled, wasmUrl]);

  const cancel = useCallback(() => {
    revision.current += 1;
    activeRequest.current?.abort();
    activeRequest.current = null;
    setPending(false);
  }, []);

  const resolveGeometry = useCallback(
    async (
      operation: OfficeKernelPresentationGeometryOperation,
      elements: readonly WorkSlideElement[],
    ): Promise<OfficeKernelPresentationGeometryResult | null> => {
      const client = clientRef.current;
      if (!client) return null;
      const requestRevision = ++revision.current;
      activeRequest.current?.abort();
      const controller = new AbortController();
      activeRequest.current = controller;
      setPending(true);
      try {
        const result = await client.presentationGeometry(
          {
            revision: requestRevision,
            documentRevision: requestRevision,
            operation,
            elements: elements.map(presentationGeometryElement),
          },
          controller.signal,
        );
        if (controller.signal.aborted || requestRevision !== revision.current) {
          return null;
        }
        setEngine(result.engine);
        return result;
      } catch (error) {
        if (
          controller.signal.aborted ||
          (error instanceof DOMException && error.name === 'AbortError')
        ) {
          return null;
        }
        setEngine(null);
        return null;
      } finally {
        if (requestRevision === revision.current) {
          activeRequest.current = null;
          setPending(false);
        }
      }
    },
    [],
  );

  const alignElement = useCallback(
    async (
      element: WorkSlideElement,
      alignment: OfficeKernelPresentationAlignment,
    ): Promise<OfficeKernelPresentationGeometryElement | null> => {
      const result = await resolveGeometry(
        { type: 'alignToSlide', alignment },
        [element],
      );
      return (
        result?.elements.find((candidate) => candidate.id === element.id) ??
        null
      );
    },
    [resolveGeometry],
  );

  const snapElement = useCallback(
    async (
      element: WorkSlideElement,
      elements: readonly WorkSlideElement[],
      mode: OfficeKernelPresentationTransformMode,
      threshold: { x: number; y: number },
    ): Promise<PresentationSnapResolution | null> => {
      const candidates = uniquePresentationGeometryElements(element, elements);
      const result = await resolveGeometry(
        {
          type: 'snapElement',
          movingElementId: element.id,
          mode,
          thresholdX: threshold.x,
          thresholdY: threshold.y,
        },
        candidates,
      );
      const resolved = result?.elements.find(
        (candidate) => candidate.id === element.id,
      );
      return resolved
        ? {
            element: resolved,
            guides: result?.guides ?? [],
          }
        : null;
    },
    [resolveGeometry],
  );

  return { alignElement, cancel, engine, pending, snapElement };
}

function presentationGeometryElement(
  element: WorkSlideElement,
): OfficeKernelPresentationGeometryElement {
  return {
    id: element.id,
    x: element.x,
    y: element.y,
    width: element.width,
    height: element.height,
  };
}

function uniquePresentationGeometryElements(
  moving: WorkSlideElement,
  elements: readonly WorkSlideElement[],
): WorkSlideElement[] {
  const unique = new Map(elements.map((element) => [element.id, element]));
  unique.set(moving.id, moving);
  return [...unique.values()];
}
