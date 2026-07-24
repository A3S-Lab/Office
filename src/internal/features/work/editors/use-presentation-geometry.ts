import { useCallback, useEffect, useRef, useState } from 'react';
import {
  createOfficeKernelClient,
  type OfficeKernelClient,
} from '../../../kernel/office-kernel-client';
import type {
  OfficeKernelEngine,
  OfficeKernelPresentationAlignment,
  OfficeKernelPresentationGeometryElement,
} from '../../../kernel/office-kernel-protocol';
import type { WorkSlideElement } from '../work-types';

export interface PresentationGeometryController {
  alignElement: (
    element: WorkSlideElement,
    alignment: OfficeKernelPresentationAlignment,
  ) => Promise<OfficeKernelPresentationGeometryElement | null>;
  engine: OfficeKernelEngine | null;
  pending: boolean;
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

  const alignElement = useCallback(
    async (
      element: WorkSlideElement,
      alignment: OfficeKernelPresentationAlignment,
    ): Promise<OfficeKernelPresentationGeometryElement | null> => {
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
            operation: { type: 'alignToSlide', alignment },
            elements: [
              {
                id: element.id,
                x: element.x,
                y: element.y,
                width: element.width,
                height: element.height,
              },
            ],
          },
          controller.signal,
        );
        if (controller.signal.aborted || requestRevision !== revision.current) {
          return null;
        }
        setEngine(result.engine);
        return (
          result.elements.find((candidate) => candidate.id === element.id) ??
          null
        );
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

  return { alignElement, engine, pending };
}
