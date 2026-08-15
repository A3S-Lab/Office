import {
  createOfficeCollaborationSession,
  createOfficePdfCollaborationBinding,
  createPdfCollaborationContent,
  initializeOfficePdfCollaboration,
  type OfficeCollaborationSession,
  type PdfCollaborationAnnotation,
} from '@a3s-lab/office/core';
import { useCallback, useEffect, useRef, useState } from 'react';
import * as Y from 'yjs';

export type PlaygroundPdfAnnotationStage =
  | 'ready'
  | 'created'
  | 'updated'
  | 'deleted';

export interface PlaygroundPdfCollaborationFixture {
  readonly collaboration: OfficeCollaborationSession;
  readonly annotationStage: PlaygroundPdfAnnotationStage;
  advanceAnnotation(): void;
}

export function usePlaygroundPdfCollaborationFixture(
  options:
    | {
        artifactId: string;
        loadSource: () => Promise<Blob>;
        pageCount: number;
      }
    | undefined,
): PlaygroundPdfCollaborationFixture | undefined {
  const initialOptions = useRef(options);
  const [ownedFixture, setOwnedFixture] =
    useState<OwnedPlaygroundPdfCollaborationFixture>();
  const [annotationStage, setAnnotationStage] =
    useState<PlaygroundPdfAnnotationStage>('ready');

  useEffect(() => {
    const fixtureOptions = initialOptions.current;
    if (!fixtureOptions) return;
    let active = true;
    let current: OwnedPlaygroundPdfCollaborationFixture | undefined;
    void createPdfCollaborationFixture(fixtureOptions).then(
      (nextFixture) => {
        if (!active) {
          nextFixture.destroy();
          return;
        }
        current = nextFixture;
        setOwnedFixture(nextFixture);
      },
      (error: unknown) => {
        if (!active) return;
        console.error(
          'Unable to prepare the PDF collaboration fixture.',
          error,
        );
      },
    );
    return () => {
      active = false;
      current?.destroy();
    };
  }, []);

  const advanceAnnotation = useCallback(() => {
    if (!ownedFixture) return;
    setAnnotationStage(ownedFixture.advanceAnnotation());
  }, [ownedFixture]);

  return ownedFixture
    ? {
        collaboration: ownedFixture.collaboration,
        annotationStage,
        advanceAnnotation,
      }
    : undefined;
}

interface OwnedPlaygroundPdfCollaborationFixture {
  readonly collaboration: OfficeCollaborationSession;
  advanceAnnotation(): PlaygroundPdfAnnotationStage;
  destroy(): void;
}

async function createPdfCollaborationFixture({
  artifactId,
  loadSource,
  pageCount,
}: {
  artifactId: string;
  loadSource: () => Promise<Blob>;
  pageCount: number;
}): Promise<OwnedPlaygroundPdfCollaborationFixture> {
  const source = await loadSource();
  const subtle = globalThis.crypto?.subtle;
  if (!subtle) {
    throw new Error('SHA-256 PDF source verification is unavailable.');
  }
  const digest = new Uint8Array(
    await subtle.digest('SHA-256', await source.arrayBuffer()),
  );
  const document = new Y.Doc();
  const collaboration = createOfficeCollaborationSession({
    actor: { id: 'playground-user', name: 'Lin Cheng', color: '#047857' },
    artifactId,
    document,
    kind: 'pdf',
  });
  const agentSession = createOfficeCollaborationSession({
    actor: {
      id: 'playground-agent',
      name: 'A3S Agent',
      color: '#6d28d9',
      kind: 'agent',
    },
    artifactId,
    document,
    kind: 'pdf',
  });
  initializeOfficePdfCollaboration(
    collaboration,
    createPdfCollaborationContent({
      byteLength: source.size,
      pageCount,
      sha256: Array.from(digest, (value) =>
        value.toString(16).padStart(2, '0'),
      ).join(''),
    }),
  );
  const agentBinding = createOfficePdfCollaborationBinding(agentSession, {
    origin: agentSession.createOrigin('agent', 'playground-native-annotation'),
  });
  let stage: PlaygroundPdfAnnotationStage = 'ready';

  return {
    collaboration,
    advanceAnnotation() {
      const previous = agentBinding.content();
      const annotation = previous.annotations.find(
        ({ id }) => id === PLAYGROUND_NATIVE_ANNOTATION_ID,
      );
      if (stage === 'ready') {
        agentBinding.replace(previous, {
          ...previous,
          annotations: [
            ...previous.annotations,
            playgroundNativeAnnotation('#f59e0b', 'Created by A3S Agent'),
          ],
        });
        stage = 'created';
      } else if (stage === 'created' && annotation) {
        agentBinding.replace(previous, {
          ...previous,
          annotations: previous.annotations.map((record) =>
            record.id === PLAYGROUND_NATIVE_ANNOTATION_ID
              ? {
                  ...record,
                  annotation: {
                    ...record.annotation,
                    color: '#dc2626',
                    contents: 'Updated concurrently by A3S Agent',
                    strokeColor: '#dc2626',
                  },
                }
              : record,
          ),
        });
        stage = 'updated';
      } else if (stage === 'updated' && annotation) {
        agentBinding.replace(previous, {
          ...previous,
          annotations: previous.annotations.map((record) =>
            record.id === PLAYGROUND_NATIVE_ANNOTATION_ID
              ? { ...record, deleted: true }
              : record,
          ),
        });
        stage = 'deleted';
      }
      return stage;
    },
    destroy() {
      agentBinding.destroy();
      collaboration.destroy();
      agentSession.destroy();
      document.destroy();
    },
  };
}

const PLAYGROUND_NATIVE_ANNOTATION_ID = 'playground-native-highlight';

function playgroundNativeAnnotation(
  color: string,
  contents: string,
): PdfCollaborationAnnotation {
  const rect = {
    origin: { x: 68, y: 78 },
    size: { width: 300, height: 28 },
  };
  return {
    id: PLAYGROUND_NATIVE_ANNOTATION_ID,
    pageIndex: 0,
    source: 'created',
    annotation: {
      id: PLAYGROUND_NATIVE_ANNOTATION_ID,
      pageIndex: 0,
      type: 9,
      rect,
      segmentRects: [rect],
      color,
      strokeColor: color,
      opacity: 0.48,
      contents,
      author: 'A3S Agent',
      created: '2026-08-15T08:00:00.000Z',
    },
  };
}
