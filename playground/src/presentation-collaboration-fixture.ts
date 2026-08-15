import {
  createOfficeCollaborationSession,
  type OfficeCollaborationSession,
  readOfficePresentationCollaboration,
} from '@a3s-lab/office/core';
import { useCallback, useEffect, useState } from 'react';
import * as Y from 'yjs';
import browserPresentationFixtureBase64 from '../../tests/fixtures/browser-presentation-collaboration-update.base64';
import {
  NATIVE_PRESENTATION_CREATE_ELEMENT_BASE64,
  NATIVE_PRESENTATION_DELETE_ELEMENT_BASE64,
  NATIVE_PRESENTATION_UPDATE_ELEMENT_BASE64,
} from '../../tests/fixtures/native-presentation-element-updates';

export type PlaygroundPresentationElementStage =
  | 'ready'
  | 'updated'
  | 'created'
  | 'deleted';

export interface PlaygroundPresentationCollaborationFixture {
  readonly collaboration: OfficeCollaborationSession;
  readonly elementStage: PlaygroundPresentationElementStage;
  advanceElement(): void;
}

export function usePlaygroundPresentationCollaborationFixture(
  enabled: boolean,
): PlaygroundPresentationCollaborationFixture | undefined {
  const [ownedFixture, setOwnedFixture] =
    useState<OwnedPlaygroundPresentationCollaborationFixture>();
  const [elementStage, setElementStage] =
    useState<PlaygroundPresentationElementStage>('ready');

  useEffect(() => {
    if (!enabled) return;
    const nextFixture = createPresentationCollaborationFixture();
    setOwnedFixture(nextFixture);
    return () => nextFixture.destroy();
  }, [enabled]);

  const advanceElement = useCallback(() => {
    if (!ownedFixture) return;
    setElementStage(ownedFixture.advanceElement());
  }, [ownedFixture]);

  return ownedFixture
    ? {
        collaboration: ownedFixture.collaboration,
        elementStage,
        advanceElement,
      }
    : undefined;
}

interface OwnedPlaygroundPresentationCollaborationFixture {
  readonly collaboration: OfficeCollaborationSession;
  advanceElement(): PlaygroundPresentationElementStage;
  destroy(): void;
}

function createPresentationCollaborationFixture(): OwnedPlaygroundPresentationCollaborationFixture {
  const document = new Y.Doc();
  Y.applyUpdate(
    document,
    decodeBase64(browserPresentationFixtureBase64.trim()),
  );
  const collaboration = createOfficeCollaborationSession({
    actor: { id: 'playground-user', name: 'Lin Cheng', color: '#c85637' },
    artifactId: 'fixture-presentation',
    document,
    kind: 'presentation',
  });
  const updates = [
    NATIVE_PRESENTATION_UPDATE_ELEMENT_BASE64,
    NATIVE_PRESENTATION_CREATE_ELEMENT_BASE64,
    NATIVE_PRESENTATION_DELETE_ELEMENT_BASE64,
  ];
  const stages: PlaygroundPresentationElementStage[] = [
    'updated',
    'created',
    'deleted',
  ];
  let updateIndex = 0;

  return {
    collaboration,
    advanceElement() {
      const encoded = updates[updateIndex];
      const stage = stages[updateIndex];
      if (!encoded || !stage) return 'deleted';
      Y.applyUpdate(document, decodeBase64(encoded), 'playground-native-agent');
      assertNativeUpdateApplied(collaboration, stage);
      updateIndex += 1;
      return stage;
    },
    destroy() {
      collaboration.destroy();
      document.destroy();
    },
  };
}

function assertNativeUpdateApplied(
  collaboration: OfficeCollaborationSession,
  stage: PlaygroundPresentationElementStage,
): void {
  const content = readOfficePresentationCollaboration(collaboration);
  if (stage === 'updated') {
    const title = content.slides
      .find(({ id }) => id === 'slide-1')
      ?.elements.find(({ id }) => id === 'element-title');
    if (title?.text !== 'Native interop title' || title.x !== 16) {
      throw new Error(
        'The native Presentation scene-element update did not project.',
      );
    }
    return;
  }
  if (stage === 'created') {
    const created = content.slides
      .find(({ id }) => id === 'slide-1')
      ?.elements.find(({ id }) => id === 'element-native');
    if (created?.text !== 'Native interop object') {
      throw new Error(
        'The native Presentation scene-element creation did not project.',
      );
    }
    return;
  }
  if (stage === 'deleted') {
    const deleted = content.slides
      .find(({ id }) => id === 'slide-2')
      ?.elements.some(({ id }) => id === 'element-body');
    if (deleted) {
      throw new Error(
        'The native Presentation scene-element tombstone did not project.',
      );
    }
  }
}

function decodeBase64(value: string): Uint8Array {
  const decoded = atob(value);
  return Uint8Array.from(decoded, (character) => character.charCodeAt(0));
}
