import {
  createOfficeCollaborationSession,
  type DocumentContent,
  type OfficeCollaborationSession,
  readOfficeDocumentCollaboration,
} from '@a3s-lab/office/core';
import { useCallback, useEffect, useState } from 'react';
import * as Y from 'yjs';
import browserDocumentFixtureBase64 from '../../tests/fixtures/browser-document-suggestion-update.base64';
import {
  NATIVE_DOCUMENT_SUGGESTION_DECISION_BASE64,
  NATIVE_DOCUMENT_SUGGESTION_PROPOSAL_BASE64,
} from '../../tests/fixtures/native-document-suggestion-updates';

export type PlaygroundNativeDocumentSuggestionStage =
  | 'ready'
  | 'proposed'
  | 'accepted';

export interface PlaygroundDocumentSuggestionFixture {
  readonly artifactId: string;
  readonly content: DocumentContent;
  readonly editor: OfficeCollaborationSession;
  readonly suggester: OfficeCollaborationSession;
  readonly nativeStage: PlaygroundNativeDocumentSuggestionStage;
  readonly advanceNativeSuggestion: () => void;
  readonly updateContent: (content: DocumentContent) => void;
}

export function usePlaygroundDocumentSuggestionFixture(
  enabled: boolean,
): PlaygroundDocumentSuggestionFixture | undefined {
  const [owned, setOwned] = useState<OwnedDocumentSuggestionFixture>();
  const [content, setContent] = useState<DocumentContent>();
  const [nativeStage, setNativeStage] =
    useState<PlaygroundNativeDocumentSuggestionStage>('ready');

  useEffect(() => {
    if (!enabled) {
      setOwned(undefined);
      setContent(undefined);
      setNativeStage('ready');
      return;
    }
    const fixture = createDocumentSuggestionFixture();
    setOwned(fixture);
    setContent(fixture.content());
    setNativeStage('ready');
    return () => fixture.destroy();
  }, [enabled]);

  const advanceNativeSuggestion = useCallback(() => {
    if (!owned) return;
    const next = owned.advanceNativeSuggestion();
    setNativeStage(next.stage);
    setContent(next.content);
  }, [owned]);

  return enabled && owned && content
    ? {
        artifactId: PLAYGROUND_DOCUMENT_SUGGESTION_ARTIFACT_ID,
        content,
        editor: owned.editor,
        suggester: owned.suggester,
        nativeStage,
        advanceNativeSuggestion,
        updateContent: setContent,
      }
    : undefined;
}

interface OwnedDocumentSuggestionFixture {
  readonly editor: OfficeCollaborationSession;
  readonly suggester: OfficeCollaborationSession;
  advanceNativeSuggestion(): {
    content: DocumentContent;
    stage: PlaygroundNativeDocumentSuggestionStage;
  };
  content(): DocumentContent;
  destroy(): void;
}

const PLAYGROUND_DOCUMENT_SUGGESTION_ARTIFACT_ID = 'fixture-document';

function createDocumentSuggestionFixture(): OwnedDocumentSuggestionFixture {
  const suggestionDocument = new Y.Doc();
  Y.applyUpdate(
    suggestionDocument,
    decodeBase64(browserDocumentFixtureBase64.trim()),
  );

  const editorDocument = new Y.Doc();
  Y.applyUpdate(editorDocument, Y.encodeStateAsUpdate(suggestionDocument));
  const suggestionRelayOrigin = Symbol('playground-suggestion-relay');
  const editorRelayOrigin = Symbol('playground-editor-relay');
  const relaySuggestionUpdate = (update: Uint8Array, origin: unknown) => {
    if (origin !== editorRelayOrigin) {
      Y.applyUpdate(editorDocument, update, suggestionRelayOrigin);
    }
  };
  const relayEditorUpdate = (update: Uint8Array, origin: unknown) => {
    if (origin !== suggestionRelayOrigin) {
      Y.applyUpdate(suggestionDocument, update, editorRelayOrigin);
    }
  };
  suggestionDocument.on('update', relaySuggestionUpdate);
  editorDocument.on('update', relayEditorUpdate);

  const suggester = createOfficeCollaborationSession({
    actor: { id: 'playground-suggester', name: '林澄', color: '#047857' },
    artifactId: PLAYGROUND_DOCUMENT_SUGGESTION_ARTIFACT_ID,
    document: suggestionDocument,
    kind: 'document',
    mode: 'suggest',
  });
  const editor = createOfficeCollaborationSession({
    actor: { id: 'playground-editor', name: '周宁', color: '#6d28d9' },
    artifactId: PLAYGROUND_DOCUMENT_SUGGESTION_ARTIFACT_ID,
    document: editorDocument,
    kind: 'document',
    mode: 'edit',
  });
  const updates = [
    NATIVE_DOCUMENT_SUGGESTION_PROPOSAL_BASE64,
    NATIVE_DOCUMENT_SUGGESTION_DECISION_BASE64,
  ];
  const stages: PlaygroundNativeDocumentSuggestionStage[] = [
    'proposed',
    'accepted',
  ];
  let updateIndex = 0;

  return {
    editor,
    suggester,
    advanceNativeSuggestion() {
      const encoded = updates[updateIndex];
      const stage = stages[updateIndex];
      if (!encoded || !stage) {
        return {
          content: readOfficeDocumentCollaboration(editor),
          stage: 'accepted',
        };
      }
      Y.applyUpdate(
        updateIndex === 0 ? suggestionDocument : editorDocument,
        decodeBase64(encoded),
        'playground-native-agent',
      );
      const content = readOfficeDocumentCollaboration(editor);
      assertNativeUpdateApplied(content, stage);
      updateIndex += 1;
      return { content, stage };
    },
    content() {
      return readOfficeDocumentCollaboration(editor);
    },
    destroy() {
      suggestionDocument.off('update', relaySuggestionUpdate);
      editorDocument.off('update', relayEditorUpdate);
      suggester.destroy();
      editor.destroy();
      suggestionDocument.destroy();
      editorDocument.destroy();
    },
  };
}

function assertNativeUpdateApplied(
  content: DocumentContent,
  stage: PlaygroundNativeDocumentSuggestionStage,
): void {
  if (stage === 'proposed') {
    const suggestionCount = (content.html.match(/data-document-change=/g) ?? [])
      .length;
    if (
      suggestionCount !== 2 ||
      !content.html.includes('data-change-author="A3S Agent"') ||
      content.changeDecisions?.length
    ) {
      throw new Error('The native Document suggestion update did not project.');
    }
    return;
  }
  if (
    stage === 'accepted' &&
    (!content.html.includes('>Hello reviewed world</p>') ||
      content.html.includes('data-document-change=') ||
      content.changeDecisions?.length !== 2)
  ) {
    throw new Error('The native Document suggestion decision did not project.');
  }
}

function decodeBase64(value: string): Uint8Array {
  const decoded = atob(value);
  return Uint8Array.from(decoded, (character) => character.charCodeAt(0));
}
