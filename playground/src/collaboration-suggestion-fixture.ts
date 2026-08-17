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
  readonly nativeProjection: PlaygroundNativeDocumentSuggestionProjection;
  readonly nativeStage: PlaygroundNativeDocumentSuggestionStage;
  readonly advanceNativeSuggestion: () => void;
  readonly updateContent: (content: DocumentContent) => void;
}

export interface PlaygroundNativeDocumentSuggestionProjection {
  readonly editorHasDeletion: boolean;
  readonly editorHasInsertion: boolean;
  readonly suggesterHasDeletion: boolean;
  readonly suggesterHasInsertion: boolean;
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
    const fixture = createDocumentSuggestionFixture('ready');
    setOwned(fixture);
    setContent(fixture.content());
    setNativeStage('ready');
  }, [enabled]);

  useEffect(() => () => owned?.destroy(), [owned]);

  const advanceNativeSuggestion = useCallback(() => {
    const nextStage = nextNativeDocumentSuggestionStage(nativeStage);
    if (nextStage === nativeStage) return;
    const fixture = createDocumentSuggestionFixture(nextStage);
    setOwned(fixture);
    setContent(fixture.content());
    setNativeStage(nextStage);
  }, [nativeStage]);

  return enabled && owned && content
    ? {
        artifactId: PLAYGROUND_DOCUMENT_SUGGESTION_ARTIFACT_ID,
        content,
        editor: owned.editor,
        suggester: owned.suggester,
        nativeProjection: owned.nativeProjection(),
        nativeStage,
        advanceNativeSuggestion,
        updateContent: setContent,
      }
    : undefined;
}

interface OwnedDocumentSuggestionFixture {
  readonly editor: OfficeCollaborationSession;
  readonly suggester: OfficeCollaborationSession;
  content(): DocumentContent;
  nativeProjection(): PlaygroundNativeDocumentSuggestionProjection;
  destroy(): void;
}

const PLAYGROUND_DOCUMENT_SUGGESTION_ARTIFACT_ID = 'fixture-document';
const NATIVE_DOCUMENT_SUGGESTION_DELETION_ID = 'native-playground-deletion';
const NATIVE_DOCUMENT_SUGGESTION_INSERTION_ID = 'native-playground-insertion';

function createDocumentSuggestionFixture(
  stage: PlaygroundNativeDocumentSuggestionStage,
): OwnedDocumentSuggestionFixture {
  const suggestionDocument = new Y.Doc();
  Y.applyUpdate(
    suggestionDocument,
    decodeBase64(browserDocumentFixtureBase64.trim()),
  );
  if (stage !== 'ready') {
    Y.applyUpdate(
      suggestionDocument,
      decodeBase64(NATIVE_DOCUMENT_SUGGESTION_PROPOSAL_BASE64),
      'playground-native-agent',
    );
  }
  if (stage === 'accepted') {
    Y.applyUpdate(
      suggestionDocument,
      decodeBase64(NATIVE_DOCUMENT_SUGGESTION_DECISION_BASE64),
      'playground-native-agent',
    );
  }

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
  const content = readOfficeDocumentCollaboration(editor);
  assertNativeUpdateApplied(content, stage);

  return {
    editor,
    suggester,
    content() {
      return readOfficeDocumentCollaboration(editor);
    },
    nativeProjection() {
      const editorContent = readOfficeDocumentCollaboration(editor);
      const suggesterContent = readOfficeDocumentCollaboration(suggester);
      return {
        editorHasDeletion: hasNativeDocumentSuggestion(
          editorContent,
          NATIVE_DOCUMENT_SUGGESTION_DELETION_ID,
        ),
        editorHasInsertion: hasNativeDocumentSuggestion(
          editorContent,
          NATIVE_DOCUMENT_SUGGESTION_INSERTION_ID,
        ),
        suggesterHasDeletion: hasNativeDocumentSuggestion(
          suggesterContent,
          NATIVE_DOCUMENT_SUGGESTION_DELETION_ID,
        ),
        suggesterHasInsertion: hasNativeDocumentSuggestion(
          suggesterContent,
          NATIVE_DOCUMENT_SUGGESTION_INSERTION_ID,
        ),
      };
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

function nextNativeDocumentSuggestionStage(
  stage: PlaygroundNativeDocumentSuggestionStage,
): PlaygroundNativeDocumentSuggestionStage {
  switch (stage) {
    case 'ready':
      return 'proposed';
    case 'proposed':
      return 'accepted';
    case 'accepted':
      return 'accepted';
  }
}

function hasNativeDocumentSuggestion(
  content: DocumentContent,
  changeId: string,
): boolean {
  return content.html.includes(`data-change-id="${changeId}"`);
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
