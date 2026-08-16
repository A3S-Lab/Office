import {
  createOfficeCollaborationSession,
  type DocumentContent,
  initializeOfficeDocumentCollaboration,
  type OfficeCollaborationSession,
} from '@a3s-lab/office/core';
import { useEffect, useRef, useState } from 'react';
import * as Y from 'yjs';

export interface PlaygroundDocumentSuggestionFixture {
  readonly content: DocumentContent;
  readonly editor: OfficeCollaborationSession;
  readonly suggester: OfficeCollaborationSession;
  readonly updateContent: (content: DocumentContent) => void;
}

export function usePlaygroundDocumentSuggestionFixture(
  options:
    | {
        artifactId: string;
        content: DocumentContent;
      }
    | undefined,
): PlaygroundDocumentSuggestionFixture | undefined {
  const initialOptions = useRef(options);
  const [owned, setOwned] = useState<OwnedDocumentSuggestionFixture>();
  const [content, setContent] = useState<DocumentContent | undefined>(
    options?.content,
  );

  useEffect(() => {
    const fixtureOptions = initialOptions.current;
    if (!fixtureOptions) return;
    const fixture = createDocumentSuggestionFixture(fixtureOptions);
    setOwned(fixture);
    return () => fixture.destroy();
  }, []);

  return owned && content
    ? {
        content,
        editor: owned.editor,
        suggester: owned.suggester,
        updateContent: setContent,
      }
    : undefined;
}

interface OwnedDocumentSuggestionFixture {
  readonly editor: OfficeCollaborationSession;
  readonly suggester: OfficeCollaborationSession;
  destroy(): void;
}

function createDocumentSuggestionFixture({
  artifactId,
  content,
}: {
  artifactId: string;
  content: DocumentContent;
}): OwnedDocumentSuggestionFixture {
  const suggestionDocument = new Y.Doc();
  const bootstrap = createOfficeCollaborationSession({
    actor: {
      id: 'playground-bootstrap',
      name: 'Playground bootstrap',
      kind: 'system',
    },
    artifactId,
    document: suggestionDocument,
    kind: 'document',
  });
  initializeOfficeDocumentCollaboration(bootstrap, content);
  bootstrap.destroy();

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
    artifactId,
    document: suggestionDocument,
    kind: 'document',
    mode: 'suggest',
  });
  const editor = createOfficeCollaborationSession({
    actor: { id: 'playground-editor', name: '周宁', color: '#6d28d9' },
    artifactId,
    document: editorDocument,
    kind: 'document',
    mode: 'edit',
  });

  return {
    editor,
    suggester,
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
