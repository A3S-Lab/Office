import {
  createOfficeCollaborationSession,
  type DocumentContent,
  type OfficeCollaborationSession,
  readOfficeDocumentCollaboration,
} from '@a3s-lab/office/core';
import { useCallback, useEffect, useState } from 'react';
import * as Y from 'yjs';
import liveFrames from '../../tests/fixtures/native-document-live-frames.json';

const CHANNEL = 'a3s-office-document-live-caret';

export type PlaygroundDocumentLiveCaretPeer = 'source' | 'observer';

export interface PlaygroundDocumentLiveCaretFixture {
  readonly collaboration: OfficeCollaborationSession;
  readonly content: DocumentContent;
  readonly peer: PlaygroundDocumentLiveCaretPeer;
  readonly applied: number;
  readonly caretIndex: number | null;
  readonly done: boolean;
  advance(): void;
  updateContent(content: DocumentContent): void;
}

export function usePlaygroundDocumentLiveCaretFixture(
  enabled: boolean,
): PlaygroundDocumentLiveCaretFixture | undefined {
  const [owned, setOwned] = useState<OwnedLiveCaretFixture>();
  const [content, setContent] = useState<DocumentContent>();
  const [applied, setApplied] = useState(0);
  const [caretIndex, setCaretIndex] = useState<number | null>(null);

  useEffect(() => {
    if (!enabled) return;
    const peer = readPeer();
    const fixture = createLiveCaretFixture();
    const channel = new BroadcastChannel(CHANNEL);
    const apply = (frame: LiveCaretFrame) => {
      Y.applyUpdate(fixture.document, decodeBase64(frame.updateBase64), {
        caret: frame.caret,
      });
      setApplied((count) => count + 1);
      setCaretIndex(frame.caret.indexUtf16);
    };
    if (peer === 'observer') {
      channel.onmessage = (event: MessageEvent<LiveCaretFrame>) => {
        apply(event.data);
      };
    }
    fixture.advance = () => {
      const frame = liveFrames.frames[fixture.next];
      if (!frame || peer !== 'source') return;
      apply(frame);
      channel.postMessage(frame);
      fixture.next += 1;
    };
    setOwned(fixture);
    setContent(fixture.content);
    return () => {
      channel.close();
      fixture.collaboration.destroy();
      fixture.document.destroy();
    };
  }, [enabled]);

  const advance = useCallback(() => {
    owned?.advance();
  }, [owned]);

  return enabled && owned && content
    ? {
        collaboration: owned.collaboration,
        content,
        peer: owned.peer,
        applied,
        caretIndex,
        done: applied >= liveFrames.frames.length,
        advance,
        updateContent: setContent,
      }
    : undefined;
}

interface OwnedLiveCaretFixture {
  readonly collaboration: OfficeCollaborationSession;
  readonly content: DocumentContent;
  readonly document: Y.Doc;
  readonly peer: PlaygroundDocumentLiveCaretPeer;
  next: number;
  advance(): void;
}

interface LiveCaretFrame {
  readonly updateBase64: string;
  readonly caret: {
    readonly kind: 'document';
    readonly paragraphId: string;
    readonly textId: string;
    readonly indexUtf16: number;
  };
}

function createLiveCaretFixture(): OwnedLiveCaretFixture {
  const document = new Y.Doc();
  Y.applyUpdate(document, decodeBase64(liveFrames.initialUpdateBase64));
  const collaboration = createOfficeCollaborationSession({
    actor: {
      id: 'playground-observer',
      name: '观测端',
      color: '#6d28d9',
      kind: 'agent',
    },
    artifactId: 'fixture-document',
    document,
    kind: 'document',
    mode: 'edit',
  });
  return {
    collaboration,
    content: readOfficeDocumentCollaboration(collaboration),
    document,
    peer: readPeer(),
    next: 0,
    advance() {},
  };
}

function readPeer(): PlaygroundDocumentLiveCaretPeer {
  if (typeof window === 'undefined') return 'source';
  return new URLSearchParams(window.location.search).get('peer') === 'observer'
    ? 'observer'
    : 'source';
}

function decodeBase64(value: string): Uint8Array {
  const decoded = atob(value);
  return Uint8Array.from(decoded, (character) => character.charCodeAt(0));
}
