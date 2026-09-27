import { expect, test } from '@rstest/core';
import * as Y from 'yjs';
import {
  createOfficeCollaborationSession,
  createOfficeDocumentCollaborationBinding,
  readOfficeDocumentCollaboration,
} from '../src/core';
import liveFrames from './fixtures/native-document-live-frames.json';

test('two browser sessions follow native CJK grapheme frames and their carets', async () => {
  const firstDocument = new Y.Doc();
  const secondDocument = new Y.Doc();
  const initial = decodeBase64(liveFrames.initialUpdateBase64);
  Y.applyUpdate(firstDocument, initial);
  Y.applyUpdate(secondDocument, initial);
  const first = createOfficeCollaborationSession({
    artifactId: 'fixture-document',
    document: firstDocument,
    kind: 'document',
  });
  const second = createOfficeCollaborationSession({
    artifactId: 'fixture-document',
    document: secondDocument,
    kind: 'document',
  });
  const firstBinding = createOfficeDocumentCollaborationBinding(first);
  const secondBinding = createOfficeDocumentCollaborationBinding(second);
  const carets: unknown[] = [];
  secondBinding.subscribe((change) => {
    if (change.caret) carets.push(change.caret);
  });

  for (const frame of liveFrames.frames) {
    const update = decodeBase64(frame.updateBase64);
    const origin = { caret: frame.caret };
    Y.applyUpdate(firstDocument, update, origin);
    Y.applyUpdate(secondDocument, update, origin);
    await Promise.resolve();
  }

  const firstText = readOfficeDocumentCollaboration(first).html;
  const secondText = readOfficeDocumentCollaboration(second).html;
  expect(firstText).toContain('中文');
  expect(secondText).toBe(firstText);
  expect(carets).toEqual(liveFrames.frames.map((frame) => frame.caret));
  expect(secondText).not.toContain('indexUtf16');

  firstBinding.destroy();
  secondBinding.destroy();
  first.destroy();
  second.destroy();
  firstDocument.destroy();
  secondDocument.destroy();
});

function decodeBase64(value: string): Uint8Array {
  const decoded = atob(value);
  return Uint8Array.from(decoded, (character) => character.charCodeAt(0));
}
