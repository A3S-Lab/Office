import { expect, test } from '@rstest/core';
import * as Y from 'yjs';
import {
  createOfficeCollaborationSession,
  readOfficeDocumentCollaboration,
} from '../src/core';
import browserDocumentFixtureBase64 from './fixtures/browser-document-suggestion-update.base64';
import browserDocumentFormattingFixtureBase64 from './fixtures/browser-document-formatting-change-update.base64';
import {
  NATIVE_DOCUMENT_SUGGESTION_DECISION_BASE64,
  NATIVE_DOCUMENT_SUGGESTION_PROPOSAL_BASE64,
} from './fixtures/native-document-suggestion-updates';

test('projects native Document suggestion creation and decisions in browser Yjs', () => {
  const document = new Y.Doc();
  const session = createOfficeCollaborationSession({
    artifactId: 'fixture-document',
    document,
    kind: 'document',
    mode: 'edit',
  });
  Y.applyUpdate(document, decodeBase64(browserDocumentFixtureBase64.trim()));

  expect(readOfficeDocumentCollaboration(session).html).toContain(
    '>Hello 😀 world</p>',
  );

  Y.applyUpdate(
    document,
    decodeBase64(NATIVE_DOCUMENT_SUGGESTION_PROPOSAL_BASE64),
  );
  const proposed = readOfficeDocumentCollaboration(session);
  expect(proposed.html.match(/data-document-change=/g)).toHaveLength(2);
  expect(proposed.html.match(/data-change-author="A3S Agent"/g)).toHaveLength(
    2,
  );
  expect(proposed.html).toContain('data-change-kind="deletion"');
  expect(proposed.html).toContain('data-change-kind="insertion"');
  expect(proposed.changeDecisions).toBeUndefined();

  Y.applyUpdate(
    document,
    decodeBase64(NATIVE_DOCUMENT_SUGGESTION_DECISION_BASE64),
  );
  const accepted = readOfficeDocumentCollaboration(session);
  expect(accepted.html).toContain('>Hello reviewed world</p>');
  expect(accepted.html).not.toContain('data-document-change=');
  expect(accepted.changeDecisions).toEqual([
    expect.objectContaining({
      changeId: 'native-playground-deletion',
      changeKind: 'deletion',
      decision: 'accept',
      decidedBy: 'Native Editor',
      decidedByActorId: 'native-editor',
      suggestedBy: 'A3S Agent',
      suggestedByActorId: 'native-agent',
      text: '😀',
    }),
    expect.objectContaining({
      changeId: 'native-playground-insertion',
      changeKind: 'insertion',
      decision: 'accept',
      decidedBy: 'Native Editor',
      decidedByActorId: 'native-editor',
      suggestedBy: 'A3S Agent',
      suggestedByActorId: 'native-agent',
      text: 'reviewed',
    }),
  ]);

  session.destroy();
  document.destroy();
});

test('projects browser formatting revisions and decisions from the shared Yjs schema', () => {
  const document = new Y.Doc();
  const session = createOfficeCollaborationSession({
    artifactId: 'fixture-document-formatting',
    document,
    kind: 'document',
    mode: 'edit',
  });
  Y.applyUpdate(
    document,
    decodeBase64(browserDocumentFormattingFixtureBase64.trim()),
  );

  const content = readOfficeDocumentCollaboration(session);
  expect(content.html).toContain('data-change-kind="formatting"');
  expect(content.html).toContain('data-change-id="browser-formatting-live"');
  expect(content.html).toContain('<strong>Format</strong>');
  expect(content.changeDecisions).toEqual([
    expect.objectContaining({
      changeId: 'browser-formatting-decided',
      changeKind: 'formatting',
      decision: 'accept',
      decidedBy: 'Browser Editor',
      suggestedBy: 'Browser Reviewer',
      text: 'Reviewed',
    }),
  ]);

  session.destroy();
  document.destroy();
});

function decodeBase64(value: string): Uint8Array {
  const decoded = atob(value);
  return Uint8Array.from(decoded, (character) => character.charCodeAt(0));
}
