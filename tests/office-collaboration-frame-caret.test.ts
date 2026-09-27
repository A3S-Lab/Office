import { Schema } from '@tiptap/pm/model';
import { expect, test } from '@rstest/core';
import * as Y from 'yjs';
import {
  createOfficeCollaborationSession,
  createOfficeCollaborationTransportBinding,
  createOfficeMarkdownCollaborationBinding,
  initializeOfficeMarkdownCollaboration,
  type OfficeCollaborationTransport,
  type OfficeCollaborationTransportMessage,
} from '../src/core';
import {
  documentFrameCaretPosition,
  markdownSourceCaret,
  paintFrameCaretAfterLayout,
  parseWorkOfficeCollaborationFrameCaret,
} from '../src/internal/collaboration/office-collaboration-frame-caret';

const schema = new Schema({
  nodes: {
    doc: { content: 'block+' },
    paragraph: {
      content: 'inline*',
      group: 'block',
      attrs: {
        paragraphId: { default: null },
        textId: { default: null },
      },
    },
    text: { group: 'inline' },
  },
});

test('places a document caret inside the addressed paragraph without adding block newlines', () => {
  const document = schema.node('doc', null, [
    schema.node(
      'paragraph',
      { paragraphId: '00000001', textId: '00000002' },
      [schema.text('A😀B')],
    ),
    schema.node(
      'paragraph',
      { paragraphId: '00000003', textId: '00000004' },
      [schema.text('next')],
    ),
  ]);
  expect(
    documentFrameCaretPosition(document, {
      paragraphId: '00000001',
      textId: '00000002',
      indexUtf16: 1,
    }),
  ).toBe(2);
  expect(
    documentFrameCaretPosition(document, {
      paragraphId: '00000001',
      textId: '00000002',
      indexUtf16: 2,
    }),
  ).toBeNull();
  expect(
    documentFrameCaretPosition(document, {
      paragraphId: '00000001',
      textId: '00000099',
      indexUtf16: 0,
    }),
  ).toBeNull();
});

test('paints a wrapped caret on the line that starts at the insertion edge after layout', () => {
  const lines = [
    { from: 1, to: 7, left: 10, top: 20, right: 110, bottom: 36 },
    { from: 7, to: 9, left: 10, top: 40, right: 70, bottom: 56 },
  ];
  expect(
    paintFrameCaretAfterLayout({
      documentSize: 10,
      followUpHead: 10,
      head: 7,
      layoutSettled: true,
      lines,
    }),
  ).toEqual({ head: 7, left: 10, top: 40, width: 2, height: 16 });
  const zoomed = lines.map((line) => ({
    ...line,
    left: line.left * 1.5,
    right: line.right * 1.5,
    top: line.top * 1.5,
    bottom: line.bottom * 1.5,
  }));
  const painted = paintFrameCaretAfterLayout({
    documentSize: 10,
    head: 7,
    layoutSettled: true,
    lines: zoomed,
  });
  expect(painted?.head).toBe(7);
  expect(painted?.left).toBe(15);
  expect(
    paintFrameCaretAfterLayout({
      documentSize: 10,
      followUpHead: 10,
      head: 7,
      layoutSettled: false,
      lines,
    }),
  ).toBeNull();
  expect(
    paintFrameCaretAfterLayout({
      documentSize: 4,
      head: 9,
      layoutSettled: true,
      lines,
    }),
  ).toBeNull();
});

test('keeps a markdown source caret on the UTF-16 index and rejects a split surrogate', () => {
  expect(markdownSourceCaret('A\nB', 2)).toBe(2);
  expect(markdownSourceCaret('A😀B', 2)).toBeNull();
  expect(() =>
    parseWorkOfficeCollaborationFrameCaret({
      kind: 'pdf-field',
      fieldId: 'name',
      indexUtf16: 3,
    }),
  ).toThrow(/no text offset/);
});

test('delivers a frame caret with the update and rejects it on a sync handshake', () => {
  const transport = new MemoryTransport();
  const firstDocument = new Y.Doc();
  const first = createOfficeCollaborationSession({
    artifactId: 'frame-caret',
    document: firstDocument,
    kind: 'markdown',
  });
  initializeOfficeMarkdownCollaboration(first, {
    type: 'markdown',
    markdown: 'Before',
  });
  const second = createOfficeCollaborationSession({
    artifactId: 'frame-caret',
    document: cloneDocument(firstDocument),
    kind: 'markdown',
  });
  const firstTransport = createOfficeCollaborationTransportBinding(
    first,
    transport,
    { autoSynchronize: false },
  );
  const secondTransport = createOfficeCollaborationTransportBinding(
    second,
    transport,
    { autoSynchronize: false },
  );
  const firstContent = createOfficeMarkdownCollaborationBinding(first);
  const secondContent = createOfficeMarkdownCollaborationBinding(second);
  const received: unknown[] = [];
  const changes: unknown[] = [];
  secondTransport.subscribeFrameCaret((caret) => received.push(caret));
  secondContent.subscribe((change) => changes.push(change));

  transport.deliver = false;
  firstContent.replace('After');
  const update = transport.messages.at(-1);
  expect(update?.type).toBe('update');
  transport.deliver = true;
  const caret = { kind: 'markdown' as const, indexUtf16: 5 };
  transport.publish({ ...update!, caret });

  expect(received).toEqual([caret]);
  expect(changes.at(-1)).toMatchObject({ caret });
  const { origin: _origin, ...handshake } = update!;
  expect(() =>
    transport.publish({
      ...handshake,
      type: 'sync-step-1',
      caret,
      payload: Y.encodeStateVector(firstDocument),
    }),
  ).toThrow(/frame caret/);

  firstContent.destroy();
  secondContent.destroy();
  firstTransport.destroy();
  secondTransport.destroy();
});

class MemoryTransport implements OfficeCollaborationTransport {
  readonly messages: OfficeCollaborationTransportMessage[] = [];
  readonly #listeners = new Set<(message: unknown) => void>();
  deliver = true;

  publish(message: OfficeCollaborationTransportMessage): void {
    this.messages.push(message);
    if (!this.deliver) return;
    for (const listener of [...this.#listeners]) listener(message);
  }

  subscribe(listener: (message: unknown) => void): () => void {
    this.#listeners.add(listener);
    return () => this.#listeners.delete(listener);
  }
}

function cloneDocument(source: Y.Doc): Y.Doc {
  const clone = new Y.Doc();
  Y.applyUpdate(clone, Y.encodeStateAsUpdate(source));
  return clone;
}
