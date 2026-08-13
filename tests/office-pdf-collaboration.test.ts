import { expect, test } from '@rstest/core';
import * as Y from 'yjs';
import {
  assertPdfCollaborationSource,
  createOfficeCollaborationSession,
  createOfficePdfCollaborationBinding,
  initializeOfficePdfCollaboration,
  readOfficePdfCollaboration,
  type PdfCollaborationContent,
} from '../src/core';
import {
  PDF_COLLABORATION_SOURCE,
  pdfCollaborationFixture as fixture,
} from './fixtures/pdf-collaboration';

test('initializes typed PDF roots without storing source bytes', () => {
  const session = pdfSession('pdf-typed');
  const expected = fixture();

  expect(initializeOfficePdfCollaboration(session, expected)).toEqual({
    initialized: true,
    content: expected,
  });
  expect(readOfficePdfCollaboration(session)).toEqual(expected);
  expect(session.document.share.has(session.rootName('pdf.content'))).toBe(
    false,
  );
  expect(session.document.share.has(session.rootName('pdf.bytes'))).toBe(false);
  expect(
    session.document.getMap(session.rootName('pdf.source')).toJSON(),
  ).toEqual(PDF_COLLABORATION_SOURCE);
});

test('rejects another byte source even when the page count matches', () => {
  const session = pdfSession('pdf-source-mismatch');
  initializeOfficePdfCollaboration(session, fixture());
  const anotherSource = {
    ...PDF_COLLABORATION_SOURCE,
    sha256: 'f'.repeat(64),
  };

  expect(() => assertPdfCollaborationSource(session, anotherSource)).toThrow(
    /does not match collaboration artifact/,
  );
  expect(() =>
    initializeOfficePdfCollaboration(session, {
      ...fixture(),
      source: anotherSource,
    }),
  ).toThrow(/does not match collaboration artifact/);
});

test('validates references and destructive review records before bootstrap', () => {
  const session = pdfSession('pdf-validation');
  const content = fixture();
  expect(() =>
    initializeOfficePdfCollaboration(session, {
      ...content,
      reviewDecisions: [
        {
          id: 'decision-missing',
          targetKind: 'redaction',
          targetId: 'missing',
          decision: 'approve',
          actorId: 'reviewer-1',
          createdAt: '2026-08-14T00:00:00.000Z',
        },
      ],
    }),
  ).toThrow(/reference an existing redaction/);
  expect(session.document.getMap(session.rootName('metadata')).size).toBe(0);
});

test('merges different fields of one base annotation and independent form values', () => {
  const { first, firstDocument, second, secondDocument } = connectedPair(
    'pdf-field-convergence',
  );
  const firstBinding = createOfficePdfCollaborationBinding(first);
  const secondBinding = createOfficePdfCollaborationBinding(second);
  const firstBefore = firstBinding.content();
  const secondBefore = secondBinding.content();

  firstBinding.replace(
    firstBefore,
    updateAnnotation(firstBefore, 'annotation-base-1', (annotation) => ({
      ...annotation,
      contents: 'First client note',
    })),
  );
  secondBinding.replace(secondBefore, {
    ...updateAnnotation(secondBefore, 'annotation-base-1', (annotation) => ({
      ...annotation,
      color: '#ff0000',
    })),
    formValues: [
      ...secondBefore.formValues,
      { id: 'Applicant.Email', value: 'ada@example.test' },
    ],
  });
  exchangeUpdates(firstDocument, secondDocument);

  const converged = firstBinding.content();
  expect(secondBinding.content()).toEqual(converged);
  expect(converged.annotations[0].annotation).toMatchObject({
    contents: 'First client note',
    color: '#ff0000',
  });
  expect(converged.formValues).toContainEqual({
    id: 'Applicant.Email',
    value: 'ada@example.test',
  });
});

test('converges a deletion tombstone with an independent annotation edit', () => {
  const { first, firstDocument, second, secondDocument } = connectedPair(
    'pdf-tombstone-convergence',
  );
  const firstBinding = createOfficePdfCollaborationBinding(first);
  const secondBinding = createOfficePdfCollaborationBinding(second);
  const firstBefore = firstBinding.content();
  const secondBefore = secondBinding.content();

  firstBinding.replace(
    firstBefore,
    updateAnnotationRecord(firstBefore, 'annotation-base-1', (record) => ({
      ...record,
      deleted: true,
    })),
  );
  secondBinding.replace(
    secondBefore,
    updateAnnotation(secondBefore, 'annotation-base-1', (annotation) => ({
      ...annotation,
      opacity: 0.8,
    })),
  );
  exchangeUpdates(firstDocument, secondDocument);

  const converged = firstBinding.content();
  expect(secondBinding.content()).toEqual(converged);
  expect(converged.annotations[0]).toMatchObject({
    deleted: true,
    annotation: { opacity: 0.8 },
  });
  expect(() =>
    firstBinding.replace(converged, {
      ...converged,
      annotations: converged.annotations.map((annotation) => {
        const { deleted: _deleted, ...rest } = annotation;
        return rest;
      }),
    }),
  ).toThrow(/tombstone/);
});

test('deduplicates identical offline creations and rejects same-ID reuse', () => {
  const identical = connectedPair('pdf-identical-creation');
  const identicalFirst = createOfficePdfCollaborationBinding(identical.first);
  const identicalSecond = createOfficePdfCollaborationBinding(identical.second);
  const annotation = createdAnnotation('annotation-created-1', '#00ff00');
  const firstBefore = identicalFirst.content();
  const secondBefore = identicalSecond.content();
  identicalFirst.replace(firstBefore, {
    ...firstBefore,
    annotations: [...firstBefore.annotations, annotation],
  });
  identicalSecond.replace(secondBefore, {
    ...secondBefore,
    annotations: [...secondBefore.annotations, annotation],
  });
  exchangeUpdates(identical.firstDocument, identical.secondDocument);
  expect(
    identicalFirst
      .content()
      .annotations.filter(({ id }) => id === annotation.id),
  ).toHaveLength(1);

  const collision = connectedPair('pdf-colliding-creation');
  const collisionFirst = createOfficePdfCollaborationBinding(collision.first);
  const collisionSecond = createOfficePdfCollaborationBinding(collision.second);
  const collisionFirstBefore = collisionFirst.content();
  const collisionSecondBefore = collisionSecond.content();
  collisionFirst.replace(collisionFirstBefore, {
    ...collisionFirstBefore,
    annotations: [
      ...collisionFirstBefore.annotations,
      createdAnnotation('annotation-collision', '#ff0000'),
    ],
  });
  collisionSecond.replace(collisionSecondBefore, {
    ...collisionSecondBefore,
    annotations: [
      ...collisionSecondBefore.annotations,
      createdAnnotation('annotation-collision', '#0000ff'),
    ],
  });
  exchangeUpdates(collision.firstDocument, collision.secondDocument);
  expect(() => collisionFirst.content()).toThrow(/concurrently assigned/);
  expect(() => collisionSecond.content()).toThrow(/concurrently assigned/);
});

test('keeps append-only redaction and page-operation review records attributable', () => {
  const session = pdfSession('pdf-review-audit');
  initializeOfficePdfCollaboration(session, fixture());
  const binding = createOfficePdfCollaborationBinding(session);
  const before = binding.content();
  const proposed = {
    ...before,
    redactionProposals: [
      {
        id: 'redaction-1',
        pageIndex: 1,
        rects: [{ left: 10, top: 10, right: 80, bottom: 30 }],
        proposedBy: 'agent-7',
        proposedAt: '2026-08-14T01:00:00.000Z',
        reason: 'Personal data',
      },
    ],
    pageOperations: [
      {
        id: 'page-op-1',
        kind: 'rotate' as const,
        pageIndices: [2],
        degrees: 90 as const,
        proposedBy: 'agent-7',
        proposedAt: '2026-08-14T01:01:00.000Z',
      },
    ],
  };
  binding.replace(before, proposed);
  const current = binding.content();
  binding.replace(current, {
    ...current,
    reviewDecisions: [
      {
        id: 'decision-1',
        targetKind: 'redaction',
        targetId: 'redaction-1',
        decision: 'approve',
        actorId: 'reviewer-2',
        createdAt: '2026-08-14T02:00:00.000Z',
      },
    ],
  });
  expect(binding.content()).toMatchObject({
    redactionProposals: [{ id: 'redaction-1', proposedBy: 'agent-7' }],
    pageOperations: [{ id: 'page-op-1', proposedBy: 'agent-7' }],
    reviewDecisions: [{ id: 'decision-1', actorId: 'reviewer-2' }],
  });
  expect(binding.canUndo()).toBe(false);
  expect(() =>
    binding.replace(binding.content(), {
      ...binding.content(),
      redactionProposals: [
        ...binding.content().redactionProposals,
        {
          id: 'redaction-inserted',
          pageIndex: 0,
          rects: [{ left: 1, top: 1, right: 2, bottom: 2 }],
          proposedBy: 'agent-8',
          proposedAt: '2026-08-14T02:01:00.000Z',
        },
      ].reverse(),
    }),
  ).toThrow(/cannot be reordered/);
  expect(() =>
    binding.replace(binding.content(), {
      ...binding.content(),
      redactionProposals: binding
        .content()
        .redactionProposals.map((proposal) => ({
          ...proposal,
          reason: 'Changed after review',
        })),
    }),
  ).toThrow(/append-only audit record/);
});

test('rejects stale conflicts before writing any part of a PDF transaction', () => {
  const { first, firstDocument, second, secondDocument } =
    connectedPair('pdf-stale-atomic');
  const firstBinding = createOfficePdfCollaborationBinding(first);
  const secondBinding = createOfficePdfCollaborationBinding(second);
  const stale = secondBinding.content();
  const firstBefore = firstBinding.content();
  firstBinding.replace(
    firstBefore,
    updateAnnotation(firstBefore, 'annotation-base-1', (annotation) => ({
      ...annotation,
      contents: 'Remote content',
    })),
  );
  exchangeUpdates(firstDocument, secondDocument);

  expect(() =>
    secondBinding.replace(stale, {
      ...updateAnnotation(stale, 'annotation-base-1', (annotation) => ({
        ...annotation,
        contents: 'Stale local content',
      })),
      formValues: [{ id: 'Applicant.Name', value: 'Grace' }],
    }),
  ).toThrow(/changed concurrently/);
  expect(secondBinding.content().formValues).toEqual(stale.formValues);
});

test('detects concurrent independent PDF bootstrap', () => {
  const firstDocument = new Y.Doc();
  const secondDocument = new Y.Doc();
  const first = pdfSession('pdf-bootstrap-race', firstDocument);
  const second = pdfSession('pdf-bootstrap-race', secondDocument);
  initializeOfficePdfCollaboration(first, fixture());
  initializeOfficePdfCollaboration(second, fixture());
  exchangeUpdates(firstDocument, secondDocument);
  expect(() => readOfficePdfCollaboration(first)).toThrow(
    /Multiple clients initialized/,
  );
  expect(() => readOfficePdfCollaboration(second)).toThrow(
    /Multiple clients initialized/,
  );
});

test('keeps PDF undo local and does not undo remote changes', () => {
  const { first, firstDocument, second, secondDocument } =
    connectedPair('pdf-local-undo');
  const firstBinding = createOfficePdfCollaborationBinding(first, {
    captureTimeoutMs: 0,
  });
  const secondBinding = createOfficePdfCollaborationBinding(second, {
    captureTimeoutMs: 0,
  });
  const firstBefore = firstBinding.content();
  firstBinding.replace(firstBefore, {
    ...firstBefore,
    formValues: [{ id: 'Applicant.Name', value: 'Grace' }],
  });
  exchangeUpdates(firstDocument, secondDocument);
  const secondBefore = secondBinding.content();
  secondBinding.replace(
    secondBefore,
    updateAnnotation(secondBefore, 'annotation-base-1', (annotation) => ({
      ...annotation,
      color: '#00ff00',
    })),
  );
  exchangeUpdates(firstDocument, secondDocument);

  expect(firstBinding.undo()).toBe(true);
  exchangeUpdates(firstDocument, secondDocument);
  const converged = firstBinding.content();
  expect(secondBinding.content()).toEqual(converged);
  expect(converged.formValues).toEqual([
    { id: 'Applicant.Name', value: 'Ada' },
  ]);
  expect(converged.annotations[0].annotation.color).toBe('#00ff00');
});

test('rejects PDF mutation outside edit mode', () => {
  const document = new Y.Doc();
  const writable = pdfSession('pdf-view', document);
  initializeOfficePdfCollaboration(writable, fixture());
  const readOnly = createOfficeCollaborationSession({
    artifactId: 'pdf-view',
    document,
    kind: 'pdf',
    mode: 'view',
  });
  const binding = createOfficePdfCollaborationBinding(readOnly);
  const before = binding.content();
  expect(() =>
    binding.replace(before, {
      ...before,
      formValues: [{ id: 'Applicant.Name', value: 'Grace' }],
    }),
  ).toThrow(/cannot modify canonical content/);
  expect(() => binding.undo()).toThrow(/cannot modify canonical content/);
});

function connectedPair(artifactId: string) {
  const firstDocument = new Y.Doc();
  const first = pdfSession(artifactId, firstDocument);
  initializeOfficePdfCollaboration(first, fixture());
  const secondDocument = cloneDocument(firstDocument);
  return {
    first,
    firstDocument,
    second: pdfSession(artifactId, secondDocument),
    secondDocument,
  };
}

function pdfSession(artifactId: string, document = new Y.Doc()) {
  return createOfficeCollaborationSession({
    artifactId,
    document,
    kind: 'pdf',
  });
}

function updateAnnotation(
  content: PdfCollaborationContent,
  id: string,
  update: (annotation: Record<string, unknown>) => Record<string, unknown>,
): PdfCollaborationContent {
  return updateAnnotationRecord(content, id, (record) => ({
    ...record,
    annotation: update(record.annotation),
  }));
}

function updateAnnotationRecord(
  content: PdfCollaborationContent,
  id: string,
  update: (
    annotation: PdfCollaborationContent['annotations'][number],
  ) => PdfCollaborationContent['annotations'][number],
): PdfCollaborationContent {
  return {
    ...content,
    annotations: content.annotations.map((annotation) =>
      annotation.id === id ? update(annotation) : annotation,
    ),
  };
}

function createdAnnotation(id: string, color: string) {
  return {
    id,
    pageIndex: 1,
    source: 'created' as const,
    annotation: {
      id,
      pageIndex: 1,
      type: 9,
      rect: { left: 20, top: 40, right: 160, bottom: 58 },
      color,
      opacity: 0.5,
    },
  };
}

function cloneDocument(source: Y.Doc): Y.Doc {
  const clone = new Y.Doc();
  Y.applyUpdate(clone, Y.encodeStateAsUpdate(source));
  return clone;
}

function exchangeUpdates(first: Y.Doc, second: Y.Doc): void {
  const firstUpdate = Y.encodeStateAsUpdate(first, Y.encodeStateVector(second));
  const secondUpdate = Y.encodeStateAsUpdate(
    second,
    Y.encodeStateVector(first),
  );
  Y.applyUpdate(first, secondUpdate, 'test-network');
  Y.applyUpdate(second, firstUpdate, 'test-network');
}
