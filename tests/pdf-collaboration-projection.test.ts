import type {
  AnnotationEvent,
  PdfAnnotationObject,
  PluginRegistry,
} from '@embedpdf/react-pdf-viewer';
import { expect, test } from '@rstest/core';
import * as Y from 'yjs';
import {
  createOfficeCollaborationSession,
  createOfficePdfCollaborationBinding,
  createPdfCollaborationContent,
  initializeOfficePdfCollaboration,
  type OfficeCollaborationSession,
  type PdfCollaborationContent,
  readOfficePdfCollaboration,
} from '../src/core';
import { createWorkPdfCollaborationProjection } from '../src/internal/features/work/editors/pdf-collaboration-projection';
import {
  BROWSER_PDF_FIXTURE_BASE64,
  NATIVE_PDF_CREATE_ANNOTATION_BASE64,
  NATIVE_PDF_DELETE_ANNOTATION_BASE64,
  NATIVE_PDF_UPDATE_ANNOTATION_BASE64,
} from './fixtures/native-pdf-annotation-updates';
import { PDF_COLLABORATION_SOURCE } from './fixtures/pdf-collaboration';

const DOCUMENT_ID = 'pdf-document';
const RENDERER_DEFAULT_AUTHOR = 'A3S Office User';

test('projects exact native Rust Highlight updates into the EmbedPDF annotation scope', async () => {
  const document = new Y.Doc();
  Y.applyUpdate(document, decodeBase64(BROWSER_PDF_FIXTURE_BASE64));
  Y.applyUpdate(document, decodeBase64(NATIVE_PDF_CREATE_ANNOTATION_BASE64));
  const session = createOfficeCollaborationSession({
    artifactId: 'fixture-pdf',
    document,
    kind: 'pdf',
  });
  const viewer = createPdfViewerHarness([], { 'Applicant.Name': 'Ada' });
  const projection = createWorkPdfCollaborationProjection(
    viewer.registry,
    session,
  );
  await projection.ready;
  await flushMicrotasks();

  expect(viewer.annotation('annotation-native-interop')).toMatchObject({
    id: 'annotation-native-interop',
    pageIndex: 1,
    type: 9,
    rect: {
      origin: { x: 68, y: 78 },
      size: { width: 300, height: 28 },
    },
    segmentRects: [
      {
        origin: { x: 68, y: 78 },
        size: { width: 300, height: 28 },
      },
    ],
    strokeColor: '#ffd400',
    color: '#ffd400',
    opacity: 0.48,
    contents: 'Native note',
    author: 'A3S Agent',
    created: new Date('2026-08-15T08:00:00.000Z'),
  });

  Y.applyUpdate(document, decodeBase64(NATIVE_PDF_UPDATE_ANNOTATION_BASE64));
  await flushMicrotasks();
  expect(viewer.annotation('annotation-native-interop')).toMatchObject({
    strokeColor: '#ff0000',
    color: '#ff0000',
  });

  Y.applyUpdate(document, decodeBase64(NATIVE_PDF_DELETE_ANNOTATION_BASE64));
  await flushMicrotasks();
  expect(viewer.annotation('annotation-native-interop')).toBeUndefined();
  expect(
    readOfficePdfCollaboration(session).annotations.find(
      ({ id }) => id === 'annotation-native-interop',
    ),
  ).toMatchObject({ deleted: true });

  projection.destroy();
  session.destroy();
  document.destroy();
});

test('projects shared PDF overlays after capturing the immutable viewer baseline', async () => {
  const baseHighlight = viewerAnnotation('base-highlight', '#ffff00', 0);
  const removedBase = viewerAnnotation('removed-base', '#cccccc', 1);
  const sharedHighlight = viewerAnnotation('base-highlight', '#ff0000', 0);
  const sharedCreated = viewerAnnotation('shared-created', '#00ff00', 2);
  const initial = {
    ...createPdfCollaborationContent(PDF_COLLABORATION_SOURCE),
    annotations: [
      collaborationAnnotation(sharedHighlight, 'base'),
      {
        ...collaborationAnnotation(removedBase, 'base'),
        deleted: true as const,
      },
      collaborationAnnotation(sharedCreated, 'created'),
    ],
    formValues: [{ id: 'Applicant.Name', value: 'Grace' }],
  } satisfies PdfCollaborationContent;
  const session = pdfSession('projection-initial');
  initializeOfficePdfCollaboration(session, initial);
  const viewer = createPdfViewerHarness([baseHighlight, removedBase], {
    'Applicant.Name': 'Ada',
  });

  const projection = createWorkPdfCollaborationProjection(
    viewer.registry,
    session,
  );
  await projection.ready;
  await flushMicrotasks();

  expect(viewer.annotation('base-highlight')).toMatchObject({
    author: RENDERER_DEFAULT_AUTHOR,
    color: '#ff0000',
  });
  expect(projection.binding.canUndo()).toBe(false);
  expect(viewer.annotation('shared-created')).toMatchObject({
    color: '#00ff00',
  });
  expect(viewer.annotation('removed-base')).toBeUndefined();
  expect(viewer.formValues()).toEqual({ 'Applicant.Name': 'Grace' });
  expect(viewer.historyPurges()).toBeGreaterThan(0);
  expect(readOfficePdfCollaboration(session)).toEqual(initial);

  projection.destroy();
});

test('waits for the source annotation baseline before projecting shared overlays', async () => {
  const base = viewerAnnotation('base-highlight', '#ffff00', 0);
  const shared = viewerAnnotation('base-highlight', '#ff0000', 0);
  const session = pdfSession('projection-baseline-race');
  initializeOfficePdfCollaboration(session, {
    ...createPdfCollaborationContent(PDF_COLLABORATION_SOURCE),
    annotations: [collaborationAnnotation(shared, 'base')],
  });
  const viewer = createPdfViewerHarness(
    [base],
    {},
    {
      annotationsLoaded: false,
    },
  );

  const projection = createWorkPdfCollaborationProjection(
    viewer.registry,
    session,
  );
  await flushMicrotasks();
  expect(viewer.annotation('base-highlight')).toBeUndefined();

  viewer.finishAnnotationLoad();
  await projection.ready;
  await flushMicrotasks();

  expect(viewer.annotation('base-highlight')).toMatchObject({
    color: '#ff0000',
  });
  projection.destroy();
});

test('reports a mismatched PDF baseline without leaking the binding', async () => {
  const expected = viewerAnnotation('expected', '#ffff00', 0);
  const actual = viewerAnnotation('unexpected', '#ff0000', 0);
  const session = pdfSession('projection-baseline-mismatch');
  initializeOfficePdfCollaboration(
    session,
    createPdfCollaborationContent(PDF_COLLABORATION_SOURCE),
  );
  const viewer = createPdfViewerHarness(
    [expected],
    {},
    {
      annotationsLoaded: false,
    },
  );
  const projection = createWorkPdfCollaborationProjection(
    viewer.registry,
    session,
  );
  const errors: unknown[] = [];
  projection.subscribeError((error) => errors.push(error));

  await flushMicrotasks();
  viewer.finishAnnotationLoad([actual]);

  await expect(projection.ready).rejects.toThrow(/source baseline/);
  expect(errors).toHaveLength(1);
  expect(() => projection.binding.content()).toThrow(/destroyed/);
  projection.destroy();
});

test('fails closed when shared form state targets an unknown PDF field', async () => {
  const session = pdfSession('projection-form-mismatch');
  initializeOfficePdfCollaboration(session, {
    ...createPdfCollaborationContent(PDF_COLLABORATION_SOURCE),
    formValues: [{ id: 'Missing.Field', value: 'Grace' }],
  });
  const viewer = createPdfViewerHarness([], { 'Applicant.Name': 'Ada' });
  const projection = createWorkPdfCollaborationProjection(
    viewer.registry,
    session,
  );

  await expect(projection.ready).rejects.toThrow(
    /form field 'Missing\.Field' does not exist/,
  );
  projection.destroy();
});

test('fails closed for PDF annotations that need unsynchronized binary context', async () => {
  const stamp = { ...viewerAnnotation('shared-stamp', '#ff0000', 0), type: 13 };
  const session = pdfSession('projection-unsupported-annotation');
  initializeOfficePdfCollaboration(session, {
    ...createPdfCollaborationContent(PDF_COLLABORATION_SOURCE),
    annotations: [collaborationAnnotation(stamp, 'created')],
  });
  const viewer = createPdfViewerHarness([], {});
  const projection = createWorkPdfCollaborationProjection(
    viewer.registry,
    session,
  );

  await expect(projection.ready).rejects.toThrow(
    /does not support annotation type '13'/,
  );
  projection.destroy();
});

test('fails closed when shared annotation source identity disagrees with the PDF', async () => {
  const source = viewerAnnotation('source-highlight', '#ffff00', 0);
  const missing = viewerAnnotation('missing-base', '#ff0000', 0);
  const session = pdfSession('projection-source-identity');
  initializeOfficePdfCollaboration(session, {
    ...createPdfCollaborationContent(PDF_COLLABORATION_SOURCE),
    annotations: [collaborationAnnotation(missing, 'base')],
  });
  const viewer = createPdfViewerHarness([source], {});
  const projection = createWorkPdfCollaborationProjection(
    viewer.registry,
    session,
  );

  await expect(projection.ready).rejects.toThrow(
    /base annotation 'missing-base' does not exist/,
  );
  projection.destroy();
});

test('commits viewer annotation and form events without echoing remote projection', async () => {
  const base = viewerAnnotation('base-highlight', '#ffff00', 0);
  const initial = createPdfCollaborationContent(PDF_COLLABORATION_SOURCE);
  const session = pdfSession('projection-local');
  initializeOfficePdfCollaboration(session, initial);
  const viewer = createPdfViewerHarness([base], { 'Applicant.Name': 'Ada' });
  const projection = createWorkPdfCollaborationProjection(
    viewer.registry,
    session,
  );
  await projection.ready;

  viewer.emitLocalAnnotation('update', {
    ...base,
    color: '#336699',
    created: new Date('2026-08-14T04:05:06.000Z'),
  });
  viewer.emitLocalFormValues({ 'Applicant.Name': 'Lin' });
  await flushMicrotasks();

  const content = readOfficePdfCollaboration(session);
  expect(content.annotations).toHaveLength(1);
  expect(content.annotations[0]).toMatchObject({
    id: 'base-highlight',
    source: 'base',
    annotation: {
      color: '#336699',
      created: '2026-08-14T04:05:06.000Z',
    },
  });
  expect(content.formValues).toEqual([{ id: 'Applicant.Name', value: 'Lin' }]);

  const remote = createOfficePdfCollaborationBinding(session, {
    origin: session.createOrigin('agent', 'remote-update'),
  });
  const before = remote.content();
  const next = {
    ...before,
    annotations: before.annotations.map((record) => ({
      ...record,
      annotation: { ...record.annotation, color: '#8844cc' },
    })),
    formValues: [{ id: 'Applicant.Name', value: 'Remote' }],
  };
  remote.replace(before, next);
  await flushMicrotasks();

  expect(viewer.annotation('base-highlight')).toMatchObject({
    author: RENDERER_DEFAULT_AUTHOR,
    color: '#8844cc',
  });
  expect(viewer.formValues()).toEqual({ 'Applicant.Name': 'Remote' });
  expect(readOfficePdfCollaboration(session)).toEqual(next);
  expect(projection.binding.canUndo()).toBe(true);

  remote.destroy();
  projection.destroy();
});

test('removes a form overlay when the viewer restores the source value', async () => {
  const session = pdfSession('projection-form-restore');
  initializeOfficePdfCollaboration(session, {
    ...createPdfCollaborationContent(PDF_COLLABORATION_SOURCE),
    formValues: [{ id: 'Applicant.Name', value: 'Grace' }],
  });
  const viewer = createPdfViewerHarness([], { 'Applicant.Name': 'Ada' });
  const projection = createWorkPdfCollaborationProjection(
    viewer.registry,
    session,
  );
  await projection.ready;
  await flushMicrotasks();

  viewer.emitLocalFormValues({ 'Applicant.Name': 'Ada' });
  await flushMicrotasks();

  expect(readOfficePdfCollaboration(session).formValues).toEqual([]);
  projection.destroy();
});

test('uses durable tombstones for viewer deletion and Yjs undo for created overlays', async () => {
  const base = viewerAnnotation('base-highlight', '#ffff00', 0);
  const session = pdfSession('projection-history');
  initializeOfficePdfCollaboration(
    session,
    createPdfCollaborationContent(PDF_COLLABORATION_SOURCE),
  );
  const viewer = createPdfViewerHarness([base], {});
  const projection = createWorkPdfCollaborationProjection(
    viewer.registry,
    session,
  );
  await projection.ready;

  viewer.emitLocalAnnotation('delete', base);
  await flushMicrotasks();
  expect(readOfficePdfCollaboration(session).annotations[0]).toMatchObject({
    id: 'base-highlight',
    source: 'base',
    deleted: true,
  });
  expect(projection.binding.canUndo()).toBe(false);
  expect(projection.binding.undo()).toBe(false);

  const secondSession = pdfSession('projection-created-undo');
  initializeOfficePdfCollaboration(
    secondSession,
    createPdfCollaborationContent(PDF_COLLABORATION_SOURCE),
  );
  const secondViewer = createPdfViewerHarness([], {});
  const secondProjection = createWorkPdfCollaborationProjection(
    secondViewer.registry,
    secondSession,
  );
  await secondProjection.ready;
  const created = viewerAnnotation('local-created', '#00aa88', 1);
  secondViewer.emitLocalAnnotation('create', created);
  await flushMicrotasks();
  expect(secondProjection.binding.canUndo()).toBe(true);
  expect(secondViewer.annotation('local-created')).toBeDefined();

  expect(secondProjection.binding.undo()).toBe(true);
  await flushMicrotasks();
  expect(secondViewer.annotation('local-created')).toBeUndefined();
  expect(readOfficePdfCollaboration(secondSession).annotations).toEqual([]);

  secondProjection.destroy();
  projection.destroy();
});

function createPdfViewerHarness(
  sourceAnnotations: readonly PdfAnnotationObject[],
  sourceFormValues: Record<string, string>,
  options: { annotationsLoaded?: boolean } = {},
) {
  const annotations = new Map(
    ((options.annotationsLoaded ?? true) ? sourceAnnotations : []).map(
      (annotation) => [annotation.id, annotation],
    ),
  );
  let formValues = { ...sourceFormValues };
  let historyPurges = 0;
  const annotationEvents = createEvent<AnnotationEvent>();
  const fieldEvents = createEvent<void>();
  const formReadyEvents = createEvent<void>();
  const annotationScope = {
    getAnnotations: () =>
      [...annotations.values()].map((object) => ({
        commitState: 'synced',
        object,
      })),
    onAnnotationEvent: annotationEvents.subscribe,
    importAnnotations: (items: Array<{ annotation: PdfAnnotationObject }>) => {
      for (const { annotation } of items) {
        annotations.set(annotation.id, annotation);
        queueMicrotask(() =>
          annotationEvents.emit({
            type: 'create',
            documentId: DOCUMENT_ID,
            annotation,
            pageIndex: annotation.pageIndex,
            committed: true,
          }),
        );
      }
    },
    updateAnnotation: (
      _pageIndex: number,
      id: string,
      patch: Partial<PdfAnnotationObject>,
    ) => {
      const previous = annotations.get(id);
      if (!previous) return;
      const annotation = {
        ...previous,
        ...patch,
        author: patch.author ?? RENDERER_DEFAULT_AUTHOR,
      } as PdfAnnotationObject;
      annotations.set(id, annotation);
      queueMicrotask(() =>
        annotationEvents.emit({
          type: 'update',
          documentId: DOCUMENT_ID,
          annotation,
          pageIndex: annotation.pageIndex,
          patch,
          committed: true,
        }),
      );
    },
    deleteAnnotation: (pageIndex: number, id: string) => {
      const annotation = annotations.get(id);
      if (!annotation) return;
      annotations.delete(id);
      queueMicrotask(() =>
        annotationEvents.emit({
          type: 'delete',
          documentId: DOCUMENT_ID,
          annotation,
          pageIndex,
          committed: true,
        }),
      );
    },
  };
  const formScope = {
    getFormValues: () => ({ ...formValues }),
    getFormFields: () =>
      Object.entries(formValues).map(([name, value]) => ({ name, value })),
    setFormValues: (values: Record<string, string>) => {
      formValues = { ...formValues, ...values };
      queueMicrotask(() => fieldEvents.emit(undefined));
      return resolvedPdfTask(true);
    },
    onFieldValueChange: fieldEvents.subscribe,
    onFormReady: formReadyEvents.subscribe,
  };
  const sourceByPage = Object.groupBy(
    sourceAnnotations,
    ({ pageIndex }) => pageIndex,
  );
  const capabilities = {
    annotation: {
      forDocument: () => annotationScope,
    },
    'document-manager': {
      getActiveDocumentId: () => DOCUMENT_ID,
      getDocument: () => ({
        id: DOCUMENT_ID,
        pageCount: PDF_COLLABORATION_SOURCE.pageCount,
        pages: Array.from(
          { length: PDF_COLLABORATION_SOURCE.pageCount },
          (_, index) => ({ index }),
        ),
      }),
    },
    form: {
      forDocument: () => formScope,
    },
    history: {
      forDocument: () => ({
        purgeByMetadata: () => {
          historyPurges += 1;
          return 1;
        },
      }),
    },
  };
  const registry = {
    pluginsReady: () => Promise.resolve(),
    getPlugin: (id: keyof typeof capabilities) => {
      const capability = capabilities[id];
      return capability ? { provides: () => capability } : undefined;
    },
    getEngine: () => ({
      getAllAnnotations: () => resolvedPdfTask(sourceByPage),
    }),
  } as unknown as PluginRegistry;

  return {
    registry,
    annotation: (id: string) => annotations.get(id),
    formValues: () => ({ ...formValues }),
    historyPurges: () => historyPurges,
    finishAnnotationLoad(
      loadedAnnotations: readonly PdfAnnotationObject[] = sourceAnnotations,
    ) {
      annotations.clear();
      for (const annotation of loadedAnnotations) {
        annotations.set(annotation.id, annotation);
      }
      annotationEvents.emit({
        type: 'loaded',
        documentId: DOCUMENT_ID,
        total: loadedAnnotations.length,
      });
      formReadyEvents.emit(undefined);
    },
    emitLocalAnnotation(
      type: 'create' | 'update' | 'delete',
      annotation: PdfAnnotationObject,
    ) {
      if (type === 'delete') annotations.delete(annotation.id);
      else annotations.set(annotation.id, annotation);
      const event = {
        type,
        documentId: DOCUMENT_ID,
        annotation,
        pageIndex: annotation.pageIndex,
        committed: true,
        ...(type === 'update' ? { patch: annotation } : {}),
      } as AnnotationEvent;
      annotationEvents.emit(event);
    },
    emitLocalFormValues(values: Record<string, string>) {
      formValues = { ...values };
      fieldEvents.emit(undefined);
    },
  };
}

function viewerAnnotation(
  id: string,
  color: string,
  pageIndex: number,
): PdfAnnotationObject {
  return {
    id,
    pageIndex,
    type: 9,
    rect: {
      origin: { x: 10, y: 20 },
      size: { width: 100, height: 16 },
    },
    color,
    opacity: 0.5,
  } as PdfAnnotationObject;
}

function collaborationAnnotation(
  annotation: PdfAnnotationObject,
  source: 'base' | 'created',
): PdfCollaborationContent['annotations'][number] {
  return {
    id: annotation.id,
    pageIndex: annotation.pageIndex,
    source,
    annotation: structuredClone(annotation),
  };
}

function pdfSession(artifactId: string): OfficeCollaborationSession {
  return createOfficeCollaborationSession({
    artifactId,
    document: new Y.Doc(),
    kind: 'pdf',
  });
}

function resolvedPdfTask<T>(value: T) {
  return {
    wait(onResolve: (resolved: T) => void) {
      queueMicrotask(() => onResolve(value));
    },
    toPromise: () => Promise.resolve(value),
  };
}

function decodeBase64(value: string): Uint8Array {
  return Uint8Array.from(atob(value), (character) => character.charCodeAt(0));
}

function createEvent<T>() {
  const listeners = new Set<(value: T) => void>();
  return {
    emit(value: T) {
      for (const listener of listeners) listener(value);
    },
    subscribe(listener: (value: T) => void) {
      listeners.add(listener);
      return () => listeners.delete(listener);
    },
  };
}

async function flushMicrotasks(): Promise<void> {
  await Promise.resolve();
  await Promise.resolve();
  await Promise.resolve();
}
