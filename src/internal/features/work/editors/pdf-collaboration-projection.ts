import {
  type AnnotationEvent,
  AnnotationPlugin,
  type AnnotationScope,
  DocumentManagerPlugin,
  FormPlugin,
  type FormScope,
  HistoryPlugin,
  type HistoryScope,
  type PdfAnnotationObject,
  PdfAnnotationSubtype,
  type PluginRegistry,
} from '@embedpdf/react-pdf-viewer';
import type { WorkOfficeCollaborationSession } from '../../../collaboration/office-collaboration';
import {
  createWorkOfficePdfCollaborationBinding,
  type WorkOfficePdfCollaborationBinding,
} from '../../../collaboration/office-pdf-collaboration';
import type {
  WorkPdfCollaborationAnnotation,
  WorkPdfCollaborationContent,
} from '../../../collaboration/office-pdf-collaboration-types';

export interface WorkPdfCollaborationProjection {
  readonly binding: WorkOfficePdfCollaborationBinding;
  readonly ready: Promise<void>;
  subscribeError(listener: (error: unknown) => void): () => void;
  destroy(): void;
}

interface PdfProjectionScopes {
  annotations: AnnotationScope;
  baseAnnotations: Map<string, PdfAnnotationObject>;
  baseFormValues: Record<string, string>;
  documentId: string;
  form: FormScope | null;
  history: HistoryScope | null;
}

interface PdfProjectionPlugins {
  annotation: ReturnType<AnnotationPlugin['provides']>;
  documentId: string;
  form: FormScope | null;
  history: HistoryScope | null;
}

interface PdfFormProjection {
  expected: Record<string, string>;
  updates: Record<string, string>;
}

interface PendingPdfAnnotationProjection {
  expected: PdfAnnotationObject | 'deleted';
  previous?: PdfAnnotationObject;
}

const COLLABORATIVE_PDF_ANNOTATION_TYPES = new Set<PdfAnnotationSubtype>([
  PdfAnnotationSubtype.FREETEXT,
  PdfAnnotationSubtype.HIGHLIGHT,
  PdfAnnotationSubtype.UNDERLINE,
  PdfAnnotationSubtype.STRIKEOUT,
  PdfAnnotationSubtype.INK,
]);

export function createWorkPdfCollaborationProjection(
  registry: PluginRegistry,
  session: WorkOfficeCollaborationSession,
): WorkPdfCollaborationProjection {
  const binding = createWorkOfficePdfCollaborationBinding(session);
  const initialization = new AbortController();
  let destroyed = false;
  let failed = false;
  let failure: unknown;
  let projecting = false;
  let scopes: PdfProjectionScopes | null = null;
  let previousRecords = new Map<string, WorkPdfCollaborationAnnotation>();
  const pendingAnnotations = new Map<string, PendingPdfAnnotationProjection>();
  let pendingFormValues: Record<string, string> | null = null;
  const errorListeners = new Set<(error: unknown) => void>();
  const stop: Array<() => void> = [];
  const close = () => {
    initialization.abort();
    scopes = null;
    pendingAnnotations.clear();
    pendingFormValues = null;
    for (const unsubscribe of stop.splice(0)) unsubscribe();
    try {
      binding.destroy();
    } catch {
      // A failed initialization may already have closed the binding.
    }
  };
  const reportError = (error: unknown) => {
    if (destroyed || failed) return;
    failed = true;
    failure = error;
    close();
    for (const listener of errorListeners) listener(error);
  };
  const project = (content: WorkPdfCollaborationContent) => {
    if (!scopes || destroyed || failed) return;
    projecting = true;
    try {
      projectAnnotations(
        scopes.annotations,
        scopes.baseAnnotations,
        previousRecords,
        content.annotations,
        pendingAnnotations,
      );
      previousRecords = new Map(
        content.annotations.map((record) => [record.id, record]),
      );
      const projectedFormValues = projectedPdfFormValues(
        scopes.form,
        scopes.baseFormValues,
        content.formValues,
      );
      if (projectedFormValues && scopes.form) {
        pendingFormValues = projectedFormValues.expected;
        applyPdfFormValues(
          scopes.form,
          projectedFormValues.updates,
          reportError,
        );
      }
      purgeViewerHistory(scopes.history);
    } finally {
      projecting = false;
    }
  };
  const ready = registry
    .pluginsReady()
    .then(async () => {
      if (destroyed || failed) return;
      const plugins = pdfProjectionPlugins(registry);
      const annotationsLoaded = await waitForPdfAnnotationsLoaded(
        registry,
        plugins,
        initialization.signal,
      );
      if (!annotationsLoaded || destroyed || failed) return;
      scopes = pdfProjectionScopes(plugins);
      const activeScopes = scopes;
      stop.push(
        activeScopes.annotations.onAnnotationEvent((event) => {
          if (destroyed || projecting || event.type === 'loaded') return;
          if (!event.committed || event.documentId !== activeScopes.documentId)
            return;
          try {
            const pending = pendingAnnotations.get(event.annotation.id);
            if (projectedAnnotationEventMatches(pending, event)) {
              pendingAnnotations.delete(event.annotation.id);
              purgeViewerHistory(activeScopes.history);
              return;
            }
            commitAnnotationEvent(binding, event);
            purgeViewerHistory(activeScopes.history);
          } catch (error) {
            reportError(error);
          }
        }),
      );
      if (activeScopes.form) {
        const form = activeScopes.form;
        stop.push(
          form.onFieldValueChange(() => {
            if (destroyed || projecting) return;
            try {
              const values = form.getFormValues();
              if (pendingFormValues) {
                if (pdfJsonEqual(values, pendingFormValues)) {
                  pendingFormValues = null;
                  purgeViewerHistory(activeScopes.history);
                }
                return;
              }
              commitFormValues(
                binding,
                activeScopes.baseFormValues,
                form,
                values,
              );
              purgeViewerHistory(activeScopes.history);
            } catch (error) {
              reportError(error);
            }
          }),
        );
        stop.push(
          form.onFormReady(() => {
            if (destroyed) return;
            try {
              const projectedFormValues = projectedPdfFormValues(
                form,
                activeScopes.baseFormValues,
                binding.content().formValues,
              );
              if (projectedFormValues) {
                pendingFormValues = projectedFormValues.expected;
                applyPdfFormValues(
                  form,
                  projectedFormValues.updates,
                  reportError,
                );
              }
              purgeViewerHistory(activeScopes.history);
            } catch (error) {
              reportError(error);
            }
          }),
        );
      }
      project(binding.content());
    })
    .catch((error: unknown) => {
      reportError(error);
      throw error;
    });
  stop.push(
    binding.subscribe(({ content }) => {
      project(content);
    }),
  );
  stop.push(binding.subscribeError(reportError));

  return {
    binding,
    ready,
    subscribeError(listener) {
      if (failed) {
        queueMicrotask(() => listener(failure));
        return () => undefined;
      }
      errorListeners.add(listener);
      return () => errorListeners.delete(listener);
    },
    destroy() {
      if (destroyed) return;
      destroyed = true;
      close();
      errorListeners.clear();
    },
  };
}

function pdfProjectionPlugins(registry: PluginRegistry): PdfProjectionPlugins {
  const annotation = registry
    .getPlugin<AnnotationPlugin>(AnnotationPlugin.id)
    ?.provides();
  const documentId = registry
    .getPlugin<DocumentManagerPlugin>(DocumentManagerPlugin.id)
    ?.provides()
    .getActiveDocumentId();
  if (!annotation || !documentId) {
    throw new Error(
      'PDF collaboration requires an open EmbedPDF annotation document.',
    );
  }
  const form =
    registry
      .getPlugin<FormPlugin>(FormPlugin.id)
      ?.provides()
      .forDocument(documentId) ?? null;
  const history =
    registry
      .getPlugin<HistoryPlugin>(HistoryPlugin.id)
      ?.provides()
      .forDocument(documentId) ?? null;
  return { annotation, documentId, form, history };
}

function pdfProjectionScopes({
  annotation,
  documentId,
  form,
  history,
}: PdfProjectionPlugins): PdfProjectionScopes {
  const annotations = annotation.forDocument(documentId);
  return {
    annotations,
    baseAnnotations: new Map(
      annotations
        .getAnnotations()
        .map(({ object }) => [
          object.id,
          viewerAnnotation(portableViewerAnnotation(object)),
        ]),
    ),
    baseFormValues: form?.getFormValues() ?? {},
    documentId,
    form,
    history,
  };
}

async function waitForPdfAnnotationsLoaded(
  registry: PluginRegistry,
  { annotation, documentId }: PdfProjectionPlugins,
  signal: AbortSignal,
): Promise<boolean> {
  if (signal.aborted) return false;
  const document = registry
    .getPlugin<DocumentManagerPlugin>(DocumentManagerPlugin.id)
    ?.provides()
    .getDocument(documentId);
  if (!document) {
    throw new Error(
      'PDF collaboration requires an open EmbedPDF annotation document.',
    );
  }
  const scope = annotation.forDocument(documentId);
  let resolveLoaded: () => void = () => undefined;
  const loaded = new Promise<void>((resolve) => {
    resolveLoaded = resolve;
  });
  const stopLoaded = scope.onAnnotationEvent((event) => {
    if (event.type === 'loaded' && event.documentId === documentId) {
      resolveLoaded();
    }
  });
  let resolveAborted: () => void = () => undefined;
  const aborted = new Promise<void>((resolve) => {
    resolveAborted = resolve;
  });
  const onAbort = () => resolveAborted();
  signal.addEventListener('abort', onAbort, { once: true });
  if (signal.aborted) onAbort();
  try {
    const expected = await Promise.race([
      registry.getEngine().getAllAnnotations(document).toPromise(),
      aborted.then(() => null),
    ]);
    if (!expected || signal.aborted) return false;
    const expectedIds = new Set(
      Object.values(expected).flatMap((records) => records.map(({ id }) => id)),
    );
    if (pdfAnnotationIdsMatch(scope, expectedIds)) return true;
    await Promise.race([loaded, aborted]);
    if (signal.aborted) return false;
    if (pdfAnnotationIdsMatch(scope, expectedIds)) return true;
    throw new Error('PDF collaboration could not capture the source baseline.');
  } finally {
    signal.removeEventListener('abort', onAbort);
    stopLoaded();
  }
}

function pdfAnnotationIdsMatch(
  scope: AnnotationScope,
  expectedIds: ReadonlySet<string>,
): boolean {
  const actualIds = new Set(
    scope.getAnnotations().map(({ object }) => object.id),
  );
  return (
    expectedIds.size === actualIds.size &&
    [...expectedIds].every((id) => actualIds.has(id))
  );
}

function purgeViewerHistory(history: HistoryScope | null): void {
  history?.purgeByMetadata(() => true);
}

function commitAnnotationEvent(
  binding: WorkOfficePdfCollaborationBinding,
  event: Exclude<AnnotationEvent, { type: 'loaded' }>,
): void {
  const previous = binding.content();
  const current = previous.annotations.find(
    ({ id }) => id === event.annotation.id,
  );
  assertCollaborativePdfAnnotation(event.annotation);
  let annotations = previous.annotations;
  if (event.type === 'delete') {
    if (current?.deleted) return;
    annotations = current
      ? previous.annotations.map((record) =>
          record.id === current.id ? { ...record, deleted: true } : record,
        )
      : [
          ...previous.annotations,
          {
            id: event.annotation.id,
            pageIndex: event.pageIndex,
            source: 'base',
            annotation: portableViewerAnnotation(event.annotation),
            deleted: true,
          },
        ];
  } else {
    const annotation = portableViewerAnnotation(event.annotation);
    if (current) {
      annotations = previous.annotations.map((record) =>
        record.id === current.id
          ? { ...record, annotation, pageIndex: event.pageIndex }
          : record,
      );
    } else {
      annotations = [
        ...previous.annotations,
        {
          id: event.annotation.id,
          pageIndex: event.pageIndex,
          source: event.type === 'create' ? 'created' : 'base',
          annotation,
        },
      ];
    }
  }
  binding.replace(previous, { ...previous, annotations });
}

function commitFormValues(
  binding: WorkOfficePdfCollaborationBinding,
  baseValues: Readonly<Record<string, string>>,
  scope: FormScope,
  values: Record<string, string>,
): void {
  const previous = binding.content();
  const fields = new Map(
    scope.getFormFields().map((field) => [field.name, field]),
  );
  const formValues = Object.entries(values)
    .filter(([id, value]) => {
      const field = fields.get(id);
      if (!field) {
        throw new Error(`PDF collaboration form field '${id}' does not exist.`);
      }
      if (field.readOnly && value !== baseValues[id]) {
        throw new Error(`PDF collaboration form field '${id}' is read-only.`);
      }
      return value !== baseValues[id];
    })
    .map(([id, value]) => ({ id, value }))
    .sort((left, right) => left.id.localeCompare(right.id));
  binding.replace(previous, { ...previous, formValues });
}

function projectAnnotations(
  scope: AnnotationScope,
  baseAnnotations: ReadonlyMap<string, PdfAnnotationObject>,
  previous: ReadonlyMap<string, WorkPdfCollaborationAnnotation>,
  records: readonly WorkPdfCollaborationAnnotation[],
  pending: Map<string, PendingPdfAnnotationProjection>,
): void {
  assertPdfAnnotationSources(baseAnnotations, records);
  const current = new Map(
    scope.getAnnotations().map(({ object }) => [object.id, object]),
  );
  const nextIds = new Set(records.map(({ id }) => id));
  for (const [id, record] of previous) {
    if (nextIds.has(id)) continue;
    const existing = current.get(id);
    if (record.source === 'created') {
      if (existing) {
        pending.set(id, { expected: 'deleted', previous: existing });
        scope.deleteAnnotation(existing.pageIndex, id);
      }
      continue;
    }
    const base = baseAnnotations.get(id);
    if (!base || pdfJsonEqual(existing, base)) continue;
    pending.set(id, { expected: base, previous: existing });
    if (existing) scope.updateAnnotation(existing.pageIndex, id, base);
    else scope.importAnnotations([{ annotation: base }]);
  }
  for (const record of records) {
    const existing = current.get(record.id);
    if (record.deleted) {
      if (existing) {
        pending.set(record.id, {
          expected: 'deleted',
          previous: existing,
        });
        scope.deleteAnnotation(existing.pageIndex, existing.id);
      }
      continue;
    }
    const expected = viewerAnnotation(record.annotation);
    assertCollaborativePdfAnnotation(expected);
    if (!existing) {
      pending.set(record.id, { expected });
      scope.importAnnotations([{ annotation: expected }]);
      continue;
    }
    if (
      projectedPdfAnnotationEqual(
        existing,
        expected,
        record.source === 'base' ? baseAnnotations.get(record.id) : undefined,
      )
    ) {
      continue;
    }
    pending.set(record.id, { expected, previous: existing });
    scope.updateAnnotation(existing.pageIndex, existing.id, {
      ...expected,
      id: existing.id,
      pageIndex: existing.pageIndex,
    });
  }
}

function assertPdfAnnotationSources(
  baseAnnotations: ReadonlyMap<string, PdfAnnotationObject>,
  records: readonly WorkPdfCollaborationAnnotation[],
): void {
  for (const record of records) {
    const inSource = baseAnnotations.has(record.id);
    if (record.source === 'base' && !inSource) {
      throw new Error(
        `PDF collaboration base annotation '${record.id}' does not exist in the source PDF.`,
      );
    }
    if (record.source === 'created' && inSource) {
      throw new Error(
        `PDF collaboration annotation '${record.id}' conflicts with a source annotation.`,
      );
    }
    const base = baseAnnotations.get(record.id);
    if (
      base &&
      (base.pageIndex !== record.pageIndex ||
        base.type !== record.annotation.type)
    ) {
      throw new Error(
        `PDF collaboration base annotation '${record.id}' has a different source identity.`,
      );
    }
  }
}

function assertCollaborativePdfAnnotation(
  annotation: PdfAnnotationObject,
): void {
  if (
    COLLABORATIVE_PDF_ANNOTATION_TYPES.has(
      annotation.type as PdfAnnotationSubtype,
    )
  ) {
    return;
  }
  throw new Error(
    `PDF collaboration does not support annotation type '${annotation.type}'.`,
  );
}

function projectedPdfFormValues(
  scope: FormScope | null,
  baseValues: Record<string, string>,
  records: WorkPdfCollaborationContent['formValues'],
): PdfFormProjection | null {
  if (!scope) return null;
  const fields = new Map(
    scope.getFormFields().map((field) => [field.name, field]),
  );
  for (const { id, value } of records) {
    const field = fields.get(id);
    if (!field) {
      throw new Error(`PDF collaboration form field '${id}' does not exist.`);
    }
    if (field.readOnly && value !== baseValues[id]) {
      throw new Error(`PDF collaboration form field '${id}' is read-only.`);
    }
  }
  const expected = {
    ...baseValues,
    ...Object.fromEntries(records.map(({ id, value }) => [id, value])),
  };
  const current = scope.getFormValues();
  const updates = Object.fromEntries(
    Object.entries(expected).filter(([id, value]) => current[id] !== value),
  );
  if (Object.keys(updates).length === 0) return null;
  return { expected, updates };
}

function applyPdfFormValues(
  scope: FormScope,
  values: Record<string, string>,
  onError: (error: unknown) => void,
): void {
  void scope.setFormValues(values).toPromise().catch(onError);
}

function projectedAnnotationEventMatches(
  pending: PendingPdfAnnotationProjection | undefined,
  event: Exclude<AnnotationEvent, { type: 'loaded' }>,
): boolean {
  if (pending === undefined) return false;
  if (pending.expected === 'deleted') return event.type === 'delete';
  return (
    event.type !== 'delete' &&
    projectedPdfAnnotationEqual(
      event.annotation,
      pending.expected,
      pending.previous,
    )
  );
}

/**
 * EmbedPDF merges an update patch into its current object and supplies its
 * configured author when the patch does not contain one. Those renderer-owned
 * additions must not turn a remote projection into a local Yjs edit. Shared
 * fields still compare exactly, and an extra field is ignored only when it is
 * the renderer author or is unchanged from the object that was projected.
 */
function projectedPdfAnnotationEqual(
  actual: PdfAnnotationObject,
  expected: PdfAnnotationObject,
  previous?: PdfAnnotationObject,
): boolean {
  const normalizedActual = portableViewerAnnotation(actual);
  const normalizedExpected = portableViewerAnnotation(expected);
  const normalizedPrevious = previous
    ? portableViewerAnnotation(previous)
    : undefined;
  for (const key of Object.keys(normalizedActual)) {
    if (Object.hasOwn(normalizedExpected, key)) continue;
    if (
      key === 'author' ||
      (normalizedPrevious &&
        Object.hasOwn(normalizedPrevious, key) &&
        pdfJsonEqual(normalizedActual[key], normalizedPrevious[key]))
    ) {
      delete normalizedActual[key];
    }
  }
  return pdfJsonEqual(normalizedActual, normalizedExpected);
}

function portableViewerAnnotation(
  annotation: PdfAnnotationObject,
): Record<string, unknown> {
  return clonePdfValue(annotation) as Record<string, unknown>;
}

function viewerAnnotation(
  annotation: Record<string, unknown>,
): PdfAnnotationObject {
  return revivePdfDates(annotation) as PdfAnnotationObject;
}

function clonePdfValue(value: unknown): unknown {
  if (value instanceof Date) return value.toISOString();
  if (Array.isArray(value)) return value.map(clonePdfValue);
  if (!value || typeof value !== 'object') return value;
  return Object.fromEntries(
    Object.entries(value).flatMap(([key, child]) =>
      child === undefined ? [] : [[key, clonePdfValue(child)]],
    ),
  );
}

function revivePdfDates(value: unknown, key?: string): unknown {
  if (Array.isArray(value)) return value.map((child) => revivePdfDates(child));
  if (!value || typeof value !== 'object') {
    if (
      (key === 'created' || key === 'modified') &&
      typeof value === 'string'
    ) {
      const timestamp = Date.parse(value);
      if (Number.isFinite(timestamp)) return new Date(timestamp);
    }
    return value;
  }
  return Object.fromEntries(
    Object.entries(value).map(([childKey, child]) => [
      childKey,
      revivePdfDates(child, childKey),
    ]),
  );
}

function pdfJsonEqual(left: unknown, right: unknown): boolean {
  return canonicalPdfJson(left) === canonicalPdfJson(right);
}

function canonicalPdfJson(value: unknown): string {
  return JSON.stringify(sortPdfJson(clonePdfValue(value)));
}

function sortPdfJson(value: unknown): unknown {
  if (Array.isArray(value)) return value.map(sortPdfJson);
  if (!value || typeof value !== 'object') return value;
  return Object.fromEntries(
    Object.entries(value)
      .sort(([left], [right]) => left.localeCompare(right))
      .map(([key, child]) => [key, sortPdfJson(child)]),
  );
}
