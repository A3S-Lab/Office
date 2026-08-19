import type {
  WorkDocumentContent,
  WorkDocumentMark,
  WorkDocumentModel,
  WorkDocumentNode,
} from './work-types';
import {
  documentHtmlFingerprint,
  documentHtmlFingerprintMatches,
} from './work-document-html-fingerprint';
import {
  prepareLazyDocumentEditorSource,
  transferLazyDocumentModelState,
} from './work-document-lazy-model';

export { documentHtmlFingerprint } from './work-document-html-fingerprint';

const DOCUMENT_MODEL_SCHEMA = 'a3s.office.document';
const DOCUMENT_MODEL_VERSION = 1;
const MAX_DOCUMENT_MODEL_DEPTH = 256;
// A 100,000-row, three-column Word table contains roughly one million
// semantic nodes even when its DOM is windowed. These limits remain bounded
// against hostile snapshots while admitting the supported large-file scale.
const MAX_DOCUMENT_MODEL_NODES = 2_000_000;
const MAX_DOCUMENT_MODEL_MARKS = 2_000_000;
const MAX_DOCUMENT_MODEL_ATTRIBUTE_VALUES = 10_000_000;

interface SchemaValidatedDocumentModelOptions {
  initialIntegrityFeatures?: number;
  previous?: WorkDocumentModel | null;
}

interface SchemaValidatedDocumentModel {
  html: string;
  htmlFingerprint: string;
  initialIntegrityFeatures: number | null;
  root: WorkDocumentNode;
}

const schemaValidatedDocumentModels = new WeakMap<
  WorkDocumentModel,
  SchemaValidatedDocumentModel
>();

export function createWorkDocumentModel(
  html: string,
  root: WorkDocumentNode,
  previous?: WorkDocumentModel | null,
): WorkDocumentModel {
  return createWorkDocumentModelWithFingerprint(
    root,
    documentHtmlFingerprint(html),
    previous,
  );
}

function createWorkDocumentModelWithFingerprint(
  root: WorkDocumentNode,
  htmlFingerprint: string,
  previous?: WorkDocumentModel | null,
): WorkDocumentModel {
  const model: WorkDocumentModel = {
    schema: DOCUMENT_MODEL_SCHEMA,
    version: DOCUMENT_MODEL_VERSION,
    revision: nextRevision(previous),
    htmlFingerprint,
    root,
  };
  transferLazyDocumentModelState(previous, model);
  return model;
}

/**
 * Creates a model for a root assembled by a schema-constrained parser. The
 * trust record is process-local and is lost across cloning or persistence, so
 * external snapshots still receive the complete structural validation.
 */
export function createSchemaValidatedWorkDocumentModel(
  html: string,
  root: WorkDocumentNode,
  options: SchemaValidatedDocumentModelOptions = {},
): WorkDocumentModel {
  const model = createWorkDocumentModel(html, root, options.previous);
  schemaValidatedDocumentModels.set(model, {
    html,
    htmlFingerprint: model.htmlFingerprint,
    initialIntegrityFeatures: nonNegativeInteger(
      options.initialIntegrityFeatures,
    ),
    root,
  });
  return model;
}

/**
 * Records a root emitted by the live TipTap schema without rescanning up to two
 * million JSON nodes on the next controlled render.
 */
export function createSchemaDerivedWorkDocumentModel(
  html: string,
  root: WorkDocumentNode,
  previous?: WorkDocumentModel | null,
  htmlFingerprint = documentHtmlFingerprint(html),
): WorkDocumentModel {
  const model = createWorkDocumentModelWithFingerprint(
    root,
    htmlFingerprint,
    previous,
  );
  schemaValidatedDocumentModels.set(model, {
    html,
    htmlFingerprint: model.htmlFingerprint,
    initialIntegrityFeatures:
      (previous &&
        schemaValidatedDocumentModels.get(previous)
          ?.initialIntegrityFeatures) ??
      null,
    root,
  });
  return model;
}

export function documentModelHasTrustedInitialIntegrityFeatures(
  model: WorkDocumentModel | null,
): boolean {
  if (!model) return false;
  const validated = schemaValidatedDocumentModels.get(model);
  return (
    validated?.root === model.root &&
    validated.initialIntegrityFeatures !== null
  );
}

export function documentModelForContent(
  content: WorkDocumentContent,
): WorkDocumentModel | null {
  return documentModelForHtml(content.model, content.html);
}

export function documentModelForHtml(
  candidate: unknown,
  html: string,
): WorkDocumentModel | null {
  if (!isRecord(candidate)) return null;
  const validated = schemaValidatedDocumentModels.get(
    candidate as unknown as WorkDocumentModel,
  );
  const trustedHtmlMatches = Boolean(
    validated &&
      validated.root === candidate.root &&
      validated.html === html &&
      validated.htmlFingerprint === candidate.htmlFingerprint,
  );
  if (
    candidate.schema !== DOCUMENT_MODEL_SCHEMA ||
    candidate.version !== DOCUMENT_MODEL_VERSION ||
    !Number.isSafeInteger(candidate.revision) ||
    Number(candidate.revision) < 1 ||
    typeof candidate.htmlFingerprint !== 'string' ||
    (!trustedHtmlMatches &&
      !documentHtmlFingerprintMatches(html, candidate.htmlFingerprint)) ||
    (validated?.root !== candidate.root && !isDocumentRoot(candidate.root))
  ) {
    return null;
  }
  return candidate as unknown as WorkDocumentModel;
}

export interface WorkDocumentEditorInput {
  model: WorkDocumentModel | null;
  source: WorkDocumentNode | string;
  sourceKey: string;
  revision: number;
}

export function resolveWorkDocumentEditorInput(
  content: WorkDocumentContent,
  fallbackHtml: string,
  validatedModel: WorkDocumentModel | null = documentModelForContent(content),
): WorkDocumentEditorInput {
  // The structured model is synchronized to the controlled `content.html`
  // value. `fallbackHtml` may be a browser-normalized legacy projection, so
  // validating against it would discard a valid model and turn mount-time HTML
  // normalization into a false user edit.
  const model = validatedModel;
  if (!model) {
    return {
      model: null,
      source: fallbackHtml,
      sourceKey: `html:${documentHtmlFingerprint(fallbackHtml)}`,
      revision: 0,
    };
  }
  return {
    model,
    source:
      prepareLazyDocumentEditorSource(
        model,
        documentModelHasTrustedInitialIntegrityFeatures(model),
        content.html,
      )?.root ?? model.root,
    sourceKey: `model:${model.revision}:${model.htmlFingerprint}`,
    revision: model.revision,
  };
}

function nextRevision(previous: WorkDocumentModel | null | undefined): number {
  const revision =
    previous?.schema === DOCUMENT_MODEL_SCHEMA &&
    previous.version === DOCUMENT_MODEL_VERSION &&
    Number.isSafeInteger(previous.revision) &&
    previous.revision > 0
      ? previous.revision
      : 0;
  return revision >= Number.MAX_SAFE_INTEGER ? 1 : revision + 1;
}

function nonNegativeInteger(value: unknown): number | null {
  const number = Number(value);
  return Number.isSafeInteger(number) && number >= 0 ? number : null;
}

function isDocumentRoot(value: unknown): value is WorkDocumentNode {
  if (!isRecord(value) || value.type !== 'doc') return false;
  const nodes: Array<{ value: unknown; depth: number }> = [{ value, depth: 0 }];
  const seen = new WeakSet<object>();
  let nodeCount = 0;
  let markCount = 0;
  let attributeValueCount = 0;

  while (nodes.length) {
    const current = nodes.pop();
    if (!current || !isRecord(current.value)) return false;
    if (current.depth > MAX_DOCUMENT_MODEL_DEPTH || seen.has(current.value)) {
      return false;
    }
    seen.add(current.value);
    nodeCount += 1;
    if (nodeCount > MAX_DOCUMENT_MODEL_NODES) return false;

    if (
      typeof current.value.type !== 'string' ||
      !current.value.type.trim() ||
      (current.value.text !== undefined &&
        typeof current.value.text !== 'string')
    ) {
      return false;
    }
    if (
      current.value.type === 'text' &&
      typeof current.value.text !== 'string'
    ) {
      return false;
    }

    if (current.value.attrs !== undefined) {
      const result = validateAttributes(
        current.value.attrs,
        attributeValueCount,
      );
      if (!result.valid) return false;
      attributeValueCount = result.count;
    }

    if (current.value.marks !== undefined) {
      if (!Array.isArray(current.value.marks)) return false;
      markCount += current.value.marks.length;
      if (markCount > MAX_DOCUMENT_MODEL_MARKS) return false;
      for (const mark of current.value.marks) {
        if (!isDocumentMark(mark)) return false;
        if (mark.attrs !== undefined) {
          const result = validateAttributes(mark.attrs, attributeValueCount);
          if (!result.valid) return false;
          attributeValueCount = result.count;
        }
      }
    }

    if (current.value.content !== undefined) {
      if (!Array.isArray(current.value.content)) return false;
      for (
        let index = current.value.content.length - 1;
        index >= 0;
        index -= 1
      ) {
        nodes.push({
          value: current.value.content[index],
          depth: current.depth + 1,
        });
      }
    }
  }

  return true;
}

function isDocumentMark(value: unknown): value is WorkDocumentMark {
  return (
    isRecord(value) &&
    typeof value.type === 'string' &&
    Boolean(value.type.trim()) &&
    (value.attrs === undefined || isRecord(value.attrs))
  );
}

function validateAttributes(
  value: unknown,
  initialCount: number,
): { valid: boolean; count: number } {
  if (!isRecord(value)) return { valid: false, count: initialCount };
  const values = Object.values(value);
  const stack = values.map((item) => ({ item, depth: 0 }));
  const seen = new WeakSet<object>();
  let count = initialCount;

  while (stack.length) {
    const current = stack.pop();
    if (!current) break;
    count += 1;
    if (count > MAX_DOCUMENT_MODEL_ATTRIBUTE_VALUES || current.depth > 32) {
      return { valid: false, count };
    }
    if (
      current.item === null ||
      typeof current.item === 'string' ||
      typeof current.item === 'boolean' ||
      (typeof current.item === 'number' && Number.isFinite(current.item))
    ) {
      continue;
    }
    if (typeof current.item !== 'object') {
      return { valid: false, count };
    }
    if (seen.has(current.item)) return { valid: false, count };
    seen.add(current.item);
    const nested = Array.isArray(current.item)
      ? current.item
      : Object.values(current.item);
    for (const item of nested) {
      stack.push({ item, depth: current.depth + 1 });
    }
  }

  return { valid: true, count };
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}
