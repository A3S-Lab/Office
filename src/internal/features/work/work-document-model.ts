import type {
  WorkDocumentContent,
  WorkDocumentMark,
  WorkDocumentModel,
  WorkDocumentNode,
} from './work-types';

const DOCUMENT_MODEL_SCHEMA = 'a3s.office.document';
const DOCUMENT_MODEL_VERSION = 1;
const MAX_DOCUMENT_MODEL_DEPTH = 256;
const MAX_DOCUMENT_MODEL_NODES = 100_000;
const MAX_DOCUMENT_MODEL_MARKS = 500_000;
const MAX_DOCUMENT_MODEL_ATTRIBUTE_VALUES = 1_000_000;

export function createWorkDocumentModel(
  html: string,
  root: WorkDocumentNode,
  previous?: WorkDocumentModel | null,
): WorkDocumentModel {
  return {
    schema: DOCUMENT_MODEL_SCHEMA,
    version: DOCUMENT_MODEL_VERSION,
    revision: nextRevision(previous),
    htmlFingerprint: documentHtmlFingerprint(html),
    root,
  };
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
  if (
    candidate.schema !== DOCUMENT_MODEL_SCHEMA ||
    candidate.version !== DOCUMENT_MODEL_VERSION ||
    !Number.isSafeInteger(candidate.revision) ||
    Number(candidate.revision) < 1 ||
    candidate.htmlFingerprint !== documentHtmlFingerprint(html) ||
    !isDocumentRoot(candidate.root)
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
): WorkDocumentEditorInput {
  const model = documentModelForHtml(content.model, fallbackHtml);
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
    source: model.root,
    sourceKey: `model:${model.revision}:${model.htmlFingerprint}`,
    revision: model.revision,
  };
}

export function documentHtmlFingerprint(html: string): string {
  let hash = 0x811c9dc5;
  for (let index = 0; index < html.length; index += 1) {
    hash ^= html.charCodeAt(index);
    hash = Math.imul(hash, 0x01000193);
  }
  return `${html.length.toString(36)}:${(hash >>> 0).toString(36)}`;
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
