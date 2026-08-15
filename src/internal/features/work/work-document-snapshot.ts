import { documentModelForContent } from './work-document-model';
import type { WorkDocumentContent } from './work-types';

export const WORK_DOCUMENT_SNAPSHOT_SCHEMA =
  'a3s.office.document.snapshot' as const;
export const WORK_DOCUMENT_SNAPSHOT_VERSION = 1 as const;
export const WORK_DOCUMENT_SNAPSHOT_MEDIA_TYPE =
  'application/vnd.a3s.office.document-snapshot+json;version=1' as const;

const MAX_DOCUMENT_SNAPSHOT_BYTES = 16 * 1024 * 1024;
const MAX_DOCUMENT_SNAPSHOT_DEPTH = 512;
const MAX_DOCUMENT_SNAPSHOT_VALUES = 2_000_000;

export interface WorkDocumentSnapshot {
  schema: typeof WORK_DOCUMENT_SNAPSHOT_SCHEMA;
  version: typeof WORK_DOCUMENT_SNAPSHOT_VERSION;
  content: WorkDocumentContent;
}

/**
 * Encode the complete controlled document value as deterministic JSON.
 *
 * The snapshot requires a synchronized structured model. Optional properties
 * whose value is `undefined` are omitted because undefined is outside the JSON
 * data model; every JSON-defined field is retained exactly.
 */
export function encodeWorkDocumentSnapshot(
  content: WorkDocumentContent,
): string {
  assertStructuredDocumentContent(content);
  let encoded: string;
  try {
    encoded = canonicalJson({
      schema: WORK_DOCUMENT_SNAPSHOT_SCHEMA,
      version: WORK_DOCUMENT_SNAPSHOT_VERSION,
      content,
    } satisfies WorkDocumentSnapshot);
  } catch (error) {
    throw new Error(
      'A3S Office document snapshots must contain only JSON-compatible values.',
      { cause: error },
    );
  }
  assertSnapshotSize(encoded);
  return encoded;
}

/** Decode and validate one versioned structured-document snapshot. */
export function decodeWorkDocumentSnapshot(
  encoded: string,
): WorkDocumentContent {
  if (typeof encoded !== 'string') {
    throw new TypeError('An A3S Office document snapshot must be a string.');
  }
  assertSnapshotSize(encoded);
  let candidate: unknown;
  try {
    candidate = JSON.parse(encoded);
  } catch (error) {
    throw new Error('The A3S Office document snapshot is not valid JSON.', {
      cause: error,
    });
  }
  if (!isRecord(candidate)) {
    throw new Error('The A3S Office document snapshot envelope is invalid.');
  }
  if (candidate.schema !== WORK_DOCUMENT_SNAPSHOT_SCHEMA) {
    throw new Error('The A3S Office document snapshot schema is unsupported.');
  }
  if (candidate.version !== WORK_DOCUMENT_SNAPSHOT_VERSION) {
    throw new Error(
      `A3S Office document snapshot version ${String(candidate.version)} is unsupported.`,
    );
  }
  const envelopeKeys = Object.keys(candidate);
  if (
    envelopeKeys.length !== 3 ||
    !envelopeKeys.includes('schema') ||
    !envelopeKeys.includes('version') ||
    !envelopeKeys.includes('content')
  ) {
    throw new Error('The A3S Office document snapshot envelope is invalid.');
  }
  assertStructuredDocumentContent(candidate.content);
  return candidate.content;
}

function assertStructuredDocumentContent(
  candidate: unknown,
): asserts candidate is WorkDocumentContent {
  if (
    !isRecord(candidate) ||
    candidate.type !== 'document' ||
    typeof candidate.html !== 'string' ||
    (candidate.pageSize !== 'a4' && candidate.pageSize !== 'letter')
  ) {
    throw new Error(
      'An A3S Office document snapshot must contain a valid document value.',
    );
  }
  if (
    !candidate.model ||
    !documentModelForContent(candidate as unknown as WorkDocumentContent)
  ) {
    throw new Error(
      'An A3S Office document snapshot requires a synchronized structured model.',
    );
  }
}

function assertSnapshotSize(encoded: string): void {
  if (
    new TextEncoder().encode(encoded).byteLength > MAX_DOCUMENT_SNAPSHOT_BYTES
  ) {
    throw new Error(
      `The A3S Office document snapshot exceeds ${MAX_DOCUMENT_SNAPSHOT_BYTES} bytes.`,
    );
  }
}

function canonicalJson(value: unknown): string {
  const ancestors = new Set<object>();
  const budget = { values: 0 };
  return encodeCanonicalValue(value, ancestors, budget, 0);
}

function encodeCanonicalValue(
  value: unknown,
  ancestors: Set<object>,
  budget: { values: number },
  depth: number,
): string {
  budget.values += 1;
  if (
    budget.values > MAX_DOCUMENT_SNAPSHOT_VALUES ||
    depth > MAX_DOCUMENT_SNAPSHOT_DEPTH
  ) {
    throw new TypeError('The JSON value graph exceeds the snapshot limits.');
  }
  if (
    value === null ||
    typeof value === 'string' ||
    typeof value === 'boolean'
  ) {
    return JSON.stringify(value);
  }
  if (typeof value === 'number') {
    if (!Number.isFinite(value)) {
      throw new TypeError('Non-finite numbers are not JSON-compatible.');
    }
    return JSON.stringify(value);
  }
  if (typeof value !== 'object' || ancestors.has(value)) {
    throw new TypeError('The value is not an acyclic JSON value.');
  }

  ancestors.add(value);
  try {
    if (Array.isArray(value)) {
      return `[${value
        .map((item) => {
          if (item === undefined) {
            throw new TypeError('Undefined array entries are not supported.');
          }
          return encodeCanonicalValue(item, ancestors, budget, depth + 1);
        })
        .join(',')}]`;
    }

    const prototype = Object.getPrototypeOf(value);
    if (prototype !== Object.prototype && prototype !== null) {
      throw new TypeError('Only plain JSON objects are supported.');
    }
    const symbols = Object.getOwnPropertySymbols(value);
    if (symbols.length) {
      throw new TypeError('Symbol properties are not JSON-compatible.');
    }
    const entries: string[] = [];
    for (const key of Object.keys(value).sort()) {
      const descriptor = Object.getOwnPropertyDescriptor(value, key);
      if (!descriptor || descriptor.get || descriptor.set) {
        throw new TypeError('Accessor properties are not JSON-compatible.');
      }
      if (descriptor.value === undefined) continue;
      entries.push(
        `${JSON.stringify(key)}:${encodeCanonicalValue(
          descriptor.value,
          ancestors,
          budget,
          depth + 1,
        )}`,
      );
    }
    return `{${entries.join(',')}}`;
  } finally {
    ancestors.delete(value);
  }
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}
