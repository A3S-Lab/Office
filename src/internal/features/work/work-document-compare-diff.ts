export type DocumentSequenceDiff<T> =
  | { kind: 'equal'; left: T[]; right: T[] }
  | { kind: 'delete'; left: T[]; right: [] }
  | { kind: 'insert'; left: []; right: T[] };

export type DocumentSequenceAlignment<T> =
  | { kind: 'equal' | 'substitute'; left: T; right: T }
  | { kind: 'delete'; left: T; right: null }
  | { kind: 'insert'; left: null; right: T };

const GAP_COST = 1;
const ALIGNMENT_EPSILON = 1e-9;

/**
 * Computes a deterministic LCS diff while bounding the allocated matrix.
 * Callers can fall back to one replacement when the requested matrix is too
 * large instead of allowing adversarial documents to allocate without limit.
 */
export function boundedDocumentSequenceDiff<T>(
  left: readonly T[],
  right: readonly T[],
  equal: (left: T, right: T) => boolean,
  maximumCells: number,
): DocumentSequenceDiff<T>[] | null {
  if ((left.length + 1) * (right.length + 1) > maximumCells) return null;
  const columns = right.length + 1;
  const CellArray =
    Math.max(left.length, right.length) <= 65_535 ? Uint16Array : Uint32Array;
  const matrix = new CellArray((left.length + 1) * columns);
  for (let leftIndex = left.length - 1; leftIndex >= 0; leftIndex -= 1) {
    for (let rightIndex = right.length - 1; rightIndex >= 0; rightIndex -= 1) {
      const offset = leftIndex * columns + rightIndex;
      matrix[offset] = equal(left[leftIndex] as T, right[rightIndex] as T)
        ? 1 + matrix[(leftIndex + 1) * columns + rightIndex + 1]
        : Math.max(
            matrix[(leftIndex + 1) * columns + rightIndex],
            matrix[leftIndex * columns + rightIndex + 1],
          );
    }
  }

  const steps: DocumentSequenceDiff<T>[] = [];
  let leftIndex = 0;
  let rightIndex = 0;
  while (leftIndex < left.length || rightIndex < right.length) {
    if (
      leftIndex < left.length &&
      rightIndex < right.length &&
      equal(left[leftIndex] as T, right[rightIndex] as T)
    ) {
      appendSequenceDiff(
        steps,
        'equal',
        left[leftIndex] as T,
        right[rightIndex] as T,
      );
      leftIndex += 1;
      rightIndex += 1;
      continue;
    }
    const deleteScore =
      leftIndex < left.length
        ? matrix[(leftIndex + 1) * columns + rightIndex]
        : -1;
    const insertScore =
      rightIndex < right.length
        ? matrix[leftIndex * columns + rightIndex + 1]
        : -1;
    if (leftIndex < left.length && deleteScore >= insertScore) {
      appendSequenceDiff(steps, 'delete', left[leftIndex] as T, null);
      leftIndex += 1;
    } else {
      appendSequenceDiff(steps, 'insert', null, right[rightIndex] as T);
      rightIndex += 1;
    }
  }
  return steps;
}

/**
 * Aligns document blocks with a weighted substitution cost. A substitution
 * cost below two pairs related blocks; a cost of two or more prefers an
 * explicit deletion plus insertion.
 */
export function boundedDocumentSequenceAlignment<T>(
  left: readonly T[],
  right: readonly T[],
  substitutionCost: (left: T, right: T) => number,
  maximumCells: number,
): DocumentSequenceAlignment<T>[] | null {
  if ((left.length + 1) * (right.length + 1) > maximumCells) return null;
  const columns = right.length + 1;
  const matrix = new Float64Array((left.length + 1) * columns);
  for (let leftIndex = 1; leftIndex <= left.length; leftIndex += 1) {
    matrix[leftIndex * columns] = leftIndex * GAP_COST;
  }
  for (let rightIndex = 1; rightIndex <= right.length; rightIndex += 1) {
    matrix[rightIndex] = rightIndex * GAP_COST;
  }
  for (let leftIndex = 1; leftIndex <= left.length; leftIndex += 1) {
    for (let rightIndex = 1; rightIndex <= right.length; rightIndex += 1) {
      const substitution =
        matrix[(leftIndex - 1) * columns + rightIndex - 1] +
        substitutionCost(left[leftIndex - 1] as T, right[rightIndex - 1] as T);
      const deletion =
        matrix[(leftIndex - 1) * columns + rightIndex] + GAP_COST;
      const insertion = matrix[leftIndex * columns + rightIndex - 1] + GAP_COST;
      matrix[leftIndex * columns + rightIndex] = Math.min(
        substitution,
        deletion,
        insertion,
      );
    }
  }

  const reversed: DocumentSequenceAlignment<T>[] = [];
  let leftIndex = left.length;
  let rightIndex = right.length;
  while (leftIndex > 0 || rightIndex > 0) {
    const current = matrix[leftIndex * columns + rightIndex];
    if (leftIndex > 0 && rightIndex > 0) {
      const cost = substitutionCost(
        left[leftIndex - 1] as T,
        right[rightIndex - 1] as T,
      );
      const substitution =
        matrix[(leftIndex - 1) * columns + rightIndex - 1] + cost;
      if (approximatelyEqual(current, substitution)) {
        reversed.push({
          kind: cost === 0 ? 'equal' : 'substitute',
          left: left[leftIndex - 1] as T,
          right: right[rightIndex - 1] as T,
        });
        leftIndex -= 1;
        rightIndex -= 1;
        continue;
      }
    }
    if (
      leftIndex > 0 &&
      approximatelyEqual(
        current,
        matrix[(leftIndex - 1) * columns + rightIndex] + GAP_COST,
      )
    ) {
      reversed.push({
        kind: 'delete',
        left: left[leftIndex - 1] as T,
        right: null,
      });
      leftIndex -= 1;
      continue;
    }
    reversed.push({
      kind: 'insert',
      left: null,
      right: right[rightIndex - 1] as T,
    });
    rightIndex -= 1;
  }
  return reversed.reverse();
}

function appendSequenceDiff<T>(
  changes: DocumentSequenceDiff<T>[],
  kind: DocumentSequenceDiff<T>['kind'],
  left: T | null,
  right: T | null,
): void {
  const previous = changes.at(-1);
  if (previous?.kind === kind) {
    if (previous.kind === 'equal') {
      previous.left.push(left as T);
      previous.right.push(right as T);
    } else if (previous.kind === 'delete') {
      previous.left.push(left as T);
    } else {
      previous.right.push(right as T);
    }
    return;
  }
  if (kind === 'equal') {
    changes.push({ kind, left: [left as T], right: [right as T] });
  } else if (kind === 'delete') {
    changes.push({ kind, left: [left as T], right: [] });
  } else {
    changes.push({ kind, left: [], right: [right as T] });
  }
}

function approximatelyEqual(left: number, right: number): boolean {
  return Math.abs(left - right) <= ALIGNMENT_EPSILON;
}
