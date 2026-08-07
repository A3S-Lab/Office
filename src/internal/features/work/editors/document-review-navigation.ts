import type { WorkDocumentChange } from '../work-document-changes';

interface DocumentSelectionRange {
  from: number;
  to: number;
}

export function selectedDocumentChangeIndex(
  changes: readonly WorkDocumentChange[],
  selection: DocumentSelectionRange,
): number | null {
  const index = changes.findIndex((change) =>
    selection.from === selection.to
      ? change.from <= selection.from && selection.from <= change.to
      : change.from < selection.to && change.to > selection.from,
  );
  return index >= 0 ? index : null;
}

export function adjacentDocumentChangeIndex(
  changes: readonly WorkDocumentChange[],
  selection: DocumentSelectionRange,
  direction: -1 | 1,
): number | null {
  const selectedIndex = selectedDocumentChangeIndex(changes, selection);
  if (selectedIndex !== null) {
    const adjacentIndex = selectedIndex + direction;
    return adjacentIndex >= 0 && adjacentIndex < changes.length
      ? adjacentIndex
      : null;
  }

  if (direction > 0) {
    const index = changes.findIndex((change) => change.from >= selection.to);
    return index >= 0 ? index : null;
  }

  for (let index = changes.length - 1; index >= 0; index -= 1) {
    if (changes[index].to <= selection.from) return index;
  }
  return null;
}

export function actionableDocumentChangeIndex(
  changes: readonly WorkDocumentChange[],
  selection: DocumentSelectionRange,
): number | null {
  return (
    selectedDocumentChangeIndex(changes, selection) ??
    adjacentDocumentChangeIndex(changes, selection, 1) ??
    adjacentDocumentChangeIndex(changes, selection, -1)
  );
}
