import type { DocxIgnorableExtensionPair } from './work-docx-ignorable-extension-preservation';
import {
  alignStaticRunBoundaries,
  createContentControlLimits,
  readStaticBlockControl,
  readStaticContentBlock,
  staticRunSpanKey,
  type BlockContentControlRecord,
  type StaticContentBlockRecord,
  type StaticContentRunRecord,
} from './work-docx-note-comment-content-control-model';
import { createStaticContentControlShell } from './work-docx-note-comment-content-control-properties';
import { preserveDocxNoteCommentRunProperties } from './work-docx-note-comment-run-properties';
import { cloneXmlElement } from './work-docx-settings-xml';

type NoteCommentKind = 'comment' | 'note';

export interface BlockPlanUnit {
  controls: BlockContentControlRecord[];
  generated: StaticContentBlockRecord[];
  kind: 'block';
  source: BlockContentControlRecord;
  targets: Element[];
}

export interface ContentControlAppliedPairs {
  content: DocxIgnorableExtensionPair[];
  runs: DocxIgnorableExtensionPair[];
  structure: DocxIgnorableExtensionPair[];
}

interface PreparedBlock {
  content: DocxIgnorableExtensionPair[];
  element: Element;
  runs: DocxIgnorableExtensionPair[];
}

export function blockCandidates(
  scope: Element,
  generated: readonly StaticContentBlockRecord[],
  source: readonly StaticContentBlockRecord[],
): StaticContentBlockRecord[][] {
  const result: StaticContentBlockRecord[][] = [];
  for (let index = 0; index <= generated.length - source.length; index += 1) {
    const candidate = generated.slice(index, index + source.length);
    if (
      candidate.every(
        (block, offset) => block.fingerprint === source[offset].fingerprint,
      ) &&
      contiguousDirectChildren(
        scope,
        candidate.map((block) => block.element),
      )
    ) {
      result.push(candidate);
    }
  }
  return result;
}

export function applyBlockUnit(
  document: Document,
  unit: BlockPlanUnit,
  assignedIds: ReadonlyMap<Element, string | null>,
  kind: NoteCommentKind,
  applied: ContentControlAppliedPairs,
): void {
  const prepared: PreparedBlock[] = [];
  for (const [index, generated] of unit.generated.entries()) {
    const block = prepareBlock(
      document,
      generated,
      unit.source.blocks[index],
      kind,
    );
    if (!block) return;
    prepared.push(block);
  }
  const shell = createStaticContentControlShell(
    document,
    unit.source,
    assignedIds.get(unit.source.control) ?? null,
    unit.targets[0],
    kind,
  );
  if (!shell) return;
  for (const block of prepared) shell.content.append(block.element);
  const restored = readStaticBlockControl(
    shell.control,
    createContentControlLimits(),
  );
  if (
    !restored ||
    restored.type !== unit.source.type ||
    restored.blocks.length !== unit.source.blocks.length ||
    restored.blocks.some(
      (block, index) =>
        block.fingerprint !== unit.source.blocks[index].fingerprint,
    ) ||
    restored.id !== (assignedIds.get(unit.source.control) ?? null)
  ) {
    return;
  }
  const first = unit.targets[0];
  const parent = first?.parentNode;
  if (!parent || !contiguousDirectChildren(parent as Element, unit.targets)) {
    return;
  }
  parent.insertBefore(shell.control, first);
  for (const target of unit.targets) target.remove();
  applied.structure.push(...shell.pairs);
  for (const block of prepared) {
    applied.content.push(...block.content);
    applied.runs.push(...block.runs);
  }
}

function prepareBlock(
  document: Document,
  generated: StaticContentBlockRecord,
  source: StaticContentBlockRecord,
  kind: NoteCommentKind,
): PreparedBlock | null {
  if (
    generated.kind !== source.kind ||
    generated.fingerprint !== source.fingerprint
  ) {
    return null;
  }
  const clone = cloneXmlElement(document, generated.element);
  let cloneRecord = readStaticContentBlock(
    clone,
    'generated',
    createContentControlLimits(),
  );
  if (
    !cloneRecord ||
    cloneRecord.kind !== source.kind ||
    cloneRecord.fingerprint !== source.fingerprint ||
    cloneRecord.paragraphs.length !== source.paragraphs.length
  ) {
    return null;
  }
  for (const [index, paragraph] of cloneRecord.paragraphs.entries()) {
    const sourceParagraph = source.paragraphs[index];
    const boundaries = sourceParagraph.runs.flatMap((run) => [
      run.start,
      run.end,
    ]);
    if (!alignStaticRunBoundaries(paragraph.element, boundaries)) return null;
  }
  cloneRecord = readStaticContentBlock(
    clone,
    'generated',
    createContentControlLimits(),
  );
  if (
    !cloneRecord ||
    cloneRecord.fingerprint !== source.fingerprint ||
    cloneRecord.paragraphs.length !== source.paragraphs.length
  ) {
    return null;
  }
  const runs: DocxIgnorableExtensionPair[] = [];
  const content: DocxIgnorableExtensionPair[] = [];
  for (const [index, paragraph] of cloneRecord.paragraphs.entries()) {
    const sourceParagraph = source.paragraphs[index];
    const pairs = uniqueRunPairs(paragraph.runs, sourceParagraph.runs);
    if (!pairs || pairs.length !== sourceParagraph.runs.length) return null;
    for (const pair of pairs) {
      preserveDocxNoteCommentRunProperties(
        document,
        pair.generated,
        pair.source,
        kind,
      );
    }
    content.push({
      generated: paragraph.element,
      source: sourceParagraph.element,
    });
    runs.push(...pairs);
  }
  if (cloneRecord.kind === 'table') {
    const generatedScopes = new Map(
      cloneRecord.extensionScopes.map((scope) => [scope.key, scope.element]),
    );
    for (const scope of source.extensionScopes) {
      const target = generatedScopes.get(scope.key);
      if (target?.localName === scope.element.localName) {
        content.push({ generated: target, source: scope.element });
      }
    }
  }
  return { content, element: clone, runs };
}

function uniqueRunPairs(
  generated: readonly StaticContentRunRecord[],
  source: readonly StaticContentRunRecord[],
): DocxIgnorableExtensionPair[] | null {
  const generatedBySpan = groupRuns(generated);
  const sourceBySpan = groupRuns(source);
  const result: DocxIgnorableExtensionPair[] = [];
  for (const [span, sourceRuns] of sourceBySpan) {
    const generatedRuns = generatedBySpan.get(span) ?? [];
    if (sourceRuns.length !== 1 || generatedRuns.length !== 1) return null;
    result.push({
      generated: generatedRuns[0].element,
      source: sourceRuns[0].element,
    });
  }
  return generatedBySpan.size === sourceBySpan.size ? result : null;
}

function groupRuns(
  runs: readonly StaticContentRunRecord[],
): Map<string, StaticContentRunRecord[]> {
  const result = new Map<string, StaticContentRunRecord[]>();
  for (const run of runs) {
    const key = staticRunSpanKey(run);
    const matches = result.get(key) ?? [];
    matches.push(run);
    result.set(key, matches);
  }
  return result;
}

function contiguousDirectChildren(
  parent: Element,
  children: readonly Element[],
): boolean {
  if (
    !children.length ||
    children.some((child) => child.parentElement !== parent)
  ) {
    return false;
  }
  const elements = Array.from(parent.children);
  const first = elements.indexOf(children[0]);
  return (
    first >= 0 &&
    children.every((child, index) => elements[first + index] === child)
  );
}
