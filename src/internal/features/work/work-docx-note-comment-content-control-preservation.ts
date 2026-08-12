import {
  mergeDocxIgnorableExtensionsAtPairs,
  type DocxExtensionDocumentRole,
  type DocxIgnorableExtensionPair,
} from './work-docx-ignorable-extension-preservation';
import {
  alignStaticRunBoundaries,
  createContentControlLimits,
  directWordControls,
  directWordParagraphs,
  readStaticBlockControl,
  readStaticInlineParagraph,
  staticRunSpanKey,
  type BlockContentControlRecord,
  type InlineContentControlRecord,
  type StaticContentParagraphRecord,
  type StaticContentRunRecord,
} from './work-docx-note-comment-content-control-model';
import {
  createStaticContentControlShell,
  readStaticContentControlDefinition,
} from './work-docx-note-comment-content-control-properties';
import {
  CONTENT_CONTROL_WORD_2012_NAMESPACE,
  ensureIgnorableContentControlNamespace,
} from './work-docx-note-comment-content-control-xml';
import { isKnownOoxmlNamespace } from './work-docx-note-comment-hyperlink-content';
import { preserveDocxNoteCommentRunProperties } from './work-docx-note-comment-run-properties';
import { descendants } from './work-ooxml-package';

type NoteCommentKind = 'comment' | 'note';

interface InlinePlanUnit {
  controls: InlineContentControlRecord[];
  generated: StaticContentParagraphRecord;
  kind: 'inline';
  source: StaticContentParagraphRecord;
  targets: Element[];
}

interface BlockPlanUnit {
  controls: BlockContentControlRecord[];
  generated: StaticContentParagraphRecord[];
  kind: 'block';
  source: BlockContentControlRecord;
  targets: Element[];
}

type PlanUnit = BlockPlanUnit | InlinePlanUnit;

interface AppliedPairs {
  paragraphs: DocxIgnorableExtensionPair[];
  runs: DocxIgnorableExtensionPair[];
  structure: DocxIgnorableExtensionPair[];
}

export function preserveDocxNoteCommentContentControls(
  generatedDocument: Document,
  sourceDocument: Document,
  scopes: readonly DocxIgnorableExtensionPair[],
  kind: NoteCommentKind,
): void {
  if (!scopes.length) return;
  const limits = createContentControlLimits();
  const planned = scopes.flatMap((scope) => planScope(scope, limits));
  const units = failClosedUnits(planned);
  const assignedIds = assignContentControlIds(generatedDocument, units);
  const applied: AppliedPairs = { paragraphs: [], runs: [], structure: [] };
  for (const unit of units.filter((item) => item.kind === 'inline')) {
    applyInlineUnit(
      generatedDocument,
      unit as InlinePlanUnit,
      assignedIds,
      kind,
      applied,
    );
  }
  for (const unit of units.filter((item) => item.kind === 'block').reverse()) {
    applyBlockUnit(
      generatedDocument,
      unit as BlockPlanUnit,
      assignedIds,
      kind,
      applied,
    );
  }
  mergeExtensions(
    generatedDocument,
    sourceDocument,
    [...applied.structure, ...applied.paragraphs],
    0,
  );
  mergeExtensions(generatedDocument, sourceDocument, applied.runs, 4);
  if (usesWord2012ContentControlExtensions(generatedDocument)) {
    ensureIgnorableContentControlNamespace(
      generatedDocument.documentElement,
      CONTENT_CONTROL_WORD_2012_NAMESPACE,
      'w15',
    );
  }
}

function planScope(
  scope: DocxIgnorableExtensionPair,
  limits: ReturnType<typeof createContentControlLimits>,
): PlanUnit[] {
  const generatedParagraphs = directWordParagraphs(scope.generated).flatMap(
    (paragraph) => {
      const record = readStaticInlineParagraph(paragraph, 'generated', limits);
      return record ? [record] : [];
    },
  );
  const units: PlanUnit[] = [];
  for (const paragraph of directWordParagraphs(scope.source)) {
    if (!directWordControls(paragraph).length) continue;
    const source = readStaticInlineParagraph(paragraph, 'source', limits);
    if (!source?.controls.length) continue;
    const matches = generatedParagraphs.filter(
      (generated) => generated.text === source.text,
    );
    if (matches.length !== 1) continue;
    units.push({
      controls: source.controls,
      generated: matches[0],
      kind: 'inline',
      source,
      targets: [matches[0].element],
    });
  }
  for (const control of directWordControls(scope.source)) {
    const source = readStaticBlockControl(control, limits);
    if (!source) continue;
    const matches = blockCandidates(
      scope.generated,
      generatedParagraphs,
      source.paragraphs.map((paragraph) => paragraph.text),
    );
    if (matches.length !== 1) continue;
    units.push({
      controls: [source],
      generated: matches[0],
      kind: 'block',
      source,
      targets: matches[0].map((paragraph) => paragraph.element),
    });
  }
  return units;
}

function blockCandidates(
  scope: Element,
  generated: readonly StaticContentParagraphRecord[],
  sourceText: readonly string[],
): StaticContentParagraphRecord[][] {
  const result: StaticContentParagraphRecord[][] = [];
  for (
    let index = 0;
    index <= generated.length - sourceText.length;
    index += 1
  ) {
    const candidate = generated.slice(index, index + sourceText.length);
    if (
      candidate.every(
        (paragraph, offset) => paragraph.text === sourceText[offset],
      ) &&
      contiguousDirectChildren(
        scope,
        candidate.map((paragraph) => paragraph.element),
      )
    ) {
      result.push(candidate);
    }
  }
  return result;
}

function failClosedUnits(units: readonly PlanUnit[]): PlanUnit[] {
  const invalid = new Set<PlanUnit>();
  const unitsById = new Map<string, PlanUnit[]>();
  const unitsByTarget = new Map<Element, PlanUnit[]>();
  for (const unit of units) {
    for (const control of unit.controls) {
      if (!control.id) continue;
      const matches = unitsById.get(control.id) ?? [];
      matches.push(unit);
      unitsById.set(control.id, matches);
    }
    for (const target of unit.targets) {
      const matches = unitsByTarget.get(target) ?? [];
      matches.push(unit);
      unitsByTarget.set(target, matches);
    }
  }
  for (const matches of [...unitsById.values(), ...unitsByTarget.values()]) {
    if (matches.length > 1) {
      for (const unit of matches) invalid.add(unit);
    }
  }
  return units.filter((unit) => !invalid.has(unit));
}

function assignContentControlIds(
  generatedDocument: Document,
  units: readonly PlanUnit[],
): Map<Element, string | null> {
  const used = new Set<string>();
  for (const element of descendants(generatedDocument, 'sdt')) {
    const definition = readStaticContentControlDefinition(element);
    if (definition?.id) used.add(definition.id);
  }
  const result = new Map<Element, string | null>();
  for (const unit of units) {
    for (const control of unit.controls) {
      if (!control.id) {
        result.set(control.control, null);
      } else if (!used.has(control.id)) {
        used.add(control.id);
        result.set(control.control, control.id);
      }
    }
  }
  for (const unit of units) {
    for (const control of unit.controls) {
      if (!control.id || result.has(control.control)) continue;
      const id = nextContentControlId(used);
      used.add(id);
      result.set(control.control, id);
    }
  }
  return result;
}

function applyInlineUnit(
  document: Document,
  unit: InlinePlanUnit,
  assignedIds: ReadonlyMap<Element, string | null>,
  kind: NoteCommentKind,
  applied: AppliedPairs,
): void {
  const working = unit.generated.element.cloneNode(true) as Element;
  const boundaries = unit.source.runs.flatMap((run) => [run.start, run.end]);
  if (!alignStaticRunBoundaries(working, boundaries)) return;
  const generated = readStaticInlineParagraph(
    working,
    'generated',
    createContentControlLimits(),
  );
  if (!generated || generated.text !== unit.source.text) return;
  const runPairs = uniqueRunPairs(generated.runs, unit.source.runs);
  if (!runPairs || runPairs.length !== unit.source.runs.length) return;
  for (const pair of runPairs) {
    preserveDocxNoteCommentRunProperties(
      document,
      pair.generated,
      pair.source,
      kind,
    );
  }
  const generatedBySpan = new Map(
    generated.runs.map((run) => [staticRunSpanKey(run), run.element]),
  );
  const structure: DocxIgnorableExtensionPair[] = [];
  for (const sourceControl of [...unit.controls].sort(
    (left, right) => right.start - left.start,
  )) {
    const selected = sourceControl.runs.map((run) =>
      generatedBySpan.get(staticRunSpanKey(run)),
    );
    if (
      selected.some((run) => !run) ||
      !contiguousDirectChildren(working, selected as Element[])
    ) {
      return;
    }
    const shell = createStaticContentControlShell(
      document,
      sourceControl,
      assignedIds.get(sourceControl.control) ?? null,
      working,
      kind,
    );
    const first = selected[0];
    if (!shell || !first?.parentNode) return;
    first.parentNode.insertBefore(shell.control, first);
    for (const run of selected as Element[]) shell.content.append(run);
    structure.push(...shell.pairs);
  }
  const restored = readStaticInlineParagraph(
    working,
    'source',
    createContentControlLimits(),
  );
  if (!stableInlineControls(restored, unit.controls, assignedIds)) return;
  replaceChildren(unit.generated.element, Array.from(working.childNodes));
  applied.structure.push(...structure);
  applied.paragraphs.push({
    generated: unit.generated.element,
    source: unit.source.element,
  });
  applied.runs.push(...runPairs);
}

function applyBlockUnit(
  document: Document,
  unit: BlockPlanUnit,
  assignedIds: ReadonlyMap<Element, string | null>,
  kind: NoteCommentKind,
  applied: AppliedPairs,
): void {
  const paragraphPairs: DocxIgnorableExtensionPair[] = [];
  const runPairs: DocxIgnorableExtensionPair[] = [];
  const clones: Element[] = [];
  for (const [index, generatedParagraph] of unit.generated.entries()) {
    const sourceParagraph = unit.source.paragraphs[index];
    const clone = generatedParagraph.element.cloneNode(true) as Element;
    const boundaries = sourceParagraph.runs.flatMap((run) => [
      run.start,
      run.end,
    ]);
    if (!alignStaticRunBoundaries(clone, boundaries)) return;
    const generated = readStaticInlineParagraph(
      clone,
      'generated',
      createContentControlLimits(),
    );
    if (!generated || generated.text !== sourceParagraph.text) return;
    const pairs = uniqueRunPairs(generated.runs, sourceParagraph.runs);
    if (!pairs || pairs.length !== sourceParagraph.runs.length) return;
    for (const pair of pairs) {
      preserveDocxNoteCommentRunProperties(
        document,
        pair.generated,
        pair.source,
        kind,
      );
    }
    clones.push(clone);
    paragraphPairs.push({ generated: clone, source: sourceParagraph.element });
    runPairs.push(...pairs);
  }
  const shell = createStaticContentControlShell(
    document,
    unit.source,
    assignedIds.get(unit.source.control) ?? null,
    unit.targets[0],
    kind,
  );
  if (!shell) return;
  for (const paragraph of clones) shell.content.append(paragraph);
  const restored = readStaticBlockControl(
    shell.control,
    createContentControlLimits(),
  );
  if (
    !restored ||
    restored.paragraphs.map((paragraph) => paragraph.text).join('\u0000') !==
      unit.source.paragraphs
        .map((paragraph) => paragraph.text)
        .join('\u0000') ||
    restored.id !== (assignedIds.get(unit.source.control) ?? null)
  ) {
    return;
  }
  const first = unit.targets[0];
  const parent = first?.parentNode;
  if (!parent || !contiguousDirectChildren(parent as Element, unit.targets))
    return;
  parent.insertBefore(shell.control, first);
  for (const target of unit.targets) target.remove();
  applied.structure.push(...shell.pairs);
  applied.paragraphs.push(...paragraphPairs);
  applied.runs.push(...runPairs);
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

function stableInlineControls(
  restored: StaticContentParagraphRecord | null,
  source: readonly InlineContentControlRecord[],
  assignedIds: ReadonlyMap<Element, string | null>,
): boolean {
  return Boolean(
    restored &&
      restored.controls.length === source.length &&
      source.every((control) =>
        restored.controls.some(
          (item) =>
            item.start === control.start &&
            item.end === control.end &&
            item.text === control.text &&
            item.type === control.type &&
            item.id === (assignedIds.get(control.control) ?? null),
        ),
      ),
  );
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

function replaceChildren(element: Element, children: readonly Node[]): void {
  while (element.firstChild) element.removeChild(element.firstChild);
  for (const child of children) element.append(child);
}

function nextContentControlId(used: ReadonlySet<string>): string {
  let id = 1;
  while (used.has(String(id))) id += 1;
  return String(id);
}

function mergeExtensions(
  generated: Document,
  source: Document,
  pairs: readonly DocxIgnorableExtensionPair[],
  maximumDepth: number,
): void {
  if (!pairs.length) return;
  mergeDocxIgnorableExtensionsAtPairs(generated, source, pairs, {
    semanticKey: contentSemanticKey,
    isAdditionalSemanticNamespace: isKnownOoxmlNamespace,
    allowExtensionNamespace: (namespace) => !isKnownOoxmlNamespace(namespace),
    allowMatchedElementMerge: (_generated, _source, depth) =>
      depth <= maximumDepth,
  });
}

function contentSemanticKey(
  element: Element,
  _role: DocxExtensionDocumentRole,
): string {
  return isKnownOoxmlNamespace(element.namespaceURI ?? '')
    ? `{semantic}${element.localName}`
    : `{${element.namespaceURI ?? ''}}${element.localName}`;
}

function usesWord2012ContentControlExtensions(document: Document): boolean {
  return descendants(document, 'sdt').some((control) =>
    [control, ...Array.from(control.querySelectorAll('*'))].some(
      (element) => element.namespaceURI === CONTENT_CONTROL_WORD_2012_NAMESPACE,
    ),
  );
}
