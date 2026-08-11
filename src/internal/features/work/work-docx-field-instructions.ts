import { attribute, descendants } from './work-ooxml-package';

export interface DocxFieldOccurrence {
  instruction: string;
  start: Element;
  end: Element;
  syntax: 'simple' | 'complex' | 'orphan';
  complete: boolean;
  nested: boolean;
  containsNested: boolean;
  sameParagraph: boolean;
  inDeletion: boolean;
}

interface ParsedDocxFields {
  occurrences: DocxFieldOccurrence[];
  hasUnmatchedEnd: boolean;
  hasUnclosedBegin: boolean;
}

interface UndecoratedDocxFieldOccurrence {
  instruction: string;
  start: Element;
  end: Element;
  syntax: DocxFieldOccurrence['syntax'];
  closed: boolean;
}

export function docxFieldInstructions(root: ParentNode): string[] {
  return docxFieldOccurrences(root).map((field) => field.instruction);
}

export function docxFieldOccurrences(root: ParentNode): DocxFieldOccurrence[] {
  return parseDocxFields(root).occurrences.filter((field) =>
    Boolean(field.instruction),
  );
}

export function docxFieldOccurrenceIsInlineEditable(
  field: DocxFieldOccurrence,
): boolean {
  return (
    field.syntax !== 'orphan' &&
    field.complete &&
    !field.nested &&
    !field.containsNested &&
    field.sameParagraph &&
    !field.inDeletion
  );
}

export function hasInvalidDocxFieldStructure(root: ParentNode): boolean {
  const parsed = parseDocxFields(root);
  return (
    parsed.hasUnmatchedEnd ||
    parsed.hasUnclosedBegin ||
    parsed.occurrences.some(
      (field) =>
        (field.syntax !== 'orphan' &&
          (!field.complete ||
            field.nested ||
            field.containsNested ||
            !field.sameParagraph ||
            field.inDeletion)) ||
        (field.syntax === 'orphan' && field.inDeletion),
    )
  );
}

function parseDocxFields(root: ParentNode): ParsedDocxFields {
  const elements = Array.from(root.querySelectorAll('*'));
  const fields: UndecoratedDocxFieldOccurrence[] = descendants(
    root,
    'fldSimple',
  ).map((field) => ({
    instruction: attribute(field, 'instr') ?? '',
    start: field,
    end: field,
    syntax: 'simple',
    closed: true,
  }));
  const stack: Array<{
    instruction: string;
    start: Element;
  }> = [];
  let hasUnmatchedEnd = false;
  for (const element of elements) {
    if (element.localName === 'fldChar') {
      const fieldType = attribute(element, 'fldCharType');
      if (fieldType === 'begin') {
        stack.push({ instruction: '', start: element });
      } else if (fieldType === 'end' && stack.length) {
        const field = stack.pop();
        if (field) {
          fields.push({
            ...field,
            end: element,
            syntax: 'complex',
            closed: true,
          });
        }
      } else if (fieldType === 'end') {
        hasUnmatchedEnd = true;
      }
      continue;
    }
    if (element.localName === 'instrText' && stack.length) {
      stack[stack.length - 1].instruction += element.textContent ?? '';
    } else if (element.localName === 'instrText') {
      fields.push({
        instruction: element.textContent ?? '',
        start: element,
        end: element,
        syntax: 'orphan',
        closed: false,
      });
    }
  }
  const hasUnclosedBegin = stack.length > 0;
  fields.push(
    ...stack.map((field) => ({
      ...field,
      end: field.start,
      syntax: 'complex' as const,
      closed: false,
    })),
  );
  return {
    occurrences: decorateFieldOccurrences(elements, fields),
    hasUnmatchedEnd,
    hasUnclosedBegin,
  };
}

export function docxFieldResultText(field: DocxFieldOccurrence): string {
  if (field.start.localName === 'fldSimple') {
    return descendants(field.start, 't')
      .map((element) => element.textContent ?? '')
      .join('');
  }
  const root =
    closestAncestor(field.start, 'p') ??
    field.start.ownerDocument?.documentElement;
  if (!root) return '';
  let inside = false;
  let separated = false;
  const values: string[] = [];
  for (const element of Array.from(root.querySelectorAll('*'))) {
    if (element === field.start) inside = true;
    if (!inside) continue;
    if (
      element.localName === 'fldChar' &&
      attribute(element, 'fldCharType') === 'separate'
    ) {
      separated = true;
    } else if (
      separated &&
      (element.localName === 't' || element.localName === 'delText')
    ) {
      values.push(element.textContent ?? '');
    }
    if (element === field.end) break;
  }
  return values.join('');
}

function closestAncestor(element: Element, localName: string): Element | null {
  let current: Element | null = element;
  while (current) {
    if (current.localName === localName) return current;
    current = current.parentElement;
  }
  return null;
}

function decorateFieldOccurrences(
  elements: Element[],
  fields: UndecoratedDocxFieldOccurrence[],
): DocxFieldOccurrence[] {
  const indexes = new Map(elements.map((element, index) => [element, index]));
  const ranges = fields.map((field) => fieldRange(field, indexes));
  return fields.map((field, index) => {
    const range = ranges[index] ?? { start: -1, end: -1 };
    const startParagraph = closestAncestor(field.start, 'p');
    const endParagraph = closestAncestor(field.end, 'p');
    const instruction = field.instruction.trim();
    return {
      instruction,
      start: field.start,
      end: field.end,
      syntax: field.syntax,
      complete:
        field.syntax !== 'orphan' && field.closed && Boolean(instruction),
      nested: ranges.some(
        (candidate, candidateIndex) =>
          candidateIndex !== index && containsRange(candidate, range),
      ),
      containsNested: ranges.some(
        (candidate, candidateIndex) =>
          candidateIndex !== index && containsRange(range, candidate),
      ),
      sameParagraph: Boolean(startParagraph) && startParagraph === endParagraph,
      inDeletion: elements
        .slice(range.start, range.end + 1)
        .some((element) => Boolean(closestAncestor(element, 'del'))),
    };
  });
}

function fieldRange(
  field: UndecoratedDocxFieldOccurrence,
  indexes: Map<Element, number>,
): { start: number; end: number } {
  const start = indexes.get(field.start) ?? -1;
  if (field.syntax !== 'simple') {
    return { start, end: indexes.get(field.end) ?? start };
  }
  const children = Array.from(field.start.querySelectorAll('*'));
  const last = children.at(-1) ?? field.start;
  return { start, end: indexes.get(last) ?? start };
}

function containsRange(
  outer: { start: number; end: number },
  inner: { start: number; end: number },
): boolean {
  return (
    outer.start >= 0 &&
    inner.start >= 0 &&
    outer.start < inner.start &&
    outer.end >= inner.end
  );
}
