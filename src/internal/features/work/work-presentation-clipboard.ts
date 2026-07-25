import { createWorkId } from './work-templates';
import { remapPresentationGroupPaths } from './work-presentation-groups';
import type { WorkSlide, WorkSlideElement } from './work-types';

export type WorkPresentationClipboardPayload =
  | { kind: 'element'; element: WorkSlideElement }
  | { kind: 'elements'; elements: WorkSlideElement[] }
  | { kind: 'slide'; slide: WorkSlide };

export interface WorkPresentationClipboardRead {
  payload: WorkPresentationClipboardPayload;
  offset: number;
}

let clipboard: {
  payload: WorkPresentationClipboardPayload;
  pasteCount: number;
} | null = null;

export function copyPresentationElement(element: WorkSlideElement): void {
  clipboard = {
    payload: { kind: 'element', element: structuredCopy(element) },
    pasteCount: 0,
  };
  writeSystemClipboardText(presentationElementPlainText(element));
}

export function copyPresentationElements(
  elements: readonly WorkSlideElement[],
): void {
  if (!elements.length) return;
  clipboard = {
    payload: { kind: 'elements', elements: structuredCopy([...elements]) },
    pasteCount: 0,
  };
  writeSystemClipboardText(
    elements.map(presentationElementPlainText).filter(Boolean).join('\n'),
  );
}

export function copyPresentationSlide(slide: WorkSlide): void {
  clipboard = {
    payload: { kind: 'slide', slide: structuredCopy(slide) },
    pasteCount: 0,
  };
  writeSystemClipboardText(presentationSlidePlainText(slide));
}

export function takePresentationClipboard(): WorkPresentationClipboardRead | null {
  if (!clipboard) return null;
  clipboard.pasteCount += 1;
  return {
    payload: structuredCopy(clipboard.payload),
    offset: Math.min(clipboard.pasteCount * 2, 12),
  };
}

export function hasPresentationClipboard(): boolean {
  return clipboard !== null;
}

export function clearPresentationClipboard(): void {
  clipboard = null;
}

export function clonePresentationElementForPaste(
  element: WorkSlideElement,
  offset: number,
): WorkSlideElement {
  const copy = remapPresentationGroupPaths([structuredCopy(element)])[0];
  return {
    ...copy,
    id: createWorkId('element'),
    x: clamp(copy.x + offset, 0, Math.max(0, 100 - copy.width)),
    y: clamp(copy.y + offset, 0, Math.max(0, 100 - copy.height)),
    placeholder: undefined,
  };
}

export function clonePresentationElementsForPaste(
  elements: readonly WorkSlideElement[],
  offset: number,
): WorkSlideElement[] {
  if (!elements.length) return [];
  const right = Math.max(
    ...elements.map((element) => element.x + element.width),
  );
  const bottom = Math.max(
    ...elements.map((element) => element.y + element.height),
  );
  const offsetX = clamp(offset, 0, Math.max(0, 100 - right));
  const offsetY = clamp(offset, 0, Math.max(0, 100 - bottom));
  return remapPresentationGroupPaths(structuredCopy(elements)).map(
    (element) => ({
      ...element,
      id: createWorkId('element'),
      x: element.x + offsetX,
      y: element.y + offsetY,
      placeholder: undefined,
    }),
  );
}

export function clonePresentationSlideForPaste(
  slide: WorkSlide,
  existingSlideNames: readonly string[] = [],
): WorkSlide {
  const copy = structuredCopy(slide);
  return {
    ...copy,
    id: createWorkId('slide'),
    name: nextPresentationSlideCopyName(slide.name, existingSlideNames),
    elements: remapPresentationGroupPaths(copy.elements).map((element) => ({
      ...element,
      id: createWorkId('element'),
    })),
    comments: copy.comments?.map((comment) => ({
      ...comment,
      id: createWorkId('slide-comment'),
    })),
  };
}

export function nextPresentationSlideCopyName(
  sourceName: string,
  existingSlideNames: readonly string[],
): string {
  const baseName = presentationSlideCopyBaseName(sourceName);
  let highestOrdinal = 0;
  for (const existingName of [sourceName, ...existingSlideNames]) {
    if (existingName === `${baseName} 副本`) {
      highestOrdinal = Math.max(highestOrdinal, 1);
      continue;
    }
    const prefix = `${baseName} 副本 `;
    if (!existingName.startsWith(prefix)) continue;
    const ordinal = Number(existingName.slice(prefix.length));
    if (Number.isInteger(ordinal) && ordinal > 1) {
      highestOrdinal = Math.max(highestOrdinal, ordinal);
    }
  }
  return highestOrdinal === 0
    ? `${baseName} 副本`
    : `${baseName} 副本 ${highestOrdinal + 1}`;
}

function presentationSlideCopyBaseName(sourceName: string): string {
  let baseName = sourceName.trim() || '幻灯片';
  while (true) {
    const numberedCopy = baseName.match(/^(.*) 副本 \d+$/);
    if (numberedCopy?.[1]) {
      baseName = numberedCopy[1].trim();
      continue;
    }
    const copy = baseName.match(/^(.*) 副本$/);
    if (!copy?.[1]) return baseName;
    baseName = copy[1].trim();
  }
}

function presentationElementPlainText(element: WorkSlideElement): string {
  if (element.table)
    return element.table.rows.map((row) => row.join('\t')).join('\n');
  if (element.textRuns?.length)
    return element.textRuns.map((run) => run.text).join('');
  if (element.text.trim()) return element.text;
  if (element.chart?.title?.trim()) return element.chart.title;
  return (
    element.altText?.trim() || (element.type === 'image' ? '图片' : '演示元素')
  );
}

function presentationSlidePlainText(slide: WorkSlide): string {
  return [
    slide.name,
    ...slide.elements.map(presentationElementPlainText),
    slide.notes ?? '',
  ]
    .map((value) => value.trim())
    .filter(Boolean)
    .join('\n');
}

function writeSystemClipboardText(value: string): void {
  if (typeof navigator === 'undefined' || !navigator.clipboard?.writeText)
    return;
  void navigator.clipboard.writeText(value).catch(() => undefined);
}

function clamp(value: number, minimum: number, maximum: number): number {
  return Math.min(Math.max(value, minimum), maximum);
}

function structuredCopy<T>(value: T): T {
  if (typeof structuredClone === 'function') return structuredClone(value);
  return JSON.parse(JSON.stringify(value)) as T;
}
