import { remapPresentationGroupPaths } from './work-presentation-groups';
import { remapWorkSlideAnimations } from './work-presentation-animation';
import { createWorkId } from './work-templates';
import type {
  WorkSlide,
  WorkSlideAnimation,
  WorkSlideElement,
} from './work-types';

export const PRESENTATION_OBJECT_OFFSET_STEP = 5;
const PRESENTATION_CLIPBOARD_MAX_OFFSET = 20;

export type WorkPresentationClipboardPayload =
  | {
      kind: 'elements';
      elements: WorkSlideElement[];
      animations?: WorkSlideAnimation[];
    }
  | { kind: 'slide'; slide: WorkSlide };

export interface WorkPresentationClipboardRead {
  payload: WorkPresentationClipboardPayload;
  offset: number;
}

let clipboard: {
  payload: WorkPresentationClipboardPayload;
  pasteCount: number;
} | null = null;

export interface WorkPresentationElementPaste {
  animations?: WorkSlideAnimation[];
  elements: WorkSlideElement[];
}

export function copyPresentationElements(
  elements: readonly WorkSlideElement[],
  animations: readonly WorkSlideAnimation[] = [],
): void {
  if (!elements.length) return;
  const elementIds = new Set(elements.map((element) => element.id));
  const selectedAnimations = animations.filter((animation) =>
    elementIds.has(animation.elementId),
  );
  clipboard = {
    payload: {
      kind: 'elements',
      elements: structuredCopy([...elements]),
      ...(selectedAnimations.length
        ? { animations: structuredCopy(selectedAnimations) }
        : {}),
    },
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
    offset: Math.min(
      clipboard.pasteCount * PRESENTATION_OBJECT_OFFSET_STEP,
      PRESENTATION_CLIPBOARD_MAX_OFFSET,
    ),
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
  return clonePresentationElementsAndAnimationsForPaste([element], [], offset)
    .elements[0];
}

export function clonePresentationElementsForPaste(
  elements: readonly WorkSlideElement[],
  offset: number,
): WorkSlideElement[] {
  return clonePresentationElementsAndAnimationsForPaste(elements, [], offset)
    .elements;
}

export function clonePresentationElementsAndAnimationsForPaste(
  elements: readonly WorkSlideElement[],
  animations: readonly WorkSlideAnimation[],
  offset: number,
): WorkPresentationElementPaste {
  if (!elements.length) return { elements: [] };
  const right = Math.max(
    ...elements.map((element) => element.x + element.width),
  );
  const bottom = Math.max(
    ...elements.map((element) => element.y + element.height),
  );
  const offsetX = clamp(offset, 0, Math.max(0, 100 - right));
  const offsetY = clamp(offset, 0, Math.max(0, 100 - bottom));
  const elementIds = new Map<string, string>();
  const copies = remapPresentationGroupPaths(structuredCopy(elements)).map(
    (element) => {
      const id = createWorkId('element');
      elementIds.set(element.id, id);
      return {
        ...element,
        id,
        x: element.x + offsetX,
        y: element.y + offsetY,
        placeholder: undefined,
      };
    },
  );
  return {
    elements: copies,
    animations: remapWorkSlideAnimations(animations, elementIds),
  };
}

export function clonePresentationSlideForPaste(
  slide: WorkSlide,
  existingSlideNames: readonly string[] = [],
): WorkSlide {
  const copy = structuredCopy(slide);
  const elementIds = new Map<string, string>();
  const elements = remapPresentationGroupPaths(copy.elements).map((element) => {
    const id = createWorkId('element');
    elementIds.set(element.id, id);
    return { ...element, id };
  });
  return {
    ...copy,
    id: createWorkId('slide'),
    name: nextPresentationSlideCopyName(slide.name, existingSlideNames),
    elements,
    animations: remapWorkSlideAnimations(copy.animations, elementIds),
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
