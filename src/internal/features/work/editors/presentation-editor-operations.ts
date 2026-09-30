import {
  officeMessage,
  resolveOfficeMessages,
} from '../../../i18n/office-locale';
import { withPresentationDesign } from '../work-presentation-layouts';
import { createWorkId } from '../work-templates';
import type {
  WorkPresentationContent,
  WorkSlide,
  WorkSlideElement,
} from '../work-types';
import { readOfficeFileAsDataUrl } from './office-file-data';
import type { PresentationDesignMode } from './presentation-editor-types';

export function updatePresentationElements(
  content: WorkPresentationContent,
  mode: PresentationDesignMode,
  targetId: string,
  update: (elements: WorkSlideElement[]) => WorkSlideElement[],
  onChange: (content: WorkPresentationContent) => void,
) {
  if (mode === 'slide') {
    updateSlide(
      content,
      targetId,
      (slide) => ({ ...slide, elements: update(slide.elements) }),
      onChange,
    );
    return;
  }
  const normalized = withPresentationDesign(content);
  if (mode === 'layout') {
    onChange({
      ...normalized,
      layouts: normalized.layouts?.map((layout) =>
        layout.id === targetId
          ? { ...layout, elements: update(structuredCopy(layout.elements)) }
          : layout,
      ),
    });
    return;
  }
  onChange({
    ...normalized,
    masters: normalized.masters?.map((master) =>
      master.id === targetId
        ? { ...master, elements: update(structuredCopy(master.elements)) }
        : master,
    ),
  });
}

export function updateSlide(
  content: WorkPresentationContent,
  slideId: string,
  update: (slide: WorkSlide) => WorkSlide,
  onChange: (content: WorkPresentationContent) => void,
) {
  onChange({
    ...content,
    slides: content.slides.map((slide) =>
      slide.id === slideId ? update(structuredCopy(slide)) : slide,
    ),
  });
}

/** `insertionIndex` is the gap before that slide, counted before the source is removed. */
export function movePresentationSlide<T>(
  slides: readonly T[],
  fromIndex: number,
  insertionIndex: number,
): T[] {
  if (
    fromIndex < 0 ||
    insertionIndex < 0 ||
    fromIndex >= slides.length ||
    insertionIndex > slides.length
  ) {
    return slides.slice();
  }
  const next = slides.slice();
  const [moved] = next.splice(fromIndex, 1);
  if (moved === undefined) return slides.slice();
  const insertAt =
    insertionIndex > fromIndex ? insertionIndex - 1 : insertionIndex;
  next.splice(insertAt, 0, moved);
  return next;
}

export function newSlide(number: number): WorkSlide {
  return {
    id: createWorkId('slide'),
    name: officeMessage(resolveOfficeMessages(), 'presentation.ops.slideName', {
      number: String(number),
    }),
    background: '#ffffff',
    elements: [
      {
        id: createWorkId('element'),
        type: 'text',
        x: 9,
        y: 12,
        width: 82,
        height: 16,
        text: '',
        fontSize: 30,
        color: '#172033',
        fill: 'transparent',
        bold: true,
        align: 'left',
        placeholder: {
          key: 'title',
          type: 'title',
          prompt: officeMessage(
            resolveOfficeMessages(),
            'presentation.ops.titlePrompt',
          ),
        },
      },
    ],
  };
}

export function newPresentationElement(
  type: 'shape' | 'text',
): WorkSlideElement {
  return {
    id: createWorkId('element'),
    type,
    x: 30,
    y: 34,
    width: 40,
    height: type === 'text' ? 14 : 20,
    text: '',
    fontSize: type === 'text' ? 24 : 14,
    color: '#172033',
    fill: type === 'text' ? 'transparent' : '#dce6fb',
    bold: false,
    align: 'center',
    radius: type === 'shape' ? 3 : 0,
  };
}

export function newPresentationTableElement(
  requestedRows = 3,
  requestedColumns = 3,
): WorkSlideElement {
  const rows = normalizeTableDimension(requestedRows);
  const columns = normalizeTableDimension(requestedColumns);
  return {
    id: createWorkId('element'),
    type: 'table',
    x: 15,
    y: 24,
    width: 70,
    height: Math.min(68, Math.max(24, 12 + rows * 10)),
    text: '',
    fontSize: 14,
    color: '#172033',
    fill: '#ffffff',
    bold: false,
    align: 'left',
    borderColor: '#cbd2de',
    borderWidth: 1,
    table: {
      headerRows: 1,
      rows: Array.from({ length: rows }, (_, rowIndex) =>
        Array.from({ length: columns }, (_, columnIndex) =>
          rowIndex === 0
            ? officeMessage(
                resolveOfficeMessages(),
                'presentation.ops.tableHeader',
                {
                  column: String(columnIndex + 1),
                },
              )
            : officeMessage(
                resolveOfficeMessages(),
                'presentation.ops.tableBody',
              ),
        ),
      ),
    },
  };
}

function normalizeTableDimension(value: number): number {
  if (!Number.isFinite(value)) return 1;
  return Math.min(20, Math.max(1, Math.floor(value)));
}

export async function newPresentationImageElement(
  file: File,
): Promise<WorkSlideElement> {
  return {
    id: createWorkId('element'),
    type: 'image',
    x: 20,
    y: 20,
    width: 60,
    height: 55,
    text: '',
    fontSize: 12,
    color: '#172033',
    fill: 'transparent',
    bold: false,
    align: 'center',
    altText: file.name,
    image: {
      dataUrl: await readOfficeFileAsDataUrl(file),
      contentType: file.type || 'application/octet-stream',
      name: file.name,
    },
  };
}

export function clamp(value: number, minimum: number, maximum: number): number {
  return Math.min(Math.max(value, minimum), maximum);
}

export function structuredCopy<T>(value: T): T {
  if (typeof structuredClone === 'function') return structuredClone(value);
  return JSON.parse(JSON.stringify(value)) as T;
}
