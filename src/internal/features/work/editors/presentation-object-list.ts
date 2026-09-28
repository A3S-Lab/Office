import type { WorkSlideElement } from '../work-types';

export type PresentationObjectListItem = {
  id: string;
  label: string;
  type: WorkSlideElement['type'];
};

const TYPE_LABELS: Record<WorkSlideElement['type'], string> = {
  text: 'Text',
  shape: 'Shape',
  image: 'Image',
  table: 'Table',
  chart: 'Chart',
  line: 'Line',
};

/**
 * Stable accessible name for a slide element. Prefer author alt text, then
 * visible text, then a typed fallback so the object list never exposes bare ids.
 */
export function presentationObjectLabel(
  element: WorkSlideElement,
  options?: { typeLabels?: Partial<Record<WorkSlideElement['type'], string>> },
): string {
  const alt = element.altText?.trim();
  if (alt) return alt;
  const text = element.text?.trim();
  if (text) return truncateLabel(text);
  const prompt = element.placeholder?.prompt?.trim();
  if (prompt) return truncateLabel(prompt);
  const typeLabel =
    options?.typeLabels?.[element.type] ?? TYPE_LABELS[element.type];
  return typeLabel;
}

export function presentationObjectListItems(
  elements: readonly WorkSlideElement[],
  options?: { typeLabels?: Partial<Record<WorkSlideElement['type'], string>> },
): PresentationObjectListItem[] {
  return elements.map((element) => ({
    id: element.id,
    type: element.type,
    label: presentationObjectLabel(element, options),
  }));
}

function truncateLabel(value: string, max = 48): string {
  if (value.length <= max) return value;
  return `${value.slice(0, max - 1)}…`;
}
