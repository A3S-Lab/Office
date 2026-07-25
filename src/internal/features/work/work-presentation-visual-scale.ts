import type { WorkSlideElement } from './work-types';

export function scaledPresentationVisuals(
  element: WorkSlideElement,
  scale: number,
): Pick<WorkSlideElement, 'fontSize'> &
  Partial<Pick<WorkSlideElement, 'borderWidth' | 'textRuns'>> {
  const fontSize = presentationElementScalesTypography(element)
    ? scaledPresentationFontSize(element.fontSize, scale)
    : element.fontSize;
  const borderWidth =
    element.borderWidth === undefined
      ? undefined
      : scaledPresentationBorderWidth(element.borderWidth, scale);
  const textRuns = scaledPresentationTextRuns(element, scale);
  return {
    fontSize,
    ...(borderWidth !== undefined ? { borderWidth } : {}),
    ...(textRuns !== element.textRuns ? { textRuns } : {}),
  };
}

function presentationElementScalesTypography(
  element: WorkSlideElement,
): boolean {
  return (
    element.type === 'text' ||
    element.type === 'shape' ||
    element.type === 'table' ||
    Boolean(element.text || element.textRuns?.length)
  );
}

function scaledPresentationTextRuns(
  element: WorkSlideElement,
  scale: number,
): WorkSlideElement['textRuns'] {
  if (!element.textRuns?.some((run) => run.fontSize !== undefined)) {
    return element.textRuns;
  }
  return element.textRuns.map((run) =>
    run.fontSize === undefined
      ? run
      : { ...run, fontSize: scaledPresentationFontSize(run.fontSize, scale) },
  );
}

function scaledPresentationFontSize(value: number, scale: number): number {
  return roundedPresentationMetric(clamp(value * scale, 1, 400));
}

function scaledPresentationBorderWidth(value: number, scale: number): number {
  if (value === 0) return 0;
  return roundedPresentationMetric(Math.max(0.1, value * scale));
}

function roundedPresentationMetric(value: number): number {
  return Math.round(value * 1_000) / 1_000;
}

function clamp(value: number, minimum: number, maximum: number): number {
  return Math.min(maximum, Math.max(minimum, value));
}
