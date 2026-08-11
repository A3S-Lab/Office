export type WorkDocumentImageWrapSide =
  | 'bothSides'
  | 'left'
  | 'right'
  | 'largest';

export interface WorkDocumentImageWrapPoint {
  x: number;
  y: number;
}

export interface WorkDocumentImageWrapContour {
  edited: boolean;
  points: WorkDocumentImageWrapPoint[];
}

const DEFAULT_WRAP_SIDE: WorkDocumentImageWrapSide = 'bothSides';
const WRAP_COORDINATE_SIZE = 21_600;
const MAXIMUM_WRAP_COORDINATE = 2_147_483_647;
const MAXIMUM_WRAP_POINTS = 512;

export function normalizeDocumentImageWrapSide(
  value: unknown,
): WorkDocumentImageWrapSide {
  return value === 'left' || value === 'right' || value === 'largest'
    ? value
    : DEFAULT_WRAP_SIDE;
}

export function normalizeDocumentImageWrapContour(
  value: Partial<
    Record<'edited' | 'points' | 'wrapPolygon' | 'wrapPolygonEdited', unknown>
  >,
): WorkDocumentImageWrapContour | null {
  const points = normalizeDocumentImageWrapPoints(
    value.points ?? value.wrapPolygon,
  );
  if (!points) return null;
  return {
    edited: normalizeBoolean(value.edited ?? value.wrapPolygonEdited),
    points,
  };
}

export function defaultDocumentImageWrapContour(): WorkDocumentImageWrapContour {
  return {
    edited: false,
    points: [
      { x: 0, y: 0 },
      { x: 0, y: WRAP_COORDINATE_SIZE },
      { x: WRAP_COORDINATE_SIZE, y: WRAP_COORDINATE_SIZE },
      { x: WRAP_COORDINATE_SIZE, y: 0 },
      { x: 0, y: 0 },
    ],
  };
}

export function documentImageWrapContourFromElement(
  element: Element,
): WorkDocumentImageWrapContour | null {
  return normalizeDocumentImageWrapContour({
    wrapPolygon: element.getAttribute('data-office-image-wrap-polygon'),
    wrapPolygonEdited: element.getAttribute(
      'data-office-image-wrap-polygon-edited',
    ),
  });
}

export function serializeDocumentImageWrapPoints(
  points: readonly WorkDocumentImageWrapPoint[],
): string {
  return points.map((point) => `${point.x},${point.y}`).join(';');
}

export function applyDocumentImageWrapContourToElement(
  element: HTMLElement,
  contour: WorkDocumentImageWrapContour | null,
): void {
  if (!contour) {
    delete element.dataset.officeImageWrapPolygon;
    delete element.dataset.officeImageWrapPolygonEdited;
    element.style.removeProperty('--work-document-image-wrap-contour');
    return;
  }
  element.dataset.officeImageWrapPolygon = serializeDocumentImageWrapPoints(
    contour.points,
  );
  element.dataset.officeImageWrapPolygonEdited = String(contour.edited);
  element.style.setProperty(
    '--work-document-image-wrap-contour',
    documentImageWrapContourCss(contour),
  );
}

export function documentImageWrapContourCss(
  contour: WorkDocumentImageWrapContour,
): string {
  return `polygon(${contour.points
    .map(
      (point) => `${formatPercentage(point.x)}% ${formatPercentage(point.y)}%`,
    )
    .join(', ')})`;
}

function normalizeDocumentImageWrapPoints(
  value: unknown,
): WorkDocumentImageWrapPoint[] | null {
  const source =
    typeof value === 'string'
      ? value
          .split(';')
          .filter(Boolean)
          .map((point) => {
            const [x, y, ...rest] = point.split(',');
            return rest.length || x === undefined || y === undefined
              ? null
              : { x, y };
          })
      : Array.isArray(value)
        ? value
        : [];
  if (source.length < 3 || source.length > MAXIMUM_WRAP_POINTS) return null;
  const points: WorkDocumentImageWrapPoint[] = [];
  for (const point of source) {
    if (!point || typeof point !== 'object') return null;
    const candidate = point as { x?: unknown; y?: unknown };
    const x = wrapCoordinate(candidate.x);
    const y = wrapCoordinate(candidate.y);
    if (x === null || y === null) return null;
    points.push({ x, y });
  }
  return points;
}

function wrapCoordinate(value: unknown): number | null {
  if (value === null || value === undefined || value === '') return null;
  const number = typeof value === 'number' ? value : Number(value);
  if (!Number.isFinite(number)) return null;
  const rounded = Math.round(number);
  return Math.abs(rounded) <= MAXIMUM_WRAP_COORDINATE ? rounded : null;
}

function normalizeBoolean(value: unknown): boolean {
  return value === true || value === 1 || value === '1' || value === 'true';
}

function formatPercentage(value: number): string {
  return Number(((value / WRAP_COORDINATE_SIZE) * 100).toFixed(4)).toString();
}
