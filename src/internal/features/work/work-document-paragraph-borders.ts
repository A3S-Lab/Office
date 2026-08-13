import { normalizeCssColor } from './work-document-paragraph-shading';
import {
  type DocxThemeColorReference,
  parseDocxThemeReference,
  serializeDocxThemeReference,
} from './work-docx-theme-reference';

export const DOCUMENT_PARAGRAPH_BORDERS_ATTRIBUTE =
  'data-office-paragraph-borders';

export const DOCUMENT_PARAGRAPH_BORDER_EDGES = [
  'top',
  'left',
  'bottom',
  'right',
  'between',
  'bar',
] as const;

export type DocumentParagraphBorderEdge =
  (typeof DOCUMENT_PARAGRAPH_BORDER_EDGES)[number];

export const DOCUMENT_PARAGRAPH_BORDER_STYLES = [
  'nil',
  'none',
  'single',
  'thick',
  'double',
  'dotted',
  'dashed',
  'dotDash',
  'dotDotDash',
  'triple',
  'thinThickSmallGap',
  'thickThinSmallGap',
  'thinThickThinSmallGap',
  'thinThickMediumGap',
  'thickThinMediumGap',
  'thinThickThinMediumGap',
  'thinThickLargeGap',
  'thickThinLargeGap',
  'thinThickThinLargeGap',
  'wave',
  'doubleWave',
  'dashSmallGap',
  'dashDotStroked',
  'threeDEmboss',
  'threeDEngrave',
  'outset',
  'inset',
  'apples',
  'archedScallops',
  'babyPacifier',
  'babyRattle',
  'balloons3Colors',
  'balloonsHotAir',
  'basicBlackDashes',
  'basicBlackDots',
  'basicBlackSquares',
  'basicThinLines',
  'basicWhiteDashes',
  'basicWhiteDots',
  'basicWhiteSquares',
  'basicWideInline',
  'basicWideMidline',
  'basicWideOutline',
  'bats',
  'birds',
  'birdsFlight',
  'cabins',
  'cakeSlice',
  'candyCorn',
  'celticKnotwork',
  'certificateBanner',
  'chainLink',
  'champagneBottle',
  'checkedBarBlack',
  'checkedBarColor',
  'checkered',
  'christmasTree',
  'circlesLines',
  'circlesRectangles',
  'classicalWave',
  'clocks',
  'compass',
  'confetti',
  'confettiGrays',
  'confettiOutline',
  'confettiStreamers',
  'confettiWhite',
  'cornerTriangles',
  'couponCutoutDashes',
  'couponCutoutDots',
  'crazyMaze',
  'creaturesButterfly',
  'creaturesFish',
  'creaturesInsects',
  'creaturesLadyBug',
  'crossStitch',
  'cup',
  'decoArch',
  'decoArchColor',
  'decoBlocks',
  'diamondsGray',
  'doubleD',
  'doubleDiamonds',
  'earth1',
  'earth2',
  'eclipsingSquares1',
  'eclipsingSquares2',
  'eggsBlack',
  'fans',
  'film',
  'firecrackers',
  'flowersBlockPrint',
  'flowersDaisies',
  'flowersModern1',
  'flowersModern2',
  'flowersPansy',
  'flowersRedRose',
  'flowersRoses',
  'flowersTeacup',
  'flowersTiny',
  'gems',
  'gingerbreadMan',
  'gradient',
  'handmade1',
  'handmade2',
  'heartBalloon',
  'heartGray',
  'hearts',
  'heebieJeebies',
  'holly',
  'houseFunky',
  'hypnotic',
  'iceCreamCones',
  'lightBulb',
  'lightning1',
  'lightning2',
  'mapPins',
  'mapleLeaf',
  'mapleMuffins',
  'marquee',
  'marqueeToothed',
  'moons',
  'mosaic',
  'musicNotes',
  'northwest',
  'ovals',
  'packages',
  'palmsBlack',
  'palmsColor',
  'paperClips',
  'papyrus',
  'partyFavor',
  'partyGlass',
  'pencils',
  'people',
  'peopleWaving',
  'peopleHats',
  'poinsettias',
  'postageStamp',
  'pumpkin1',
  'pushPinNote2',
  'pushPinNote1',
  'pyramids',
  'pyramidsAbove',
  'quadrants',
  'rings',
  'safari',
  'sawtooth',
  'sawtoothGray',
  'scaredCat',
  'seattle',
  'shadowedSquares',
  'sharksTeeth',
  'shorebirdTracks',
  'skyrocket',
  'snowflakeFancy',
  'snowflakes',
  'sombrero',
  'southwest',
  'stars',
  'starsTop',
  'stars3d',
  'starsBlack',
  'starsShadowed',
  'sun',
  'swirligig',
  'tornPaper',
  'tornPaperBlack',
  'trees',
  'triangleParty',
  'triangles',
  'tribal1',
  'tribal2',
  'tribal3',
  'tribal4',
  'tribal5',
  'tribal6',
  'triangle1',
  'triangle2',
  'triangleCircle1',
  'triangleCircle2',
  'shapes1',
  'shapes2',
  'twistedLines1',
  'twistedLines2',
  'vine',
  'waveline',
  'weavingAngles',
  'weavingBraid',
  'weavingRibbon',
  'weavingStrips',
  'whiteFlowers',
  'woodwork',
  'xIllusions',
  'zanyTriangles',
  'zigZag',
  'zigZagStitch',
] as const;

export type DocumentParagraphBorderStyle =
  (typeof DOCUMENT_PARAGRAPH_BORDER_STYLES)[number];

export interface DocumentParagraphBorderColor {
  value: 'auto' | `#${string}`;
  theme?: DocxThemeColorReference;
}

export interface DocumentParagraphBorder {
  style: DocumentParagraphBorderStyle;
  color?: DocumentParagraphBorderColor;
  size?: number;
  space?: number;
  shadow?: boolean;
  frame?: boolean;
}

export type DocumentParagraphBorders = Partial<
  Record<DocumentParagraphBorderEdge, DocumentParagraphBorder>
>;

const BORDER_STYLE_SET = new Set<DocumentParagraphBorderStyle>(
  DOCUMENT_PARAGRAPH_BORDER_STYLES,
);
const BORDER_EDGE_SET = new Set<string>(DOCUMENT_PARAGRAPH_BORDER_EDGES);
const BORDER_PROPERTY_SET = new Set([
  'style',
  'color',
  'size',
  'space',
  'shadow',
  'frame',
]);
const LINE_BORDER_STYLES = new Set<DocumentParagraphBorderStyle>(
  DOCUMENT_PARAGRAPH_BORDER_STYLES.slice(0, 27),
);
const DASHED_BORDER_STYLES = new Set<DocumentParagraphBorderStyle>([
  'dashed',
  'dashSmallGap',
  'dashDotStroked',
  'dotDash',
  'dotDotDash',
]);
const DOUBLE_BORDER_STYLES = new Set<DocumentParagraphBorderStyle>([
  'double',
  'triple',
  'thinThickSmallGap',
  'thickThinSmallGap',
  'thinThickThinSmallGap',
  'thinThickMediumGap',
  'thickThinMediumGap',
  'thinThickThinMediumGap',
  'thinThickLargeGap',
  'thickThinLargeGap',
  'thinThickThinLargeGap',
  'doubleWave',
]);
const MAX_SERIALIZED_PARAGRAPH_BORDERS = 32_768;
const POINTS_TO_PIXELS = 96 / 72;

export function normalizeDocumentParagraphBorders(
  source: unknown,
): DocumentParagraphBorders | null {
  if (!source || typeof source !== 'object' || Array.isArray(source))
    return null;
  const record = source as Record<string, unknown>;
  if (Object.keys(record).some((key) => !BORDER_EDGE_SET.has(key))) return null;
  const borders: DocumentParagraphBorders = {};
  for (const edge of DOCUMENT_PARAGRAPH_BORDER_EDGES) {
    if (record[edge] === undefined) continue;
    const border = normalizeDocumentParagraphBorder(record[edge]);
    if (!border) return null;
    borders[edge] = border;
  }
  return Object.keys(borders).length ? borders : null;
}

export function normalizeDocumentParagraphBorder(
  source: unknown,
): DocumentParagraphBorder | null {
  if (!source || typeof source !== 'object' || Array.isArray(source))
    return null;
  const record = source as Record<string, unknown>;
  if (Object.keys(record).some((key) => !BORDER_PROPERTY_SET.has(key)))
    return null;
  const style = record.style;
  if (
    typeof style !== 'string' ||
    !BORDER_STYLE_SET.has(style as DocumentParagraphBorderStyle)
  )
    return null;
  const normalizedStyle = style as DocumentParagraphBorderStyle;
  const color = normalizeBorderColor(record.color);
  if (record.color !== undefined && !color) return null;
  const size = optionalInteger(record.size);
  if (
    size === null ||
    (size !== undefined && !validBorderSize(normalizedStyle, size))
  ) {
    return null;
  }
  const space = optionalInteger(record.space);
  if (space === null || (space !== undefined && (space < 0 || space > 31)))
    return null;
  const shadow = optionalBoolean(record.shadow);
  const frame = optionalBoolean(record.frame);
  if (shadow === null || frame === null) return null;
  return {
    style: normalizedStyle,
    ...(color ? { color } : {}),
    ...(size !== undefined ? { size } : {}),
    ...(space !== undefined ? { space } : {}),
    ...(shadow !== undefined ? { shadow } : {}),
    ...(frame !== undefined ? { frame } : {}),
  };
}

export function parseDocumentParagraphBorders(
  source: unknown,
): DocumentParagraphBorders | null {
  if (typeof source !== 'string')
    return normalizeDocumentParagraphBorders(source);
  if (!source.trim() || source.length > MAX_SERIALIZED_PARAGRAPH_BORDERS)
    return null;
  try {
    return normalizeDocumentParagraphBorders(JSON.parse(source));
  } catch {
    return null;
  }
}

export function serializeDocumentParagraphBorders(
  source: unknown,
): string | undefined {
  const borders = normalizeDocumentParagraphBorders(source);
  if (!borders) return undefined;
  return JSON.stringify(
    Object.fromEntries(
      DOCUMENT_PARAGRAPH_BORDER_EDGES.flatMap((edge) => {
        const border = borders[edge];
        return border ? [[edge, serializedBorder(border)]] : [];
      }),
    ),
  );
}

export function parseDocumentParagraphBordersElement(
  element: HTMLElement,
): DocumentParagraphBorders | null {
  const semantic = parseDocumentParagraphBorders(
    element.getAttribute(DOCUMENT_PARAGRAPH_BORDERS_ATTRIBUTE),
  );
  if (!semantic) return paragraphBordersFromCss(element);
  const edited: DocumentParagraphBorders = { ...semantic };
  for (const edge of ['top', 'left', 'bottom', 'right'] as const) {
    const border = semantic[edge];
    const css = cssBorder(element, edge);
    if (!css || (border && sameBorderPresentation(border, css))) continue;
    if (css.style === 'none') {
      edited[edge] = { style: 'nil' };
      continue;
    }
    edited[edge] = {
      ...(border ?? { style: 'single' as const }),
      style: cssStyleToBorderStyle(css.style),
      color: { value: css.color },
      size: cssWidthToEighthPoints(css.width),
    };
  }
  return normalizeDocumentParagraphBorders(edited);
}

export function documentParagraphBordersDomAttributes(
  source: unknown,
): Record<string, string> {
  const borders = normalizeDocumentParagraphBorders(source);
  const serialized = serializeDocumentParagraphBorders(borders);
  if (!borders || !serialized) return {};
  const styles: string[] = [];
  const shadows: string[] = [];
  for (const edge of ['top', 'left', 'bottom', 'right'] as const) {
    const border = borders[edge];
    if (!border) continue;
    const presentation = borderPresentation(border);
    styles.push(
      `border-${edge}: ${formatPixels(presentation.width)}px ${presentation.style} ${presentation.color}`,
    );
    if (border.space)
      styles.push(
        `padding-${edge}: ${formatPixels(border.space * POINTS_TO_PIXELS)}px`,
      );
    if (border.shadow && presentation.width > 0) {
      shadows.push(`2px 2px 0 ${presentation.color}`);
    }
  }
  const between = borders.between ? borderPresentation(borders.between) : null;
  if (between && between.width > 0) {
    shadows.push(
      `inset 0 -${formatPixels(between.width)}px 0 ${between.color}`,
    );
  }
  const bar = borders.bar ? borderPresentation(borders.bar) : null;
  if (bar && bar.width > 0) {
    shadows.push(`inset ${formatPixels(bar.width)}px 0 0 ${bar.color}`);
  }
  if (shadows.length) styles.push(`box-shadow: ${shadows.join(', ')}`);
  return {
    [DOCUMENT_PARAGRAPH_BORDERS_ATTRIBUTE]: serialized,
    ...(styles.length ? { style: styles.join('; ') } : {}),
  };
}

export function isDocumentParagraphArtBorderStyle(
  style: DocumentParagraphBorderStyle,
): boolean {
  return !LINE_BORDER_STYLES.has(style);
}

function normalizeBorderColor(
  source: unknown,
): DocumentParagraphBorderColor | null {
  if (!source || typeof source !== 'object' || Array.isArray(source))
    return null;
  const record = source as Record<string, unknown>;
  if (Object.keys(record).some((key) => key !== 'value' && key !== 'theme'))
    return null;
  const theme = parseDocxThemeReference(
    typeof record.theme === 'string'
      ? record.theme
      : record.theme
        ? JSON.stringify(record.theme)
        : undefined,
  );
  const direct =
    record.value === 'auto'
      ? ('auto' as const)
      : typeof record.value === 'string'
        ? normalizeCssColor(record.value)
        : null;
  const resolved = direct ?? theme?.resolved ?? null;
  if (!resolved || resolved === 'transparent') return null;
  if (
    theme &&
    resolved !== theme.resolved &&
    !(
      resolved === 'auto' &&
      theme.theme === 'none' &&
      theme.resolved === '#000000'
    )
  ) {
    return null;
  }
  return {
    value: resolved as DocumentParagraphBorderColor['value'],
    ...(theme ? { theme } : {}),
  };
}

function serializedBorder(
  border: DocumentParagraphBorder,
): Record<string, unknown> {
  const color = border.color ? serializedBorderColor(border.color) : undefined;
  return {
    style: border.style,
    ...(color ? { color } : {}),
    ...(border.size !== undefined ? { size: border.size } : {}),
    ...(border.space !== undefined ? { space: border.space } : {}),
    ...(border.shadow !== undefined ? { shadow: border.shadow } : {}),
    ...(border.frame !== undefined ? { frame: border.frame } : {}),
  };
}

function serializedBorderColor(
  color: DocumentParagraphBorderColor,
): Record<string, unknown> {
  const theme = serializeDocxThemeReference(color.theme ?? null);
  return {
    value: color.value,
    ...(theme ? { theme: JSON.parse(theme) as Record<string, unknown> } : {}),
  };
}

function optionalInteger(value: unknown): number | null | undefined {
  if (value === undefined) return undefined;
  return typeof value === 'number' && Number.isSafeInteger(value)
    ? value
    : null;
}

function optionalBoolean(value: unknown): boolean | null | undefined {
  if (value === undefined) return undefined;
  return typeof value === 'boolean' ? value : null;
}

function validBorderSize(
  style: DocumentParagraphBorderStyle,
  size: number,
): boolean {
  if (style === 'nil' || style === 'none') return size >= 0 && size <= 96;
  return isDocumentParagraphArtBorderStyle(style)
    ? size >= 1 && size <= 31
    : size >= 2 && size <= 96;
}

interface BorderPresentation {
  color: `#${string}` | 'transparent';
  style: 'none' | 'solid' | 'dashed' | 'dotted' | 'double' | 'inset' | 'outset';
  width: number;
}

function borderPresentation(
  border: DocumentParagraphBorder,
): BorderPresentation {
  if (border.style === 'nil' || border.style === 'none' || !border.size) {
    return { color: 'transparent', style: 'none', width: 0 };
  }
  const width = Math.min(
    16,
    isDocumentParagraphArtBorderStyle(border.style)
      ? border.size * POINTS_TO_PIXELS
      : border.size / 6,
  );
  return {
    color:
      !border.color || border.color.value === 'auto'
        ? '#000000'
        : border.color.value,
    style: borderStyleToCssStyle(border.style),
    width,
  };
}

function borderStyleToCssStyle(
  style: DocumentParagraphBorderStyle,
): BorderPresentation['style'] {
  if (style === 'nil' || style === 'none') return 'none';
  if (style === 'dotted') return 'dotted';
  if (DASHED_BORDER_STYLES.has(style)) return 'dashed';
  if (DOUBLE_BORDER_STYLES.has(style)) return 'double';
  if (style === 'inset' || style === 'threeDEngrave') return 'inset';
  if (style === 'outset' || style === 'threeDEmboss') return 'outset';
  return 'solid';
}

function paragraphBordersFromCss(
  element: HTMLElement,
): DocumentParagraphBorders | null {
  const borders: DocumentParagraphBorders = {};
  for (const edge of ['top', 'left', 'bottom', 'right'] as const) {
    const css = cssBorder(element, edge);
    if (!css || css.style === 'none') continue;
    borders[edge] = {
      style: cssStyleToBorderStyle(css.style),
      color: { value: css.color },
      size: cssWidthToEighthPoints(css.width),
    };
  }
  return normalizeDocumentParagraphBorders(borders);
}

function cssBorder(
  element: HTMLElement,
  edge: 'top' | 'left' | 'bottom' | 'right',
): {
  color: `#${string}`;
  style: string;
  width: number;
} | null {
  const style = element.style.getPropertyValue(`border-${edge}-style`).trim();
  const width = Number.parseFloat(
    element.style.getPropertyValue(`border-${edge}-width`),
  );
  const color = normalizeCssColor(
    element.style.getPropertyValue(`border-${edge}-color`),
  );
  if (!style) return null;
  if (style === 'none' || style === 'hidden') {
    return { color: '#000000', style: 'none', width: 0 };
  }
  return color && color !== 'transparent' && Number.isFinite(width) && width > 0
    ? { color, style, width }
    : null;
}

function sameBorderPresentation(
  border: DocumentParagraphBorder,
  css: { color: `#${string}`; style: string; width: number },
): boolean {
  const expected = borderPresentation(border);
  if (expected.style === 'none') {
    return css.style === 'none' && css.width === 0;
  }
  return (
    expected.color === css.color &&
    expected.style === css.style &&
    Math.abs(expected.width - css.width) < 0.01
  );
}

function cssStyleToBorderStyle(style: string): DocumentParagraphBorderStyle {
  if (style === 'double') return 'double';
  if (style === 'dashed') return 'dashed';
  if (style === 'dotted') return 'dotted';
  if (style === 'inset' || style === 'groove') return 'inset';
  if (style === 'outset' || style === 'ridge') return 'outset';
  return 'single';
}

function cssWidthToEighthPoints(width: number): number {
  return Math.max(2, Math.min(96, Math.round(width * 6)));
}

function formatPixels(value: number): string {
  return Number(value.toFixed(3)).toString();
}
