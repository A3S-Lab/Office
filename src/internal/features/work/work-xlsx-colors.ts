import {
  attribute,
  directChild,
  directChildren,
  firstDescendant,
} from './work-ooxml-package';

export interface XlsxColorResolver {
  indexed: readonly string[];
  theme: readonly string[];
}

const defaultThemeColors = [
  'ffffff',
  '000000',
  'e7e6e6',
  '44546a',
  '4472c4',
  'ed7d31',
  'a5a5a5',
  'ffc000',
  '5b9bd5',
  '70ad47',
  '0563c1',
  '954f72',
] as const;

const themeColorNames = [
  'lt1',
  'dk1',
  'lt2',
  'dk2',
  'accent1',
  'accent2',
  'accent3',
  'accent4',
  'accent5',
  'accent6',
  'hlink',
  'folHlink',
] as const;

const defaultIndexedColors = [
  '000000',
  'ffffff',
  'ff0000',
  '00ff00',
  '0000ff',
  'ffff00',
  'ff00ff',
  '00ffff',
  '000000',
  'ffffff',
  'ff0000',
  '00ff00',
  '0000ff',
  'ffff00',
  'ff00ff',
  '00ffff',
  '800000',
  '008000',
  '000080',
  '808000',
  '800080',
  '008080',
  'c0c0c0',
  '808080',
  '9999ff',
  '993366',
  'ffffcc',
  'ccffff',
  '660066',
  'ff8080',
  '0066cc',
  'ccccff',
  '000080',
  'ff00ff',
  'ffff00',
  '00ffff',
  '800080',
  '800000',
  '008080',
  '0000ff',
  '00ccff',
  'ccffff',
  'ccffcc',
  'ffff99',
  '99ccff',
  'ff99cc',
  'cc99ff',
  'ffcc99',
  '3366ff',
  '33cccc',
  '99cc00',
  'ffcc00',
  'ff9900',
  'ff6600',
  '666699',
  '969696',
  '003366',
  '339966',
  '003300',
  '333300',
  '993300',
  '993366',
  '333399',
  '333333',
] as const;

export function createXlsxColorResolver(
  styles: Document | null,
  theme: Document | null,
): XlsxColorResolver {
  return {
    indexed: styles ? readIndexedColors(styles) : defaultIndexedColors,
    theme: readThemeColors(theme),
  };
}

export function resolveXlsxColor(
  element: Element | undefined,
  resolver: XlsxColorResolver,
): string | undefined {
  if (!element) return undefined;
  const direct = normalizedHexColor(attribute(element, 'rgb'));
  const themeIndex = boundedInteger(attribute(element, 'theme'), 11);
  const indexed = boundedInteger(attribute(element, 'indexed'), 65_535);
  const automatic = booleanAttribute(element, 'auto');
  const source =
    direct ??
    (themeIndex === null ? undefined : resolver.theme[themeIndex]) ??
    (indexed === null ? undefined : resolver.indexed[indexed]) ??
    (automatic ? '000000' : undefined);
  if (!source) return undefined;
  const tint = Number(attribute(element, 'tint'));
  return `#${
    Number.isFinite(tint) && tint >= -1 && tint <= 1
      ? tintXlsxColor(source, tint)
      : source
  }`;
}

function readThemeColors(theme: Document | null): readonly string[] {
  if (!theme) return defaultThemeColors;
  const scheme = firstDescendant(theme, 'clrScheme');
  if (!scheme) return defaultThemeColors;
  return themeColorNames.map((name, index) => {
    const entry = directChild(scheme, name);
    const color = entry ? drawingColor(directChildren(entry)[0]) : undefined;
    return color ?? defaultThemeColors[index];
  });
}

function readIndexedColors(styles: Document): readonly string[] {
  const indexed = firstDescendant(styles, 'indexedColors');
  if (!indexed) return defaultIndexedColors;
  const colors = directChildren(indexed, 'rgbColor').map((element) =>
    normalizedHexColor(attribute(element, 'rgb')),
  );
  return colors.every(Boolean) ? (colors as string[]) : defaultIndexedColors;
}

function drawingColor(element: Element | undefined): string | undefined {
  if (!element) return undefined;
  return normalizedHexColor(
    element.localName === 'sysClr'
      ? attribute(element, 'lastClr')
      : attribute(element, 'val'),
  );
}

function tintXlsxColor(color: string, tint: number): string {
  const [hue, saturation, lightness] = rgbToHsl(
    [0, 2, 4].map(
      (offset) => Number.parseInt(color.slice(offset, offset + 2), 16) / 255,
    ) as [number, number, number],
  );
  const tintedLightness =
    tint < 0 ? lightness * (1 + tint) : 1 - (1 - lightness) * (1 - tint);
  return hslToRgb(hue, saturation, tintedLightness)
    .map((channel) => Math.round(channel * 255))
    .map((channel) => channel.toString(16).padStart(2, '0'))
    .join('');
}

function rgbToHsl([red, green, blue]: [number, number, number]): [
  number,
  number,
  number,
] {
  const maximum = Math.max(red, green, blue);
  const minimum = Math.min(red, green, blue);
  const delta = maximum - minimum;
  const lightness = (maximum + minimum) / 2;
  if (delta === 0) return [0, 0, lightness];

  const saturation = delta / (1 - Math.abs(2 * lightness - 1));
  const hue =
    maximum === red
      ? ((green - blue) / delta + (green < blue ? 6 : 0)) / 6
      : maximum === green
        ? ((blue - red) / delta + 2) / 6
        : ((red - green) / delta + 4) / 6;
  return [hue, saturation, lightness];
}

function hslToRgb(
  hue: number,
  saturation: number,
  lightness: number,
): [number, number, number] {
  if (saturation === 0) return [lightness, lightness, lightness];
  const chroma = saturation * 2 * (lightness < 0.5 ? lightness : 1 - lightness);
  const minimum = lightness - chroma / 2;
  const channels: [number, number, number] = [minimum, minimum, minimum];
  const sector = 6 * hue;
  switch (Math.floor(sector)) {
    case 0:
    case 6:
      channels[0] += chroma;
      channels[1] += chroma * sector;
      break;
    case 1:
      channels[0] += chroma * (2 - sector);
      channels[1] += chroma;
      break;
    case 2:
      channels[1] += chroma;
      channels[2] += chroma * (sector - 2);
      break;
    case 3:
      channels[1] += chroma * (4 - sector);
      channels[2] += chroma;
      break;
    case 4:
      channels[0] += chroma * (sector - 4);
      channels[2] += chroma;
      break;
    case 5:
      channels[0] += chroma;
      channels[2] += chroma * (6 - sector);
      break;
  }
  return channels;
}

function normalizedHexColor(value: string | null): string | undefined {
  if (!value || !/^[0-9a-f]{6,8}$/i.test(value)) return undefined;
  return value.slice(-6).toLowerCase();
}

function boundedInteger(value: string | null, maximum: number): number | null {
  if (value === null || !/^\d+$/.test(value)) return null;
  const parsed = Number(value);
  return Number.isSafeInteger(parsed) && parsed <= maximum ? parsed : null;
}

function booleanAttribute(element: Element, name: string): boolean {
  const value = attribute(element, name)?.trim().toLowerCase();
  return value === '1' || value === 'true' || value === 'on';
}
