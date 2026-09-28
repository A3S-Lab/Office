import type { Editor } from '@tiptap/core';
import {
  OFFICE_DOCUMENT_LAYOUT_ARABIC_FONT_ID,
  OFFICE_DOCUMENT_LAYOUT_FONT_ID,
  OFFICE_DOCUMENT_LAYOUT_HEBREW_FONT_ID,
  OFFICE_DOCUMENT_LAYOUT_LATIN_FONT_ID,
  type WorkDocumentLayoutFont,
} from '../work-document-fonts';
import {
  officeMessage,
  resolveOfficeMessages,
} from '../../../i18n/office-locale';
import type { OfficeMessageCatalog } from '../../../i18n/office-messages';
import {
  normalizeOfficeFontFamily,
  officeFontFamilies,
  officeFontFamilyGroupLabel,
  officeFontFamilyLabel,
  officeFontFamilyLocalizedLabel,
} from './office-font-families';
import type { OfficeSelectOption } from './office-select';

export function documentFontFamilyOptions(
  messages: OfficeMessageCatalog = resolveOfficeMessages(),
): readonly OfficeSelectOption[] {
  return [
    {
      value: 'default',
      label: officeMessage(messages, 'document.font.family.default'),
    },
    ...officeFontFamilies.map((family) => {
      const label = officeFontFamilyLocalizedLabel(family, messages);
      return {
        value: family.cssFamily,
        group: officeFontFamilyGroupLabel(family.group, messages),
        label,
        previewStyle: { fontFamily: family.cssFamily },
        searchText: `${family.name} ${label}`,
      };
    }),
  ];
}

export const documentFontSizeOptions = [
  { value: 'default', label: '10.5' },
  { value: '9pt', label: '9' },
  { value: '12pt', label: '12' },
  { value: '14pt', label: '14' },
  { value: '16pt', label: '16' },
  { value: '18pt', label: '18' },
  { value: '22pt', label: '22' },
  { value: '24pt', label: '24' },
  { value: '36pt', label: '36' },
  { value: '48pt', label: '48' },
  { value: '72pt', label: '72' },
] as const;

const documentFontSizeSteps = [9, 10.5, 12, 14, 16, 18, 22, 24, 36, 48, 72];

export function documentFontFamilyValue(
  editor: Editor,
  layoutFonts: readonly WorkDocumentLayoutFont[] = [],
  messages: OfficeMessageCatalog = resolveOfficeMessages(),
): string {
  const value = editor.getAttributes('textStyle').fontFamily;
  if (typeof value !== 'string' || !value.trim()) return 'default';
  const normalized = normalizeOfficeFontFamily(value);
  return (
    documentFontFamilyOptionsForLayoutFonts(layoutFonts, messages).find(
      (option) =>
        option.value !== 'default' &&
        normalizeOfficeFontFamily(option.value) === normalized,
    )?.value ?? value.trim()
  );
}

export function documentFontFamilyOptionsForValue(
  value: string,
  layoutFonts: readonly WorkDocumentLayoutFont[] = [],
  messages: OfficeMessageCatalog = resolveOfficeMessages(),
) {
  const options = documentFontFamilyOptionsForLayoutFonts(layoutFonts, messages);
  if (value === 'default' || options.some((option) => option.value === value)) {
    return options;
  }
  return [
    ...options,
    {
      value,
      group: officeMessage(messages, 'document.font.familyGroup.document'),
      label: officeFontFamilyLabel(value),
      previewStyle: { fontFamily: value },
      searchText: value,
    },
  ];
}

function documentFontFamilyOptionsForLayoutFonts(
  layoutFonts: readonly WorkDocumentLayoutFont[],
  messages: OfficeMessageCatalog = resolveOfficeMessages(),
): readonly OfficeSelectOption[] {
  const base = documentFontFamilyOptions(messages);
  if (!layoutFonts.length) return base;
  const layoutOptions = documentLayoutFontFamilyOptions(layoutFonts, messages);
  const layoutFamilies = new Set(
    layoutOptions.map((option) => normalizeOfficeFontFamily(option.value)),
  );
  return [
    base[0] as OfficeSelectOption,
    ...layoutOptions,
    ...base
      .slice(1)
      .filter(
        (option) =>
          !layoutFamilies.has(normalizeOfficeFontFamily(option.value)),
      ),
  ];
}

function documentLayoutFontFamilyOptions(
  layoutFonts: readonly WorkDocumentLayoutFont[],
  messages: OfficeMessageCatalog,
): OfficeSelectOption[] {
  const families = new Set<string>();
  return layoutFonts.flatMap((font) => {
    const normalized = normalizeOfficeFontFamily(font.family);
    if (!normalized || families.has(normalized)) return [];
    families.add(normalized);
    const bundledLabel =
      font.id === OFFICE_DOCUMENT_LAYOUT_FONT_ID
        ? officeMessage(messages, 'document.font.family.sourceHanSans')
        : font.id === OFFICE_DOCUMENT_LAYOUT_LATIN_FONT_ID
          ? 'Noto Sans'
          : font.id === OFFICE_DOCUMENT_LAYOUT_ARABIC_FONT_ID
            ? 'Noto Naskh Arabic'
            : font.id === OFFICE_DOCUMENT_LAYOUT_HEBREW_FONT_ID
              ? 'Noto Sans Hebrew'
              : null;
    const value = cssFontFamilyValue(font.family);
    const label = bundledLabel ?? font.family;
    return [
      {
        value,
        group: officeMessage(
          messages,
          bundledLabel
            ? 'document.font.familyGroup.bundled'
            : 'document.font.familyGroup.project',
        ),
        label,
        previewStyle: { fontFamily: value },
        searchText: `${font.family} ${label}`,
      },
    ];
  });
}

function cssFontFamilyValue(family: string): string {
  const trimmed = family.trim();
  return /[\s,'"]/.test(trimmed)
    ? `"${trimmed.replaceAll('"', '\\"')}"`
    : trimmed;
}

/**
 * CSS `font-size` for headings when `textStyle.fontSize` is unset
 * (`work-document-page.css`). Converted to pt via the same 0.75 factor as
 * `parseFontSizePoints` so closed labels match painted size.
 */
export const DOCUMENT_HEADING_CSS_FONT_SIZE_PX: Readonly<
  Partial<Record<1 | 2 | 3 | 4 | 5 | 6, number>>
> = {
  1: 24,
  2: 20,
  3: 17,
};

/** CSS `line-height` for headings when the node has no explicit lineHeight. */
export const DOCUMENT_HEADING_CSS_LINE_HEIGHT: Readonly<
  Partial<Record<1 | 2 | 3 | 4 | 5 | 6, string>>
> = {
  1: '1.32',
  2: '1.4',
};

export function documentFontSizeValue(editor: Editor): string {
  const value = editor.getAttributes('textStyle').fontSize;
  if (value && value !== '10.5pt') {
    if (documentFontSizeOptions.some((option) => option.value === value)) {
      return value;
    }
    return parseFontSizePoints(value) === null ? 'default' : value;
  }
  const headingPoints = documentHeadingCssFontSizePoints(editor);
  if (headingPoints !== null) {
    return pointsToDocumentFontSizeValue(headingPoints);
  }
  return 'default';
}

export function documentLineHeightValue(editor: Editor): string {
  const attributes = editor.isActive('heading')
    ? editor.getAttributes('heading')
    : editor.getAttributes('paragraph');
  const value = attributes.lineHeight;
  if (typeof value === 'string' && value.trim() && value !== 'normal') {
    return value.trim();
  }
  const level = activeDocumentHeadingLevel(editor);
  if (level !== null) {
    const css = DOCUMENT_HEADING_CSS_LINE_HEIGHT[level];
    if (css) return css;
  }
  return 'default';
}

export function documentFontSizeOptionsForValue(value: string) {
  if (
    value === 'default' ||
    documentFontSizeOptions.some((option) => option.value === value)
  ) {
    return documentFontSizeOptions;
  }
  const points = parseFontSizePoints(value);
  if (points === null) return documentFontSizeOptions;
  return [
    ...documentFontSizeOptions,
    { value, label: formatFontSizePoints(points) },
  ];
}

export function changeDocumentFontSize(
  editor: Editor,
  direction: -1 | 1,
): boolean {
  const current = effectiveDocumentFontSizePoints(editor);
  const next = nextDocumentFontSize(current, direction);
  if (next === null) return false;
  return editor.commands.setFontSize(`${next}pt`);
}

export function canChangeDocumentFontSize(
  editor: Editor,
  direction: -1 | 1,
): boolean {
  return (
    nextDocumentFontSize(effectiveDocumentFontSizePoints(editor), direction) !==
    null
  );
}

function nextDocumentFontSize(
  current: number,
  direction: -1 | 1,
): number | null {
  return direction > 0
    ? (documentFontSizeSteps.find((size) => size > current) ?? null)
    : ([...documentFontSizeSteps].reverse().find((size) => size < current) ??
        null);
}

function effectiveDocumentFontSizePoints(editor: Editor): number {
  const mark = parseFontSizePoints(editor.getAttributes('textStyle').fontSize);
  if (mark !== null) return mark;
  return documentHeadingCssFontSizePoints(editor) ?? 10.5;
}

function documentHeadingCssFontSizePoints(editor: Editor): number | null {
  const level = activeDocumentHeadingLevel(editor);
  if (level === null) return null;
  const px = DOCUMENT_HEADING_CSS_FONT_SIZE_PX[level];
  return px === undefined ? null : px * 0.75;
}

function activeDocumentHeadingLevel(
  editor: Editor,
): 1 | 2 | 3 | 4 | 5 | 6 | null {
  for (const level of [1, 2, 3, 4, 5, 6] as const) {
    if (editor.isActive('heading', { level })) return level;
  }
  return null;
}

function pointsToDocumentFontSizeValue(points: number): string {
  return `${formatFontSizePoints(points)}pt`;
}

function parseFontSizePoints(value: unknown): number | null {
  if (typeof value !== 'string') return null;
  const match = /^(\d+(?:\.\d+)?)(px|pt)?$/i.exec(value.trim());
  if (!match) return null;
  const amount = Number(match[1]);
  if (!Number.isFinite(amount) || amount <= 0) return null;
  return match[2]?.toLowerCase() === 'px' ? amount * 0.75 : amount;
}

function formatFontSizePoints(value: number): string {
  return Number.isInteger(value)
    ? String(value)
    : String(Number(value.toFixed(2)));
}
