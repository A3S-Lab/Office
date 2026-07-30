import type { Editor } from '@tiptap/core';
import {
  OFFICE_DOCUMENT_LAYOUT_ARABIC_FONT_ID,
  OFFICE_DOCUMENT_LAYOUT_FONT_ID,
  OFFICE_DOCUMENT_LAYOUT_HEBREW_FONT_ID,
  OFFICE_DOCUMENT_LAYOUT_LATIN_FONT_ID,
  type WorkDocumentLayoutFont,
} from '../work-document-fonts';
import {
  normalizeOfficeFontFamily,
  officeFontFamilies,
  officeFontFamilyLabel,
} from './office-font-families';
import type { OfficeSelectOption } from './office-select';

export const documentFontFamilyOptions: readonly OfficeSelectOption[] = [
  { value: 'default', label: '默认字体' },
  ...officeFontFamilies.map(({ cssFamily, group, label, name }) => ({
    value: cssFamily,
    group,
    label,
    previewStyle: { fontFamily: cssFamily },
    searchText: `${name} ${label}`,
  })),
];

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
): string {
  const value = editor.getAttributes('textStyle').fontFamily;
  if (typeof value !== 'string' || !value.trim()) return 'default';
  const normalized = normalizeOfficeFontFamily(value);
  return (
    documentFontFamilyOptionsForLayoutFonts(layoutFonts).find(
      (option) =>
        option.value !== 'default' &&
        normalizeOfficeFontFamily(option.value) === normalized,
    )?.value ?? value.trim()
  );
}

export function documentFontFamilyOptionsForValue(
  value: string,
  layoutFonts: readonly WorkDocumentLayoutFont[] = [],
) {
  const options = documentFontFamilyOptionsForLayoutFonts(layoutFonts);
  if (value === 'default' || options.some((option) => option.value === value)) {
    return options;
  }
  return [
    ...options,
    {
      value,
      group: '文档字体',
      label: officeFontFamilyLabel(value),
      previewStyle: { fontFamily: value },
      searchText: value,
    },
  ];
}

function documentFontFamilyOptionsForLayoutFonts(
  layoutFonts: readonly WorkDocumentLayoutFont[],
): readonly OfficeSelectOption[] {
  if (!layoutFonts.length) return documentFontFamilyOptions;
  const layoutOptions = documentLayoutFontFamilyOptions(layoutFonts);
  const layoutFamilies = new Set(
    layoutOptions.map((option) => normalizeOfficeFontFamily(option.value)),
  );
  return [
    documentFontFamilyOptions[0] as OfficeSelectOption,
    ...layoutOptions,
    ...documentFontFamilyOptions
      .slice(1)
      .filter(
        (option) =>
          !layoutFamilies.has(normalizeOfficeFontFamily(option.value)),
      ),
  ];
}

const bundledDocumentFontLabels = new Map<string, string>([
  [OFFICE_DOCUMENT_LAYOUT_LATIN_FONT_ID, 'Noto Sans'],
  [OFFICE_DOCUMENT_LAYOUT_FONT_ID, '思源黑体'],
  [OFFICE_DOCUMENT_LAYOUT_ARABIC_FONT_ID, 'Noto Naskh Arabic'],
  [OFFICE_DOCUMENT_LAYOUT_HEBREW_FONT_ID, 'Noto Sans Hebrew'],
]);

function documentLayoutFontFamilyOptions(
  layoutFonts: readonly WorkDocumentLayoutFont[],
): OfficeSelectOption[] {
  const families = new Set<string>();
  return layoutFonts.flatMap((font) => {
    const normalized = normalizeOfficeFontFamily(font.family);
    if (!normalized || families.has(normalized)) return [];
    families.add(normalized);
    const bundledLabel = bundledDocumentFontLabels.get(font.id);
    const value = cssFontFamilyValue(font.family);
    const label = bundledLabel ?? font.family;
    return [
      {
        value,
        group: bundledLabel ? '内置字体' : '项目字体',
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

export function documentFontSizeValue(editor: Editor): string {
  const value = editor.getAttributes('textStyle').fontSize;
  if (value === '10.5pt' || !value) return 'default';
  if (documentFontSizeOptions.some((option) => option.value === value)) {
    return value;
  }
  return parseFontSizePoints(value) === null ? 'default' : value;
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
  const current = fontSizePoints(editor.getAttributes('textStyle').fontSize);
  const next = nextDocumentFontSize(current, direction);
  if (next === null) return false;
  return editor.chain().focus().setFontSize(`${next}pt`).run();
}

export function canChangeDocumentFontSize(
  editor: Editor,
  direction: -1 | 1,
): boolean {
  return (
    nextDocumentFontSize(
      fontSizePoints(editor.getAttributes('textStyle').fontSize),
      direction,
    ) !== null
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

function fontSizePoints(value: unknown): number {
  return parseFontSizePoints(value) ?? 10.5;
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
