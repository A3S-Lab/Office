import type { Editor } from '@tiptap/core';
import {
  normalizeOfficeFontFamily,
  officeFontFamilies,
  officeFontFamilyLabel,
} from './office-font-families';

export const documentFontFamilyOptions = [
  { value: 'default', label: '默认字体' },
  ...officeFontFamilies
    .filter(({ id }) => id !== 'aptos')
    .map(({ cssFamily, label }) => ({
      value: cssFamily,
      label,
      previewStyle: { fontFamily: cssFamily },
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

export function documentFontFamilyValue(editor: Editor): string {
  const value = editor.getAttributes('textStyle').fontFamily;
  if (typeof value !== 'string' || !value.trim()) return 'default';
  const normalized = normalizeOfficeFontFamily(value);
  return (
    documentFontFamilyOptions.find(
      (option) =>
        option.value !== 'default' &&
        normalizeOfficeFontFamily(option.value) === normalized,
    )?.value ?? value.trim()
  );
}

export function documentFontFamilyOptionsForValue(value: string) {
  if (
    value === 'default' ||
    documentFontFamilyOptions.some((option) => option.value === value)
  ) {
    return documentFontFamilyOptions;
  }
  return [
    ...documentFontFamilyOptions,
    {
      value,
      label: officeFontFamilyLabel(value),
      previewStyle: { fontFamily: value },
    },
  ];
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
