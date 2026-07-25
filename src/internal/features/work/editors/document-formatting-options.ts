import type { Editor } from '@tiptap/core';

export const documentFontFamilyOptions = [
  { value: 'default', label: '默认字体' },
  {
    value: '"Microsoft YaHei", "PingFang SC", sans-serif',
    label: '微软雅黑',
  },
  { value: 'SimSun, "Songti SC", serif', label: '宋体' },
  { value: 'SimHei, "Heiti SC", sans-serif', label: '黑体' },
  { value: 'KaiTi, "Kaiti SC", serif', label: '楷体' },
  { value: 'Arial, sans-serif', label: 'Arial' },
  { value: '"Times New Roman", serif', label: 'Times New Roman' },
] as const;

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
): (typeof documentFontFamilyOptions)[number]['value'] {
  const value = editor.getAttributes('textStyle').fontFamily;
  return documentFontFamilyOptions.some((option) => option.value === value)
    ? (value as (typeof documentFontFamilyOptions)[number]['value'])
    : 'default';
}

export function documentFontSizeValue(
  editor: Editor,
): (typeof documentFontSizeOptions)[number]['value'] {
  const value = editor.getAttributes('textStyle').fontSize;
  if (value === '10.5pt' || !value) return 'default';
  return documentFontSizeOptions.some((option) => option.value === value)
    ? (value as (typeof documentFontSizeOptions)[number]['value'])
    : 'default';
}

export function changeDocumentFontSize(
  editor: Editor,
  direction: -1 | 1,
): boolean {
  const current = fontSizePoints(editor.getAttributes('textStyle').fontSize);
  const next =
    direction > 0
      ? (documentFontSizeSteps.find((size) => size > current) ??
        documentFontSizeSteps.at(-1))
      : ([...documentFontSizeSteps].reverse().find((size) => size < current) ??
        documentFontSizeSteps[0]);
  return editor
    .chain()
    .focus()
    .setFontSize(`${next ?? 10.5}pt`)
    .run();
}

function fontSizePoints(value: unknown): number {
  if (typeof value !== 'string') return 10.5;
  const match = /^(\d+(?:\.\d+)?)(px|pt)?$/i.exec(value.trim());
  if (!match) return 10.5;
  const amount = Number(match[1]);
  if (!Number.isFinite(amount)) return 10.5;
  return match[2]?.toLowerCase() === 'px' ? amount * 0.75 : amount;
}
