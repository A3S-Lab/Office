import { Extension, type Editor } from '@tiptap/core';

export const DOCUMENT_INDENT_STEP_PX = 24;
export const MAX_DOCUMENT_INDENT_LEVEL = 8;

export const DocumentParagraphFormatting = Extension.create({
  name: 'documentParagraphFormatting',

  addGlobalAttributes() {
    return [
      {
        types: ['paragraph', 'heading'],
        attributes: {
          lineHeight: {
            default: null,
            parseHTML: (element: HTMLElement) =>
              normalizedLineHeight(element.style.lineHeight),
            renderHTML: (attributes: Record<string, unknown>) => {
              const lineHeight = normalizedLineHeight(attributes.lineHeight);
              return lineHeight ? { style: `line-height: ${lineHeight}` } : {};
            },
          },
          indentLevel: {
            default: 0,
            parseHTML: (element: HTMLElement) =>
              parsedIndentLevel(
                element.dataset.officeIndentLevel,
                element.style.marginLeft,
              ),
            renderHTML: (attributes: Record<string, unknown>) => {
              const indentLevel = normalizedIndentLevel(attributes.indentLevel);
              return indentLevel
                ? {
                    'data-office-indent-level': String(indentLevel),
                    style: `margin-left: ${
                      indentLevel * DOCUMENT_INDENT_STEP_PX
                    }px`,
                  }
                : {};
            },
          },
        },
      },
    ];
  },
});

export function setDocumentLineHeight(
  editor: Editor,
  lineHeight: string | null,
): boolean {
  const value = normalizedLineHeight(lineHeight);
  return editor
    .chain()
    .focus()
    .updateAttributes('paragraph', { lineHeight: value })
    .updateAttributes('heading', { lineHeight: value })
    .run();
}

export function changeDocumentIndent(
  editor: Editor,
  direction: -1 | 1,
): boolean {
  if (editor.isActive('listItem')) {
    const canChangeListLevel =
      direction > 0
        ? editor.can().chain().focus().sinkListItem('listItem').run()
        : editor.can().chain().focus().liftListItem('listItem').run();
    if (canChangeListLevel) {
      return direction > 0
        ? editor.chain().focus().sinkListItem('listItem').run()
        : editor.chain().focus().liftListItem('listItem').run();
    }
  }

  const attributes = editor.isActive('heading')
    ? editor.getAttributes('heading')
    : editor.getAttributes('paragraph');
  const indentLevel = normalizedIndentLevel(attributes.indentLevel);
  const nextLevel = normalizedIndentLevel(indentLevel + direction);
  return editor
    .chain()
    .focus()
    .updateAttributes('paragraph', { indentLevel: nextLevel })
    .updateAttributes('heading', { indentLevel: nextLevel })
    .run();
}

export function clearDocumentFormatting(editor: Editor): boolean {
  return editor
    .chain()
    .focus()
    .unsetAllMarks()
    .setParagraph()
    .unsetTextAlign()
    .updateAttributes('paragraph', {
      indentLevel: 0,
      lineHeight: null,
    })
    .removeEmptyTextStyle()
    .run();
}

function parsedIndentLevel(
  dataValue: string | undefined,
  marginLeft: string,
): number {
  if (dataValue) return normalizedIndentLevel(dataValue);
  const pixels = cssPixels(marginLeft);
  return normalizedIndentLevel(
    pixels ? Math.round(pixels / DOCUMENT_INDENT_STEP_PX) : 0,
  );
}

function normalizedIndentLevel(value: unknown): number {
  const number = Number(value);
  if (!Number.isFinite(number)) return 0;
  return Math.min(MAX_DOCUMENT_INDENT_LEVEL, Math.max(0, Math.round(number)));
}

function normalizedLineHeight(value: unknown): string | null {
  if (typeof value !== 'string') return null;
  const normalized = value.trim();
  if (!normalized || normalized === 'normal') return null;
  return /^(?:\d+(?:\.\d+)?|\d+(?:\.\d+)?(?:px|pt|%))$/i.test(normalized)
    ? normalized
    : null;
}

function cssPixels(value: string): number {
  const match = /^(-?\d+(?:\.\d+)?)(px|pt)?$/i.exec(value.trim());
  if (!match) return 0;
  const amount = Number(match[1]);
  if (!Number.isFinite(amount)) return 0;
  return match[2]?.toLowerCase() === 'pt' ? amount / 0.75 : amount;
}
