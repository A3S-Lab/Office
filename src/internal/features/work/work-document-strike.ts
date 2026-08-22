import { mergeAttributes, type Editor } from '@tiptap/core';
import Strike from '@tiptap/extension-strike';

export const workDocumentStrikeStyles = ['none', 'single', 'double'] as const;

export type WorkDocumentStrikeStyle = (typeof workDocumentStrikeStyles)[number];

export interface WorkDocumentStrikeFormatting {
  style: WorkDocumentStrikeStyle;
}

export const DOCUMENT_STRIKE_STYLE_ATTRIBUTE = 'data-office-strike-style';

const STRIKE_STYLE_BY_NORMALIZED_VALUE = new Map<
  string,
  WorkDocumentStrikeStyle
>(workDocumentStrikeStyles.map((style) => [style, style]));

declare module '@tiptap/core' {
  interface Commands<ReturnType> {
    documentStrike: {
      setDocumentStrike: (style: WorkDocumentStrikeStyle) => ReturnType;
      toggleDocumentStrikeStyle: (
        style: Exclude<WorkDocumentStrikeStyle, 'none'>,
      ) => ReturnType;
    };
  }
}

export const DocumentStrike = Strike.extend({
  addAttributes() {
    return {
      strikeStyle: {
        default: 'single',
        parseHTML: (element: HTMLElement) =>
          documentStrikeFormattingFromElement(element)?.style ?? 'single',
        renderHTML: () => ({}),
      },
    };
  },

  renderHTML({ mark, HTMLAttributes }) {
    return [
      's',
      mergeAttributes(
        this.options.HTMLAttributes,
        HTMLAttributes,
        documentStrikeDomAttributes(
          documentStrikeFormattingFromMarkAttributes(mark.attrs),
        ),
      ),
      0,
    ];
  },

  addCommands() {
    return {
      setStrike:
        () =>
        ({ commands }) =>
          commands.setMark(this.name, { strikeStyle: 'single' }),
      toggleStrike:
        () =>
        ({ commands, editor }) =>
          commands.setMark(this.name, {
            strikeStyle:
              documentStrikeStyle(editor) === 'none' ? 'single' : 'none',
          }),
      unsetStrike:
        () =>
        ({ commands }) =>
          commands.unsetMark(this.name),
      setDocumentStrike:
        (style) =>
        ({ commands }) =>
          commands.setMark(this.name, { strikeStyle: style }),
      toggleDocumentStrikeStyle:
        (style) =>
        ({ commands, editor }) =>
          commands.setMark(this.name, {
            strikeStyle: documentStrikeStyle(editor) === style ? 'none' : style,
          }),
    };
  },

  addKeyboardShortcuts() {
    return {};
  },
});

export function normalizeDocumentStrikeStyle(
  value: unknown,
): WorkDocumentStrikeStyle | null {
  if (typeof value !== 'string') return null;
  const normalized = value.trim().toLowerCase();
  if (normalized === '0' || normalized === 'false' || normalized === 'off') {
    return 'none';
  }
  return STRIKE_STYLE_BY_NORMALIZED_VALUE.get(normalized) ?? null;
}

export function documentStrikeStyle(editor: Editor): WorkDocumentStrikeStyle {
  return documentStrikeFormatting(editor)?.style ?? 'none';
}

export function documentStrikeFormatting(
  editor: Editor,
): WorkDocumentStrikeFormatting | null {
  if (editor.isDestroyed || !editor.isActive('strike')) return null;
  return documentStrikeFormattingFromMarkAttributes(
    editor.getAttributes('strike'),
  );
}

export function documentStrikeFormattingFromElement(
  element: HTMLElement,
): WorkDocumentStrikeFormatting | null {
  const declared = normalizeDocumentStrikeStyle(
    element.getAttribute(DOCUMENT_STRIKE_STYLE_ATTRIBUTE),
  );
  const lines = element.style.textDecorationLine
    .trim()
    .toLowerCase()
    .split(/\s+/);
  const style =
    declared ??
    (lines.includes('none')
      ? 'none'
      : lines.includes('line-through')
        ? element.style.textDecorationStyle.trim().toLowerCase() === 'double'
          ? 'double'
          : 'single'
        : ['del', 's', 'strike'].includes(element.tagName.toLowerCase())
          ? 'single'
          : null);
  return style ? { style } : null;
}

export function documentStrikeDomAttributes(
  formatting: WorkDocumentStrikeFormatting,
): Record<string, string> {
  const style = normalizeDocumentStrikeStyle(formatting.style) ?? 'single';
  return {
    [DOCUMENT_STRIKE_STYLE_ATTRIBUTE]: style,
    style:
      style === 'none'
        ? 'text-decoration-line: none'
        : `text-decoration-line: line-through; text-decoration-style: ${
            style === 'double' ? 'double' : 'solid'
          }; text-decoration-skip-ink: none`,
  };
}

function documentStrikeFormattingFromMarkAttributes(
  attributes: Record<string, unknown>,
): WorkDocumentStrikeFormatting {
  return {
    style: normalizeDocumentStrikeStyle(attributes.strikeStyle) ?? 'single',
  };
}
