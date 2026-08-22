import { mergeAttributes, type Editor } from '@tiptap/core';
import Underline from '@tiptap/extension-underline';
import {
  type DocxThemeColorReference,
  parseDocxThemeReference,
  serializeDocxThemeReference,
} from './work-docx-theme-reference';

export const workDocumentUnderlineStyles = [
  'none',
  'single',
  'words',
  'double',
  'thick',
  'dotted',
  'dottedHeavy',
  'dash',
  'dashedHeavy',
  'dashLong',
  'dashLongHeavy',
  'dotDash',
  'dashDotHeavy',
  'dotDotDash',
  'dashDotDotHeavy',
  'wave',
  'wavyHeavy',
  'wavyDouble',
] as const;

export type WorkDocumentUnderlineStyle =
  (typeof workDocumentUnderlineStyles)[number];

export interface WorkDocumentUnderlineFormatting {
  style: WorkDocumentUnderlineStyle;
  color?: string;
  themeColor?: DocxThemeColorReference;
}

export interface WorkDocumentUnderlineCommandOptions {
  color?: string | null;
  themeColor?: DocxThemeColorReference | string | null;
}

export const DOCUMENT_UNDERLINE_STYLE_ATTRIBUTE = 'data-office-underline-style';
export const DOCUMENT_UNDERLINE_COLOR_ATTRIBUTE = 'data-office-underline-color';
export const DOCUMENT_UNDERLINE_THEME_COLOR_ATTRIBUTE =
  'data-office-underline-theme-color';

export const documentUnderlineKeyboardShortcuts = {
  double: 'Mod-Shift-d',
  words: 'Mod-Shift-w',
} as const;

const UNDERLINE_STYLE_BY_NORMALIZED_VALUE = new Map(
  workDocumentUnderlineStyles.map((style) => [style.toLowerCase(), style]),
);
const HEAVY_UNDERLINE_STYLES = new Set<WorkDocumentUnderlineStyle>([
  'thick',
  'dottedHeavy',
  'dashedHeavy',
  'dashLongHeavy',
  'dashDotHeavy',
  'dashDotDotHeavy',
  'wavyHeavy',
]);

declare module '@tiptap/core' {
  interface Commands<ReturnType> {
    documentUnderline: {
      setDocumentUnderline: (
        style: WorkDocumentUnderlineStyle,
        options?: WorkDocumentUnderlineCommandOptions,
      ) => ReturnType;
      setDocumentUnderlineColor: (
        color: string | null,
        themeColor?: DocxThemeColorReference | string | null,
      ) => ReturnType;
      toggleDocumentUnderlineStyle: (
        style: Exclude<WorkDocumentUnderlineStyle, 'none'>,
      ) => ReturnType;
    };
  }
}

export const DocumentUnderline = Underline.extend({
  addAttributes() {
    return {
      underlineStyle: {
        default: 'single',
        parseHTML: (element: HTMLElement) =>
          documentUnderlineFormattingFromElement(element)?.style ?? 'single',
        renderHTML: () => ({}),
      },
      underlineColor: {
        default: null,
        parseHTML: (element: HTMLElement) =>
          documentUnderlineFormattingFromElement(element)?.color ?? null,
        renderHTML: () => ({}),
      },
      underlineThemeColor: {
        default: null,
        parseHTML: (element: HTMLElement) =>
          serializeDocxThemeReference(
            documentUnderlineFormattingFromElement(element)?.themeColor ?? null,
          ) ?? null,
        renderHTML: () => ({}),
      },
    };
  },

  renderHTML({ mark, HTMLAttributes }) {
    const attributes = documentUnderlineDomAttributes(
      documentUnderlineFormattingFromMarkAttributes(mark.attrs),
    );
    return [
      'u',
      mergeAttributes(this.options.HTMLAttributes, HTMLAttributes, attributes),
      0,
    ];
  },

  addCommands() {
    return {
      setUnderline:
        () =>
        ({ commands, editor }) =>
          commands.setMark(
            this.name,
            documentUnderlineMarkAttributes(editor, 'single'),
          ),
      toggleUnderline:
        () =>
        ({ commands, editor }) => {
          const style =
            documentUnderlineStyle(editor) === 'none' ? 'single' : 'none';
          return commands.setMark(
            this.name,
            documentUnderlineMarkAttributes(editor, style),
          );
        },
      unsetUnderline:
        () =>
        ({ commands }) =>
          commands.unsetMark(this.name),
      setDocumentUnderline:
        (style, options = {}) =>
        ({ commands, editor }) =>
          commands.setMark(
            this.name,
            documentUnderlineMarkAttributes(editor, style, options),
          ),
      setDocumentUnderlineColor:
        (color, themeColor = null) =>
        ({ commands, editor }) => {
          const normalizedColor = normalizeDocumentUnderlineColor(color);
          if (color !== null && !normalizedColor) return false;
          const style = documentUnderlineStyle(editor);
          return commands.setMark(
            this.name,
            documentUnderlineMarkAttributes(
              editor,
              style === 'none' ? 'single' : style,
              { color: normalizedColor, themeColor },
            ),
          );
        },
      toggleDocumentUnderlineStyle:
        (style) =>
        ({ commands, editor }) => {
          const next =
            documentUnderlineStyle(editor) === style ? 'none' : style;
          return commands.setMark(
            this.name,
            documentUnderlineMarkAttributes(editor, next),
          );
        },
    };
  },

  addKeyboardShortcuts() {
    return {
      'Mod-u': () => this.editor.commands.toggleUnderline(),
      'Mod-U': () => this.editor.commands.toggleUnderline(),
      [documentUnderlineKeyboardShortcuts.double]: () =>
        this.editor.commands.toggleDocumentUnderlineStyle('double'),
      [documentUnderlineKeyboardShortcuts.words]: () =>
        this.editor.commands.toggleDocumentUnderlineStyle('words'),
    };
  },
});

export function normalizeDocumentUnderlineStyle(
  value: unknown,
): WorkDocumentUnderlineStyle | null {
  if (typeof value !== 'string') return null;
  const normalized = value.trim().toLowerCase();
  if (normalized === '0' || normalized === 'false' || normalized === 'off') {
    return 'none';
  }
  return UNDERLINE_STYLE_BY_NORMALIZED_VALUE.get(normalized) ?? null;
}

export function normalizeDocumentUnderlineColor(value: unknown): string | null {
  if (typeof value !== 'string') return null;
  const source = value.trim();
  const hex = /^#?([\da-f]{6})$/i.exec(source)?.[1];
  if (hex) return `#${hex.toLowerCase()}`;
  const shortHex = /^#([\da-f]{3})$/i.exec(source)?.[1];
  if (shortHex) {
    return `#${[...shortHex]
      .map((channel) => `${channel}${channel}`)
      .join('')
      .toLowerCase()}`;
  }
  const rgb =
    /^rgba?\(\s*(\d{1,3})\s*,\s*(\d{1,3})\s*,\s*(\d{1,3})(?:\s*,\s*(?:1(?:\.0+)?|0?\.\d+))?\s*\)$/i.exec(
      source,
    );
  if (!rgb) return null;
  const channels = rgb.slice(1, 4).map(Number);
  if (channels.some((channel) => channel < 0 || channel > 255)) return null;
  return `#${channels
    .map((channel) => channel.toString(16).padStart(2, '0'))
    .join('')}`;
}

export function documentUnderlineStyle(
  editor: Editor,
): WorkDocumentUnderlineStyle {
  return documentUnderlineFormatting(editor)?.style ?? 'none';
}

export function documentUnderlineColor(editor: Editor): string | null {
  return documentUnderlineFormatting(editor)?.color ?? null;
}

export function documentUnderlineFormatting(
  editor: Editor,
): WorkDocumentUnderlineFormatting | null {
  if (editor.isDestroyed || !editor.isActive('underline')) return null;
  return documentUnderlineFormattingFromMarkAttributes(
    editor.getAttributes('underline'),
  );
}

export function documentUnderlineFormattingFromElement(
  element: HTMLElement,
): WorkDocumentUnderlineFormatting | null {
  const declaredStyle = normalizeDocumentUnderlineStyle(
    element.getAttribute(DOCUMENT_UNDERLINE_STYLE_ATTRIBUTE),
  );
  const cssLine = element.style.textDecorationLine
    .trim()
    .toLowerCase()
    .split(/\s+/);
  const style =
    declaredStyle ??
    documentUnderlineStyleFromCss(element) ??
    (element.tagName.toLowerCase() === 'u'
      ? 'single'
      : cssLine.includes('none')
        ? 'none'
        : null);
  if (!style) return null;
  const themeColor = parseDocxThemeReference(
    element.getAttribute(DOCUMENT_UNDERLINE_THEME_COLOR_ATTRIBUTE),
  );
  const color =
    normalizeDocumentUnderlineColor(
      element.getAttribute(DOCUMENT_UNDERLINE_COLOR_ATTRIBUTE),
    ) ??
    normalizeDocumentUnderlineColor(element.style.textDecorationColor) ??
    themeColor?.resolved;
  return {
    style,
    ...(color ? { color } : {}),
    ...(themeColor ? { themeColor } : {}),
  };
}

export function documentUnderlineDomAttributes(
  formatting: WorkDocumentUnderlineFormatting,
): Record<string, string> {
  const style = normalizeDocumentUnderlineStyle(formatting.style) ?? 'single';
  const color = normalizeDocumentUnderlineColor(formatting.color);
  const themeColor = serializeDocxThemeReference(formatting.themeColor ?? null);
  const declarations = documentUnderlineCssDeclarations({
    style,
    ...(color ? { color } : {}),
  });
  return {
    [DOCUMENT_UNDERLINE_STYLE_ATTRIBUTE]: style,
    ...(color ? { [DOCUMENT_UNDERLINE_COLOR_ATTRIBUTE]: color } : {}),
    ...(themeColor
      ? { [DOCUMENT_UNDERLINE_THEME_COLOR_ATTRIBUTE]: themeColor }
      : {}),
    style: declarations.join('; '),
  };
}

function documentUnderlineMarkAttributes(
  editor: Editor,
  style: WorkDocumentUnderlineStyle,
  options: WorkDocumentUnderlineCommandOptions = {},
): Record<string, string | null> {
  const current = documentUnderlineFormatting(editor);
  const color =
    options.color === undefined
      ? (current?.color ?? null)
      : normalizeDocumentUnderlineColor(options.color);
  const requestedTheme =
    options.themeColor === undefined
      ? (current?.themeColor ?? null)
      : documentUnderlineThemeReference(options.themeColor);
  const themeColor =
    requestedTheme && color === requestedTheme.resolved
      ? (serializeDocxThemeReference(requestedTheme) ?? null)
      : null;
  return {
    underlineStyle: style,
    underlineColor: color,
    underlineThemeColor: themeColor,
  };
}

function documentUnderlineFormattingFromMarkAttributes(
  attributes: Record<string, unknown>,
): WorkDocumentUnderlineFormatting {
  const style =
    normalizeDocumentUnderlineStyle(attributes.underlineStyle) ?? 'single';
  const themeColor =
    typeof attributes.underlineThemeColor === 'string'
      ? parseDocxThemeReference(attributes.underlineThemeColor)
      : null;
  const color =
    normalizeDocumentUnderlineColor(attributes.underlineColor) ??
    themeColor?.resolved;
  return {
    style,
    ...(color ? { color } : {}),
    ...(themeColor ? { themeColor } : {}),
  };
}

function documentUnderlineThemeReference(
  value: DocxThemeColorReference | string | null,
): DocxThemeColorReference | null {
  if (typeof value === 'string') return parseDocxThemeReference(value);
  return value
    ? parseDocxThemeReference(serializeDocxThemeReference(value))
    : null;
}

function documentUnderlineStyleFromCss(
  element: HTMLElement,
): WorkDocumentUnderlineStyle | null {
  const line = element.style.textDecorationLine
    .trim()
    .toLowerCase()
    .split(/\s+/);
  if (line.includes('none')) return 'none';
  if (!line.includes('underline')) return null;
  switch (element.style.textDecorationStyle.trim().toLowerCase()) {
    case 'double':
      return 'double';
    case 'dotted':
      return 'dotted';
    case 'dashed':
      return 'dash';
    case 'wavy':
      return 'wave';
    default:
      return 'single';
  }
}

function documentUnderlineCssDeclarations(
  formatting: Pick<WorkDocumentUnderlineFormatting, 'color' | 'style'>,
): string[] {
  const { style, color } = formatting;
  if (style === 'none') return ['text-decoration-line: none'];
  const declarations = [
    'text-decoration-line: underline',
    `text-decoration-style: ${documentUnderlineCssStyle(style)}`,
    'text-underline-offset: 0.12em',
    'text-decoration-skip-ink: auto',
  ];
  if (color) declarations.push(`text-decoration-color: ${color}`);
  if (HEAVY_UNDERLINE_STYLES.has(style)) {
    declarations.push('text-decoration-thickness: 0.14em');
  }
  return declarations;
}

function documentUnderlineCssStyle(
  style: Exclude<WorkDocumentUnderlineStyle, 'none'>,
): 'dashed' | 'dotted' | 'double' | 'solid' | 'wavy' {
  if (style === 'double') return 'double';
  if (style === 'dotted' || style === 'dottedHeavy') return 'dotted';
  if (
    style === 'dash' ||
    style === 'dashedHeavy' ||
    style === 'dashLong' ||
    style === 'dashLongHeavy' ||
    style === 'dotDash' ||
    style === 'dashDotHeavy' ||
    style === 'dotDotDash' ||
    style === 'dashDotDotHeavy'
  ) {
    return 'dashed';
  }
  if (style === 'wave' || style === 'wavyHeavy' || style === 'wavyDouble') {
    return 'wavy';
  }
  return 'solid';
}
