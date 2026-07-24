import { Extension, type Editor } from '@tiptap/core';

export const DOCUMENT_INDENT_STEP_PX = 24;
export const MAX_DOCUMENT_INDENT_LEVEL = 8;
export const DOCUMENT_RULER_INDENT_STEP_PX = 6;
export const MAX_DOCUMENT_INDENT_PX =
  DOCUMENT_INDENT_STEP_PX * MAX_DOCUMENT_INDENT_LEVEL;

export interface DocumentParagraphIndent {
  left: number;
  right: number;
  firstLine: number;
}

export interface DocumentParagraphPagination {
  keepLines: boolean;
  keepWithNext: boolean;
  pageBreakBefore: boolean;
  widowControl: boolean;
}

export type DocumentParagraphDirection = 'ltr' | 'rtl';

export type DocumentParagraphLineRule = 'auto' | 'exact' | 'atLeast';

export interface DocumentParagraphSpacing {
  before: number | null;
  after: number | null;
  lineHeight: string | null;
  lineRule: DocumentParagraphLineRule | null;
}

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
          lineRule: {
            default: null,
            parseHTML: (element: HTMLElement) =>
              normalizedLineRule(element.dataset.officeLineRule),
            renderHTML: (attributes: Record<string, unknown>) => {
              const lineRule = normalizedLineRule(attributes.lineRule);
              return lineRule ? { 'data-office-line-rule': lineRule } : {};
            },
          },
          spaceBefore: pointSpacingAttribute(
            'spaceBefore',
            'officeSpaceBefore',
            'data-office-space-before',
            'margin-top',
          ),
          spaceAfter: pointSpacingAttribute(
            'spaceAfter',
            'officeSpaceAfter',
            'data-office-space-after',
            'margin-bottom',
          ),
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
          rightIndent: {
            default: 0,
            parseHTML: (element: HTMLElement) =>
              parsedIndentPixels(
                element.dataset.officeIndentRight,
                element.style.marginRight,
              ),
            renderHTML: (attributes: Record<string, unknown>) => {
              const rightIndent = normalizedIndentPixels(
                attributes.rightIndent,
              );
              return rightIndent
                ? {
                    'data-office-indent-right': formatPixelValue(rightIndent),
                    style: `margin-right: ${rightIndent}px`,
                  }
                : {};
            },
          },
          firstLineIndent: {
            default: 0,
            parseHTML: (element: HTMLElement) => {
              const left = parsedIndentLevel(
                element.dataset.officeIndentLevel,
                element.style.marginLeft,
              );
              return normalizedFirstLineIndent(
                parsedSignedIndentPixels(
                  element.dataset.officeIndentFirstLine,
                  element.style.textIndent,
                ),
                left * DOCUMENT_INDENT_STEP_PX,
              );
            },
            renderHTML: (attributes: Record<string, unknown>) => {
              const left =
                normalizedIndentLevel(attributes.indentLevel) *
                DOCUMENT_INDENT_STEP_PX;
              const firstLineIndent = normalizedFirstLineIndent(
                attributes.firstLineIndent,
                left,
              );
              return firstLineIndent
                ? {
                    'data-office-indent-first-line':
                      formatPixelValue(firstLineIndent),
                    style: `text-indent: ${firstLineIndent}px`,
                  }
                : {};
            },
          },
          keepLines: directBooleanAttribute(
            'keepLines',
            'officeKeepLines',
            'data-office-keep-lines',
          ),
          keepWithNext: directBooleanAttribute(
            'keepWithNext',
            'officeKeepWithNext',
            'data-office-keep-with-next',
          ),
          pageBreakBefore: directBooleanAttribute(
            'pageBreakBefore',
            'officePageBreakBefore',
            'data-office-page-break-before',
          ),
          widowControl: directBooleanAttribute(
            'widowControl',
            'officeWidowControl',
            'data-office-widow-control',
          ),
        },
      },
      {
        types: ['paragraph', 'heading', 'listItem'],
        attributes: {
          paragraphDirection: paragraphDirectionAttribute(),
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
  const attributes = {
    lineHeight: value,
    lineRule: value ? lineRuleForLineHeight(value) : null,
  };
  return editor
    .chain()
    .focus()
    .updateAttributes('paragraph', attributes)
    .updateAttributes('heading', attributes)
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

  return setDocumentIndentLevel(
    editor,
    documentIndentLevel(editor) + direction,
  );
}

export function documentIndentLevel(editor: Editor): number {
  return documentParagraphIndent(editor).left / DOCUMENT_INDENT_STEP_PX;
}

export function documentParagraphIndent(
  editor: Editor,
): DocumentParagraphIndent {
  const attributes = editor.isActive('heading')
    ? editor.getAttributes('heading')
    : editor.getAttributes('paragraph');
  return normalizeDocumentParagraphIndent({
    left:
      normalizedIndentLevel(attributes.indentLevel) * DOCUMENT_INDENT_STEP_PX,
    right: normalizedIndentPixels(attributes.rightIndent),
    firstLine: normalizedSignedIndentPixels(attributes.firstLineIndent),
  });
}

export function setDocumentIndentLevel(
  editor: Editor,
  indentLevel: number,
  options: { restoreFocus?: boolean } = {},
): boolean {
  const current = documentParagraphIndent(editor);
  return setDocumentParagraphIndent(
    editor,
    {
      ...current,
      left: normalizedIndentLevel(indentLevel) * DOCUMENT_INDENT_STEP_PX,
    },
    options,
  );
}

export function setDocumentParagraphIndent(
  editor: Editor,
  indent: DocumentParagraphIndent,
  options: { restoreFocus?: boolean } = {},
): boolean {
  const normalized = normalizeDocumentParagraphIndent(indent);
  const attributes = {
    indentLevel: normalized.left / DOCUMENT_INDENT_STEP_PX,
    rightIndent: normalized.right,
    firstLineIndent: normalized.firstLine,
  };
  const nodeTypes = activeParagraphNodeTypes(editor);
  if (!nodeTypes.length) return false;
  const chain = editor.chain();
  if (options.restoreFocus !== false) chain.focus();
  for (const nodeType of nodeTypes)
    chain.updateAttributes(nodeType, attributes);
  return chain.run();
}

export function documentParagraphPagination(
  editor: Editor,
): DocumentParagraphPagination {
  const attributes = activeParagraphAttributes(editor);
  return {
    keepLines: directBoolean(attributes.keepLines) ?? false,
    keepWithNext:
      directBoolean(attributes.keepWithNext) ?? editor.isActive('heading'),
    pageBreakBefore: directBoolean(attributes.pageBreakBefore) ?? false,
    widowControl: directBoolean(attributes.widowControl) ?? true,
  };
}

export function documentParagraphDirection(
  editor: Editor,
): DocumentParagraphDirection {
  const attributes = editor.isActive('listItem')
    ? editor.getAttributes('listItem')
    : activeParagraphAttributes(editor);
  return (
    normalizeDocumentParagraphDirection(attributes.paragraphDirection) ?? 'ltr'
  );
}

export function setDocumentParagraphDirection(
  editor: Editor,
  direction: DocumentParagraphDirection,
  options: { restoreFocus?: boolean } = {},
): boolean {
  const nodeTypes = activeParagraphDirectionNodeTypes(editor);
  if (!nodeTypes.length) return false;
  const value = normalizeDocumentParagraphDirection(direction);
  if (!value) return false;
  const chain = editor.chain();
  if (options.restoreFocus !== false) chain.focus();
  for (const nodeType of nodeTypes)
    chain.updateAttributes(nodeType, { paragraphDirection: value });
  return chain.run();
}

export function setDocumentParagraphPagination(
  editor: Editor,
  pagination: DocumentParagraphPagination,
  options: { restoreFocus?: boolean } = {},
): boolean {
  const nodeTypes = activeParagraphNodeTypes(editor);
  if (!nodeTypes.length) return false;
  const attributes = {
    keepLines: Boolean(pagination.keepLines),
    keepWithNext: Boolean(pagination.keepWithNext),
    pageBreakBefore: Boolean(pagination.pageBreakBefore),
    widowControl: Boolean(pagination.widowControl),
  };
  const chain = editor.chain();
  if (options.restoreFocus !== false) chain.focus();
  for (const nodeType of nodeTypes)
    chain.updateAttributes(nodeType, attributes);
  return chain.run();
}

export function documentParagraphSpacing(
  editor: Editor,
): DocumentParagraphSpacing {
  const attributes = activeParagraphAttributes(editor);
  return {
    before: normalizedPointSpacing(attributes.spaceBefore),
    after: normalizedPointSpacing(attributes.spaceAfter),
    lineHeight: normalizedLineHeight(attributes.lineHeight),
    lineRule: normalizedLineRule(attributes.lineRule),
  };
}

export function setDocumentParagraphSpacing(
  editor: Editor,
  spacing: DocumentParagraphSpacing,
  options: { restoreFocus?: boolean } = {},
): boolean {
  const nodeTypes = activeParagraphNodeTypes(editor);
  if (!nodeTypes.length) return false;
  const lineHeight = normalizedLineHeight(spacing.lineHeight);
  const attributes = {
    spaceBefore: normalizedPointSpacing(spacing.before),
    spaceAfter: normalizedPointSpacing(spacing.after),
    lineHeight,
    lineRule:
      normalizedLineRule(spacing.lineRule) ??
      (lineHeight ? lineRuleForLineHeight(lineHeight) : null),
  };
  const chain = editor.chain();
  if (options.restoreFocus !== false) chain.focus();
  for (const nodeType of nodeTypes)
    chain.updateAttributes(nodeType, attributes);
  return chain.run();
}

export function normalizeDocumentParagraphIndent(
  indent: DocumentParagraphIndent,
): DocumentParagraphIndent {
  const left =
    normalizedIndentLevel(indent.left / DOCUMENT_INDENT_STEP_PX) *
    DOCUMENT_INDENT_STEP_PX;
  return {
    left,
    right: normalizedIndentPixels(indent.right),
    firstLine: normalizedFirstLineIndent(indent.firstLine, left),
  };
}

export function normalizeDocumentParagraphDirection(
  value: unknown,
): DocumentParagraphDirection | null {
  if (typeof value !== 'string') return null;
  const normalized = value.trim().toLowerCase();
  return normalized === 'ltr' || normalized === 'rtl' ? normalized : null;
}

export function clearDocumentFormatting(editor: Editor): boolean {
  const chain = editor.chain().focus().unsetAllMarks();
  if (!editor.isActive('listItem')) chain.setParagraph();
  return chain
    .unsetTextAlign()
    .updateAttributes('paragraph', {
      indentLevel: 0,
      rightIndent: 0,
      firstLineIndent: 0,
      lineHeight: null,
      lineRule: null,
      paragraphDirection: null,
      spaceBefore: null,
      spaceAfter: null,
      keepLines: null,
      keepWithNext: null,
      pageBreakBefore: null,
      widowControl: null,
      tabStops: null,
    })
    .updateAttributes('listItem', {
      paragraphDirection: null,
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
  return Math.min(
    MAX_DOCUMENT_INDENT_LEVEL,
    Math.max(0, Math.round(number * 4) / 4),
  );
}

function normalizedLineHeight(value: unknown): string | null {
  if (typeof value !== 'string') return null;
  const normalized = value.trim();
  if (!normalized || normalized === 'normal') return null;
  return /^(?:\d+(?:\.\d+)?|\d+(?:\.\d+)?(?:px|pt|%))$/i.test(normalized)
    ? normalized
    : null;
}

function normalizedLineRule(value: unknown): DocumentParagraphLineRule | null {
  if (value === 'auto' || value === 'exact' || value === 'atLeast')
    return value;
  if (value === 'exactly') return 'exact';
  return null;
}

function lineRuleForLineHeight(value: string): DocumentParagraphLineRule {
  return /(?:px|pt)$/i.test(value) ? 'exact' : 'auto';
}

function pointSpacingAttribute(
  modelKey: 'spaceBefore' | 'spaceAfter',
  datasetKey: 'officeSpaceBefore' | 'officeSpaceAfter',
  htmlName: 'data-office-space-before' | 'data-office-space-after',
  cssProperty: 'margin-top' | 'margin-bottom',
) {
  return {
    default: null,
    parseHTML: (element: HTMLElement) =>
      parsedPointSpacing(
        element.dataset[datasetKey],
        cssProperty === 'margin-top'
          ? element.style.marginTop
          : element.style.marginBottom,
      ),
    renderHTML: (attributes: Record<string, unknown>) => {
      const value = normalizedPointSpacing(attributes[modelKey]);
      if (value === null) return {};
      const formatted = formatPointValue(value);
      return {
        [htmlName]: formatted,
        style: `${cssProperty}: ${formatted}pt`,
      };
    },
  };
}

function parsedPointSpacing(
  dataValue: string | undefined,
  cssValue: string,
): number | null {
  if (dataValue !== undefined) return normalizedPointSpacing(dataValue);
  const match = /^(-?\d+(?:\.\d+)?)(px|pt|mm|cm|in)?$/i.exec(cssValue.trim());
  if (!match) return null;
  const amount = Number(match[1]);
  if (!Number.isFinite(amount)) return null;
  const unit = match[2]?.toLowerCase() ?? 'px';
  const points =
    unit === 'pt'
      ? amount
      : unit === 'mm'
        ? (amount * 72) / 25.4
        : unit === 'cm'
          ? (amount * 72) / 2.54
          : unit === 'in'
            ? amount * 72
            : amount * 0.75;
  return normalizedPointSpacing(points);
}

function normalizedPointSpacing(value: unknown): number | null {
  if (value === null || value === undefined || value === '') return null;
  const number = Number(value);
  if (!Number.isFinite(number)) return null;
  return Math.min(720, Math.max(0, Math.round(number * 4) / 4));
}

function formatPointValue(value: number): string {
  return Number(value.toFixed(2)).toString();
}

function cssPixels(value: string): number {
  const match = /^(-?\d+(?:\.\d+)?)(px|pt)?$/i.exec(value.trim());
  if (!match) return 0;
  const amount = Number(match[1]);
  if (!Number.isFinite(amount)) return 0;
  return match[2]?.toLowerCase() === 'pt' ? amount / 0.75 : amount;
}

function parsedIndentPixels(
  dataValue: string | undefined,
  cssValue: string,
): number {
  return normalizedIndentPixels(
    dataValue === undefined ? cssPixels(cssValue) : dataValue,
  );
}

function parsedSignedIndentPixels(
  dataValue: string | undefined,
  cssValue: string,
): number {
  return normalizedSignedIndentPixels(
    dataValue === undefined ? cssPixels(cssValue) : dataValue,
  );
}

function normalizedIndentPixels(value: unknown): number {
  const number = Number(value);
  if (!Number.isFinite(number)) return 0;
  return Math.min(MAX_DOCUMENT_INDENT_PX, Math.max(0, Math.round(number)));
}

function normalizedSignedIndentPixels(value: unknown): number {
  const number = Number(value);
  if (!Number.isFinite(number)) return 0;
  return Math.min(
    MAX_DOCUMENT_INDENT_PX,
    Math.max(-MAX_DOCUMENT_INDENT_PX, Math.round(number)),
  );
}

function normalizedFirstLineIndent(value: unknown, leftIndent: number): number {
  return Math.max(
    -leftIndent,
    Math.min(MAX_DOCUMENT_INDENT_PX, normalizedSignedIndentPixels(value)),
  );
}

function formatPixelValue(value: number): string {
  return Number(value.toFixed(2)).toString();
}

function activeParagraphAttributes(editor: Editor): Record<string, unknown> {
  return editor.isActive('heading')
    ? editor.getAttributes('heading')
    : editor.getAttributes('paragraph');
}

function paragraphDirectionAttribute() {
  return {
    default: null,
    parseHTML: (element: HTMLElement) =>
      normalizeDocumentParagraphDirection(
        element.getAttribute('dir') ?? element.style.direction,
      ),
    renderHTML: (attributes: Record<string, unknown>) => {
      const direction = normalizeDocumentParagraphDirection(
        attributes.paragraphDirection,
      );
      return direction ? { dir: direction } : {};
    },
  };
}

function directBooleanAttribute(
  modelKey: 'keepLines' | 'keepWithNext' | 'pageBreakBefore' | 'widowControl',
  datasetKey:
    | 'officeKeepLines'
    | 'officeKeepWithNext'
    | 'officePageBreakBefore'
    | 'officeWidowControl',
  htmlName:
    | 'data-office-keep-lines'
    | 'data-office-keep-with-next'
    | 'data-office-page-break-before'
    | 'data-office-widow-control',
) {
  return {
    default: null,
    parseHTML: (element: HTMLElement) =>
      directBoolean(element.dataset[datasetKey]),
    renderHTML: (attributes: Record<string, unknown>) => {
      const value = directBoolean(attributes[modelKey]);
      return value === null ? {} : { [htmlName]: String(value) };
    },
  };
}

function directBoolean(value: unknown): boolean | null {
  if (value === true || value === 'true' || value === '1') return true;
  if (value === false || value === 'false' || value === '0') return false;
  return null;
}

function activeParagraphNodeTypes(
  editor: Editor,
): Array<'paragraph' | 'heading'> {
  const nodeTypes = new Set<'paragraph' | 'heading'>();
  if (editor.isActive('paragraph')) nodeTypes.add('paragraph');
  if (editor.isActive('heading')) nodeTypes.add('heading');
  const { from, to } = editor.state.selection;
  if (from !== to) {
    editor.state.doc.nodesBetween(from, to, (node) => {
      if (node.type.name === 'paragraph') nodeTypes.add('paragraph');
      if (node.type.name === 'heading') nodeTypes.add('heading');
    });
  }
  return [...nodeTypes];
}

function activeParagraphDirectionNodeTypes(
  editor: Editor,
): Array<'paragraph' | 'heading' | 'listItem'> {
  const { from, to } = editor.state.selection;
  if (from === to) {
    return editor.isActive('listItem')
      ? ['listItem']
      : activeParagraphNodeTypes(editor);
  }
  const nodeTypes = new Set<'paragraph' | 'heading' | 'listItem'>();
  editor.state.doc.nodesBetween(from, to, (node, _position, parent) => {
    if (node.type.name === 'listItem') nodeTypes.add('listItem');
    if (node.type.name === 'paragraph' && parent?.type.name !== 'listItem') {
      nodeTypes.add('paragraph');
    }
    if (node.type.name === 'heading') nodeTypes.add('heading');
  });
  return [...nodeTypes];
}
