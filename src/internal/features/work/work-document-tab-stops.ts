import { Extension, type Editor } from '@tiptap/core';

export const DEFAULT_DOCUMENT_TAB_INTERVAL_PX = 48;
export const DOCUMENT_TAB_RULER_STEP_PX = 6;
export const MAX_DOCUMENT_TAB_POSITION_PX = 4_096;
export const MAX_DOCUMENT_TAB_STOPS = 64;

export type DocumentTabAlignment = 'left' | 'center' | 'right' | 'decimal';
export type DocumentTabLeader =
  | 'none'
  | 'dot'
  | 'hyphen'
  | 'underscore'
  | 'middleDot';

export interface DocumentTabStop {
  position: number;
  alignment: DocumentTabAlignment;
  leader: DocumentTabLeader;
}

export const DocumentParagraphTabStops = Extension.create({
  name: 'documentParagraphTabStops',

  addGlobalAttributes() {
    return [
      {
        types: ['paragraph', 'heading'],
        attributes: {
          tabStops: {
            default: null,
            parseHTML: (element: HTMLElement) =>
              normalizeDocumentTabStops(element.dataset.officeTabStops),
            renderHTML: (attributes: Record<string, unknown>) => {
              const tabStops = normalizeDocumentTabStops(attributes.tabStops);
              return tabStops.length
                ? {
                    'data-office-tab-stops':
                      serializeDocumentTabStops(tabStops),
                  }
                : {};
            },
          },
        },
      },
    ];
  },
});

export function documentParagraphTabStops(editor: Editor): DocumentTabStop[] {
  return normalizeDocumentTabStops(activeParagraphAttributes(editor).tabStops);
}

export function setDocumentParagraphTabStops(
  editor: Editor,
  tabStops: readonly DocumentTabStop[],
  options: { restoreFocus?: boolean } = {},
): boolean {
  const nodeTypes = activeParagraphNodeTypes(editor);
  if (!nodeTypes.length) return false;
  const normalized = normalizeDocumentTabStops(tabStops);
  const chain = editor.chain();
  if (options.restoreFocus !== false) chain.focus();
  for (const nodeType of nodeTypes) {
    chain.updateAttributes(nodeType, {
      tabStops: normalized.length ? normalized : null,
    });
  }
  return chain.run();
}

export function normalizeDocumentTabStops(value: unknown): DocumentTabStop[] {
  const source = parsedTabStopSource(value);
  const byPosition = new Map<number, DocumentTabStop>();
  for (const candidate of source.slice(0, MAX_DOCUMENT_TAB_STOPS * 4)) {
    if (!isRecord(candidate)) continue;
    const rawPosition = Number(candidate.position);
    if (!Number.isFinite(rawPosition) || rawPosition <= 0) continue;
    const position = normalizedTabPosition(rawPosition);
    if (position <= 0) continue;
    byPosition.set(position, {
      position,
      alignment: normalizedTabAlignment(candidate.alignment),
      leader: normalizedTabLeader(candidate.leader),
    });
  }
  return Array.from(byPosition.values())
    .sort((left, right) => left.position - right.position)
    .slice(0, MAX_DOCUMENT_TAB_STOPS);
}

export function serializeDocumentTabStops(
  value: readonly DocumentTabStop[],
): string {
  return JSON.stringify(normalizeDocumentTabStops(value));
}

export function normalizedTabPosition(value: number): number {
  if (!Number.isFinite(value)) return 0;
  return Math.min(
    MAX_DOCUMENT_TAB_POSITION_PX,
    Math.max(0, Math.round(value * 100) / 100),
  );
}

export function nextDocumentTabAlignment(
  alignment: DocumentTabAlignment,
): DocumentTabAlignment {
  if (alignment === 'left') return 'center';
  if (alignment === 'center') return 'right';
  if (alignment === 'right') return 'decimal';
  return 'left';
}

function parsedTabStopSource(value: unknown): unknown[] {
  if (Array.isArray(value)) return value;
  if (typeof value !== 'string' || !value.trim()) return [];
  try {
    const parsed: unknown = JSON.parse(value);
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

function normalizedTabAlignment(value: unknown): DocumentTabAlignment {
  return value === 'center' || value === 'right' || value === 'decimal'
    ? value
    : 'left';
}

function normalizedTabLeader(value: unknown): DocumentTabLeader {
  return value === 'dot' ||
    value === 'hyphen' ||
    value === 'underscore' ||
    value === 'middleDot'
    ? value
    : 'none';
}

function activeParagraphAttributes(editor: Editor): Record<string, unknown> {
  return editor.isActive('heading')
    ? editor.getAttributes('heading')
    : editor.getAttributes('paragraph');
}

function activeParagraphNodeTypes(
  editor: Editor,
): Array<'paragraph' | 'heading'> {
  const types: Array<'paragraph' | 'heading'> = [];
  if (editor.isActive('paragraph')) types.push('paragraph');
  if (editor.isActive('heading')) types.push('heading');
  return types;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}
