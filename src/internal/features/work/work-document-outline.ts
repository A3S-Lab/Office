import type { Node as ProseMirrorNode } from '@tiptap/pm/model';

export interface WorkDocumentOutlineItem {
  id: string;
  text: string;
  level: number;
  depth: number;
  from: number;
  to: number;
  parentId?: string;
  hasChildren: boolean;
}

interface MutableWorkDocumentOutlineItem
  extends Omit<WorkDocumentOutlineItem, 'hasChildren'> {
  hasChildren?: boolean;
}

export function collectWorkDocumentOutline(
  document: ProseMirrorNode,
): WorkDocumentOutlineItem[] {
  const items: MutableWorkDocumentOutlineItem[] = [];
  const hierarchy: MutableWorkDocumentOutlineItem[] = [];

  document.descendants((node, position) => {
    if (node.type.name !== 'heading') return;
    const level = documentHeadingLevel(node.attrs.level);
    while ((hierarchy.at(-1)?.level ?? 0) >= level) hierarchy.pop();
    const parent = hierarchy.at(-1);
    const from = position + 1;
    const item: MutableWorkDocumentOutlineItem = {
      id: `heading-${position}`,
      text: normalizedDocumentHeadingText(node.textContent),
      level,
      depth: hierarchy.length,
      from,
      to: from + node.content.size,
      ...(parent ? { parentId: parent.id } : {}),
    };
    items.push(item);
    hierarchy.push(item);
  });

  return items.map((item, index) => ({
    ...item,
    hasChildren: (items[index + 1]?.depth ?? -1) > item.depth,
  }));
}

export function currentWorkDocumentOutlineItem(
  items: readonly WorkDocumentOutlineItem[],
  position: number,
): WorkDocumentOutlineItem | null {
  let current: WorkDocumentOutlineItem | null = null;
  for (const item of items) {
    if (item.from > position) break;
    current = item;
  }
  return current;
}

export function visibleWorkDocumentOutlineItems(
  items: readonly WorkDocumentOutlineItem[],
  collapsedIds: ReadonlySet<string>,
  rawQuery: string,
): WorkDocumentOutlineItem[] {
  const query = normalizedDocumentOutlineQuery(rawQuery);
  if (query) {
    return items.filter((item) =>
      normalizedDocumentOutlineQuery(item.text).includes(query),
    );
  }

  const visible: WorkDocumentOutlineItem[] = [];
  let hiddenBelowDepth: number | null = null;
  for (const item of items) {
    if (hiddenBelowDepth !== null && item.depth <= hiddenBelowDepth) {
      hiddenBelowDepth = null;
    }
    if (hiddenBelowDepth !== null) continue;
    visible.push(item);
    if (item.hasChildren && collapsedIds.has(item.id)) {
      hiddenBelowDepth = item.depth;
    }
  }
  return visible;
}

function documentHeadingLevel(value: unknown): number {
  return typeof value === 'number' &&
    Number.isInteger(value) &&
    value >= 1 &&
    value <= 6
    ? value
    : 1;
}

function normalizedDocumentHeadingText(value: string): string {
  return value.replace(/\s+/g, ' ').trim() || '未命名标题';
}

function normalizedDocumentOutlineQuery(value: string): string {
  return value.normalize('NFKC').trim().toLocaleLowerCase();
}
