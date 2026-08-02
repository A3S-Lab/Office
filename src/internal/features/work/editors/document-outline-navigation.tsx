import { ChevronDown, ChevronRight } from 'lucide-react';
import {
  type CSSProperties,
  type KeyboardEvent,
  forwardRef,
  useImperativeHandle,
  useMemo,
} from 'react';
import type { WorkDocumentOutlineItem } from '../work-document-outline';
import {
  DOCUMENT_NAVIGATION_COLLECTION_WINDOW_LIMIT,
  DocumentNavigationWindowSpacer,
  type DocumentNavigationListHandle,
  useDocumentNavigationWindow,
} from './document-navigation-window';

const DOCUMENT_OUTLINE_ITEM_HEIGHT = 31;
const DOCUMENT_OUTLINE_LIST_PADDING_TOP = 6;

export const DocumentOutlineNavigation = forwardRef<
  DocumentNavigationListHandle,
  {
    activeItemId: string | null;
    collapsedIds: ReadonlySet<string>;
    items: readonly WorkDocumentOutlineItem[];
    rovingId: string | null;
    onNavigate: (position: number) => void;
    onRovingIdChange: (itemId: string) => void;
    onToggleCollapsed: (itemId: string) => void;
  }
>(function DocumentOutlineNavigation(
  {
    activeItemId,
    collapsedIds,
    items,
    rovingId,
    onNavigate,
    onRovingIdChange,
    onToggleCollapsed,
  },
  ref,
) {
  const keys = useMemo(() => items.map((item) => item.id), [items]);
  const pinnedKeys = useMemo(() => [activeItemId], [activeItemId]);
  const navigation = useDocumentNavigationWindow({
    estimatedItemHeight: DOCUMENT_OUTLINE_ITEM_HEIGHT,
    itemGap: 0,
    keys,
    listPaddingTop: DOCUMENT_OUTLINE_LIST_PADDING_TOP,
    onRovingKeyChange: onRovingIdChange,
    pinnedKeys,
    rovingKey: rovingId,
  });

  useImperativeHandle(
    ref,
    () => ({
      focusFirst: () => navigation.focusAt(0),
      focusLast: () => navigation.focusAt(items.length - 1),
    }),
    [items.length, navigation.focusAt],
  );

  const handleItemKeyDown = (
    event: KeyboardEvent<HTMLButtonElement>,
    item: WorkDocumentOutlineItem,
    index: number,
  ) => {
    if (event.key === 'ArrowDown') {
      event.preventDefault();
      navigation.focusAt(Math.min(items.length - 1, index + 1));
    } else if (event.key === 'ArrowUp') {
      event.preventDefault();
      navigation.focusAt(Math.max(0, index - 1));
    } else if (event.key === 'Home') {
      event.preventDefault();
      navigation.focusAt(0);
    } else if (event.key === 'End') {
      event.preventDefault();
      navigation.focusAt(items.length - 1);
    } else if (
      event.key === 'ArrowRight' &&
      item.hasChildren &&
      collapsedIds.has(item.id)
    ) {
      event.preventDefault();
      onToggleCollapsed(item.id);
    } else if (event.key === 'ArrowRight' && item.hasChildren) {
      event.preventDefault();
      const child = items[index + 1];
      if (child?.depth > item.depth) navigation.focusAt(index + 1);
    } else if (event.key === 'ArrowLeft') {
      event.preventDefault();
      if (item.hasChildren && !collapsedIds.has(item.id)) {
        onToggleCollapsed(item.id);
      } else if (item.parentId) {
        navigation.focusAt(
          items.findIndex((candidate) => candidate.id === item.parentId),
        );
      }
    }
  };

  return (
    <nav
      ref={navigation.viewportRef}
      className="work-document-task-pane-body work-document-outline"
      aria-label="文档标题"
      data-document-navigation-collection="outline"
      data-document-navigation-item-count={items.length}
      data-document-navigation-mounted-count={navigation.mountedCount}
      data-document-navigation-window-end={navigation.range.end}
      data-document-navigation-window-limit={
        DOCUMENT_NAVIGATION_COLLECTION_WINDOW_LIMIT
      }
      data-document-navigation-window-start={navigation.range.start}
      data-document-navigation-windowed={
        navigation.range.windowed ? 'true' : 'false'
      }
      onScroll={navigation.onScroll}
    >
      {items.length ? (
        <ol>
          {navigation.entries.map((entry) => {
            if (entry.kind === 'spacer') {
              return (
                <DocumentNavigationWindowSpacer
                  entry={entry}
                  key={`spacer-${entry.start}-${entry.end}`}
                />
              );
            }
            const item = items[entry.index];
            if (!item) return null;
            const collapsed = collapsedIds.has(item.id);
            return (
              <li
                aria-posinset={entry.index + 1}
                aria-setsize={items.length}
                className={activeItemId === item.id ? 'active' : undefined}
                data-document-navigation-item={entry.index + 1}
                key={item.id}
                style={
                  {
                    '--work-document-outline-indent': `${item.depth * 14}px`,
                  } as CSSProperties
                }
              >
                {item.hasChildren ? (
                  <button
                    type="button"
                    tabIndex={-1}
                    className="work-document-outline-toggle"
                    aria-label={`${collapsed ? '展开' : '折叠'} ${item.text}`}
                    title={collapsed ? '展开下级标题' : '折叠下级标题'}
                    onClick={() => onToggleCollapsed(item.id)}
                  >
                    {collapsed ? (
                      <ChevronRight size={13} />
                    ) : (
                      <ChevronDown size={13} />
                    )}
                  </button>
                ) : (
                  <span className="work-document-outline-toggle-spacer" />
                )}
                <button
                  ref={(element) => navigation.registerItem(item.id, element)}
                  type="button"
                  className="work-document-outline-item"
                  tabIndex={navigation.rovingIndex === entry.index ? 0 : -1}
                  aria-current={
                    activeItemId === item.id ? 'location' : undefined
                  }
                  aria-expanded={item.hasChildren ? !collapsed : undefined}
                  title={item.text}
                  onFocus={() => navigation.onItemFocus(entry.index)}
                  onKeyDown={(event) =>
                    handleItemKeyDown(event, item, entry.index)
                  }
                  onClick={() => onNavigate(item.from)}
                >
                  {item.text}
                </button>
              </li>
            );
          })}
        </ol>
      ) : (
        <div className="work-document-outline-empty">
          应用标题样式后，可在这里快速跳转
        </div>
      )}
    </nav>
  );
});
