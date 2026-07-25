import type { Editor } from '@tiptap/core';
import { ChevronDown, ChevronRight, Search } from 'lucide-react';
import {
  type CSSProperties,
  type KeyboardEvent,
  useCallback,
  useEffect,
  useMemo,
  useRef,
  useState,
} from 'react';
import {
  collectWorkDocumentOutline,
  currentWorkDocumentOutlineItem,
  visibleWorkDocumentOutlineItems,
  type WorkDocumentOutlineItem,
} from '../work-document-outline';
import { OfficeTextField } from './office-controls';
import { DocumentTaskPane } from './document-task-pane';

export function DocumentNavigationPanel({
  editor,
  onClose,
}: {
  editor: Editor;
  onClose: () => void;
}) {
  const searchRef = useRef<HTMLInputElement>(null);
  const itemRefs = useRef(new Map<string, HTMLButtonElement>());
  const [, renderEditorState] = useState(0);
  const [query, setQuery] = useState('');
  const [collapsedIds, setCollapsedIds] = useState<Set<string>>(
    () => new Set(),
  );
  const [rovingId, setRovingId] = useState<string | null>(null);

  useEffect(() => {
    const refresh = () => renderEditorState((current) => current + 1);
    editor.on('transaction', refresh);
    return () => {
      editor.off('transaction', refresh);
    };
  }, [editor]);

  const document = editor.state.doc;
  const outline = useMemo(
    () => collectWorkDocumentOutline(document),
    [document],
  );
  const activeItem = currentWorkDocumentOutlineItem(
    outline,
    editor.state.selection.from,
  );
  const visibleItems = useMemo(
    () => visibleWorkDocumentOutlineItems(outline, collapsedIds, query),
    [collapsedIds, outline, query],
  );

  useEffect(() => {
    searchRef.current?.focus({ preventScroll: true });
  }, []);

  useEffect(() => {
    const validIds = new Set(
      outline.filter((item) => item.hasChildren).map((item) => item.id),
    );
    setCollapsedIds((current) => {
      const next = new Set(
        [...current].filter((itemId) => validIds.has(itemId)),
      );
      return next.size === current.size ? current : next;
    });
  }, [outline]);

  useEffect(() => {
    if (visibleItems.some((item) => item.id === rovingId)) return;
    const nextId =
      visibleItems.find((item) => item.id === activeItem?.id)?.id ??
      visibleItems[0]?.id ??
      null;
    setRovingId(nextId);
  }, [activeItem?.id, rovingId, visibleItems]);

  const focusItem = useCallback((itemId: string) => {
    setRovingId(itemId);
    requestAnimationFrame(() => {
      itemRefs.current.get(itemId)?.focus({ preventScroll: true });
    });
  }, []);

  const toggleCollapsed = useCallback((itemId: string) => {
    setCollapsedIds((current) => {
      const next = new Set(current);
      if (next.has(itemId)) next.delete(itemId);
      else next.add(itemId);
      return next;
    });
  }, []);

  const handleItemKeyDown = (
    event: KeyboardEvent<HTMLButtonElement>,
    item: WorkDocumentOutlineItem,
  ) => {
    const index = visibleItems.findIndex(
      (candidate) => candidate.id === item.id,
    );
    const focusAt = (requestedIndex: number) => {
      const next = visibleItems[requestedIndex];
      if (next) focusItem(next.id);
    };
    if (event.key === 'ArrowDown') {
      event.preventDefault();
      focusAt(Math.min(visibleItems.length - 1, index + 1));
    } else if (event.key === 'ArrowUp') {
      event.preventDefault();
      focusAt(Math.max(0, index - 1));
    } else if (event.key === 'Home') {
      event.preventDefault();
      focusAt(0);
    } else if (event.key === 'End') {
      event.preventDefault();
      focusAt(visibleItems.length - 1);
    } else if (
      event.key === 'ArrowRight' &&
      !query &&
      item.hasChildren &&
      collapsedIds.has(item.id)
    ) {
      event.preventDefault();
      toggleCollapsed(item.id);
    } else if (event.key === 'ArrowRight' && !query && item.hasChildren) {
      event.preventDefault();
      const child = visibleItems[index + 1];
      if (child?.depth > item.depth) focusItem(child.id);
    } else if (event.key === 'ArrowLeft' && !query) {
      event.preventDefault();
      if (item.hasChildren && !collapsedIds.has(item.id)) {
        toggleCollapsed(item.id);
      } else if (item.parentId) {
        focusItem(item.parentId);
      }
    }
  };

  return (
    <DocumentTaskPane
      ariaLabel="文档导航"
      className="work-document-navigation-panel"
      title="导航窗格"
      description="按标题浏览文档"
      closeLabel="关闭导航窗格"
      onClose={onClose}
    >
      <div className="work-document-navigation-search">
        <Search size={14} aria-hidden="true" />
        <OfficeTextField
          ref={searchRef}
          type="search"
          aria-label="搜索标题"
          placeholder="搜索标题"
          value={query}
          onChange={(event) => setQuery(event.target.value)}
          onKeyDown={(event) => {
            if (event.key !== 'ArrowDown' || !visibleItems.length) return;
            event.preventDefault();
            focusItem(visibleItems[0]?.id ?? '');
          }}
        />
      </div>
      <div className="work-document-navigation-summary" aria-live="polite">
        {query ? `${visibleItems.length} 个匹配` : `${outline.length} 个标题`}
      </div>
      <nav
        className="work-document-task-pane-body work-document-outline"
        aria-label="文档标题"
      >
        {visibleItems.length ? (
          <ol>
            {visibleItems.map((item) => {
              const collapsed = collapsedIds.has(item.id);
              return (
                <li
                  key={item.id}
                  className={activeItem?.id === item.id ? 'active' : undefined}
                  style={
                    {
                      '--work-document-outline-indent': `${item.depth * 14}px`,
                    } as CSSProperties
                  }
                >
                  {!query && item.hasChildren ? (
                    <button
                      type="button"
                      tabIndex={-1}
                      className="work-document-outline-toggle"
                      aria-label={`${collapsed ? '展开' : '折叠'} ${item.text}`}
                      title={collapsed ? '展开下级标题' : '折叠下级标题'}
                      onClick={() => toggleCollapsed(item.id)}
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
                    ref={(element) => {
                      if (element) itemRefs.current.set(item.id, element);
                      else itemRefs.current.delete(item.id);
                    }}
                    type="button"
                    className="work-document-outline-item"
                    tabIndex={item.id === rovingId ? 0 : -1}
                    aria-current={
                      activeItem?.id === item.id ? 'location' : undefined
                    }
                    aria-expanded={
                      !query && item.hasChildren ? !collapsed : undefined
                    }
                    title={item.text}
                    onFocus={() => setRovingId(item.id)}
                    onKeyDown={(event) => handleItemKeyDown(event, item)}
                    onClick={() => {
                      editor
                        .chain()
                        .focus()
                        .setTextSelection(item.from)
                        .scrollIntoView()
                        .run();
                    }}
                  >
                    {item.text}
                  </button>
                </li>
              );
            })}
          </ol>
        ) : (
          <div className="work-document-outline-empty">
            {outline.length
              ? '没有匹配的标题'
              : '应用标题样式后，可在这里快速跳转'}
          </div>
        )}
      </nav>
    </DocumentTaskPane>
  );
}
