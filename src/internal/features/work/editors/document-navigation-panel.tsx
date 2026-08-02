import type { Editor } from '@tiptap/core';
import { TextSelection } from '@tiptap/pm/state';
import { ChevronDown, ChevronRight, Search } from 'lucide-react';
import {
  type CSSProperties,
  type KeyboardEvent,
  useCallback,
  useEffect,
  useId,
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
import {
  documentTextMatches,
  type DocumentTextMatch,
} from '../work-document-search';
import { Tabs } from '../../../design-system/primitives';
import {
  registerDocumentFindHighlight,
  unregisterDocumentFindHighlight,
  updateDocumentFindHighlights,
} from './document-find-highlight';
import { OfficeTextField } from './office-controls';
import {
  DocumentPageNavigation,
  type DocumentNavigationPage,
} from './document-page-navigation';
import type { WorkDocumentPageThumbnailSource } from './document-page-thumbnail';
import { DocumentTaskPane } from './document-task-pane';

type DocumentNavigationView = 'headings' | 'pages';

export function DocumentNavigationPanel({
  currentPage = 1,
  editor,
  modal = false,
  pages = [],
  pageThumbnailSource,
  onClose,
}: {
  currentPage?: number;
  editor: Editor;
  modal?: boolean;
  pages?: readonly DocumentNavigationPage[];
  pageThumbnailSource?: WorkDocumentPageThumbnailSource;
  onClose: () => void | Promise<void>;
}) {
  const tabsId = useId();
  const searchRef = useRef<HTMLInputElement>(null);
  const itemRefs = useRef(new Map<string, HTMLButtonElement>());
  const resultRefs = useRef(new Map<string, HTMLButtonElement>());
  const [, renderEditorState] = useState(0);
  const [navigationView, setNavigationView] =
    useState<DocumentNavigationView>('headings');
  const [query, setQuery] = useState('');
  const [collapsedIds, setCollapsedIds] = useState<Set<string>>(
    () => new Set(),
  );
  const [rovingId, setRovingId] = useState<string | null>(null);
  const [rovingResultId, setRovingResultId] = useState<string | null>(null);

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
    () => visibleWorkDocumentOutlineItems(outline, collapsedIds, ''),
    [collapsedIds, outline],
  );
  const normalizedQuery = query.trim();
  const matches = useMemo(
    () => documentTextMatches(document, normalizedQuery),
    [document, normalizedQuery],
  );
  const selectedMatchIndex = matches.findIndex(
    (match) =>
      editor.state.selection.from === match.from &&
      editor.state.selection.to === match.to,
  );

  useEffect(() => {
    searchRef.current?.focus({ preventScroll: true });
  }, []);

  useEffect(() => {
    registerDocumentFindHighlight(editor);
    return () => unregisterDocumentFindHighlight(editor);
  }, [editor]);

  useEffect(() => {
    updateDocumentFindHighlights(editor, matches, selectedMatchIndex);
  }, [editor, matches, selectedMatchIndex]);

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

  useEffect(() => {
    if (!normalizedQuery) return;
    const validResultIds = new Set(matches.map(documentMatchId));
    setRovingResultId((current) =>
      current && validResultIds.has(current)
        ? current
        : matches[0]
          ? documentMatchId(matches[0])
          : null,
    );
  }, [matches, normalizedQuery]);

  const focusItem = useCallback((itemId: string) => {
    setRovingId(itemId);
    requestAnimationFrame(() => {
      itemRefs.current.get(itemId)?.focus({ preventScroll: true });
    });
  }, []);

  const focusResult = useCallback((resultId: string) => {
    setRovingResultId(resultId);
    requestAnimationFrame(() => {
      resultRefs.current.get(resultId)?.focus({ preventScroll: true });
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
      !normalizedQuery &&
      item.hasChildren &&
      collapsedIds.has(item.id)
    ) {
      event.preventDefault();
      toggleCollapsed(item.id);
    } else if (
      event.key === 'ArrowRight' &&
      !normalizedQuery &&
      item.hasChildren
    ) {
      event.preventDefault();
      const child = visibleItems[index + 1];
      if (child?.depth > item.depth) focusItem(child.id);
    } else if (event.key === 'ArrowLeft' && !normalizedQuery) {
      event.preventDefault();
      if (item.hasChildren && !collapsedIds.has(item.id)) {
        toggleCollapsed(item.id);
      } else if (item.parentId) {
        focusItem(item.parentId);
      }
    }
  };

  const navigateTo = async (
    selection: number | { from: number; to: number },
    focusEditor = true,
  ) => {
    const resolvedSelection =
      typeof selection === 'number'
        ? TextSelection.near(
            editor.state.doc.resolve(
              Math.min(editor.state.doc.content.size, Math.max(0, selection)),
            ),
            1,
          ).from
        : selection;
    const chain = editor.chain();
    if (!modal && focusEditor) chain.focus();
    chain.setTextSelection(resolvedSelection).scrollIntoView().run();
    if (!modal) return;
    await onClose();
    requestAnimationFrame(() => {
      if (!editor.isDestroyed) editor.chain().focus().scrollIntoView().run();
    });
  };

  const selectMatch = (match: DocumentTextMatch) => {
    void navigateTo({ from: match.from, to: match.to });
  };

  const handleResultKeyDown = (
    event: KeyboardEvent<HTMLButtonElement>,
    match: DocumentTextMatch,
  ) => {
    const index = matches.findIndex(
      (candidate) => documentMatchId(candidate) === documentMatchId(match),
    );
    const focusAt = (requestedIndex: number) => {
      const next = matches[requestedIndex];
      if (next) focusResult(documentMatchId(next));
    };
    if (event.key === 'ArrowDown') {
      event.preventDefault();
      focusAt(Math.min(matches.length - 1, index + 1));
    } else if (event.key === 'ArrowUp') {
      event.preventDefault();
      focusAt(Math.max(0, index - 1));
    } else if (event.key === 'Home') {
      event.preventDefault();
      focusAt(0);
    } else if (event.key === 'End') {
      event.preventDefault();
      focusAt(matches.length - 1);
    }
  };

  return (
    <DocumentTaskPane
      ariaLabel="文档导航"
      className="work-document-navigation-panel"
      title="导航窗格"
      description="浏览标题、页面或搜索正文"
      closeLabel="关闭导航窗格"
      onClose={onClose}
    >
      <Tabs
        ariaLabel="导航视图"
        className="work-document-navigation-tabs"
        value={navigationView}
        variant="line"
        size="compact"
        items={[
          {
            id: 'headings',
            label: '标题',
            tabId: `${tabsId}-headings-tab`,
            panelId: `${tabsId}-headings-panel`,
          },
          {
            id: 'pages',
            label: '页面',
            tabId: `${tabsId}-pages-tab`,
            panelId: `${tabsId}-pages-panel`,
          },
        ]}
        onChange={setNavigationView}
      />
      {navigationView === 'pages' ? (
        <div
          id={`${tabsId}-pages-panel`}
          className="work-document-navigation-tab-panel"
          role="tabpanel"
          aria-labelledby={`${tabsId}-pages-tab`}
        >
          <div className="work-document-navigation-summary" aria-live="polite">
            {pages.length ? `${pages.length} 页` : '正在生成页面'}
          </div>
          <DocumentPageNavigation
            currentPage={currentPage}
            pages={pages}
            thumbnailSource={pageThumbnailSource}
            onSelectPage={(page) => navigateTo(page.selectionPosition, false)}
          />
        </div>
      ) : (
        <div
          id={`${tabsId}-headings-panel`}
          className="work-document-navigation-tab-panel"
          role="tabpanel"
          aria-labelledby={`${tabsId}-headings-tab`}
        >
          <div className="work-document-navigation-search">
            <Search size={14} aria-hidden="true" />
            <OfficeTextField
              ref={searchRef}
              type="search"
              aria-label="搜索文档"
              placeholder="搜索文档"
              value={query}
              onChange={(event) => setQuery(event.target.value)}
              onKeyDown={(event) => {
                if (event.key === 'Enter' && matches[0]) {
                  event.preventDefault();
                  selectMatch(matches[0]);
                } else if (event.key === 'ArrowDown') {
                  const firstTarget = normalizedQuery
                    ? matches[0]
                      ? documentMatchId(matches[0])
                      : null
                    : visibleItems[0]?.id;
                  if (!firstTarget) return;
                  event.preventDefault();
                  if (normalizedQuery) focusResult(firstTarget);
                  else focusItem(firstTarget);
                }
              }}
            />
          </div>
          <div className="work-document-navigation-summary" aria-live="polite">
            {normalizedQuery
              ? matches.length
                ? `${matches.length} 个匹配`
                : '没有匹配内容'
              : `${outline.length} 个标题`}
          </div>
          {normalizedQuery ? (
            <nav
              className="work-document-task-pane-body work-document-search-results"
              aria-label="文档搜索结果"
            >
              {matches.length ? (
                <ol>
                  {matches.map((match, index) => {
                    const resultId = documentMatchId(match);
                    const section =
                      currentWorkDocumentOutlineItem(outline, match.from)
                        ?.text ?? '文档开头';
                    return (
                      <li key={resultId}>
                        <button
                          ref={(element) => {
                            if (element)
                              resultRefs.current.set(resultId, element);
                            else resultRefs.current.delete(resultId);
                          }}
                          type="button"
                          className="work-document-search-result"
                          tabIndex={resultId === rovingResultId ? 0 : -1}
                          aria-label={`第 ${index + 1} 个匹配：${match.matchedText}`}
                          aria-current={
                            selectedMatchIndex === index
                              ? 'location'
                              : undefined
                          }
                          onFocus={() => setRovingResultId(resultId)}
                          onKeyDown={(event) =>
                            handleResultKeyDown(event, match)
                          }
                          onClick={() => selectMatch(match)}
                        >
                          <span className="work-document-search-section">
                            {section}
                          </span>
                          <span className="work-document-search-excerpt">
                            {match.truncatedBefore && '…'}
                            {match.before}
                            <mark>{match.matchedText}</mark>
                            {match.after}
                            {match.truncatedAfter && '…'}
                          </span>
                        </button>
                      </li>
                    );
                  })}
                </ol>
              ) : (
                <div className="work-document-outline-empty">
                  尝试更短或不同的文字
                </div>
              )}
            </nav>
          ) : (
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
                        className={
                          activeItem?.id === item.id ? 'active' : undefined
                        }
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
                            item.hasChildren ? !collapsed : undefined
                          }
                          title={item.text}
                          onFocus={() => setRovingId(item.id)}
                          onKeyDown={(event) => handleItemKeyDown(event, item)}
                          onClick={() => {
                            void navigateTo(item.from);
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
                  应用标题样式后，可在这里快速跳转
                </div>
              )}
            </nav>
          )}
        </div>
      )}
    </DocumentTaskPane>
  );
}

function documentMatchId(match: DocumentTextMatch): string {
  return `match-${match.from}-${match.to}`;
}
