import type { Editor } from '@tiptap/core';
import { TextSelection } from '@tiptap/pm/state';
import { Search } from 'lucide-react';
import {
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
import { DocumentOutlineNavigation } from './document-outline-navigation';
import {
  DocumentPageNavigation,
  type DocumentNavigationPage,
} from './document-page-navigation';
import type { WorkDocumentPageThumbnailSource } from './document-page-thumbnail';
import {
  documentNavigationMatchId,
  DocumentSearchNavigation,
} from './document-search-navigation';
import { DocumentTaskPane } from './document-task-pane';
import type { DocumentNavigationListHandle } from './document-navigation-window';

type DocumentNavigationView = 'headings' | 'pages';

interface InstantDocumentScrollState {
  previousBehavior: string;
  restoreFrame: number | null;
}

const instantDocumentScrollStates = new WeakMap<
  HTMLElement,
  InstantDocumentScrollState
>();

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
  const outlineNavigationRef = useRef<DocumentNavigationListHandle>(null);
  const searchNavigationRef = useRef<DocumentNavigationListHandle>(null);
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
    const validResultIds = new Set(matches.map(documentNavigationMatchId));
    setRovingResultId((current) =>
      current && validResultIds.has(current)
        ? current
        : matches[0]
          ? documentNavigationMatchId(matches[0])
          : null,
    );
  }, [matches, normalizedQuery]);

  const toggleCollapsed = useCallback((itemId: string) => {
    setCollapsedIds((current) => {
      const next = new Set(current);
      if (next.has(itemId)) next.delete(itemId);
      else next.add(itemId);
      return next;
    });
  }, []);

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
    runWithInstantDocumentScroll(editor, () =>
      chain.setTextSelection(resolvedSelection).scrollIntoView().run(),
    );
    if (!modal) return;
    await onClose();
    requestAnimationFrame(() => {
      if (!editor.isDestroyed) {
        runWithInstantDocumentScroll(editor, () =>
          editor.chain().focus().scrollIntoView().run(),
        );
      }
    });
  };

  const selectMatch = (match: DocumentTextMatch) => {
    void navigateTo({ from: match.from, to: match.to });
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
                  const target = normalizedQuery
                    ? searchNavigationRef.current
                    : outlineNavigationRef.current;
                  if (!target) return;
                  event.preventDefault();
                  target.focusFirst();
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
            <DocumentSearchNavigation
              ref={searchNavigationRef}
              matches={matches}
              outline={outline}
              rovingId={rovingResultId}
              selectedMatchIndex={selectedMatchIndex}
              onRovingIdChange={setRovingResultId}
              onSelect={selectMatch}
            />
          ) : (
            <DocumentOutlineNavigation
              ref={outlineNavigationRef}
              activeItemId={activeItem?.id ?? null}
              collapsedIds={collapsedIds}
              items={visibleItems}
              rovingId={rovingId}
              onNavigate={(position) => void navigateTo(position)}
              onRovingIdChange={setRovingId}
              onToggleCollapsed={toggleCollapsed}
            />
          )}
        </div>
      )}
    </DocumentTaskPane>
  );
}

function runWithInstantDocumentScroll<T>(editor: Editor, run: () => T): T {
  const scrollSurface = editor.view.dom.closest<HTMLElement>(
    '.work-document-scroll',
  );
  if (!scrollSurface) return run();
  let state = instantDocumentScrollStates.get(scrollSurface);
  if (!state) {
    state = {
      previousBehavior: scrollSurface.style.scrollBehavior,
      restoreFrame: null,
    };
    instantDocumentScrollStates.set(scrollSurface, state);
  } else if (
    state.restoreFrame !== null &&
    typeof cancelAnimationFrame === 'function'
  ) {
    cancelAnimationFrame(state.restoreFrame);
    state.restoreFrame = null;
  }
  scrollSurface.style.scrollBehavior = 'auto';
  scrollSurface.getBoundingClientRect();
  try {
    return run();
  } finally {
    const restore = () => {
      if (instantDocumentScrollStates.get(scrollSurface) !== state) return;
      scrollSurface.style.scrollBehavior = state.previousBehavior;
      instantDocumentScrollStates.delete(scrollSurface);
    };
    if (typeof requestAnimationFrame === 'function') {
      state.restoreFrame = requestAnimationFrame(restore);
    } else {
      restore();
    }
  }
}
