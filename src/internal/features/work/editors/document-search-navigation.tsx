import {
  type KeyboardEvent,
  forwardRef,
  useImperativeHandle,
  useMemo,
} from 'react';
import {
  currentWorkDocumentOutlineItem,
  type WorkDocumentOutlineItem,
} from '../work-document-outline';
import type { DocumentTextMatch } from '../work-document-search';
import {
  DOCUMENT_NAVIGATION_COLLECTION_WINDOW_LIMIT,
  DocumentNavigationWindowSpacer,
  type DocumentNavigationListHandle,
  useDocumentNavigationWindow,
} from './document-navigation-window';

const DOCUMENT_SEARCH_RESULT_HEIGHT = 56;
const DOCUMENT_SEARCH_RESULT_GAP = 5;
const DOCUMENT_SEARCH_LIST_PADDING_TOP = 7;

export const DocumentSearchNavigation = forwardRef<
  DocumentNavigationListHandle,
  {
    matches: readonly DocumentTextMatch[];
    outline: readonly WorkDocumentOutlineItem[];
    rovingId: string | null;
    selectedMatchIndex: number;
    onRovingIdChange: (resultId: string) => void;
    onSelect: (match: DocumentTextMatch) => void;
  }
>(function DocumentSearchNavigation(
  {
    matches,
    outline,
    rovingId,
    selectedMatchIndex,
    onRovingIdChange,
    onSelect,
  },
  ref,
) {
  const keys = useMemo(() => matches.map(documentNavigationMatchId), [matches]);
  const selectedMatch = matches[selectedMatchIndex];
  const selectedKey = selectedMatch
    ? documentNavigationMatchId(selectedMatch)
    : null;
  const pinnedKeys = useMemo(() => [selectedKey], [selectedKey]);
  const navigation = useDocumentNavigationWindow({
    estimatedItemHeight: DOCUMENT_SEARCH_RESULT_HEIGHT,
    itemGap: DOCUMENT_SEARCH_RESULT_GAP,
    keys,
    listPaddingTop: DOCUMENT_SEARCH_LIST_PADDING_TOP,
    onRovingKeyChange: onRovingIdChange,
    pinnedKeys,
    rovingKey: rovingId,
  });

  useImperativeHandle(
    ref,
    () => ({
      focusFirst: () => navigation.focusAt(0),
      focusLast: () => navigation.focusAt(matches.length - 1),
    }),
    [matches.length, navigation.focusAt],
  );

  const handleResultKeyDown = (
    event: KeyboardEvent<HTMLButtonElement>,
    index: number,
  ) => {
    if (event.key === 'ArrowDown') {
      event.preventDefault();
      navigation.focusAt(Math.min(matches.length - 1, index + 1));
    } else if (event.key === 'ArrowUp') {
      event.preventDefault();
      navigation.focusAt(Math.max(0, index - 1));
    } else if (event.key === 'Home') {
      event.preventDefault();
      navigation.focusAt(0);
    } else if (event.key === 'End') {
      event.preventDefault();
      navigation.focusAt(matches.length - 1);
    }
  };

  return (
    <nav
      ref={navigation.viewportRef}
      className="work-document-task-pane-body work-document-search-results"
      aria-label="文档搜索结果"
      data-document-navigation-collection="search"
      data-document-navigation-item-count={matches.length}
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
      {matches.length ? (
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
            const match = matches[entry.index];
            if (!match) return null;
            const resultId = documentNavigationMatchId(match);
            const section =
              currentWorkDocumentOutlineItem(outline, match.from)?.text ??
              '文档开头';
            return (
              <li
                aria-posinset={entry.index + 1}
                aria-setsize={matches.length}
                data-document-navigation-item={entry.index + 1}
                key={resultId}
              >
                <button
                  ref={(element) => navigation.registerItem(resultId, element)}
                  type="button"
                  className="work-document-search-result"
                  tabIndex={navigation.rovingIndex === entry.index ? 0 : -1}
                  aria-label={`第 ${entry.index + 1} 个匹配：${match.matchedText}`}
                  aria-current={
                    selectedMatchIndex === entry.index ? 'location' : undefined
                  }
                  onFocus={() => navigation.onItemFocus(entry.index)}
                  onKeyDown={(event) => handleResultKeyDown(event, entry.index)}
                  onClick={() => onSelect(match)}
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
        <div className="work-document-outline-empty">尝试更短或不同的文字</div>
      )}
    </nav>
  );
});

export function documentNavigationMatchId(match: DocumentTextMatch): string {
  return `match-${match.from}-${match.to}`;
}
