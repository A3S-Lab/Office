import type { Editor } from '@tiptap/core';
import { ArrowDown, ArrowUp, Replace, ReplaceAll, Search } from 'lucide-react';
import {
  type KeyboardEvent,
  useEffect,
  useId,
  useMemo,
  useRef,
  useState,
} from 'react';
import { Button, IconButton } from '../../../design-system/primitives';
import { officeMessage } from '../../../i18n/office-locale';
import { documentTextMatches } from '../work-document-search';
import {
  registerDocumentFindHighlight,
  unregisterDocumentFindHighlight,
  updateDocumentFindHighlights,
} from './document-find-highlight';
import { DocumentTaskPane } from './document-task-pane';
import { OfficeTextField } from './office-controls';
import { useOfficeMessages } from './office-messages-context';

export type DocumentFindReplaceMode = 'find' | 'replace';

export function DocumentFindReplacePanel({
  editor,
  mode,
  focusRequest = 0,
  onModeChange,
  onReplaceText,
  onClose,
}: {
  editor: Editor;
  mode: DocumentFindReplaceMode;
  focusRequest?: number;
  onModeChange: (mode: DocumentFindReplaceMode) => void;
  onReplaceText: (from: number, to: number, replacement: string) => boolean;
  onClose: () => void;
}) {
  const messages = useOfficeMessages();
  const queryRef = useRef<HTMLInputElement>(null);
  const queryId = useId();
  const replacementId = useId();
  const [query, setQuery] = useState('');
  const [replacement, setReplacement] = useState('');
  const [activeIndex, setActiveIndex] = useState(-1);
  const [announcement, setAnnouncement] = useState('');
  const [, renderEditorState] = useState(0);

  useEffect(() => {
    const refresh = () => renderEditorState((current) => current + 1);
    editor.on('transaction', refresh);
    return () => {
      editor.off('transaction', refresh);
    };
  }, [editor]);

  const editorDocument = editor.state.doc;
  const matches = useMemo(
    () => documentTextMatches(editorDocument, query),
    [editor, editorDocument, query],
  );

  useEffect(() => {
    registerDocumentFindHighlight(editor);
    return () => unregisterDocumentFindHighlight(editor);
  }, [editor]);

  useEffect(() => {
    updateDocumentFindHighlights(editor, matches, activeIndex);
  }, [activeIndex, editor, matches]);

  useEffect(() => {
    queryRef.current?.focus({ preventScroll: true });
    queryRef.current?.select();
  }, [focusRequest, mode]);

  useEffect(() => {
    setActiveIndex((current) =>
      current >= 0 && matches.length
        ? Math.min(current, matches.length - 1)
        : -1,
    );
  }, [matches.length]);

  useEffect(() => {
    setActiveIndex(-1);
    setAnnouncement('');
  }, [query]);

  const selectMatch = (requestedIndex: number) => {
    if (!matches.length) return;
    const index = (requestedIndex + matches.length * 2) % matches.length;
    const match = matches[index];
    if (!match) return;
    setActiveIndex(index);
    editor
      .chain()
      .setTextSelection({ from: match.from, to: match.to })
      .scrollIntoView()
      .run();
    setAnnouncement(
      officeMessage(messages, 'document.find.announce.position', {
        current: String(index + 1),
        total: String(matches.length),
      }),
    );
  };

  const moveToMatch = (direction: -1 | 1) => {
    selectMatch(
      activeIndex < 0
        ? direction > 0
          ? 0
          : matches.length - 1
        : activeIndex + direction,
    );
  };

  const replaceCurrent = () => {
    const match = matches[activeIndex >= 0 ? activeIndex : 0];
    const focusTarget =
      document.activeElement instanceof HTMLElement
        ? document.activeElement
        : null;
    if (!match || !onReplaceText(match.from, match.to, replacement)) return;
    setAnnouncement(
      officeMessage(messages, 'document.find.announce.replaced'),
    );
    const remaining = documentTextMatches(editor.state.doc, query);
    const next =
      remaining[Math.min(Math.max(activeIndex, 0), remaining.length - 1)];
    if (next) {
      editor
        .chain()
        .setTextSelection({ from: next.from, to: next.to })
        .scrollIntoView()
        .run();
    }
    if (!remaining.length && focusTarget instanceof HTMLButtonElement) {
      queryRef.current?.focus({ preventScroll: true });
      queryRef.current?.select();
      return;
    }
    if (focusTarget?.isConnected && !focusTarget.matches(':disabled')) {
      focusTarget.focus({ preventScroll: true });
    } else {
      queryRef.current?.focus({ preventScroll: true });
    }
  };

  const replaceAll = () => {
    let replaced = 0;
    for (const match of [...matches].reverse()) {
      if (onReplaceText(match.from, match.to, replacement)) replaced += 1;
    }
    setAnnouncement(
      replaced
        ? officeMessage(messages, 'document.find.announce.replacedAll', {
            count: String(replaced),
          })
        : officeMessage(messages, 'document.find.announce.none'),
    );
    setActiveIndex(-1);
    queryRef.current?.focus({ preventScroll: true });
  };

  const handleCommandKeyDown = (event: KeyboardEvent<HTMLElement>) => {
    if (!(event.metaKey || event.ctrlKey) || event.altKey) return;
    const key = event.key.toLocaleLowerCase();
    if (key !== 'f' && key !== 'h') return;
    event.preventDefault();
    event.stopPropagation();
    onModeChange(key === 'h' ? 'replace' : 'find');
    requestAnimationFrame(() => {
      queryRef.current?.focus({ preventScroll: true });
      queryRef.current?.select();
    });
  };

  return (
    <DocumentTaskPane
      className="work-document-find-panel"
      title={officeMessage(
        messages,
        mode === 'replace'
          ? 'document.find.titleReplace'
          : 'document.find.title',
      )}
      description={officeMessage(messages, 'document.find.description')}
      closeLabel={officeMessage(messages, 'document.find.close')}
      onClose={onClose}
      onKeyDown={handleCommandKeyDown}
    >
      <div
        className="work-document-find-tabs"
        role="tablist"
        aria-label={officeMessage(messages, 'document.find.modeAria')}
      >
        <button
          type="button"
          role="tab"
          aria-selected={mode === 'find'}
          onClick={() => onModeChange('find')}
        >
          {officeMessage(messages, 'document.find.tabFind')}
        </button>
        <button
          type="button"
          role="tab"
          aria-selected={mode === 'replace'}
          onClick={() => onModeChange('replace')}
        >
          {officeMessage(messages, 'document.find.tabReplace')}
        </button>
      </div>
      <div className="work-document-task-pane-body work-document-find-body">
        <label htmlFor={queryId}>
          <span>{officeMessage(messages, 'document.find.query')}</span>
          <span className="work-document-find-field">
            <Search size={14} aria-hidden="true" />
            <OfficeTextField
              id={queryId}
              ref={queryRef}
              data-document-find-query=""
              aria-label={officeMessage(messages, 'document.find.query')}
              value={query}
              onChange={(event) => setQuery(event.target.value)}
              onKeyDown={(event) => {
                if (event.key !== 'Enter') return;
                event.preventDefault();
                moveToMatch(event.shiftKey ? -1 : 1);
              }}
            />
          </span>
        </label>
        {mode === 'replace' && (
          <label htmlFor={replacementId}>
            <span>{officeMessage(messages, 'document.find.replacement')}</span>
            <OfficeTextField
              id={replacementId}
              data-document-find-replacement=""
              aria-label={officeMessage(messages, 'document.find.replacement')}
              value={replacement}
              onChange={(event) => setReplacement(event.target.value)}
              onKeyDown={(event) => {
                if (event.key === 'Enter') {
                  event.preventDefault();
                  replaceCurrent();
                }
              }}
            />
          </label>
        )}
        <div className="work-document-find-results">
          <output aria-live="polite">
            {!query
              ? officeMessage(messages, 'document.find.prompt')
              : matches.length
                ? officeMessage(messages, 'document.find.matchCount', {
                    count: String(matches.length),
                  })
                : officeMessage(messages, 'document.find.noMatches')}
          </output>
          <div>
            <IconButton
              label={officeMessage(messages, 'document.find.previous')}
              disabled={!matches.length}
              onClick={() => moveToMatch(-1)}
            >
              <ArrowUp size={14} />
            </IconButton>
            <IconButton
              label={officeMessage(messages, 'document.find.next')}
              disabled={!matches.length}
              onClick={() => moveToMatch(1)}
            >
              <ArrowDown size={14} />
            </IconButton>
          </div>
        </div>
        {mode === 'replace' && (
          <div className="work-document-find-actions">
            <Button
              size="compact"
              tone="secondary"
              disabled={!matches.length}
              onClick={replaceCurrent}
            >
              <Replace size={13} />
              {officeMessage(messages, 'document.find.replace')}
            </Button>
            <Button
              size="compact"
              tone="primary"
              disabled={!matches.length}
              onClick={replaceAll}
            >
              <ReplaceAll size={13} />
              {officeMessage(messages, 'document.find.replaceAll')}
            </Button>
          </div>
        )}
        {announcement && (
          <output
            className="work-document-find-announcement"
            aria-live="polite"
          >
            {announcement}
          </output>
        )}
      </div>
    </DocumentTaskPane>
  );
}
