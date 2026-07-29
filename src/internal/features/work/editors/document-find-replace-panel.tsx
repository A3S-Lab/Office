import type { Editor } from '@tiptap/core';
import { ArrowDown, ArrowUp, Replace, ReplaceAll, Search } from 'lucide-react';
import { useEffect, useId, useMemo, useRef, useState } from 'react';
import { Button, IconButton } from '../../../design-system/primitives';
import { OfficeTextField } from './office-controls';
import { DocumentTaskPane } from './document-task-pane';

export type DocumentFindReplaceMode = 'find' | 'replace';

interface DocumentTextMatch {
  from: number;
  to: number;
}

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
  const queryRef = useRef<HTMLInputElement>(null);
  const queryId = useId();
  const replacementId = useId();
  const [query, setQuery] = useState('');
  const [replacement, setReplacement] = useState('');
  const [activeIndex, setActiveIndex] = useState(-1);
  const [announcement, setAnnouncement] = useState('');
  const document = editor.state.doc;
  const matches = useMemo(
    () => documentTextMatches(editor, query),
    [document, editor, query],
  );

  useEffect(() => {
    queryRef.current?.focus({ preventScroll: true });
    queryRef.current?.select();
  }, [focusRequest, mode]);

  useEffect(() => {
    setActiveIndex((current) =>
      matches.length ? Math.min(Math.max(current, 0), matches.length - 1) : -1,
    );
  }, [matches.length]);

  useEffect(() => setAnnouncement(''), [query]);

  const selectMatch = (requestedIndex: number) => {
    if (!matches.length) return;
    const index = (requestedIndex + matches.length * 2) % matches.length;
    const match = matches[index];
    if (!match) return;
    setActiveIndex(index);
    editor
      .chain()
      .focus()
      .setTextSelection({ from: match.from, to: match.to })
      .run();
    setAnnouncement(`第 ${index + 1} 个，共 ${matches.length} 个`);
  };

  const replaceCurrent = () => {
    const match = matches[activeIndex >= 0 ? activeIndex : 0];
    if (!match || !onReplaceText(match.from, match.to, replacement)) return;
    setAnnouncement('已替换当前匹配');
    const remaining = documentTextMatches(editor, query);
    const next =
      remaining[Math.min(Math.max(activeIndex, 0), remaining.length - 1)];
    if (next) {
      editor
        .chain()
        .focus()
        .setTextSelection({ from: next.from, to: next.to })
        .run();
    } else {
      queryRef.current?.focus({ preventScroll: true });
    }
  };

  const replaceAll = () => {
    let replaced = 0;
    for (const match of [...matches].reverse()) {
      if (onReplaceText(match.from, match.to, replacement)) replaced += 1;
    }
    setAnnouncement(replaced ? `已替换 ${replaced} 处` : '没有可替换的内容');
    setActiveIndex(-1);
    queryRef.current?.focus({ preventScroll: true });
  };

  return (
    <DocumentTaskPane
      className="work-document-find-panel"
      title={mode === 'replace' ? '查找和替换' : '查找'}
      description="在当前文档中定位文字"
      closeLabel="关闭查找"
      onClose={onClose}
    >
      <div
        className="work-document-find-tabs"
        role="tablist"
        aria-label="查找方式"
      >
        <button
          type="button"
          role="tab"
          aria-selected={mode === 'find'}
          onClick={() => onModeChange('find')}
        >
          查找
        </button>
        <button
          type="button"
          role="tab"
          aria-selected={mode === 'replace'}
          onClick={() => onModeChange('replace')}
        >
          替换
        </button>
      </div>
      <div className="work-document-task-pane-body work-document-find-body">
        <label htmlFor={queryId}>
          <span>查找内容</span>
          <span className="work-document-find-field">
            <Search size={14} aria-hidden="true" />
            <OfficeTextField
              id={queryId}
              ref={queryRef}
              aria-label="查找内容"
              value={query}
              onChange={(event) => setQuery(event.target.value)}
              onKeyDown={(event) => {
                if (event.key !== 'Enter') return;
                event.preventDefault();
                selectMatch(activeIndex + (event.shiftKey ? -1 : 1));
              }}
            />
          </span>
        </label>
        {mode === 'replace' && (
          <label htmlFor={replacementId}>
            <span>替换为</span>
            <OfficeTextField
              id={replacementId}
              aria-label="替换为"
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
              ? '输入文字开始查找'
              : matches.length
                ? `${matches.length} 个匹配`
                : '没有匹配内容'}
          </output>
          <div>
            <IconButton
              label="上一个匹配"
              disabled={!matches.length}
              onClick={() => selectMatch(activeIndex - 1)}
            >
              <ArrowUp size={14} />
            </IconButton>
            <IconButton
              label="下一个匹配"
              disabled={!matches.length}
              onClick={() => selectMatch(activeIndex + 1)}
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
              替换
            </Button>
            <Button
              size="compact"
              tone="primary"
              disabled={!matches.length}
              onClick={replaceAll}
            >
              <ReplaceAll size={13} />
              全部替换
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

export function documentTextMatches(
  editor: Editor,
  rawQuery: string,
): DocumentTextMatch[] {
  const query = rawQuery.toLocaleLowerCase();
  if (!query) return [];
  const runs: Array<{ from: number; text: string }> = [];
  editor.state.doc.descendants((node, position) => {
    if (!node.isText || !node.text) return;
    const previous = runs.at(-1);
    if (previous && previous.from + previous.text.length === position) {
      previous.text += node.text;
    } else {
      runs.push({ from: position, text: node.text });
    }
  });
  return runs.flatMap((run) => {
    const matches: DocumentTextMatch[] = [];
    const text = run.text.toLocaleLowerCase();
    let offset = 0;
    while (offset <= text.length - query.length) {
      const index = text.indexOf(query, offset);
      if (index < 0) break;
      matches.push({
        from: run.from + index,
        to: run.from + index + query.length,
      });
      offset = index + Math.max(1, query.length);
    }
    return matches;
  });
}
