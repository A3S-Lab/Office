import { ArrowDown, ArrowUp, Search, X } from 'lucide-react';
import { useEffect, useMemo, useRef, useState } from 'react';
import { IconButton } from '../../../design-system/primitives';
import type { WorkSpreadsheetContent } from '../work-types';
import { OfficeTextField } from './office-controls';
import {
  spreadsheetFindMatches,
  type SpreadsheetFindMatch,
} from './spreadsheet-find';

export function SpreadsheetFindBar({
  sheet,
  focusRequest,
  onClose,
  onSelectMatch,
}: {
  sheet: WorkSpreadsheetContent['sheets'][number] | undefined;
  focusRequest: number;
  onClose: () => void;
  onSelectMatch: (match: SpreadsheetFindMatch) => void;
}) {
  const queryRef = useRef<HTMLInputElement>(null);
  const [query, setQuery] = useState('');
  const [activeIndex, setActiveIndex] = useState(-1);
  const matches = useMemo(
    () => spreadsheetFindMatches(sheet, query),
    [query, sheet],
  );

  useEffect(() => {
    queryRef.current?.focus({ preventScroll: true });
    queryRef.current?.select();
  }, [focusRequest]);

  useEffect(() => setActiveIndex(-1), [query, sheet?.id]);

  useEffect(() => {
    setActiveIndex((current) =>
      current >= matches.length ? matches.length - 1 : current,
    );
  }, [matches.length]);

  const focusQuery = () => queryRef.current?.focus({ preventScroll: true });
  const moveToMatch = (direction: -1 | 1) => {
    if (!matches.length) return;
    const requestedIndex =
      activeIndex < 0
        ? direction > 0
          ? 0
          : matches.length - 1
        : activeIndex + direction;
    const index = (requestedIndex + matches.length * 2) % matches.length;
    const match = matches[index];
    if (!match) return;
    setActiveIndex(index);
    onSelectMatch(match);
  };

  const resultText =
    activeIndex >= 0 && matches.length
      ? `${activeIndex + 1}/${matches.length}`
      : matches.length
        ? `${matches.length} 个匹配`
        : query
          ? '没有匹配'
          : '';

  return (
    <search className="work-spreadsheet-find-bar" aria-label="查找当前工作表">
      <Search size={15} aria-hidden="true" />
      <OfficeTextField
        ref={queryRef}
        aria-label="查找当前工作表"
        placeholder="查找"
        value={query}
        onChange={(event) => setQuery(event.currentTarget.value)}
        onKeyDown={(event) => {
          if (event.key === 'Escape') {
            event.preventDefault();
            event.stopPropagation();
            onClose();
          } else if (event.key === 'Enter' && !event.nativeEvent.isComposing) {
            event.preventDefault();
            event.stopPropagation();
            moveToMatch(event.shiftKey ? -1 : 1);
          }
        }}
      />
      <output aria-live="polite">{resultText}</output>
      <span className="work-spreadsheet-find-actions">
        <IconButton
          label="上一个匹配"
          disabled={!matches.length}
          onClick={() => {
            moveToMatch(-1);
            focusQuery();
          }}
        >
          <ArrowUp size={14} />
        </IconButton>
        <IconButton
          label="下一个匹配"
          disabled={!matches.length}
          onClick={() => {
            moveToMatch(1);
            focusQuery();
          }}
        >
          <ArrowDown size={14} />
        </IconButton>
        <IconButton label="关闭查找" onClick={onClose}>
          <X size={14} />
        </IconButton>
      </span>
    </search>
  );
}
