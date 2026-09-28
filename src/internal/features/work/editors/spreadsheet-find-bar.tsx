import { ArrowDown, ArrowUp, Search, X } from 'lucide-react';
import { useEffect, useMemo, useRef, useState } from 'react';
import { IconButton } from '../../../design-system/primitives';
import { officeMessage } from '../../../i18n/office-locale';
import type { WorkSpreadsheetContent } from '../work-types';
import { OfficeTextField } from './office-controls';
import { useOfficeMessages } from './office-messages-context';
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
  const messages = useOfficeMessages();
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
        ? officeMessage(messages, 'spreadsheet.findBar.matchCount', {
            count: String(matches.length),
          })
        : query
          ? officeMessage(messages, 'spreadsheet.findBar.noMatch')
          : '';

  return (
    <search
      className="work-spreadsheet-find-bar"
      aria-label={officeMessage(messages, 'spreadsheet.findBar.aria')}
    >
      <Search size={15} aria-hidden="true" />
      <OfficeTextField
        ref={queryRef}
        aria-label={officeMessage(messages, 'spreadsheet.findBar.aria')}
        placeholder={officeMessage(messages, 'spreadsheet.findBar.placeholder')}
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
          label={officeMessage(messages, 'spreadsheet.findBar.previous')}
          disabled={!matches.length}
          onClick={() => {
            moveToMatch(-1);
            focusQuery();
          }}
        >
          <ArrowUp size={14} />
        </IconButton>
        <IconButton
          label={officeMessage(messages, 'spreadsheet.findBar.next')}
          disabled={!matches.length}
          onClick={() => {
            moveToMatch(1);
            focusQuery();
          }}
        >
          <ArrowDown size={14} />
        </IconButton>
        <IconButton
          label={officeMessage(messages, 'spreadsheet.findBar.close')}
          onClick={onClose}
        >
          <X size={14} />
        </IconButton>
      </span>
    </search>
  );
}
