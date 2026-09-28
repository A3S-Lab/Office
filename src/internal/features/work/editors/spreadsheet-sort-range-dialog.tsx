import { type FormEvent, useId, useState } from 'react';
import { officeMessage } from '../../../i18n/office-locale';
import { Button, Dialog } from '../../../design-system/primitives';
import { useOfficeMessages } from './office-messages-context';
import type {
  SpreadsheetSortRangeChoice,
  SpreadsheetSortRangeDialogSource,
} from './spreadsheet-sort';

export function SpreadsheetSortRangeDialog({
  source,
  restoreFocusTarget,
  onApply,
  onClose,
}: {
  source: SpreadsheetSortRangeDialogSource;
  restoreFocusTarget: () => HTMLElement | null;
  onApply: (choice: SpreadsheetSortRangeChoice) => boolean;
  onClose: () => void;
}) {
  const messages = useOfficeMessages();
  const [choice, setChoice] = useState<SpreadsheetSortRangeChoice>(() =>
    source.canSortExpandedRange ? 'expand' : 'selection',
  );
  const formId = useId();
  const ownedScope = source.ownedScope;
  const description = ownedScope
    ? officeMessage(messages, 'spreadsheet.sort.range.descOwned', {
        sheet: source.sheetName,
        scope:
          ownedScope.kind === 'table'
            ? officeMessage(messages, 'spreadsheet.sort.range.scopeTable')
            : officeMessage(messages, 'spreadsheet.sort.range.scopeFilter'),
      })
    : officeMessage(messages, 'spreadsheet.sort.range.descAdjacent', {
        sheet: source.sheetName,
      });
  const expandedTitle =
    ownedScope?.kind === 'table'
      ? officeMessage(messages, 'spreadsheet.sort.range.expandTable')
      : ownedScope?.kind === 'auto-filter'
        ? officeMessage(messages, 'spreadsheet.sort.range.expandFilter')
        : officeMessage(messages, 'spreadsheet.sort.range.expandAdjacent');
  const expandedHint =
    ownedScope?.kind === 'table'
      ? officeMessage(messages, 'spreadsheet.sort.range.hintTable')
      : ownedScope?.kind === 'auto-filter'
        ? officeMessage(messages, 'spreadsheet.sort.range.hintFilter')
        : officeMessage(messages, 'spreadsheet.sort.range.hintAdjacent');
  const submit = (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    if (onApply(choice)) onClose();
  };

  return (
    <Dialog
      title={officeMessage(messages, 'spreadsheet.sort.range.title')}
      description={description}
      className="work-spreadsheet-sort-range-dialog"
      restoreFocusTarget={restoreFocusTarget}
      onClose={onClose}
      footer={
        <>
          <Button tone="quiet" onClick={onClose}>
            {officeMessage(messages, 'spreadsheet.sort.cancel')}
          </Button>
          <Button tone="primary" type="submit" form={formId}>
            {officeMessage(messages, 'spreadsheet.sort.range.sort')}
          </Button>
        </>
      }
    >
      <form id={formId} onSubmit={submit}>
        <fieldset className="work-spreadsheet-sort-range-options">
          <legend>{officeMessage(messages, 'spreadsheet.sort.range.legend')}</legend>
          <label>
            <input
              type="radio"
              name="spreadsheet-sort-range"
              value="expand"
              checked={choice === 'expand'}
              disabled={!source.canSortExpandedRange}
              onChange={() => setChoice('expand')}
            />
            <span>
              <strong>{expandedTitle}</strong>
              <small>{expandedHint}</small>
              <code>{source.expandedRangeReference}</code>
            </span>
          </label>
          <label>
            <input
              type="radio"
              name="spreadsheet-sort-range"
              value="selection"
              checked={choice === 'selection'}
              disabled={!source.canSortSelection}
              onChange={() => setChoice('selection')}
            />
            <span>
              <strong>
                {officeMessage(messages, 'spreadsheet.sort.range.selection')}
              </strong>
              <small>
                {officeMessage(messages, 'spreadsheet.sort.range.selectionHint')}
              </small>
              <code>{source.selectedRangeReference}</code>
            </span>
          </label>
        </fieldset>
      </form>
    </Dialog>
  );
}
