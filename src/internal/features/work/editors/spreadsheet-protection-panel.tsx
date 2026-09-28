import type { Sheet } from '@fortune-sheet/core';
import { KeyRound, Plus, Trash2 } from 'lucide-react';
import { useEffect, useState } from 'react';
import { officeMessage } from '../../../i18n/office-locale';
import { sparseMatrixColumnCount } from '../spreadsheet-sparse';
import {
  Button,
  CollectionState,
  InlineNotice,
  StateView,
} from '../../../design-system/primitives';
import {
  editableRangeCellCount,
  editableRangeRequiresCredentials,
  sheetProtectionAuthority,
  unlockedCellCount,
  withEditableRange,
  withoutEditableRange,
  withSheetProtection,
  withSheetSelectionPermissions,
} from '../work-spreadsheet-protection';
import {
  formatSpreadsheetCellRanges,
  parseSpreadsheetCellRanges,
} from '../work-spreadsheet-ranges';
import type { WorkSpreadsheetContent } from '../work-types';
import {
  OfficeCheckbox,
  OfficeSelect,
  OfficeTextField,
} from './office-controls';
import { useOfficeMessages } from './office-messages-context';
import { useOfficeDraft } from './use-office-draft';

interface SpreadsheetProtectionPanelProps {
  content: WorkSpreadsheetContent;
  onChange: (content: WorkSpreadsheetContent) => void;
}

interface EditableRangeDraft {
  name: string;
  reference: string;
}

const MAX_EDITABLE_RANGE_CELLS = 100_000;

export function SpreadsheetProtectionPanel({
  content,
  onChange,
}: SpreadsheetProtectionPanelProps) {
  const messages = useOfficeMessages();
  const sheets = content.sheets.filter(
    (sheet): sheet is Sheet & { id: string } => Boolean(sheet.id),
  );
  const initialSheetId =
    sheets.find((sheet) => sheet.status === 1)?.id ?? sheets[0]?.id ?? '';
  const [sheetId, setSheetId] = useState(initialSheetId);
  const [selectedIndex, setSelectedIndex] = useState<number | null>(null);
  const {
    cancelDraft: resetDraft,
    dirty,
    draft,
    replaceDraft,
    setDraft,
    syncDraft,
  } = useOfficeDraft<EditableRangeDraft>(newRangeDraft);
  const [error, setError] = useState('');
  const sheet = sheets.find((item) => item.id === sheetId) ?? sheets[0];
  const authority = sheet ? sheetProtectionAuthority(sheet) : null;
  const ranges = authority?.allowRangeList ?? [];

  useEffect(() => {
    if (sheet && sheet.id !== sheetId) setSheetId(sheet.id);
  }, [sheet?.id, sheetId]);

  useEffect(() => {
    if (selectedIndex === null) return;
    const range = ranges[selectedIndex];
    if (range) {
      syncDraft({ name: range.name, reference: range.sqref });
      return;
    }
    setSelectedIndex(null);
    replaceDraft(newRangeDraft());
  }, [content.sheets, replaceDraft, sheetId, syncDraft]);

  if (!sheet || !authority) {
    return (
      <StateView
        className="work-office-panel-empty work-spreadsheet-protection-empty"
        size="compact"
        title={officeMessage(messages, 'spreadsheet.protection.emptyTitle')}
      />
    );
  }

  const updateSheet = (nextSheet: Sheet) => {
    onChange({
      ...content,
      sheets: content.sheets.map((item) =>
        item.id === nextSheet.id ? nextSheet : item,
      ),
    });
  };
  const changeSheet = (nextSheetId: string) => {
    if (nextSheetId === sheetId) return;
    if (dirty) {
      setError(officeMessage(messages, 'spreadsheet.protection.error.unsaved'));
      return;
    }
    setSheetId(nextSheetId);
    setSelectedIndex(null);
    replaceDraft(newRangeDraft());
    setError('');
  };
  const selectRange = (index: number) => {
    if (index === selectedIndex) return;
    if (dirty) {
      setError(officeMessage(messages, 'spreadsheet.protection.error.unsaved'));
      return;
    }
    const range = ranges[index];
    if (!range) return;
    setSelectedIndex(index);
    replaceDraft({ name: range.name, reference: range.sqref });
    setError('');
  };
  const startNew = () => {
    if (dirty) {
      setError(officeMessage(messages, 'spreadsheet.protection.error.unsaved'));
      return;
    }
    setSelectedIndex(null);
    replaceDraft(newRangeDraft());
    setError('');
  };
  const saveRange = () => {
    if (selectedIndex !== null && !dirty) return;
    const name = draft.name.trim();
    const parsed = parseSpreadsheetCellRanges(draft.reference);
    if (!name) {
      setError(officeMessage(messages, 'spreadsheet.protection.error.nameRequired'));
      return;
    }
    if (!parsed) {
      setError(officeMessage(messages, 'spreadsheet.protection.error.invalidRange'));
      return;
    }
    if (
      ranges.some(
        (range, index) =>
          index !== selectedIndex &&
          range.name.trim().toLowerCase() === name.toLowerCase(),
      )
    ) {
      setError(officeMessage(messages, 'spreadsheet.protection.error.duplicateName'));
      return;
    }
    if (editableRangeCellCount(parsed) > MAX_EDITABLE_RANGE_CELLS) {
      setError(officeMessage(messages, 'spreadsheet.protection.error.tooManyCells'));
      return;
    }
    const maximumRow = Math.max(1, sheet.row ?? sheet.data?.length ?? 1) - 1;
    const maximumColumn =
      Math.max(1, sheet.column ?? sparseMatrixColumnCount(sheet.data)) - 1;
    if (
      parsed.some(
        (range) => range.row[1] > maximumRow || range.column[1] > maximumColumn,
      )
    ) {
      setError(
        officeMessage(messages, 'spreadsheet.protection.error.outOfBounds', {
          rows: String(maximumRow + 1),
          cols: String(maximumColumn + 1),
        }),
      );
      return;
    }
    const index = selectedIndex ?? ranges.length;
    updateSheet(
      withEditableRange(sheet, selectedIndex, {
        name,
        sqref: formatSpreadsheetCellRanges(parsed),
      }),
    );
    setSelectedIndex(index);
    replaceDraft({ name, reference: formatSpreadsheetCellRanges(parsed) });
    setError('');
  };
  const cancelDraft = () => {
    resetDraft();
    setError('');
  };
  const deleteRange = () => {
    if (selectedIndex === null) return;
    updateSheet(withoutEditableRange(sheet, selectedIndex));
    setSelectedIndex(null);
    replaceDraft(newRangeDraft());
    setError('');
  };
  const setSelectLocked = (checked: boolean) => {
    updateSheet(
      withSheetSelectionPermissions(sheet, {
        selectLockedCells: checked,
        selectUnlockedCells: checked
          ? true
          : authority.selectunLockedCells === 1,
      }),
    );
  };
  const setSelectUnlocked = (checked: boolean) => {
    updateSheet(
      withSheetSelectionPermissions(sheet, {
        selectLockedCells: checked ? authority.selectLockedCells === 1 : false,
        selectUnlockedCells: checked,
      }),
    );
  };
  const selectedRange = selectedIndex === null ? null : ranges[selectedIndex];

  return (
    <fieldset
      className="work-spreadsheet-protection-manager"
      data-office-escape-consumer={dirty || undefined}
      onKeyDown={(event) => {
        if (event.key !== 'Escape' || event.defaultPrevented || !dirty) return;
        event.preventDefault();
        event.stopPropagation();
        cancelDraft();
      }}
    >
      <legend className="sr-only">{officeMessage(messages, 'spreadsheet.protection.legend')}</legend>
      <aside aria-label={officeMessage(messages, 'spreadsheet.protection.rangesAria')}>
        <Button className="create" tone="secondary" onClick={startNew}>
          <Plus size={13} />
          {officeMessage(messages, 'spreadsheet.protection.newRange')}
        </Button>
        <div className="work-spreadsheet-protection-list">
          {ranges.map((range, index) => (
            <button
              type="button"
              className={selectedIndex === index ? 'active' : ''}
              key={`${range.name}-${index}`}
              onClick={() => selectRange(index)}
            >
              <strong>{range.name}</strong>
              {editableRangeRequiresCredentials(range) && (
                <span className="credential">
                  <KeyRound size={10} />
                  {officeMessage(messages, 'spreadsheet.protection.sourceCredential')}
                </span>
              )}
              <small>{range.sqref}</small>
            </button>
          ))}
          {!ranges.length && (
            <CollectionState
              className="work-office-collection-empty"
              role="status"
            >
              {officeMessage(messages, 'spreadsheet.protection.noRanges')}
            </CollectionState>
          )}
        </div>
      </aside>
      <form
        onSubmit={(event) => {
          event.preventDefault();
          saveRange();
        }}
      >
        <div className="work-office-field">
          <span>{officeMessage(messages, 'spreadsheet.protection.sheet')}</span>
          <OfficeSelect
            ariaLabel={officeMessage(messages, 'spreadsheet.protection.sheetAria')}
            value={sheet.id}
            options={sheets.map((item) => ({
              value: item.id,
              label: item.name,
            }))}
            onValueChange={changeSheet}
          />
        </div>
        <OfficeCheckbox
          className="toggle"
          ariaLabel={officeMessage(messages, 'spreadsheet.protection.enableAria')}
          checked={authority.sheet === 1}
          onCheckedChange={(checked) =>
            updateSheet(withSheetProtection(sheet, checked))
          }
        >
          {officeMessage(messages, 'spreadsheet.protection.enable')}
        </OfficeCheckbox>
        <fieldset>
          <legend>{officeMessage(messages, 'spreadsheet.protection.allowSelect')}</legend>
          <OfficeCheckbox
            className="check"
            ariaLabel={officeMessage(messages, 'spreadsheet.protection.selectLockedAria')}
            checked={authority.selectLockedCells === 1}
            onCheckedChange={setSelectLocked}
          >
            {officeMessage(messages, 'spreadsheet.protection.selectLocked')}
          </OfficeCheckbox>
          <OfficeCheckbox
            className="check"
            ariaLabel={officeMessage(messages, 'spreadsheet.protection.selectUnlockedAria')}
            checked={authority.selectunLockedCells === 1}
            onCheckedChange={setSelectUnlocked}
          >
            {officeMessage(messages, 'spreadsheet.protection.selectUnlocked')}
          </OfficeCheckbox>
        </fieldset>
        <div className="work-office-field">
          <span>{officeMessage(messages, 'spreadsheet.protection.rangeName')}</span>
          <OfficeTextField
            aria-label={officeMessage(messages, 'spreadsheet.protection.rangeNameAria')}
            value={draft.name}
            maxLength={255}
            placeholder={officeMessage(messages, 'spreadsheet.protection.rangeNamePlaceholder')}
            onChange={(event) =>
              setDraft({ ...draft, name: event.target.value })
            }
          />
        </div>
        <div className="work-office-field reference">
          <span>{officeMessage(messages, 'spreadsheet.protection.rangeRef')}</span>
          <OfficeTextField
            aria-label={officeMessage(messages, 'spreadsheet.protection.rangeRefAria')}
            value={draft.reference}
            placeholder="B2:B20"
            onChange={(event) =>
              setDraft({ ...draft, reference: event.target.value })
            }
          />
        </div>
        <p>
          {officeMessage(messages, 'spreadsheet.protection.unlockedSummary', {
            n: String(unlockedCellCount(sheet)),
          })}
        </p>
        {selectedRange && editableRangeRequiresCredentials(selectedRange) && (
          <InlineNotice
            className="work-office-form-warning"
            tone="warning"
            role="note"
          >
            {officeMessage(messages, 'spreadsheet.protection.credentialConvertNote')}
          </InlineNotice>
        )}
        <div className="actions">
          {error && (
            <InlineNotice
              className="work-office-form-error"
              tone="danger"
              role="alert"
            >
              {error}
            </InlineNotice>
          )}
          <Button
            tone="danger"
            disabled={selectedIndex === null}
            onClick={deleteRange}
          >
            <Trash2 size={13} />
            {officeMessage(messages, 'spreadsheet.protection.deleteRange')}
          </Button>
          <Button tone="secondary" disabled={!dirty} onClick={cancelDraft}>
            {officeMessage(messages, 'spreadsheet.protection.cancelChanges')}
          </Button>
          <Button
            type="submit"
            tone="primary"
            disabled={selectedIndex !== null && !dirty}
          >
            {officeMessage(messages, 'spreadsheet.protection.saveRange')}
          </Button>
        </div>
      </form>
    </fieldset>
  );
}

function newRangeDraft(): EditableRangeDraft {
  return { name: '', reference: 'B2:B10' };
}
