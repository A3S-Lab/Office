import { Table2 } from 'lucide-react';
import { type FormEvent, useId, useState } from 'react';
import { Button, Dialog, Field } from '../../../design-system/primitives';
import { officeMessage } from '../../../i18n/office-locale';
import { OfficeCheckbox } from './office-controls';
import { useOfficeMessages } from './office-messages-context';
import type {
  SpreadsheetTableDialogSource,
  SpreadsheetTableDialogValue,
} from './spreadsheet-table';

export function SpreadsheetTableDialog({
  source,
  restoreFocusTarget,
  onApply,
  onClose,
  onValidate,
}: {
  source: SpreadsheetTableDialogSource;
  restoreFocusTarget: () => HTMLElement | null;
  onApply: (value: SpreadsheetTableDialogValue) => boolean;
  onClose: () => void;
  onValidate: (value: SpreadsheetTableDialogValue) => string | null;
}) {
  const messages = useOfficeMessages();
  const [value, setValue] = useState(source.value);
  const [touched, setTouched] = useState(false);
  const formId = useId();
  const validationError = onValidate(value);
  const visibleError = touched ? validationError : null;
  const dirty =
    value.rangeReference !== source.value.rangeReference ||
    value.headerRow !== source.value.headerRow ||
    Boolean(value.totalsRow) !== Boolean(source.value.totalsRow);

  const resetDraft = () => {
    setValue(source.value);
    setTouched(false);
  };

  const submit = (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    setTouched(true);
    if (validationError) return;
    if (onApply(value)) onClose();
  };

  return (
    <Dialog
      title={officeMessage(messages, 'spreadsheet.tableDialog.title')}
      description={source.sheetName}
      className="work-spreadsheet-table-dialog"
      restoreFocusTarget={restoreFocusTarget}
      onClose={onClose}
      onEscape={() => {
        if (dirty) {
          resetDraft();
          return;
        }
        onClose();
      }}
      footer={
        <>
          <Button tone="quiet" onClick={onClose}>
            {officeMessage(messages, 'spreadsheet.tableDialog.cancel')}
          </Button>
          <Button
            tone="primary"
            type="submit"
            form={formId}
            disabled={Boolean(validationError)}
          >
            {officeMessage(messages, 'spreadsheet.tableDialog.confirm')}
          </Button>
        </>
      }
    >
      <form id={formId} onSubmit={submit}>
        <div className="work-spreadsheet-table-scope">
          <span aria-hidden="true">
            <Table2 size={18} />
          </span>
          <div>
            <strong>{source.name}</strong>
            <small>
              {officeMessage(messages, 'spreadsheet.tableDialog.hint')}
            </small>
          </div>
        </div>
        <Field
          label={officeMessage(messages, 'spreadsheet.tableDialog.range')}
          required
          error={visibleError ?? undefined}
          description={officeMessage(
            messages,
            'spreadsheet.tableDialog.rangeDesc',
          )}
        >
          <input
            type="text"
            autoCapitalize="none"
            autoFocus
            spellCheck={false}
            value={value.rangeReference}
            onBlur={() => setTouched(true)}
            onChange={(event) => {
              const rangeReference = event.currentTarget.value;
              setValue((current) => ({
                ...current,
                rangeReference,
              }));
              setTouched(true);
            }}
          />
        </Field>
        <OfficeCheckbox
          ariaLabel={officeMessage(
            messages,
            'spreadsheet.tableDialog.headerRow',
          )}
          checked={value.headerRow}
          onCheckedChange={(headerRow) =>
            setValue((current) => ({ ...current, headerRow }))
          }
        >
          {officeMessage(messages, 'spreadsheet.tableDialog.headerRow')}
        </OfficeCheckbox>
        <OfficeCheckbox
          ariaLabel={officeMessage(
            messages,
            'spreadsheet.tableDialog.totalsRow',
          )}
          checked={value.totalsRow === true}
          onCheckedChange={(totalsRow) =>
            setValue((current) => ({ ...current, totalsRow }))
          }
        >
          {officeMessage(messages, 'spreadsheet.tableDialog.totalsRow')}
        </OfficeCheckbox>
      </form>
    </Dialog>
  );
}
