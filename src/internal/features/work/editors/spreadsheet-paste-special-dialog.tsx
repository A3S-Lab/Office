import { type FormEvent, useId, useState } from 'react';
import { Button, Dialog } from '../../../design-system/primitives';
import { officeMessage } from '../../../i18n/office-locale';
import { OfficeCheckbox } from './office-controls';
import { useOfficeMessages } from './office-messages-context';
import {
  spreadsheetPasteContentOptions,
  spreadsheetPasteOperationOptions,
  spreadsheetPasteSpecialModeAvailable,
  type SpreadsheetPasteContent,
  type SpreadsheetPasteSpecialOptions,
} from './spreadsheet-paste-special';
import type { SpreadsheetPasteSpecialDialogSource } from './use-spreadsheet-clipboard';

export function SpreadsheetPasteSpecialDialog({
  source,
  restoreFocusTarget,
  onApply,
  onClose,
  onValidate,
}: {
  source: SpreadsheetPasteSpecialDialogSource;
  restoreFocusTarget: () => HTMLElement | null;
  onApply: (options: SpreadsheetPasteSpecialOptions) => boolean;
  onClose: () => void;
  onValidate: (options: SpreadsheetPasteSpecialOptions) => string | null;
}) {
  const messages = useOfficeMessages();
  const [options, setOptions] = useState<SpreadsheetPasteSpecialOptions>({
    content: 'all',
    operation: 'none',
    skipBlanks: false,
    transpose: false,
  });
  const formId = useId();
  const validationError = onValidate(options);
  const contentOptions = spreadsheetPasteContentOptions(messages);
  const operationOptions = spreadsheetPasteOperationOptions(messages);
  const sourceKind =
    source.snapshot.kind === 'rich'
      ? officeMessage(messages, 'spreadsheet.pasteSpecial.richClipboard')
      : officeMessage(messages, 'spreadsheet.pasteSpecial.plainText');
  const sourceSize = officeMessage(
    messages,
    'spreadsheet.pasteSpecial.sourceSize',
    {
      rows: String(source.snapshot.rowCount),
      columns: String(source.snapshot.columnCount),
    },
  );

  const selectContent = (content: SpreadsheetPasteContent) => {
    setOptions((current) => ({
      ...current,
      content,
      operation: spreadsheetPasteContentSupportsOperation(content)
        ? current.operation
        : 'none',
      skipBlanks: content === 'column-widths' ? false : current.skipBlanks,
      transpose: content === 'column-widths' ? false : current.transpose,
    }));
  };
  const submit = (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    if (!validationError && onApply(options)) onClose();
  };

  return (
    <Dialog
      title={officeMessage(messages, 'spreadsheet.pasteSpecial.title')}
      description={officeMessage(
        messages,
        'spreadsheet.pasteSpecial.description',
      )}
      className="work-spreadsheet-paste-special-dialog"
      restoreFocusTarget={restoreFocusTarget}
      onClose={onClose}
      footer={
        <>
          <Button tone="quiet" onClick={onClose}>
            {officeMessage(messages, 'spreadsheet.pasteSpecial.cancel')}
          </Button>
          <Button
            tone="primary"
            type="submit"
            form={formId}
            disabled={Boolean(validationError)}
          >
            {officeMessage(messages, 'spreadsheet.pasteSpecial.paste')}
          </Button>
        </>
      }
    >
      <section
        className="work-spreadsheet-paste-special-source"
        aria-label={officeMessage(
          messages,
          'spreadsheet.pasteSpecial.summaryAria',
        )}
      >
        <span>{sourceKind}</span>
        <strong>{sourceSize}</strong>
      </section>
      <form id={formId} onSubmit={submit}>
        <fieldset className="work-spreadsheet-paste-special-modes">
          <legend>
            {officeMessage(messages, 'spreadsheet.pasteSpecial.contentLegend')}
          </legend>
          <div>
            {contentOptions.map((option) => {
              const available = spreadsheetPasteSpecialModeAvailable(
                source.snapshot,
                option.value,
              );
              return (
                <label
                  key={option.value}
                  className={options.content === option.value ? 'selected' : ''}
                >
                  <input
                    type="radio"
                    name="spreadsheet-paste-content"
                    value={option.value}
                    checked={options.content === option.value}
                    disabled={!available}
                    onChange={() => selectContent(option.value)}
                  />
                  <span aria-hidden="true" />
                  <span>{option.label}</span>
                  {!available && (
                    <small>
                      {officeMessage(
                        messages,
                        'spreadsheet.pasteSpecial.richOnly',
                      )}
                    </small>
                  )}
                </label>
              );
            })}
          </div>
        </fieldset>

        <fieldset className="work-spreadsheet-paste-special-operations">
          <legend>
            {officeMessage(
              messages,
              'spreadsheet.pasteSpecial.operationLegend',
            )}
          </legend>
          <div>
            {operationOptions.map((operation) => (
              <label
                key={operation.value}
                className={
                  options.operation === operation.value ? 'selected' : ''
                }
              >
                <input
                  type="radio"
                  name="spreadsheet-paste-operation"
                  value={operation.value}
                  checked={options.operation === operation.value}
                  disabled={
                    operation.value !== 'none' &&
                    !spreadsheetPasteContentSupportsOperation(options.content)
                  }
                  onChange={() =>
                    setOptions((current) => ({
                      ...current,
                      operation: operation.value,
                    }))
                  }
                />
                <span>{operation.label}</span>
              </label>
            ))}
          </div>
        </fieldset>

        <div className="work-spreadsheet-paste-special-options">
          <OfficeCheckbox
            ariaLabel={officeMessage(
              messages,
              'spreadsheet.pasteSpecial.skipBlanksAria',
            )}
            checked={options.skipBlanks}
            disabled={options.content === 'column-widths'}
            onCheckedChange={(skipBlanks) =>
              setOptions((current) => ({ ...current, skipBlanks }))
            }
          >
            {officeMessage(messages, 'spreadsheet.pasteSpecial.skipBlanks')}
          </OfficeCheckbox>
          <OfficeCheckbox
            ariaLabel={officeMessage(
              messages,
              'spreadsheet.pasteSpecial.transposeAria',
            )}
            checked={options.transpose}
            disabled={options.content === 'column-widths'}
            onCheckedChange={(transpose) =>
              setOptions((current) => ({ ...current, transpose }))
            }
          >
            {officeMessage(messages, 'spreadsheet.pasteSpecial.transpose')}
          </OfficeCheckbox>
        </div>

        {validationError && (
          <p className="work-spreadsheet-paste-special-error" role="alert">
            {validationError}
          </p>
        )}
      </form>
    </Dialog>
  );
}

function spreadsheetPasteContentSupportsOperation(
  content: SpreadsheetPasteContent,
): boolean {
  return !['formats', 'comments', 'validation', 'column-widths'].includes(
    content,
  );
}
