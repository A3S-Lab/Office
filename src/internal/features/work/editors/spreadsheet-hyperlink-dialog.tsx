import { Globe2, Grid2X2, Link2, Table2 } from 'lucide-react';
import { type FormEvent, useId, useMemo, useState } from 'react';
import { Button, Dialog, Field } from '../../../design-system/primitives';
import { officeMessage } from '../../../i18n/office-locale';
import { OfficeSelect } from './office-controls';
import { useOfficeMessages } from './office-messages-context';
import type {
  SpreadsheetHyperlinkDialogSource,
  SpreadsheetHyperlinkDialogValue,
  SpreadsheetHyperlinkType,
} from './spreadsheet-hyperlink';

export function SpreadsheetHyperlinkDialog({
  source,
  restoreFocusTarget,
  onApply,
  onClose,
  onRemove,
  onValidate,
}: {
  source: SpreadsheetHyperlinkDialogSource;
  restoreFocusTarget: () => HTMLElement | null;
  onApply: (value: SpreadsheetHyperlinkDialogValue) => boolean;
  onClose: () => void;
  onRemove: () => boolean;
  onValidate: (value: SpreadsheetHyperlinkDialogValue) => string | null;
}) {
  const messages = useOfficeMessages();
  const hyperlinkTypes = [
    {
      id: 'webpage' as const,
      label: officeMessage(messages, 'spreadsheet.hyperlink.type.webpage'),
      description: officeMessage(
        messages,
        'spreadsheet.hyperlink.type.webpageDesc',
      ),
      icon: Globe2,
    },
    {
      id: 'cellrange' as const,
      label: officeMessage(messages, 'spreadsheet.hyperlink.type.range'),
      description: officeMessage(
        messages,
        'spreadsheet.hyperlink.type.rangeDesc',
      ),
      icon: Grid2X2,
    },
    {
      id: 'sheet' as const,
      label: officeMessage(messages, 'spreadsheet.hyperlink.type.sheet'),
      description: officeMessage(
        messages,
        'spreadsheet.hyperlink.type.sheetDesc',
      ),
      icon: Table2,
    },
  ];
  const [linkType, setLinkType] = useState<SpreadsheetHyperlinkType>(
    source.link?.linkType ?? 'webpage',
  );
  const [linkAddress, setLinkAddress] = useState(
    source.link?.linkAddress ?? '',
  );
  const [displayText, setDisplayText] = useState(source.displayText);
  const [addressTouched, setAddressTouched] = useState(false);
  const formId = useId();
  const typeId = useId().replaceAll(':', '');
  const value = useMemo(
    () =>
      spreadsheetHyperlinkDialogValue(
        source,
        linkType,
        linkAddress,
        displayText,
      ),
    [displayText, linkAddress, linkType, source],
  );
  const validationError = onValidate(value);
  const addressMissing = !linkAddress.trim();
  const visibleError =
    addressTouched && !addressMissing ? validationError : undefined;

  const submit = (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    setAddressTouched(true);
    if (addressMissing || validationError) return;
    if (onApply(value)) onClose();
  };

  const changeType = (nextType: SpreadsheetHyperlinkType) => {
    setLinkType(nextType);
    setLinkAddress(defaultSpreadsheetHyperlinkAddress(source, nextType));
    setAddressTouched(false);
  };

  return (
    <Dialog
      title={officeMessage(
        messages,
        source.hasHyperlink
          ? 'spreadsheet.hyperlink.editTitle'
          : 'spreadsheet.hyperlink.insertTitle',
      )}
      description={`${source.sheetName}!${source.cellReference}`}
      className="work-spreadsheet-hyperlink-dialog"
      restoreFocusTarget={restoreFocusTarget}
      onClose={onClose}
      footer={
        <>
          {source.hasHyperlink ? (
            <Button
              tone="danger"
              className="work-spreadsheet-hyperlink-remove"
              onClick={() => {
                if (onRemove()) onClose();
              }}
            >
              {officeMessage(messages, 'spreadsheet.hyperlink.remove')}
            </Button>
          ) : null}
          <Button tone="quiet" onClick={onClose}>
            {officeMessage(messages, 'spreadsheet.hyperlink.cancel')}
          </Button>
          <Button
            tone="primary"
            type="submit"
            form={formId}
            disabled={addressMissing || Boolean(validationError)}
          >
            {officeMessage(messages, 'spreadsheet.hyperlink.confirm')}
          </Button>
        </>
      }
    >
      <form id={formId} onSubmit={submit}>
        <fieldset
          className="work-spreadsheet-hyperlink-types"
          aria-label={officeMessage(messages, 'spreadsheet.hyperlink.linkType')}
        >
          <legend>
            {officeMessage(messages, 'spreadsheet.hyperlink.linkType')}
          </legend>
          <div
            role="radiogroup"
            aria-label={officeMessage(
              messages,
              'spreadsheet.hyperlink.linkType',
            )}
          >
            {hyperlinkTypes.map((item) => {
              const Icon = item.icon;
              const checked = linkType === item.id;
              const inputId = `${typeId}-${item.id}`;
              return (
                <label key={item.id} className={checked ? 'selected' : ''}>
                  <input
                    id={inputId}
                    type="radio"
                    aria-label={item.label}
                    name={`${typeId}-link-type`}
                    value={item.id}
                    checked={checked}
                    onChange={() => changeType(item.id)}
                  />
                  <Icon size={17} aria-hidden="true" />
                  <span>
                    <strong>{item.label}</strong>
                    <small>{item.description}</small>
                  </span>
                </label>
              );
            })}
          </div>
        </fieldset>

        <Field
          label={officeMessage(messages, 'spreadsheet.hyperlink.displayText')}
          description={officeMessage(
            messages,
            source.displayTextEditable
              ? 'spreadsheet.hyperlink.displayHintEmpty'
              : 'spreadsheet.hyperlink.displayHintFormula',
          )}
        >
          <input
            type="text"
            value={displayText}
            disabled={!source.displayTextEditable}
            maxLength={32_767}
            onChange={(event) => setDisplayText(event.currentTarget.value)}
          />
        </Field>

        {linkType === 'sheet' ? (
          <Field
            label={officeMessage(messages, 'spreadsheet.hyperlink.sheet')}
            required
            error={visibleError}
          >
            <OfficeSelect
              ariaLabel={officeMessage(
                messages,
                'spreadsheet.hyperlink.sheetAria',
              )}
              value={linkAddress}
              options={source.sheetOptions.map((sheet) => ({
                value: sheet.name,
                label: sheet.name,
              }))}
              onValueChange={(next) => {
                setLinkAddress(next);
                setAddressTouched(true);
              }}
            />
          </Field>
        ) : (
          <Field
            label={officeMessage(
              messages,
              linkType === 'webpage'
                ? 'spreadsheet.hyperlink.address'
                : 'spreadsheet.hyperlink.cellOrRange',
            )}
            required
            error={visibleError}
            description={officeMessage(
              messages,
              linkType === 'webpage'
                ? 'spreadsheet.hyperlink.webPlaceholder'
                : 'spreadsheet.hyperlink.rangePlaceholder',
            )}
          >
            <input
              type="text"
              inputMode={linkType === 'webpage' ? 'url' : 'text'}
              autoCapitalize="none"
              spellCheck={false}
              value={linkAddress}
              onBlur={() => setAddressTouched(true)}
              onChange={(event) => {
                setLinkAddress(event.currentTarget.value);
                setAddressTouched(true);
              }}
            />
          </Field>
        )}

        <div className="work-spreadsheet-hyperlink-preview" aria-live="polite">
          <span aria-hidden="true">
            <Link2 size={17} />
          </span>
          <div>
            <strong>
              {(value.displayText ?? source.displayText) ||
                officeMessage(messages, 'spreadsheet.hyperlink.defaultDisplay')}
            </strong>
            <small>
              {linkAddress.trim() ||
                officeMessage(messages, 'spreadsheet.hyperlink.awaitingTarget')}
            </small>
          </div>
        </div>
      </form>
    </Dialog>
  );
}

function spreadsheetHyperlinkDialogValue(
  source: SpreadsheetHyperlinkDialogSource,
  linkType: SpreadsheetHyperlinkType,
  linkAddress: string,
  displayText: string,
): SpreadsheetHyperlinkDialogValue {
  const normalizedDisplay = displayText.trim() || linkAddress.trim();
  const shouldChangeDisplay =
    source.displayTextEditable && normalizedDisplay !== source.displayText;
  return {
    linkType,
    linkAddress,
    ...(shouldChangeDisplay ? { displayText: normalizedDisplay } : {}),
  };
}

function defaultSpreadsheetHyperlinkAddress(
  source: SpreadsheetHyperlinkDialogSource,
  linkType: SpreadsheetHyperlinkType,
): string {
  if (source.link?.linkType === linkType) return source.link.linkAddress;
  if (linkType !== 'sheet') return '';
  return (
    source.sheetOptions.find((sheet) => sheet.id !== source.sheetId)?.name ??
    source.sheetOptions[0]?.name ??
    ''
  );
}
