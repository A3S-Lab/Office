import { useId } from 'react';
import { Button, Dialog } from '../../../design-system/primitives';
import { officeMessage } from '../../../i18n/office-locale';
import type { WorkDocumentIndexEntryDraft } from '../work-document-index';
import { OfficeCheckbox, OfficeTextField } from './office-controls';
import { useOfficeMessages } from './office-messages-context';

export interface DocumentIndexEntryDialogProps {
  editing: boolean;
  value: WorkDocumentIndexEntryDraft;
  restoreFocusTarget: () => HTMLElement | null;
  onCancel: () => void;
  onChange: (value: WorkDocumentIndexEntryDraft) => void;
  onSubmit: () => void;
}

export function DocumentIndexEntryDialog({
  editing,
  value,
  restoreFocusTarget,
  onCancel,
  onChange,
  onSubmit,
}: DocumentIndexEntryDialogProps) {
  const messages = useOfficeMessages();
  const update = (patch: Partial<WorkDocumentIndexEntryDraft>) =>
    onChange({ ...value, ...patch });
  const crossReferenceEnabled = Boolean(value.crossReference);
  const mainEntryId = useId();
  const subEntryId = useId();
  const crossReferenceId = useId();

  return (
    <Dialog
      title={officeMessage(
        messages,
        editing
          ? 'document.indexEntry.title.edit'
          : 'document.indexEntry.title.insert',
      )}
      description={officeMessage(messages, 'document.indexEntry.description')}
      className="work-document-index-entry-dialog"
      restoreFocusTarget={restoreFocusTarget}
      onClose={onCancel}
      footer={
        <>
          <Button tone="quiet" onClick={onCancel}>
            {officeMessage(messages, 'document.indexEntry.cancel')}
          </Button>
          <Button
            tone="primary"
            disabled={!value.mainEntry.trim()}
            onClick={onSubmit}
          >
            {officeMessage(
              messages,
              editing
                ? 'document.indexEntry.apply'
                : 'document.indexEntry.mark',
            )}
          </Button>
        </>
      }
    >
      <div className="work-document-index-entry-dialog-fields">
        <label htmlFor={mainEntryId}>
          <span>{officeMessage(messages, 'document.indexEntry.main')}</span>
          <OfficeTextField
            id={mainEntryId}
            data-autofocus
            aria-label={officeMessage(messages, 'document.indexEntry.main')}
            value={value.mainEntry}
            maxLength={240}
            onChange={(event) => update({ mainEntry: event.target.value })}
            onKeyDown={(event) => {
              if (event.key !== 'Enter' || !value.mainEntry.trim()) return;
              event.preventDefault();
              onSubmit();
            }}
          />
        </label>
        <label htmlFor={subEntryId}>
          <span>{officeMessage(messages, 'document.indexEntry.sub')}</span>
          <OfficeTextField
            id={subEntryId}
            aria-label={officeMessage(messages, 'document.indexEntry.sub')}
            value={value.subEntry}
            maxLength={240}
            placeholder={officeMessage(
              messages,
              'document.indexEntry.optional',
            )}
            onChange={(event) => update({ subEntry: event.target.value })}
          />
        </label>
      </div>
      <div className="work-document-index-entry-dialog-mode">
        <OfficeCheckbox
          ariaLabel={officeMessage(
            messages,
            'document.indexEntry.useCrossRef',
          )}
          checked={crossReferenceEnabled}
          onCheckedChange={(enabled) =>
            update({
              crossReference: enabled
                ? value.crossReference ||
                  officeMessage(
                    messages,
                    'document.indexEntry.crossRefDefault',
                  )
                : '',
              pageBold: enabled ? false : value.pageBold,
              pageItalic: enabled ? false : value.pageItalic,
            })
          }
        >
          {officeMessage(messages, 'document.indexEntry.useCrossRef')}
        </OfficeCheckbox>
        {crossReferenceEnabled && (
          <label htmlFor={crossReferenceId}>
            <span>
              {officeMessage(messages, 'document.indexEntry.crossRefTarget')}
            </span>
            <OfficeTextField
              id={crossReferenceId}
              aria-label={officeMessage(
                messages,
                'document.indexEntry.crossRefTargetAria',
              )}
              value={value.crossReference}
              maxLength={240}
              placeholder={officeMessage(
                messages,
                'document.indexEntry.crossRefPlaceholder',
              )}
              onChange={(event) =>
                update({ crossReference: event.target.value })
              }
            />
          </label>
        )}
      </div>
      <fieldset
        className="work-document-index-entry-dialog-page-style"
        disabled={crossReferenceEnabled}
      >
        <legend>
          {officeMessage(messages, 'document.indexEntry.pageFormatLegend')}
        </legend>
        <OfficeCheckbox
          ariaLabel={officeMessage(messages, 'document.indexEntry.pageBold')}
          checked={value.pageBold}
          disabled={crossReferenceEnabled}
          onCheckedChange={(pageBold) => update({ pageBold })}
        >
          {officeMessage(messages, 'document.indexEntry.pageBold')}
        </OfficeCheckbox>
        <OfficeCheckbox
          ariaLabel={officeMessage(messages, 'document.indexEntry.pageItalic')}
          checked={value.pageItalic}
          disabled={crossReferenceEnabled}
          onCheckedChange={(pageItalic) => update({ pageItalic })}
        >
          {officeMessage(messages, 'document.indexEntry.pageItalic')}
        </OfficeCheckbox>
      </fieldset>
    </Dialog>
  );
}
