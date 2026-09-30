import { useState } from 'react';
import { Button, Dialog } from '../../../design-system/primitives';
import { officeMessage } from '../../../i18n/office-locale';
import {
  parseMailMergeTable,
  type WorkDocumentMailMergeSource,
} from '../work-document-mail-merge';
import { OfficeTextArea } from './office-controls';
import { useOfficeMessages } from './office-messages-context';

export interface DocumentMailMergeImportDialogProps {
  restoreFocusTarget: () => HTMLElement | null;
  onCancel: () => void;
  onImport: (source: WorkDocumentMailMergeSource) => void;
}

export function DocumentMailMergeImportDialog({
  restoreFocusTarget,
  onCancel,
  onImport,
}: DocumentMailMergeImportDialogProps) {
  const messages = useOfficeMessages();
  const [draft, setDraft] = useState('');
  const [invalid, setInvalid] = useState(false);

  const submit = () => {
    const parsed = parseMailMergeTable(draft);
    if (!parsed) {
      setInvalid(true);
      return;
    }
    onImport(parsed);
  };

  return (
    <Dialog
      title={officeMessage(messages, 'document.mailMerge.import.title')}
      description={officeMessage(
        messages,
        'document.mailMerge.import.description',
      )}
      className="work-document-mail-merge-import-dialog"
      focusKey="document-mail-merge-import"
      restoreFocusTarget={restoreFocusTarget}
      onClose={onCancel}
      footer={
        <>
          <Button tone="quiet" onClick={onCancel}>
            {officeMessage(messages, 'document.mailMerge.import.cancel')}
          </Button>
          <Button tone="primary" onClick={submit}>
            {officeMessage(messages, 'document.mailMerge.import.confirm')}
          </Button>
        </>
      }
    >
      <label
        className="work-document-mail-merge-import-field"
        htmlFor="mail-merge-import-table"
      >
        <span>
          {officeMessage(messages, 'document.mailMerge.import.field')}
        </span>
        <OfficeTextArea
          id="mail-merge-import-table"
          aria-label={officeMessage(
            messages,
            'document.mailMerge.import.field',
          )}
          aria-invalid={invalid || undefined}
          rows={8}
          value={draft}
          placeholder={officeMessage(
            messages,
            'document.mailMerge.import.placeholder',
          )}
          onChange={(event) => {
            setDraft(event.currentTarget.value);
            setInvalid(false);
          }}
        />
      </label>
      {invalid ? (
        <p className="work-document-mail-merge-import-invalid" role="alert">
          {officeMessage(messages, 'document.mailMerge.import.invalid')}
        </p>
      ) : null}
    </Dialog>
  );
}
