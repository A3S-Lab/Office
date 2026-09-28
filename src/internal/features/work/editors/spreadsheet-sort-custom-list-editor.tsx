import { useEffect, useRef } from 'react';
import { officeMessage } from '../../../i18n/office-locale';
import { Button } from '../../../design-system/primitives';
import { useOfficeMessages } from './office-messages-context';

export function SpreadsheetSortCustomListEditor({
  error,
  level,
  text,
  onCancel,
  onChange,
  onUse,
}: {
  error: string | null;
  level: number;
  text: string;
  onCancel: () => void;
  onChange: (text: string) => void;
  onUse: () => void;
}) {
  const messages = useOfficeMessages();
  const textareaRef = useRef<HTMLTextAreaElement>(null);
  useEffect(() => textareaRef.current?.focus(), []);

  return (
    <div className="work-spreadsheet-sort-custom-list-editor">
      <label>
        <span>{officeMessage(messages, 'spreadsheet.sort.listEditor.label')}</span>
        <textarea
          ref={textareaRef}
          aria-label={officeMessage(messages, 'spreadsheet.sort.listEditor.aria', {
            level: String(level),
          })}
          aria-invalid={error ? true : undefined}
          rows={5}
          value={text}
          onChange={(event) => onChange(event.currentTarget.value)}
        />
      </label>
      <p>{officeMessage(messages, 'spreadsheet.sort.listEditor.hint')}</p>
      {error ? (
        <p className="work-spreadsheet-sort-custom-list-error" role="alert">
          {error}
        </p>
      ) : null}
      <div className="work-spreadsheet-sort-custom-list-actions">
        <Button tone="primary" type="button" onClick={onUse}>
          {officeMessage(messages, 'spreadsheet.sort.listEditor.use')}
        </Button>
        <Button tone="quiet" type="button" onClick={onCancel}>
          {officeMessage(messages, 'spreadsheet.sort.listEditor.cancelEdit')}
        </Button>
      </div>
    </div>
  );
}
