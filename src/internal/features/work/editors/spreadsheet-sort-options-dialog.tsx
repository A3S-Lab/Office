import { type FormEvent, useId, useState } from 'react';
import { officeMessage } from '../../../i18n/office-locale';
import { Button, Dialog } from '../../../design-system/primitives';
import { useOfficeMessages } from './office-messages-context';
import type { SpreadsheetSortOptions } from './spreadsheet-sort';

export function SpreadsheetSortOptionsDialog({
  value,
  orientationLocked = false,
  restoreFocusTarget,
  onApply,
  onClose,
}: {
  value: SpreadsheetSortOptions;
  orientationLocked?: boolean;
  restoreFocusTarget: () => HTMLElement | null;
  onApply: (value: SpreadsheetSortOptions) => void;
  onClose: () => void;
}) {
  const messages = useOfficeMessages();
  const [draft, setDraft] = useState(value);
  const formId = useId();
  const dirty =
    draft.caseSensitive !== value.caseSensitive ||
    draft.textMethod !== value.textMethod ||
    draft.orientation !== value.orientation;

  const resetDraft = () => {
    setDraft(value);
  };

  const submit = (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    onApply(draft);
  };

  return (
    <Dialog
      title={officeMessage(messages, 'spreadsheet.sort.options.title')}
      description={officeMessage(messages, 'spreadsheet.sort.options.desc')}
      className="work-spreadsheet-sort-options-dialog"
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
            {officeMessage(messages, 'spreadsheet.sort.cancel')}
          </Button>
          <Button tone="primary" type="submit" form={formId}>
            {officeMessage(messages, 'spreadsheet.sort.ok')}
          </Button>
        </>
      }
    >
      <form
        id={formId}
        className="work-spreadsheet-sort-options-form"
        onSubmit={submit}
      >
        <fieldset className="work-spreadsheet-sort-options">
          <legend>
            {officeMessage(messages, 'spreadsheet.sort.options.textCompare')}
          </legend>
          <label>
            <input
              type="checkbox"
              aria-label={officeMessage(
                messages,
                'spreadsheet.sort.options.caseSensitive',
              )}
              checked={draft.caseSensitive}
              onChange={(event) => {
                const caseSensitive = event.currentTarget.checked;
                setDraft((current) => ({
                  ...current,
                  caseSensitive,
                }));
              }}
            />
            <span>
              <strong>
                {officeMessage(
                  messages,
                  'spreadsheet.sort.options.caseSensitive',
                )}
              </strong>
              <small>
                {officeMessage(
                  messages,
                  'spreadsheet.sort.options.caseSensitiveHint',
                )}
              </small>
            </span>
          </label>
        </fieldset>
        <fieldset className="work-spreadsheet-sort-options">
          <legend>
            {officeMessage(messages, 'spreadsheet.sort.options.method')}
          </legend>
          <label>
            <input
              type="radio"
              name="spreadsheet-sort-text-method"
              value="pinyin"
              aria-label={officeMessage(
                messages,
                'spreadsheet.sort.options.pinyin',
              )}
              checked={draft.textMethod === 'pinyin'}
              onChange={() =>
                setDraft((current) => ({
                  ...current,
                  textMethod: 'pinyin',
                }))
              }
            />
            <span>
              <strong>
                {officeMessage(messages, 'spreadsheet.sort.options.pinyin')}
              </strong>
              <small>
                {officeMessage(messages, 'spreadsheet.sort.options.pinyinHint')}
              </small>
            </span>
          </label>
          <label>
            <input
              type="radio"
              name="spreadsheet-sort-text-method"
              value="stroke"
              aria-label={officeMessage(
                messages,
                'spreadsheet.sort.options.stroke',
              )}
              checked={draft.textMethod === 'stroke'}
              onChange={() =>
                setDraft((current) => ({
                  ...current,
                  textMethod: 'stroke',
                }))
              }
            />
            <span>
              <strong>
                {officeMessage(messages, 'spreadsheet.sort.options.stroke')}
              </strong>
              <small>
                {officeMessage(messages, 'spreadsheet.sort.options.strokeHint')}
              </small>
            </span>
          </label>
        </fieldset>
        <fieldset className="work-spreadsheet-sort-options">
          <legend>
            {officeMessage(messages, 'spreadsheet.sort.options.direction')}
          </legend>
          {orientationLocked ? (
            <p className="work-spreadsheet-sort-options-lock-note">
              {officeMessage(
                messages,
                'spreadsheet.sort.options.structuredOnlyColumns',
              )}
            </p>
          ) : null}
          <label>
            <input
              type="radio"
              name="spreadsheet-sort-orientation"
              value="top-to-bottom"
              checked={draft.orientation === 'top-to-bottom'}
              data-autofocus={
                draft.orientation === 'top-to-bottom' ? '' : undefined
              }
              onChange={() =>
                setDraft((current) => ({
                  ...current,
                  orientation: 'top-to-bottom',
                }))
              }
            />
            <span>
              <strong>
                {officeMessage(messages, 'spreadsheet.sort.options.byColumn')}
              </strong>
              <small>
                {officeMessage(
                  messages,
                  'spreadsheet.sort.options.byColumnHint',
                )}
              </small>
            </span>
          </label>
          <label>
            <input
              type="radio"
              name="spreadsheet-sort-orientation"
              value="left-to-right"
              disabled={orientationLocked}
              checked={draft.orientation === 'left-to-right'}
              data-autofocus={
                draft.orientation === 'left-to-right' ? '' : undefined
              }
              onChange={() =>
                setDraft((current) => ({
                  ...current,
                  orientation: 'left-to-right',
                }))
              }
            />
            <span>
              <strong>
                {officeMessage(messages, 'spreadsheet.sort.options.byRow')}
              </strong>
              <small>
                {officeMessage(messages, 'spreadsheet.sort.options.byRowHint')}
              </small>
            </span>
          </label>
        </fieldset>
      </form>
    </Dialog>
  );
}
