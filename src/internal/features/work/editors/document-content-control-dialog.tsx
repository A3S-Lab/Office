import { Button, Dialog } from '../../../design-system/primitives';
import { officeMessage } from '../../../i18n/office-locale';
import type { OfficeMessageCatalog } from '../../../i18n/office-messages';
import type {
  WorkDocumentContentControlAppearance,
  WorkDocumentContentControlListItem,
  WorkDocumentContentControlLock,
  WorkDocumentContentControlProperties,
  WorkDocumentContentControlType,
} from '../work-document-content-control';
import {
  contentControlDateInputValue,
  contentControlFullDateFromInput,
  DOCUMENT_CONTENT_CONTROL_DATE_FORMATS,
  DOCUMENT_CONTENT_CONTROL_DEFAULT_DATE_FORMAT,
  DOCUMENT_CONTENT_CONTROL_DEFAULT_DATE_LANGUAGE,
  isListStyleContentControl,
} from '../work-document-content-control';
import {
  OfficeCheckbox,
  OfficeSelect,
  OfficeTextField,
} from './office-controls';
import { useOfficeMessages } from './office-messages-context';

export interface DocumentContentControlDialogProps {
  editing: boolean;
  properties: WorkDocumentContentControlProperties;
  restoreFocusTarget: () => HTMLElement | null;
  onCancel: () => void;
  onChange: (properties: WorkDocumentContentControlProperties) => void;
  onSubmit: () => void;
}

const TYPE_KEYS = [
  'text',
  'richText',
  'checkbox',
  'dropDownList',
  'comboBox',
  'date',
] as const satisfies readonly WorkDocumentContentControlType[];

const LOCK_KEYS = [
  'unlocked',
  'contentLocked',
  'sdtLocked',
  'sdtContentLocked',
] as const satisfies readonly WorkDocumentContentControlLock[];

const APPEARANCE_KEYS = [
  'boundingBox',
  'tags',
  'hidden',
] as const satisfies readonly WorkDocumentContentControlAppearance[];

export function DocumentContentControlDialog({
  editing,
  properties,
  restoreFocusTarget,
  onCancel,
  onChange,
  onSubmit,
}: DocumentContentControlDialogProps) {
  const messages = useOfficeMessages();
  const update = (patch: Partial<WorkDocumentContentControlProperties>) =>
    onChange({ ...properties, ...patch });
  const aliasId = 'document-content-control-alias';
  const tagId = 'document-content-control-tag';
  const optionsId = 'document-content-control-options';
  const dateId = 'document-content-control-date';
  const isCheckbox = properties.type === 'checkbox';
  const isList = isListStyleContentControl(properties.type);
  const isComboBox = properties.type === 'comboBox';
  const isDate = properties.type === 'date';
  const optionsText = listItemsToLines(properties.options);
  const canSubmit =
    (!isList || properties.options.length > 0) &&
    (!isDate || Boolean(properties.fullDate));
  const dateFormatOptions = DOCUMENT_CONTENT_CONTROL_DATE_FORMATS.map(
    (value) => ({ value, label: value }),
  );
  const dateFormatSelectOptions = dateFormatOptions.some(
    (item) => item.value === properties.dateFormat,
  )
    ? dateFormatOptions
    : [
        { value: properties.dateFormat, label: properties.dateFormat },
        ...dateFormatOptions,
      ];
  const defaultOption = officeMessage(
    messages,
    'document.contentControl.defaultOption',
  );

  return (
    <Dialog
      title={officeMessage(
        messages,
        editing
          ? 'document.contentControl.title.edit'
          : 'document.contentControl.title.insert',
      )}
      description={officeMessage(
        messages,
        'document.contentControl.description',
      )}
      className="work-document-content-control-dialog"
      restoreFocusTarget={restoreFocusTarget}
      onClose={onCancel}
      footer={
        <>
          <Button tone="quiet" onClick={onCancel}>
            {officeMessage(messages, 'document.contentControl.cancel')}
          </Button>
          <Button tone="primary" disabled={!canSubmit} onClick={onSubmit}>
            {officeMessage(
              messages,
              editing
                ? 'document.contentControl.submit.apply'
                : 'document.contentControl.submit.insert',
            )}
          </Button>
        </>
      }
    >
      <div className="work-document-content-control-dialog-fields">
        <label htmlFor={aliasId}>
          <span>
            {officeMessage(messages, 'document.contentControl.alias')}
          </span>
          <OfficeTextField
            id={aliasId}
            data-autofocus
            aria-label={officeMessage(
              messages,
              'document.contentControl.aliasAria',
            )}
            value={properties.alias}
            maxLength={255}
            placeholder={officeMessage(
              messages,
              'document.contentControl.aliasPlaceholder',
            )}
            onChange={(event) => update({ alias: event.target.value })}
            onKeyDown={(event) => {
              if (event.key !== 'Enter' || event.nativeEvent.isComposing)
                return;
              event.preventDefault();
              if (canSubmit) onSubmit();
            }}
          />
        </label>
        <label htmlFor={tagId}>
          <span>{officeMessage(messages, 'document.contentControl.tag')}</span>
          <OfficeTextField
            id={tagId}
            aria-label={officeMessage(
              messages,
              'document.contentControl.tagAria',
            )}
            value={properties.tag}
            maxLength={255}
            placeholder={officeMessage(
              messages,
              'document.contentControl.tagPlaceholder',
            )}
            onChange={(event) => update({ tag: event.target.value })}
          />
        </label>
      </div>
      <div className="work-document-content-control-dialog-selects">
        <div className="work-document-content-control-dialog-field">
          <span>
            {officeMessage(messages, 'document.contentControl.type')}
          </span>
          <OfficeSelect<WorkDocumentContentControlType>
            ariaLabel={officeMessage(
              messages,
              'document.contentControl.typeAria',
            )}
            value={properties.type}
            options={typeOptions(messages)}
            onValueChange={(type) =>
              update({
                type,
                multiLine:
                  type === 'checkbox' ||
                  type === 'date' ||
                  isListStyleContentControl(type)
                    ? false
                    : properties.multiLine,
                checked: type === 'checkbox' ? properties.checked : false,
                options: isListStyleContentControl(type)
                  ? properties.options.length
                    ? properties.options
                    : [{ displayText: defaultOption, value: defaultOption }]
                  : [],
                selectedValue: isListStyleContentControl(type)
                  ? type === 'comboBox'
                    ? properties.selectedValue &&
                      properties.options.some(
                        (item) => item.value === properties.selectedValue,
                      )
                      ? properties.selectedValue
                      : properties.options[0]?.value || defaultOption
                    : properties.selectedValue ||
                      properties.options[0]?.value ||
                      defaultOption
                  : '',
                fullDate:
                  type === 'date'
                    ? properties.fullDate || contentControlFullDateFromInput('')
                    : '',
                dateFormat:
                  type === 'date'
                    ? properties.dateFormat ||
                      DOCUMENT_CONTENT_CONTROL_DEFAULT_DATE_FORMAT
                    : DOCUMENT_CONTENT_CONTROL_DEFAULT_DATE_FORMAT,
                dateLanguage:
                  type === 'date'
                    ? properties.dateLanguage ||
                      DOCUMENT_CONTENT_CONTROL_DEFAULT_DATE_LANGUAGE
                    : DOCUMENT_CONTENT_CONTROL_DEFAULT_DATE_LANGUAGE,
                dateMapping:
                  type === 'date' ? properties.dateMapping : 'dateTime',
              })
            }
          />
        </div>
        <div className="work-document-content-control-dialog-field">
          <span>
            {officeMessage(messages, 'document.contentControl.lock')}
          </span>
          <OfficeSelect<WorkDocumentContentControlLock>
            ariaLabel={officeMessage(
              messages,
              'document.contentControl.lockAria',
            )}
            value={properties.lock}
            options={lockOptions(messages)}
            onValueChange={(lock) => update({ lock })}
          />
        </div>
        <div className="work-document-content-control-dialog-field">
          <span>
            {officeMessage(messages, 'document.contentControl.appearance')}
          </span>
          <OfficeSelect<WorkDocumentContentControlAppearance>
            ariaLabel={officeMessage(
              messages,
              'document.contentControl.appearanceAria',
            )}
            value={properties.appearance}
            options={appearanceOptions(messages)}
            onValueChange={(appearance) => update({ appearance })}
          />
        </div>
      </div>
      {isList ? (
        <div className="work-document-content-control-dialog-fields">
          <label htmlFor={optionsId}>
            <span>
              {officeMessage(messages, 'document.contentControl.options')}
            </span>
            <textarea
              id={optionsId}
              aria-label={officeMessage(
                messages,
                isComboBox
                  ? 'document.contentControl.optionsAria.comboBox'
                  : 'document.contentControl.optionsAria.dropDown',
              )}
              className="work-document-content-control-options"
              rows={4}
              value={optionsText}
              placeholder={officeMessage(
                messages,
                'document.contentControl.optionsPlaceholder',
              )}
              onChange={(event) => {
                const options = linesToListItems(event.target.value);
                update({
                  options,
                  selectedValue: options.some(
                    (item) => item.value === properties.selectedValue,
                  )
                    ? properties.selectedValue
                    : (options[0]?.value ?? ''),
                });
              }}
            />
          </label>
          <div className="work-document-content-control-dialog-field">
            <span>
              {officeMessage(
                messages,
                isComboBox
                  ? 'document.contentControl.selected.comboBox'
                  : 'document.contentControl.selected.dropDown',
              )}
            </span>
            <OfficeSelect<string>
              ariaLabel={officeMessage(
                messages,
                isComboBox
                  ? 'document.contentControl.selectedAria.comboBox'
                  : 'document.contentControl.selectedAria.dropDown',
              )}
              value={
                properties.selectedValue || properties.options[0]?.value || ''
              }
              options={properties.options.map((item) => ({
                value: item.value,
                label: item.displayText,
              }))}
              disabled={!properties.options.length}
              onValueChange={(selectedValue) => update({ selectedValue })}
            />
          </div>
        </div>
      ) : null}
      {isDate ? (
        <div className="work-document-content-control-dialog-fields">
          <label htmlFor={dateId}>
            <span>
              {officeMessage(messages, 'document.contentControl.date')}
            </span>
            <input
              id={dateId}
              type="date"
              aria-label={officeMessage(
                messages,
                'document.contentControl.dateAria',
              )}
              className="work-document-content-control-date"
              value={contentControlDateInputValue(properties.fullDate)}
              onChange={(event) =>
                update({
                  fullDate: contentControlFullDateFromInput(event.target.value),
                })
              }
            />
          </label>
          <div className="work-document-content-control-dialog-field">
            <span>
              {officeMessage(messages, 'document.contentControl.dateFormat')}
            </span>
            <OfficeSelect<string>
              ariaLabel={officeMessage(
                messages,
                'document.contentControl.dateFormatAria',
              )}
              value={properties.dateFormat}
              options={dateFormatSelectOptions}
              onValueChange={(dateFormat) => update({ dateFormat })}
            />
          </div>
        </div>
      ) : null}
      <div className="work-document-content-control-dialog-options">
        {isCheckbox ? (
          <OfficeCheckbox
            ariaLabel={officeMessage(
              messages,
              'document.contentControl.checkedAria',
            )}
            checked={properties.checked}
            onCheckedChange={(checked) => update({ checked })}
          >
            {officeMessage(messages, 'document.contentControl.checked')}
          </OfficeCheckbox>
        ) : isList || isDate ? null : (
          <OfficeCheckbox
            ariaLabel={officeMessage(
              messages,
              'document.contentControl.multiLineAria',
            )}
            checked={properties.multiLine}
            onCheckedChange={(multiLine) => update({ multiLine })}
          >
            {officeMessage(messages, 'document.contentControl.multiLine')}
          </OfficeCheckbox>
        )}
        <label className="work-document-content-control-color">
          <span>
            {officeMessage(messages, 'document.contentControl.color')}
          </span>
          <input
            type="color"
            aria-label={officeMessage(
              messages,
              'document.contentControl.colorAria',
            )}
            value={properties.color ?? '#2f6fed'}
            onChange={(event) => update({ color: event.target.value })}
          />
          <button
            type="button"
            className="work-document-content-control-color-reset"
            onClick={() => update({ color: null })}
          >
            {officeMessage(messages, 'document.contentControl.colorTheme')}
          </button>
        </label>
      </div>
    </Dialog>
  );
}

function typeOptions(messages: OfficeMessageCatalog) {
  return TYPE_KEYS.map((value) => ({
    value,
    label: officeMessage(messages, `document.contentControl.type.${value}`),
  }));
}

function lockOptions(messages: OfficeMessageCatalog) {
  return LOCK_KEYS.map((value) => ({
    value,
    label: officeMessage(messages, `document.contentControl.lock.${value}`),
  }));
}

function appearanceOptions(messages: OfficeMessageCatalog) {
  return APPEARANCE_KEYS.map((value) => ({
    value,
    label: officeMessage(
      messages,
      `document.contentControl.appearance.${value}`,
    ),
  }));
}

function listItemsToLines(
  options: WorkDocumentContentControlListItem[],
): string {
  return options
    .map((item) =>
      item.displayText === item.value
        ? item.displayText
        : `${item.displayText}|${item.value}`,
    )
    .join('\n');
}

function linesToListItems(
  source: string,
): WorkDocumentContentControlListItem[] {
  const items: WorkDocumentContentControlListItem[] = [];
  const seen = new Set<string>();
  for (const line of source.split(/\r?\n/u)) {
    const trimmed = line.trim();
    if (!trimmed) continue;
    const separator = trimmed.indexOf('|');
    const displayText =
      separator >= 0 ? trimmed.slice(0, separator).trim() : trimmed;
    const value =
      separator >= 0
        ? trimmed.slice(separator + 1).trim() || displayText
        : displayText;
    if (!displayText || !value || seen.has(value) || items.length >= 32)
      continue;
    seen.add(value);
    items.push({ displayText, value });
  }
  return items;
}
