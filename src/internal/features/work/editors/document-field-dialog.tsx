import { Button, Dialog } from '../../../design-system/primitives';
import { officeMessage } from '../../../i18n/office-locale';
import type { OfficeMessageCatalog } from '../../../i18n/office-messages';
import type {
  WorkDocumentClockFieldFormat,
  WorkDocumentFieldDraft,
  WorkDocumentFieldKind,
  WorkDocumentNumericFieldFormat,
} from '../work-document-fields';
import { DOCUMENT_FIELD_DATE_FORMAT_ZH } from '../work-document-fields';
import {
  OfficeCheckbox,
  OfficeSelect,
  type OfficeSelectOption,
  OfficeTextField,
} from './office-controls';
import { useOfficeMessages } from './office-messages-context';

export interface DocumentFieldTargetOption {
  id: string;
  name: string;
}

export interface DocumentFieldDialogProps {
  editing: boolean;
  draft: WorkDocumentFieldDraft;
  preview: string;
  targets: readonly DocumentFieldTargetOption[];
  restoreFocusTarget: () => HTMLElement | null;
  onCancel: () => void;
  onChange: (draft: WorkDocumentFieldDraft) => void;
  onSubmit: () => void;
}

const MERGE_FIELD_NAME_PATTERN = /^[\p{L}_][\p{L}\p{N}_]*$/u;

const FIELD_KIND_KEYS = [
  'page',
  'numPages',
  'section',
  'sectionPages',
  'date',
  'time',
  'createDate',
  'saveDate',
  'printDate',
  'wordCount',
  'characterCount',
  'fileName',
  'author',
  'title',
  'subject',
  'keywords',
  'lastSavedBy',
  'comments',
  'mergeField',
  'pageReference',
] as const satisfies readonly WorkDocumentFieldKind[];

const NUMERIC_FORMAT_KEYS = [
  'arabic',
  'roman',
  'romanLower',
  'alphabetic',
  'alphabeticLower',
  'ordinal',
] as const satisfies readonly WorkDocumentNumericFieldFormat[];

export function DocumentFieldDialog({
  editing,
  draft,
  preview,
  targets,
  restoreFocusTarget,
  onCancel,
  onChange,
  onSubmit,
}: DocumentFieldDialogProps) {
  const messages = useOfficeMessages();
  const numeric = isNumericField(draft.kind);
  const clockKind =
    draft.kind === 'date' ||
    draft.kind === 'createDate' ||
    draft.kind === 'saveDate' ||
    draft.kind === 'printDate'
      ? 'date'
      : draft.kind === 'time'
        ? 'time'
        : null;
  const clock = clockKind !== null;
  const hasTarget =
    draft.kind !== 'pageReference' ||
    targets.some(
      (target) =>
        Boolean(draft.targetName) &&
        target.id === draft.targetId &&
        target.name === draft.targetName,
    );
  const hasMergeFieldName =
    draft.kind !== 'mergeField' ||
    MERGE_FIELD_NAME_PATTERN.test(draft.targetName.trim());
  const targetValue = `${draft.targetId}:${draft.targetName}`;
  const canSubmit =
    (draft.kind !== 'pageReference' || hasTarget) && hasMergeFieldName;

  return (
    <Dialog
      title={officeMessage(
        messages,
        editing ? 'document.field.title.edit' : 'document.field.title.insert',
      )}
      description={officeMessage(messages, 'document.field.description')}
      className="work-document-field-dialog"
      restoreFocusTarget={restoreFocusTarget}
      onClose={onCancel}
      footer={
        <>
          <Button tone="quiet" onClick={onCancel}>
            {officeMessage(messages, 'document.field.cancel')}
          </Button>
          <Button tone="primary" disabled={!canSubmit} onClick={onSubmit}>
            {officeMessage(
              messages,
              editing
                ? 'document.field.submit.apply'
                : 'document.field.submit.insert',
            )}
          </Button>
        </>
      }
    >
      <div className="work-document-field-dialog-grid">
        <div className="work-document-dialog-field">
          <span>{officeMessage(messages, 'document.field.kind')}</span>
          <OfficeSelect
            initialFocus
            ariaLabel={officeMessage(messages, 'document.field.kindAria')}
            value={draft.kind}
            options={fieldKindOptions(messages)}
            onValueChange={(value) => {
              if (!value) return;
              const kind = value as WorkDocumentFieldKind;
              onChange({
                ...draft,
                kind,
                format: defaultFormat(kind),
                targetId: kind === 'pageReference' ? draft.targetId : '',
                targetName:
                  kind === 'pageReference' || kind === 'mergeField'
                    ? draft.targetName
                    : '',
                hyperlink: kind === 'pageReference' ? draft.hyperlink : false,
              });
            }}
          />
        </div>

        {numeric && (
          <div className="work-document-dialog-field">
            <span>
              {officeMessage(messages, 'document.field.numericFormat')}
            </span>
            <OfficeSelect
              ariaLabel={officeMessage(
                messages,
                'document.field.numericFormatAria',
              )}
              value={
                draft.format.kind === 'numeric' ? draft.format.value : 'arabic'
              }
              options={numericFormatOptions(messages)}
              onValueChange={(value) => {
                if (!value) return;
                onChange({
                  ...draft,
                  format: {
                    kind: 'numeric',
                    value: value as WorkDocumentNumericFieldFormat,
                  },
                });
              }}
            />
          </div>
        )}

        {clock && (
          <div className="work-document-dialog-field">
            <span>
              {officeMessage(
                messages,
                draft.kind === 'time'
                  ? 'document.field.timeFormat'
                  : 'document.field.dateFormat',
              )}
            </span>
            <OfficeSelect
              ariaLabel={officeMessage(
                messages,
                clockKind === 'date'
                  ? 'document.field.dateFormat'
                  : 'document.field.timeFormat',
              )}
              value={
                draft.format.kind === 'clock'
                  ? draft.format.source
                    ? '__preserved__'
                    : draft.format.value
                  : defaultClockFormat(clockKind ?? 'date')
              }
              options={clockFormatOptions(messages, clockKind ?? 'date', draft)}
              onValueChange={(value) => {
                if (!value || value === '__preserved__') return;
                onChange({
                  ...draft,
                  format: {
                    kind: 'clock',
                    value: value as WorkDocumentClockFieldFormat,
                  },
                });
              }}
            />
          </div>
        )}

        {draft.kind === 'mergeField' && (
          <div className="work-document-dialog-field">
            <span>{officeMessage(messages, 'document.field.mergeName')}</span>
            <OfficeTextField
              aria-label={officeMessage(
                messages,
                'document.field.mergeNameAria',
              )}
              value={draft.targetName}
              maxLength={64}
              placeholder={officeMessage(
                messages,
                'document.field.mergeNamePlaceholder',
              )}
              onChange={(event) =>
                onChange({ ...draft, targetName: event.target.value })
              }
            />
            {!hasMergeFieldName && (
              <small className="work-document-field-dialog-help">
                {officeMessage(messages, 'document.field.mergeNameHelp')}
              </small>
            )}
          </div>
        )}

        {draft.kind === 'pageReference' && (
          <div className="work-document-dialog-field">
            <span>{officeMessage(messages, 'document.field.target')}</span>
            <OfficeSelect
              ariaLabel={officeMessage(messages, 'document.field.targetAria')}
              value={targetValue}
              options={targets.map((target) => ({
                value: `${target.id}:${target.name}`,
                label: target.name,
              }))}
              disabled={!targets.length}
              onValueChange={(value) => {
                const target = targets.find(
                  (candidate) => `${candidate.id}:${candidate.name}` === value,
                );
                if (!target) return;
                onChange({
                  ...draft,
                  targetId: target.id,
                  targetName: target.name,
                });
              }}
            />
            {!hasTarget && (
              <small className="work-document-field-dialog-help">
                {officeMessage(
                  messages,
                  targets.length
                    ? 'document.field.targetMissing'
                    : 'document.field.targetEmpty',
                )}
              </small>
            )}
          </div>
        )}
      </div>

      <div className="work-document-field-dialog-options">
        {draft.kind === 'pageReference' && (
          <OfficeCheckbox
            ariaLabel={officeMessage(messages, 'document.field.hyperlink')}
            checked={draft.hyperlink}
            onCheckedChange={(hyperlink) => onChange({ ...draft, hyperlink })}
          >
            {officeMessage(messages, 'document.field.hyperlink')}
          </OfficeCheckbox>
        )}

        <OfficeCheckbox
          ariaLabel={officeMessage(messages, 'document.field.mergeFormat')}
          checked={draft.mergeFormat}
          onCheckedChange={(mergeFormat) => onChange({ ...draft, mergeFormat })}
        >
          {officeMessage(messages, 'document.field.mergeFormat')}
        </OfficeCheckbox>
      </div>

      <div className="work-document-field-dialog-preview">
        <span>{officeMessage(messages, 'document.field.preview')}</span>
        <output
          aria-label={officeMessage(messages, 'document.field.previewAria')}
          aria-live="polite"
        >
          {preview || '—'}
        </output>
      </div>
      <p className="work-document-field-dialog-note">
        {draft.kind === 'mergeField'
          ? draft.mergeFormat
            ? officeMessage(messages, 'document.field.note.merge.mergeFormat')
            : officeMessage(messages, 'document.field.note.merge.plain')
          : draft.mergeFormat
            ? officeMessage(messages, 'document.field.note.mergeFormat')
            : officeMessage(messages, 'document.field.note.plain')}
      </p>
    </Dialog>
  );
}

function fieldKindOptions(
  messages: OfficeMessageCatalog,
): readonly OfficeSelectOption<WorkDocumentFieldKind>[] {
  return FIELD_KIND_KEYS.map((value) => ({
    value,
    label: officeMessage(messages, `document.field.kind.${value}`),
  }));
}

function numericFormatOptions(
  messages: OfficeMessageCatalog,
): readonly OfficeSelectOption<WorkDocumentNumericFieldFormat>[] {
  return NUMERIC_FORMAT_KEYS.map((value) => ({
    value,
    label: officeMessage(messages, `document.field.numeric.${value}`),
  }));
}

function dateFormatOptions(
  messages: OfficeMessageCatalog,
): readonly OfficeSelectOption<WorkDocumentClockFieldFormat>[] {
  return [
    {
      value: DOCUMENT_FIELD_DATE_FORMAT_ZH,
      label: officeMessage(messages, 'document.field.dateExample.zh'),
    },
    {
      value: 'yyyy-MM-dd',
      label: officeMessage(messages, 'document.field.dateExample.iso'),
    },
    {
      value: 'MMMM d, yyyy',
      label: officeMessage(messages, 'document.field.dateExample.long'),
    },
    {
      value: 'dddd, MMMM d, yyyy',
      label: officeMessage(messages, 'document.field.dateExample.full'),
    },
  ];
}

function timeFormatOptions(): readonly OfficeSelectOption<WorkDocumentClockFieldFormat>[] {
  return [
    { value: 'HH:mm', label: '14:05' },
    { value: 'HH:mm:ss', label: '14:05:09' },
    { value: 'h:mm AM/PM', label: '2:05 PM' },
  ];
}

function isNumericField(kind: WorkDocumentFieldKind): boolean {
  return (
    kind === 'page' ||
    kind === 'numPages' ||
    kind === 'section' ||
    kind === 'sectionPages' ||
    kind === 'pageReference'
  );
}

function defaultFormat(kind: WorkDocumentFieldKind) {
  if (isNumericField(kind))
    return { kind: 'numeric' as const, value: 'arabic' as const };
  if (
    kind === 'date' ||
    kind === 'createDate' ||
    kind === 'saveDate' ||
    kind === 'printDate'
  ) {
    return {
      kind: 'clock' as const,
      value: DOCUMENT_FIELD_DATE_FORMAT_ZH,
    };
  }
  if (kind === 'time') {
    return { kind: 'clock' as const, value: 'HH:mm' as const };
  }
  return { kind: 'none' as const };
}

function defaultClockFormat(
  kind: 'date' | 'time',
): WorkDocumentClockFieldFormat {
  return kind === 'date' ? DOCUMENT_FIELD_DATE_FORMAT_ZH : 'HH:mm';
}

function clockFormatOptions(
  messages: OfficeMessageCatalog,
  kind: 'date' | 'time',
  draft: WorkDocumentFieldDraft,
): readonly OfficeSelectOption<
  WorkDocumentClockFieldFormat | '__preserved__'
>[] {
  const options =
    kind === 'date' ? dateFormatOptions(messages) : timeFormatOptions();
  return draft.format.kind === 'clock' && draft.format.source
    ? [
        {
          value: '__preserved__' as const,
          label: officeMessage(messages, 'document.field.preserveFormat', {
            source: draft.format.source,
          }),
        },
        ...options,
      ]
    : options;
}
