import { Button, Dialog } from '../../../design-system/primitives';
import { officeMessage } from '../../../i18n/office-locale';
import type { OfficeMessageCatalog } from '../../../i18n/office-messages';
import type {
  WorkDocumentMailMergeFilterOperator,
  WorkDocumentMailMergeFilterRule,
  WorkDocumentMailMergeRecipientFilter,
  WorkDocumentMailMergeSource,
} from '../work-document-mail-merge';
import {
  DOCUMENT_MAIL_MERGE_MAX_FILTER_RULES,
  filteredMailMergeRecords,
  mailMergeFieldNamesFromRecords,
  normalizeMailMergeRecipientFilter,
  setMailMergeRecipientFilter,
} from '../work-document-mail-merge';
import {
  OfficeSelect,
  type OfficeSelectOption,
  OfficeTextField,
} from './office-controls';
import { useOfficeMessages } from './office-messages-context';

export interface DocumentMailMergeRecipientFilterDialogProps {
  source: WorkDocumentMailMergeSource;
  fieldNames?: readonly string[];
  restoreFocusTarget: () => HTMLElement | null;
  onCancel: () => void;
  onChange: (source: WorkDocumentMailMergeSource) => void;
  onSubmit: () => void;
}

export function DocumentMailMergeRecipientFilterDialog({
  source,
  fieldNames,
  restoreFocusTarget,
  onCancel,
  onChange,
  onSubmit,
}: DocumentMailMergeRecipientFilterDialogProps) {
  const messages = useOfficeMessages();
  const columns =
    fieldNames && fieldNames.length
      ? [...fieldNames]
      : mailMergeFieldNamesFromRecords(source.records);
  const filter = normalizeMailMergeRecipientFilter(source.filter) ?? {
    rules: [] as WorkDocumentMailMergeFilterRule[],
  };
  const visibleCount = filteredMailMergeRecords(source).length;
  const fieldOptions: OfficeSelectOption[] = columns.map((name) => ({
    value: name,
    label: name,
  }));
  const canAdd =
    columns.length > 0 &&
    filter.rules.length < DOCUMENT_MAIL_MERGE_MAX_FILTER_RULES;

  const updateFilter = (next: WorkDocumentMailMergeRecipientFilter | null) => {
    onChange(setMailMergeRecipientFilter(source, next));
  };

  const updateRule = (
    index: number,
    patch: Partial<WorkDocumentMailMergeFilterRule>,
  ) => {
    const rules = filter.rules.map((rule, ruleIndex) => {
      if (ruleIndex !== index) return rule;
      const operator = patch.operator ?? rule.operator;
      const next: WorkDocumentMailMergeFilterRule = {
        field: patch.field ?? rule.field,
        operator,
        ...(operator === 'isBlank' || operator === 'isNotBlank'
          ? {}
          : { value: patch.value ?? rule.value ?? '' }),
      };
      return next;
    });
    updateFilter(rules.length ? { rules } : null);
  };

  const removeRule = (index: number) => {
    const rules = filter.rules.filter((_, ruleIndex) => ruleIndex !== index);
    updateFilter(rules.length ? { rules } : null);
  };

  const addRule = () => {
    if (!canAdd) return;
    const field = columns[0]!;
    updateFilter({
      rules: [...filter.rules, { field, operator: 'equals', value: '' }],
    });
  };

  return (
    <Dialog
      title={officeMessage(messages, 'document.mailMerge.filter.title')}
      description={officeMessage(
        messages,
        'document.mailMerge.filter.description',
      )}
      className="work-document-mail-merge-filter-dialog"
      focusKey="document-mail-merge-filter"
      restoreFocusTarget={restoreFocusTarget}
      onClose={onCancel}
      footer={
        <>
          <Button tone="quiet" onClick={onCancel}>
            {officeMessage(messages, 'document.mailMerge.filter.cancel')}
          </Button>
          <Button
            tone="quiet"
            onClick={() => updateFilter(null)}
            disabled={!filter.rules.length}
          >
            {officeMessage(messages, 'document.mailMerge.filter.clear')}
          </Button>
          <Button tone="primary" onClick={onSubmit}>
            {officeMessage(messages, 'document.mailMerge.filter.confirm')}
          </Button>
        </>
      }
    >
      <p className="work-document-mail-merge-filter-summary">
        {officeMessage(messages, 'document.mailMerge.filter.summary', {
          visible: String(visibleCount),
          total: String(source.records.length),
          max: String(DOCUMENT_MAIL_MERGE_MAX_FILTER_RULES),
        })}
      </p>
      {columns.length === 0 ? (
        <p className="work-document-mail-merge-filter-empty">
          {officeMessage(messages, 'document.mailMerge.filter.noFields')}
        </p>
      ) : (
        <div className="work-document-mail-merge-filter-rules">
          {filter.rules.map((rule, index) => {
            const needsValue =
              rule.operator !== 'isBlank' && rule.operator !== 'isNotBlank';
            const n = String(index + 1);
            return (
              <div
                key={`mail-merge-filter-rule-${index}`}
                className="work-document-mail-merge-filter-rule"
              >
                <div className="work-document-mail-merge-filter-field">
                  <span>
                    {officeMessage(messages, 'document.mailMerge.filter.field')}
                  </span>
                  <OfficeSelect
                    id={`mail-merge-filter-field-${index}`}
                    ariaLabel={officeMessage(
                      messages,
                      'document.mailMerge.filter.fieldAria',
                      { n },
                    )}
                    value={rule.field}
                    options={fieldOptions}
                    onValueChange={(field) => updateRule(index, { field })}
                  />
                </div>
                <div className="work-document-mail-merge-filter-field">
                  <span>
                    {officeMessage(
                      messages,
                      'document.mailMerge.filter.condition',
                    )}
                  </span>
                  <OfficeSelect<WorkDocumentMailMergeFilterOperator>
                    id={`mail-merge-filter-operator-${index}`}
                    ariaLabel={officeMessage(
                      messages,
                      'document.mailMerge.filter.conditionAria',
                      { n },
                    )}
                    value={rule.operator}
                    options={operatorOptions(messages)}
                    onValueChange={(operator) =>
                      updateRule(index, { operator })
                    }
                  />
                </div>
                {needsValue ? (
                  <label
                    className="work-document-mail-merge-filter-field"
                    htmlFor={`mail-merge-filter-value-${index}`}
                  >
                    <span>
                      {officeMessage(
                        messages,
                        'document.mailMerge.filter.value',
                      )}
                    </span>
                    <OfficeTextField
                      id={`mail-merge-filter-value-${index}`}
                      aria-label={officeMessage(
                        messages,
                        'document.mailMerge.filter.valueAria',
                        { n },
                      )}
                      value={rule.value ?? ''}
                      onChange={(event) =>
                        updateRule(index, { value: event.target.value })
                      }
                    />
                  </label>
                ) : null}
                <Button tone="quiet" onClick={() => removeRule(index)}>
                  {officeMessage(messages, 'document.mailMerge.filter.delete')}
                </Button>
              </div>
            );
          })}
          <Button tone="quiet" onClick={addRule} disabled={!canAdd}>
            {officeMessage(messages, 'document.mailMerge.filter.add')}
          </Button>
        </div>
      )}
    </Dialog>
  );
}

function operatorOptions(messages: OfficeMessageCatalog) {
  return [
    {
      value: 'equals' as const,
      label: officeMessage(messages, 'document.mailMerge.filter.equals'),
    },
    {
      value: 'notEquals' as const,
      label: officeMessage(messages, 'document.mailMerge.filter.notEquals'),
    },
    {
      value: 'contains' as const,
      label: officeMessage(messages, 'document.mailMerge.filter.contains'),
    },
    {
      value: 'notContains' as const,
      label: officeMessage(messages, 'document.mailMerge.filter.notContains'),
    },
    {
      value: 'startsWith' as const,
      label: officeMessage(messages, 'document.mailMerge.filter.startsWith'),
    },
    {
      value: 'endsWith' as const,
      label: officeMessage(messages, 'document.mailMerge.filter.endsWith'),
    },
    {
      value: 'isBlank' as const,
      label: officeMessage(messages, 'document.mailMerge.filter.isBlank'),
    },
    {
      value: 'isNotBlank' as const,
      label: officeMessage(messages, 'document.mailMerge.filter.isNotBlank'),
    },
  ];
}
