import { Button, Dialog } from '../../../design-system/primitives';
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
  OfficeTextField,
  type OfficeSelectOption,
} from './office-controls';

export interface DocumentMailMergeRecipientFilterDialogProps {
  source: WorkDocumentMailMergeSource;
  fieldNames?: readonly string[];
  restoreFocusTarget: () => HTMLElement | null;
  onCancel: () => void;
  onChange: (source: WorkDocumentMailMergeSource) => void;
  onSubmit: () => void;
}

const OPERATOR_OPTIONS = [
  { value: 'equals', label: '等于' },
  { value: 'notEquals', label: '不等于' },
  { value: 'contains', label: '包含' },
  { value: 'notContains', label: '不包含' },
  { value: 'startsWith', label: '开头是' },
  { value: 'endsWith', label: '结尾是' },
  { value: 'isBlank', label: '为空' },
  { value: 'isNotBlank', label: '不为空' },
] as const satisfies readonly {
  value: WorkDocumentMailMergeFilterOperator;
  label: string;
}[];

export function DocumentMailMergeRecipientFilterDialog({
  source,
  fieldNames,
  restoreFocusTarget,
  onCancel,
  onChange,
  onSubmit,
}: DocumentMailMergeRecipientFilterDialogProps) {
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
      title="筛选收件人"
      description="按字段条件筛选邮件合并收件人。多条规则为并且关系。"
      className="work-document-mail-merge-filter-dialog"
      focusKey="document-mail-merge-filter"
      restoreFocusTarget={restoreFocusTarget}
      onClose={onCancel}
      footer={
        <>
          <Button tone="quiet" onClick={onCancel}>
            取消
          </Button>
          <Button
            tone="quiet"
            onClick={() => updateFilter(null)}
            disabled={!filter.rules.length}
          >
            清除筛选
          </Button>
          <Button tone="primary" onClick={onSubmit}>
            确定
          </Button>
        </>
      }
    >
      <p className="work-document-mail-merge-filter-summary">
        当前可见 {visibleCount} / {source.records.length} 位收件人（最多{' '}
        {DOCUMENT_MAIL_MERGE_MAX_FILTER_RULES} 条规则）。
      </p>
      {columns.length === 0 ? (
        <p className="work-document-mail-merge-filter-empty">
          当前数据源没有可用字段。
        </p>
      ) : (
        <div className="work-document-mail-merge-filter-rules">
          {filter.rules.map((rule, index) => {
            const needsValue =
              rule.operator !== 'isBlank' && rule.operator !== 'isNotBlank';
            return (
              <div
                key={`mail-merge-filter-rule-${index}`}
                className="work-document-mail-merge-filter-rule"
              >
                <div className="work-document-mail-merge-filter-field">
                  <span>字段</span>
                  <OfficeSelect
                    id={`mail-merge-filter-field-${index}`}
                    ariaLabel={`筛选字段 ${index + 1}`}
                    value={rule.field}
                    options={fieldOptions}
                    onValueChange={(field) => updateRule(index, { field })}
                  />
                </div>
                <div className="work-document-mail-merge-filter-field">
                  <span>条件</span>
                  <OfficeSelect<WorkDocumentMailMergeFilterOperator>
                    id={`mail-merge-filter-operator-${index}`}
                    ariaLabel={`筛选条件 ${index + 1}`}
                    value={rule.operator}
                    options={[...OPERATOR_OPTIONS]}
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
                    <span>值</span>
                    <OfficeTextField
                      id={`mail-merge-filter-value-${index}`}
                      aria-label={`筛选值 ${index + 1}`}
                      value={rule.value ?? ''}
                      onChange={(event) =>
                        updateRule(index, { value: event.target.value })
                      }
                    />
                  </label>
                ) : null}
                <Button tone="quiet" onClick={() => removeRule(index)}>
                  删除
                </Button>
              </div>
            );
          })}
          <Button tone="quiet" onClick={addRule} disabled={!canAdd}>
            添加条件
          </Button>
        </div>
      )}
    </Dialog>
  );
}
