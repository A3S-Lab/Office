import { ListChecks, ShieldCheck } from 'lucide-react';
import { type FormEvent, useId, useState } from 'react';
import { Button, Dialog, Field } from '../../../design-system/primitives';
import { OfficeCheckbox } from './office-controls';
import {
  spreadsheetDataValidationOperators,
  type SpreadsheetDataValidationDialogSource,
  type SpreadsheetDataValidationDialogValue,
  type SpreadsheetDataValidationOperator,
  type SpreadsheetDataValidationType,
} from './spreadsheet-data-validation';

const validationTypes: readonly {
  label: string;
  value: SpreadsheetDataValidationType;
}[] = [
  { value: 'dropdown', label: '序列' },
  { value: 'number_integer', label: '整数' },
  { value: 'number', label: '小数' },
  { value: 'date', label: '日期' },
  { value: 'text_length', label: '文本长度' },
];

const operatorLabels: Record<SpreadsheetDataValidationOperator, string> = {
  between: '介于',
  notBetween: '未介于',
  equal: '等于',
  notEqualTo: '不等于',
  moreThanThe: '大于',
  lessThan: '小于',
  greaterOrEqualTo: '大于或等于',
  lessThanOrEqualTo: '小于或等于',
  earlierThan: '早于',
  noEarlierThan: '不早于',
  laterThan: '晚于',
  noLaterThan: '不晚于',
};

export function SpreadsheetDataValidationDialog({
  source,
  restoreFocusTarget,
  onApply,
  onClose,
  onRemove,
  onValidate,
}: {
  source: SpreadsheetDataValidationDialogSource;
  restoreFocusTarget: () => HTMLElement | null;
  onApply: (value: SpreadsheetDataValidationDialogValue) => boolean;
  onClose: () => void;
  onRemove: () => boolean;
  onValidate: (value: SpreadsheetDataValidationDialogValue) => string | null;
}) {
  const [value, setValue] = useState(source.value);
  const [touched, setTouched] = useState(false);
  const formId = useId();
  const operators = spreadsheetDataValidationOperators(value.type);
  const validationError = onValidate(value);
  const dirty =
    source.mixed || !sameSpreadsheetDataValidationValue(value, source.value);
  const visibleError = touched ? validationError : null;

  const update = (patch: Partial<SpreadsheetDataValidationDialogValue>) => {
    setValue((current) => ({ ...current, ...patch }));
  };
  const changeType = (type: SpreadsheetDataValidationType) => {
    const nextOperators = spreadsheetDataValidationOperators(type);
    setValue((current) => ({
      ...current,
      type,
      type2: nextOperators[0] ?? '',
      value1: '',
      value2: '',
    }));
    setTouched(false);
  };
  const submit = (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    setTouched(true);
    if (!dirty || validationError) return;
    if (onApply(value)) onClose();
  };

  return (
    <Dialog
      title="数据验证"
      description={`${source.sheetName}!${source.rangeReference}`}
      className="work-spreadsheet-data-validation-dialog"
      restoreFocusTarget={restoreFocusTarget}
      onClose={onClose}
      footer={
        <>
          {source.hasValidation && (
            <Button
              tone="danger"
              className="work-spreadsheet-data-validation-remove"
              onClick={() => {
                if (onRemove()) onClose();
              }}
            >
              全部清除
            </Button>
          )}
          <Button tone="quiet" onClick={onClose}>
            取消
          </Button>
          <Button
            tone="primary"
            type="submit"
            form={formId}
            disabled={!dirty || Boolean(validationError)}
          >
            确定
          </Button>
        </>
      }
    >
      <form id={formId} onSubmit={submit}>
        <div className="work-spreadsheet-data-validation-scope">
          <span aria-hidden="true">
            <ListChecks size={17} />
          </span>
          <div>
            <strong>{source.rangeReference}</strong>
            <small>
              {source.ranges.length === 1
                ? '当前连续区域'
                : `${source.ranges.length} 个选定区域`}
            </small>
          </div>
        </div>

        {source.mixed && (
          <p className="work-spreadsheet-data-validation-mixed" role="status">
            所选区域包含不同的数据验证规则。确定后将统一为当前设置。
          </p>
        )}

        <section aria-labelledby={`${formId}-condition`}>
          <div className="work-spreadsheet-data-validation-section-heading">
            <ShieldCheck size={16} aria-hidden="true" />
            <h3 id={`${formId}-condition`}>验证条件</h3>
          </div>
          <div className="work-spreadsheet-data-validation-condition-grid">
            <Field label="允许">
              <select
                value={value.type}
                onChange={(event) =>
                  changeType(
                    event.currentTarget.value as SpreadsheetDataValidationType,
                  )
                }
              >
                {validationTypes.map((option) => (
                  <option key={option.value} value={option.value}>
                    {option.label}
                  </option>
                ))}
              </select>
            </Field>

            {value.type !== 'dropdown' && (
              <Field label="数据">
                <select
                  value={value.type2}
                  onChange={(event) => {
                    const type2 = event.currentTarget
                      .value as SpreadsheetDataValidationOperator;
                    update({
                      type2,
                      ...(!spreadsheetDataValidationNeedsSecondValue(type2)
                        ? { value2: '' }
                        : {}),
                    });
                    setTouched(true);
                  }}
                >
                  {operators.map((operator) => (
                    <option key={operator} value={operator}>
                      {operatorLabels[operator]}
                    </option>
                  ))}
                </select>
              </Field>
            )}
          </div>

          {value.type === 'dropdown' ? (
            <Field
              label="来源"
              required
              description="输入逗号分隔项目，或单行/单列区域，例如 Ready,Blocked 或 Lists!A1:A8。"
              error={visibleError ?? undefined}
            >
              <input
                type="text"
                autoCapitalize="none"
                spellCheck={false}
                value={value.value1}
                onBlur={() => setTouched(true)}
                onChange={(event) => {
                  update({ value1: event.currentTarget.value });
                  setTouched(true);
                }}
              />
            </Field>
          ) : (
            <SpreadsheetDataValidationBoundaryFields
              type={value.type}
              type2={value.type2}
              value1={value.value1}
              value2={value.value2}
              error={visibleError}
              onTouched={() => setTouched(true)}
              onValueChange={(patch) => update(patch)}
            />
          )}
        </section>

        <section aria-labelledby={`${formId}-behavior`}>
          <h3 id={`${formId}-behavior`}>输入行为</h3>
          <div className="work-spreadsheet-data-validation-behavior">
            <OfficeCheckbox
              ariaLabel="输入无效数据时阻止提交"
              checked={value.prohibitInput}
              onCheckedChange={(prohibitInput) => update({ prohibitInput })}
            >
              输入无效数据时阻止提交
            </OfficeCheckbox>
            <OfficeCheckbox
              ariaLabel="选中单元格时显示输入信息"
              checked={value.hintShow}
              onCheckedChange={(hintShow) => update({ hintShow })}
            >
              选中单元格时显示输入信息
            </OfficeCheckbox>
          </div>
          {value.hintShow && (
            <Field
              label="输入信息"
              description="最多 255 个字符，将在用户选中受验证单元格时显示。"
            >
              <textarea
                rows={3}
                maxLength={255}
                value={value.hintValue}
                onChange={(event) =>
                  update({ hintValue: event.currentTarget.value })
                }
              />
            </Field>
          )}
        </section>

        {visibleError && value.type !== 'dropdown' && (
          <p className="work-spreadsheet-data-validation-error" role="alert">
            {visibleError}
          </p>
        )}
      </form>
    </Dialog>
  );
}

function SpreadsheetDataValidationBoundaryFields({
  error,
  onTouched,
  onValueChange,
  type,
  type2,
  value1,
  value2,
}: {
  error: string | null;
  onTouched: () => void;
  onValueChange: (value: Partial<SpreadsheetDataValidationDialogValue>) => void;
  type: Exclude<SpreadsheetDataValidationType, 'dropdown'>;
  type2: SpreadsheetDataValidationDialogValue['type2'];
  value1: string;
  value2: string;
}) {
  const needsSecond = spreadsheetDataValidationNeedsSecondValue(type2);
  const inputMode = type === 'date' ? 'text' : 'decimal';
  const firstLabel = needsSecond
    ? type === 'date'
      ? '开始日期'
      : '最小值'
    : type === 'date'
      ? '日期'
      : '值';
  const secondLabel = type === 'date' ? '结束日期' : '最大值';
  return (
    <div
      className={
        needsSecond
          ? 'work-spreadsheet-data-validation-values two'
          : 'work-spreadsheet-data-validation-values'
      }
    >
      <Field
        label={firstLabel}
        required
        description={
          type === 'date'
            ? '使用 YYYY-MM-DD、Excel 日期序号或 DATE(...)。'
            : undefined
        }
      >
        <input
          type="text"
          inputMode={inputMode}
          value={value1}
          onBlur={onTouched}
          onChange={(event) => {
            onValueChange({ value1: event.currentTarget.value });
            onTouched();
          }}
        />
      </Field>
      {needsSecond && (
        <Field label={secondLabel} required>
          <input
            type="text"
            inputMode={inputMode}
            value={value2}
            onBlur={onTouched}
            onChange={(event) => {
              onValueChange({ value2: event.currentTarget.value });
              onTouched();
            }}
          />
        </Field>
      )}
      {error && (
        <span className="work-spreadsheet-data-validation-inline-error">
          {error}
        </span>
      )}
    </div>
  );
}

function spreadsheetDataValidationNeedsSecondValue(
  type2: SpreadsheetDataValidationDialogValue['type2'],
): boolean {
  return type2 === 'between' || type2 === 'notBetween';
}

function sameSpreadsheetDataValidationValue(
  left: SpreadsheetDataValidationDialogValue,
  right: SpreadsheetDataValidationDialogValue,
): boolean {
  return (
    left.type === right.type &&
    left.type2 === right.type2 &&
    left.value1 === right.value1 &&
    left.value2 === right.value2 &&
    left.prohibitInput === right.prohibitInput &&
    left.hintShow === right.hintShow &&
    left.hintValue === right.hintValue
  );
}
