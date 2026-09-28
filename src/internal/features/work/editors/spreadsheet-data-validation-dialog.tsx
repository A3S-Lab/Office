import {
  Info,
  ListChecks,
  ShieldCheck,
  Sigma,
  TriangleAlert,
} from 'lucide-react';
import { type FormEvent, useId, useMemo, useState } from 'react';
import { officeMessage } from '../../../i18n/office-locale';
import type { OfficeMessageCatalog } from '../../../i18n/office-messages';
import { Button, Dialog, Field } from '../../../design-system/primitives';
import { OfficeCheckbox, OfficeSelect } from './office-controls';
import { useOfficeMessages } from './office-messages-context';
import {
  SPREADSHEET_DATA_VALIDATION_ERROR_LIMIT,
  SPREADSHEET_DATA_VALIDATION_FORMULA_LIMIT,
  SPREADSHEET_DATA_VALIDATION_HINT_LIMIT,
  SPREADSHEET_DATA_VALIDATION_TITLE_LIMIT,
} from '../work-spreadsheet-data-validation';
import {
  spreadsheetDataValidationOperators,
  type SpreadsheetDataValidationDialogSource,
  type SpreadsheetDataValidationDialogValue,
  type SpreadsheetDataValidationOperator,
  type SpreadsheetDataValidationType,
} from './spreadsheet-data-validation';
import { isSpreadsheetDependentListFormula } from './spreadsheet-data-validation-list';

function validationTypeOptions(messages: OfficeMessageCatalog): readonly {
  label: string;
  value: SpreadsheetDataValidationType;
}[] {
  return [
    {
      value: 'custom',
      label: officeMessage(messages, 'spreadsheet.dv.type.custom'),
    },
    {
      value: 'dropdown',
      label: officeMessage(messages, 'spreadsheet.dv.type.dropdown'),
    },
    {
      value: 'number_integer',
      label: officeMessage(messages, 'spreadsheet.dv.type.integer'),
    },
    {
      value: 'number',
      label: officeMessage(messages, 'spreadsheet.dv.type.decimal'),
    },
    {
      value: 'date',
      label: officeMessage(messages, 'spreadsheet.dv.type.date'),
    },
    {
      value: 'text_length',
      label: officeMessage(messages, 'spreadsheet.dv.type.textLength'),
    },
  ];
}

function operatorLabel(
  messages: OfficeMessageCatalog,
  operator: SpreadsheetDataValidationOperator,
): string {
  switch (operator) {
    case 'between':
      return officeMessage(messages, 'spreadsheet.dv.op.between');
    case 'notBetween':
      return officeMessage(messages, 'spreadsheet.dv.op.notBetween');
    case 'equal':
      return officeMessage(messages, 'spreadsheet.dv.op.equal');
    case 'notEqualTo':
      return officeMessage(messages, 'spreadsheet.dv.op.notEqualTo');
    case 'moreThanThe':
      return officeMessage(messages, 'spreadsheet.dv.op.moreThanThe');
    case 'lessThan':
      return officeMessage(messages, 'spreadsheet.dv.op.lessThan');
    case 'greaterOrEqualTo':
      return officeMessage(messages, 'spreadsheet.dv.op.greaterOrEqualTo');
    case 'lessThanOrEqualTo':
      return officeMessage(messages, 'spreadsheet.dv.op.lessThanOrEqualTo');
    case 'earlierThan':
      return officeMessage(messages, 'spreadsheet.dv.op.earlierThan');
    case 'noEarlierThan':
      return officeMessage(messages, 'spreadsheet.dv.op.noEarlierThan');
    case 'laterThan':
      return officeMessage(messages, 'spreadsheet.dv.op.laterThan');
    case 'noLaterThan':
      return officeMessage(messages, 'spreadsheet.dv.op.noLaterThan');
  }
}

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
  const messages = useOfficeMessages();
  const [value, setValue] = useState(source.value);
  const [touched, setTouched] = useState(false);
  const formId = useId();
  const operators = spreadsheetDataValidationOperators(value.type);
  const validationError = onValidate(value);
  const valueDirty = !sameSpreadsheetDataValidationValue(value, source.value);
  const dirty = source.mixed || valueDirty;
  const visibleError = touched ? validationError : null;
  const validationTypes = useMemo(
    () => validationTypeOptions(messages),
    [messages],
  );
  const allowLabel = officeMessage(messages, 'spreadsheet.dv.allow');
  const dataLabel = officeMessage(messages, 'spreadsheet.dv.data');
  const allowBlankLabel = officeMessage(messages, 'spreadsheet.dv.allowBlank');
  const showDropdownLabel = officeMessage(
    messages,
    'spreadsheet.dv.showDropdown',
  );
  const sourceLabel = officeMessage(messages, 'spreadsheet.dv.source');
  const inputTitleLabel = officeMessage(messages, 'spreadsheet.dv.inputTitle');
  const inputMessageLabel = officeMessage(
    messages,
    'spreadsheet.dv.inputMessage',
  );
  const errorStyleLabel = officeMessage(messages, 'spreadsheet.dv.errorStyle');
  const errorTitleLabel = officeMessage(messages, 'spreadsheet.dv.errorTitle');
  const errorMessageLabel = officeMessage(
    messages,
    'spreadsheet.dv.errorMessage',
  );

  const resetDraft = () => {
    setValue(source.value);
    setTouched(false);
  };

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
      title={officeMessage(messages, 'spreadsheet.dv.title')}
      description={`${source.sheetName}!${source.rangeReference}`}
      className="work-spreadsheet-data-validation-dialog"
      restoreFocusTarget={restoreFocusTarget}
      onClose={onClose}
      onEscape={() => {
        if (valueDirty) {
          resetDraft();
          return;
        }
        onClose();
      }}
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
              {officeMessage(messages, 'spreadsheet.dv.clearAll')}
            </Button>
          )}
          <Button tone="quiet" onClick={onClose}>
            {officeMessage(messages, 'spreadsheet.dv.cancel')}
          </Button>
          <Button
            tone="primary"
            type="submit"
            form={formId}
            disabled={!dirty || Boolean(validationError)}
          >
            {officeMessage(messages, 'spreadsheet.dv.ok')}
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
                ? officeMessage(messages, 'spreadsheet.dv.scope.contiguous')
                : officeMessage(messages, 'spreadsheet.dv.scope.multi', {
                    n: String(source.ranges.length),
                  })}
            </small>
          </div>
        </div>

        {source.mixed ? (
          <p className="work-spreadsheet-data-validation-mixed" role="status">
            {officeMessage(messages, 'spreadsheet.dv.mixedRules')}
          </p>
        ) : null}

        <section aria-labelledby={`${formId}-condition`}>
          <div className="work-spreadsheet-data-validation-section-heading">
            <ShieldCheck size={16} aria-hidden="true" />
            <h3 id={`${formId}-condition`}>
              {officeMessage(messages, 'spreadsheet.dv.section.condition')}
            </h3>
          </div>
          <div className="work-spreadsheet-data-validation-condition-grid">
            <Field label={allowLabel}>
              <OfficeSelect
                ariaLabel={allowLabel}
                value={value.type}
                options={validationTypes}
                onValueChange={(type) => changeType(type)}
              />
            </Field>

            {value.type !== 'dropdown' && value.type !== 'custom' ? (
              <Field label={dataLabel}>
                <OfficeSelect
                  ariaLabel={dataLabel}
                  value={value.type2}
                  options={operators.map((operator) => ({
                    value: operator,
                    label: operatorLabel(messages, operator),
                  }))}
                  onValueChange={(type2) => {
                    update({
                      type2,
                      ...(!spreadsheetDataValidationNeedsSecondValue(type2)
                        ? { value2: '' }
                        : {}),
                    });
                    setTouched(true);
                  }}
                />
              </Field>
            ) : null}
          </div>

          <div className="work-spreadsheet-data-validation-behavior">
            <OfficeCheckbox
              ariaLabel={allowBlankLabel}
              checked={value.allowBlank}
              onCheckedChange={(allowBlank) => update({ allowBlank })}
            >
              {allowBlankLabel}
            </OfficeCheckbox>
            {value.type === 'dropdown' ? (
              <OfficeCheckbox
                ariaLabel={showDropdownLabel}
                checked={value.showDropdownArrow}
                onCheckedChange={(showDropdownArrow) =>
                  update({ showDropdownArrow })
                }
              >
                {showDropdownLabel}
              </OfficeCheckbox>
            ) : null}
          </div>

          {value.type === 'dropdown' ? (
            <div className="work-spreadsheet-data-validation-list-source">
              <Field
                label={sourceLabel}
                required
                description={
                  isSpreadsheetDependentListFormula(value.value1)
                    ? officeMessage(
                        messages,
                        'spreadsheet.dv.source.desc.dependent',
                      )
                    : officeMessage(
                        messages,
                        'spreadsheet.dv.source.desc.static',
                      )
                }
                error={visibleError ?? undefined}
              >
                <input
                  type="text"
                  aria-label={sourceLabel}
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
              {isSpreadsheetDependentListFormula(value.value1) ? (
                <p className="work-spreadsheet-data-validation-formula-note">
                  <Sigma size={14} aria-hidden="true" />
                  <span>
                    {officeMessage(
                      messages,
                      'spreadsheet.dv.source.note.dependent',
                    )}
                  </span>
                </p>
              ) : null}
            </div>
          ) : value.type === 'custom' ? (
            <SpreadsheetDataValidationCustomFormulaField
              messages={messages}
              value={value.value1}
              error={visibleError}
              onTouched={() => setTouched(true)}
              onValueChange={(value1) => update({ value1 })}
            />
          ) : (
            <SpreadsheetDataValidationBoundaryFields
              messages={messages}
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

        <section aria-labelledby={`${formId}-input-message`}>
          <div className="work-spreadsheet-data-validation-section-heading">
            <Info size={16} aria-hidden="true" />
            <h3 id={`${formId}-input-message`}>
              {officeMessage(messages, 'spreadsheet.dv.section.inputMessage')}
            </h3>
          </div>
          <OfficeCheckbox
            ariaLabel={officeMessage(
              messages,
              'spreadsheet.dv.showInputMessage',
            )}
            checked={value.hintShow}
            onCheckedChange={(hintShow) => update({ hintShow })}
          >
            {officeMessage(messages, 'spreadsheet.dv.showInputMessage')}
          </OfficeCheckbox>
          {value.hintShow ? (
            <div className="work-spreadsheet-data-validation-message-grid">
              <Field label={inputTitleLabel}>
                <input
                  type="text"
                  aria-label={inputTitleLabel}
                  maxLength={SPREADSHEET_DATA_VALIDATION_TITLE_LIMIT}
                  value={value.hintTitle}
                  onChange={(event) =>
                    update({ hintTitle: event.currentTarget.value })
                  }
                />
              </Field>
              <Field
                label={inputMessageLabel}
                className="message"
                description={officeMessage(
                  messages,
                  'spreadsheet.dv.inputMessage.desc',
                  { n: String(SPREADSHEET_DATA_VALIDATION_HINT_LIMIT) },
                )}
              >
                <textarea
                  aria-label={inputMessageLabel}
                  rows={3}
                  maxLength={SPREADSHEET_DATA_VALIDATION_HINT_LIMIT}
                  value={value.hintValue}
                  onChange={(event) =>
                    update({ hintValue: event.currentTarget.value })
                  }
                />
              </Field>
            </div>
          ) : null}
        </section>

        <section aria-labelledby={`${formId}-error-alert`}>
          <div className="work-spreadsheet-data-validation-section-heading">
            <TriangleAlert size={16} aria-hidden="true" />
            <h3 id={`${formId}-error-alert`}>
              {officeMessage(messages, 'spreadsheet.dv.section.errorAlert')}
            </h3>
          </div>
          <OfficeCheckbox
            ariaLabel={officeMessage(messages, 'spreadsheet.dv.showErrorAlert')}
            checked={value.prohibitInput}
            onCheckedChange={(prohibitInput) => update({ prohibitInput })}
          >
            {officeMessage(messages, 'spreadsheet.dv.showErrorAlert')}
          </OfficeCheckbox>
          {value.prohibitInput ? (
            <div className="work-spreadsheet-data-validation-message-grid">
              <Field
                label={errorStyleLabel}
                description={officeMessage(
                  messages,
                  'spreadsheet.dv.errorStyle.desc',
                )}
              >
                <OfficeSelect
                  ariaLabel={errorStyleLabel}
                  value={value.errorStyle}
                  options={[
                    {
                      value: 'stop',
                      label: officeMessage(
                        messages,
                        'spreadsheet.dv.errorStyle.stop',
                      ),
                    },
                    {
                      value: 'warning',
                      label: officeMessage(
                        messages,
                        'spreadsheet.dv.errorStyle.warning',
                      ),
                    },
                    {
                      value: 'information',
                      label: officeMessage(
                        messages,
                        'spreadsheet.dv.errorStyle.information',
                      ),
                    },
                  ]}
                  onValueChange={(errorStyle) => update({ errorStyle })}
                />
              </Field>
              <Field label={errorTitleLabel}>
                <input
                  type="text"
                  aria-label={errorTitleLabel}
                  maxLength={SPREADSHEET_DATA_VALIDATION_TITLE_LIMIT}
                  value={value.errorTitle}
                  onChange={(event) =>
                    update({ errorTitle: event.currentTarget.value })
                  }
                />
              </Field>
              <Field
                label={errorMessageLabel}
                className="message"
                description={officeMessage(
                  messages,
                  'spreadsheet.dv.errorMessage.desc',
                  { n: String(SPREADSHEET_DATA_VALIDATION_ERROR_LIMIT) },
                )}
              >
                <textarea
                  aria-label={errorMessageLabel}
                  rows={3}
                  maxLength={SPREADSHEET_DATA_VALIDATION_ERROR_LIMIT}
                  value={value.errorMessage}
                  onChange={(event) =>
                    update({ errorMessage: event.currentTarget.value })
                  }
                />
              </Field>
            </div>
          ) : null}
        </section>

        {visibleError && value.type !== 'dropdown' ? (
          <p className="work-spreadsheet-data-validation-error" role="alert">
            {visibleError}
          </p>
        ) : null}
      </form>
    </Dialog>
  );
}

function SpreadsheetDataValidationCustomFormulaField({
  error,
  messages,
  onTouched,
  onValueChange,
  value,
}: {
  error: string | null;
  messages: OfficeMessageCatalog;
  onTouched: () => void;
  onValueChange: (value: string) => void;
  value: string;
}) {
  const formulaLabel = officeMessage(messages, 'spreadsheet.dv.formula');
  return (
    <div className="work-spreadsheet-data-validation-custom-formula">
      <Field
        label={formulaLabel}
        required
        description={officeMessage(messages, 'spreadsheet.dv.formula.desc')}
        error={error ?? undefined}
      >
        <textarea
          aria-label={formulaLabel}
          rows={2}
          maxLength={SPREADSHEET_DATA_VALIDATION_FORMULA_LIMIT}
          placeholder={officeMessage(
            messages,
            'spreadsheet.dv.formula.placeholder',
          )}
          autoCapitalize="none"
          spellCheck={false}
          value={value}
          onBlur={onTouched}
          onChange={(event) => {
            onValueChange(event.currentTarget.value);
            onTouched();
          }}
        />
      </Field>
      <p className="work-spreadsheet-data-validation-formula-note">
        <Sigma size={14} aria-hidden="true" />
        <span>{officeMessage(messages, 'spreadsheet.dv.formula.note')}</span>
      </p>
    </div>
  );
}

function SpreadsheetDataValidationBoundaryFields({
  error,
  messages,
  onTouched,
  onValueChange,
  type,
  type2,
  value1,
  value2,
}: {
  error: string | null;
  messages: OfficeMessageCatalog;
  onTouched: () => void;
  onValueChange: (value: Partial<SpreadsheetDataValidationDialogValue>) => void;
  type: Exclude<SpreadsheetDataValidationType, 'custom' | 'dropdown'>;
  type2: SpreadsheetDataValidationDialogValue['type2'];
  value1: string;
  value2: string;
}) {
  const needsSecond = spreadsheetDataValidationNeedsSecondValue(type2);
  const inputMode = type === 'date' ? 'text' : 'decimal';
  const firstLabel = needsSecond
    ? type === 'date'
      ? officeMessage(messages, 'spreadsheet.dv.bound.startDate')
      : officeMessage(messages, 'spreadsheet.dv.bound.min')
    : type === 'date'
      ? officeMessage(messages, 'spreadsheet.dv.bound.date')
      : officeMessage(messages, 'spreadsheet.dv.bound.value');
  const secondLabel =
    type === 'date'
      ? officeMessage(messages, 'spreadsheet.dv.bound.endDate')
      : officeMessage(messages, 'spreadsheet.dv.bound.max');
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
            ? officeMessage(messages, 'spreadsheet.dv.bound.dateDesc')
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
      {needsSecond ? (
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
      ) : null}
      {error ? (
        <span className="work-spreadsheet-data-validation-inline-error">
          {error}
        </span>
      ) : null}
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
    left.allowBlank === right.allowBlank &&
    left.showDropdownArrow === right.showDropdownArrow &&
    left.prohibitInput === right.prohibitInput &&
    left.errorStyle === right.errorStyle &&
    left.errorTitle === right.errorTitle &&
    left.errorMessage === right.errorMessage &&
    left.hintShow === right.hintShow &&
    left.hintTitle === right.hintTitle &&
    left.hintValue === right.hintValue
  );
}
