import { Plus, Trash2 } from 'lucide-react';
import { officeMessage } from '../../../i18n/office-locale';
import { CollectionState } from '../../../design-system/primitives';
import {
  type WorkSpreadsheetChartType,
  type WorkSpreadsheetErrorBarDirection,
  type WorkSpreadsheetErrorBars,
  type WorkSpreadsheetErrorBarType,
  type WorkSpreadsheetErrorBarValueType,
  workSpreadsheetChartUsesNumericXAxis,
} from '../work-types';
import { useOfficeMessages } from './office-messages-context';
import {
  CommittedOfficeNumberField,
  CommittedOfficeTextField,
  OfficeCheckbox,
  OfficeSelect,
  OfficeTextField,
} from './office-controls';
import { normalizeRequiredOfficeNumber } from './office-number-normalization';

interface SpreadsheetErrorBarEditorProps {
  chartType: WorkSpreadsheetChartType;
  seriesNumber: number;
  errorBars: WorkSpreadsheetErrorBars[];
  onChange: (errorBars: WorkSpreadsheetErrorBars[]) => void;
  customInput?: 'references' | 'values';
}

export function SpreadsheetErrorBarEditor({
  chartType,
  seriesNumber,
  errorBars,
  onChange,
  customInput = 'references',
}: SpreadsheetErrorBarEditorProps) {
  const messages = useOfficeMessages();
  const replaceErrorBars = (
    index: number,
    change: Partial<WorkSpreadsheetErrorBars>,
  ) => {
    onChange(
      errorBars.map((item, candidate) =>
        candidate === index ? { ...item, ...change } : item,
      ),
    );
  };
  const hasDirection = (direction: WorkSpreadsheetErrorBarDirection) =>
    errorBars.some((item) => item.direction === direction);
  const addErrorBars = (direction: WorkSpreadsheetErrorBarDirection) =>
    onChange([
      ...errorBars,
      { direction, barType: 'both', valueType: 'standardError' },
    ]);

  return (
    <section
      className="work-spreadsheet-error-bars"
      aria-label={officeMessage(messages, 'spreadsheet.chart.errorBar.sectionAria', {
        n: String(seriesNumber),
      })}
    >
      <header>
        <strong>
          {officeMessage(messages, 'spreadsheet.chart.errorBar.title')}
        </strong>
        <div>
          {workSpreadsheetChartUsesNumericXAxis(chartType) ? (
            <button
              type="button"
              aria-label={officeMessage(
                messages,
                'spreadsheet.chart.errorBar.addXAria',
                { n: String(seriesNumber) },
              )}
              disabled={hasDirection('x')}
              onClick={() => addErrorBars('x')}
            >
              <Plus size={11} />X
            </button>
          ) : null}
          <button
            type="button"
            aria-label={officeMessage(
              messages,
              'spreadsheet.chart.errorBar.addYAria',
              { n: String(seriesNumber) },
            )}
            disabled={hasDirection('y')}
            onClick={() => addErrorBars('y')}
          >
            <Plus size={11} />Y
          </button>
        </div>
      </header>
      {!errorBars.length ? (
        <CollectionState
          className="work-spreadsheet-error-bars-empty"
          role="status"
        >
          {officeMessage(messages, 'spreadsheet.chart.errorBar.empty')}
        </CollectionState>
      ) : null}
      {errorBars.map((item, index) => {
        const errorBarNumber = index + 1;
        const labelPrefix = officeMessage(
          messages,
          'spreadsheet.chart.errorBar.itemPrefix',
          { n: String(seriesNumber), index: String(errorBarNumber) },
        );
        return (
          <fieldset key={`${seriesNumber}-${errorBarNumber}`}>
            <legend>
              {officeMessage(messages, 'spreadsheet.chart.errorBar.itemLegend', {
                direction: item.direction.toUpperCase(),
                index: String(errorBarNumber),
              })}
            </legend>
            <div className="work-office-field">
              <span>
                {officeMessage(messages, 'spreadsheet.chart.errorBar.direction')}
              </span>
              <OfficeSelect
                ariaLabel={officeMessage(
                  messages,
                  'spreadsheet.chart.errorBar.directionAria',
                  { prefix: labelPrefix },
                )}
                value={item.direction}
                options={[
                  ...(workSpreadsheetChartUsesNumericXAxis(chartType)
                    ? [{ value: 'x', label: 'X' } as const]
                    : []),
                  { value: 'y', label: 'Y' },
                ]}
                onValueChange={(direction) =>
                  replaceErrorBars(index, {
                    direction: direction as WorkSpreadsheetErrorBarDirection,
                  })
                }
              />
            </div>
            <div className="work-office-field">
              <span>
                {officeMessage(messages, 'spreadsheet.chart.errorBar.barType')}
              </span>
              <OfficeSelect
                ariaLabel={officeMessage(
                  messages,
                  'spreadsheet.chart.errorBar.barTypeAria',
                  { prefix: labelPrefix },
                )}
                value={item.barType}
                options={[
                  {
                    value: 'both',
                    label: officeMessage(
                      messages,
                      'spreadsheet.chart.errorBar.barType.both',
                    ),
                  },
                  {
                    value: 'plus',
                    label: officeMessage(
                      messages,
                      'spreadsheet.chart.errorBar.barType.plus',
                    ),
                  },
                  {
                    value: 'minus',
                    label: officeMessage(
                      messages,
                      'spreadsheet.chart.errorBar.barType.minus',
                    ),
                  },
                ]}
                onValueChange={(barType) =>
                  replaceErrorBars(index, {
                    barType: barType as WorkSpreadsheetErrorBarType,
                  })
                }
              />
            </div>
            <div className="work-office-field">
              <span>
                {officeMessage(messages, 'spreadsheet.chart.errorBar.valueType')}
              </span>
              <OfficeSelect
                ariaLabel={officeMessage(
                  messages,
                  'spreadsheet.chart.errorBar.valueTypeAria',
                  { prefix: labelPrefix },
                )}
                value={item.valueType}
                options={[
                  {
                    value: 'fixedValue',
                    label: officeMessage(
                      messages,
                      'spreadsheet.chart.errorBar.valueType.fixedValue',
                    ),
                  },
                  {
                    value: 'percentage',
                    label: officeMessage(
                      messages,
                      'spreadsheet.chart.errorBar.valueType.percentage',
                    ),
                  },
                  {
                    value: 'standardDeviation',
                    label: officeMessage(
                      messages,
                      'spreadsheet.chart.errorBar.valueType.standardDeviation',
                    ),
                  },
                  {
                    value: 'standardError',
                    label: officeMessage(
                      messages,
                      'spreadsheet.chart.errorBar.valueType.standardError',
                    ),
                  },
                  {
                    value: 'custom',
                    label: officeMessage(
                      messages,
                      'spreadsheet.chart.errorBar.valueType.customRange',
                    ),
                  },
                ]}
                onValueChange={(valueType) =>
                  replaceErrorBars(
                    index,
                    errorBarsWithValueType(
                      item,
                      valueType as WorkSpreadsheetErrorBarValueType,
                    ),
                  )
                }
              />
            </div>
            {item.valueType === 'fixedValue' ||
            item.valueType === 'percentage' ||
            item.valueType === 'standardDeviation' ? (
              <div className="work-office-field">
                <span>
                  {item.valueType === 'percentage'
                    ? officeMessage(
                        messages,
                        'spreadsheet.chart.errorBar.percentLabel',
                      )
                    : officeMessage(
                        messages,
                        'spreadsheet.chart.errorBar.valueLabel',
                      )}
                </span>
                <CommittedOfficeNumberField
                  ariaLabel={officeMessage(
                    messages,
                    'spreadsheet.chart.errorBar.valueAria',
                    { prefix: labelPrefix },
                  )}
                  min={0}
                  step={0.1}
                  value={
                    item.value ?? (item.valueType === 'percentage' ? 5 : 1)
                  }
                  normalizeValue={(value) =>
                    normalizeRequiredOfficeNumber(value, { minimum: 0 })
                  }
                  onValueCommit={(value) => replaceErrorBars(index, { value })}
                />
              </div>
            ) : null}
            {item.valueType === 'custom' &&
            item.barType !== 'minus' &&
            customInput === 'references' ? (
              <div className="work-office-field error-reference">
                <span>
                  {officeMessage(messages, 'spreadsheet.chart.errorBar.plusRef')}
                </span>
                <OfficeTextField
                  aria-label={officeMessage(
                    messages,
                    'spreadsheet.chart.errorBar.plusRefAria',
                    { prefix: labelPrefix },
                  )}
                  value={item.plusReference ?? ''}
                  placeholder={officeMessage(
                    messages,
                    'spreadsheet.chart.errorBar.refPlaceholder',
                  )}
                  onChange={(event) =>
                    replaceErrorBars(index, {
                      plusReference: event.target.value,
                    })
                  }
                />
              </div>
            ) : null}
            {item.valueType === 'custom' &&
            item.barType !== 'plus' &&
            customInput === 'references' ? (
              <div className="work-office-field error-reference">
                <span>
                  {officeMessage(
                    messages,
                    'spreadsheet.chart.errorBar.minusRef',
                  )}
                </span>
                <OfficeTextField
                  aria-label={officeMessage(
                    messages,
                    'spreadsheet.chart.errorBar.minusRefAria',
                    { prefix: labelPrefix },
                  )}
                  value={item.minusReference ?? ''}
                  placeholder={officeMessage(
                    messages,
                    'spreadsheet.chart.errorBar.minusRefPlaceholder',
                  )}
                  onChange={(event) =>
                    replaceErrorBars(index, {
                      minusReference: event.target.value,
                    })
                  }
                />
              </div>
            ) : null}
            {item.valueType === 'custom' &&
            item.barType !== 'minus' &&
            customInput === 'values' ? (
              <div className="work-office-field error-reference">
                <span>
                  {officeMessage(
                    messages,
                    'spreadsheet.chart.errorBar.plusValue',
                  )}
                </span>
                <CustomErrorValuesInput
                  label={officeMessage(
                    messages,
                    'spreadsheet.chart.errorBar.plusValueAria',
                    { prefix: labelPrefix },
                  )}
                  id={`work-error-values-${seriesNumber}-${errorBarNumber}-plus`}
                  values={item.plusValues}
                  reference={item.plusReference}
                  keptRefPlaceholder={officeMessage(
                    messages,
                    'spreadsheet.chart.errorBar.keptRefPlaceholder',
                  )}
                  onCommit={(plusValues) =>
                    replaceErrorBars(index, {
                      plusValues,
                      plusReference: undefined,
                    })
                  }
                />
                {item.plusReference ? (
                  <small title={item.plusReference}>
                    {officeMessage(
                      messages,
                      'spreadsheet.chart.errorBar.keptPlusRef',
                      { ref: item.plusReference },
                    )}
                  </small>
                ) : null}
              </div>
            ) : null}
            {item.valueType === 'custom' &&
            item.barType !== 'plus' &&
            customInput === 'values' ? (
              <div className="work-office-field error-reference">
                <span>
                  {officeMessage(
                    messages,
                    'spreadsheet.chart.errorBar.minusValue',
                  )}
                </span>
                <CustomErrorValuesInput
                  label={officeMessage(
                    messages,
                    'spreadsheet.chart.errorBar.minusValueAria',
                    { prefix: labelPrefix },
                  )}
                  id={`work-error-values-${seriesNumber}-${errorBarNumber}-minus`}
                  values={item.minusValues}
                  reference={item.minusReference}
                  keptRefPlaceholder={officeMessage(
                    messages,
                    'spreadsheet.chart.errorBar.keptRefPlaceholder',
                  )}
                  onCommit={(minusValues) =>
                    replaceErrorBars(index, {
                      minusValues,
                      minusReference: undefined,
                    })
                  }
                />
                {item.minusReference ? (
                  <small title={item.minusReference}>
                    {officeMessage(
                      messages,
                      'spreadsheet.chart.errorBar.keptMinusRef',
                      { ref: item.minusReference },
                    )}
                  </small>
                ) : null}
              </div>
            ) : null}
            <OfficeCheckbox
              className="check"
              ariaLabel={officeMessage(
                messages,
                'spreadsheet.chart.errorBar.endCapsAria',
                { prefix: labelPrefix },
              )}
              checked={item.showEndCaps !== false}
              onCheckedChange={(showEndCaps) =>
                replaceErrorBars(index, { showEndCaps })
              }
            >
              {officeMessage(messages, 'spreadsheet.chart.errorBar.endCaps')}
            </OfficeCheckbox>
            <button
              type="button"
              className="remove-error-bars"
              aria-label={officeMessage(
                messages,
                'spreadsheet.chart.errorBar.deleteAria',
                { prefix: labelPrefix },
              )}
              onClick={() =>
                onChange(
                  errorBars.filter((_, candidate) => candidate !== index),
                )
              }
            >
              <Trash2 size={12} />
            </button>
          </fieldset>
        );
      })}
    </section>
  );
}

function errorBarsWithValueType(
  errorBars: WorkSpreadsheetErrorBars,
  valueType: WorkSpreadsheetErrorBarValueType,
): Partial<WorkSpreadsheetErrorBars> {
  if (valueType === 'standardError' || valueType === 'custom')
    return { valueType, value: undefined };
  return {
    valueType,
    value: errorBars.value ?? (valueType === 'percentage' ? 5 : 1),
  };
}

function CustomErrorValuesInput({
  label,
  id,
  values,
  reference,
  keptRefPlaceholder,
  onCommit,
}: {
  label: string;
  id: string;
  values: number[] | undefined;
  reference: string | undefined;
  keptRefPlaceholder: string;
  onCommit: (values: number[] | undefined) => void;
}) {
  return (
    <CommittedOfficeTextField
      id={id}
      aria-label={label}
      value={values}
      formatValue={(items) => items?.join(', ') ?? ''}
      parseValue={parseCustomValues}
      placeholder={reference ? keptRefPlaceholder : '1, 2, 1.5'}
      onValueCommit={onCommit}
    />
  );
}

function parseCustomValues(value: string): number[] | undefined | null {
  if (!value.trim()) return undefined;
  const tokens = value
    .split(/[\s,;，；]+/)
    .map((item) => item.trim())
    .filter(Boolean);
  if (!tokens.length || tokens.length > 256) return null;
  const values = tokens.map(Number);
  return values.every((item) => Number.isFinite(item) && item >= 0)
    ? values
    : null;
}
