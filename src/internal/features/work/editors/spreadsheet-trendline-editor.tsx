import { Plus, Trash2 } from 'lucide-react';
import { officeMessage } from '../../../i18n/office-locale';
import type {
  WorkSpreadsheetTrendline,
  WorkSpreadsheetTrendlineType,
} from '../work-types';
import { useOfficeMessages } from './office-messages-context';
import {
  CommittedOfficeNumberField,
  OfficeCheckbox,
  OfficeSelect,
  OfficeTextField,
} from './office-controls';
import { normalizeRequiredOfficeNumber } from './office-number-normalization';

interface SpreadsheetTrendlineEditorProps {
  seriesNumber: number;
  trendlines: WorkSpreadsheetTrendline[];
  onChange: (trendlines: WorkSpreadsheetTrendline[]) => void;
}

export function SpreadsheetTrendlineEditor({
  seriesNumber,
  trendlines,
  onChange,
}: SpreadsheetTrendlineEditorProps) {
  const messages = useOfficeMessages();
  const replaceTrendline = (
    index: number,
    change: Partial<WorkSpreadsheetTrendline>,
  ) => {
    onChange(
      trendlines.map((trendline, candidate) =>
        candidate === index ? { ...trendline, ...change } : trendline,
      ),
    );
  };

  return (
    <section
      className="work-spreadsheet-trendlines"
      aria-label={officeMessage(
        messages,
        'spreadsheet.chart.trendline.sectionAria',
        { n: String(seriesNumber) },
      )}
    >
      <header>
        <strong>
          {officeMessage(messages, 'spreadsheet.chart.trendline.title')}
        </strong>
        <button
          type="button"
          aria-label={officeMessage(
            messages,
            'spreadsheet.chart.trendline.addAria',
            { n: String(seriesNumber) },
          )}
          onClick={() => onChange([...trendlines, { type: 'linear' }])}
        >
          <Plus size={11} />
          {officeMessage(messages, 'spreadsheet.chart.trendline.add')}
        </button>
      </header>
      {!trendlines.length ? (
        <p>{officeMessage(messages, 'spreadsheet.chart.trendline.empty')}</p>
      ) : null}
      {trendlines.map((trendline, index) => {
        const trendlineNumber = index + 1;
        const labelPrefix = officeMessage(
          messages,
          'spreadsheet.chart.trendline.itemPrefix',
          {
            n: String(seriesNumber),
            index: String(trendlineNumber),
          },
        );
        const hasIntercept = trendline.intercept !== undefined;
        return (
          <fieldset key={`${seriesNumber}-${trendlineNumber}`}>
            <legend>
              {officeMessage(messages, 'spreadsheet.chart.trendline.itemLegend', {
                index: String(trendlineNumber),
              })}
            </legend>
            <div className="work-office-field">
              <span>
                {officeMessage(messages, 'spreadsheet.chart.trendline.type')}
              </span>
              <OfficeSelect
                ariaLabel={officeMessage(
                  messages,
                  'spreadsheet.chart.trendline.typeAria',
                  { prefix: labelPrefix },
                )}
                value={trendline.type}
                options={[
                  {
                    value: 'linear',
                    label: officeMessage(
                      messages,
                      'spreadsheet.chart.trendline.type.linear',
                    ),
                  },
                  {
                    value: 'exponential',
                    label: officeMessage(
                      messages,
                      'spreadsheet.chart.trendline.type.exponential',
                    ),
                  },
                  {
                    value: 'logarithmic',
                    label: officeMessage(
                      messages,
                      'spreadsheet.chart.trendline.type.logarithmic',
                    ),
                  },
                  {
                    value: 'polynomial',
                    label: officeMessage(
                      messages,
                      'spreadsheet.chart.trendline.type.polynomial',
                    ),
                  },
                  {
                    value: 'power',
                    label: officeMessage(
                      messages,
                      'spreadsheet.chart.trendline.type.power',
                    ),
                  },
                  {
                    value: 'movingAverage',
                    label: officeMessage(
                      messages,
                      'spreadsheet.chart.trendline.type.movingAverage',
                    ),
                  },
                ]}
                onValueChange={(value) =>
                  onChange(
                    trendlines.map((item, candidate) =>
                      candidate === index
                        ? trendlineWithType(
                            item,
                            value as WorkSpreadsheetTrendlineType,
                          )
                        : item,
                    ),
                  )
                }
              />
            </div>
            <div className="work-office-field">
              <span>
                {officeMessage(messages, 'spreadsheet.chart.trendline.name')}
              </span>
              <OfficeTextField
                aria-label={officeMessage(
                  messages,
                  'spreadsheet.chart.trendline.nameAria',
                  { prefix: labelPrefix },
                )}
                value={trendline.name ?? ''}
                maxLength={255}
                placeholder={officeMessage(
                  messages,
                  'spreadsheet.chart.trendline.namePlaceholder',
                  { index: String(trendlineNumber) },
                )}
                onChange={(event) =>
                  replaceTrendline(index, { name: event.target.value })
                }
              />
            </div>
            {trendline.type === 'polynomial' ? (
              <div className="work-office-field">
                <span>
                  {officeMessage(messages, 'spreadsheet.chart.trendline.order')}
                </span>
                <CommittedOfficeNumberField
                  ariaLabel={officeMessage(
                    messages,
                    'spreadsheet.chart.trendline.orderAria',
                    { prefix: labelPrefix },
                  )}
                  min={2}
                  max={6}
                  step={1}
                  value={trendline.order ?? 2}
                  normalizeValue={(value) =>
                    normalizeRequiredOfficeNumber(value, {
                      integer: true,
                      minimum: 2,
                      maximum: 6,
                    })
                  }
                  onValueCommit={(order) => replaceTrendline(index, { order })}
                />
              </div>
            ) : null}
            {trendline.type === 'movingAverage' ? (
              <div className="work-office-field">
                <span>
                  {officeMessage(messages, 'spreadsheet.chart.trendline.period')}
                </span>
                <CommittedOfficeNumberField
                  ariaLabel={officeMessage(
                    messages,
                    'spreadsheet.chart.trendline.periodAria',
                    { prefix: labelPrefix },
                  )}
                  min={2}
                  max={255}
                  step={1}
                  value={trendline.period ?? 2}
                  normalizeValue={(value) =>
                    normalizeRequiredOfficeNumber(value, {
                      integer: true,
                      minimum: 2,
                      maximum: 255,
                    })
                  }
                  onValueCommit={(period) =>
                    replaceTrendline(index, { period })
                  }
                />
              </div>
            ) : null}
            <div className="work-office-field">
              <span>
                {officeMessage(messages, 'spreadsheet.chart.trendline.forward')}
              </span>
              <CommittedOfficeNumberField
                ariaLabel={officeMessage(
                  messages,
                  'spreadsheet.chart.trendline.forwardAria',
                  { prefix: labelPrefix },
                )}
                min={0}
                step={0.1}
                value={trendline.forward ?? 0}
                normalizeValue={(value) =>
                  normalizeRequiredOfficeNumber(value, { minimum: 0 })
                }
                onValueCommit={(forward) =>
                  replaceTrendline(index, { forward })
                }
              />
            </div>
            <div className="work-office-field">
              <span>
                {officeMessage(messages, 'spreadsheet.chart.trendline.backward')}
              </span>
              <CommittedOfficeNumberField
                ariaLabel={officeMessage(
                  messages,
                  'spreadsheet.chart.trendline.backwardAria',
                  { prefix: labelPrefix },
                )}
                min={0}
                step={0.1}
                value={trendline.backward ?? 0}
                normalizeValue={(value) =>
                  normalizeRequiredOfficeNumber(value, { minimum: 0 })
                }
                onValueCommit={(backward) =>
                  replaceTrendline(index, { backward })
                }
              />
            </div>
            <OfficeCheckbox
              className="check intercept-toggle"
              ariaLabel={officeMessage(
                messages,
                'spreadsheet.chart.trendline.interceptAria',
                { prefix: labelPrefix },
              )}
              checked={hasIntercept}
              onCheckedChange={(checked) =>
                replaceTrendline(index, { intercept: checked ? 0 : undefined })
              }
            >
              {officeMessage(messages, 'spreadsheet.chart.trendline.intercept')}
            </OfficeCheckbox>
            <div className="work-office-field">
              <span>
                {officeMessage(
                  messages,
                  'spreadsheet.chart.trendline.interceptValue',
                )}
              </span>
              <CommittedOfficeNumberField
                ariaLabel={officeMessage(
                  messages,
                  'spreadsheet.chart.trendline.interceptValueAria',
                  { prefix: labelPrefix },
                )}
                step={0.1}
                disabled={!hasIntercept}
                value={trendline.intercept ?? 0}
                normalizeValue={normalizeRequiredOfficeNumber}
                onValueCommit={(intercept) =>
                  replaceTrendline(index, { intercept })
                }
              />
            </div>
            <OfficeCheckbox
              className="check"
              ariaLabel={officeMessage(
                messages,
                'spreadsheet.chart.trendline.showEqAria',
                { prefix: labelPrefix },
              )}
              checked={trendline.displayEquation === true}
              onCheckedChange={(displayEquation) =>
                replaceTrendline(index, { displayEquation })
              }
            >
              {officeMessage(messages, 'spreadsheet.chart.trendline.showEq')}
            </OfficeCheckbox>
            <OfficeCheckbox
              className="check"
              ariaLabel={officeMessage(
                messages,
                'spreadsheet.chart.trendline.showR2Aria',
                { prefix: labelPrefix },
              )}
              checked={trendline.displayRSquared === true}
              onCheckedChange={(displayRSquared) =>
                replaceTrendline(index, { displayRSquared })
              }
            >
              {officeMessage(messages, 'spreadsheet.chart.trendline.showR2')}
            </OfficeCheckbox>
            <button
              type="button"
              className="remove-trendline"
              aria-label={officeMessage(
                messages,
                'spreadsheet.chart.trendline.deleteAria',
                { prefix: labelPrefix },
              )}
              onClick={() =>
                onChange(
                  trendlines.filter((_, candidate) => candidate !== index),
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

function trendlineWithType(
  trendline: WorkSpreadsheetTrendline,
  type: WorkSpreadsheetTrendlineType,
): WorkSpreadsheetTrendline {
  const next = { ...trendline, type, order: undefined, period: undefined };
  if (type === 'polynomial') return { ...next, order: trendline.order ?? 2 };
  if (type === 'movingAverage')
    return { ...next, period: trendline.period ?? 2, intercept: undefined };
  return next;
}
