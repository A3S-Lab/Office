import { officeMessage } from '../../../i18n/office-locale';
import {
  normalizeDocumentColumns,
  setCustomDocumentColumns,
  updateDocumentColumnWidth,
} from '../work-document-columns';
import type { WorkDocumentColumns } from '../work-types';
import { CommittedOfficeNumberField, OfficeCheckbox } from './office-controls';
import { useOfficeMessages } from './office-messages-context';

export function DocumentColumnsPanel({
  columns,
  onChange,
}: {
  columns: WorkDocumentColumns;
  onChange: (columns: WorkDocumentColumns) => void;
}) {
  const messages = useOfficeMessages();
  const normalized = normalizeDocumentColumns(columns);
  const custom = normalized.custom;
  const updateCustomSpacing = (index: number, spacing: number) => {
    if (!custom) return;
    onChange(
      normalizeDocumentColumns({
        ...normalized,
        custom: custom.map((column, columnIndex) =>
          columnIndex === index ? { ...column, spacing } : column,
        ),
      }),
    );
  };

  return (
    <fieldset className="work-document-columns-panel">
      <legend>{officeMessage(messages, 'document.columns.legend')}</legend>
      <div className="work-office-field">
        <span>{officeMessage(messages, 'document.columns.count')}</span>
        <CommittedOfficeNumberField
          min={1}
          max={6}
          step={1}
          ariaLabel={officeMessage(messages, 'document.columns.countAria')}
          value={normalized.count}
          normalizeValue={(value) => normalizeRequiredInteger(value, 1, 6)}
          onValueCommit={(count) =>
            onChange(
              normalizeDocumentColumns({
                ...normalized,
                count,
              }),
            )
          }
        />
      </div>
      {!normalized.custom && (
        <div className="work-office-field">
          <span>{officeMessage(messages, 'document.columns.spacing')}</span>
          <CommittedOfficeNumberField
            min={0}
            max={30}
            step={0.5}
            ariaLabel={officeMessage(messages, 'document.columns.spacingAria')}
            value={normalized.spacing}
            normalizeValue={(value) => normalizeRequiredDecimal(value, 0, 30)}
            onValueCommit={(spacing) =>
              onChange(
                normalizeDocumentColumns({
                  ...normalized,
                  spacing,
                }),
              )
            }
          />
        </div>
      )}
      <OfficeCheckbox
        className="work-document-column-option"
        ariaLabel={officeMessage(messages, 'document.columns.customWidths')}
        disabled={normalized.count < 2}
        checked={Boolean(normalized.custom)}
        onCheckedChange={(checked) =>
          onChange(setCustomDocumentColumns(normalized, checked))
        }
      >
        {officeMessage(messages, 'document.columns.customWidths')}
      </OfficeCheckbox>
      <OfficeCheckbox
        className="work-document-column-option"
        ariaLabel={officeMessage(messages, 'document.columns.separatorAria')}
        checked={normalized.separator}
        onCheckedChange={(checked) =>
          onChange({ ...normalized, separator: checked })
        }
      >
        {officeMessage(messages, 'document.columns.separator')}
      </OfficeCheckbox>
      {custom && (
        <div className="work-document-custom-columns">
          {custom.map((column, index) => {
            const n = String(index + 1);
            return (
              <div
                key={`document-column-${index + 1}`}
                className="work-document-custom-column"
              >
                <strong>
                  {officeMessage(messages, 'document.columns.columnN', { n })}
                </strong>
                <div className="work-office-field">
                  <span>
                    {officeMessage(messages, 'document.columns.widthPercent')}
                  </span>
                  <CommittedOfficeNumberField
                    min={5}
                    max={100 - (normalized.count - 1) * 5}
                    step={0.5}
                    ariaLabel={officeMessage(
                      messages,
                      'document.columns.widthAria',
                      { n },
                    )}
                    value={column.widthPercent}
                    normalizeValue={(value) =>
                      normalizeRequiredDecimal(
                        value,
                        5,
                        100 - (normalized.count - 1) * 5,
                      )
                    }
                    onValueCommit={(widthPercent) =>
                      onChange(
                        updateDocumentColumnWidth(
                          normalized,
                          index,
                          widthPercent,
                        ),
                      )
                    }
                  />
                </div>
                {index < custom.length - 1 && (
                  <div className="work-office-field">
                    <span>
                      {officeMessage(
                        messages,
                        'document.columns.afterSpacing',
                      )}
                    </span>
                    <CommittedOfficeNumberField
                      min={0}
                      max={30}
                      step={0.5}
                      ariaLabel={officeMessage(
                        messages,
                        'document.columns.afterSpacingAria',
                        { n },
                      )}
                      value={column.spacing}
                      normalizeValue={(value) =>
                        normalizeRequiredDecimal(value, 0, 30)
                      }
                      onValueCommit={(spacing) =>
                        updateCustomSpacing(index, spacing)
                      }
                    />
                  </div>
                )}
              </div>
            );
          })}
          <p>{officeMessage(messages, 'document.columns.hint')}</p>
        </div>
      )}
    </fieldset>
  );
}

function normalizeRequiredInteger(
  value: string,
  minimum: number,
  maximum: number,
): number | null {
  if (!value.trim()) return null;
  const number = Number(value);
  return Number.isFinite(number)
    ? Math.min(maximum, Math.max(minimum, Math.round(number)))
    : null;
}

function normalizeRequiredDecimal(
  value: string,
  minimum: number,
  maximum: number,
): number | null {
  if (!value.trim()) return null;
  const number = Number(value);
  return Number.isFinite(number)
    ? Math.min(maximum, Math.max(minimum, Math.round(number * 10) / 10))
    : null;
}
