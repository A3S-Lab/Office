import { officeMessage } from '../../../i18n/office-locale';
import type { SpreadsheetConditionalThresholdType } from '../work-spreadsheet-conditional-values';
import {
  OfficeCheckbox,
  OfficeNumberField,
  OfficeSelect,
} from './office-controls';
import { useOfficeMessages } from './office-messages-context';
import type { SpreadsheetConditionalThresholdDraft } from './spreadsheet-conditional-format-model';

interface SpreadsheetConditionalThresholdFieldsProps {
  label: string;
  thresholds: SpreadsheetConditionalThresholdDraft[];
  startIndex?: number;
  showEquality?: boolean;
  onChange: (
    index: number,
    patch: Partial<SpreadsheetConditionalThresholdDraft>,
  ) => void;
}

export function SpreadsheetConditionalThresholdFields({
  label,
  thresholds,
  startIndex = 0,
  showEquality = false,
  onChange,
}: SpreadsheetConditionalThresholdFieldsProps) {
  const messages = useOfficeMessages();
  return (
    <fieldset className="work-spreadsheet-conditional-thresholds">
      <legend>
        {officeMessage(messages, 'spreadsheet.cf.threshold.legend', { label })}
      </legend>
      {thresholds.slice(startIndex).map((threshold, offset) => {
        const index = startIndex + offset;
        const indexLabel = String(index + 1);
        const valueRequired =
          threshold.type !== 'min' && threshold.type !== 'max';
        return (
          <div key={index}>
            <span>
              {officeMessage(messages, 'spreadsheet.cf.threshold.start', {
                index: indexLabel,
              })}
            </span>
            <OfficeSelect
              ariaLabel={officeMessage(
                messages,
                'spreadsheet.cf.threshold.typeAria',
                { label, index: indexLabel },
              )}
              value={threshold.type}
              options={[
                {
                  value: 'percent',
                  label: officeMessage(
                    messages,
                    'spreadsheet.cf.threshold.type.percent',
                  ),
                },
                {
                  value: 'percentile',
                  label: officeMessage(
                    messages,
                    'spreadsheet.cf.threshold.type.percentile',
                  ),
                },
                {
                  value: 'num',
                  label: officeMessage(
                    messages,
                    'spreadsheet.cf.threshold.type.num',
                  ),
                },
                {
                  value: 'min',
                  label: officeMessage(
                    messages,
                    'spreadsheet.cf.threshold.type.min',
                  ),
                },
                {
                  value: 'max',
                  label: officeMessage(
                    messages,
                    'spreadsheet.cf.threshold.type.max',
                  ),
                },
              ]}
              onValueChange={(type) =>
                onChange(index, {
                  type: type as SpreadsheetConditionalThresholdType,
                })
              }
            />
            <OfficeNumberField
              ariaLabel={officeMessage(
                messages,
                'spreadsheet.cf.threshold.valueAria',
                { label, index: indexLabel },
              )}
              value={valueRequired ? threshold.value : ''}
              disabled={!valueRequired}
              onValueChange={(value) => onChange(index, { value })}
            />
            {showEquality ? (
              <OfficeCheckbox
                className="threshold-gte"
                ariaLabel={officeMessage(
                  messages,
                  'spreadsheet.cf.threshold.includeEqualAria',
                  { label, index: indexLabel },
                )}
                checked={threshold.gte}
                onCheckedChange={(gte) => onChange(index, { gte })}
              >
                {officeMessage(
                  messages,
                  'spreadsheet.cf.threshold.includeEqual',
                )}
              </OfficeCheckbox>
            ) : (
              <span />
            )}
          </div>
        );
      })}
    </fieldset>
  );
}
