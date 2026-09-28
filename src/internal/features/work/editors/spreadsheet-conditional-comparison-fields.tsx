import { officeMessage } from '../../../i18n/office-locale';
import {
  type SpreadsheetConditionalComparisonOperator,
  spreadsheetConditionalComparisonNeedsUpperValue,
} from '../work-spreadsheet-conditional-comparisons';
import {
  OfficeCheckbox,
  OfficeColorPicker,
  OfficeSelect,
  OfficeTextField,
} from './office-controls';
import { useOfficeMessages } from './office-messages-context';
import {
  type ConditionalRuleDraft,
  SPREADSHEET_CONDITIONAL_COMPARISONS,
} from './spreadsheet-conditional-format-model';

interface SpreadsheetConditionalComparisonFieldsProps {
  draft: ConditionalRuleDraft;
  onChange: (patch: Partial<ConditionalRuleDraft>) => void;
}

export function SpreadsheetConditionalComparisonFields({
  draft,
  onChange,
}: SpreadsheetConditionalComparisonFieldsProps) {
  const messages = useOfficeMessages();
  const needsUpperValue = spreadsheetConditionalComparisonNeedsUpperValue(
    draft.comparisonOperator,
  );
  return (
    <>
      <div className="work-office-field">
        <span>
          {officeMessage(messages, 'spreadsheet.cf.comparison.method')}
        </span>
        <OfficeSelect
          ariaLabel={officeMessage(
            messages,
            'spreadsheet.cf.comparison.operatorAria',
          )}
          value={draft.comparisonOperator}
          options={SPREADSHEET_CONDITIONAL_COMPARISONS.map((comparison) => ({
            value: comparison.name,
            label: comparison.label,
          }))}
          onValueChange={(comparisonOperator) =>
            onChange({
              comparisonOperator:
                comparisonOperator as SpreadsheetConditionalComparisonOperator,
            })
          }
        />
      </div>
      <div className="work-office-field">
        <span>
          {needsUpperValue
            ? officeMessage(messages, 'spreadsheet.cf.comparison.lower')
            : officeMessage(messages, 'spreadsheet.cf.comparison.value')}
        </span>
        <OfficeTextField
          aria-label={
            needsUpperValue
              ? officeMessage(messages, 'spreadsheet.cf.comparison.lowerAria')
              : officeMessage(messages, 'spreadsheet.cf.comparison.valueAria')
          }
          value={draft.comparisonValue}
          onChange={(event) =>
            onChange({ comparisonValue: event.target.value })
          }
        />
      </div>
      {needsUpperValue ? (
        <div className="work-office-field">
          <span>
            {officeMessage(messages, 'spreadsheet.cf.comparison.upper')}
          </span>
          <OfficeTextField
            aria-label={officeMessage(
              messages,
              'spreadsheet.cf.comparison.upperAria',
            )}
            value={draft.comparisonUpperValue}
            onChange={(event) =>
              onChange({ comparisonUpperValue: event.target.value })
            }
          />
        </div>
      ) : null}
      <OfficeCheckbox
        className="toggle"
        ariaLabel={officeMessage(
          messages,
          'spreadsheet.cf.comparison.setTextColor',
        )}
        checked={draft.comparisonUseTextColor}
        onCheckedChange={(comparisonUseTextColor) =>
          onChange({ comparisonUseTextColor })
        }
      >
        {officeMessage(messages, 'spreadsheet.cf.comparison.setTextColor')}
      </OfficeCheckbox>
      <div className="work-office-field color">
        <span>
          {officeMessage(messages, 'spreadsheet.cf.comparison.textColor')}
        </span>
        <OfficeColorPicker
          ariaLabel={officeMessage(
            messages,
            'spreadsheet.cf.comparison.textColorAria',
          )}
          value={draft.comparisonTextColor}
          disabled={!draft.comparisonUseTextColor}
          onValueChange={(comparisonTextColor) =>
            onChange({ comparisonTextColor })
          }
        />
      </div>
      <OfficeCheckbox
        className="toggle"
        ariaLabel={officeMessage(
          messages,
          'spreadsheet.cf.comparison.setFillColor',
        )}
        checked={draft.comparisonUseCellColor}
        onCheckedChange={(comparisonUseCellColor) =>
          onChange({ comparisonUseCellColor })
        }
      >
        {officeMessage(messages, 'spreadsheet.cf.comparison.setFillColor')}
      </OfficeCheckbox>
      <div className="work-office-field color">
        <span>
          {officeMessage(messages, 'spreadsheet.cf.comparison.fillColor')}
        </span>
        <OfficeColorPicker
          ariaLabel={officeMessage(
            messages,
            'spreadsheet.cf.comparison.fillColorAria',
          )}
          value={draft.comparisonCellColor}
          disabled={!draft.comparisonUseCellColor}
          onValueChange={(comparisonCellColor) =>
            onChange({ comparisonCellColor })
          }
        />
      </div>
    </>
  );
}
