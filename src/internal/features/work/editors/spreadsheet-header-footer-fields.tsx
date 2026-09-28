import { officeMessage } from '../../../i18n/office-locale';
import type { EffectiveSpreadsheetPageSetup } from '../work-spreadsheet-page-setup';
import { useOfficeMessages } from './office-messages-context';
import {
  CommittedOfficeNumberField,
  OfficeCheckbox,
  OfficeSelect,
  OfficeTextField,
} from './office-controls';

export function SpreadsheetHeaderFooterFields({
  pageSetup,
  onChange,
}: {
  pageSetup: EffectiveSpreadsheetPageSetup;
  onChange: (pageSetup: EffectiveSpreadsheetPageSetup) => void;
}) {
  const messages = useOfficeMessages();
  const updateSection = (
    area: 'header' | 'footer',
    section: 'left' | 'center' | 'right',
    value: string,
  ) => {
    onChange({
      ...pageSetup,
      [area]: { ...pageSetup[area], [section]: value },
    });
  };
  return (
    <fieldset className="work-spreadsheet-header-footer-fields">
      <legend>{officeMessage(messages, 'spreadsheet.print.hf.legend')}</legend>
      <TemplateField
        label={officeMessage(messages, 'spreadsheet.print.hf.headerLeft')}
        value={pageSetup.header.left}
        onChange={(value) => updateSection('header', 'left', value)}
      />
      <TemplateField
        label={officeMessage(messages, 'spreadsheet.print.hf.headerCenter')}
        value={pageSetup.header.center}
        onChange={(value) => updateSection('header', 'center', value)}
      />
      <TemplateField
        label={officeMessage(messages, 'spreadsheet.print.hf.headerRight')}
        value={pageSetup.header.right}
        showPagePlaceholder
        onChange={(value) => updateSection('header', 'right', value)}
      />
      <TemplateField
        label={officeMessage(messages, 'spreadsheet.print.hf.footerLeft')}
        value={pageSetup.footer.left}
        onChange={(value) => updateSection('footer', 'left', value)}
      />
      <TemplateField
        label={officeMessage(messages, 'spreadsheet.print.hf.footerCenter')}
        value={pageSetup.footer.center}
        onChange={(value) => updateSection('footer', 'center', value)}
      />
      <TemplateField
        label={officeMessage(messages, 'spreadsheet.print.hf.footerRight')}
        value={pageSetup.footer.right}
        showPagePlaceholder
        onChange={(value) => updateSection('footer', 'right', value)}
      />
      <div className="work-office-field">
        <span>{officeMessage(messages, 'spreadsheet.print.hf.pageStart')}</span>
        <CommittedOfficeNumberField
          ariaLabel={officeMessage(messages, 'spreadsheet.print.hf.pageStart')}
          min={1}
          max={32767}
          value={pageSetup.pageNumberStart}
          normalizeValue={normalizeStartingPageNumber}
          onValueCommit={(pageNumberStart) =>
            onChange({ ...pageSetup, pageNumberStart })
          }
        />
      </div>
      <div className="work-office-field">
        <span>{officeMessage(messages, 'spreadsheet.print.hf.pageOrder')}</span>
        <OfficeSelect
          ariaLabel={officeMessage(messages, 'spreadsheet.print.hf.pageOrder')}
          value={pageSetup.pageOrder}
          options={[
            {
              value: 'overThenDown',
              label: officeMessage(
                messages,
                'spreadsheet.print.hf.overThenDown',
              ),
            },
            {
              value: 'downThenOver',
              label: officeMessage(
                messages,
                'spreadsheet.print.hf.downThenOver',
              ),
            },
          ]}
          onValueChange={(pageOrder) =>
            onChange({
              ...pageSetup,
              pageOrder:
                pageOrder as EffectiveSpreadsheetPageSetup['pageOrder'],
            })
          }
        />
      </div>
      <OfficeCheckbox
        className="toggle"
        ariaLabel={officeMessage(messages, 'spreadsheet.print.hf.scaleWithDoc')}
        checked={pageSetup.scaleWithDocument}
        onCheckedChange={(scaleWithDocument) =>
          onChange({ ...pageSetup, scaleWithDocument })
        }
      >
        {officeMessage(messages, 'spreadsheet.print.hf.scaleWithDoc')}
      </OfficeCheckbox>
      <OfficeCheckbox
        className="toggle"
        ariaLabel={officeMessage(messages, 'spreadsheet.print.hf.alignMargins')}
        checked={pageSetup.alignWithMargins}
        onCheckedChange={(alignWithMargins) =>
          onChange({ ...pageSetup, alignWithMargins })
        }
      >
        {officeMessage(messages, 'spreadsheet.print.hf.alignMargins')}
      </OfficeCheckbox>
      <p className="tokens">
        {officeMessage(messages, 'spreadsheet.print.hf.tokens')}
      </p>
    </fieldset>
  );
}

function normalizeStartingPageNumber(value: string): number | null {
  if (!value.trim()) return null;
  const number = Number(value);
  return Number.isFinite(number)
    ? Math.min(32_767, Math.max(1, Math.round(number)))
    : null;
}

function TemplateField({
  label,
  value,
  showPagePlaceholder = false,
  onChange,
}: {
  label: string;
  value: string;
  showPagePlaceholder?: boolean;
  onChange: (value: string) => void;
}) {
  return (
    <div className="work-office-field">
      <span>{label}</span>
      <OfficeTextField
        aria-label={label}
        value={value}
        placeholder={
          showPagePlaceholder ? 'Page {page} of {pages}' : ''
        }
        onChange={(event) => onChange(event.target.value)}
      />
    </div>
  );
}
