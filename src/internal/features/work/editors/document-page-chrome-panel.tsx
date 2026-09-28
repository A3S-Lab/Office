import { useState } from 'react';
import { officeMessage } from '../../../i18n/office-locale';
import {
  normalizeDocumentPageChrome,
  updateDocumentPageChromeVariant,
} from '../work-document-page-chrome';
import type {
  WorkDocumentPageChrome,
  WorkDocumentPageChromeContent,
  WorkDocumentPageChromeVariant,
} from '../work-types';
import { DocumentPageChromeRichTextEditor } from './document-page-chrome-editor';
import { OfficeCheckbox, OfficeSelect } from './office-controls';
import { useOfficeMessages } from './office-messages-context';

export function DocumentPageChromePanel({
  pageChrome,
  onChange,
}: {
  pageChrome: WorkDocumentPageChrome;
  onChange: (pageChrome: WorkDocumentPageChrome) => void;
}) {
  const messages = useOfficeMessages();
  const chrome = normalizeDocumentPageChrome(pageChrome);
  const [variant, setVariant] =
    useState<WorkDocumentPageChromeVariant>('default');
  const variantLabel = officeMessage(
    messages,
    variant === 'first'
      ? 'document.pageChrome.panel.variant.first'
      : variant === 'even'
        ? 'document.pageChrome.panel.variant.even'
        : 'document.pageChrome.panel.variant.default',
  );
  const updateVariant = (patch: Partial<WorkDocumentPageChromeContent>) => {
    onChange(updateDocumentPageChromeVariant(chrome, variant, patch));
  };
  const toggleFirstPage = (enabled: boolean) => {
    onChange({
      ...chrome,
      differentFirstPage: enabled,
      first:
        enabled && emptyPageChrome(chrome.first)
          ? { ...chrome.default }
          : chrome.first,
    });
    if (!enabled && variant === 'first') setVariant('default');
  };
  const toggleOddEvenPages = (enabled: boolean) => {
    onChange({
      ...chrome,
      differentOddEvenPages: enabled,
      even:
        enabled && emptyPageChrome(chrome.even)
          ? { ...chrome.default }
          : chrome.even,
    });
    if (!enabled && variant === 'even') setVariant('default');
  };

  return (
    <fieldset className="work-document-page-chrome-panel">
      <legend>
        {officeMessage(messages, 'document.pageChrome.panel.legend')}
      </legend>
      <div className="work-document-page-chrome-options">
        <OfficeCheckbox
          ariaLabel={officeMessage(
            messages,
            'document.pageChrome.panel.differentFirstAria',
          )}
          checked={chrome.differentFirstPage}
          onCheckedChange={toggleFirstPage}
        >
          {officeMessage(messages, 'document.pageChrome.panel.differentFirst')}
        </OfficeCheckbox>
        <OfficeCheckbox
          ariaLabel={officeMessage(
            messages,
            'document.pageChrome.panel.differentOddEvenAria',
          )}
          checked={chrome.differentOddEvenPages}
          onCheckedChange={toggleOddEvenPages}
        >
          {officeMessage(
            messages,
            'document.pageChrome.panel.differentOddEven',
          )}
        </OfficeCheckbox>
        <div className="work-office-field">
          <span>
            {officeMessage(messages, 'document.pageChrome.panel.edit')}
          </span>
          <OfficeSelect
            ariaLabel={officeMessage(
              messages,
              'document.pageChrome.panel.variantAria',
            )}
            value={variant}
            options={[
              {
                value: 'default',
                label: officeMessage(
                  messages,
                  'document.pageChrome.panel.variant.default',
                ),
              },
              {
                value: 'first',
                label: officeMessage(
                  messages,
                  'document.pageChrome.panel.variant.first',
                ),
                disabled: !chrome.differentFirstPage,
              },
              {
                value: 'even',
                label: officeMessage(
                  messages,
                  'document.pageChrome.panel.variant.even',
                ),
                disabled: !chrome.differentOddEvenPages,
              },
            ]}
            onValueChange={setVariant}
          />
        </div>
      </div>
      <DocumentPageChromeRichTextEditor
        key={`${variant}-header`}
        label={officeMessage(messages, 'document.pageChrome.panel.headerLabel', {
          variant: variantLabel,
        })}
        value={chrome[variant].headerHtml}
        onChange={(headerHtml) => updateVariant({ headerHtml })}
      />
      <DocumentPageChromeRichTextEditor
        key={`${variant}-footer`}
        label={officeMessage(messages, 'document.pageChrome.panel.footerLabel', {
          variant: variantLabel,
        })}
        value={chrome[variant].footerHtml}
        onChange={(footerHtml) => updateVariant({ footerHtml })}
      />
      <OfficeCheckbox
        className="work-document-page-number-option"
        ariaLabel={officeMessage(
          messages,
          'document.pageChrome.panel.showPageNumberAria',
          { variant: variantLabel },
        )}
        checked={chrome[variant].showPageNumber}
        onCheckedChange={(showPageNumber) => updateVariant({ showPageNumber })}
      >
        {officeMessage(messages, 'document.pageChrome.panel.showPageNumber')}
      </OfficeCheckbox>
    </fieldset>
  );
}

function emptyPageChrome(content: WorkDocumentPageChromeContent): boolean {
  return !content.headerHtml && !content.footerHtml && !content.showPageNumber;
}
