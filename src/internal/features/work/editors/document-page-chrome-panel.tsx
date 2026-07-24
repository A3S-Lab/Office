import { useState } from 'react';
import {
  normalizeDocumentPageChrome,
  updateDocumentPageChromeVariant,
} from '../work-document-page-chrome';
import type {
  WorkDocumentPageChrome,
  WorkDocumentPageChromeContent,
  WorkDocumentPageChromeVariant,
} from '../work-types';
import { OfficeCheckbox, OfficeSelect } from './office-controls';
import { DocumentPageChromeRichTextEditor } from './document-page-chrome-editor';

export function DocumentPageChromePanel({
  pageChrome,
  onChange,
}: {
  pageChrome: WorkDocumentPageChrome;
  onChange: (pageChrome: WorkDocumentPageChrome) => void;
}) {
  const chrome = normalizeDocumentPageChrome(pageChrome);
  const [variant, setVariant] =
    useState<WorkDocumentPageChromeVariant>('default');
  const label =
    variant === 'first' ? '首页' : variant === 'even' ? '偶数页' : '默认页';
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
      <legend>页眉和页脚</legend>
      <div className="work-document-page-chrome-options">
        <OfficeCheckbox
          ariaLabel="首页页眉页脚不同"
          checked={chrome.differentFirstPage}
          onCheckedChange={toggleFirstPage}
        >
          首页不同
        </OfficeCheckbox>
        <OfficeCheckbox
          ariaLabel="奇偶页页眉页脚不同"
          checked={chrome.differentOddEvenPages}
          onCheckedChange={toggleOddEvenPages}
        >
          奇偶页不同
        </OfficeCheckbox>
        <div className="work-office-field">
          <span>编辑</span>
          <OfficeSelect
            ariaLabel="页眉页脚页面类型"
            value={variant}
            options={[
              { value: 'default', label: '默认页' },
              {
                value: 'first',
                label: '首页',
                disabled: !chrome.differentFirstPage,
              },
              {
                value: 'even',
                label: '偶数页',
                disabled: !chrome.differentOddEvenPages,
              },
            ]}
            onValueChange={setVariant}
          />
        </div>
      </div>
      <DocumentPageChromeRichTextEditor
        key={`${variant}-header`}
        label={`${label}页眉`}
        value={chrome[variant].headerHtml}
        onChange={(headerHtml) => updateVariant({ headerHtml })}
      />
      <DocumentPageChromeRichTextEditor
        key={`${variant}-footer`}
        label={`${label}页脚`}
        value={chrome[variant].footerHtml}
        onChange={(footerHtml) => updateVariant({ footerHtml })}
      />
      <OfficeCheckbox
        className="work-document-page-number-option"
        ariaLabel={`${label}显示页码`}
        checked={chrome[variant].showPageNumber}
        onCheckedChange={(showPageNumber) => updateVariant({ showPageNumber })}
      >
        在本页面类型的页脚中显示页码
      </OfficeCheckbox>
    </fieldset>
  );
}

function emptyPageChrome(content: WorkDocumentPageChromeContent): boolean {
  return !content.headerHtml && !content.footerHtml && !content.showPageNumber;
}
