import { useId, useState } from 'react';
import { Button, Tabs } from '../../../design-system/primitives';
import { officeMessage } from '../../../i18n/office-locale';
import type { OfficeMessageCatalog } from '../../../i18n/office-messages';
import { clampDocumentMargin } from '../work-document-layout';
import {
  documentPageMarginsForLayout,
  type WorkDocumentGutterPosition,
  type WorkDocumentPageMarginMode,
  twipsToMillimeters,
  updateDocumentGutterPosition,
  updateDocumentMirrorMargins,
  updateDocumentPageMarginMillimeters,
  updateDocumentPageMarginMode,
} from '../work-document-page-margins';
import {
  normalizeDocumentPageGeometry,
  pageTwipsToMillimeters,
  resolveDocumentPageSize,
  updateDocumentCustomPageMillimeters,
  updateDocumentPageOrientation,
  updateDocumentPaperSizePreset,
} from '../work-document-page-size';
import {
  documentPageChromeLegacyFields,
  normalizeDocumentPageChrome,
} from '../work-document-page-chrome';
import type {
  WorkDocumentMargins,
  WorkDocumentPaperSize,
  WorkDocumentSectionLayout,
} from '../work-types';
import { DocumentColumnsPanel } from './document-columns-panel';
import { DocumentPageChromePanel } from './document-page-chrome-panel';
import { DocumentTaskPane } from './document-task-pane';
import { CommittedOfficeNumberField, OfficeSelect } from './office-controls';
import { useOfficeMessages } from './office-messages-context';

export type DocumentLayoutPanelTab = 'columns' | 'headerFooter' | 'page';

export function DocumentLayoutPanel({
  layout,
  activeTab: controlledActiveTab,
  sectionIndex,
  sectionCount,
  onActiveTabChange,
  onChange,
  onInsertSection,
  onMergeSection,
  onClose,
}: {
  layout: WorkDocumentSectionLayout;
  activeTab?: DocumentLayoutPanelTab;
  sectionIndex: number;
  sectionCount: number;
  onActiveTabChange?: (tab: DocumentLayoutPanelTab) => void;
  onChange: (layout: WorkDocumentSectionLayout) => void;
  onInsertSection: () => void;
  onMergeSection: () => void;
  onClose: () => void;
}) {
  const messages = useOfficeMessages();
  const tabsId = useId();
  const [internalActiveTab, setInternalActiveTab] =
    useState<DocumentLayoutPanelTab>('page');
  const [customPaperSelected, setCustomPaperSelected] = useState(
    layout.pageSize === 'custom',
  );
  const activeTab = controlledActiveTab ?? internalActiveTab;
  const changeActiveTab = (tab: DocumentLayoutPanelTab) => {
    if (controlledActiveTab === undefined) setInternalActiveTab(tab);
    onActiveTabChange?.(tab);
  };
  const marginFields = documentLayoutMarginFields(messages);
  const pageMargins = documentPageMarginsForLayout(layout);
  const exactPageGeometry = normalizeDocumentPageGeometry(layout.pageGeometry);
  const resolvedPageSize = resolveDocumentPageSize(layout);
  const customPageDimensions = exactPageGeometry
    ? {
        width: pageTwipsToMillimeters(exactPageGeometry.width),
        height: pageTwipsToMillimeters(exactPageGeometry.height),
      }
    : { width: resolvedPageSize.width, height: resolvedPageSize.height };
  const gutterPosition: WorkDocumentGutterPosition = pageMargins.gutterAtTop
    ? 'top'
    : pageMargins.gutterOnRight
      ? 'right'
      : 'left';
  const update = (patch: Partial<WorkDocumentSectionLayout>) =>
    onChange({ ...layout, ...patch });
  const pageNumberStartSuffix = officeMessage(
    messages,
    'document.layout.pageNumber.start',
  );
  return (
    <DocumentTaskPane
      className="work-document-layout-panel"
      title={officeMessage(messages, 'document.layout.title')}
      description={officeMessage(messages, 'document.layout.description', {
        section: String(sectionIndex + 1),
        count: String(sectionCount),
      })}
      closeLabel={officeMessage(messages, 'document.layout.close')}
      onClose={onClose}
    >
      <Tabs
        ariaLabel={officeMessage(messages, 'document.layout.tabsAria')}
        className="work-document-layout-tabs"
        value={activeTab}
        variant="line"
        size="compact"
        items={[
          {
            id: 'page',
            label: officeMessage(messages, 'document.layout.tab.page'),
            tabId: `${tabsId}-page-tab`,
            panelId: `${tabsId}-page-panel`,
          },
          {
            id: 'columns',
            label: officeMessage(messages, 'document.layout.tab.columns'),
            tabId: `${tabsId}-columns-tab`,
            panelId: `${tabsId}-columns-panel`,
          },
          {
            id: 'headerFooter',
            label: officeMessage(messages, 'document.layout.tab.headerFooter'),
            tabId: `${tabsId}-header-footer-tab`,
            panelId: `${tabsId}-header-footer-panel`,
          },
        ]}
        onChange={changeActiveTab}
      />
      <div className="work-document-task-pane-body work-document-layout-body">
        {activeTab === 'page' && (
          <div
            id={`${tabsId}-page-panel`}
            className="work-document-layout-tab-panel"
            role="tabpanel"
            aria-labelledby={`${tabsId}-page-tab`}
          >
            <section
              className="work-document-layout-group"
              aria-label={officeMessage(
                messages,
                'document.layout.paper.sectionAria',
              )}
            >
              <h3>
                {officeMessage(messages, 'document.layout.paper.heading')}
              </h3>
              <div className="work-document-layout-paired-fields">
                <div className="work-office-field">
                  <span>
                    {officeMessage(messages, 'document.layout.paper.size')}
                  </span>
                  <OfficeSelect<WorkDocumentPaperSize>
                    ariaLabel={officeMessage(
                      messages,
                      'document.layout.paper.sizeAria',
                    )}
                    value={customPaperSelected ? 'custom' : layout.pageSize}
                    options={[
                      { value: 'a3', label: 'A3' },
                      { value: 'a4', label: 'A4' },
                      { value: 'a5', label: 'A5' },
                      { value: 'letter', label: 'Letter' },
                      { value: 'legal', label: 'Legal' },
                      { value: 'tabloid', label: 'Tabloid' },
                      {
                        value: 'custom',
                        label: officeMessage(
                          messages,
                          'document.layout.paper.custom',
                        ),
                      },
                    ]}
                    onValueChange={(pageSize) => {
                      if (pageSize === 'custom') {
                        setCustomPaperSelected(true);
                        return;
                      }
                      setCustomPaperSelected(false);
                      onChange(updateDocumentPaperSizePreset(layout, pageSize));
                    }}
                  />
                </div>
                <div className="work-office-field">
                  <span>
                    {officeMessage(
                      messages,
                      'document.layout.paper.orientation',
                    )}
                  </span>
                  <OfficeSelect
                    ariaLabel={officeMessage(
                      messages,
                      'document.layout.paper.orientationAria',
                    )}
                    value={layout.orientation}
                    options={[
                      {
                        value: 'portrait',
                        label: officeMessage(
                          messages,
                          'document.layout.paper.portrait',
                        ),
                      },
                      {
                        value: 'landscape',
                        label: officeMessage(
                          messages,
                          'document.layout.paper.landscape',
                        ),
                      },
                    ]}
                    onValueChange={(orientation) =>
                      onChange(
                        updateDocumentPageOrientation(layout, orientation),
                      )
                    }
                  />
                </div>
              </div>
              {(customPaperSelected || layout.pageSize === 'custom') && (
                <div className="work-document-layout-paired-fields">
                  {(
                    [
                      ['width', 'document.layout.paper.width'],
                      ['height', 'document.layout.paper.height'],
                    ] as const
                  ).map(([dimension, labelKey]) => {
                    const label = officeMessage(messages, labelKey);
                    return (
                      <div className="work-office-field" key={dimension}>
                        <span>
                          {officeMessage(
                            messages,
                            'document.layout.paper.dimensionMm',
                            { label },
                          )}
                        </span>
                        <CommittedOfficeNumberField
                          min={25.4}
                          max={558.8}
                          step={0.1}
                          ariaLabel={officeMessage(
                            messages,
                            'document.layout.paper.dimensionAria',
                            { label },
                          )}
                          value={customPageDimensions[dimension]}
                          normalizeValue={normalizeDocumentPageDimensionInput}
                          onValueCommit={(value) => {
                            setCustomPaperSelected(false);
                            onChange(
                              updateDocumentCustomPageMillimeters(
                                layout,
                                dimension,
                                value,
                              ),
                            );
                          }}
                        />
                      </div>
                    );
                  })}
                </div>
              )}
            </section>
            <fieldset className="work-document-layout-group">
              <legend>
                {officeMessage(messages, 'document.layout.margins.legend')}
              </legend>
              <div className="work-document-layout-margin-grid">
                {marginFields.map(([side, label]) => (
                  <div className="work-office-field" key={side}>
                    <span>{label}</span>
                    <CommittedOfficeNumberField
                      min={5}
                      max={60}
                      step={1}
                      ariaLabel={officeMessage(
                        messages,
                        'document.layout.margins.sideAria',
                        { label },
                      )}
                      value={layout.margins[side]}
                      normalizeValue={normalizeDocumentMarginInput}
                      onValueCommit={(value) =>
                        onChange(
                          updateDocumentPageMarginMillimeters(
                            layout,
                            side,
                            value,
                          ),
                        )
                      }
                    />
                  </div>
                ))}
              </div>
            </fieldset>
            <section
              className="work-document-layout-group"
              aria-label={officeMessage(
                messages,
                'document.layout.advanced.sectionAria',
              )}
            >
              <h3>
                {officeMessage(messages, 'document.layout.advanced.heading')}
              </h3>
              <div className="work-document-layout-margin-grid">
                {(
                  [
                    ['header', 'document.layout.advanced.header'],
                    ['footer', 'document.layout.advanced.footer'],
                    ['gutter', 'document.layout.advanced.gutter'],
                  ] as const
                ).map(([key, labelKey]) => {
                  const label = officeMessage(messages, labelKey);
                  return (
                    <div className="work-office-field" key={key}>
                      <span>{label}</span>
                      <CommittedOfficeNumberField
                        min={0}
                        max={60}
                        step={0.1}
                        ariaLabel={label}
                        value={twipsToMillimeters(pageMargins[key])}
                        normalizeValue={normalizeDocumentPageOffsetInput}
                        onValueCommit={(value) =>
                          onChange(
                            updateDocumentPageMarginMillimeters(
                              layout,
                              key,
                              value,
                            ),
                          )
                        }
                      />
                    </div>
                  );
                })}
              </div>
              <div className="work-document-layout-paired-fields">
                <div className="work-office-field">
                  <span>
                    {officeMessage(
                      messages,
                      'document.layout.advanced.topRule',
                    )}
                  </span>
                  <OfficeSelect<WorkDocumentPageMarginMode>
                    ariaLabel={officeMessage(
                      messages,
                      'document.layout.advanced.topRuleAria',
                    )}
                    value={pageMargins.top < 0 ? 'fromPageEdge' : 'clearChrome'}
                    options={[
                      {
                        value: 'clearChrome',
                        label: officeMessage(
                          messages,
                          'document.layout.advanced.clearHeader',
                        ),
                      },
                      {
                        value: 'fromPageEdge',
                        label: officeMessage(
                          messages,
                          'document.layout.advanced.allowOverlap',
                        ),
                      },
                    ]}
                    onValueChange={(mode) =>
                      onChange(
                        updateDocumentPageMarginMode(layout, 'top', mode),
                      )
                    }
                  />
                </div>
                <div className="work-office-field">
                  <span>
                    {officeMessage(
                      messages,
                      'document.layout.advanced.bottomRule',
                    )}
                  </span>
                  <OfficeSelect<WorkDocumentPageMarginMode>
                    ariaLabel={officeMessage(
                      messages,
                      'document.layout.advanced.bottomRuleAria',
                    )}
                    value={
                      pageMargins.bottom < 0 ? 'fromPageEdge' : 'clearChrome'
                    }
                    options={[
                      {
                        value: 'clearChrome',
                        label: officeMessage(
                          messages,
                          'document.layout.advanced.clearFooter',
                        ),
                      },
                      {
                        value: 'fromPageEdge',
                        label: officeMessage(
                          messages,
                          'document.layout.advanced.allowOverlap',
                        ),
                      },
                    ]}
                    onValueChange={(mode) =>
                      onChange(
                        updateDocumentPageMarginMode(layout, 'bottom', mode),
                      )
                    }
                  />
                </div>
              </div>
              <div className="work-document-layout-paired-fields">
                <div className="work-office-field">
                  <span>
                    {officeMessage(
                      messages,
                      'document.layout.advanced.multiPage',
                    )}
                  </span>
                  <OfficeSelect<'mirrored' | 'normal'>
                    ariaLabel={officeMessage(
                      messages,
                      'document.layout.advanced.mirrorAria',
                    )}
                    value={pageMargins.mirrorMargins ? 'mirrored' : 'normal'}
                    options={[
                      {
                        value: 'normal',
                        label: officeMessage(
                          messages,
                          'document.layout.advanced.normal',
                        ),
                      },
                      {
                        value: 'mirrored',
                        label: officeMessage(
                          messages,
                          'document.layout.advanced.mirrored',
                        ),
                      },
                    ]}
                    onValueChange={(value) =>
                      onChange(
                        updateDocumentMirrorMargins(
                          layout,
                          value === 'mirrored',
                        ),
                      )
                    }
                  />
                </div>
                <div className="work-office-field">
                  <span>
                    {officeMessage(
                      messages,
                      'document.layout.advanced.gutterPosition',
                    )}
                  </span>
                  <OfficeSelect<WorkDocumentGutterPosition>
                    ariaLabel={officeMessage(
                      messages,
                      'document.layout.advanced.gutterPositionAria',
                    )}
                    value={gutterPosition}
                    options={[
                      {
                        value: 'left',
                        label: officeMessage(
                          messages,
                          'document.layout.advanced.gutterLeft',
                        ),
                      },
                      {
                        value: 'right',
                        label: officeMessage(
                          messages,
                          'document.layout.advanced.gutterRight',
                        ),
                      },
                      {
                        value: 'top',
                        label: officeMessage(
                          messages,
                          'document.layout.advanced.gutterTop',
                        ),
                      },
                    ]}
                    onValueChange={(position) =>
                      onChange(updateDocumentGutterPosition(layout, position))
                    }
                  />
                </div>
              </div>
            </section>
          </div>
        )}
        {activeTab === 'columns' && (
          <div
            id={`${tabsId}-columns-panel`}
            className="work-document-layout-tab-panel"
            role="tabpanel"
            aria-labelledby={`${tabsId}-columns-tab`}
          >
            <DocumentColumnsPanel
              columns={layout.columns}
              onChange={(columns) => update({ columns })}
            />
            <section
              className="work-document-layout-group"
              aria-label={officeMessage(
                messages,
                'document.layout.section.sectionAria',
              )}
            >
              <h3>
                {officeMessage(messages, 'document.layout.section.heading')}
              </h3>
              <div className="work-office-field">
                <span>
                  {officeMessage(messages, 'document.layout.section.after')}
                </span>
                <OfficeSelect
                  ariaLabel={officeMessage(
                    messages,
                    'document.layout.section.breakAria',
                  )}
                  value={layout.breakAfter}
                  options={[
                    {
                      value: 'nextPage',
                      label: officeMessage(
                        messages,
                        'document.layout.section.nextPage',
                      ),
                    },
                    {
                      value: 'continuous',
                      label: officeMessage(
                        messages,
                        'document.layout.section.continuous',
                      ),
                    },
                    {
                      value: 'evenPage',
                      label: officeMessage(
                        messages,
                        'document.layout.section.evenPage',
                      ),
                    },
                    {
                      value: 'oddPage',
                      label: officeMessage(
                        messages,
                        'document.layout.section.oddPage',
                      ),
                    },
                    {
                      value: 'nextColumn',
                      label: officeMessage(
                        messages,
                        'document.layout.section.nextColumn',
                      ),
                    },
                  ]}
                  onValueChange={(breakAfter) => update({ breakAfter })}
                />
              </div>
              <div className="work-document-section-actions">
                <Button size="compact" onClick={onInsertSection}>
                  {officeMessage(messages, 'document.layout.section.insert')}
                </Button>
                <Button
                  size="compact"
                  tone="quiet"
                  disabled={sectionIndex === 0}
                  onClick={onMergeSection}
                >
                  {officeMessage(messages, 'document.layout.section.merge')}
                </Button>
              </div>
            </section>
          </div>
        )}
        {activeTab === 'headerFooter' && (
          <div
            id={`${tabsId}-header-footer-panel`}
            className="work-document-layout-tab-panel"
            role="tabpanel"
            aria-labelledby={`${tabsId}-header-footer-tab`}
          >
            <DocumentPageChromePanel
              pageChrome={normalizeDocumentPageChrome(
                layout.pageChrome,
                layout,
              )}
              onChange={(pageChrome) =>
                update({
                  pageChrome,
                  ...documentPageChromeLegacyFields(pageChrome),
                })
              }
            />
            <div className="work-office-field work-document-page-number-option">
              <span>
                {officeMessage(messages, 'document.layout.pageNumber.from')}
              </span>
              <CommittedOfficeNumberField
                min={1}
                max={9999}
                ariaLabel={officeMessage(
                  messages,
                  'document.layout.pageNumber.aria',
                )}
                value={Math.max(1, layout.pageNumberStart ?? 1)}
                normalizeValue={normalizeDocumentPageNumberInput}
                onValueCommit={(pageNumberStart) => update({ pageNumberStart })}
              />
              {pageNumberStartSuffix ? (
                <span>{pageNumberStartSuffix}</span>
              ) : null}
            </div>
          </div>
        )}
      </div>
    </DocumentTaskPane>
  );
}

function documentLayoutMarginFields(
  messages: OfficeMessageCatalog,
): Array<[keyof WorkDocumentMargins, string]> {
  return [
    ['top', officeMessage(messages, 'document.layout.margins.top')],
    ['right', officeMessage(messages, 'document.layout.margins.right')],
    ['bottom', officeMessage(messages, 'document.layout.margins.bottom')],
    ['left', officeMessage(messages, 'document.layout.margins.left')],
  ];
}

function normalizeDocumentMarginInput(value: string): number | null {
  if (!value.trim()) return null;
  const number = Number(value);
  return Number.isFinite(number) ? clampDocumentMargin(number) : null;
}

function normalizeDocumentPageOffsetInput(value: string): number | null {
  if (!value.trim()) return null;
  const number = Number(value);
  return Number.isFinite(number)
    ? Math.min(60, Math.max(0, Math.round(number * 10) / 10))
    : null;
}

function normalizeDocumentPageDimensionInput(value: string): number | null {
  if (!value.trim()) return null;
  const number = Number(value);
  return Number.isFinite(number)
    ? Math.min(558.8, Math.max(25.4, Math.round(number * 10) / 10))
    : null;
}

function normalizeDocumentPageNumberInput(value: string): number | null {
  if (!value.trim()) return null;
  const number = Number(value);
  return Number.isFinite(number)
    ? Math.min(9999, Math.max(1, Math.round(number)))
    : null;
}
