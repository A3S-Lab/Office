import { useId, useState } from 'react';
import { Button, Tabs } from '../../../design-system/primitives';
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
  documentPageChromeLegacyFields,
  normalizeDocumentPageChrome,
} from '../work-document-page-chrome';
import type {
  WorkDocumentMargins,
  WorkDocumentSectionLayout,
} from '../work-types';
import { DocumentColumnsPanel } from './document-columns-panel';
import { DocumentPageChromePanel } from './document-page-chrome-panel';
import { DocumentTaskPane } from './document-task-pane';
import { CommittedOfficeNumberField, OfficeSelect } from './office-controls';

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
  const tabsId = useId();
  const [internalActiveTab, setInternalActiveTab] =
    useState<DocumentLayoutPanelTab>('page');
  const activeTab = controlledActiveTab ?? internalActiveTab;
  const changeActiveTab = (tab: DocumentLayoutPanelTab) => {
    if (controlledActiveTab === undefined) setInternalActiveTab(tab);
    onActiveTabChange?.(tab);
  };
  const marginFields: Array<[keyof WorkDocumentMargins, string]> = [
    ['top', '上'],
    ['right', '右'],
    ['bottom', '下'],
    ['left', '左'],
  ];
  const pageMargins = documentPageMarginsForLayout(layout);
  const gutterPosition: WorkDocumentGutterPosition = pageMargins.gutterAtTop
    ? 'top'
    : pageMargins.gutterOnRight
      ? 'right'
      : 'left';
  const update = (patch: Partial<WorkDocumentSectionLayout>) =>
    onChange({ ...layout, ...patch });
  return (
    <DocumentTaskPane
      className="work-document-layout-panel"
      title="页面设置"
      description={`第 ${sectionIndex + 1} 节 · 共 ${sectionCount} 节`}
      closeLabel="关闭页面设置"
      onClose={onClose}
    >
      <Tabs
        ariaLabel="页面设置类别"
        className="work-document-layout-tabs"
        value={activeTab}
        variant="line"
        size="compact"
        items={[
          {
            id: 'page',
            label: '页面',
            tabId: `${tabsId}-page-tab`,
            panelId: `${tabsId}-page-panel`,
          },
          {
            id: 'columns',
            label: '分栏与分节',
            tabId: `${tabsId}-columns-tab`,
            panelId: `${tabsId}-columns-panel`,
          },
          {
            id: 'headerFooter',
            label: '页眉页脚',
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
            <section className="work-document-layout-group" aria-label="纸张">
              <h3>纸张</h3>
              <div className="work-document-layout-paired-fields">
                <div className="work-office-field">
                  <span>大小</span>
                  <OfficeSelect
                    ariaLabel="纸张大小"
                    value={layout.pageSize}
                    options={[
                      { value: 'a4', label: 'A4' },
                      { value: 'letter', label: 'Letter' },
                    ]}
                    onValueChange={(pageSize) => update({ pageSize })}
                  />
                </div>
                <div className="work-office-field">
                  <span>方向</span>
                  <OfficeSelect
                    ariaLabel="页面方向"
                    value={layout.orientation}
                    options={[
                      { value: 'portrait', label: '纵向' },
                      { value: 'landscape', label: '横向' },
                    ]}
                    onValueChange={(orientation) => update({ orientation })}
                  />
                </div>
              </div>
            </section>
            <fieldset className="work-document-layout-group">
              <legend>页边距（毫米）</legend>
              <div className="work-document-layout-margin-grid">
                {marginFields.map(([side, label]) => (
                  <div className="work-office-field" key={side}>
                    <span>{label}</span>
                    <CommittedOfficeNumberField
                      min={5}
                      max={60}
                      step={1}
                      ariaLabel={`${label}页边距`}
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
              aria-label="高级页边距"
            >
              <h3>高级页边距</h3>
              <div className="work-document-layout-margin-grid">
                {(
                  [
                    ['header', '页眉距顶端'],
                    ['footer', '页脚距底端'],
                    ['gutter', '装订线'],
                  ] as const
                ).map(([key, label]) => (
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
                ))}
              </div>
              <div className="work-document-layout-paired-fields">
                <div className="work-office-field">
                  <span>上边距规则</span>
                  <OfficeSelect<WorkDocumentPageMarginMode>
                    ariaLabel="上边距与页眉关系"
                    value={pageMargins.top < 0 ? 'fromPageEdge' : 'clearChrome'}
                    options={[
                      { value: 'clearChrome', label: '避让页眉' },
                      { value: 'fromPageEdge', label: '允许重叠' },
                    ]}
                    onValueChange={(mode) =>
                      onChange(
                        updateDocumentPageMarginMode(layout, 'top', mode),
                      )
                    }
                  />
                </div>
                <div className="work-office-field">
                  <span>下边距规则</span>
                  <OfficeSelect<WorkDocumentPageMarginMode>
                    ariaLabel="下边距与页脚关系"
                    value={
                      pageMargins.bottom < 0 ? 'fromPageEdge' : 'clearChrome'
                    }
                    options={[
                      { value: 'clearChrome', label: '避让页脚' },
                      { value: 'fromPageEdge', label: '允许重叠' },
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
                  <span>多页方式</span>
                  <OfficeSelect<'mirrored' | 'normal'>
                    ariaLabel="镜像页边距"
                    value={pageMargins.mirrorMargins ? 'mirrored' : 'normal'}
                    options={[
                      { value: 'normal', label: '普通' },
                      { value: 'mirrored', label: '镜像页边距' },
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
                  <span>装订线位置</span>
                  <OfficeSelect<WorkDocumentGutterPosition>
                    ariaLabel="装订线位置"
                    value={gutterPosition}
                    options={[
                      { value: 'left', label: '左侧' },
                      { value: 'right', label: '右侧' },
                      { value: 'top', label: '顶部' },
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
            <section className="work-document-layout-group" aria-label="分节">
              <h3>分节</h3>
              <div className="work-office-field">
                <span>本节之后</span>
                <OfficeSelect
                  ariaLabel="分节方式"
                  value={layout.breakAfter}
                  options={[
                    { value: 'nextPage', label: '下一页' },
                    { value: 'continuous', label: '连续' },
                    { value: 'evenPage', label: '下一偶数页' },
                    { value: 'oddPage', label: '下一奇数页' },
                    { value: 'nextColumn', label: '下一栏（预览按连续节）' },
                  ]}
                  onValueChange={(breakAfter) => update({ breakAfter })}
                />
              </div>
              <div className="work-document-section-actions">
                <Button size="compact" onClick={onInsertSection}>
                  插入新节
                </Button>
                <Button
                  size="compact"
                  tone="quiet"
                  disabled={sectionIndex === 0}
                  onClick={onMergeSection}
                >
                  与上一节合并
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
              <span>本节页码从</span>
              <CommittedOfficeNumberField
                min={1}
                max={9999}
                ariaLabel="起始页码"
                value={Math.max(1, layout.pageNumberStart ?? 1)}
                normalizeValue={normalizeDocumentPageNumberInput}
                onValueCommit={(pageNumberStart) => update({ pageNumberStart })}
              />
              <span>开始</span>
            </div>
          </div>
        )}
      </div>
    </DocumentTaskPane>
  );
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

function normalizeDocumentPageNumberInput(value: string): number | null {
  if (!value.trim()) return null;
  const number = Number(value);
  return Number.isFinite(number)
    ? Math.min(9999, Math.max(1, Math.round(number)))
    : null;
}
