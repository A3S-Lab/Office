import { clampDocumentMargin } from '../work-document-layout';
import { Button } from '../../../design-system/primitives';
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
import { OfficeNumberField, OfficeSelect } from './office-controls';

export function DocumentLayoutPanel({
  layout,
  sectionIndex,
  sectionCount,
  onChange,
  onInsertSection,
  onMergeSection,
  onClose,
}: {
  layout: WorkDocumentSectionLayout;
  sectionIndex: number;
  sectionCount: number;
  onChange: (layout: WorkDocumentSectionLayout) => void;
  onInsertSection: () => void;
  onMergeSection: () => void;
  onClose: () => void;
}) {
  const marginFields: Array<[keyof WorkDocumentMargins, string]> = [
    ['top', '上'],
    ['right', '右'],
    ['bottom', '下'],
    ['left', '左'],
  ];
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
      <div className="work-document-task-pane-body work-document-layout-body">
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
        <DocumentColumnsPanel
          columns={layout.columns}
          onChange={(columns) => update({ columns })}
        />
        <fieldset className="work-document-layout-group">
          <legend>页边距（毫米）</legend>
          <div className="work-document-layout-margin-grid">
            {marginFields.map(([side, label]) => (
              <div className="work-office-field" key={side}>
                <span>{label}</span>
                <OfficeNumberField
                  min={5}
                  max={60}
                  step={1}
                  ariaLabel={`${label}页边距`}
                  value={layout.margins[side]}
                  onValueChange={(value) =>
                    update({
                      margins: {
                        ...layout.margins,
                        [side]: clampDocumentMargin(Number(value)),
                      },
                    })
                  }
                />
              </div>
            ))}
          </div>
        </fieldset>
        <DocumentPageChromePanel
          pageChrome={normalizeDocumentPageChrome(layout.pageChrome, layout)}
          onChange={(pageChrome) =>
            update({
              pageChrome,
              ...documentPageChromeLegacyFields(pageChrome),
            })
          }
        />
        <div className="work-office-field work-document-page-number-option">
          <span>本节页码从</span>
          <OfficeNumberField
            min={1}
            max={9999}
            ariaLabel="起始页码"
            value={Math.max(1, layout.pageNumberStart ?? 1)}
            onValueChange={(value) =>
              update({
                pageNumberStart: Math.min(
                  9999,
                  Math.max(1, Math.round(Number(value) || 1)),
                ),
              })
            }
          />
          <span>开始</span>
        </div>
      </div>
    </DocumentTaskPane>
  );
}
