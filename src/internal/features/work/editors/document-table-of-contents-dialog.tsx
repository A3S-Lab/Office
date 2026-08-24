import { Button, Dialog } from '../../../design-system/primitives';
import type {
  WorkDocumentTableOfContentsLeader,
  WorkDocumentTableOfContentsOptions,
} from '../work-document-table-of-contents';
import {
  OfficeCheckbox,
  OfficeSelect,
  type OfficeSelectOption,
} from './office-controls';

export interface DocumentTableOfContentsDialogProps {
  editing: boolean;
  options: WorkDocumentTableOfContentsOptions;
  restoreFocusTarget: () => HTMLElement | null;
  onCancel: () => void;
  onOptionsChange: (options: WorkDocumentTableOfContentsOptions) => void;
  onSubmit: () => void;
}

export function DocumentTableOfContentsDialog({
  editing,
  options,
  restoreFocusTarget,
  onCancel,
  onOptionsChange,
  onSubmit,
}: DocumentTableOfContentsDialogProps) {
  const updateOptions = (update: Partial<WorkDocumentTableOfContentsOptions>) =>
    onOptionsChange({ ...options, ...update });

  return (
    <Dialog
      title={editing ? '自定义目录' : '插入目录'}
      description="从文档标题生成可更新目录，并保留原生 DOCX 目录域。"
      className="work-document-table-of-contents-dialog"
      restoreFocusTarget={restoreFocusTarget}
      onClose={onCancel}
      footer={
        <>
          <Button tone="quiet" onClick={onCancel}>
            取消
          </Button>
          <Button tone="primary" onClick={onSubmit}>
            {editing ? '应用' : '插入目录'}
          </Button>
        </>
      }
    >
      <div className="work-document-table-of-contents-dialog-grid">
        <div className="work-document-table-of-contents-dialog-field">
          <span>起始标题级别</span>
          <OfficeSelect
            initialFocus
            ariaLabel="起始标题级别"
            value={String(options.minLevel)}
            options={tableOfContentsLevelOptions}
            onValueChange={(value) => {
              const minLevel = Number(value);
              updateOptions({
                minLevel,
                maxLevel: Math.max(minLevel, options.maxLevel),
              });
            }}
          />
        </div>
        <div className="work-document-table-of-contents-dialog-field">
          <span>结束标题级别</span>
          <OfficeSelect
            ariaLabel="结束标题级别"
            value={String(options.maxLevel)}
            options={tableOfContentsLevelOptions.map((option) => ({
              ...option,
              disabled: Number(option.value) < options.minLevel,
            }))}
            onValueChange={(value) =>
              updateOptions({ maxLevel: Number(value) })
            }
          />
        </div>
      </div>
      <div className="work-document-table-of-contents-dialog-options">
        <OfficeCheckbox
          ariaLabel="目录项使用超链接"
          checked={options.hyperlinks}
          onCheckedChange={(hyperlinks) => updateOptions({ hyperlinks })}
        >
          目录项使用超链接
        </OfficeCheckbox>
        <OfficeCheckbox
          ariaLabel="显示页码"
          checked={options.showPageNumbers}
          onCheckedChange={(showPageNumbers) =>
            updateOptions({
              showPageNumbers,
              rightAlignPageNumbers:
                showPageNumbers && options.rightAlignPageNumbers,
            })
          }
        >
          显示页码
        </OfficeCheckbox>
        <OfficeCheckbox
          ariaLabel="页码右对齐"
          checked={options.rightAlignPageNumbers}
          disabled={!options.showPageNumbers}
          onCheckedChange={(rightAlignPageNumbers) =>
            updateOptions({ rightAlignPageNumbers })
          }
        >
          页码右对齐
        </OfficeCheckbox>
      </div>
      <div className="work-document-table-of-contents-dialog-leader">
        <span>前导符</span>
        <OfficeSelect<WorkDocumentTableOfContentsLeader>
          ariaLabel="目录前导符"
          value={options.leader}
          options={tableOfContentsLeaderOptions}
          disabled={!options.showPageNumbers || !options.rightAlignPageNumbers}
          onValueChange={(leader) => updateOptions({ leader })}
        />
      </div>
      <fieldset
        className="work-document-table-of-contents-dialog-preview"
        aria-label="目录预览"
        data-toc-leader={options.leader}
        data-toc-right-align-page-numbers={String(
          options.rightAlignPageNumbers,
        )}
      >
        <legend>目录</legend>
        <span>
          项目概述 <i aria-hidden="true" />
          {options.showPageNumbers ? '1' : ''}
        </span>
        <span className="nested">
          实施计划 <i aria-hidden="true" />
          {options.showPageNumbers ? '3' : ''}
        </span>
      </fieldset>
    </Dialog>
  );
}

const tableOfContentsLevelOptions = Array.from({ length: 9 }, (_, index) => ({
  value: String(index + 1),
  label: `标题 ${index + 1}`,
})) satisfies readonly OfficeSelectOption[];

const tableOfContentsLeaderOptions = [
  { value: 'dot', label: '点线（……）' },
  { value: 'dash', label: '短横线（----）' },
  { value: 'underline', label: '下划线（____）' },
  { value: 'none', label: '无' },
] as const satisfies readonly OfficeSelectOption<WorkDocumentTableOfContentsLeader>[];
