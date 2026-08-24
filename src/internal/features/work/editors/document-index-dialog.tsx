import { Button, Dialog } from '../../../design-system/primitives';
import type {
  WorkDocumentIndexFormat,
  WorkDocumentIndexLeader,
  WorkDocumentIndexOptions,
} from '../work-document-index';
import {
  OfficeCheckbox,
  OfficeSelect,
  type OfficeSelectOption,
} from './office-controls';

export interface DocumentIndexDialogProps {
  editing: boolean;
  options: WorkDocumentIndexOptions;
  restoreFocusTarget: () => HTMLElement | null;
  onCancel: () => void;
  onOptionsChange: (options: WorkDocumentIndexOptions) => void;
  onSubmit: () => void;
}

export function DocumentIndexDialog({
  editing,
  options,
  restoreFocusTarget,
  onCancel,
  onOptionsChange,
  onSubmit,
}: DocumentIndexDialogProps) {
  const update = (patch: Partial<WorkDocumentIndexOptions>) =>
    onOptionsChange({ ...options, ...patch });

  return (
    <Dialog
      title={editing ? '自定义索引' : '插入索引'}
      description="从已标记的索引项生成可更新索引，并保留原生 DOCX INDEX 域。"
      className="work-document-index-dialog"
      restoreFocusTarget={restoreFocusTarget}
      onClose={onCancel}
      footer={
        <>
          <Button tone="quiet" onClick={onCancel}>
            取消
          </Button>
          <Button tone="primary" onClick={onSubmit}>
            {editing ? '应用' : '插入索引'}
          </Button>
        </>
      }
    >
      <div className="work-document-index-dialog-grid">
        <div className="work-document-index-dialog-field">
          <span>栏数</span>
          <OfficeSelect
            initialFocus
            ariaLabel="索引栏数"
            value={String(options.columns)}
            options={indexColumnOptions}
            onValueChange={(columns) => update({ columns: Number(columns) })}
          />
        </div>
        <div className="work-document-index-dialog-field">
          <span>次索引项布局</span>
          <OfficeSelect<WorkDocumentIndexFormat>
            ariaLabel="索引布局"
            value={options.format}
            options={indexFormatOptions}
            onValueChange={(format) => update({ format })}
          />
        </div>
      </div>
      <div className="work-document-index-dialog-options">
        <OfficeCheckbox
          ariaLabel="页码右对齐"
          checked={options.rightAlignPageNumbers}
          onCheckedChange={(rightAlignPageNumbers) =>
            update({ rightAlignPageNumbers })
          }
        >
          页码右对齐
        </OfficeCheckbox>
        <div className="work-document-index-dialog-field">
          <span>前导符</span>
          <OfficeSelect<WorkDocumentIndexLeader>
            ariaLabel="索引前导符"
            value={options.leader}
            options={indexLeaderOptions}
            disabled={!options.rightAlignPageNumbers}
            onValueChange={(leader) => update({ leader })}
          />
        </div>
      </div>
      <fieldset
        className="work-document-index-dialog-preview"
        aria-label="索引预览"
        data-index-format={options.format}
        data-index-leader={options.leader}
        data-index-right-align-page-numbers={String(
          options.rightAlignPageNumbers,
        )}
      >
        <legend>索引</legend>
        <span>
          Architecture <i aria-hidden="true" /> 1, 3
        </span>
        <span className="nested">
          Runtime <i aria-hidden="true" /> 5
        </span>
      </fieldset>
    </Dialog>
  );
}

const indexColumnOptions = [
  { value: '1', label: '一栏' },
  { value: '2', label: '两栏' },
  { value: '3', label: '三栏' },
  { value: '4', label: '四栏' },
] as const satisfies readonly OfficeSelectOption[];

const indexFormatOptions = [
  { value: 'indented', label: '缩进式' },
  { value: 'run-in', label: '连续式' },
] as const satisfies readonly OfficeSelectOption<WorkDocumentIndexFormat>[];

const indexLeaderOptions = [
  { value: 'dot', label: '点线（……）' },
  { value: 'dash', label: '短横线（----）' },
  { value: 'underline', label: '下划线（____）' },
  { value: 'none', label: '无' },
] as const satisfies readonly OfficeSelectOption<WorkDocumentIndexLeader>[];
