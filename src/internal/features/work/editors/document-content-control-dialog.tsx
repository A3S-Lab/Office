import { Button, Dialog } from '../../../design-system/primitives';
import type {
  WorkDocumentContentControlAppearance,
  WorkDocumentContentControlLock,
  WorkDocumentContentControlProperties,
  WorkDocumentContentControlType,
} from '../work-document-content-control';
import {
  OfficeCheckbox,
  OfficeSelect,
  OfficeTextField,
} from './office-controls';

export interface DocumentContentControlDialogProps {
  editing: boolean;
  properties: WorkDocumentContentControlProperties;
  restoreFocusTarget: () => HTMLElement | null;
  onCancel: () => void;
  onChange: (properties: WorkDocumentContentControlProperties) => void;
  onSubmit: () => void;
}

const typeOptions = [
  { value: 'text', label: '纯文本' },
  { value: 'richText', label: '富文本' },
] as const satisfies readonly {
  value: WorkDocumentContentControlType;
  label: string;
}[];

const lockOptions = [
  { value: 'unlocked', label: '可编辑' },
  { value: 'contentLocked', label: '锁定内容' },
  { value: 'sdtLocked', label: '锁定控件' },
  { value: 'sdtContentLocked', label: '锁定控件和内容' },
] as const satisfies readonly {
  value: WorkDocumentContentControlLock;
  label: string;
}[];

const appearanceOptions = [
  { value: 'boundingBox', label: '边框' },
  { value: 'tags', label: '标签' },
  { value: 'hidden', label: '隐藏边框' },
] as const satisfies readonly {
  value: WorkDocumentContentControlAppearance;
  label: string;
}[];

export function DocumentContentControlDialog({
  editing,
  properties,
  restoreFocusTarget,
  onCancel,
  onChange,
  onSubmit,
}: DocumentContentControlDialogProps) {
  const update = (patch: Partial<WorkDocumentContentControlProperties>) =>
    onChange({ ...properties, ...patch });
  const aliasId = 'document-content-control-alias';
  const tagId = 'document-content-control-tag';

  return (
    <Dialog
      title={editing ? '编辑内容控件' : '插入内容控件'}
      description="仅支持可编辑的内联纯文本或富文本控件；绑定、下拉、日期和重复区域不会被伪装成普通文本。"
      className="work-document-content-control-dialog"
      restoreFocusTarget={restoreFocusTarget}
      onClose={onCancel}
      footer={
        <>
          <Button tone="quiet" onClick={onCancel}>
            取消
          </Button>
          <Button tone="primary" onClick={onSubmit}>
            {editing ? '应用' : '插入控件'}
          </Button>
        </>
      }
    >
      <div className="work-document-content-control-dialog-fields">
        <label htmlFor={aliasId}>
          <span>显示名称</span>
          <OfficeTextField
            id={aliasId}
            data-autofocus
            aria-label="内容控件显示名称"
            value={properties.alias}
            maxLength={255}
            placeholder="例如：客户名称"
            onChange={(event) => update({ alias: event.target.value })}
            onKeyDown={(event) => {
              if (event.key !== 'Enter' || event.nativeEvent.isComposing)
                return;
              event.preventDefault();
              onSubmit();
            }}
          />
        </label>
        <label htmlFor={tagId}>
          <span>程序标签</span>
          <OfficeTextField
            id={tagId}
            aria-label="内容控件程序标签"
            value={properties.tag}
            maxLength={255}
            placeholder="可选，供文档自动化识别"
            onChange={(event) => update({ tag: event.target.value })}
          />
        </label>
      </div>
      <div className="work-document-content-control-dialog-selects">
        <div className="work-document-content-control-dialog-field">
          <span>控件类型</span>
          <OfficeSelect<WorkDocumentContentControlType>
            ariaLabel="内容控件类型"
            value={properties.type}
            options={typeOptions}
            onValueChange={(type) => update({ type })}
          />
        </div>
        <div className="work-document-content-control-dialog-field">
          <span>锁定方式</span>
          <OfficeSelect<WorkDocumentContentControlLock>
            ariaLabel="内容控件锁定方式"
            value={properties.lock}
            options={lockOptions}
            onValueChange={(lock) => update({ lock })}
          />
        </div>
        <div className="work-document-content-control-dialog-field">
          <span>外观</span>
          <OfficeSelect<WorkDocumentContentControlAppearance>
            ariaLabel="内容控件外观"
            value={properties.appearance}
            options={appearanceOptions}
            onValueChange={(appearance) => update({ appearance })}
          />
        </div>
      </div>
      <div className="work-document-content-control-dialog-options">
        <OfficeCheckbox
          ariaLabel="允许多行文字"
          checked={properties.multiLine}
          onCheckedChange={(multiLine) => update({ multiLine })}
        >
          允许多行文字
        </OfficeCheckbox>
        <label className="work-document-content-control-color">
          <span>控件颜色</span>
          <input
            type="color"
            aria-label="内容控件颜色"
            value={properties.color ?? '#2f6fed'}
            onChange={(event) => update({ color: event.target.value })}
          />
          <button
            type="button"
            className="work-document-content-control-color-reset"
            onClick={() => update({ color: null })}
          >
            使用主题色
          </button>
        </label>
      </div>
    </Dialog>
  );
}
