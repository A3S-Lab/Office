import { Button, Dialog } from '../../../design-system/primitives';
import { useId } from 'react';
import type { WorkDocumentIndexEntryDraft } from '../work-document-index';
import { OfficeCheckbox, OfficeTextField } from './office-controls';

export interface DocumentIndexEntryDialogProps {
  editing: boolean;
  value: WorkDocumentIndexEntryDraft;
  restoreFocusTarget: () => HTMLElement | null;
  onCancel: () => void;
  onChange: (value: WorkDocumentIndexEntryDraft) => void;
  onSubmit: () => void;
}

export function DocumentIndexEntryDialog({
  editing,
  value,
  restoreFocusTarget,
  onCancel,
  onChange,
  onSubmit,
}: DocumentIndexEntryDialogProps) {
  const update = (patch: Partial<WorkDocumentIndexEntryDraft>) =>
    onChange({ ...value, ...patch });
  const crossReferenceEnabled = Boolean(value.crossReference);
  const mainEntryId = useId();
  const subEntryId = useId();
  const crossReferenceId = useId();

  return (
    <Dialog
      title={editing ? '编辑索引项' : '标记索引项'}
      description="标记主索引项、次索引项或交叉引用，并保留原生 DOCX XE 域。"
      className="work-document-index-entry-dialog"
      restoreFocusTarget={restoreFocusTarget}
      onClose={onCancel}
      footer={
        <>
          <Button tone="quiet" onClick={onCancel}>
            取消
          </Button>
          <Button
            tone="primary"
            disabled={!value.mainEntry.trim()}
            onClick={onSubmit}
          >
            {editing ? '应用' : '标记'}
          </Button>
        </>
      }
    >
      <div className="work-document-index-entry-dialog-fields">
        <label htmlFor={mainEntryId}>
          <span>主索引项</span>
          <OfficeTextField
            id={mainEntryId}
            data-autofocus
            aria-label="主索引项"
            value={value.mainEntry}
            maxLength={240}
            onChange={(event) => update({ mainEntry: event.target.value })}
            onKeyDown={(event) => {
              if (event.key !== 'Enter' || !value.mainEntry.trim()) return;
              event.preventDefault();
              onSubmit();
            }}
          />
        </label>
        <label htmlFor={subEntryId}>
          <span>次索引项</span>
          <OfficeTextField
            id={subEntryId}
            aria-label="次索引项"
            value={value.subEntry}
            maxLength={240}
            placeholder="可选"
            onChange={(event) => update({ subEntry: event.target.value })}
          />
        </label>
      </div>
      <div className="work-document-index-entry-dialog-mode">
        <OfficeCheckbox
          ariaLabel="使用交叉引用"
          checked={crossReferenceEnabled}
          onCheckedChange={(enabled) =>
            update({
              crossReference: enabled ? value.crossReference || '参见' : '',
              pageBold: enabled ? false : value.pageBold,
              pageItalic: enabled ? false : value.pageItalic,
            })
          }
        >
          使用交叉引用
        </OfficeCheckbox>
        {crossReferenceEnabled && (
          <label htmlFor={crossReferenceId}>
            <span>引用目标</span>
            <OfficeTextField
              id={crossReferenceId}
              aria-label="交叉引用目标"
              value={value.crossReference}
              maxLength={240}
              placeholder="例如：Architecture"
              onChange={(event) =>
                update({ crossReference: event.target.value })
              }
            />
          </label>
        )}
      </div>
      <fieldset
        className="work-document-index-entry-dialog-page-style"
        disabled={crossReferenceEnabled}
      >
        <legend>当前页码格式</legend>
        <OfficeCheckbox
          ariaLabel="页码加粗"
          checked={value.pageBold}
          disabled={crossReferenceEnabled}
          onCheckedChange={(pageBold) => update({ pageBold })}
        >
          页码加粗
        </OfficeCheckbox>
        <OfficeCheckbox
          ariaLabel="页码倾斜"
          checked={value.pageItalic}
          disabled={crossReferenceEnabled}
          onCheckedChange={(pageItalic) => update({ pageItalic })}
        >
          页码倾斜
        </OfficeCheckbox>
      </fieldset>
    </Dialog>
  );
}
