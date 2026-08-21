import { Table2 } from 'lucide-react';
import { type FormEvent, useId, useState } from 'react';
import { Button, Dialog, Field } from '../../../design-system/primitives';
import { OfficeCheckbox } from './office-controls';
import type {
  SpreadsheetTableDialogSource,
  SpreadsheetTableDialogValue,
} from './spreadsheet-table';

export function SpreadsheetTableDialog({
  source,
  restoreFocusTarget,
  onApply,
  onClose,
  onValidate,
}: {
  source: SpreadsheetTableDialogSource;
  restoreFocusTarget: () => HTMLElement | null;
  onApply: (value: SpreadsheetTableDialogValue) => boolean;
  onClose: () => void;
  onValidate: (value: SpreadsheetTableDialogValue) => string | null;
}) {
  const [value, setValue] = useState(source.value);
  const [touched, setTouched] = useState(false);
  const formId = useId();
  const validationError = onValidate(value);
  const visibleError = touched ? validationError : null;

  const submit = (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    setTouched(true);
    if (validationError) return;
    if (onApply(value)) onClose();
  };

  return (
    <Dialog
      title="创建表格"
      description={source.sheetName}
      className="work-spreadsheet-table-dialog"
      restoreFocusTarget={restoreFocusTarget}
      onClose={onClose}
      footer={
        <>
          <Button tone="quiet" onClick={onClose}>
            取消
          </Button>
          <Button
            tone="primary"
            type="submit"
            form={formId}
            disabled={Boolean(validationError)}
          >
            确定
          </Button>
        </>
      }
    >
      <form id={formId} onSubmit={submit}>
        <div className="work-spreadsheet-table-scope">
          <span aria-hidden="true">
            <Table2 size={18} />
          </span>
          <div>
            <strong>{source.name}</strong>
            <small>将连续数据区域转换为原生工作表表格</small>
          </div>
        </div>
        <Field
          label="表格区域"
          required
          error={visibleError ?? undefined}
          description="输入一个连续区域，例如 A1:C20。"
        >
          <input
            type="text"
            autoCapitalize="none"
            autoFocus
            spellCheck={false}
            value={value.rangeReference}
            onBlur={() => setTouched(true)}
            onChange={(event) => {
              const rangeReference = event.currentTarget.value;
              setValue((current) => ({
                ...current,
                rangeReference,
              }));
              setTouched(true);
            }}
          />
        </Field>
        <OfficeCheckbox
          ariaLabel="表包含标题"
          checked={value.headerRow}
          onCheckedChange={(headerRow) =>
            setValue((current) => ({ ...current, headerRow }))
          }
        >
          表包含标题
        </OfficeCheckbox>
      </form>
    </Dialog>
  );
}
