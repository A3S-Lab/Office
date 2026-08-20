import { type FormEvent, useId, useState } from 'react';
import { Button, Dialog } from '../../../design-system/primitives';
import { OfficeCheckbox } from './office-controls';
import {
  spreadsheetPasteContentOptions,
  spreadsheetPasteOperationOptions,
  spreadsheetPasteSpecialModeAvailable,
  type SpreadsheetPasteContent,
  type SpreadsheetPasteSpecialOptions,
} from './spreadsheet-paste-special';
import type { SpreadsheetPasteSpecialDialogSource } from './use-spreadsheet-clipboard';

export function SpreadsheetPasteSpecialDialog({
  source,
  restoreFocusTarget,
  onApply,
  onClose,
  onValidate,
}: {
  source: SpreadsheetPasteSpecialDialogSource;
  restoreFocusTarget: () => HTMLElement | null;
  onApply: (options: SpreadsheetPasteSpecialOptions) => boolean;
  onClose: () => void;
  onValidate: (options: SpreadsheetPasteSpecialOptions) => string | null;
}) {
  const [options, setOptions] = useState<SpreadsheetPasteSpecialOptions>({
    content: 'all',
    operation: 'none',
    skipBlanks: false,
    transpose: false,
  });
  const formId = useId();
  const validationError = onValidate(options);
  const sourceKind =
    source.snapshot.kind === 'rich' ? 'A3S 富剪贴板' : '纯文本';
  const sourceSize = `${source.snapshot.rowCount} 行 × ${source.snapshot.columnCount} 列`;

  const selectContent = (content: SpreadsheetPasteContent) => {
    setOptions((current) => ({
      ...current,
      content,
      operation: spreadsheetPasteContentSupportsOperation(content)
        ? current.operation
        : 'none',
      skipBlanks: content === 'column-widths' ? false : current.skipBlanks,
      transpose: content === 'column-widths' ? false : current.transpose,
    }));
  };
  const submit = (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    if (!validationError && onApply(options)) onClose();
  };

  return (
    <Dialog
      title="选择性粘贴"
      description="精确选择要保留的单元格内容、格式和粘贴运算。"
      className="work-spreadsheet-paste-special-dialog"
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
            粘贴
          </Button>
        </>
      }
    >
      <section
        className="work-spreadsheet-paste-special-source"
        aria-label="剪贴板摘要"
      >
        <span>{sourceKind}</span>
        <strong>{sourceSize}</strong>
      </section>
      <form id={formId} onSubmit={submit}>
        <fieldset className="work-spreadsheet-paste-special-modes">
          <legend>粘贴内容</legend>
          <div>
            {spreadsheetPasteContentOptions.map((option) => {
              const available = spreadsheetPasteSpecialModeAvailable(
                source.snapshot,
                option.value,
              );
              return (
                <label
                  key={option.value}
                  className={options.content === option.value ? 'selected' : ''}
                >
                  <input
                    type="radio"
                    name="spreadsheet-paste-content"
                    value={option.value}
                    checked={options.content === option.value}
                    disabled={!available}
                    onChange={() => selectContent(option.value)}
                  />
                  <span aria-hidden="true" />
                  <span>{option.label}</span>
                  {!available && <small>仅限同一编辑器复制</small>}
                </label>
              );
            })}
          </div>
        </fieldset>

        <fieldset className="work-spreadsheet-paste-special-operations">
          <legend>运算</legend>
          <div>
            {spreadsheetPasteOperationOptions.map((operation) => (
              <label
                key={operation.value}
                className={
                  options.operation === operation.value ? 'selected' : ''
                }
              >
                <input
                  type="radio"
                  name="spreadsheet-paste-operation"
                  value={operation.value}
                  checked={options.operation === operation.value}
                  disabled={
                    operation.value !== 'none' &&
                    !spreadsheetPasteContentSupportsOperation(options.content)
                  }
                  onChange={() =>
                    setOptions((current) => ({
                      ...current,
                      operation: operation.value,
                    }))
                  }
                />
                <span>{operation.label}</span>
              </label>
            ))}
          </div>
        </fieldset>

        <div className="work-spreadsheet-paste-special-options">
          <OfficeCheckbox
            ariaLabel="跳过空白单元格"
            checked={options.skipBlanks}
            disabled={options.content === 'column-widths'}
            onCheckedChange={(skipBlanks) =>
              setOptions((current) => ({ ...current, skipBlanks }))
            }
          >
            跳过空白单元格
          </OfficeCheckbox>
          <OfficeCheckbox
            ariaLabel="转置行列"
            checked={options.transpose}
            disabled={options.content === 'column-widths'}
            onCheckedChange={(transpose) =>
              setOptions((current) => ({ ...current, transpose }))
            }
          >
            转置
          </OfficeCheckbox>
        </div>

        {validationError && (
          <p className="work-spreadsheet-paste-special-error" role="alert">
            {validationError}
          </p>
        )}
      </form>
    </Dialog>
  );
}

function spreadsheetPasteContentSupportsOperation(
  content: SpreadsheetPasteContent,
): boolean {
  return !['formats', 'comments', 'validation', 'column-widths'].includes(
    content,
  );
}
