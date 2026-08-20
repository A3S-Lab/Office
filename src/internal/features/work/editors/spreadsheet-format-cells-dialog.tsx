import { type FormEvent, useId, useMemo, useState } from 'react';
import { Button, Dialog, Tabs } from '../../../design-system/primitives';
import type { SpreadsheetCellFormatPatch } from './spreadsheet-cell-format';
import { SpreadsheetFormatCellsPanel } from './spreadsheet-format-cells-dialog-panels';
import {
  createSpreadsheetFormatCellsDraft,
  spreadsheetFormatCellsDraftErrors,
  spreadsheetFormatCellsPatch,
  spreadsheetFormatCellsTabs,
  type SpreadsheetFormatCellsDialogSource,
  type SpreadsheetFormatCellsTabId,
  type SpreadsheetFormatCellsTouched,
} from './spreadsheet-format-cells-dialog-model';

export function SpreadsheetFormatCellsDialog({
  source,
  restoreFocusTarget,
  onApply,
  onClose,
}: {
  source: SpreadsheetFormatCellsDialogSource;
  restoreFocusTarget: () => HTMLElement | null;
  onApply: (patch: SpreadsheetCellFormatPatch) => boolean;
  onClose: () => void;
}) {
  const [draft, setDraft] = useState(() =>
    createSpreadsheetFormatCellsDraft(source),
  );
  const [touched, setTouched] = useState<SpreadsheetFormatCellsTouched>({});
  const [activeTab, setActiveTab] =
    useState<SpreadsheetFormatCellsTabId>('number');
  const formId = useId();
  const idBase = `spreadsheet-format-cells-${useId().replaceAll(':', '')}`;
  const errors = spreadsheetFormatCellsDraftErrors(draft);
  const patch = spreadsheetFormatCellsPatch(source, draft, touched);
  const hasChanges = Object.keys(patch).length > 0;
  const invalid = Object.keys(errors).length > 0;
  const tabs = useMemo(
    () =>
      spreadsheetFormatCellsTabs.map((tab) => ({
        ...tab,
        tabId: `${idBase}-${tab.id}-tab`,
        panelId: `${idBase}-${tab.id}-panel`,
      })),
    [idBase],
  );

  const submit = (event?: FormEvent<HTMLFormElement>) => {
    event?.preventDefault();
    if (!hasChanges || invalid) return;
    if (onApply(patch)) onClose();
  };

  return (
    <Dialog
      title="设置单元格格式"
      description={formatCellsSelectionDescription(source)}
      className="work-spreadsheet-format-cells-dialog"
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
            disabled={!hasChanges || invalid}
          >
            应用
          </Button>
        </>
      }
    >
      <Tabs
        ariaLabel="单元格格式分类"
        value={activeTab}
        items={tabs}
        variant="line"
        size="compact"
        className="work-spreadsheet-format-cells-tabs"
        onChange={setActiveTab}
      />
      <form id={formId} onSubmit={submit}>
        <SpreadsheetFormatCellsPanel
          activeTab={activeTab}
          idBase={idBase}
          source={source}
          draft={draft}
          errors={errors}
          touched={touched}
          setDraft={setDraft}
          touch={(field) =>
            setTouched((current) => ({ ...current, [field]: true }))
          }
        />
      </form>
    </Dialog>
  );
}

function formatCellsSelectionDescription(
  source: SpreadsheetFormatCellsDialogSource,
): string {
  const rows = source.range.row[1] - source.range.row[0] + 1;
  const columns = source.range.column[1] - source.range.column[0] + 1;
  return `设置当前选区的数字、对齐、字体、边框、填充和保护属性（${rows} 行 × ${columns} 列）。`;
}
