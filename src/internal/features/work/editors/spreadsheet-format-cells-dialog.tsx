import { type FormEvent, useId, useMemo, useState } from 'react';
import { officeMessage } from '../../../i18n/office-locale';
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
import {
  defaultSpreadsheetFormatCellsOpenIntent,
  type SpreadsheetFormatCellsOpenIntent,
} from './spreadsheet-format-cells-intent';
import { useOfficeMessages } from './office-messages-context';

export function SpreadsheetFormatCellsDialog({
  source,
  restoreFocusTarget,
  openIntent = defaultSpreadsheetFormatCellsOpenIntent,
  onApply,
  onClose,
}: {
  source: SpreadsheetFormatCellsDialogSource;
  restoreFocusTarget: () => HTMLElement | null;
  openIntent?: SpreadsheetFormatCellsOpenIntent;
  onApply: (patch: SpreadsheetCellFormatPatch) => boolean;
  onClose: () => void;
}) {
  const messages = useOfficeMessages();
  const [draft, setDraft] = useState(() =>
    createSpreadsheetFormatCellsDraft(source),
  );
  const [touched, setTouched] = useState<SpreadsheetFormatCellsTouched>({});
  const [activeTab, setActiveTab] = useState<SpreadsheetFormatCellsTabId>(
    openIntent.tab,
  );
  const formId = useId();
  const idBase = `spreadsheet-format-cells-${useId().replaceAll(':', '')}`;
  const errors = spreadsheetFormatCellsDraftErrors(draft);
  const patch = spreadsheetFormatCellsPatch(source, draft, touched);
  const hasChanges = Object.keys(patch).length > 0;
  const invalid = Object.keys(errors).length > 0;
  const tabs = useMemo(() => {
    const tabLabelKey = {
      number: 'spreadsheet.formatCells.tab.number',
      alignment: 'spreadsheet.formatCells.tab.alignment',
      font: 'spreadsheet.formatCells.tab.font',
      border: 'spreadsheet.formatCells.tab.border',
      fill: 'spreadsheet.formatCells.tab.fill',
      protection: 'spreadsheet.formatCells.tab.protection',
    } as const;
    return spreadsheetFormatCellsTabs.map((tab) => ({
      ...tab,
      label: officeMessage(messages, tabLabelKey[tab.id]),
      tabId: `${idBase}-${tab.id}-tab`,
      panelId: `${idBase}-${tab.id}-panel`,
    }));
  }, [idBase, messages]);

  const submit = (event?: FormEvent<HTMLFormElement>) => {
    event?.preventDefault();
    if (!hasChanges || invalid) return;
    if (onApply(patch)) onClose();
  };

  return (
    <Dialog
      title={officeMessage(messages, 'spreadsheet.formatCells.title')}
      description={formatCellsSelectionDescription(source, messages)}
      className="work-spreadsheet-format-cells-dialog"
      restoreFocusTarget={restoreFocusTarget}
      onClose={onClose}
      footer={
        <>
          <Button tone="quiet" onClick={onClose}>
            {officeMessage(messages, 'spreadsheet.formatCells.cancel')}
          </Button>
          <Button
            tone="primary"
            type="submit"
            form={formId}
            disabled={!hasChanges || invalid}
          >
            {officeMessage(messages, 'spreadsheet.formatCells.apply')}
          </Button>
        </>
      }
    >
      <Tabs
        ariaLabel={officeMessage(messages, 'spreadsheet.formatCells.tabsAria')}
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
          initialFocus={
            openIntent.tab === 'font' ? openIntent.focus : undefined
          }
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
  messages: ReturnType<typeof useOfficeMessages>,
): string {
  const rows = source.range.row[1] - source.range.row[0] + 1;
  const columns = source.range.column[1] - source.range.column[0] + 1;
  return officeMessage(messages, 'spreadsheet.formatCells.description', {
    rows: String(rows),
    columns: String(columns),
  });
}
