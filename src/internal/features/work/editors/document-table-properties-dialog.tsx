import type { Editor } from '@tiptap/core';
import { TableProperties } from 'lucide-react';
import {
  type FormEvent,
  type MouseEvent,
  useId,
  useRef,
  useState,
} from 'react';
import { Button, Dialog } from '../../../design-system/primitives';
import { documentTableCellFormat } from '../work-document-table-cell-formatting';
import {
  canSetDocumentTableRowRepeatHeader,
  documentTableRowOptions,
} from '../work-document-table-row';
import { documentTableSizing } from '../work-document-table-sizing';
import {
  createDocumentTablePropertiesDraft,
  documentTablePropertiesErrors,
  documentTablePropertyChanges,
  hasDocumentTablePropertiesErrors,
  type DocumentTablePropertiesSource,
} from './document-table-properties-dialog-model';
import {
  DocumentTablePropertiesPanel,
  DocumentTablePropertiesTabs,
} from './document-table-properties-dialog-sections';
import { WorkOfficeRibbonButton } from './work-office-chrome';

export function DocumentTablePropertiesControl({
  editor,
  renderedTableWidth,
  renderedRowHeight,
  renderedColumnWidth,
  renderedColumnWidths,
}: {
  editor: Editor;
  renderedTableWidth?: number;
  renderedRowHeight?: number;
  renderedColumnWidth?: number;
  renderedColumnWidths?: readonly number[];
}) {
  const [source, setSource] = useState<DocumentTablePropertiesSource | null>(
    null,
  );
  const triggerRef = useRef<HTMLButtonElement | null>(null);

  const openDialog = (event: MouseEvent<HTMLButtonElement>) => {
    const sizing = documentTableSizing(editor.state);
    const cellFormat = documentTableCellFormat(editor.state);
    if (!sizing || !cellFormat) return;
    triggerRef.current = event.currentTarget;
    setSource({
      sizing,
      cellFormat,
      rowOptions: documentTableRowOptions(editor),
      canRepeatHeader: canSetDocumentTableRowRepeatHeader(editor),
      renderedTableWidth,
      renderedRowHeight,
      renderedColumnWidth,
      renderedColumnWidths,
    });
  };

  return (
    <>
      <WorkOfficeRibbonButton
        label="表格属性"
        visibleLabel="表格属性"
        disabled={!documentTableSizing(editor.state)}
        onClick={openDialog}
      >
        <TableProperties size={18} />
      </WorkOfficeRibbonButton>
      {source && (
        <DocumentTablePropertiesDialog
          editor={editor}
          source={source}
          restoreFocusTarget={() => triggerRef.current}
          onClose={() => setSource(null)}
        />
      )}
    </>
  );
}

function DocumentTablePropertiesDialog({
  editor,
  source,
  restoreFocusTarget,
  onClose,
}: {
  editor: Editor;
  source: DocumentTablePropertiesSource;
  restoreFocusTarget: () => HTMLElement | null;
  onClose: () => void;
}) {
  const [initial] = useState(() => createDocumentTablePropertiesDraft(source));
  const [draft, setDraft] = useState(initial);
  const [activeTab, setActiveTab] = useState<
    'table' | 'row' | 'column' | 'cell'
  >('table');
  const formId = useId();
  const idBase = `document-table-properties-${useId().replaceAll(':', '')}`;
  const errors = documentTablePropertiesErrors(draft);
  const invalid = hasDocumentTablePropertiesErrors(errors);

  const submit = (event?: FormEvent<HTMLFormElement>) => {
    event?.preventDefault();
    if (invalid) return;
    const changes = documentTablePropertyChanges(initial, draft, source);
    if (!changes || editor.commands.setDocumentTablePropertyChanges(changes)) {
      onClose();
    }
  };

  return (
    <Dialog
      title="表格属性"
      description="设置当前表格、行、列和单元格。"
      className="work-document-table-properties-dialog"
      restoreFocusTarget={restoreFocusTarget}
      onClose={onClose}
      footer={
        <>
          <Button tone="quiet" onClick={onClose}>
            取消
          </Button>
          <Button tone="primary" type="submit" form={formId} disabled={invalid}>
            确定
          </Button>
        </>
      }
    >
      <DocumentTablePropertiesTabs
        activeTab={activeTab}
        idBase={idBase}
        onTabChange={setActiveTab}
      />
      <form id={formId} onSubmit={submit}>
        <DocumentTablePropertiesPanel
          activeTab={activeTab}
          idBase={idBase}
          draft={draft}
          setDraft={setDraft}
          errors={errors}
          source={source}
        />
      </form>
    </Dialog>
  );
}
