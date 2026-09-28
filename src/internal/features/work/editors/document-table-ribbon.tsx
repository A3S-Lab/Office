import type { Editor } from '@tiptap/core';
import { officeMessage } from '../../../i18n/office-locale';
import type { OfficeMessageCatalog } from '../../../i18n/office-messages';
import { NodeSelection } from '@tiptap/pm/state';
import { isInTable, selectedRect, TableMap } from '@tiptap/pm/tables';
import {
  AlignCenter,
  AlignLeft,
  AlignRight,
  AlignVerticalJustifyCenter,
  AlignVerticalJustifyEnd,
  AlignVerticalJustifyStart,
  BetweenHorizontalEnd,
  BetweenHorizontalStart,
  BetweenVerticalEnd,
  BetweenVerticalStart,
  Columns3,
  Grid2X2,
  PanelTop,
  Rows3,
  TableCellsMerge,
  TableCellsSplit,
  Trash2,
} from 'lucide-react';
import {
  type CSSProperties,
  type KeyboardEvent,
  type MouseEvent as ReactMouseEvent,
  type ReactNode,
  useEffect,
  useId,
  useRef,
  useState,
} from 'react';
import {
  activeDocumentTableStyle,
  DEFAULT_DOCUMENT_TABLE_CELL_FORMAT,
  DOCUMENT_TABLE_STYLE_OPTIONS,
  type DocumentTableBorder,
  type DocumentTableBorderStyle,
  type DocumentTableBorderTarget,
  documentTableCellFormat,
  documentTableHorizontalAlignment,
  type DocumentTableStyleOption,
} from '../work-document-table-cell-formatting';
import {
  canSetDocumentTableRowRepeatHeader,
  documentTableRowOptions,
} from '../work-document-table-row';
import {
  documentTableSizing,
  type DocumentTableLayoutMode,
  type DocumentTableSizingState,
} from '../work-document-table-sizing';
import {
  OfficeColorPicker,
  OfficeNumberField,
  OfficeSelect,
  type OfficeSelectOption,
} from './office-controls';
import {
  WorkOfficeRibbonButton,
  WorkOfficeRibbonGroup,
} from './work-office-chrome';
import { useOfficeDraft } from './use-office-draft';
import { useOfficeMessages } from './office-messages-context';
import { DocumentTableMarginsPopover } from './document-table-margins-popover';
import { DocumentTablePropertiesControl } from './document-table-properties-dialog';

type DocumentTableBorderAction = DocumentTableBorderTarget | 'clear';
type DocumentTableLayoutControlValue = DocumentTableLayoutMode | 'current';

function borderOptions(messages: OfficeMessageCatalog) {
  return [
    {
      value: 'solid-1',
      label: officeMessage(messages, 'document.table.border.solid1'),
      style: 'solid' as const,
      width: 1,
    },
    {
      value: 'solid-2',
      label: officeMessage(messages, 'document.table.border.solid2'),
      style: 'solid' as const,
      width: 2,
    },
    {
      value: 'dashed-1',
      label: officeMessage(messages, 'document.table.border.dashed'),
      style: 'dashed' as const,
      width: 1,
    },
    {
      value: 'dotted-1',
      label: officeMessage(messages, 'document.table.border.dotted'),
      style: 'dotted' as const,
      width: 1,
    },
    {
      value: 'double-2',
      label: officeMessage(messages, 'document.table.border.double'),
      style: 'double' as const,
      width: 2,
    },
  ];
}

function borderTargetOptions(messages: OfficeMessageCatalog) {
  return [
    {
      value: 'clear' as const,
      label: officeMessage(messages, 'document.table.border.clear'),
    },
    {
      value: 'all' as const,
      label: officeMessage(messages, 'document.table.border.all'),
    },
    {
      value: 'outside' as const,
      label: officeMessage(messages, 'document.table.border.outside'),
    },
    {
      value: 'inside' as const,
      label: officeMessage(messages, 'document.table.border.inside'),
    },
    {
      value: 'top' as const,
      label: officeMessage(messages, 'document.table.border.top'),
    },
    {
      value: 'bottom' as const,
      label: officeMessage(messages, 'document.table.border.bottom'),
    },
    {
      value: 'left' as const,
      label: officeMessage(messages, 'document.table.border.left'),
    },
    {
      value: 'right' as const,
      label: officeMessage(messages, 'document.table.border.right'),
    },
    {
      value: 'insideHorizontal' as const,
      label: officeMessage(messages, 'document.table.border.insideHorizontal'),
    },
    {
      value: 'insideVertical' as const,
      label: officeMessage(messages, 'document.table.border.insideVertical'),
    },
  ];
}

function tableLayoutOptions(messages: OfficeMessageCatalog) {
  return [
    {
      value: 'window' as const,
      label: officeMessage(messages, 'document.table.layout.window'),
    },
    {
      value: 'contents' as const,
      label: officeMessage(messages, 'document.table.layout.contents'),
    },
    {
      value: 'fixed' as const,
      label: officeMessage(messages, 'document.table.layout.fixed'),
    },
  ];
}

function currentTableLayoutOption(
  messages: OfficeMessageCatalog,
): OfficeSelectOption<DocumentTableLayoutControlValue> {
  return {
    value: 'current',
    label: officeMessage(messages, 'document.table.layout.current'),
    disabled: true,
  };
}

const PIXELS_PER_CENTIMETER = 96 / 2.54;

export function DocumentTableDesignRibbon({ editor }: { editor: Editor }) {
  const messages = useOfficeMessages();
  const format =
    documentTableCellFormat(editor.state) ?? DEFAULT_DOCUMENT_TABLE_CELL_FORMAT;
  const [borderPen, setBorderPen] = useState<DocumentTableBorder>(() => ({
    color: format.borderColor,
    style: format.borderStyle === 'none' ? 'solid' : format.borderStyle,
    width:
      format.borderStyle === 'none' || format.borderWidth === 0
        ? 1
        : format.borderWidth,
  }));
  const [borderTarget, setBorderTarget] =
    useState<DocumentTableBorderAction>('all');
  const borderValue = borderOptionValue(borderPen.style, borderPen.width);
  const currentBorderOptions = borderOptionsForFormat(
    messages,
    borderPen.style,
    borderPen.width,
  );
  const targetOptions = borderTargetOptions(messages);
  return (
    <>
      <RibbonGroup label={officeMessage(messages, 'document.table.group.style')}>
        <DocumentTableStyleGallery editor={editor} />
      </RibbonGroup>
      <RibbonGroup label={officeMessage(messages, 'document.table.group.options')}>
        <RibbonButton
          label={officeMessage(messages, 'document.table.headerRow')}
          visibleLabel={officeMessage(messages, 'document.table.headerRow')}
          active={editor.isActive('tableHeader')}
          disabled={!editor.can().toggleHeaderRow()}
          onMouseDown={(event) => event.preventDefault()}
          onClick={() => {
            // Keep ribbon focus on the header-row trigger.
            // chain().focus() schedules into the editor and breaks L2 loops.
            editor.commands.toggleHeaderRow();
          }}
        >
          <PanelTop size={18} />
        </RibbonButton>
      </RibbonGroup>
      <RibbonGroup label={officeMessage(messages, 'document.table.group.shading')}>
        <OfficeColorPicker
          ariaLabel={officeMessage(messages, 'document.table.shadingAria')}
          className="work-document-table-color-picker"
          triggerLabel={officeMessage(messages, 'document.table.shading')}
          value={format.backgroundColor}
          onValueChange={(backgroundColor) => {
            // Keep ribbon focus on the color trigger via Popover restore.
            // chain().focus() schedules into the editor and breaks L2 loops.
            editor.commands.setDocumentTableCellFormat({ backgroundColor });
          }}
        />
      </RibbonGroup>
      <RibbonGroup label={officeMessage(messages, 'document.table.group.borders')}>
        <OfficeColorPicker
          ariaLabel={officeMessage(messages, 'document.table.borderColorAria')}
          className="work-document-table-color-picker"
          compact
          value={borderPen.color}
          onValueChange={(borderColor) =>
            setBorderPen((current) => ({ ...current, color: borderColor }))
          }
        />
        <OfficeSelect
          ariaLabel={officeMessage(messages, 'document.table.borderStyleAria')}
          className="work-document-table-border-select"
          value={borderValue}
          options={currentBorderOptions}
          onValueChange={(value) => {
            const option = currentBorderOptions.find(
              (candidate) => candidate.value === value,
            );
            if (!option) return;
            setBorderPen((current) => ({
              ...current,
              style: option.style,
              width: option.width,
            }));
          }}
        />
        <OfficeSelect
          ariaLabel={officeMessage(messages, 'document.table.borderApplyAria')}
          className="work-document-table-border-target-select"
          value={borderTarget}
          options={targetOptions}
          onValueChange={(value) => {
            setBorderTarget(value);
            const clear = value === 'clear';
            // Keep ribbon focus on the border-target combobox via Popover
            // restore. chain().focus() schedules into the editor and breaks
            // L2 loops (same class as cell fill).
            editor.commands.setDocumentTableBorders(
              clear ? 'all' : value,
              clear
                ? { color: borderPen.color, style: 'none', width: 0 }
                : borderPen,
            );
          }}
        />
      </RibbonGroup>
    </>
  );
}

export function DocumentTableLayoutRibbon({ editor }: { editor: Editor }) {
  const messages = useOfficeMessages();
  const rowOptions = documentTableRowOptions(editor);
  const canSetRowOptions = editor
    .can()
    .setDocumentTableRowOptions(rowOptions, { restoreFocus: false });
  const sizing = documentTableSizing(editor.state);
  const cellFormat =
    documentTableCellFormat(editor.state) ?? DEFAULT_DOCUMENT_TABLE_CELL_FORMAT;
  const horizontalAlignment = documentTableHorizontalAlignment(editor.state);
  return (
    <>
      <RibbonGroup label={officeMessage(messages, 'document.table.group.rows')}>
        <RibbonButton
          label={officeMessage(messages, 'document.table.insertRowAbove')}
          visibleLabel={officeMessage(messages, 'document.table.insertRowAboveShort')}
          disabled={!editor.can().addRowBefore()}
          onMouseDown={(event) => event.preventDefault()}
          onClick={() => {
            // Keep ribbon focus on the insert-row trigger.
            // chain().focus() schedules into the editor and breaks L2 loops.
            editor.commands.addRowBefore();
          }}
        >
          <BetweenHorizontalStart size={18} />
        </RibbonButton>
        <RibbonButton
          label={officeMessage(messages, 'document.table.insertRowBelow')}
          visibleLabel={officeMessage(messages, 'document.table.insertRowBelowShort')}
          disabled={!editor.can().addRowAfter()}
          onMouseDown={(event) => event.preventDefault()}
          onClick={() => {
            // Keep ribbon focus on the insert-row trigger.
            // chain().focus() schedules into the editor and breaks L2 loops.
            editor.commands.addRowAfter();
          }}
        >
          <BetweenHorizontalEnd size={18} />
        </RibbonButton>
        <RibbonButton
          label={officeMessage(messages, 'document.table.deleteRow')}
          visibleLabel={officeMessage(messages, 'document.table.deleteRowShort')}
          disabled={!editor.can().deleteRow()}
          onMouseDown={(event) => event.preventDefault()}
          onClick={() => {
            // Keep ribbon focus on the delete-row trigger.
            // chain().focus() schedules into the editor and breaks L2 loops.
            editor.commands.deleteRow();
          }}
        >
          <Rows3 size={18} />
        </RibbonButton>
      </RibbonGroup>
      <RibbonGroup label={officeMessage(messages, 'document.table.group.columns')}>
        <RibbonButton
          label={officeMessage(messages, 'document.table.insertColLeft')}
          visibleLabel={officeMessage(messages, 'document.table.insertColLeftShort')}
          disabled={!editor.can().addColumnBefore()}
          onMouseDown={(event) => event.preventDefault()}
          onClick={() => {
            // Keep ribbon focus on the insert-column trigger.
            // chain().focus() schedules into the editor and breaks L2 loops.
            editor.commands.addColumnBefore();
          }}
        >
          <BetweenVerticalStart size={18} />
        </RibbonButton>
        <RibbonButton
          label={officeMessage(messages, 'document.table.insertColRight')}
          visibleLabel={officeMessage(messages, 'document.table.insertColRightShort')}
          disabled={!editor.can().addColumnAfter()}
          onMouseDown={(event) => event.preventDefault()}
          onClick={() => {
            // Keep ribbon focus on the insert-column trigger.
            // chain().focus() schedules into the editor and breaks L2 loops.
            editor.commands.addColumnAfter();
          }}
        >
          <BetweenVerticalEnd size={18} />
        </RibbonButton>
        <RibbonButton
          label={officeMessage(messages, 'document.table.deleteCol')}
          visibleLabel={officeMessage(messages, 'document.table.deleteColShort')}
          disabled={!editor.can().deleteColumn()}
          onMouseDown={(event) => event.preventDefault()}
          onClick={() => {
            // Keep ribbon focus on the delete-column trigger.
            // chain().focus() schedules into the editor and breaks L2 loops.
            editor.commands.deleteColumn();
          }}
        >
          <Trash2 size={18} />
        </RibbonButton>
      </RibbonGroup>
      <RibbonGroup label={officeMessage(messages, 'document.table.group.merge')}>
        <RibbonButton
          label={officeMessage(messages, 'document.table.mergeCells')}
          visibleLabel={officeMessage(messages, 'document.table.mergeShort')}
          disabled={!editor.can().mergeCells()}
          onMouseDown={(event) => event.preventDefault()}
          onClick={() => {
            // Keep ribbon focus on the merge trigger.
            // chain().focus() schedules into the editor and breaks L2 loops.
            editor.commands.mergeCells();
          }}
        >
          <TableCellsMerge size={18} />
        </RibbonButton>
        <RibbonButton
          label={officeMessage(messages, 'document.table.splitCells')}
          visibleLabel={officeMessage(messages, 'document.table.splitShort')}
          disabled={!editor.can().splitCell()}
          onMouseDown={(event) => event.preventDefault()}
          onClick={() => {
            // Keep ribbon focus on the split trigger.
            // chain().focus() schedules into the editor and breaks L2 loops.
            editor.commands.splitCell();
          }}
        >
          <TableCellsSplit size={18} />
        </RibbonButton>
      </RibbonGroup>
      <DocumentTableSizeRibbonGroup editor={editor} sizing={sizing} />
      <RibbonGroup label={officeMessage(messages, 'document.table.group.properties')}>
        <DocumentTablePropertiesControl
          editor={editor}
          renderedTableWidth={measuredTableWidth(editor)}
          renderedRowHeight={
            measuredCurrentTableDimension(editor, 'rows') ?? undefined
          }
          renderedColumnWidth={
            measuredCurrentTableDimension(editor, 'columns') ?? undefined
          }
          renderedColumnWidths={measuredTableColumnWidths(editor)}
        />
        <DocumentTableMarginsPopover editor={editor} />
      </RibbonGroup>
      <RibbonGroup label={officeMessage(messages, 'document.table.group.alignment')}>
        <RibbonButton
          label={officeMessage(messages, 'document.table.alignLeft')}
          visibleLabel={officeMessage(messages, 'document.table.alignLeftShort')}
          active={horizontalAlignment === 'left'}
          onMouseDown={(event) => event.preventDefault()}
          onClick={() => {
            // Keep ribbon focus on the alignment trigger.
            // chain().focus() schedules into the editor and breaks L2 loops.
            editor.commands.setDocumentTableHorizontalAlignment('left');
          }}
        >
          <AlignLeft size={18} />
        </RibbonButton>
        <RibbonButton
          label={officeMessage(messages, 'document.table.alignCenter')}
          visibleLabel={officeMessage(messages, 'document.table.alignCenterShort')}
          active={horizontalAlignment === 'center'}
          onMouseDown={(event) => event.preventDefault()}
          onClick={() => {
            // Keep ribbon focus on the alignment trigger.
            // chain().focus() schedules into the editor and breaks L2 loops.
            editor.commands.setDocumentTableHorizontalAlignment('center');
          }}
        >
          <AlignCenter size={18} />
        </RibbonButton>
        <RibbonButton
          label={officeMessage(messages, 'document.table.alignRight')}
          visibleLabel={officeMessage(messages, 'document.table.alignRightShort')}
          active={horizontalAlignment === 'right'}
          onMouseDown={(event) => event.preventDefault()}
          onClick={() => {
            // Keep ribbon focus on the alignment trigger.
            // chain().focus() schedules into the editor and breaks L2 loops.
            editor.commands.setDocumentTableHorizontalAlignment('right');
          }}
        >
          <AlignRight size={18} />
        </RibbonButton>
        <RibbonButton
          label={officeMessage(messages, 'document.table.alignTop')}
          visibleLabel={officeMessage(messages, 'document.table.alignTopShort')}
          active={cellFormat.verticalAlign === 'top'}
          onMouseDown={(event) => event.preventDefault()}
          onClick={() => setVerticalAlignment(editor, 'top')}
        >
          <AlignVerticalJustifyStart size={18} />
        </RibbonButton>
        <RibbonButton
          label={officeMessage(messages, 'document.table.alignMiddle')}
          visibleLabel={officeMessage(messages, 'document.table.alignMiddleShort')}
          active={cellFormat.verticalAlign === 'middle'}
          onMouseDown={(event) => event.preventDefault()}
          onClick={() => setVerticalAlignment(editor, 'middle')}
        >
          <AlignVerticalJustifyCenter size={18} />
        </RibbonButton>
        <RibbonButton
          label={officeMessage(messages, 'document.table.alignBottom')}
          visibleLabel={officeMessage(messages, 'document.table.alignBottomShort')}
          active={cellFormat.verticalAlign === 'bottom'}
          onMouseDown={(event) => event.preventDefault()}
          onClick={() => setVerticalAlignment(editor, 'bottom')}
        >
          <AlignVerticalJustifyEnd size={18} />
        </RibbonButton>
      </RibbonGroup>
      <RibbonGroup label={officeMessage(messages, 'document.table.group.options')}>
        <RibbonButton
          label={officeMessage(messages, 'document.table.repeatHeader')}
          visibleLabel={officeMessage(messages, 'document.table.repeatHeaderShort')}
          active={rowOptions.repeatHeader}
          disabled={!canSetDocumentTableRowRepeatHeader(editor)}
          onClick={() =>
            editor.commands.setDocumentTableRowOptions({
              ...documentTableRowOptions(editor),
              repeatHeader: !documentTableRowOptions(editor).repeatHeader,
            })
          }
        >
          <Rows3 size={18} />
        </RibbonButton>
        <RibbonButton
          label={officeMessage(messages, 'document.table.keepTogether')}
          visibleLabel={officeMessage(messages, 'document.table.keepTogetherShort')}
          active={rowOptions.cantSplit}
          disabled={!canSetRowOptions}
          onClick={() =>
            editor.commands.setDocumentTableRowOptions({
              ...documentTableRowOptions(editor),
              cantSplit: !documentTableRowOptions(editor).cantSplit,
            })
          }
        >
          <BetweenHorizontalEnd size={18} />
        </RibbonButton>
      </RibbonGroup>
      <RibbonGroup label={officeMessage(messages, 'document.table.group.table')}>
        <RibbonButton
          label={officeMessage(messages, 'document.table.deleteTable')}
          visibleLabel={officeMessage(messages, 'document.table.deleteTable')}
          disabled={!editor.can().deleteTable()}
          onClick={() => {
            // Deleting the table removes the contextual ribbon, so restore
            // editor focus. Other Layout structural controls omit focus() to
            // keep L2 loops on the ribbon trigger.
            editor.chain().focus().deleteTable().run();
          }}
        >
          <Trash2 size={18} />
        </RibbonButton>
      </RibbonGroup>
    </>
  );
}

function DocumentTableSizeRibbonGroup({
  editor,
  sizing,
}: {
  editor: Editor;
  sizing: DocumentTableSizingState | null;
}) {
  const messages = useOfficeMessages();
  const measuredRowSelection = measuredTableSelectionSize(editor, 'rows');
  const measuredColumnSelection = measuredTableSelectionSize(editor, 'columns');
  const measuredColumnWidth = measuredCurrentTableDimension(editor, 'columns');
  const displayedColumnWidth =
    sizing?.layoutAlgorithm === 'autofit'
      ? (measuredColumnWidth ?? sizing.columnWidth)
      : (sizing?.columnWidth ?? measuredColumnWidth);
  const tableLayoutValue: DocumentTableLayoutControlValue =
    sizing?.layoutAlgorithm === 'autofit' && sizing.layoutMode === 'fixed'
      ? 'current'
      : (sizing?.layoutMode ?? 'window');
  const displayedTableLayoutOptions: readonly OfficeSelectOption<DocumentTableLayoutControlValue>[] =
    tableLayoutValue === 'current'
      ? [currentTableLayoutOption(messages), ...tableLayoutOptions(messages)]
      : tableLayoutOptions(messages);
  const canDistributeRows = editor
    .can()
    .distributeDocumentTableRows(measuredRowSelection);
  const canDistributeColumns = editor
    .can()
    .distributeDocumentTableColumns(measuredColumnSelection);
  return (
    <RibbonGroup label={officeMessage(messages, 'document.table.group.cellSize')}>
      <div className="work-document-table-size-fields">
        <TableDimensionField
          label={officeMessage(messages, 'document.table.rowHeight')}
          ariaLabel={officeMessage(messages, 'document.table.rowHeightAria')}
          value={
            sizing?.rowHeight ?? measuredCurrentTableDimension(editor, 'rows')
          }
          onValueChange={(height) =>
            editor.commands.setDocumentTableRowHeight(
              height,
              sizing?.rowHeightRule ?? 'atLeast',
            )
          }
        />
        <TableDimensionField
          label={officeMessage(messages, 'document.table.colWidth')}
          ariaLabel={officeMessage(messages, 'document.table.colWidthAria')}
          value={displayedColumnWidth}
          onValueChange={(width) =>
            editor.commands.setDocumentTableColumnWidth(
              width,
              measuredTableColumnWidths(editor),
            )
          }
        />
      </div>
      <OfficeSelect
        ariaLabel={officeMessage(messages, 'document.table.autoFitAria')}
        className="work-document-table-layout-select"
        value={tableLayoutValue}
        options={displayedTableLayoutOptions}
        onValueChange={(layoutMode) => {
          if (layoutMode === 'current') return;
          // Keep ribbon focus on the layout combobox via Popover restore.
          // chain().focus() schedules into the editor and breaks L2 loops.
          editor.commands.setDocumentTableLayoutMode(
            layoutMode,
            layoutMode === 'fixed' ? measuredTableWidth(editor) : undefined,
          );
        }}
      />
      <RibbonButton
        label={officeMessage(messages, 'document.table.distributeRows')}
        visibleLabel={officeMessage(messages, 'document.table.distributeRowsShort')}
        disabled={!canDistributeRows}
        onMouseDown={(event) => event.preventDefault()}
        onClick={() => {
          // Keep ribbon focus on the distribute trigger.
          // chain().focus() schedules into the editor and breaks L2 loops.
          editor.commands.distributeDocumentTableRows(measuredRowSelection);
        }}
      >
        <Rows3 size={18} />
      </RibbonButton>
      <RibbonButton
        label={officeMessage(messages, 'document.table.distributeCols')}
        visibleLabel={officeMessage(messages, 'document.table.distributeColsShort')}
        disabled={!canDistributeColumns}
        onMouseDown={(event) => event.preventDefault()}
        onClick={() => {
          // Keep ribbon focus on the distribute trigger.
          // chain().focus() schedules into the editor and breaks L2 loops.
          editor.commands.distributeDocumentTableColumns(
            measuredColumnSelection,
          );
        }}
      >
        <Columns3 size={18} />
      </RibbonButton>
    </RibbonGroup>
  );
}

function TableDimensionField({
  label,
  ariaLabel,
  value,
  onValueChange,
}: {
  label: string;
  ariaLabel: string;
  value: number | null;
  onValueChange: (value: number) => boolean;
}) {
  const messages = useOfficeMessages();
  const formattedValue =
    value === null ? '' : formatCentimeters(value / PIXELS_PER_CENTIMETER);
  const { cancelDraft, dirty, draft, replaceDraft, setDraft, syncDraft } =
    useOfficeDraft(() => formattedValue);
  useEffect(() => syncDraft(formattedValue), [formattedValue, syncDraft]);
  return (
    <div className="work-document-table-size-field">
      <span>{label}</span>
      <OfficeNumberField
        ariaLabel={ariaLabel}
        value={draft}
        min={0.5}
        max={30}
        step={0.1}
        placeholder="—"
        escapeConsumer={dirty}
        onValueChange={setDraft}
        onCancel={dirty ? cancelDraft : undefined}
        onCommit={(nextValue) => {
          const centimeters = Number(nextValue);
          if (
            !Number.isFinite(centimeters) ||
            centimeters < 0.5 ||
            centimeters > 30
          ) {
            cancelDraft();
            return;
          }
          const normalizedDraft = formatCentimeters(centimeters);
          const committed = onValueChange(
            Math.round(centimeters * PIXELS_PER_CENTIMETER * 100) / 100,
          );
          if (committed) replaceDraft(normalizedDraft);
          else cancelDraft();
        }}
      />
      <small aria-hidden="true">{officeMessage(messages, 'document.table.unit.cm')}</small>
    </div>
  );
}

function measuredTableWidth(editor: Editor): number | undefined {
  const table = selectedTableElement(editor);
  const width = table?.getBoundingClientRect().width || table?.offsetWidth || 0;
  return width > 0 ? width : undefined;
}

function measuredCurrentTableDimension(
  editor: Editor,
  axis: 'columns' | 'rows',
): number | null {
  const table = selectedTableElement(editor);
  const rectangle = selectedTableRectangle(editor);
  if (!table || !rectangle) return null;
  if (axis === 'rows') {
    const row = table.rows[rectangle.top];
    const height =
      row?.getBoundingClientRect().height || row?.offsetHeight || 0;
    return height > 0 ? height : null;
  }
  return measuredTableColumnWidths(editor)?.[rectangle.left] ?? null;
}

function measuredTableColumnWidths(editor: Editor): number[] | undefined {
  const table = selectedTableElement(editor);
  if (!table) return undefined;
  const columns = Array.from(table.querySelectorAll('colgroup > col'));
  const tableWidth = table.getBoundingClientRect().width || table.offsetWidth;
  const widths = columns.map(
    (column) =>
      column.getBoundingClientRect().width ||
      Number.parseFloat((column as HTMLElement).style.width) ||
      Number.parseFloat((column as HTMLElement).style.minWidth) ||
      (tableWidth > 0 && columns.length ? tableWidth / columns.length : 0),
  );
  return widths.length && widths.every((width) => width > 0)
    ? widths
    : undefined;
}

function measuredTableSelectionSize(
  editor: Editor,
  axis: 'columns' | 'rows',
): number | undefined {
  const table = selectedTableElement(editor);
  const sizing = documentTableSizing(editor.state);
  if (!table || !sizing) return undefined;
  if (axis === 'rows') {
    const rows = Array.from(table.tBodies[0]?.rows ?? []);
    const selectedRows =
      sizing.selectedRowCount > 1
        ? rows.slice(
            selectedTableRectangle(editor)?.top ?? 0,
            selectedTableRectangle(editor)?.bottom ?? rows.length,
          )
        : rows;
    const heights = selectedRows.map(
      (row) => row.getBoundingClientRect().height || row.offsetHeight,
    );
    return heights.every((height) => height > 0)
      ? heights.reduce((sum, height) => sum + height, 0)
      : undefined;
  }
  const columns = Array.from(table.querySelectorAll('colgroup > col'));
  const rectangle = selectedTableRectangle(editor);
  const selectedColumns =
    sizing.selectedColumnCount > 1 && rectangle
      ? columns.slice(rectangle.left, rectangle.right)
      : columns;
  const measuredWidths = measuredTableColumnWidths(editor);
  const widths = measuredWidths
    ? sizing.selectedColumnCount > 1 && rectangle
      ? measuredWidths.slice(rectangle.left, rectangle.right)
      : measuredWidths
    : selectedColumns.map(() => 0);
  if (widths.length && widths.every((width) => width > 0)) {
    return widths.reduce((sum, width) => sum + width, 0);
  }
  const tableWidth = measuredTableWidth(editor);
  return tableWidth && columns.length
    ? (tableWidth * selectedColumns.length) / columns.length
    : undefined;
}

function selectedTableRectangle(editor: Editor): {
  left: number;
  right: number;
  top: number;
  bottom: number;
} | null {
  const selection = editor.state.selection;
  if (
    selection instanceof NodeSelection &&
    selection.node.type.spec.tableRole === 'table'
  ) {
    const map = TableMap.get(selection.node);
    return { left: 0, right: map.width, top: 0, bottom: map.height };
  }
  if (!isInTable(editor.state)) return null;
  const rectangle = selectedRect(editor.state);
  return {
    left: rectangle.left,
    right: rectangle.right,
    top: rectangle.top,
    bottom: rectangle.bottom,
  };
}

function selectedTableElement(editor: Editor): HTMLTableElement | null {
  const selection = editor.state.selection;
  if (
    selection instanceof NodeSelection &&
    selection.node.type.spec.tableRole === 'table'
  ) {
    const nodeDom = editor.view.nodeDOM(selection.from);
    if (nodeDom instanceof HTMLTableElement) return nodeDom;
    if (nodeDom instanceof HTMLElement) {
      return nodeDom.querySelector(':scope > table');
    }
    return null;
  }
  for (let depth = selection.$from.depth; depth > 0; depth -= 1) {
    if (selection.$from.node(depth).type.spec.tableRole !== 'table') continue;
    const nodeDom = editor.view.nodeDOM(selection.$from.before(depth));
    if (nodeDom instanceof HTMLTableElement) return nodeDom;
    if (nodeDom instanceof HTMLElement) {
      return nodeDom.querySelector(':scope > table');
    }
  }
  return null;
}

function formatCentimeters(value: number): string {
  return value.toFixed(2).replace(/\.?0+$/, '');
}

function DocumentTableStyleGallery({ editor }: { editor: Editor }) {
  const messages = useOfficeMessages();
  const groupName = useId();
  const inputsRef = useRef<Array<HTMLInputElement | null>>([]);
  const activeStyle = activeDocumentTableStyle(editor.state);

  const applyStyle = (style: DocumentTableStyleOption) => {
    // Keep radiogroup focus on the style option.
    // chain().focus() schedules into the editor and breaks L2 loops.
    editor.commands.applyDocumentTableStyle(style.id);
  };
  const moveSelection = (
    event: KeyboardEvent<HTMLInputElement>,
    nextIndex: number,
  ) => {
    event.preventDefault();
    const normalizedIndex =
      (nextIndex + DOCUMENT_TABLE_STYLE_OPTIONS.length) %
      DOCUMENT_TABLE_STYLE_OPTIONS.length;
    const style = DOCUMENT_TABLE_STYLE_OPTIONS[normalizedIndex];
    if (!style) return;
    applyStyle(style);
    inputsRef.current[normalizedIndex]?.focus({ preventScroll: true });
  };

  return (
    <div
      className="work-document-table-style-gallery"
      role="radiogroup"
      aria-label={officeMessage(messages, 'document.table.styleGalleryAria')}
    >
      {DOCUMENT_TABLE_STYLE_OPTIONS.map((style, index) => {
        const active = style.id === activeStyle;
        return (
          <label key={style.id} className={active ? 'active' : ''}>
            <input
              ref={(input) => {
                inputsRef.current[index] = input;
              }}
              type="radio"
              name={groupName}
              aria-label={officeMessage(messages, 'document.table.applyStyleAria', {
                label: style.label,
              })}
              checked={active}
              tabIndex={active || (!activeStyle && index === 0) ? 0 : -1}
              onChange={() => applyStyle(style)}
              onKeyDown={(event) => {
                if (event.key === 'ArrowRight' || event.key === 'ArrowDown') {
                  moveSelection(event, index + 1);
                } else if (
                  event.key === 'ArrowLeft' ||
                  event.key === 'ArrowUp'
                ) {
                  moveSelection(event, index - 1);
                } else if (event.key === 'Home') {
                  moveSelection(event, 0);
                } else if (event.key === 'End') {
                  moveSelection(event, DOCUMENT_TABLE_STYLE_OPTIONS.length - 1);
                }
              }}
            />
            <span
              className="work-document-table-style-preview"
              style={
                {
                  '--work-table-style-header': style.headerColor,
                  '--work-table-style-body': style.bodyColor,
                  '--work-table-style-alternate': style.alternateColor,
                  '--work-table-style-border':
                    style.borderStyle === 'none'
                      ? 'transparent'
                      : style.borderColor,
                } as CSSProperties
              }
              aria-hidden="true"
            >
              <Grid2X2 size={24} strokeWidth={1.35} />
            </span>
            <span>{style.label}</span>
          </label>
        );
      })}
    </div>
  );
}

function setVerticalAlignment(
  editor: Editor,
  verticalAlign: 'top' | 'middle' | 'bottom',
) {
  // Keep ribbon focus on the vertical-align trigger.
  // chain().focus() schedules into the editor and breaks L2 loops.
  editor.commands.setDocumentTableCellFormat({ verticalAlign });
}

function borderOptionValue(
  style: DocumentTableBorderStyle,
  width: number,
): string {
  return `${style}-${width}`;
}

function borderOptionsForFormat(
  messages: OfficeMessageCatalog,
  style: DocumentTableBorderStyle,
  width: number,
) {
  const options = borderOptions(messages);
  const value = borderOptionValue(style, width);
  if (options.some((option) => option.value === value)) {
    return options;
  }
  return [
    ...options,
    {
      value,
      label:
        style === 'none'
          ? officeMessage(messages, 'document.table.border.none')
          : officeMessage(messages, 'document.table.border.summary', {
              width: String(width),
              style: documentTableBorderStyleLabel(messages, style),
            }),
      style,
      width,
    },
  ];
}

function documentTableBorderStyleLabel(
  messages: OfficeMessageCatalog,
  style: Exclude<DocumentTableBorderStyle, 'none'>,
): string {
  if (style === 'dashed') {
    return officeMessage(messages, 'document.table.border.dashed');
  }
  if (style === 'dotted') {
    return officeMessage(messages, 'document.table.border.dotted');
  }
  if (style === 'double') {
    return officeMessage(messages, 'document.table.border.double');
  }
  return officeMessage(messages, 'document.table.border.solid');
}

function RibbonButton({
  label,
  visibleLabel,
  active = false,
  disabled = false,
  onClick,
  onMouseDown,
  children,
}: {
  label: string;
  visibleLabel: string;
  active?: boolean;
  disabled?: boolean;
  onClick: () => void;
  onMouseDown?: (event: ReactMouseEvent<HTMLButtonElement>) => void;
  children: ReactNode;
}) {
  return (
    <WorkOfficeRibbonButton
      label={label}
      visibleLabel={visibleLabel}
      active={active}
      disabled={disabled}
      onMouseDown={onMouseDown}
      onClick={onClick}
    >
      {children}
    </WorkOfficeRibbonButton>
  );
}

const RibbonGroup = WorkOfficeRibbonGroup;
