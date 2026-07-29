import type { Editor } from '@tiptap/core';
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
  type ReactNode,
  useEffect,
  useId,
  useRef,
} from 'react';
import {
  activeDocumentTableStyle,
  DEFAULT_DOCUMENT_TABLE_CELL_FORMAT,
  DOCUMENT_TABLE_STYLE_OPTIONS,
  type DocumentTableBorderStyle,
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
} from './office-controls';
import {
  WorkOfficeRibbonButton,
  WorkOfficeRibbonGroup,
} from './work-office-chrome';
import { useOfficeDraft } from './use-office-draft';

const borderOptions = [
  { value: 'solid-1', label: '细实线', style: 'solid', width: 1 },
  { value: 'solid-2', label: '粗实线', style: 'solid', width: 2 },
  { value: 'dashed-1', label: '虚线', style: 'dashed', width: 1 },
  { value: 'dotted-1', label: '点线', style: 'dotted', width: 1 },
  { value: 'double-2', label: '双线', style: 'double', width: 2 },
  { value: 'none-0', label: '无边框', style: 'none', width: 0 },
] as const satisfies readonly {
  value: string;
  label: string;
  style: DocumentTableBorderStyle;
  width: number;
}[];

const tableLayoutOptions = [
  { value: 'window', label: '适应窗口' },
  { value: 'contents', label: '适应内容' },
  { value: 'fixed', label: '固定列宽' },
] as const satisfies readonly {
  value: DocumentTableLayoutMode;
  label: string;
}[];

const PIXELS_PER_CENTIMETER = 96 / 2.54;

export function DocumentTableDesignRibbon({ editor }: { editor: Editor }) {
  const format =
    documentTableCellFormat(editor.state) ?? DEFAULT_DOCUMENT_TABLE_CELL_FORMAT;
  const borderValue = borderOptionValue(format.borderStyle, format.borderWidth);
  const currentBorderOptions = borderOptionsForFormat(
    format.borderStyle,
    format.borderWidth,
  );
  return (
    <>
      <RibbonGroup label="表格样式">
        <DocumentTableStyleGallery editor={editor} />
      </RibbonGroup>
      <RibbonGroup label="表格选项">
        <RibbonButton
          label="标题行"
          visibleLabel="标题行"
          active={editor.isActive('tableHeader')}
          disabled={!editor.can().chain().focus().toggleHeaderRow().run()}
          onClick={() => editor.chain().focus().toggleHeaderRow().run()}
        >
          <PanelTop size={18} />
        </RibbonButton>
      </RibbonGroup>
      <RibbonGroup label="底纹">
        <OfficeColorPicker
          ariaLabel="单元格底纹"
          className="work-document-table-color-picker"
          triggerLabel="底纹"
          value={format.backgroundColor}
          onValueChange={(backgroundColor) =>
            editor
              .chain()
              .focus()
              .setDocumentTableCellFormat({ backgroundColor })
              .run()
          }
        />
      </RibbonGroup>
      <RibbonGroup label="边框">
        <OfficeColorPicker
          ariaLabel="边框颜色"
          className="work-document-table-color-picker"
          compact
          value={format.borderColor}
          onValueChange={(borderColor) =>
            editor
              .chain()
              .focus()
              .setDocumentTableCellFormat({ borderColor })
              .run()
          }
        />
        <OfficeSelect
          ariaLabel="边框样式"
          className="work-document-table-border-select"
          value={borderValue}
          options={currentBorderOptions}
          onValueChange={(value) => {
            const option = currentBorderOptions.find(
              (candidate) => candidate.value === value,
            );
            if (!option) return;
            editor
              .chain()
              .focus()
              .setDocumentTableCellFormat({
                borderStyle: option.style,
                borderWidth: option.width,
              })
              .run();
          }}
        />
      </RibbonGroup>
    </>
  );
}

export function DocumentTableLayoutRibbon({ editor }: { editor: Editor }) {
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
      <RibbonGroup label="行">
        <RibbonButton
          label="在上方插入行"
          visibleLabel="上方插入"
          disabled={!editor.can().chain().focus().addRowBefore().run()}
          onClick={() => editor.chain().focus().addRowBefore().run()}
        >
          <BetweenHorizontalStart size={18} />
        </RibbonButton>
        <RibbonButton
          label="在下方插入行"
          visibleLabel="下方插入"
          disabled={!editor.can().chain().focus().addRowAfter().run()}
          onClick={() => editor.chain().focus().addRowAfter().run()}
        >
          <BetweenHorizontalEnd size={18} />
        </RibbonButton>
        <RibbonButton
          label="删除当前行"
          visibleLabel="删除行"
          disabled={!editor.can().chain().focus().deleteRow().run()}
          onClick={() => editor.chain().focus().deleteRow().run()}
        >
          <Rows3 size={18} />
        </RibbonButton>
      </RibbonGroup>
      <RibbonGroup label="列">
        <RibbonButton
          label="在左侧插入列"
          visibleLabel="左侧插入"
          disabled={!editor.can().chain().focus().addColumnBefore().run()}
          onClick={() => editor.chain().focus().addColumnBefore().run()}
        >
          <BetweenVerticalStart size={18} />
        </RibbonButton>
        <RibbonButton
          label="在右侧插入列"
          visibleLabel="右侧插入"
          disabled={!editor.can().chain().focus().addColumnAfter().run()}
          onClick={() => editor.chain().focus().addColumnAfter().run()}
        >
          <BetweenVerticalEnd size={18} />
        </RibbonButton>
        <RibbonButton
          label="删除当前列"
          visibleLabel="删除列"
          disabled={!editor.can().chain().focus().deleteColumn().run()}
          onClick={() => editor.chain().focus().deleteColumn().run()}
        >
          <Trash2 size={18} />
        </RibbonButton>
      </RibbonGroup>
      <RibbonGroup label="合并">
        <RibbonButton
          label="合并单元格"
          visibleLabel="合并"
          disabled={!editor.can().chain().focus().mergeCells().run()}
          onClick={() => editor.chain().focus().mergeCells().run()}
        >
          <TableCellsMerge size={18} />
        </RibbonButton>
        <RibbonButton
          label="拆分单元格"
          visibleLabel="拆分"
          disabled={!editor.can().chain().focus().splitCell().run()}
          onClick={() => editor.chain().focus().splitCell().run()}
        >
          <TableCellsSplit size={18} />
        </RibbonButton>
      </RibbonGroup>
      <DocumentTableSizeRibbonGroup editor={editor} sizing={sizing} />
      <RibbonGroup label="单元格对齐">
        <RibbonButton
          label="单元格水平左对齐"
          visibleLabel="左对齐"
          active={horizontalAlignment === 'left'}
          onClick={() =>
            editor
              .chain()
              .focus()
              .setDocumentTableHorizontalAlignment('left')
              .run()
          }
        >
          <AlignLeft size={18} />
        </RibbonButton>
        <RibbonButton
          label="单元格水平居中"
          visibleLabel="水平居中"
          active={horizontalAlignment === 'center'}
          onClick={() =>
            editor
              .chain()
              .focus()
              .setDocumentTableHorizontalAlignment('center')
              .run()
          }
        >
          <AlignCenter size={18} />
        </RibbonButton>
        <RibbonButton
          label="单元格水平右对齐"
          visibleLabel="右对齐"
          active={horizontalAlignment === 'right'}
          onClick={() =>
            editor
              .chain()
              .focus()
              .setDocumentTableHorizontalAlignment('right')
              .run()
          }
        >
          <AlignRight size={18} />
        </RibbonButton>
        <RibbonButton
          label="单元格顶端对齐"
          visibleLabel="顶端"
          active={cellFormat.verticalAlign === 'top'}
          onClick={() => setVerticalAlignment(editor, 'top')}
        >
          <AlignVerticalJustifyStart size={18} />
        </RibbonButton>
        <RibbonButton
          label="单元格垂直居中"
          visibleLabel="垂直居中"
          active={cellFormat.verticalAlign === 'middle'}
          onClick={() => setVerticalAlignment(editor, 'middle')}
        >
          <AlignVerticalJustifyCenter size={18} />
        </RibbonButton>
        <RibbonButton
          label="单元格底端对齐"
          visibleLabel="底端"
          active={cellFormat.verticalAlign === 'bottom'}
          onClick={() => setVerticalAlignment(editor, 'bottom')}
        >
          <AlignVerticalJustifyEnd size={18} />
        </RibbonButton>
      </RibbonGroup>
      <RibbonGroup label="表格选项">
        <RibbonButton
          label="跨页重复标题"
          visibleLabel="重复标题"
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
          label="整行不跨页"
          visibleLabel="整行换页"
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
      <RibbonGroup label="表格">
        <RibbonButton
          label="删除表格"
          visibleLabel="删除表格"
          disabled={!editor.can().chain().focus().deleteTable().run()}
          onClick={() => editor.chain().focus().deleteTable().run()}
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
  const measuredRowSelection = measuredTableSelectionSize(editor, 'rows');
  const measuredColumnSelection = measuredTableSelectionSize(editor, 'columns');
  const canDistributeRows = editor
    .can()
    .chain()
    .focus()
    .distributeDocumentTableRows(measuredRowSelection)
    .run();
  const canDistributeColumns = editor
    .can()
    .chain()
    .focus()
    .distributeDocumentTableColumns(measuredColumnSelection)
    .run();
  return (
    <RibbonGroup label="单元格大小">
      <div className="work-document-table-size-fields">
        <TableDimensionField
          label="高度"
          ariaLabel="行高（厘米）"
          value={
            sizing?.rowHeight ?? measuredCurrentTableDimension(editor, 'rows')
          }
          onValueChange={(height) =>
            editor
              .chain()
              .focus()
              .setDocumentTableRowHeight(
                height,
                sizing?.rowHeightRule ?? 'atLeast',
              )
              .run()
          }
        />
        <TableDimensionField
          label="宽度"
          ariaLabel="列宽（厘米）"
          value={
            sizing?.columnWidth ??
            measuredCurrentTableDimension(editor, 'columns')
          }
          onValueChange={(width) =>
            editor
              .chain()
              .focus()
              .setDocumentTableColumnWidth(
                width,
                measuredTableColumnWidths(editor),
              )
              .run()
          }
        />
      </div>
      <OfficeSelect
        ariaLabel="表格自动调整"
        className="work-document-table-layout-select"
        value={sizing?.layoutMode ?? 'window'}
        options={tableLayoutOptions}
        onValueChange={(layoutMode) =>
          editor
            .chain()
            .focus()
            .setDocumentTableLayoutMode(
              layoutMode,
              layoutMode === 'fixed' ? measuredTableWidth(editor) : undefined,
            )
            .run()
        }
      />
      <RibbonButton
        label="平均分布行"
        visibleLabel="平均行高"
        disabled={!canDistributeRows}
        onClick={() =>
          editor
            .chain()
            .focus()
            .distributeDocumentTableRows(measuredRowSelection)
            .run()
        }
      >
        <Rows3 size={18} />
      </RibbonButton>
      <RibbonButton
        label="平均分布列"
        visibleLabel="平均列宽"
        disabled={!canDistributeColumns}
        onClick={() =>
          editor
            .chain()
            .focus()
            .distributeDocumentTableColumns(measuredColumnSelection)
            .run()
        }
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
  const formattedValue =
    value === null ? '' : formatCentimeters(value / PIXELS_PER_CENTIMETER);
  const {
    cancelDraft,
    dirty,
    draft,
    replaceDraft,
    setDraft,
    syncDraft,
  } = useOfficeDraft(() => formattedValue);
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
      <small aria-hidden="true">厘米</small>
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
  const groupName = useId();
  const inputsRef = useRef<Array<HTMLInputElement | null>>([]);
  const activeStyle = activeDocumentTableStyle(editor.state);

  const applyStyle = (style: DocumentTableStyleOption) =>
    editor.chain().focus().applyDocumentTableStyle(style.id).run();
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
      aria-label="表格样式库"
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
              aria-label={`应用表格样式：${style.label}`}
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
  editor.chain().focus().setDocumentTableCellFormat({ verticalAlign }).run();
}

function borderOptionValue(
  style: DocumentTableBorderStyle,
  width: number,
): string {
  return `${style}-${width}`;
}

function borderOptionsForFormat(
  style: DocumentTableBorderStyle,
  width: number,
) {
  const value = borderOptionValue(style, width);
  if (borderOptions.some((option) => option.value === value)) {
    return borderOptions;
  }
  return [
    ...borderOptions,
    {
      value,
      label:
        style === 'none'
          ? '无边框'
          : `${width} 像素${documentTableBorderStyleLabel(style)}`,
      style,
      width,
    },
  ];
}

function documentTableBorderStyleLabel(
  style: Exclude<DocumentTableBorderStyle, 'none'>,
): string {
  if (style === 'dashed') return '虚线';
  if (style === 'dotted') return '点线';
  if (style === 'double') return '双线';
  return '实线';
}

function RibbonButton({
  label,
  visibleLabel,
  active = false,
  disabled = false,
  onClick,
  children,
}: {
  label: string;
  visibleLabel: string;
  active?: boolean;
  disabled?: boolean;
  onClick: () => void;
  children: ReactNode;
}) {
  return (
    <WorkOfficeRibbonButton
      label={label}
      visibleLabel={visibleLabel}
      active={active}
      disabled={disabled}
      onClick={onClick}
    >
      {children}
    </WorkOfficeRibbonButton>
  );
}

const RibbonGroup = WorkOfficeRibbonGroup;
