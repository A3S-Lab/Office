import type { Editor } from '@tiptap/core';
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
import { OfficeColorPicker, OfficeSelect } from './office-controls';
import {
  WorkOfficeRibbonButton,
  WorkOfficeRibbonGroup,
} from './work-office-chrome';

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

export function DocumentTableDesignRibbon({ editor }: { editor: Editor }) {
  const format =
    documentTableCellFormat(editor.state) ?? DEFAULT_DOCUMENT_TABLE_CELL_FORMAT;
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
          value={borderOptionValue(format.borderStyle, format.borderWidth)}
          options={borderOptions}
          onValueChange={(value) => {
            const option = borderOptions.find(
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
          onClick={() =>
            editor.commands.setDocumentTableRowOptions({
              ...documentTableRowOptions(editor),
              cantSplit: !documentTableRowOptions(editor).cantSplit,
            })
          }
        >
          <BetweenHorizontalEnd size={18} />
        </RibbonButton>
        <RibbonButton
          label="平均分布列"
          visibleLabel="平均列宽"
          onClick={() =>
            editor.chain().focus().distributeDocumentTableColumns().run()
          }
        >
          <Columns3 size={18} />
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
              onChange={() => undefined}
              onClick={() => applyStyle(style)}
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
): (typeof borderOptions)[number]['value'] {
  return (
    borderOptions.find(
      (option) => option.style === style && option.width === width,
    ) ?? borderOptions[0]
  ).value;
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
