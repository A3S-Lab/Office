import type { Editor } from '@tiptap/core';
import {
  BetweenHorizontalEnd,
  BetweenHorizontalStart,
  BetweenVerticalEnd,
  BetweenVerticalStart,
  PanelTop,
  Rows3,
  TableCellsMerge,
  TableCellsSplit,
  Trash2,
} from 'lucide-react';
import type { ReactNode } from 'react';
import {
  canSetDocumentTableRowRepeatHeader,
  documentTableRowOptions,
  setDocumentTableRowOptions,
} from '../work-document-table-row';
import {
  WorkOfficeRibbonButton,
  WorkOfficeRibbonGroup,
} from './work-office-chrome';

export function DocumentTableRibbon({ editor }: { editor: Editor }) {
  const rowOptions = documentTableRowOptions(editor);
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
      <RibbonGroup label="单元格">
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
        <RibbonButton
          label="跨页重复标题"
          visibleLabel="重复标题"
          active={rowOptions.repeatHeader}
          disabled={!canSetDocumentTableRowRepeatHeader(editor)}
          onClick={() =>
            setDocumentTableRowOptions(editor, {
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
            setDocumentTableRowOptions(editor, {
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
