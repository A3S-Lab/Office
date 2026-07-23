import type { Editor } from '@tiptap/core';
import {
  AArrowDown,
  AArrowUp,
  AlignCenter,
  AlignJustify,
  AlignLeft,
  AlignRight,
  Bold,
  Eraser,
  Highlighter,
  IndentDecrease,
  IndentIncrease,
  Italic,
  List,
  ListOrdered,
  Redo2,
  Replace,
  Search,
  Strikethrough,
  Underline as UnderlineIcon,
  Undo2,
} from 'lucide-react';
import type { ReactNode } from 'react';
import {
  changeDocumentIndent,
  clearDocumentFormatting,
  setDocumentLineHeight,
} from '../work-document-paragraph-formatting';
import { OfficeColorPicker, OfficeSelect } from './office-controls';
import {
  WorkOfficeRibbonButton,
  WorkOfficeRibbonGroup,
} from './work-office-chrome';

const documentFontFamilyOptions = [
  { value: 'default', label: '默认字体' },
  {
    value: '"Microsoft YaHei", "PingFang SC", sans-serif',
    label: '微软雅黑',
  },
  { value: 'SimSun, "Songti SC", serif', label: '宋体' },
  { value: 'SimHei, "Heiti SC", sans-serif', label: '黑体' },
  { value: 'KaiTi, "Kaiti SC", serif', label: '楷体' },
  { value: 'Arial, sans-serif', label: 'Arial' },
  { value: '"Times New Roman", serif', label: 'Times New Roman' },
] as const;

const documentFontSizeOptions = [
  { value: 'default', label: '10.5' },
  { value: '9pt', label: '9' },
  { value: '12pt', label: '12' },
  { value: '14pt', label: '14' },
  { value: '16pt', label: '16' },
  { value: '18pt', label: '18' },
  { value: '22pt', label: '22' },
  { value: '24pt', label: '24' },
  { value: '36pt', label: '36' },
  { value: '48pt', label: '48' },
  { value: '72pt', label: '72' },
] as const;

const documentFontSizeSteps = [9, 10.5, 12, 14, 16, 18, 22, 24, 36, 48, 72];

const documentLineHeightOptions = [
  { value: 'default', label: '默认行距' },
  { value: '1', label: '单倍' },
  { value: '1.15', label: '1.15 倍' },
  { value: '1.5', label: '1.5 倍' },
  { value: '2', label: '2 倍' },
] as const;

export function DocumentHomeRibbon({
  editor,
  onFindText,
}: {
  editor: Editor;
  onFindText: (replace: boolean) => void;
}) {
  return (
    <>
      <RibbonGroup label="撤销">
        <ToolbarButton
          label="撤销"
          shortcut="Cmd/Ctrl+Z"
          disabled={!editor.can().chain().focus().undo().run()}
          onClick={() => editor.chain().focus().undo().run()}
        >
          <Undo2 size={16} />
        </ToolbarButton>
        <ToolbarButton
          label="重做"
          shortcut="Cmd/Ctrl+Shift+Z"
          disabled={!editor.can().chain().focus().redo().run()}
          onClick={() => editor.chain().focus().redo().run()}
        >
          <Redo2 size={16} />
        </ToolbarButton>
      </RibbonGroup>
      <RibbonGroup label="样式">
        <OfficeSelect
          ariaLabel="段落样式"
          value={
            editor.isActive('heading', { level: 1 })
              ? 'h1'
              : editor.isActive('heading', { level: 2 })
                ? 'h2'
                : editor.isActive('heading', { level: 3 })
                  ? 'h3'
                  : 'paragraph'
          }
          options={[
            { value: 'paragraph', label: '正文' },
            { value: 'h1', label: '标题 1' },
            { value: 'h2', label: '标题 2' },
            { value: 'h3', label: '标题 3' },
          ]}
          onValueChange={(value) => {
            if (value === 'paragraph')
              editor.chain().focus().setParagraph().run();
            else
              editor
                .chain()
                .focus()
                .toggleHeading({
                  level: Number(value.slice(1)) as 1 | 2 | 3,
                })
                .run();
          }}
        />
      </RibbonGroup>
      <RibbonGroup label="字体">
        <div className="work-document-font-tools">
          <div className="work-document-font-selects">
            <OfficeSelect
              ariaLabel="字体"
              className="work-document-font-family-select"
              value={documentFontFamilyValue(editor)}
              options={documentFontFamilyOptions}
              onValueChange={(value) => {
                if (value === 'default')
                  editor.chain().focus().unsetFontFamily().run();
                else editor.chain().focus().setFontFamily(value).run();
              }}
            />
            <OfficeSelect
              ariaLabel="字号"
              className="work-document-font-size-select"
              value={documentFontSizeValue(editor)}
              options={documentFontSizeOptions}
              onValueChange={(value) => {
                if (value === 'default')
                  editor.chain().focus().unsetFontSize().run();
                else editor.chain().focus().setFontSize(value).run();
              }}
            />
            <ToolbarButton
              label="增大字号"
              onClick={() => changeDocumentFontSize(editor, 1)}
            >
              <AArrowUp size={16} />
            </ToolbarButton>
            <ToolbarButton
              label="减小字号"
              onClick={() => changeDocumentFontSize(editor, -1)}
            >
              <AArrowDown size={16} />
            </ToolbarButton>
          </div>
          <div className="work-document-font-actions">
            <ToolbarButton
              label="加粗"
              shortcut="Cmd/Ctrl+B"
              active={editor.isActive('bold')}
              onClick={() => editor.chain().focus().toggleBold().run()}
            >
              <Bold size={16} />
            </ToolbarButton>
            <ToolbarButton
              label="斜体"
              shortcut="Cmd/Ctrl+I"
              active={editor.isActive('italic')}
              onClick={() => editor.chain().focus().toggleItalic().run()}
            >
              <Italic size={16} />
            </ToolbarButton>
            <ToolbarButton
              label="下划线"
              shortcut="Cmd/Ctrl+U"
              active={editor.isActive('underline')}
              onClick={() => editor.chain().focus().toggleUnderline().run()}
            >
              <UnderlineIcon size={16} />
            </ToolbarButton>
            <ToolbarButton
              label="删除线"
              active={editor.isActive('strike')}
              onClick={() => editor.chain().focus().toggleStrike().run()}
            >
              <Strikethrough size={16} />
            </ToolbarButton>
            <OfficeColorPicker
              compact
              className="work-color-tool"
              value={editor.getAttributes('textStyle').color ?? '#172033'}
              ariaLabel="文字颜色"
              onValueChange={(color) =>
                editor.chain().focus().setColor(color).run()
              }
            />
            <ToolbarButton
              label="突出显示"
              active={editor.isActive('highlight')}
              onClick={() =>
                editor
                  .chain()
                  .focus()
                  .toggleHighlight({ color: '#fff0a6' })
                  .run()
              }
            >
              <Highlighter size={16} />
            </ToolbarButton>
            <ToolbarButton
              label="清除格式"
              onClick={() => clearDocumentFormatting(editor)}
            >
              <Eraser size={16} />
            </ToolbarButton>
          </div>
        </div>
      </RibbonGroup>
      <RibbonGroup label="段落">
        <div className="work-document-paragraph-tools">
          <div className="work-document-paragraph-actions">
            <ToolbarButton
              label="项目符号"
              active={editor.isActive('bulletList')}
              onClick={() => editor.chain().focus().toggleBulletList().run()}
            >
              <List size={16} />
            </ToolbarButton>
            <ToolbarButton
              label="编号"
              active={editor.isActive('orderedList')}
              onClick={() => editor.chain().focus().toggleOrderedList().run()}
            >
              <ListOrdered size={16} />
            </ToolbarButton>
            <ToolbarButton
              label="减少缩进"
              onClick={() => changeDocumentIndent(editor, -1)}
            >
              <IndentDecrease size={16} />
            </ToolbarButton>
            <ToolbarButton
              label="增加缩进"
              onClick={() => changeDocumentIndent(editor, 1)}
            >
              <IndentIncrease size={16} />
            </ToolbarButton>
            <OfficeSelect
              ariaLabel="行距"
              className="work-document-line-height-select"
              value={documentLineHeightValue(editor)}
              options={documentLineHeightOptions}
              onValueChange={(value) =>
                setDocumentLineHeight(
                  editor,
                  value === 'default' ? null : value,
                )
              }
            />
          </div>
          <div className="work-document-alignment-actions">
            <ToolbarButton
              label="左对齐"
              active={editor.isActive({ textAlign: 'left' })}
              onClick={() => editor.chain().focus().setTextAlign('left').run()}
            >
              <AlignLeft size={16} />
            </ToolbarButton>
            <ToolbarButton
              label="居中"
              active={editor.isActive({ textAlign: 'center' })}
              onClick={() =>
                editor.chain().focus().setTextAlign('center').run()
              }
            >
              <AlignCenter size={16} />
            </ToolbarButton>
            <ToolbarButton
              label="右对齐"
              active={editor.isActive({ textAlign: 'right' })}
              onClick={() => editor.chain().focus().setTextAlign('right').run()}
            >
              <AlignRight size={16} />
            </ToolbarButton>
            <ToolbarButton
              label="两端对齐"
              active={editor.isActive({ textAlign: 'justify' })}
              onClick={() =>
                editor.chain().focus().setTextAlign('justify').run()
              }
            >
              <AlignJustify size={16} />
            </ToolbarButton>
          </div>
        </div>
      </RibbonGroup>
      <RibbonGroup label="编辑">
        <ToolbarButton
          label="查找"
          shortcut="Cmd/Ctrl+F"
          onClick={() => onFindText(false)}
        >
          <Search size={16} />
        </ToolbarButton>
        <ToolbarButton
          label="替换"
          shortcut="Cmd/Ctrl+H"
          onClick={() => onFindText(true)}
        >
          <Replace size={16} />
        </ToolbarButton>
      </RibbonGroup>
      {editor.isActive('table') && (
        <RibbonGroup label="表格">
          <ToolbarButton
            label="添加行"
            onClick={() => editor.chain().focus().addRowAfter().run()}
          >
            + 行
          </ToolbarButton>
          <ToolbarButton
            label="添加列"
            onClick={() => editor.chain().focus().addColumnAfter().run()}
          >
            + 列
          </ToolbarButton>
          <ToolbarButton
            label="删除表格"
            onClick={() => editor.chain().focus().deleteTable().run()}
          >
            × 表
          </ToolbarButton>
        </RibbonGroup>
      )}
    </>
  );
}

function ToolbarButton({
  label,
  shortcut,
  active = false,
  disabled = false,
  onClick,
  children,
}: {
  label: string;
  shortcut?: string;
  active?: boolean;
  disabled?: boolean;
  onClick: () => void;
  children: ReactNode;
}) {
  return (
    <WorkOfficeRibbonButton
      label={label}
      title={shortcut ? `${label}（${shortcut}）` : label}
      active={active}
      displayLabel={false}
      disabled={disabled}
      onClick={onClick}
    >
      {children}
    </WorkOfficeRibbonButton>
  );
}

const RibbonGroup = WorkOfficeRibbonGroup;

function documentFontFamilyValue(
  editor: Editor,
): (typeof documentFontFamilyOptions)[number]['value'] {
  const value = editor.getAttributes('textStyle').fontFamily;
  return documentFontFamilyOptions.some((option) => option.value === value)
    ? (value as (typeof documentFontFamilyOptions)[number]['value'])
    : 'default';
}

function documentFontSizeValue(
  editor: Editor,
): (typeof documentFontSizeOptions)[number]['value'] {
  const value = editor.getAttributes('textStyle').fontSize;
  if (value === '10.5pt' || !value) return 'default';
  return documentFontSizeOptions.some((option) => option.value === value)
    ? (value as (typeof documentFontSizeOptions)[number]['value'])
    : 'default';
}

function documentLineHeightValue(
  editor: Editor,
): (typeof documentLineHeightOptions)[number]['value'] {
  const attributes = editor.isActive('heading')
    ? editor.getAttributes('heading')
    : editor.getAttributes('paragraph');
  const value = attributes.lineHeight;
  return documentLineHeightOptions.some((option) => option.value === value)
    ? (value as (typeof documentLineHeightOptions)[number]['value'])
    : 'default';
}

function changeDocumentFontSize(editor: Editor, direction: -1 | 1): boolean {
  const current = fontSizePoints(editor.getAttributes('textStyle').fontSize);
  const next =
    direction > 0
      ? (documentFontSizeSteps.find((size) => size > current) ??
        documentFontSizeSteps.at(-1))
      : ([...documentFontSizeSteps].reverse().find((size) => size < current) ??
        documentFontSizeSteps[0]);
  return editor
    .chain()
    .focus()
    .setFontSize(`${next ?? 10.5}pt`)
    .run();
}

function fontSizePoints(value: unknown): number {
  if (typeof value !== 'string') return 10.5;
  const match = /^(\d+(?:\.\d+)?)(px|pt)?$/i.exec(value.trim());
  if (!match) return 10.5;
  const amount = Number(match[1]);
  if (!Number.isFinite(amount)) return 10.5;
  return match[2]?.toLowerCase() === 'px' ? amount * 0.75 : amount;
}
