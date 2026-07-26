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
  PilcrowLeft,
  PilcrowRight,
  Redo2,
  Replace,
  Search,
  Strikethrough,
  Subscript as SubscriptIcon,
  Superscript as SuperscriptIcon,
  Underline as UnderlineIcon,
  Undo2,
} from 'lucide-react';
import type { ReactNode } from 'react';
import { documentParagraphDirection } from '../work-document-paragraph-formatting';
import {
  changeDocumentFontSize,
  documentFontFamilyOptions,
  documentFontFamilyValue,
  documentFontSizeOptions,
  documentFontSizeValue,
} from './document-formatting-options';
import { OfficeColorPicker, OfficeSelect } from './office-controls';
import {
  WorkOfficeRibbonButton,
  WorkOfficeRibbonGroup,
} from './work-office-chrome';
import type { DocumentFindReplaceMode } from './document-find-replace-panel';
import { DocumentListGallery } from './document-list-gallery';
import { DocumentStyleGallery } from './document-style-gallery';

const documentLineHeightOptions = [
  { value: 'default', label: '默认行距' },
  { value: '1', label: '单倍' },
  { value: '1.15', label: '1.15 倍' },
  { value: '1.5', label: '1.5 倍' },
  { value: '2', label: '2 倍' },
] as const;

export function DocumentHomeRibbon({
  editor,
  findReplaceMode,
  onFindText,
}: {
  editor: Editor;
  findReplaceMode: DocumentFindReplaceMode | null;
  onFindText: (replace: boolean) => void;
}) {
  if (editor.isDestroyed) return null;

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
            <ToolbarButton
              label="下标"
              shortcut="Cmd/Ctrl+,"
              active={editor.isActive('subscript')}
              onClick={() => editor.commands.toggleDocumentSubscript()}
            >
              <SubscriptIcon size={16} />
            </ToolbarButton>
            <ToolbarButton
              label="上标"
              shortcut="Cmd/Ctrl+."
              active={editor.isActive('superscript')}
              onClick={() => editor.commands.toggleDocumentSuperscript()}
            >
              <SuperscriptIcon size={16} />
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
              onClick={() => editor.commands.clearDocumentFormatting()}
            >
              <Eraser size={16} />
            </ToolbarButton>
          </div>
        </div>
      </RibbonGroup>
      <RibbonGroup label="段落">
        <div className="work-document-paragraph-tools">
          <div className="work-document-paragraph-actions">
            <DocumentListGallery editor={editor} />
            <ToolbarButton
              label="减少缩进"
              onClick={() => editor.commands.changeDocumentIndent(-1)}
            >
              <IndentDecrease size={16} />
            </ToolbarButton>
            <ToolbarButton
              label="增加缩进"
              onClick={() => editor.commands.changeDocumentIndent(1)}
            >
              <IndentIncrease size={16} />
            </ToolbarButton>
            <OfficeSelect
              ariaLabel="行距"
              className="work-document-line-height-select"
              value={documentLineHeightValue(editor)}
              options={documentLineHeightOptions}
              onValueChange={(value) =>
                editor.commands.setDocumentLineHeight(
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
            <ToolbarButton
              label="从左向右"
              active={documentParagraphDirection(editor) === 'ltr'}
              onClick={() =>
                editor.commands.setDocumentParagraphDirection('ltr')
              }
            >
              <PilcrowRight size={16} />
            </ToolbarButton>
            <ToolbarButton
              label="从右向左"
              active={documentParagraphDirection(editor) === 'rtl'}
              onClick={() =>
                editor.commands.setDocumentParagraphDirection('rtl')
              }
            >
              <PilcrowLeft size={16} />
            </ToolbarButton>
          </div>
        </div>
      </RibbonGroup>
      <RibbonGroup label="样式">
        <DocumentStyleGallery editor={editor} />
      </RibbonGroup>
      <RibbonGroup label="编辑">
        <ToolbarButton
          label="查找"
          shortcut="Cmd/Ctrl+F"
          active={findReplaceMode === 'find'}
          onClick={() => onFindText(false)}
        >
          <Search size={16} />
        </ToolbarButton>
        <ToolbarButton
          label="替换"
          shortcut="Cmd/Ctrl+H"
          active={findReplaceMode === 'replace'}
          onClick={() => onFindText(true)}
        >
          <Replace size={16} />
        </ToolbarButton>
      </RibbonGroup>
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
