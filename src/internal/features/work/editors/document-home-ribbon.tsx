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
  Replace,
  Search,
  Subscript as SubscriptIcon,
  Superscript as SuperscriptIcon,
} from 'lucide-react';
import { type ReactNode, useCallback, useSyncExternalStore } from 'react';
import {
  canChangeDocumentIndent,
  documentParagraphDirection,
} from '../work-document-paragraph-formatting';
import type { WorkDocumentLayoutFont } from '../work-document-fonts';
import {
  canChangeDocumentFontSize,
  changeDocumentFontSize,
  documentFontFamilyOptionsForValue,
  documentFontFamilyValue,
  documentFontSizeOptionsForValue,
  documentFontSizeValue,
} from './document-formatting-options';
import {
  type DocumentCommandId,
  getDocumentCommandDefinition,
} from './document-command-catalog';
import { DocumentFormatTools } from './document-format-tools';
import { OfficeColorPicker, OfficeSelect } from './office-controls';
import {
  WorkOfficeRibbonButton,
  WorkOfficeRibbonGroup,
} from './work-office-chrome';
import type { DocumentFindReplaceMode } from './document-find-replace-panel';
import { DocumentListGallery } from './document-list-gallery';
import { DocumentStyleGallery } from './document-style-gallery';
import { DocumentTextCaseRibbon } from './document-text-case-ribbon';
import { DocumentStrikeRibbon } from './document-strike-ribbon';
import { DocumentUnderlineRibbon } from './document-underline-ribbon';

const documentLineHeightOptions = [
  { value: 'default', label: '默认行距' },
  {
    value: '1',
    label: '单倍',
    meta: getDocumentCommandDefinition('lineSpacingSingle').shortcut?.label,
  },
  { value: '1.15', label: '1.15 倍' },
  {
    value: '1.5',
    label: '1.5 倍',
    meta: getDocumentCommandDefinition('lineSpacingOneAndHalf').shortcut?.label,
  },
  {
    value: '2',
    label: '2 倍',
    meta: getDocumentCommandDefinition('lineSpacingDouble').shortcut?.label,
  },
] as const;

export function DocumentHomeRibbon({
  editor,
  findReplaceMode,
  layoutFonts = [],
  onFindText,
}: {
  editor: Editor;
  findReplaceMode: DocumentFindReplaceMode | null;
  layoutFonts?: readonly WorkDocumentLayoutFont[];
  onFindText: (replace: boolean) => void;
}) {
  const subscribe = useCallback(
    (notify: () => void) => {
      if (editor.isDestroyed) return () => undefined;
      editor.on('transaction', notify);
      return () => editor.off('transaction', notify);
    },
    [editor],
  );
  useSyncExternalStore(
    subscribe,
    () => editor.state,
    () => editor.state,
  );
  if (editor.isDestroyed) return null;
  const fontFamilyValue = documentFontFamilyValue(editor, layoutFonts);
  const fontSizeValue = documentFontSizeValue(editor);
  const lineHeightValue = documentLineHeightValue(editor);

  return (
    <>
      <DocumentFormatTools editor={editor} />
      <RibbonGroup label="字体" priority="high">
        <div className="work-document-font-tools">
          <div className="work-document-font-selects">
            <OfficeSelect
              ariaLabel="字体"
              className="work-document-font-family-select"
              value={fontFamilyValue}
              options={documentFontFamilyOptionsForValue(
                fontFamilyValue,
                layoutFonts,
              )}
              onValueChange={(value) => {
                if (value === 'default')
                  editor.chain().focus().unsetFontFamily().run();
                else editor.chain().focus().setFontFamily(value).run();
              }}
            />
            <OfficeSelect
              ariaLabel="字号"
              className="work-document-font-size-select"
              value={fontSizeValue}
              options={documentFontSizeOptionsForValue(fontSizeValue)}
              onValueChange={(value) => {
                if (value === 'default')
                  editor.chain().focus().unsetFontSize().run();
                else editor.chain().focus().setFontSize(value).run();
              }}
            />
            <ToolbarButton
              label="增大字号"
              {...commandShortcut('growFont')}
              disabled={!canChangeDocumentFontSize(editor, 1)}
              onClick={() => changeDocumentFontSize(editor, 1)}
            >
              <AArrowUp size={16} />
            </ToolbarButton>
            <ToolbarButton
              label="减小字号"
              {...commandShortcut('shrinkFont')}
              disabled={!canChangeDocumentFontSize(editor, -1)}
              onClick={() => changeDocumentFontSize(editor, -1)}
            >
              <AArrowDown size={16} />
            </ToolbarButton>
          </div>
          <div className="work-document-font-actions">
            <ToolbarButton
              label="加粗"
              {...commandShortcut('bold')}
              active={editor.isActive('bold')}
              onClick={() => editor.chain().focus().toggleBold().run()}
            >
              <Bold size={16} />
            </ToolbarButton>
            <ToolbarButton
              label="斜体"
              {...commandShortcut('italic')}
              active={editor.isActive('italic')}
              onClick={() => editor.chain().focus().toggleItalic().run()}
            >
              <Italic size={16} />
            </ToolbarButton>
            <DocumentUnderlineRibbon editor={editor} />
            <DocumentStrikeRibbon editor={editor} />
            <DocumentTextCaseRibbon editor={editor} />
            <ToolbarButton
              label="下标"
              {...commandShortcut('subscript')}
              active={editor.isActive('subscript')}
              onClick={() => editor.commands.toggleDocumentSubscript()}
            >
              <SubscriptIcon size={16} />
            </ToolbarButton>
            <ToolbarButton
              label="上标"
              {...commandShortcut('superscript')}
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
      <RibbonGroup label="段落" priority="high">
        <div className="work-document-paragraph-tools">
          <div className="work-document-paragraph-actions">
            <DocumentListGallery editor={editor} />
            <ToolbarButton
              label="减少缩进"
              disabled={!canChangeDocumentIndent(editor, -1)}
              onClick={() => editor.commands.changeDocumentIndent(-1)}
            >
              <IndentDecrease size={16} />
            </ToolbarButton>
            <ToolbarButton
              label="增加缩进"
              disabled={!canChangeDocumentIndent(editor, 1)}
              onClick={() => editor.commands.changeDocumentIndent(1)}
            >
              <IndentIncrease size={16} />
            </ToolbarButton>
            <OfficeSelect
              ariaLabel="行距"
              className="work-document-line-height-select"
              value={lineHeightValue}
              options={documentLineHeightOptionsForValue(lineHeightValue)}
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
              {...commandShortcut('alignLeft')}
              active={editor.isActive({ textAlign: 'left' })}
              onClick={() => editor.chain().focus().setTextAlign('left').run()}
            >
              <AlignLeft size={16} />
            </ToolbarButton>
            <ToolbarButton
              label="居中"
              {...commandShortcut('alignCenter')}
              active={editor.isActive({ textAlign: 'center' })}
              onClick={() =>
                editor.chain().focus().setTextAlign('center').run()
              }
            >
              <AlignCenter size={16} />
            </ToolbarButton>
            <ToolbarButton
              label="右对齐"
              {...commandShortcut('alignRight')}
              active={editor.isActive({ textAlign: 'right' })}
              onClick={() => editor.chain().focus().setTextAlign('right').run()}
            >
              <AlignRight size={16} />
            </ToolbarButton>
            <ToolbarButton
              label="两端对齐"
              {...commandShortcut('alignJustify')}
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
      <RibbonGroup label="样式" priority="low">
        <DocumentStyleGallery editor={editor} />
      </RibbonGroup>
      <RibbonGroup label="编辑" priority="low">
        <ToolbarButton
          label="查找"
          {...commandShortcut('find')}
          active={findReplaceMode === 'find'}
          onClick={() => onFindText(false)}
        >
          <Search size={16} />
        </ToolbarButton>
        <ToolbarButton
          label="替换"
          {...commandShortcut('replace')}
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
  ariaKeyShortcuts,
  active = false,
  disabled = false,
  onClick,
  children,
}: {
  label: string;
  shortcut?: string;
  ariaKeyShortcuts?: string;
  active?: boolean;
  disabled?: boolean;
  onClick: () => void;
  children: ReactNode;
}) {
  return (
    <WorkOfficeRibbonButton
      label={label}
      title={shortcut ? `${label}（${shortcut}）` : label}
      aria-keyshortcuts={ariaKeyShortcuts}
      active={active}
      displayLabel={false}
      disabled={disabled}
      onClick={onClick}
    >
      {children}
    </WorkOfficeRibbonButton>
  );
}

function commandShortcut(commandId: DocumentCommandId): {
  shortcut?: string;
  ariaKeyShortcuts?: string;
} {
  const shortcut = getDocumentCommandDefinition(commandId).shortcut;
  return shortcut
    ? { shortcut: shortcut.label, ariaKeyShortcuts: shortcut.aria }
    : {};
}

const RibbonGroup = WorkOfficeRibbonGroup;

function documentLineHeightValue(editor: Editor): string {
  const attributes = editor.isActive('heading')
    ? editor.getAttributes('heading')
    : editor.getAttributes('paragraph');
  const value = attributes.lineHeight;
  if (typeof value !== 'string' || !value.trim() || value === 'normal') {
    return 'default';
  }
  return value.trim();
}

function documentLineHeightOptionsForValue(value: string) {
  if (documentLineHeightOptions.some((option) => option.value === value)) {
    return documentLineHeightOptions;
  }
  const numeric = Number(value);
  return [
    ...documentLineHeightOptions,
    {
      value,
      label: Number.isFinite(numeric) && numeric > 0 ? `${numeric} 倍` : value,
    },
  ];
}
