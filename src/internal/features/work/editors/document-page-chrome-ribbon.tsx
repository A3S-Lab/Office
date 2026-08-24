import type { Editor } from '@tiptap/core';
import {
  AlignCenter,
  AlignJustify,
  AlignLeft,
  AlignRight,
  Bold,
  Hash,
  Image as ImageIcon,
  Italic,
  Link2,
  Languages,
  PanelBottom,
  PanelTop,
  Redo2,
  Subscript as SubscriptIcon,
  Superscript as SuperscriptIcon,
  Undo2,
  X,
} from 'lucide-react';
import { type ReactNode, useEffect, useRef, useState } from 'react';
import { DOCUMENT_LINK_VALIDATION_MESSAGE } from '../work-document-links';
import {
  type DocumentPageChromeAlignment,
  documentPageChromeEditorState,
  loadDocumentPageChromeImage,
  normalizeDocumentPageChromeHref,
} from './document-page-chrome-editor';
import {
  type DocumentCommandId,
  getDocumentCommandDefinition,
} from './document-command-catalog';
import {
  OfficeColorPicker,
  OfficeFileInput,
  useOfficeDialog,
} from './office-controls';
import {
  WorkOfficeRibbonButton,
  WorkOfficeRibbonGroup,
} from './work-office-chrome';
import { DocumentStrikeRibbon } from './document-strike-ribbon';
import { DocumentUnderlineRibbon } from './document-underline-ribbon';

export type DocumentPageChromeEditingPart = 'footer' | 'header';

export function DocumentPageChromeRibbon({
  editor,
  editingPart,
  showPageNumber,
  onEditingPartChange,
  onTogglePageNumber,
  onOpenFontDialog,
  onOpenProofingDialog,
  onClose,
}: {
  editor: Editor;
  editingPart: DocumentPageChromeEditingPart;
  showPageNumber: boolean;
  onEditingPartChange: (part: DocumentPageChromeEditingPart) => void;
  onTogglePageNumber: () => void;
  onOpenFontDialog?: () => void;
  onOpenProofingDialog?: () => void;
  onClose: () => void;
}) {
  const imageInputRef = useRef<HTMLInputElement>(null);
  const [, setRevision] = useState(0);
  const officeDialog = useOfficeDialog();

  useEffect(() => {
    const refresh = () => setRevision((revision) => revision + 1);
    editor.on('transaction', refresh);
    return () => {
      editor.off('transaction', refresh);
    };
  }, [editor]);

  const state = documentPageChromeEditorState(editor);
  const editLink = async () => {
    if (state.link) {
      editor.chain().focus().setDocumentPageChromeLink(null).run();
      return;
    }
    const href = await officeDialog.prompt({
      title: '添加链接',
      fieldLabel: '链接地址',
      initialValue: 'https://',
      placeholder: 'https://',
      inputMode: 'url',
      confirmLabel: '添加链接',
      required: '请输入链接地址。',
      validate: (value) =>
        normalizeDocumentPageChromeHref(value)
          ? null
          : DOCUMENT_LINK_VALIDATION_MESSAGE,
      restoreFocusTarget: () => editor.view.dom,
    });
    if (href === null) return;
    const normalized = normalizeDocumentPageChromeHref(href);
    if (normalized && !editor.isDestroyed) {
      editor.chain().focus().setDocumentPageChromeLink(normalized).run();
    }
  };
  const insertImage = async (file: File | undefined) => {
    if (!file || editor.isDestroyed) return;
    const image = await loadDocumentPageChromeImage(file);
    if (!image.ok) {
      await officeDialog.notice({
        title: image.title,
        description: image.description,
      });
      return;
    }
    if (!editor.isDestroyed) {
      editor
        .chain()
        .focus()
        .insertDocumentPageChromeImage({
          alt: image.alt,
          source: image.source,
        })
        .run();
    }
  };

  return (
    <>
      <WorkOfficeRibbonGroup label="位置">
        <PageChromeRibbonButton
          label="切换到页眉"
          displayLabel
          active={editingPart === 'header'}
          onClick={() => onEditingPartChange('header')}
        >
          <PanelTop size={18} />
        </PageChromeRibbonButton>
        <PageChromeRibbonButton
          label="切换到页脚"
          displayLabel
          active={editingPart === 'footer'}
          onClick={() => onEditingPartChange('footer')}
        >
          <PanelBottom size={18} />
        </PageChromeRibbonButton>
      </WorkOfficeRibbonGroup>
      <WorkOfficeRibbonGroup label="撤销">
        <PageChromeRibbonButton
          label="撤销页眉页脚编辑"
          shortcut="Cmd/Ctrl+Z"
          ariaKeyShortcuts="Control+Z Meta+Z"
          disabled={!state.canUndo}
          onClick={() => editor.chain().focus().undo().run()}
        >
          <Undo2 size={16} />
        </PageChromeRibbonButton>
        <PageChromeRibbonButton
          label="重做页眉页脚编辑"
          shortcut="Cmd/Ctrl+Shift+Z 或 Cmd/Ctrl+Y"
          ariaKeyShortcuts="Control+Shift+Z Meta+Shift+Z Control+Y Meta+Y"
          disabled={!state.canRedo}
          onClick={() => editor.chain().focus().redo().run()}
        >
          <Redo2 size={16} />
        </PageChromeRibbonButton>
      </WorkOfficeRibbonGroup>
      <WorkOfficeRibbonGroup
        label="文字"
        dialogLauncher={
          onOpenFontDialog
            ? {
                label: '页眉页脚字体高级设置',
                ...pageChromeCommandShortcut('fontDialog'),
                onClick: onOpenFontDialog,
              }
            : undefined
        }
      >
        <PageChromeRibbonButton
          label="页眉页脚加粗"
          shortcut="Cmd/Ctrl+B"
          ariaKeyShortcuts="Control+B Meta+B"
          active={state.bold}
          onClick={() => editor.chain().focus().toggleBold().run()}
        >
          <Bold size={16} />
        </PageChromeRibbonButton>
        <PageChromeRibbonButton
          label="页眉页脚斜体"
          shortcut="Cmd/Ctrl+I"
          ariaKeyShortcuts="Control+I Meta+I"
          active={state.italic}
          onClick={() => editor.chain().focus().toggleItalic().run()}
        >
          <Italic size={16} />
        </PageChromeRibbonButton>
        <DocumentUnderlineRibbon
          editor={editor}
          label="页眉页脚下划线"
          menuLabel="页眉页脚下划线样式"
          colorLabel="页眉页脚下划线颜色"
          className="work-document-page-chrome-underline"
        />
        <DocumentStrikeRibbon
          editor={editor}
          label="页眉页脚删除线"
          menuLabel="页眉页脚删除线样式"
          className="work-document-page-chrome-strike"
        />
        <PageChromeRibbonButton
          label="页眉页脚下标"
          {...pageChromeCommandShortcut('subscript')}
          active={state.subscript}
          onClick={() => editor.chain().focus().toggleDocumentSubscript().run()}
        >
          <SubscriptIcon size={16} />
        </PageChromeRibbonButton>
        <PageChromeRibbonButton
          label="页眉页脚上标"
          {...pageChromeCommandShortcut('superscript')}
          active={state.superscript}
          onClick={() =>
            editor.chain().focus().toggleDocumentSuperscript().run()
          }
        >
          <SuperscriptIcon size={16} />
        </PageChromeRibbonButton>
        <OfficeColorPicker
          compact
          className="work-document-page-chrome-ribbon-color"
          ariaLabel="页眉页脚文字颜色"
          value={pickerColor(state.color)}
          onValueChange={(color) =>
            editor.chain().focus().setColor(color).run()
          }
        />
        {onOpenProofingDialog && (
          <PageChromeRibbonButton
            label="页眉页脚校对语言"
            onClick={onOpenProofingDialog}
          >
            <Languages size={16} />
          </PageChromeRibbonButton>
        )}
      </WorkOfficeRibbonGroup>
      <WorkOfficeRibbonGroup label="对齐">
        {(['left', 'center', 'right', 'justify'] as const).map((alignment) => (
          <PageChromeRibbonButton
            key={alignment}
            label={alignmentLabel(alignment)}
            {...pageChromeCommandShortcut(
              pageChromeAlignmentCommandIds[alignment],
            )}
            active={state.alignment === alignment}
            onClick={() => editor.chain().focus().setTextAlign(alignment).run()}
          >
            {alignmentIcon(alignment)}
          </PageChromeRibbonButton>
        ))}
      </WorkOfficeRibbonGroup>
      <WorkOfficeRibbonGroup label="插入">
        <PageChromeRibbonButton
          label={state.link ? '移除页眉页脚链接' : '添加页眉页脚链接'}
          active={Boolean(state.link)}
          onClick={() => void editLink()}
        >
          <Link2 size={17} />
        </PageChromeRibbonButton>
        <PageChromeRibbonButton
          label="插入页眉页脚图片"
          onClick={() => imageInputRef.current?.click()}
        >
          <ImageIcon size={17} />
        </PageChromeRibbonButton>
      </WorkOfficeRibbonGroup>
      <WorkOfficeRibbonGroup label="页码">
        <PageChromeRibbonButton
          label="显示页码"
          displayLabel
          active={showPageNumber}
          onClick={onTogglePageNumber}
        >
          <Hash size={18} />
        </PageChromeRibbonButton>
      </WorkOfficeRibbonGroup>
      <WorkOfficeRibbonGroup label="关闭">
        <PageChromeRibbonButton
          label="关闭页眉和页脚"
          displayLabel
          onClick={onClose}
        >
          <X size={18} />
        </PageChromeRibbonButton>
      </WorkOfficeRibbonGroup>
      <OfficeFileInput
        ref={imageInputRef}
        accept="image/bmp,image/gif,image/jpeg,image/png,image/webp"
        aria-label="页眉页脚图片文件"
        onFileSelect={insertImage}
      />
      {officeDialog.dialog}
    </>
  );
}

function PageChromeRibbonButton({
  label,
  shortcut,
  ariaKeyShortcuts,
  active = false,
  disabled = false,
  displayLabel = false,
  onClick,
  children,
}: {
  label: string;
  shortcut?: string;
  ariaKeyShortcuts?: string;
  active?: boolean;
  disabled?: boolean;
  displayLabel?: boolean;
  onClick: () => void;
  children: ReactNode;
}) {
  return (
    <WorkOfficeRibbonButton
      label={label}
      title={shortcut ? `${label}（${shortcut}）` : label}
      aria-keyshortcuts={ariaKeyShortcuts}
      active={active}
      disabled={disabled}
      displayLabel={displayLabel}
      onMouseDown={(event) => event.preventDefault()}
      onClick={onClick}
    >
      {children}
    </WorkOfficeRibbonButton>
  );
}

function pickerColor(color: string): string {
  return /^#[\da-f]{6}$/i.test(color) ? color : '#4d5668';
}

function pageChromeCommandShortcut(commandId: DocumentCommandId): {
  shortcut?: string;
  ariaKeyShortcuts?: string;
} {
  const shortcut = getDocumentCommandDefinition(commandId).shortcut;
  return shortcut
    ? { shortcut: shortcut.label, ariaKeyShortcuts: shortcut.aria }
    : {};
}

const pageChromeAlignmentCommandIds = {
  center: 'alignCenter',
  justify: 'alignJustify',
  left: 'alignLeft',
  right: 'alignRight',
} as const satisfies Record<DocumentPageChromeAlignment, DocumentCommandId>;

function alignmentLabel(
  alignment: 'center' | 'justify' | 'left' | 'right',
): string {
  switch (alignment) {
    case 'center':
      return '页眉页脚居中';
    case 'justify':
      return '页眉页脚两端对齐';
    case 'left':
      return '页眉页脚左对齐';
    case 'right':
      return '页眉页脚右对齐';
  }
}

function alignmentIcon(
  alignment: 'center' | 'justify' | 'left' | 'right',
): ReactNode {
  switch (alignment) {
    case 'center':
      return <AlignCenter size={16} />;
    case 'justify':
      return <AlignJustify size={16} />;
    case 'left':
      return <AlignLeft size={16} />;
    case 'right':
      return <AlignRight size={16} />;
  }
}
