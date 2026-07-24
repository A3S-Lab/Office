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
  PanelBottom,
  PanelTop,
  Redo2,
  Underline,
  Undo2,
  X,
} from 'lucide-react';
import { type ReactNode, useEffect, useRef, useState } from 'react';
import {
  applyDocumentPageChromeEditorCommand,
  documentPageChromeEditorState,
  loadDocumentPageChromeImage,
  normalizeDocumentPageChromeHref,
} from './document-page-chrome-editor';
import {
  OfficeColorPicker,
  OfficeFileInput,
  useOfficeDialog,
} from './office-controls';
import {
  WorkOfficeRibbonButton,
  WorkOfficeRibbonGroup,
} from './work-office-chrome';

export type DocumentPageChromeEditingPart = 'footer' | 'header';

export function DocumentPageChromeRibbon({
  editor,
  editingPart,
  showPageNumber,
  onEditingPartChange,
  onTogglePageNumber,
  onClose,
}: {
  editor: Editor;
  editingPart: DocumentPageChromeEditingPart;
  showPageNumber: boolean;
  onEditingPartChange: (part: DocumentPageChromeEditingPart) => void;
  onTogglePageNumber: () => void;
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
      applyDocumentPageChromeEditorCommand(editor, {
        type: 'setLink',
        href: null,
      });
      return;
    }
    const href = await officeDialog.prompt({
      title: '链接地址',
      initialValue: 'https://',
      placeholder: 'https://',
      confirmLabel: '添加链接',
    });
    if (href === null) return;
    const normalized = normalizeDocumentPageChromeHref(href);
    if (!normalized) {
      await officeDialog.notice({
        title: '无法添加链接',
        description: '请输入 http、https、mailto 或文档内 # 锚点地址。',
      });
      return;
    }
    if (!editor.isDestroyed) {
      applyDocumentPageChromeEditorCommand(editor, {
        type: 'setLink',
        href: normalized,
      });
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
      applyDocumentPageChromeEditorCommand(editor, {
        type: 'insertImage',
        alt: image.alt,
        source: image.source,
      });
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
          disabled={!state.canUndo}
          onClick={() =>
            applyDocumentPageChromeEditorCommand(editor, { type: 'undo' })
          }
        >
          <Undo2 size={16} />
        </PageChromeRibbonButton>
        <PageChromeRibbonButton
          label="重做页眉页脚编辑"
          disabled={!state.canRedo}
          onClick={() =>
            applyDocumentPageChromeEditorCommand(editor, { type: 'redo' })
          }
        >
          <Redo2 size={16} />
        </PageChromeRibbonButton>
      </WorkOfficeRibbonGroup>
      <WorkOfficeRibbonGroup label="文字">
        <PageChromeRibbonButton
          label="页眉页脚加粗"
          active={state.bold}
          onClick={() =>
            applyDocumentPageChromeEditorCommand(editor, {
              type: 'toggleBold',
            })
          }
        >
          <Bold size={16} />
        </PageChromeRibbonButton>
        <PageChromeRibbonButton
          label="页眉页脚斜体"
          active={state.italic}
          onClick={() =>
            applyDocumentPageChromeEditorCommand(editor, {
              type: 'toggleItalic',
            })
          }
        >
          <Italic size={16} />
        </PageChromeRibbonButton>
        <PageChromeRibbonButton
          label="页眉页脚下划线"
          active={state.underline}
          onClick={() =>
            applyDocumentPageChromeEditorCommand(editor, {
              type: 'toggleUnderline',
            })
          }
        >
          <Underline size={16} />
        </PageChromeRibbonButton>
        <OfficeColorPicker
          compact
          className="work-document-page-chrome-ribbon-color"
          ariaLabel="页眉页脚文字颜色"
          value={pickerColor(state.color)}
          onValueChange={(color) =>
            applyDocumentPageChromeEditorCommand(editor, {
              type: 'setColor',
              color,
            })
          }
        />
      </WorkOfficeRibbonGroup>
      <WorkOfficeRibbonGroup label="对齐">
        {(['left', 'center', 'right', 'justify'] as const).map((alignment) => (
          <PageChromeRibbonButton
            key={alignment}
            label={alignmentLabel(alignment)}
            active={state.alignment === alignment}
            onClick={() =>
              applyDocumentPageChromeEditorCommand(editor, {
                type: 'setAlignment',
                alignment,
              })
            }
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
        onChange={(event) => {
          const file = event.target.files?.[0];
          event.target.value = '';
          void insertImage(file);
        }}
      />
      {officeDialog.dialog}
    </>
  );
}

function PageChromeRibbonButton({
  label,
  active = false,
  disabled = false,
  displayLabel = false,
  onClick,
  children,
}: {
  label: string;
  active?: boolean;
  disabled?: boolean;
  displayLabel?: boolean;
  onClick: () => void;
  children: ReactNode;
}) {
  return (
    <WorkOfficeRibbonButton
      label={label}
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
