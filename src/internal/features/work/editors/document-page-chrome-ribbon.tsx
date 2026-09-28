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
import { officeMessage } from '../../../i18n/office-locale';
import type { OfficeMessageCatalog } from '../../../i18n/office-messages';
import { useOfficeMessages } from './office-messages-context';

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
  const messages = useOfficeMessages();
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
      // Keep ribbon focus on the link trigger. chain().focus() schedules into
      // the editor and breaks L2 loops.
      editor.commands.setDocumentPageChromeLink(null);
      return;
    }
    const href = await officeDialog.prompt({
      title: officeMessage(messages, 'document.pageChrome.link.title'),
      fieldLabel: officeMessage(messages, 'document.pageChrome.link.field'),
      initialValue: 'https://',
      placeholder: 'https://',
      inputMode: 'url',
      confirmLabel: officeMessage(messages, 'document.pageChrome.link.confirm'),
      required: officeMessage(messages, 'document.pageChrome.link.required'),
      validate: (value) =>
        normalizeDocumentPageChromeHref(value)
          ? null
          : DOCUMENT_LINK_VALIDATION_MESSAGE,
      restoreFocusTarget: () => editor.view.dom,
    });
    if (href === null) return;
    const normalized = normalizeDocumentPageChromeHref(href);
    if (normalized && !editor.isDestroyed) {
      // Dialog restore targets the chrome editor; focus there after apply.
      editor.chain().focus().setDocumentPageChromeLink(normalized).run();
    }
  };
  const insertImage = async (file: File | undefined) => {
    if (!file || editor.isDestroyed) return;
    const image = await loadDocumentPageChromeImage(file, messages);
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
      <WorkOfficeRibbonGroup label={officeMessage(messages, 'document.pageChrome.group.position')}>
        <PageChromeRibbonButton
          label={officeMessage(messages, 'document.pageChrome.switchHeader')}
          displayLabel
          active={editingPart === 'header'}
          onClick={() => onEditingPartChange('header')}
        >
          <PanelTop size={18} />
        </PageChromeRibbonButton>
        <PageChromeRibbonButton
          label={officeMessage(messages, 'document.pageChrome.switchFooter')}
          displayLabel
          active={editingPart === 'footer'}
          onClick={() => onEditingPartChange('footer')}
        >
          <PanelBottom size={18} />
        </PageChromeRibbonButton>
      </WorkOfficeRibbonGroup>
      <WorkOfficeRibbonGroup label={officeMessage(messages, 'document.pageChrome.group.undo')}>
        <PageChromeRibbonButton
          label={officeMessage(messages, 'document.pageChrome.undo')}
          shortcut="Cmd/Ctrl+Z"
          ariaKeyShortcuts="Control+Z Meta+Z"
          disabled={!state.canUndo}
          onClick={() => {
            // Keep ribbon focus on the undo trigger.
            // chain().focus() schedules into the editor and breaks L2 loops.
            editor.commands.undo();
          }}
        >
          <Undo2 size={16} />
        </PageChromeRibbonButton>
        <PageChromeRibbonButton
          label={officeMessage(messages, 'document.pageChrome.redo')}
          shortcut={officeMessage(messages, 'document.pageChrome.redoShortcut')}
          ariaKeyShortcuts="Control+Shift+Z Meta+Shift+Z Control+Y Meta+Y"
          disabled={!state.canRedo}
          onClick={() => {
            // Keep ribbon focus on the redo trigger.
            // chain().focus() schedules into the editor and breaks L2 loops.
            editor.commands.redo();
          }}
        >
          <Redo2 size={16} />
        </PageChromeRibbonButton>
      </WorkOfficeRibbonGroup>
      <WorkOfficeRibbonGroup
        label={officeMessage(messages, 'document.pageChrome.group.text')}
        dialogLauncher={
          onOpenFontDialog
            ? {
                label: officeMessage(messages, 'document.pageChrome.fontDialog'),
                ...pageChromeCommandShortcut('fontDialog'),
                onClick: onOpenFontDialog,
              }
            : undefined
        }
      >
        <PageChromeRibbonButton
          label={officeMessage(messages, 'document.pageChrome.bold')}
          shortcut="Cmd/Ctrl+B"
          ariaKeyShortcuts="Control+B Meta+B"
          active={state.bold}
          onClick={() => {
            // Keep ribbon focus on the bold trigger.
            // chain().focus() schedules into the editor and breaks L2 loops.
            editor.commands.toggleBold();
          }}
        >
          <Bold size={16} />
        </PageChromeRibbonButton>
        <PageChromeRibbonButton
          label={officeMessage(messages, 'document.pageChrome.italic')}
          shortcut="Cmd/Ctrl+I"
          ariaKeyShortcuts="Control+I Meta+I"
          active={state.italic}
          onClick={() => {
            // Keep ribbon focus on the italic trigger.
            // chain().focus() schedules into the editor and breaks L2 loops.
            editor.commands.toggleItalic();
          }}
        >
          <Italic size={16} />
        </PageChromeRibbonButton>
        <DocumentUnderlineRibbon
          editor={editor}
          label={officeMessage(messages, 'document.pageChrome.underline')}
          menuLabel={officeMessage(messages, 'document.pageChrome.underlineStyle')}
          colorLabel={officeMessage(messages, 'document.pageChrome.underlineColor')}
          className="work-document-page-chrome-underline"
        />
        <DocumentStrikeRibbon
          editor={editor}
          label={officeMessage(messages, 'document.pageChrome.strike')}
          menuLabel={officeMessage(messages, 'document.pageChrome.strikeStyle')}
          className="work-document-page-chrome-strike"
        />
        <PageChromeRibbonButton
          label={officeMessage(messages, 'document.pageChrome.subscript')}
          {...pageChromeCommandShortcut('subscript')}
          active={state.subscript}
          onClick={() => {
            // Keep ribbon focus on the subscript trigger.
            // chain().focus() schedules into the editor and breaks L2 loops.
            editor.commands.toggleDocumentSubscript();
          }}
        >
          <SubscriptIcon size={16} />
        </PageChromeRibbonButton>
        <PageChromeRibbonButton
          label={officeMessage(messages, 'document.pageChrome.superscript')}
          {...pageChromeCommandShortcut('superscript')}
          active={state.superscript}
          onClick={() => {
            // Keep ribbon focus on the superscript trigger.
            // chain().focus() schedules into the editor and breaks L2 loops.
            editor.commands.toggleDocumentSuperscript();
          }}
        >
          <SuperscriptIcon size={16} />
        </PageChromeRibbonButton>
        <OfficeColorPicker
          compact
          className="work-document-page-chrome-ribbon-color"
          ariaLabel={officeMessage(messages, 'document.pageChrome.textColorAria')}
          value={pickerColor(state.color)}
          onValueChange={(color) => {
            // Keep ribbon focus on the color trigger via Popover restore.
            // chain().focus() schedules into the editor and breaks L2 loops.
            editor.commands.setColor(color);
          }}
        />
        {onOpenProofingDialog && (
          <PageChromeRibbonButton
            label={officeMessage(messages, 'document.pageChrome.proofing')}
            onClick={onOpenProofingDialog}
          >
            <Languages size={16} />
          </PageChromeRibbonButton>
        )}
      </WorkOfficeRibbonGroup>
      <WorkOfficeRibbonGroup label={officeMessage(messages, 'document.pageChrome.group.align')}>
        {(['left', 'center', 'right', 'justify'] as const).map((alignment) => (
          <PageChromeRibbonButton
            key={alignment}
            label={alignmentLabel(messages, alignment)}
            {...pageChromeCommandShortcut(
              pageChromeAlignmentCommandIds[alignment],
            )}
            active={state.alignment === alignment}
            onClick={() => {
              // Keep ribbon focus on the alignment trigger.
              // chain().focus() schedules into the editor and breaks L2 loops.
              editor.commands.setTextAlign(alignment);
            }}
          >
            {alignmentIcon(alignment)}
          </PageChromeRibbonButton>
        ))}
      </WorkOfficeRibbonGroup>
      <WorkOfficeRibbonGroup label={officeMessage(messages, 'document.pageChrome.group.insert')}>
        <PageChromeRibbonButton
          label={officeMessage(messages, state.link ? 'document.pageChrome.link.remove' : 'document.pageChrome.link.add')}
          active={Boolean(state.link)}
          onClick={() => void editLink()}
        >
          <Link2 size={17} />
        </PageChromeRibbonButton>
        <PageChromeRibbonButton
          label={officeMessage(messages, 'document.pageChrome.insertImage')}
          onClick={() => imageInputRef.current?.click()}
        >
          <ImageIcon size={17} />
        </PageChromeRibbonButton>
      </WorkOfficeRibbonGroup>
      <WorkOfficeRibbonGroup label={officeMessage(messages, 'document.pageChrome.group.pageNumber')}>
        <PageChromeRibbonButton
          label={officeMessage(messages, 'document.pageChrome.showPageNumber')}
          displayLabel
          active={showPageNumber}
          onClick={onTogglePageNumber}
        >
          <Hash size={18} />
        </PageChromeRibbonButton>
      </WorkOfficeRibbonGroup>
      <WorkOfficeRibbonGroup label={officeMessage(messages, 'document.pageChrome.group.close')}>
        <PageChromeRibbonButton
          label={officeMessage(messages, 'document.pageChrome.close')}
          displayLabel
          onClick={onClose}
        >
          <X size={18} />
        </PageChromeRibbonButton>
      </WorkOfficeRibbonGroup>
      <OfficeFileInput
        ref={imageInputRef}
        accept="image/bmp,image/gif,image/jpeg,image/png,image/webp"
        aria-label={officeMessage(messages, 'document.pageChrome.imageFileAria')}
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
  const messages = useOfficeMessages();
  return (
    <WorkOfficeRibbonButton
      label={label}
      title={
        shortcut
          ? officeMessage(messages, 'document.pageChrome.buttonTitleWithShortcut', {
              label,
              shortcut,
            })
          : label
      }
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
  messages: OfficeMessageCatalog,
  alignment: 'center' | 'justify' | 'left' | 'right',
): string {
  switch (alignment) {
    case 'center':
      return officeMessage(messages, 'document.pageChrome.align.center');
    case 'justify':
      return officeMessage(messages, 'document.pageChrome.align.justify');
    case 'left':
      return officeMessage(messages, 'document.pageChrome.align.left');
    case 'right':
      return officeMessage(messages, 'document.pageChrome.align.right');
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
