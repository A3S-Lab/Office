import type { Editor } from '@tiptap/core';
import {
  AArrowDown,
  AArrowUp,
  AlignCenter,
  AlignHorizontalDistributeCenter,
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
  Square,
  Subscript as SubscriptIcon,
  Superscript as SuperscriptIcon,
} from 'lucide-react';
import {
  type MouseEvent as ReactMouseEvent,
  type ReactNode,
  useCallback,
  useSyncExternalStore,
} from 'react';
import { officeMessage } from '../../../i18n/office-locale';
import type { OfficeMessageCatalog } from '../../../i18n/office-messages';
import type { WorkDocumentLayoutFont } from '../work-document-fonts';
import {
  canChangeDocumentIndent,
  documentParagraphDirection,
} from '../work-document-paragraph-formatting';
import {
  documentRunBorderIsVisible,
  parseDocumentRunBorder,
} from '../work-document-run-border';
import {
  type DocumentCommandId,
  getDocumentCommandDefinition,
} from './document-command-catalog';
import { documentCommandLabel } from './document-command-i18n';
import type { DocumentFindReplaceMode } from './document-find-replace-panel';
import { DocumentFormatTools } from './document-format-tools';
import {
  canChangeDocumentFontSize,
  changeDocumentFontSize,
  documentFontFamilyOptionsForValue,
  documentFontFamilyValue,
  documentFontSizeOptionsForValue,
  documentFontSizeValue,
  documentLineHeightValue,
} from './document-formatting-options';
import { DocumentListGallery } from './document-list-gallery';
import { DocumentStrikeRibbon } from './document-strike-ribbon';
import { DocumentStyleGallery } from './document-style-gallery';
import { DocumentTextCaseRibbon } from './document-text-case-ribbon';
import { DocumentUnderlineRibbon } from './document-underline-ribbon';
import { OfficeColorPicker, OfficeSelect } from './office-controls';
import { useOfficeMessages } from './office-messages-context';
import {
  WorkOfficeRibbonButton,
  WorkOfficeRibbonGroup,
} from './work-office-chrome';

const lineSpacingAriaKeyShortcuts = [
  getDocumentCommandDefinition('lineSpacingSingle').shortcut?.aria,
  getDocumentCommandDefinition('lineSpacingOneAndHalf').shortcut?.aria,
  getDocumentCommandDefinition('lineSpacingDouble').shortcut?.aria,
]
  .filter((value): value is string => Boolean(value))
  .join(' ');

function documentLineHeightOptions(messages: OfficeMessageCatalog) {
  return [
    { value: 'default', label: messages['document.lineHeight.default'] },
    {
      value: '1',
      label: messages['document.lineHeight.single'],
      meta: getDocumentCommandDefinition('lineSpacingSingle').shortcut?.label,
    },
    { value: '1.15', label: messages['document.lineHeight.1_15'] },
    {
      value: '1.5',
      label: messages['document.lineHeight.1_5'],
      meta: getDocumentCommandDefinition('lineSpacingOneAndHalf').shortcut
        ?.label,
    },
    {
      value: '2',
      label: messages['document.lineHeight.double'],
      meta: getDocumentCommandDefinition('lineSpacingDouble').shortcut?.label,
    },
  ] as const;
}

function documentLineHeightOptionsForValue(
  messages: OfficeMessageCatalog,
  value: string,
) {
  const options = documentLineHeightOptions(messages);
  if (options.some((option) => option.value === value)) {
    return options;
  }
  const numeric = Number(value);
  return [
    ...options,
    {
      value,
      label:
        Number.isFinite(numeric) && numeric > 0
          ? officeMessage(messages, 'document.lineHeight.custom', {
              value: String(numeric),
            })
          : value,
    },
  ];
}

export function DocumentHomeRibbon({
  editor,
  findReplaceMode,
  layoutFonts = [],
  onFindText,
  onOpenFontDialog,
}: {
  editor: Editor;
  findReplaceMode: DocumentFindReplaceMode | null;
  layoutFonts?: readonly WorkDocumentLayoutFont[];
  onFindText: (replace: boolean) => void;
  onOpenFontDialog?: () => void;
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
  const messages = useOfficeMessages();
  if (editor.isDestroyed) return null;
  const fontFamilyValue = documentFontFamilyValue(
    editor,
    layoutFonts,
    messages,
  );
  const fontSizeValue = documentFontSizeValue(editor);
  const lineHeightValue = documentLineHeightValue(editor);

  return (
    <>
      <DocumentFormatTools editor={editor} />
      <RibbonGroup
        label={messages['document.group.font']}
        priority="high"
        dialogLauncher={
          onOpenFontDialog
            ? {
                label: documentCommandLabel('fontDialog', messages),
                ...commandShortcut('fontDialog'),
                onClick: onOpenFontDialog,
              }
            : undefined
        }
      >
        <div className="work-document-font-tools">
          <div className="work-document-font-selects">
            <OfficeSelect
              ariaLabel={officeMessage(messages, 'document.home.fontFamilyAria')}
              className="work-document-font-family-select"
              value={fontFamilyValue}
              options={documentFontFamilyOptionsForValue(
                fontFamilyValue,
                layoutFonts,
                messages,
              )}
              onValueChange={(value) => {
                // Keep ribbon focus on the font combobox via Popover restore.
                // chain().focus() schedules into the editor and breaks L2 loops.
                if (value === 'default') editor.commands.unsetFontFamily();
                else editor.commands.setFontFamily(value);
              }}
            />
            <OfficeSelect
              ariaLabel={officeMessage(messages, 'document.home.fontSizeAria')}
              className="work-document-font-size-select"
              value={fontSizeValue}
              options={documentFontSizeOptionsForValue(fontSizeValue)}
              onValueChange={(value) => {
                // Keep ribbon focus on the size combobox via Popover restore.
                // chain().focus() schedules into the editor and breaks L2 loops.
                if (value === 'default') editor.commands.unsetFontSize();
                else editor.commands.setFontSize(value);
              }}
            />
            <ToolbarButton
              label={documentCommandLabel('growFont', messages)}
              {...commandShortcut('growFont')}
              disabled={!canChangeDocumentFontSize(editor, 1)}
              onMouseDown={(event) => event.preventDefault()}
              onClick={() => changeDocumentFontSize(editor, 1)}
            >
              <AArrowUp size={16} />
            </ToolbarButton>
            <ToolbarButton
              label={documentCommandLabel('shrinkFont', messages)}
              {...commandShortcut('shrinkFont')}
              disabled={!canChangeDocumentFontSize(editor, -1)}
              onMouseDown={(event) => event.preventDefault()}
              onClick={() => changeDocumentFontSize(editor, -1)}
            >
              <AArrowDown size={16} />
            </ToolbarButton>
          </div>
          <div className="work-document-font-actions">
            <ToolbarButton
              label={documentCommandLabel('bold', messages)}
              {...commandShortcut('bold')}
              active={editor.isActive('bold')}
              onMouseDown={(event) => event.preventDefault()}
              onClick={() => {
                // Keep ribbon focus on the bold trigger.
                // chain().focus() schedules into the editor and breaks L2 loops.
                editor.commands.toggleBold();
              }}
            >
              <Bold size={16} />
            </ToolbarButton>
            <ToolbarButton
              label={documentCommandLabel('italic', messages)}
              {...commandShortcut('italic')}
              active={editor.isActive('italic')}
              onMouseDown={(event) => event.preventDefault()}
              onClick={() => {
                // Keep ribbon focus on the italic trigger.
                // chain().focus() schedules into the editor and breaks L2 loops.
                editor.commands.toggleItalic();
              }}
            >
              <Italic size={16} />
            </ToolbarButton>
            <DocumentUnderlineRibbon editor={editor} />
            <DocumentStrikeRibbon editor={editor} />
            <DocumentTextCaseRibbon editor={editor} />
            <ToolbarButton
              label={documentCommandLabel('subscript', messages)}
              {...commandShortcut('subscript')}
              active={editor.isActive('subscript')}
              onClick={() => editor.commands.toggleDocumentSubscript()}
            >
              <SubscriptIcon size={16} />
            </ToolbarButton>
            <ToolbarButton
              label={documentCommandLabel('superscript', messages)}
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
              ariaLabel={officeMessage(messages, 'document.home.textColorAria')}
              onValueChange={(color) => {
                // Keep ribbon focus on the color trigger via Popover restore.
                // chain().focus() schedules into the editor and breaks L2 loops.
                editor.commands.setColor(color);
              }}
            />
            <ToolbarButton
              label={documentCommandLabel('highlight', messages)}
              active={editor.isActive('highlight')}
              onClick={() => {
                // Keep ribbon focus on the highlight trigger.
                // chain().focus() schedules into the editor and breaks L2 loops.
                editor.commands.toggleHighlight({ color: '#fff0a6' });
              }}
            >
              <Highlighter size={16} />
            </ToolbarButton>
            <ToolbarButton
              label={documentCommandLabel('runBorder', messages)}
              active={documentRunBorderIsVisible(
                parseDocumentRunBorder(
                  editor.getAttributes('textStyle').runBorder,
                ),
              )}
              onClick={() => editor.commands.toggleDocumentRunBorder()}
            >
              <Square size={16} />
            </ToolbarButton>
            <ToolbarButton
              label={documentCommandLabel('clearFormatting', messages)}
              {...commandShortcut('clearFormatting')}
              onClick={() => editor.commands.clearDocumentFormatting()}
            >
              <Eraser size={16} />
            </ToolbarButton>
          </div>
        </div>
      </RibbonGroup>
      <RibbonGroup label={messages['document.group.paragraph']} priority="high">
        <div className="work-document-paragraph-tools">
          <div className="work-document-paragraph-actions">
            <DocumentListGallery editor={editor} />
            <ToolbarButton
              label={documentCommandLabel('decreaseIndent', messages)}
              {...commandShortcut('decreaseIndent')}
              disabled={!canChangeDocumentIndent(editor, -1)}
              onClick={() => editor.commands.changeDocumentIndent(-1)}
            >
              <IndentDecrease size={16} />
            </ToolbarButton>
            <ToolbarButton
              label={documentCommandLabel('increaseIndent', messages)}
              {...commandShortcut('increaseIndent')}
              disabled={!canChangeDocumentIndent(editor, 1)}
              onClick={() => editor.commands.changeDocumentIndent(1)}
            >
              <IndentIncrease size={16} />
            </ToolbarButton>
            <OfficeSelect
              ariaLabel={officeMessage(
                messages,
                'document.home.lineSpacingAria',
              )}
              ariaKeyShortcuts={lineSpacingAriaKeyShortcuts}
              className="work-document-line-height-select"
              value={lineHeightValue}
              options={documentLineHeightOptionsForValue(
                messages,
                lineHeightValue,
              )}
              onValueChange={(value) =>
                editor.commands.setDocumentLineHeight(
                  value === 'default' ? null : value,
                )
              }
            />
          </div>
          <div className="work-document-alignment-actions">
            <ToolbarButton
              label={documentCommandLabel('alignLeft', messages)}
              {...commandShortcut('alignLeft')}
              active={editor.isActive({ textAlign: 'left' })}
              onMouseDown={(event) => event.preventDefault()}
              onClick={() => {
                // Keep ribbon focus on the alignment trigger.
                // chain().focus() schedules into the editor and breaks L2 loops.
                editor.commands.setTextAlign('left');
              }}
            >
              <AlignLeft size={16} />
            </ToolbarButton>
            <ToolbarButton
              label={documentCommandLabel('alignCenter', messages)}
              {...commandShortcut('alignCenter')}
              active={editor.isActive({ textAlign: 'center' })}
              onMouseDown={(event) => event.preventDefault()}
              onClick={() => {
                // Keep ribbon focus on the alignment trigger.
                // chain().focus() schedules into the editor and breaks L2 loops.
                editor.commands.setTextAlign('center');
              }}
            >
              <AlignCenter size={16} />
            </ToolbarButton>
            <ToolbarButton
              label={documentCommandLabel('alignRight', messages)}
              {...commandShortcut('alignRight')}
              active={editor.isActive({ textAlign: 'right' })}
              onMouseDown={(event) => event.preventDefault()}
              onClick={() => {
                // Keep ribbon focus on the alignment trigger.
                // chain().focus() schedules into the editor and breaks L2 loops.
                editor.commands.setTextAlign('right');
              }}
            >
              <AlignRight size={16} />
            </ToolbarButton>
            <ToolbarButton
              label={documentCommandLabel('alignJustify', messages)}
              {...commandShortcut('alignJustify')}
              active={editor.isActive({ textAlign: 'justify' })}
              onMouseDown={(event) => event.preventDefault()}
              onClick={() => {
                // Keep ribbon focus on the alignment trigger.
                // chain().focus() schedules into the editor and breaks L2 loops.
                editor.commands.setTextAlign('justify');
              }}
            >
              <AlignJustify size={16} />
            </ToolbarButton>
            <ToolbarButton
              label={documentCommandLabel('alignDistribute', messages)}
              {...commandShortcut('alignDistribute')}
              active={editor.isActive({ textAlign: 'distribute' })}
              onMouseDown={(event) => event.preventDefault()}
              onClick={() => {
                // Keep ribbon focus on the alignment trigger.
                // chain().focus() schedules into the editor and breaks L2 loops.
                editor.commands.setTextAlign('distribute');
              }}
            >
              <AlignHorizontalDistributeCenter size={16} />
            </ToolbarButton>
            <ToolbarButton
              label={officeMessage(messages, 'document.home.ltr')}
              active={documentParagraphDirection(editor) === 'ltr'}
              onClick={() =>
                editor.commands.setDocumentParagraphDirection('ltr')
              }
            >
              <PilcrowRight size={16} />
            </ToolbarButton>
            <ToolbarButton
              label={officeMessage(messages, 'document.home.rtl')}
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
      <RibbonGroup label={messages['document.group.style']} priority="low">
        <DocumentStyleGallery editor={editor} />
      </RibbonGroup>
      <RibbonGroup label={messages['document.group.edit']} priority="low">
        <ToolbarButton
          label={documentCommandLabel('find', messages)}
          {...commandShortcut('find')}
          active={findReplaceMode === 'find'}
          onClick={() => onFindText(false)}
        >
          <Search size={16} />
        </ToolbarButton>
        <ToolbarButton
          label={documentCommandLabel('replace', messages)}
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
  onMouseDown,
  children,
}: {
  label: string;
  shortcut?: string;
  ariaKeyShortcuts?: string;
  active?: boolean;
  disabled?: boolean;
  onClick: () => void;
  onMouseDown?: (event: ReactMouseEvent<HTMLButtonElement>) => void;
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
      onMouseDown={onMouseDown}
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
