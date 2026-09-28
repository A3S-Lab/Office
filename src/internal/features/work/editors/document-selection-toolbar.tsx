import type { Editor } from '@tiptap/core';
import { TextSelection } from '@tiptap/pm/state';
import { BubbleMenu } from '@tiptap/react/menus';
import {
  Bold,
  Eraser,
  Highlighter,
  Italic,
  MessageSquarePlus,
} from 'lucide-react';
import { type ReactNode, useEffect, useRef } from 'react';
import { officeFloatingPortalRoot } from '../../../design-system/primitives/overlay/portal-root';
import { officeMessage } from '../../../i18n/office-locale';
import type { WorkDocumentLayoutFont } from '../work-document-fonts';
import { documentCommandLabel } from './document-command-i18n';
import {
  documentFontFamilyOptionsForValue,
  documentFontFamilyValue,
  documentFontSizeOptionsForValue,
  documentFontSizeValue,
} from './document-formatting-options';
import { DocumentStrikeRibbon } from './document-strike-ribbon';
import { DocumentUnderlineRibbon } from './document-underline-ribbon';
import { OfficeColorPicker, OfficeSelect } from './office-controls';
import { useOfficeMessages } from './office-messages-context';
import { moveOfficeToolbarFocus } from './office-toolbar-keyboard';

const bubbleMenuOptions = {
  strategy: 'fixed' as const,
  placement: 'top' as const,
  offset: 9,
  flip: { padding: 8 },
  shift: { padding: 8 },
  inline: true,
};

const selectionToolbarPluginKey = 'documentSelectionToolbar';

export function DocumentSelectionToolbar({
  editor,
  canInsertComment,
  layoutFonts = [],
  onInsertComment,
  reviewOnly = false,
}: {
  editor: Editor;
  canInsertComment: boolean;
  layoutFonts?: readonly WorkDocumentLayoutFont[];
  onInsertComment: () => void;
  reviewOnly?: boolean;
}) {
  const messages = useOfficeMessages();
  const toolbarRef = useRef<HTMLDivElement>(null);
  const fontFamilyValue = documentFontFamilyValue(
    editor,
    layoutFonts,
    messages,
  );
  const fontSizeValue = documentFontSizeValue(editor);

  useEffect(() => {
    const hideOutsideInteraction = (event: FocusEvent | PointerEvent) => {
      const toolbar = toolbarRef.current;
      if (
        editor.isDestroyed ||
        !(event.target instanceof Node) ||
        editor.view.dom.contains(event.target) ||
        (toolbar && selectionToolbarOwnsTarget(toolbar, event.target))
      ) {
        return;
      }
      editor.view.dispatch(
        editor.state.tr.setMeta(selectionToolbarPluginKey, 'hide'),
      );
    };

    document.addEventListener('focusin', hideOutsideInteraction, true);
    document.addEventListener('pointerdown', hideOutsideInteraction, true);
    return () => {
      document.removeEventListener('focusin', hideOutsideInteraction, true);
      document.removeEventListener('pointerdown', hideOutsideInteraction, true);
    };
  }, [editor]);

  return (
    <BubbleMenu
      ref={toolbarRef}
      editor={editor}
      pluginKey={selectionToolbarPluginKey}
      className="work-document-selection-toolbar"
      role="toolbar"
      aria-label={officeMessage(messages, 'document.selectionToolbar.aria')}
      onKeyDown={moveOfficeToolbarFocus}
      updateDelay={80}
      resizeDelay={60}
      appendTo={() =>
        officeFloatingPortalRoot(editor.view.dom.ownerDocument, editor.view.dom)
      }
      options={bubbleMenuOptions}
      shouldShow={({ element, state, view }) => {
        const { selection } = state;
        const hasFocus =
          view.hasFocus() || element.contains(document.activeElement);
        return (
          hasFocus &&
          editor.isEditable &&
          selection instanceof TextSelection &&
          !selection.empty &&
          Boolean(
            state.doc.textBetween(selection.from, selection.to, '\n').trim(),
          )
        );
      }}
      onPointerDown={(event) => {
        if (event.nativeEvent.pointerType !== 'touch') event.preventDefault();
      }}
    >
      {!reviewOnly && (
        <>
          <OfficeSelect
            ariaLabel={officeMessage(
              messages,
              'document.selectionToolbar.fontFamily',
            )}
            className="work-document-selection-font-family"
            value={fontFamilyValue}
            options={documentFontFamilyOptionsForValue(
              fontFamilyValue,
              layoutFonts,
              messages,
            )}
            onValueChange={(value) => {
              // Keep toolbar focus on the font combobox via Popover restore.
              // chain().focus() schedules into the editor and breaks L2 loops.
              if (value === 'default') editor.commands.unsetFontFamily();
              else editor.commands.setFontFamily(value);
            }}
          />
          <OfficeSelect
            ariaLabel={officeMessage(
              messages,
              'document.selectionToolbar.fontSize',
            )}
            className="work-document-selection-font-size"
            value={fontSizeValue}
            options={documentFontSizeOptionsForValue(fontSizeValue)}
            onValueChange={(value) => {
              // Keep toolbar focus on the size combobox via Popover restore.
              // chain().focus() schedules into the editor and breaks L2 loops.
              if (value === 'default') editor.commands.unsetFontSize();
              else editor.commands.setFontSize(value);
            }}
          />
          <span
            className="work-document-selection-divider"
            aria-hidden="true"
          />
          <SelectionToolbarButton
            label={documentCommandLabel('bold', messages)}
            active={editor.isActive('bold')}
            onClick={() => {
              // Keep toolbar focus on the bold trigger.
              // chain().focus() schedules into the editor and breaks L2 loops.
              editor.commands.toggleBold();
            }}
          >
            <Bold size={15} />
          </SelectionToolbarButton>
          <SelectionToolbarButton
            label={documentCommandLabel('italic', messages)}
            active={editor.isActive('italic')}
            onClick={() => {
              // Keep toolbar focus on the italic trigger.
              // chain().focus() schedules into the editor and breaks L2 loops.
              editor.commands.toggleItalic();
            }}
          >
            <Italic size={15} />
          </SelectionToolbarButton>
          <DocumentUnderlineRibbon
            editor={editor}
            menuLabel={officeMessage(
              messages,
              'document.selectionToolbar.underline',
            )}
            showColor={false}
            className="work-document-selection-underline"
          />
          <DocumentStrikeRibbon
            editor={editor}
            menuLabel={officeMessage(
              messages,
              'document.selectionToolbar.strike',
            )}
            className="work-document-selection-strike"
          />
          <OfficeColorPicker
            compact
            className="work-document-selection-color"
            value={editor.getAttributes('textStyle').color ?? '#172033'}
            ariaLabel={officeMessage(
              messages,
              'document.selectionToolbar.color',
            )}
            onValueChange={(color) => {
              // Keep toolbar focus on the color trigger via Popover restore.
              // chain().focus() schedules into the editor and breaks L2 loops.
              editor.commands.setColor(color);
            }}
          />
          <SelectionToolbarButton
            label={documentCommandLabel('highlight', messages)}
            active={editor.isActive('highlight')}
            onClick={() => {
              // Keep toolbar focus on the highlight trigger.
              // chain().focus() schedules into the editor and breaks L2 loops.
              editor.commands.toggleHighlight({ color: '#fff0a6' });
            }}
          >
            <Highlighter size={15} />
          </SelectionToolbarButton>
          <SelectionToolbarButton
            className="secondary"
            label={documentCommandLabel('clearFormatting', messages)}
            onClick={() => editor.commands.clearDocumentFormatting()}
          >
            <Eraser size={15} />
          </SelectionToolbarButton>
          <span
            className="work-document-selection-divider"
            aria-hidden="true"
          />
        </>
      )}
      <SelectionToolbarButton
        label={documentCommandLabel('insertComment', messages)}
        disabled={!canInsertComment}
        onClick={onInsertComment}
      >
        <MessageSquarePlus size={15} />
      </SelectionToolbarButton>
    </BubbleMenu>
  );
}

function selectionToolbarOwnsTarget(
  toolbar: HTMLElement,
  target: Node,
): boolean {
  if (toolbar.contains(target)) return true;
  return [...toolbar.querySelectorAll<HTMLElement>('[aria-controls]')]
    .map((controller) => controller.getAttribute('aria-controls'))
    .some((controlledId) => {
      if (!controlledId) return false;
      return document.getElementById(controlledId)?.contains(target) ?? false;
    });
}

function SelectionToolbarButton({
  active,
  className = '',
  disabled = false,
  label,
  onClick,
  children,
}: {
  active?: boolean;
  className?: string;
  disabled?: boolean;
  label: string;
  onClick: () => void;
  children: ReactNode;
}) {
  return (
    <button
      type="button"
      className={className}
      aria-label={label}
      aria-pressed={active}
      title={label}
      disabled={disabled}
      onClick={onClick}
    >
      {children}
    </button>
  );
}
