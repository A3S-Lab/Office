import type { Editor } from '@tiptap/core';
import { TextSelection } from '@tiptap/pm/state';
import { BubbleMenu } from '@tiptap/react/menus';
import {
  Bold,
  Eraser,
  Highlighter,
  Italic,
  MessageSquarePlus,
  Strikethrough,
  Underline as UnderlineIcon,
} from 'lucide-react';
import { type ReactNode, useEffect, useRef } from 'react';
import {
  documentFontFamilyOptions,
  documentFontFamilyValue,
  documentFontSizeOptions,
  documentFontSizeValue,
} from './document-formatting-options';
import { OfficeColorPicker, OfficeSelect } from './office-controls';

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
  onInsertComment,
}: {
  editor: Editor;
  canInsertComment: boolean;
  onInsertComment: () => void;
}) {
  const toolbarRef = useRef<HTMLDivElement>(null);

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
      aria-label="文本快捷工具栏"
      updateDelay={80}
      resizeDelay={60}
      appendTo={() => document.body}
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
      <OfficeSelect
        ariaLabel="快捷字体"
        className="work-document-selection-font-family"
        value={documentFontFamilyValue(editor)}
        options={documentFontFamilyOptions}
        onValueChange={(value) => {
          if (value === 'default')
            editor.chain().focus().unsetFontFamily().run();
          else editor.chain().focus().setFontFamily(value).run();
        }}
      />
      <OfficeSelect
        ariaLabel="快捷字号"
        className="work-document-selection-font-size"
        value={documentFontSizeValue(editor)}
        options={documentFontSizeOptions}
        onValueChange={(value) => {
          if (value === 'default') editor.chain().focus().unsetFontSize().run();
          else editor.chain().focus().setFontSize(value).run();
        }}
      />
      <span className="work-document-selection-divider" aria-hidden="true" />
      <SelectionToolbarButton
        label="加粗"
        active={editor.isActive('bold')}
        onClick={() => editor.chain().focus().toggleBold().run()}
      >
        <Bold size={15} />
      </SelectionToolbarButton>
      <SelectionToolbarButton
        label="斜体"
        active={editor.isActive('italic')}
        onClick={() => editor.chain().focus().toggleItalic().run()}
      >
        <Italic size={15} />
      </SelectionToolbarButton>
      <SelectionToolbarButton
        label="下划线"
        active={editor.isActive('underline')}
        onClick={() => editor.chain().focus().toggleUnderline().run()}
      >
        <UnderlineIcon size={15} />
      </SelectionToolbarButton>
      <SelectionToolbarButton
        label="删除线"
        active={editor.isActive('strike')}
        onClick={() => editor.chain().focus().toggleStrike().run()}
      >
        <Strikethrough size={15} />
      </SelectionToolbarButton>
      <OfficeColorPicker
        compact
        className="work-document-selection-color"
        value={editor.getAttributes('textStyle').color ?? '#172033'}
        ariaLabel="快捷文字颜色"
        onValueChange={(color) => editor.chain().focus().setColor(color).run()}
      />
      <SelectionToolbarButton
        label="突出显示"
        active={editor.isActive('highlight')}
        onClick={() =>
          editor.chain().focus().toggleHighlight({ color: '#fff0a6' }).run()
        }
      >
        <Highlighter size={15} />
      </SelectionToolbarButton>
      <SelectionToolbarButton
        className="secondary"
        label="清除格式"
        onClick={() => editor.commands.clearDocumentFormatting()}
      >
        <Eraser size={15} />
      </SelectionToolbarButton>
      <span className="work-document-selection-divider" aria-hidden="true" />
      <SelectionToolbarButton
        label="添加批注"
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
