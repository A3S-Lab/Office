import type { Editor } from '@tiptap/core';
import {
  copyDocumentFormatting,
  pasteDocumentFormatting,
} from './document-format-clipboard';
import { changeDocumentFontSize } from './document-formatting-options';
import { moveDocumentBlock } from './document-move-block';
import { cycleDocumentSelectionCase } from './document-selection-case';

export interface DocumentWpsShortcutCallbacks {
  canInsertComment: boolean;
  canRefreshFields: boolean;
  onInsertComment: () => void;
  onOpenFontDialog: () => void;
  onOpenWordCount: () => void;
  onRefreshFields: () => void;
  onToggleFieldCodes: () => void;
  /** WPS/Word Shift+F9: toggle codes for the selected field(s) only. */
  onToggleSelectedFieldCodes: () => boolean;
  /** WPS/Word Ctrl+Shift+F9: replace selected fields with result text. */
  onUnlinkFields: () => boolean;
  /** WPS/Word Ctrl+F11: lock selected fields against F9 refresh. */
  onLockFields: () => boolean;
  /** WPS/Word Ctrl+Shift+F11: unlock selected fields. */
  onUnlockFields: () => boolean;
  onToggleSpellcheck: () => void;
  onToggleTrackChanges: () => void;
}

export type DocumentWpsShortcutEvent = Pick<
  KeyboardEvent,
  'altKey' | 'code' | 'ctrlKey' | 'key' | 'metaKey' | 'shiftKey'
>;

export function runDocumentWpsShortcut(
  editor: Editor,
  event: DocumentWpsShortcutEvent,
  callbacks: DocumentWpsShortcutCallbacks,
): boolean {
  if (editor.isDestroyed) return false;
  const key = event.key.toLowerCase();
  const modifier = event.ctrlKey || event.metaKey;

  if (!modifier) {
    // WPS/Word Shift+F3: cycle selected text lower → UPPER → Title Case.
    if (key === 'f3' && event.shiftKey && !event.altKey) {
      return cycleDocumentSelectionCase(editor);
    }
    if (key === 'f9' && event.altKey && !event.shiftKey) {
      callbacks.onToggleFieldCodes();
      return true;
    }
    if (key === 'f9' && event.shiftKey && !event.altKey) {
      return callbacks.onToggleSelectedFieldCodes();
    }
    // WPS Shift+Alt+. / , and Alt+Shift+→ / ← : increase / decrease indent.
    // WPS/Word Alt+Shift+↑ / ↓ : move the current block among siblings.
    if (event.altKey && event.shiftKey) {
      if (event.code === 'ArrowUp' || key === 'arrowup') {
        return moveDocumentBlock(editor, -1);
      }
      if (event.code === 'ArrowDown' || key === 'arrowdown') {
        return moveDocumentBlock(editor, 1);
      }
      if (
        event.code === 'Period' ||
        key === '.' ||
        key === '>' ||
        event.code === 'ArrowRight' ||
        key === 'arrowright'
      ) {
        editor.commands.changeDocumentIndent(1);
        return true;
      }
      if (
        event.code === 'Comma' ||
        key === ',' ||
        key === '<' ||
        event.code === 'ArrowLeft' ||
        key === 'arrowleft'
      ) {
        editor.commands.changeDocumentIndent(-1);
        return true;
      }
      return false;
    }
    if (event.altKey || event.shiftKey) return false;
    if (key === 'f7') {
      callbacks.onToggleSpellcheck();
      return true;
    }
    if (key === 'f9') {
      if (callbacks.canRefreshFields) callbacks.onRefreshFields();
      return true;
    }
    return false;
  }

  if (event.altKey) {
    if (event.shiftKey) return false;
    if (key === 'm') {
      if (callbacks.canInsertComment) callbacks.onInsertComment();
      return true;
    }
    const headingLevel = headingLevelForShortcut(event.code, key);
    if (headingLevel !== null) {
      editor.chain().focus().setHeading({ level: headingLevel }).run();
      return true;
    }
    return false;
  }

  if (event.shiftKey) {
    if (key === 'f9') {
      return callbacks.onUnlinkFields();
    }
    if (key === 'f11') {
      return callbacks.onUnlockFields();
    }
    if (key === 'h') {
      editor.commands.toggleDocumentHiddenText();
      return true;
    }
    if (key === 'd') {
      editor.commands.toggleDocumentUnderlineStyle('double');
      return true;
    }
    if (key === 'w') {
      editor.commands.toggleDocumentUnderlineStyle('words');
      return true;
    }
    // WPS: Ctrl+Shift+X toggles double strikethrough.
    if (key === 'x') {
      editor.commands.toggleDocumentStrikeStyle('double');
      return true;
    }
    if (key === 'a') {
      editor.commands.toggleDocumentTextCase('all-caps');
      return true;
    }
    if (key === 'k') {
      editor.commands.toggleDocumentTextCase('small-caps');
      return true;
    }
    if (key === 'g') {
      callbacks.onOpenWordCount();
      return true;
    }
    if (key === 'c') {
      copyDocumentFormatting(editor);
      return true;
    }
    if (key === 'v') {
      pasteDocumentFormatting(editor);
      return true;
    }
    if (key === 'e') {
      callbacks.onToggleTrackChanges();
      return true;
    }
    // WPS/Word Ctrl+Shift+M: decrease paragraph indent.
    if (key === 'm') {
      editor.commands.changeDocumentIndent(-1);
      return true;
    }
    // WPS/Word Ctrl+Shift+N: apply the Normal / 正文 paragraph style.
    if (key === 'n') {
      editor.chain().focus().setParagraph().run();
      return true;
    }
    // WPS Ctrl+Shift+J: distributed paragraph alignment.
    if (key === 'j') {
      editor.chain().focus().setTextAlign('distribute').run();
      return true;
    }
    // WPS/Word Ctrl+Shift+L: toggle the default bullet list.
    if (key === 'l') {
      if (editor.isActive('bulletList')) {
        editor.chain().focus().clearDocumentList().run();
      } else {
        editor.chain().focus().applyDocumentBulletList('disc').run();
      }
      return true;
    }
    if (event.code === 'Period' || key === '>' || key === '.') {
      changeDocumentFontSize(editor, 1);
      return true;
    }
    if (event.code === 'Comma' || key === '<' || key === ',') {
      changeDocumentFontSize(editor, -1);
      return true;
    }
    return false;
  }

  if (key === 'f11') {
    return callbacks.onLockFields();
  }

  // WPS/Word Ctrl+Space: clear character and paragraph formatting on the selection.
  if (key === ' ' || event.code === 'Space') {
    editor.commands.clearDocumentFormatting();
    return true;
  }

  // WPS/Word Ctrl+M: increase paragraph indent (Ctrl+Alt+M inserts a comment).
  if (key === 'm') {
    editor.commands.changeDocumentIndent(1);
    return true;
  }

  if (key === ']') {
    changeDocumentFontSize(editor, 1);
    return true;
  }
  if (key === '[') {
    changeDocumentFontSize(editor, -1);
    return true;
  }
  if (key === 'd') {
    callbacks.onOpenFontDialog();
    return true;
  }

  const alignment = alignmentForShortcut(key);
  if (alignment) {
    editor.chain().focus().setTextAlign(alignment).run();
    return true;
  }

  const lineHeight = lineHeightForShortcut(key);
  if (lineHeight) {
    editor.commands.setDocumentLineHeight(lineHeight);
    return true;
  }

  return false;
}

function alignmentForShortcut(
  key: string,
): 'left' | 'center' | 'right' | 'justify' | null {
  if (key === 'l') return 'left';
  if (key === 'e') return 'center';
  if (key === 'r') return 'right';
  if (key === 'j') return 'justify';
  return null;
}

function lineHeightForShortcut(key: string): '1' | '1.5' | '2' | null {
  if (key === '1') return '1';
  if (key === '5') return '1.5';
  if (key === '2') return '2';
  return null;
}

function headingLevelForShortcut(code: string, key: string): 1 | 2 | 3 | null {
  const value = code.startsWith('Digit') ? code.slice(5) : key;
  if (value === '1') return 1;
  if (value === '2') return 2;
  if (value === '3') return 3;
  return null;
}
