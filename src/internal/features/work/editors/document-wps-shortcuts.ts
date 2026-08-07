import type { Editor } from '@tiptap/core';
import { changeDocumentFontSize } from './document-formatting-options';

export interface DocumentWpsShortcutCallbacks {
  canInsertComment: boolean;
  canRefreshFields: boolean;
  onInsertComment: () => void;
  onRefreshFields: () => void;
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
    if (key === 'e') {
      callbacks.onToggleTrackChanges();
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

  if (key === ']') {
    changeDocumentFontSize(editor, 1);
    return true;
  }
  if (key === '[') {
    changeDocumentFontSize(editor, -1);
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
