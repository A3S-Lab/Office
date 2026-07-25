import type { Editor } from '@tiptap/core';
import Subscript from '@tiptap/extension-subscript';
import Superscript from '@tiptap/extension-superscript';

export const DocumentSubscript = Subscript.extend({
  excludes: 'superscript',
  addKeyboardShortcuts() {
    return {
      'Mod-,': () => toggleDocumentSubscript(this.editor),
    };
  },
});

export const DocumentSuperscript = Superscript.extend({
  excludes: 'subscript',
  addKeyboardShortcuts() {
    return {
      'Mod-.': () => toggleDocumentSuperscript(this.editor),
    };
  },
});

export function toggleDocumentSubscript(editor: Editor): boolean {
  const chain = editor.chain().focus();
  return editor.isActive('subscript')
    ? chain.unsetSubscript().run()
    : chain.unsetSuperscript().setSubscript().run();
}

export function toggleDocumentSuperscript(editor: Editor): boolean {
  const chain = editor.chain().focus();
  return editor.isActive('superscript')
    ? chain.unsetSuperscript().run()
    : chain.unsetSubscript().setSuperscript().run();
}
