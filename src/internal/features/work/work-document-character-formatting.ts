import type { Editor } from '@tiptap/core';
import Subscript from '@tiptap/extension-subscript';
import Superscript from '@tiptap/extension-superscript';

declare module '@tiptap/core' {
  interface Commands<ReturnType> {
    documentCharacterFormatting: {
      toggleDocumentSubscript: () => ReturnType;
      toggleDocumentSuperscript: () => ReturnType;
    };
  }
}

export const DocumentSubscript = Subscript.extend({
  excludes: 'superscript',
  addCommands() {
    return {
      ...(this.parent?.() ?? {}),
      toggleDocumentSubscript:
        () =>
        ({ chain, editor }) => {
          const commandChain = chain().focus();
          return editor.isActive('subscript')
            ? commandChain.unsetSubscript().run()
            : commandChain.unsetSuperscript().setSubscript().run();
        },
    };
  },
  addKeyboardShortcuts() {
    return {
      'Mod-=': () => this.editor.commands.toggleDocumentSubscript(),
    };
  },
});

export const DocumentSuperscript = Superscript.extend({
  excludes: 'subscript',
  addCommands() {
    return {
      ...(this.parent?.() ?? {}),
      toggleDocumentSuperscript:
        () =>
        ({ chain, editor }) => {
          const commandChain = chain().focus();
          return editor.isActive('superscript')
            ? commandChain.unsetSuperscript().run()
            : commandChain.unsetSubscript().setSuperscript().run();
        },
    };
  },
  addKeyboardShortcuts() {
    return {
      'Mod-Shift-=': () => this.editor.commands.toggleDocumentSuperscript(),
    };
  },
});

export function toggleDocumentSubscript(editor: Editor): boolean {
  return editor.commands.toggleDocumentSubscript();
}

export function toggleDocumentSuperscript(editor: Editor): boolean {
  return editor.commands.toggleDocumentSuperscript();
}
