import { Extension } from '@tiptap/core';
import {
  copyDocumentFormatting,
  pasteDocumentFormatting,
} from './document-format-clipboard';

export const DocumentPageChromeWpsShortcuts = Extension.create({
  name: 'documentPageChromeWpsShortcuts',

  addKeyboardShortcuts() {
    return {
      'Mod-l': () => this.editor.chain().focus().setTextAlign('left').run(),
      'Mod-e': () => this.editor.chain().focus().setTextAlign('center').run(),
      'Mod-r': () => this.editor.chain().focus().setTextAlign('right').run(),
      'Mod-j': () => this.editor.chain().focus().setTextAlign('justify').run(),
      'Mod-Shift-c': () => {
        copyDocumentFormatting(this.editor);
        return true;
      },
      'Mod-Shift-v': () => {
        pasteDocumentFormatting(this.editor);
        return true;
      },
    };
  },
});
