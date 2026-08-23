import { Extension, type Editor } from '@tiptap/core';
import Subscript from '@tiptap/extension-subscript';
import Superscript from '@tiptap/extension-superscript';
import { normalizeDocumentCharacterScalePercent } from './work-document-character-scale';
import { normalizeDocumentCharacterPositionHalfPoints } from './work-document-character-position';
import { normalizeDocumentCharacterSpacingTwips } from './work-document-character-spacing';
import {
  normalizeDocumentEmphasisMark,
  type WorkDocumentEmphasisMark,
} from './work-document-emphasis';
import { normalizeDocumentKerningThresholdHalfPoints } from './work-document-kerning';
import {
  documentTextCaseKeyboardShortcuts,
  normalizeDocumentTextCase,
  type WorkDocumentTextCase,
} from './work-document-text-case';

declare module '@tiptap/core' {
  interface Commands<ReturnType> {
    documentCharacterFormatting: {
      setDocumentCharacterScale: (characterScalePercent: number) => ReturnType;
      toggleDocumentSubscript: () => ReturnType;
      toggleDocumentSuperscript: () => ReturnType;
      setDocumentCharacterSpacing: (
        characterSpacingTwips: number,
      ) => ReturnType;
      setDocumentCharacterPosition: (
        characterPositionHalfPoints: number,
      ) => ReturnType;
      setDocumentKerningThreshold: (
        kerningThresholdHalfPoints: number,
      ) => ReturnType;
      unsetDocumentKerningThreshold: () => ReturnType;
      setDocumentEmphasisMark: (
        emphasisMark: WorkDocumentEmphasisMark,
      ) => ReturnType;
      unsetDocumentEmphasisMark: () => ReturnType;
      setDocumentTextCase: (textCase: WorkDocumentTextCase) => ReturnType;
      toggleDocumentTextCase: (
        textCase: Exclude<WorkDocumentTextCase, 'none'>,
      ) => ReturnType;
    };
  }
}

export const DocumentCharacterFormatting = Extension.create({
  name: 'documentCharacterFormatting',

  addCommands() {
    return {
      setDocumentCharacterScale:
        (characterScalePercent: number) =>
        ({ commands }) => {
          const scale = normalizeDocumentCharacterScalePercent(
            characterScalePercent,
          );
          return scale === null
            ? false
            : commands.setMark('textStyle', {
                characterScalePercent: scale,
              });
        },
      setDocumentCharacterPosition:
        (characterPositionHalfPoints: number) =>
        ({ commands }) => {
          const position = normalizeDocumentCharacterPositionHalfPoints(
            characterPositionHalfPoints,
          );
          return position === null
            ? false
            : commands.setMark('textStyle', {
                characterPositionHalfPoints: position,
              });
        },
      setDocumentCharacterSpacing:
        (characterSpacingTwips: number) =>
        ({ commands }) => {
          const spacing = normalizeDocumentCharacterSpacingTwips(
            characterSpacingTwips,
          );
          return spacing === null
            ? false
            : commands.setMark('textStyle', {
                characterSpacingTwips: spacing,
              });
        },
      setDocumentKerningThreshold:
        (kerningThresholdHalfPoints: number) =>
        ({ commands }) => {
          const threshold = normalizeDocumentKerningThresholdHalfPoints(
            kerningThresholdHalfPoints,
          );
          return threshold === null
            ? false
            : commands.setMark('textStyle', {
                kerningThresholdHalfPoints: threshold,
              });
        },
      unsetDocumentKerningThreshold:
        () =>
        ({ chain }) =>
          chain()
            .setMark('textStyle', { kerningThresholdHalfPoints: null })
            .removeEmptyTextStyle()
            .run(),
      setDocumentEmphasisMark:
        (emphasisMark: WorkDocumentEmphasisMark) =>
        ({ commands }) => {
          const mark = normalizeDocumentEmphasisMark(emphasisMark);
          return mark === null
            ? false
            : commands.setMark('textStyle', { emphasisMark: mark });
        },
      unsetDocumentEmphasisMark:
        () =>
        ({ chain }) =>
          chain()
            .setMark('textStyle', { emphasisMark: null })
            .removeEmptyTextStyle()
            .run(),
      setDocumentTextCase:
        (textCase: WorkDocumentTextCase) =>
        ({ chain }) =>
          chain().focus().setMark('textStyle', { textCase }).run(),
      toggleDocumentTextCase:
        (textCase: Exclude<WorkDocumentTextCase, 'none'>) =>
        ({ commands, editor }) => {
          const current =
            normalizeDocumentTextCase(
              editor.getAttributes('textStyle').textCase,
            ) ?? 'none';
          return commands.setDocumentTextCase(
            current === textCase ? 'none' : textCase,
          );
        },
    };
  },

  addKeyboardShortcuts() {
    return {
      [documentTextCaseKeyboardShortcuts.allCaps]: () =>
        this.editor.commands.toggleDocumentTextCase('all-caps'),
      [documentTextCaseKeyboardShortcuts.smallCaps]: () =>
        this.editor.commands.toggleDocumentTextCase('small-caps'),
    };
  },
});

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
