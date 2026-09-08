import TextAlign from '@tiptap/extension-text-align';

export const DOCUMENT_DISTRIBUTE_ALIGN_ATTRIBUTE = 'data-office-text-align';
export const DOCUMENT_DISTRIBUTE_ALIGN_VALUE = 'distribute';

/**
 * TipTap TextAlign stores alignment in `style="text-align: …"`. Browsers reject
 * the non-standard `distribute` value and clear it, which TipTap then re-parses
 * as no alignment. Persist distribute via a data attribute and paint with
 * valid justify + text-align-last:justify CSS (WPS 分散对齐).
 */
export const DocumentTextAlign = TextAlign.extend({
  addGlobalAttributes() {
    return [
      {
        types: this.options.types,
        attributes: {
          textAlign: {
            default: this.options.defaultAlignment,
            parseHTML: (element) => {
              if (
                element.getAttribute(DOCUMENT_DISTRIBUTE_ALIGN_ATTRIBUTE) ===
                DOCUMENT_DISTRIBUTE_ALIGN_VALUE
              ) {
                return DOCUMENT_DISTRIBUTE_ALIGN_VALUE;
              }
              const alignment = element.style.textAlign;
              return this.options.alignments.includes(alignment)
                ? alignment
                : this.options.defaultAlignment;
            },
            renderHTML: (attributes) => {
              if (!attributes.textAlign) {
                return {};
              }
              if (attributes.textAlign === DOCUMENT_DISTRIBUTE_ALIGN_VALUE) {
                return {
                  style: 'text-align: justify; text-align-last: justify',
                  [DOCUMENT_DISTRIBUTE_ALIGN_ATTRIBUTE]:
                    DOCUMENT_DISTRIBUTE_ALIGN_VALUE,
                };
              }
              return { style: `text-align: ${attributes.textAlign}` };
            },
          },
        },
      },
    ];
  },

  // Override TipTap TextAlign Mod-Shift-* chords that collide with WPS/Word:
  // Mod-Shift-l is bullets (not left align); Mod-Shift-j is distribute.
  addKeyboardShortcuts() {
    return {
      'Mod-Shift-e': () => this.editor.commands.setTextAlign('center'),
      'Mod-Shift-r': () => this.editor.commands.setTextAlign('right'),
      'Mod-Shift-j': () =>
        this.editor.commands.setTextAlign(DOCUMENT_DISTRIBUTE_ALIGN_VALUE),
    };
  },
});
