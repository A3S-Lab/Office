import { mergeAttributes, Node } from '@tiptap/core';
import { DOCUMENT_LAZY_BLOCK_NODE } from './work-document-lazy-model';

/** Internal equal-size placeholder used only by the large DOCX editor source. */
export const DocumentLazyBlock = Node.create({
  name: DOCUMENT_LAZY_BLOCK_NODE,
  group: 'block',
  content: 'text*',
  selectable: false,

  addAttributes() {
    return {
      chunkId: { default: '' },
      contentSize: { default: 2 },
      paragraphCount: { default: 0 },
    };
  },

  parseHTML() {
    return [{ tag: 'div[data-document-lazy-block]' }];
  },

  renderHTML({ HTMLAttributes }) {
    return [
      'div',
      mergeAttributes(HTMLAttributes, {
        'data-document-lazy-block': 'true',
        'aria-hidden': 'true',
      }),
      0,
    ];
  },
});
