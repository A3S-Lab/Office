import { Extension } from '@tiptap/core';
import { normalizeDocumentHref } from './work-document-links';

export interface DocumentPageChromeImageOptions {
  alt: string;
  source: string;
}

declare module '@tiptap/core' {
  interface Commands<ReturnType> {
    documentPageChrome: {
      insertDocumentPageChromeImage: (
        options: DocumentPageChromeImageOptions,
      ) => ReturnType;
      setDocumentPageChromeLink: (href: string | null) => ReturnType;
    };
  }
}

export const DocumentPageChromeCommands = Extension.create({
  name: 'documentPageChromeCommands',

  addCommands() {
    return {
      insertDocumentPageChromeImage:
        ({ alt, source }) =>
        ({ commands }) => {
          if (!isDocumentPageChromeImageSource(source)) return false;
          const label = alt.trim();
          return commands.setImage({
            src: source,
            alt: label || 'Image',
            title: label || undefined,
          });
        },
      setDocumentPageChromeLink:
        (candidate) =>
        ({ chain }) => {
          const command = chain().extendMarkRange('link');
          if (candidate === null) return command.unsetLink().run();
          const href = normalizeDocumentPageChromeHref(candidate);
          return href ? command.setLink({ href }).run() : false;
        },
    };
  },
});

export function normalizeDocumentPageChromeHref(value: string): string | null {
  return normalizeDocumentHref(value);
}

function isDocumentPageChromeImageSource(source: string): boolean {
  return /^(?:https?:|blob:|data:image\/)/i.test(source.trim());
}
