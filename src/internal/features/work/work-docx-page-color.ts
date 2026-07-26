import JSZip from 'jszip';
import {
  DEFAULT_DOCUMENT_PAGE_COLOR,
  normalizeDocumentPageColor,
} from './work-document-page-color';
import {
  attribute,
  directChild,
  parseXml,
  xmlNamespacePrefix,
} from './work-ooxml-package';

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

export function importDocxPageColor(document: Document): string | undefined {
  const background = directChild(document.documentElement, 'background');
  const color = attribute(background ?? document.documentElement, 'color');
  if (!background || !color || color.toLowerCase() === 'auto') return undefined;
  return normalizeDocumentPageColor(`#${color}`);
}

export async function patchDocxPageColor(
  buffer: ArrayBuffer,
  value: string | undefined,
): Promise<ArrayBuffer> {
  const pageColor = normalizeDocumentPageColor(value);
  if (!pageColor || pageColor === DEFAULT_DOCUMENT_PAGE_COLOR) return buffer;

  const archive = await JSZip.loadAsync(buffer);
  const documentEntry = archive.file('word/document.xml');
  if (!documentEntry) return buffer;
  const document = parseXml(
    await documentEntry.async('text'),
    'word/document.xml',
  );
  upsertDocumentBackground(document, pageColor.slice(1).toUpperCase());
  archive.file(
    'word/document.xml',
    new XMLSerializer().serializeToString(document),
  );

  const settingsEntry = archive.file('word/settings.xml');
  if (settingsEntry) {
    const settings = parseXml(
      await settingsEntry.async('text'),
      'word/settings.xml',
    );
    if (!directChild(settings.documentElement, 'displayBackgroundShape')) {
      settings.documentElement.append(
        settings.createElementNS(
          WORD_NAMESPACE,
          wordQualifiedName(settings.documentElement, 'displayBackgroundShape'),
        ),
      );
    }
    archive.file(
      'word/settings.xml',
      new XMLSerializer().serializeToString(settings),
    );
  }

  return archive.generateAsync({ type: 'arraybuffer' });
}

function upsertDocumentBackground(document: Document, color: string): void {
  const root = document.documentElement;
  let background = directChild(root, 'background');
  if (!background) {
    background = document.createElementNS(
      WORD_NAMESPACE,
      wordQualifiedName(root, 'background'),
    );
    root.insertBefore(background, directChild(root, 'body') ?? root.firstChild);
  }
  background.setAttributeNS(
    WORD_NAMESPACE,
    wordQualifiedName(root, 'color'),
    color,
  );
}

function wordQualifiedName(root: Element, localName: string): string {
  return `${xmlNamespacePrefix(root, WORD_NAMESPACE) ?? root.prefix ?? 'w'}:${localName}`;
}
