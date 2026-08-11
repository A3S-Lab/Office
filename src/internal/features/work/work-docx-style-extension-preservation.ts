import type JSZip from 'jszip';
import {
  DOCX_WORDPROCESSING_NAMESPACES,
  mergeDocxIgnorableExtensions,
  type DocxExtensionDocumentRole,
} from './work-docx-ignorable-extension-preservation';
import { directChildren, parseXml } from './work-ooxml-package';
import {
  assertXmlRoot,
  xmlAttributeLocalName,
  xmlAttributeNamespace,
} from './work-docx-settings-xml';
import { decodeXmlBytes, serializeUtf8Xml } from './work-ooxml-xml';

const STYLES_PATH = 'word/styles.xml';
const MAX_STYLE_RECORDS = 8_192;

export async function preserveDocxStyleExtensions(
  generated: JSZip,
  source: JSZip,
  generatedPath = STYLES_PATH,
  sourcePath = STYLES_PATH,
): Promise<void> {
  const generatedEntry = generated.file(generatedPath);
  const sourceEntry = source.file(sourcePath);
  if (!generatedEntry || !sourceEntry) return;
  const [generatedBytes, sourceBytes] = await Promise.all([
    generatedEntry.async('uint8array'),
    sourceEntry.async('uint8array'),
  ]);
  const generatedDocument = parseXml(
    decodeXmlBytes(generatedBytes, `generated DOCX ${generatedPath}`),
    `generated DOCX ${generatedPath}`,
  );
  const sourceDocument = parseXml(
    decodeXmlBytes(sourceBytes, `source DOCX ${sourcePath}`),
    `source DOCX ${sourcePath}`,
  );
  assertXmlRoot(
    generatedDocument.documentElement,
    'styles',
    DOCX_WORDPROCESSING_NAMESPACES,
    'Generated DOCX word/styles.xml is not WordprocessingML styles.',
  );
  assertXmlRoot(
    sourceDocument.documentElement,
    'styles',
    DOCX_WORDPROCESSING_NAMESPACES,
    'Registered source DOCX word/styles.xml is not WordprocessingML styles.',
  );
  assertUniqueStyleIdentities(sourceDocument);
  mergeDocxIgnorableExtensions(generatedDocument, sourceDocument, {
    semanticKey: styleSemanticKey,
  });
  generated.file(generatedPath, serializeUtf8Xml(generatedDocument));
}

function assertUniqueStyleIdentities(document: Document): void {
  const styles = directChildren(document.documentElement, 'style').filter(
    (item) => DOCX_WORDPROCESSING_NAMESPACES.has(item.namespaceURI ?? ''),
  );
  if (styles.length > MAX_STYLE_RECORDS) {
    throw new Error('Registered source DOCX exceeds the style-record limit.');
  }
  const identities = new Set<string>();
  for (const style of styles) {
    const identity = styleIdentity(style);
    if (!identity) continue;
    if (identities.has(identity)) {
      throw new Error(
        'Registered source DOCX styles contain duplicate style identities.',
      );
    }
    identities.add(identity);
  }
}

function styleSemanticKey(
  element: Element,
  _role: DocxExtensionDocumentRole,
): string {
  if (!DOCX_WORDPROCESSING_NAMESPACES.has(element.namespaceURI ?? '')) {
    return `{${element.namespaceURI ?? ''}}${element.localName}`;
  }
  if (element.localName === 'style') {
    return `word:style:${styleIdentity(element) ?? 'unidentified'}`;
  }
  if (element.localName === 'tblStylePr') {
    return `word:tblStylePr:${wordAttribute(element, 'type') ?? ''}`;
  }
  if (element.localName === 'lsdException') {
    return `word:lsdException:${wordAttribute(element, 'name') ?? ''}`;
  }
  if (element.localName === 'tab') {
    return `word:tab:${wordAttribute(element, 'val') ?? ''}:${wordAttribute(element, 'pos') ?? ''}`;
  }
  return `word:${element.localName}`;
}

function styleIdentity(style: Element): string | null {
  const type = wordAttribute(style, 'type')?.toLowerCase();
  const id = wordAttribute(style, 'styleId');
  return type && id ? `${type}\u0000${id}` : null;
}

function wordAttribute(element: Element, name: string): string | null {
  for (const item of Array.from(element.attributes)) {
    if (
      DOCX_WORDPROCESSING_NAMESPACES.has(
        xmlAttributeNamespace(element, item) ?? '',
      ) &&
      xmlAttributeLocalName(item) === name
    ) {
      return item.value.trim() || null;
    }
  }
  return null;
}
