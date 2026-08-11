import type JSZip from 'jszip';
import {
  directChildren,
  parseXml,
  resolvePartTarget,
  xmlNamespacePrefix,
} from './work-ooxml-package';
import { decodeXmlBytes, serializeUtf8Xml } from './work-ooxml-xml';
import {
  assertXmlRoot,
  hasNonWhitespaceXmlText,
  xmlAttributeLocalName,
  xmlAttributeNamespace,
} from './work-docx-settings-xml';
import type { PreservedRelationshipReference } from './work-ooxml-relationship-preservation';

const FONT_TABLE_PATH = 'word/fontTable.xml';
const TRANSITIONAL_WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const TRANSITIONAL_RELATIONSHIP_NAMESPACE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const WORD_NAMESPACES = new Set([
  TRANSITIONAL_WORD_NAMESPACE,
  'http://purl.oclc.org/ooxml/wordprocessingml/main',
]);
const RELATIONSHIP_NAMESPACES = new Set([
  TRANSITIONAL_RELATIONSHIP_NAMESPACE,
  'http://purl.oclc.org/ooxml/officeDocument/relationships',
]);
const FONT_CHILD_ORDER = [
  'altName',
  'panose1',
  'charset',
  'family',
  'notTrueType',
  'pitch',
  'sig',
  'embedRegular',
  'embedBold',
  'embedItalic',
  'embedBoldItalic',
] as const;
const FONT_CHILDREN = new Set<string>(FONT_CHILD_ORDER);
const EMBEDDED_FONT_CHILDREN = new Set([
  'embedRegular',
  'embedBold',
  'embedItalic',
  'embedBoldItalic',
]);
const MAX_FONT_RECORDS = 4_096;
const MAX_FONT_NAME_LENGTH = 256;

export async function preserveDocxFontTable(
  generated: JSZip,
  source: JSZip,
  relationships: ReadonlyMap<string, PreservedRelationshipReference>,
  fontPartPaths: ReadonlySet<string>,
  generatedPath = FONT_TABLE_PATH,
  sourcePath = FONT_TABLE_PATH,
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
    'fonts',
    WORD_NAMESPACES,
    'Generated DOCX word/fontTable.xml is not a WordprocessingML font table.',
  );
  assertXmlRoot(
    sourceDocument.documentElement,
    'fonts',
    WORD_NAMESPACES,
    'Registered source DOCX word/fontTable.xml is not a WordprocessingML font table.',
  );

  const generatedFonts = indexFonts(generatedDocument, 'Generated DOCX');
  const sourceFonts = indexFonts(sourceDocument, 'Registered source DOCX');
  const generatedRoot = generatedDocument.documentElement;
  const wordNamespace =
    generatedRoot.namespaceURI ?? TRANSITIONAL_WORD_NAMESPACE;
  const wordPrefix = xmlNamespacePrefix(generatedRoot, wordNamespace) ?? 'w';
  const relationshipPrefix =
    xmlNamespacePrefix(generatedRoot, TRANSITIONAL_RELATIONSHIP_NAMESPACE) ??
    'r';

  for (const sourceFont of sourceFonts.values()) {
    const key = fontKey(fontName(sourceFont));
    let generatedFont = generatedFonts.get(key);
    if (!generatedFont) {
      generatedFont = generatedDocument.createElementNS(
        wordNamespace,
        `${wordPrefix}:font`,
      );
      generatedFont.setAttributeNS(
        wordNamespace,
        `${wordPrefix}:name`,
        fontName(sourceFont),
      );
      generatedRoot.append(generatedFont);
      generatedFonts.set(key, generatedFont);
    }
    mergeFontChildren(
      generatedDocument,
      generatedFont,
      sourceFont,
      relationships,
      fontPartPaths,
      wordNamespace,
      wordPrefix,
      relationshipPrefix,
    );
  }
  generated.file(generatedPath, serializeUtf8Xml(generatedDocument));
}

function indexFonts(document: Document, label: string): Map<string, Element> {
  const fonts = directChildren(document.documentElement, 'font').filter(
    (item) => WORD_NAMESPACES.has(item.namespaceURI ?? ''),
  );
  if (fonts.length > MAX_FONT_RECORDS) {
    throw new Error(
      `${label} font table exceeds the supported font-record limit.`,
    );
  }
  const result = new Map<string, Element>();
  for (const font of fonts) {
    const name = fontName(font);
    if (!name || name.length > MAX_FONT_NAME_LENGTH) {
      throw new Error(`${label} font table contains an invalid font name.`);
    }
    const key = fontKey(name);
    if (result.has(key)) {
      throw new Error(`${label} font table contains duplicate font names.`);
    }
    result.set(key, font);
  }
  return result;
}

function mergeFontChildren(
  document: Document,
  generated: Element,
  source: Element,
  relationships: ReadonlyMap<string, PreservedRelationshipReference>,
  fontPartPaths: ReadonlySet<string>,
  wordNamespace: string,
  wordPrefix: string,
  relationshipPrefix: string,
): void {
  const generatedNames = new Set(
    directChildren(generated)
      .filter((item) => WORD_NAMESPACES.has(item.namespaceURI ?? ''))
      .map((item) => item.localName),
  );
  const sourceChildren = supportedFontChildren(source);
  for (const localName of FONT_CHILD_ORDER) {
    if (generatedNames.has(localName)) continue;
    const sourceChild = sourceChildren.get(localName);
    if (!sourceChild) continue;
    const child = copyFontChild(
      document,
      sourceChild,
      relationships,
      fontPartPaths,
      wordNamespace,
      wordPrefix,
      relationshipPrefix,
    );
    if (!child) continue;
    insertFontChild(generated, child);
    generatedNames.add(localName);
  }
}

function supportedFontChildren(font: Element): Map<string, Element> {
  const result = new Map<string, Element>();
  for (const child of directChildren(font)) {
    if (
      !WORD_NAMESPACES.has(child.namespaceURI ?? '') ||
      !FONT_CHILDREN.has(child.localName)
    ) {
      continue;
    }
    if (result.has(child.localName)) {
      throw new Error(
        `Registered source DOCX font record ${fontName(font)} contains duplicate ${child.localName} settings.`,
      );
    }
    result.set(child.localName, child);
  }
  return result;
}

function copyFontChild(
  document: Document,
  source: Element,
  relationships: ReadonlyMap<string, PreservedRelationshipReference>,
  fontPartPaths: ReadonlySet<string>,
  wordNamespace: string,
  wordPrefix: string,
  relationshipPrefix: string,
): Element | null {
  if (directChildren(source).length || hasNonWhitespaceXmlText(source)) {
    return null;
  }
  const embedded = EMBEDDED_FONT_CHILDREN.has(source.localName);
  const sourceRelationshipId = embedded
    ? relationshipAttribute(source, 'id')
    : null;
  const relationship = sourceRelationshipId
    ? relationships.get(sourceRelationshipId)
    : undefined;
  if (
    embedded &&
    (!relationship || !isInternalFontRelationship(relationship, fontPartPaths))
  ) {
    return null;
  }
  const result = document.createElementNS(
    wordNamespace,
    `${wordPrefix}:${source.localName}`,
  );
  for (const item of Array.from(source.attributes)) {
    const namespace = xmlAttributeNamespace(source, item);
    const localName = xmlAttributeLocalName(item);
    if (namespace && WORD_NAMESPACES.has(namespace)) {
      result.setAttributeNS(
        wordNamespace,
        `${wordPrefix}:${localName}`,
        item.value,
      );
    }
  }
  if (embedded && relationship) {
    result.setAttributeNS(
      TRANSITIONAL_RELATIONSHIP_NAMESPACE,
      `${relationshipPrefix}:id`,
      relationship.id,
    );
  }
  return result;
}

function insertFontChild(font: Element, child: Element): void {
  const childOrder = FONT_CHILD_ORDER.indexOf(
    child.localName as (typeof FONT_CHILD_ORDER)[number],
  );
  const anchor = directChildren(font).find((item) => {
    const order = FONT_CHILD_ORDER.indexOf(
      item.localName as (typeof FONT_CHILD_ORDER)[number],
    );
    return order >= 0 && order > childOrder;
  });
  font.insertBefore(child, anchor ?? null);
}

function fontName(font: Element): string {
  for (const item of Array.from(font.attributes)) {
    if (
      WORD_NAMESPACES.has(xmlAttributeNamespace(font, item) ?? '') &&
      xmlAttributeLocalName(item) === 'name'
    ) {
      return item.value.trim();
    }
  }
  return '';
}

function relationshipAttribute(element: Element, localName: string): string {
  for (const item of Array.from(element.attributes)) {
    if (
      RELATIONSHIP_NAMESPACES.has(xmlAttributeNamespace(element, item) ?? '') &&
      xmlAttributeLocalName(item) === localName
    ) {
      return item.value.trim();
    }
  }
  return '';
}

function isInternalFontRelationship(
  relationship: PreservedRelationshipReference,
  fontPartPaths: ReadonlySet<string>,
): boolean {
  const targetMode = relationship.targetMode?.toLowerCase();
  return (
    (!targetMode || targetMode === 'internal') &&
    /\/relationships\/font$/i.test(relationship.type) &&
    fontPartPaths.has(
      resolvePartTarget(FONT_TABLE_PATH, relationship.target).toLowerCase(),
    )
  );
}

function fontKey(name: string): string {
  return name.toLowerCase();
}
