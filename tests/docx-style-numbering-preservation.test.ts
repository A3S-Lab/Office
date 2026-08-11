import { describe, expect, test } from '@rstest/core';
import JSZip from 'jszip';
import { createDocxBlob } from '../src/internal/features/work/work-docx-export';
import {
  attribute,
  descendants,
  directChild,
  directChildren,
  parseXml,
} from '../src/internal/features/work/work-ooxml-package';
import {
  xmlAttributeLocalName,
  xmlAttributeNamespace,
  xmlNamespaceUri,
} from '../src/internal/features/work/work-docx-settings-xml';

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const STRICT_WORD_NAMESPACE =
  'http://purl.oclc.org/ooxml/wordprocessingml/main';
const RELATIONSHIPS_NAMESPACE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const MARKUP_COMPATIBILITY_NAMESPACE =
  'http://schemas.openxmlformats.org/markup-compatibility/2006';
const XMLNS_NAMESPACE = 'http://www.w3.org/2000/xmlns/';
const VENDOR_NAMESPACE = 'urn:a3s:test:style-numbering';

describe('DOCX style and numbering extension preservation', () => {
  test('maps passive extensions onto stable generated style and list identities', async () => {
    const seed = await createDocxBlob(documentContent(listHtml(91, 71)));
    const source = await JSZip.loadAsync(await seed.arrayBuffer());
    const styles = await xmlEntry(source, 'word/styles.xml');
    const stylesRoot = styles.documentElement;
    declareIgnorableVendor(stylesRoot);
    appendVendor(stylesRoot, 'styleCatalog', 'root-style-metadata');
    const heading = styleById(styles, 'Heading1');
    heading.setAttributeNS(VENDOR_NAMESPACE, 'vnd:stable', 'heading-metadata');
    appendVendor(heading, 'styleMeta', 'heading-extension');
    const runProperties = directChild(heading, 'rPr');
    if (!runProperties) throw new Error('Generated Heading1 rPr is missing.');
    appendVendor(runProperties, 'runMeta', 'nested-extension');
    const sourceColor = descendants(heading, 'color')[0];
    if (!sourceColor) throw new Error('Generated Heading1 color is missing.');
    setWordAttribute(sourceColor, 'val', 'FF0000');
    const unsafeStyle = appendVendor(heading, 'unsafeStyle', 'drop-me');
    unsafeStyle.setAttributeNS(
      RELATIONSHIPS_NAMESPACE,
      'r:id',
      'rIdUnsafeStyle',
    );
    const sourceOnlyStyle = styles.createElementNS(WORD_NAMESPACE, 'w:style');
    setWordAttribute(sourceOnlyStyle, 'type', 'paragraph');
    setWordAttribute(sourceOnlyStyle, 'styleId', 'SourceOnlyStyle');
    appendVendor(sourceOnlyStyle, 'styleMeta', 'must-not-reattach');
    stylesRoot.append(sourceOnlyStyle);

    const numbering = await xmlEntry(source, 'word/numbering.xml');
    const numberingRoot = numbering.documentElement;
    declareIgnorableVendor(numberingRoot);
    appendVendor(numberingRoot, 'numberingCatalog', 'root-list-metadata');
    const sourceAbstract = directChildren(numberingRoot, 'abstractNum').at(-1);
    if (!sourceAbstract)
      throw new Error('Generated abstract numbering is missing.');
    const oldAbstractId = wordAttribute(sourceAbstract, 'abstractNumId');
    if (!oldAbstractId) throw new Error('Generated abstract ID is missing.');
    setWordAttribute(sourceAbstract, 'abstractNumId', '71');
    sourceAbstract.setAttributeNS(
      VENDOR_NAMESPACE,
      'vnd:stable',
      'abstract-metadata',
    );
    appendVendor(sourceAbstract, 'abstractMeta', 'abstract-extension');
    const sourceLevel = directChildren(sourceAbstract, 'lvl').find(
      (item) => wordAttribute(item, 'ilvl') === '0',
    );
    if (!sourceLevel) throw new Error('Generated level zero is missing.');
    appendVendor(sourceLevel, 'levelMeta', 'level-extension');
    const sourceFormat = directChild(sourceLevel, 'numFmt');
    if (!sourceFormat) throw new Error('Generated number format is missing.');
    setWordAttribute(sourceFormat, 'val', 'upperRoman');
    const sourceNum = directChildren(numberingRoot, 'num').find(
      (item) =>
        wordAttribute(directChild(item, 'abstractNumId'), 'val') ===
        oldAbstractId,
    );
    if (!sourceNum) throw new Error('Generated concrete numbering is missing.');
    const oldNumId = wordAttribute(sourceNum, 'numId');
    if (!oldNumId) throw new Error('Generated concrete ID is missing.');
    setWordAttribute(sourceNum, 'numId', '91');
    const sourceAbstractReference = directChild(sourceNum, 'abstractNumId');
    if (!sourceAbstractReference) {
      throw new Error('Generated abstract numbering reference is missing.');
    }
    setWordAttribute(sourceAbstractReference, 'val', '71');
    appendVendor(sourceNum, 'numMeta', 'concrete-extension');
    const unsafeNumbering = appendVendor(
      sourceAbstract,
      'unsafeNumbering',
      'drop-me',
    );
    unsafeNumbering.setAttributeNS(
      RELATIONSHIPS_NAMESPACE,
      'r:id',
      'rIdUnsafeNumbering',
    );
    const sourceOnlyAbstract = numbering.createElementNS(
      WORD_NAMESPACE,
      'w:abstractNum',
    );
    setWordAttribute(sourceOnlyAbstract, 'abstractNumId', '72');
    appendVendor(sourceOnlyAbstract, 'abstractMeta', 'must-not-reattach');
    numberingRoot.append(sourceOnlyAbstract);

    const document = await xmlEntry(source, 'word/document.xml');
    for (const numId of descendants(document, 'numId')) {
      if (wordAttribute(numId, 'val') === oldNumId) {
        setWordAttribute(numId, 'val', '91');
      }
    }
    source.file('word/document.xml', serializeXml(document));
    source.file('word/styles.xml', strictUtf16(styles));
    source.file('word/numbering.xml', strictUtf16(numbering));
    const sourceBytes = await source.generateAsync({ type: 'arraybuffer' });

    const first = await createDocxBlob(
      documentContent(listHtml(91, 71)),
      sourceBytes,
    );
    const firstArchive = await JSZip.loadAsync(await first.arrayBuffer());
    const firstStyles = await xmlEntry(firstArchive, 'word/styles.xml');
    const firstHeading = styleById(firstStyles, 'Heading1');
    expect(vendorValue(firstStyles.documentElement, 'styleCatalog')).toBe(
      'root-style-metadata',
    );
    expect(vendorAttribute(firstHeading, 'stable')).toBe('heading-metadata');
    expect(vendorValue(firstHeading, 'styleMeta')).toBe('heading-extension');
    expect(vendorValue(directChild(firstHeading, 'rPr'), 'runMeta')).toBe(
      'nested-extension',
    );
    expect(descendants(firstHeading, 'unsafeStyle')).toEqual([]);
    expect(
      wordAttribute(descendants(firstHeading, 'color')[0], 'val'),
    ).not.toBe('FF0000');
    expect(() => styleById(firstStyles, 'SourceOnlyStyle')).toThrow(
      'Missing style',
    );
    expect(ignorableNamespaces(firstStyles.documentElement)).toContain(
      VENDOR_NAMESPACE,
    );

    const firstNumbering = await xmlEntry(firstArchive, 'word/numbering.xml');
    expect(
      vendorValue(firstNumbering.documentElement, 'numberingCatalog'),
    ).toBe('root-list-metadata');
    const mappedAbstract = descendants(firstNumbering, 'abstractNum').find(
      (item) => vendorValue(item, 'abstractMeta') === 'abstract-extension',
    );
    if (!mappedAbstract)
      throw new Error('Mapped abstract numbering is missing.');
    const mappedAbstractId = wordAttribute(mappedAbstract, 'abstractNumId');
    expect(mappedAbstractId).not.toBe('71');
    expect(vendorAttribute(mappedAbstract, 'stable')).toBe('abstract-metadata');
    const mappedLevel = directChildren(mappedAbstract, 'lvl').find(
      (item) => wordAttribute(item, 'ilvl') === '0',
    );
    expect(vendorValue(mappedLevel, 'levelMeta')).toBe('level-extension');
    expect(wordAttribute(directChild(mappedLevel, 'numFmt'), 'val')).toBe(
      'decimal',
    );
    expect(descendants(mappedAbstract, 'unsafeNumbering')).toEqual([]);
    const mappedNum = directChildren(
      firstNumbering.documentElement,
      'num',
    ).find((item) => vendorValue(item, 'numMeta') === 'concrete-extension');
    if (!mappedNum) throw new Error('Mapped concrete numbering is missing.');
    const mappedNumId = wordAttribute(mappedNum, 'numId');
    expect(mappedNumId).not.toBe('91');
    expect(
      descendants(firstNumbering, 'abstractNum').some(
        (item) => wordAttribute(item, 'abstractNumId') === '72',
      ),
    ).toBe(false);
    expect(ignorableNamespaces(firstNumbering.documentElement)).toContain(
      VENDOR_NAMESPACE,
    );
    await expect(
      firstArchive.file('word/styles.xml')?.async('text'),
    ).resolves.toContain('encoding="UTF-8"');
    await expect(
      firstArchive.file('word/numbering.xml')?.async('text'),
    ).resolves.toContain('encoding="UTF-8"');

    if (!mappedNumId || !mappedAbstractId) {
      throw new Error('Mapped numbering identities are missing.');
    }
    const second = await createDocxBlob(
      documentContent(listHtml(Number(mappedNumId), Number(mappedAbstractId))),
      await first.arrayBuffer(),
    );
    const secondArchive = await JSZip.loadAsync(await second.arrayBuffer());
    const secondStyles = await xmlEntry(secondArchive, 'word/styles.xml');
    expect(vendorValue(styleById(secondStyles, 'Heading1'), 'styleMeta')).toBe(
      'heading-extension',
    );
    const secondNumbering = await xmlEntry(secondArchive, 'word/numbering.xml');
    expect(
      descendants(secondNumbering, 'abstractMeta').map(
        (item) => item.textContent,
      ),
    ).toEqual(['abstract-extension']);
  });

  test('rejects duplicate source style identities', async () => {
    const generated = await createDocxBlob(documentContent('<h1>Title</h1>'));
    const source = await JSZip.loadAsync(await generated.arrayBuffer());
    const styles = await xmlEntry(source, 'word/styles.xml');
    const heading = styleById(styles, 'Heading1');
    styles.documentElement.append(styles.importNode(heading, true));
    source.file('word/styles.xml', serializeXml(styles));
    await expect(
      createDocxBlob(
        documentContent('<h1>Edited title</h1>'),
        await source.generateAsync({ type: 'arraybuffer' }),
      ),
    ).rejects.toThrow('duplicate style identities');
  });

  test('drops abstract extensions when one source definition expands into multiple generated definitions', async () => {
    const html = `${listHtml(91, 71)}${listHtml(92, 71)}`;
    const seed = await createDocxBlob(documentContent(html));
    const source = await JSZip.loadAsync(await seed.arrayBuffer());
    const numbering = await xmlEntry(source, 'word/numbering.xml');
    const numberingRoot = numbering.documentElement;
    declareIgnorableVendor(numberingRoot);
    const sourceAbstracts = directChildren(numberingRoot, 'abstractNum').slice(
      -2,
    );
    if (sourceAbstracts.length !== 2) {
      throw new Error('Generated shared-abstract fixture is incomplete.');
    }
    const oldAbstractIds = sourceAbstracts.map((item) =>
      wordAttribute(item, 'abstractNumId'),
    );
    if (oldAbstractIds.some((item) => !item)) {
      throw new Error('Generated abstract numbering ID is missing.');
    }
    setWordAttribute(sourceAbstracts[0], 'abstractNumId', '71');
    appendVendor(sourceAbstracts[0], 'abstractMeta', 'ambiguous-extension');
    sourceAbstracts[1].remove();
    const sourceNums = oldAbstractIds.map((abstractId, index) => {
      const num = directChildren(numberingRoot, 'num').find(
        (item) =>
          wordAttribute(directChild(item, 'abstractNumId'), 'val') ===
          abstractId,
      );
      if (!num) throw new Error('Generated concrete numbering is missing.');
      setWordAttribute(num, 'numId', String(91 + index));
      const abstractReference = directChild(num, 'abstractNumId');
      if (!abstractReference) {
        throw new Error('Generated abstract numbering reference is missing.');
      }
      setWordAttribute(abstractReference, 'val', '71');
      appendVendor(num, 'numMeta', `concrete-${index + 1}`);
      return num;
    });
    expect(sourceNums).toHaveLength(2);
    source.file('word/numbering.xml', serializeXml(numbering));

    const output = await createDocxBlob(
      documentContent(html),
      await source.generateAsync({ type: 'arraybuffer' }),
    );
    const archive = await JSZip.loadAsync(await output.arrayBuffer());
    const outputNumbering = await xmlEntry(archive, 'word/numbering.xml');
    expect(descendants(outputNumbering, 'abstractMeta')).toEqual([]);
    expect(
      descendants(outputNumbering, 'numMeta')
        .map((item) => item.textContent)
        .sort(),
    ).toEqual(['concrete-1', 'concrete-2']);
  });

  test('rejects duplicate source numbering identities', async () => {
    const html = listHtml(91, 71);
    const generated = await createDocxBlob(documentContent(html));
    const source = await JSZip.loadAsync(await generated.arrayBuffer());
    const numbering = await xmlEntry(source, 'word/numbering.xml');
    const num = directChildren(numbering.documentElement, 'num').at(-1);
    if (!num) throw new Error('Generated concrete numbering is missing.');
    numbering.documentElement.append(numbering.importNode(num, true));
    source.file('word/numbering.xml', serializeXml(numbering));
    await expect(
      createDocxBlob(
        documentContent(html),
        await source.generateAsync({ type: 'arraybuffer' }),
      ),
    ).rejects.toThrow('duplicate numbering identities');
  });
});

function documentContent(html: string) {
  return { type: 'document' as const, html, pageSize: 'a4' as const };
}

function listHtml(numId: number, abstractNumId: number): string {
  return `<h1>Stable heading</h1><ol data-office-numbering-id="${numId}" data-office-abstract-numbering-id="${abstractNumId}" data-office-numbering-format="decimal"><li>Stable list</li></ol>`;
}

function declareIgnorableVendor(root: Element): void {
  root.setAttributeNS(XMLNS_NAMESPACE, 'xmlns:vnd', VENDOR_NAMESPACE);
  const current = (attribute(root, 'mc:Ignorable') ?? '').trim();
  root.setAttributeNS(
    MARKUP_COMPATIBILITY_NAMESPACE,
    'mc:Ignorable',
    [current, 'vnd'].filter(Boolean).join(' '),
  );
}

function appendVendor(
  parent: Element,
  localName: string,
  value: string,
): Element {
  const element = parent.ownerDocument.createElementNS(
    VENDOR_NAMESPACE,
    `vnd:${localName}`,
  );
  element.textContent = value;
  parent.append(element);
  return element;
}

function vendorValue(
  parent: ParentNode | null | undefined,
  localName: string,
): string | null {
  if (!parent) return null;
  return (
    Array.from(parent.children).find(
      (item) =>
        item.namespaceURI === VENDOR_NAMESPACE && item.localName === localName,
    )?.textContent ?? null
  );
}

function vendorAttribute(element: Element, localName: string): string | null {
  return (
    Array.from(element.attributes).find(
      (item) =>
        xmlAttributeLocalName(item) === localName &&
        xmlAttributeNamespace(element, item) === VENDOR_NAMESPACE,
    )?.value ?? null
  );
}

function styleById(document: Document, id: string): Element {
  const style = directChildren(document.documentElement, 'style').find(
    (item) => wordAttribute(item, 'styleId') === id,
  );
  if (!style) throw new Error(`Missing style: ${id}`);
  return style;
}

function wordAttribute(
  element: Element | null | undefined,
  name: string,
): string | null {
  if (!element) return null;
  return attribute(element, name) ?? attribute(element, `w:${name}`);
}

function setWordAttribute(element: Element, name: string, value: string): void {
  element.setAttributeNS(WORD_NAMESPACE, `w:${name}`, value);
}

function ignorableNamespaces(root: Element): string[] {
  return (attribute(root, 'mc:Ignorable') ?? '')
    .trim()
    .split(/\s+/)
    .filter(Boolean)
    .flatMap((prefix) => {
      const namespace = xmlNamespaceUri(root, prefix);
      return namespace ? [namespace] : [];
    });
}

async function xmlEntry(archive: JSZip, path: string): Promise<Document> {
  const source = await archive.file(path)?.async('text');
  if (!source) throw new Error(`Missing OOXML part: ${path}`);
  return parseXml(source, path);
}

function strictUtf16(document: Document): Uint8Array {
  return utf16LittleEndian(
    serializeXml(document)
      .replaceAll(WORD_NAMESPACE, STRICT_WORD_NAMESPACE)
      .replace(
        /^\s*<\?xml[^?]*\?>/i,
        '<?xml version="1.0" encoding="UTF-16" standalone="yes"?>',
      ),
  );
}

function serializeXml(document: Document): string {
  return new XMLSerializer().serializeToString(document);
}

function utf16LittleEndian(source: string): Uint8Array {
  const bytes = new Uint8Array(2 + source.length * 2);
  bytes[0] = 0xff;
  bytes[1] = 0xfe;
  for (let index = 0; index < source.length; index += 1) {
    const codeUnit = source.charCodeAt(index);
    bytes[2 + index * 2] = codeUnit & 0xff;
    bytes[3 + index * 2] = codeUnit >>> 8;
  }
  return bytes;
}
