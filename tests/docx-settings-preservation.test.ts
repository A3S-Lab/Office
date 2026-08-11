import { describe, expect, test } from '@rstest/core';
import JSZip from 'jszip';
import { createDocxBlob } from '../src/internal/features/work/work-docx-export';
import {
  attribute,
  descendants,
  directChildren,
  parseXml,
} from '../src/internal/features/work/work-ooxml-package';
import { preserveDocxSourcePackage } from '../src/internal/features/work/work-ooxml-package-preservation';

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const STRICT_WORD_NAMESPACE =
  'http://purl.oclc.org/ooxml/wordprocessingml/main';
const MARKUP_COMPATIBILITY_NAMESPACE =
  'http://schemas.openxmlformats.org/markup-compatibility/2006';
const OFFICE_RELATIONSHIPS_NAMESPACE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PACKAGE_RELATIONSHIPS_NAMESPACE =
  'http://schemas.openxmlformats.org/package/2006/relationships';
const XMLNS_NAMESPACE = 'http://www.w3.org/2000/xmlns/';
const WORD_2016_NAMESPACE =
  'http://schemas.microsoft.com/office/word/2016/wordml/cid';
const VENDOR_NAMESPACE = 'https://schemas.a3s.dev/word/settings/2026';

describe('DOCX settings extension preservation', () => {
  test('retains passive settings extensions while generated Word settings stay authoritative', async () => {
    const generated = await generatedDocx();
    const source = await mutateSettings(generated, (document) => {
      const root = document.documentElement;
      root.setAttributeNS(XMLNS_NAMESPACE, 'xmlns:w16cid', WORD_2016_NAMESPACE);
      root.setAttributeNS(XMLNS_NAMESPACE, 'xmlns:vendor', VENDOR_NAMESPACE);
      root.setAttributeNS(
        MARKUP_COMPATIBILITY_NAMESPACE,
        'mc:Ignorable',
        `${attribute(root, 'mc:Ignorable') ?? ''} w16cid vendor`.trim(),
      );
      root.setAttributeNS(VENDOR_NAMESPACE, 'vendor:packageRevision', '7');
      root.setAttributeNS(
        MARKUP_COMPATIBILITY_NAMESPACE,
        'mc:PreserveAttributes',
        'vendor:value missing:value',
      );
      root.setAttributeNS(
        MARKUP_COMPATIBILITY_NAMESPACE,
        'mc:PreserveElements',
        'vendor:passiveSetting',
      );
      root.setAttributeNS(
        MARKUP_COMPATIBILITY_NAMESPACE,
        'mc:ProcessContent',
        'vendor:passiveSetting malformed',
      );

      const evenAndOddHeaders = directChildren(root, 'evenAndOddHeaders')[0];
      evenAndOddHeaders?.setAttributeNS(WORD_NAMESPACE, 'w:val', 'true');

      const compatibility = directChildren(root, 'compat')[0];
      if (!compatibility) throw new Error('Generated settings lack w:compat.');
      compatibility.setAttributeNS(
        VENDOR_NAMESPACE,
        'vendor:layoutRevision',
        '3',
      );
      compatibility.append(
        document.createElementNS(WORD_NAMESPACE, 'w:useFELayout'),
      );
      compatibility.append(
        document.createElementNS(
          VENDOR_NAMESPACE,
          'vendor:compatibilityExtension',
        ),
      );

      const alternateContent = document.createElementNS(
        MARKUP_COMPATIBILITY_NAMESPACE,
        'mc:AlternateContent',
      );
      const choice = document.createElementNS(
        MARKUP_COMPATIBILITY_NAMESPACE,
        'mc:Choice',
      );
      choice.setAttribute('Requires', 'w16cid');
      const documentId = document.createElementNS(
        WORD_2016_NAMESPACE,
        'w16cid:docId',
      );
      documentId.setAttributeNS(
        WORD_2016_NAMESPACE,
        'w16cid:val',
        '{A3S-2026}',
      );
      choice.append(documentId);
      const fallback = document.createElementNS(
        MARKUP_COMPATIBILITY_NAMESPACE,
        'mc:Fallback',
      );
      fallback.append(
        document.createElementNS(WORD_NAMESPACE, 'w:doNotTrackMoves'),
      );
      alternateContent.append(choice, fallback);

      const vendorSetting = document.createElementNS(
        VENDOR_NAMESPACE,
        'vendor:passiveSetting',
      );
      vendorSetting.setAttributeNS(VENDOR_NAMESPACE, 'vendor:value', 'keep');
      root.insertBefore(alternateContent, compatibility);
      root.insertBefore(vendorSetting, compatibility);

      const attachedTemplate = document.createElementNS(
        WORD_NAMESPACE,
        'w:attachedTemplate',
      );
      attachedTemplate.setAttributeNS(
        OFFICE_RELATIONSHIPS_NAMESPACE,
        'r:id',
        'rIdTemplate',
      );
      root.insertBefore(attachedTemplate, compatibility);
      const protection = document.createElementNS(
        WORD_NAMESPACE,
        'w:documentProtection',
      );
      protection.setAttributeNS(WORD_NAMESPACE, 'w:edit', 'readOnly');
      root.insertBefore(protection, compatibility);
    });

    const output = await preserveDocxSourcePackage(generated, source);
    const settings = await settingsDocument(output);
    const root = settings.documentElement;
    expect(attribute(root, 'vendor:packageRevision')).toBe('7');
    expect(
      (attribute(root, 'mc:Ignorable') ?? '').split(/\s+/).filter(Boolean),
    ).toEqual(
      expect.arrayContaining(['w14', 'w15', 'wp14', 'w16cid', 'vendor']),
    );
    expect(directChildren(root, 'AlternateContent')).toHaveLength(1);
    expect(descendants(root, 'docId')).toHaveLength(1);
    expect(attribute(descendants(root, 'docId')[0], 'val')).toBe('{A3S-2026}');
    expect(descendants(root, 'doNotTrackMoves')).toHaveLength(1);
    expect(descendants(root, 'passiveSetting')).toHaveLength(1);
    expect(attribute(root, 'mc:PreserveAttributes')).toBe('vendor:value');
    expect(attribute(root, 'mc:PreserveElements')).toBe(
      'vendor:passiveSetting',
    );
    expect(attribute(root, 'mc:ProcessContent')).toBe('vendor:passiveSetting');

    const compatibility = directChildren(root, 'compat')[0];
    expect(
      compatibility ? attribute(compatibility, 'vendor:layoutRevision') : null,
    ).toBe('3');
    expect(descendants(compatibility, 'compatibilityExtension')).toHaveLength(
      1,
    );
    expect(descendants(compatibility, 'useFELayout')).toHaveLength(0);
    expect(directChildren(root, 'attachedTemplate')).toHaveLength(0);
    expect(directChildren(root, 'documentProtection')).toHaveLength(0);
    expect(attribute(directChildren(root, 'evenAndOddHeaders')[0], 'val')).toBe(
      'false',
    );

    const secondOutput = await preserveDocxSourcePackage(generated, output);
    const secondSettings = await settingsDocument(secondOutput);
    expect(
      directChildren(secondSettings.documentElement, 'AlternateContent'),
    ).toHaveLength(1);
    expect(descendants(secondSettings, 'passiveSetting')).toHaveLength(1);
    const ignorable = (
      attribute(secondSettings.documentElement, 'mc:Ignorable') ?? ''
    )
      .split(/\s+/)
      .filter(Boolean);
    expect(new Set(ignorable).size).toBe(ignorable.length);
    expect(
      (attribute(secondSettings.documentElement, 'mc:ProcessContent') ?? '')
        .split(/\s+/)
        .filter(Boolean),
    ).toEqual(['vendor:passiveSetting']);
  });

  test('drops malformed, relationship-bound, or duplicate settings extensions and their source relationships', async () => {
    const generated = await generatedDocx();
    const sourceArchive = await JSZip.loadAsync(generated);
    const settings = await xmlEntry(sourceArchive, 'word/settings.xml');
    const root = settings.documentElement;
    root.setAttributeNS(XMLNS_NAMESPACE, 'xmlns:vendor', VENDOR_NAMESPACE);
    root.setAttributeNS(
      MARKUP_COMPATIBILITY_NAMESPACE,
      'mc:Ignorable',
      `${attribute(root, 'mc:Ignorable') ?? ''} vendor`.trim(),
    );
    root.setAttributeNS(OFFICE_RELATIONSHIPS_NAMESPACE, 'r:id', 'rIdRoot');
    root.append(documentSetting(settings, 'zoom'));

    const linkedAlternateContent = alternateContent(
      settings,
      'linkedExtension',
      true,
    );
    root.append(linkedAlternateContent);

    const duplicateAlternateContent = settings.createElementNS(
      MARKUP_COMPATIBILITY_NAMESPACE,
      'mc:AlternateContent',
    );
    const duplicateChoice = settings.createElementNS(
      MARKUP_COMPATIBILITY_NAMESPACE,
      'mc:Choice',
    );
    duplicateChoice.setAttribute('Requires', 'w14');
    duplicateChoice.append(documentSetting(settings, 'evenAndOddHeaders'));
    const duplicateFallback = settings.createElementNS(
      MARKUP_COMPATIBILITY_NAMESPACE,
      'mc:Fallback',
    );
    duplicateFallback.append(documentSetting(settings, 'evenAndOddHeaders'));
    duplicateAlternateContent.append(duplicateChoice, duplicateFallback);
    root.append(duplicateAlternateContent);

    const missingRequires = alternateContent(
      settings,
      'missingRequires',
      false,
    );
    directChildren(missingRequires, 'Choice')[0]?.removeAttribute('Requires');
    root.append(missingRequires);

    const unexpectedAttribute = alternateContent(
      settings,
      'unexpectedAttribute',
      false,
    );
    unexpectedAttribute.setAttribute('bogus', 'true');
    root.append(unexpectedAttribute);

    const unexpectedText = alternateContent(settings, 'unexpectedText', false);
    directChildren(unexpectedText, 'Choice')[0]?.prepend(
      settings.createTextNode('invalid text'),
    );
    root.append(unexpectedText);

    const invalidOrder = settings.createElementNS(
      MARKUP_COMPATIBILITY_NAMESPACE,
      'mc:AlternateContent',
    );
    const earlyFallback = settings.createElementNS(
      MARKUP_COMPATIBILITY_NAMESPACE,
      'mc:Fallback',
    );
    const lateChoice = settings.createElementNS(
      MARKUP_COMPATIBILITY_NAMESPACE,
      'mc:Choice',
    );
    lateChoice.setAttribute('Requires', 'vendor');
    lateChoice.append(
      settings.createElementNS(VENDOR_NAMESPACE, 'vendor:lateChoice'),
    );
    invalidOrder.append(earlyFallback, lateChoice);
    root.append(invalidOrder);

    const nested = alternateContent(settings, 'outerExtension', false);
    directChildren(nested, 'Choice')[0]?.append(
      alternateContent(settings, 'nestedExtension', false),
    );
    root.append(nested);

    const linkedSetting = settings.createElementNS(
      VENDOR_NAMESPACE,
      'vendor:linkedSetting',
    );
    linkedSetting.setAttributeNS(
      OFFICE_RELATIONSHIPS_NAMESPACE,
      'r:id',
      'rIdVendor',
    );
    root.append(linkedSetting);
    const requiredSetting = settings.createElementNS(
      VENDOR_NAMESPACE,
      'vendor:requiresUnderstanding',
    );
    requiredSetting.setAttributeNS(
      MARKUP_COMPATIBILITY_NAMESPACE,
      'mc:MustUnderstand',
      'vendor',
    );
    root.append(requiredSetting);
    sourceArchive.file(
      'word/_rels/settings.xml.rels',
      `<Relationships xmlns="${PACKAGE_RELATIONSHIPS_NAMESPACE}"><Relationship Id="rIdTemplate" Type="${OFFICE_RELATIONSHIPS_NAMESPACE}/attachedTemplate" Target="https://templates.a3s.dev/remote.dotx" TargetMode="External"/></Relationships>`,
    );
    sourceArchive.file(
      'word/settings.xml',
      new XMLSerializer().serializeToString(settings),
    );

    const output = await preserveDocxSourcePackage(
      generated,
      await sourceArchive.generateAsync({ type: 'arraybuffer' }),
    );
    const outputArchive = await JSZip.loadAsync(output);
    const outputSettings = await xmlEntry(outputArchive, 'word/settings.xml');
    expect(directChildren(outputSettings.documentElement, 'zoom')).toHaveLength(
      0,
    );
    expect(
      directChildren(outputSettings.documentElement, 'AlternateContent'),
    ).toHaveLength(0);
    expect(descendants(outputSettings, 'linkedSetting')).toHaveLength(0);
    expect(descendants(outputSettings, 'requiresUnderstanding')).toHaveLength(
      0,
    );
    expect(attribute(outputSettings.documentElement, 'r:id')).toBeNull();
    expect(
      (attribute(outputSettings.documentElement, 'mc:Ignorable') ?? '').split(
        /\s+/,
      ),
    ).not.toContain('vendor');
    expect(outputArchive.file('word/_rels/settings.xml.rels')).toBeNull();
  });

  test('reads strict UTF-16 settings and emits a transitional UTF-8 result', async () => {
    const generated = await generatedDocx();
    const sourceArchive = await JSZip.loadAsync(generated);
    const sourceText = await sourceArchive
      .file('word/settings.xml')
      ?.async('text');
    if (!sourceText) throw new Error('Generated settings are missing.');
    const strictSettings = parseXml(
      sourceText.replaceAll(WORD_NAMESPACE, STRICT_WORD_NAMESPACE),
      'strict word/settings.xml',
    );
    const root = strictSettings.documentElement;
    root.setAttributeNS(XMLNS_NAMESPACE, 'xmlns:vendor', VENDOR_NAMESPACE);
    root.setAttributeNS(
      MARKUP_COMPATIBILITY_NAMESPACE,
      'mc:Ignorable',
      `${attribute(root, 'mc:Ignorable') ?? ''} vendor`.trim(),
    );
    const extension = strictSettings.createElementNS(
      VENDOR_NAMESPACE,
      'vendor:utf16Setting',
    );
    extension.setAttributeNS(VENDOR_NAMESPACE, 'vendor:value', '保留');
    root.append(extension);
    const sourceOnlyWordSetting = strictSettings.createElementNS(
      STRICT_WORD_NAMESPACE,
      'w:zoom',
    );
    sourceOnlyWordSetting.setAttributeNS(
      STRICT_WORD_NAMESPACE,
      'w:percent',
      '175',
    );
    root.append(sourceOnlyWordSetting);
    const serialized = new XMLSerializer()
      .serializeToString(strictSettings)
      .replace(
        /^\s*<\?xml[^?]*\?>/i,
        '<?xml version="1.0" encoding="UTF-16" standalone="yes"?>',
      );
    sourceArchive.file('word/settings.xml', utf16LittleEndian(serialized));

    const output = await preserveDocxSourcePackage(
      generated,
      await sourceArchive.generateAsync({ type: 'arraybuffer' }),
    );
    const outputArchive = await JSZip.loadAsync(output);
    const outputBytes = await outputArchive
      .file('word/settings.xml')
      ?.async('uint8array');
    expect(outputBytes?.slice(0, 3)).toEqual(new Uint8Array([60, 63, 120]));
    const outputText = await outputArchive
      .file('word/settings.xml')
      ?.async('text');
    expect(outputText).toContain('encoding="UTF-8"');
    const outputSettings = await xmlEntry(outputArchive, 'word/settings.xml');
    expect(outputSettings.documentElement.namespaceURI).toBe(WORD_NAMESPACE);
    expect(descendants(outputSettings, 'utf16Setting')).toHaveLength(1);
    expect(
      attribute(descendants(outputSettings, 'utf16Setting')[0], 'value'),
    ).toBe('保留');
    expect(directChildren(outputSettings.documentElement, 'zoom')).toHaveLength(
      0,
    );
  });
});

async function generatedDocx(): Promise<ArrayBuffer> {
  return (
    await createDocxBlob({
      type: 'document',
      html: '<p>Settings preservation boundary</p>',
      pageSize: 'a4',
    })
  ).arrayBuffer();
}

async function mutateSettings(
  source: ArrayBuffer,
  mutate: (document: Document) => void,
): Promise<ArrayBuffer> {
  const archive = await JSZip.loadAsync(source);
  const settings = await xmlEntry(archive, 'word/settings.xml');
  mutate(settings);
  archive.file(
    'word/settings.xml',
    new XMLSerializer().serializeToString(settings),
  );
  return archive.generateAsync({ type: 'arraybuffer' });
}

async function settingsDocument(source: ArrayBuffer): Promise<Document> {
  const archive = await JSZip.loadAsync(source);
  return xmlEntry(archive, 'word/settings.xml');
}

async function xmlEntry(archive: JSZip, path: string): Promise<Document> {
  const source = await archive.file(path)?.async('text');
  if (!source) throw new Error(`Missing OOXML part: ${path}`);
  return parseXml(source, path);
}

function documentSetting(document: Document, localName: string): Element {
  return document.createElementNS(WORD_NAMESPACE, `w:${localName}`);
}

function alternateContent(
  document: Document,
  localName: string,
  relationshipBound: boolean,
): Element {
  const container = document.createElementNS(
    MARKUP_COMPATIBILITY_NAMESPACE,
    'mc:AlternateContent',
  );
  const choice = document.createElementNS(
    MARKUP_COMPATIBILITY_NAMESPACE,
    'mc:Choice',
  );
  choice.setAttribute('Requires', 'vendor');
  const extension = document.createElementNS(
    VENDOR_NAMESPACE,
    `vendor:${localName}`,
  );
  if (relationshipBound) {
    extension.setAttributeNS(
      OFFICE_RELATIONSHIPS_NAMESPACE,
      'r:id',
      'rIdVendor',
    );
  }
  choice.append(extension);
  container.append(choice);
  return container;
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
