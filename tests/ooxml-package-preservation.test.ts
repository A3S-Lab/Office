import { describe, expect, test } from '@rstest/core';
import JSZip from 'jszip';
import {
  createArtifactBlob,
  forgetSourceBlob,
  importOfficeFile,
  registerSourceBlob,
} from '../src/core';
import { createDocxBlob } from '../src/internal/features/work/work-docx-export';
import {
  attribute,
  directChildren,
  parseXml,
} from '../src/internal/features/work/work-ooxml-package';
import { preserveDocxSourcePackage } from '../src/internal/features/work/work-ooxml-package-preservation';

const CONTENT_TYPES_NAMESPACE =
  'http://schemas.openxmlformats.org/package/2006/content-types';
const RELATIONSHIPS_NAMESPACE =
  'http://schemas.openxmlformats.org/package/2006/relationships';
const OFFICE_RELATIONSHIPS_NAMESPACE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

describe('OOXML package preservation', () => {
  test('retains safe source parts and relationships without reviving active content', async () => {
    const generated = new JSZip();
    generated.file(
      '[Content_Types].xml',
      contentTypes(
        [
          ['rels', 'application/vnd.openxmlformats-package.relationships+xml'],
          ['xml', 'application/xml'],
          ['bin', 'application/octet-stream'],
        ],
        [
          [
            '/word/document.xml',
            'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
          ],
        ],
      ),
    );
    generated.file(
      '_rels/.rels',
      relationships([
        [
          'rId1',
          `${OFFICE_RELATIONSHIPS_NAMESPACE}/officeDocument`,
          'word/document.xml',
        ],
      ]),
    );
    generated.file('word/document.xml', '<document>generated</document>');
    generated.file('word/styles.xml', '<styles/>');
    generated.file(
      'word/_rels/document.xml.rels',
      relationships([
        ['rId1', `${OFFICE_RELATIONSHIPS_NAMESPACE}/styles`, 'styles.xml'],
      ]),
    );

    const source = new JSZip();
    source.file(
      '[Content_Types].xml',
      contentTypes(
        [
          ['rels', 'application/vnd.openxmlformats-package.relationships+xml'],
          ['xml', 'application/xml'],
          ['bin', 'application/vnd.openxmlformats-officedocument.oleObject'],
        ],
        [
          [
            '/word/document.xml',
            'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
          ],
          [
            '/customXml/itemProps2.xml',
            'application/vnd.openxmlformats-officedocument.customXmlProperties+xml',
          ],
          [
            '/docProps/custom.xml',
            'application/vnd.openxmlformats-officedocument.custom-properties+xml',
          ],
          [
            '/_xmlsignatures/sig1.xml',
            'application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml',
          ],
          ['/word/vbaProject.bin', 'application/vnd.ms-office.vbaProject'],
          [
            '/word/vendorData/disguised.bin',
            'application/vnd.ms-office.vbaProject',
          ],
        ],
      ),
    );
    source.file('word/document.xml', '<document>source</document>');
    source.file('customXml/item2.xml', '<customer-record id="42"/>');
    source.file('customXml/itemProps2.xml', '<properties source="vendor"/>');
    source.file(
      'customXml/_rels/item2.xml.rels',
      relationships([
        [
          'rId1',
          `${OFFICE_RELATIONSHIPS_NAMESPACE}/customXmlProps`,
          'itemProps2.xml',
        ],
        [
          'rId2',
          'https://a3s.dev/relationships/vendor-schema',
          'https://schemas.a3s.dev/customer-record.xsd',
          'External',
        ],
      ]),
    );
    source.file(
      'docProps/custom.xml',
      '<properties><tier>gold</tier></properties>',
    );
    source.file(
      'word/embeddings/object1.bin',
      new Uint8Array([0, 255, 12, 34, 128, 64]),
    );
    source.file('word/vbaProject.bin', new Uint8Array([86, 66, 65]));
    source.file('word/vendorData/disguised.bin', new Uint8Array([77, 65, 67]));
    source.file('word/activeX/activeX1.bin', new Uint8Array([1, 2, 3]));
    source.file('customUI/customUI.xml', '<customUI/>');
    source.file('_xmlsignatures/origin.sigs', '<SignatureOrigin/>');
    source.file('_xmlsignatures/sig1.xml', '<Signature/>');
    source.file(
      '_rels/.rels',
      relationships([
        [
          'rId1',
          `${OFFICE_RELATIONSHIPS_NAMESPACE}/officeDocument`,
          'word/document.xml',
        ],
        [
          'rId2',
          'http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties',
          'docProps/custom.xml',
        ],
        [
          'rId3',
          'http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin',
          '_xmlsignatures/origin.sigs',
        ],
      ]),
    );
    source.file(
      'word/_rels/document.xml.rels',
      relationships([
        [
          'rId1',
          `${OFFICE_RELATIONSHIPS_NAMESPACE}/customXml`,
          '../customXml/item2.xml',
        ],
        [
          'rId2',
          `${OFFICE_RELATIONSHIPS_NAMESPACE}/oleObject`,
          'embeddings/object1.bin',
        ],
        [
          'rId3',
          `${OFFICE_RELATIONSHIPS_NAMESPACE}/vbaProject`,
          'vbaProject.bin',
        ],
        [
          'rId4',
          `${OFFICE_RELATIONSHIPS_NAMESPACE}/attachedTemplate`,
          'https://templates.a3s.dev/remote.dotx',
          'External',
        ],
        [
          'rId5',
          'https://a3s.dev/relationships/vendor-data',
          'vendorData/disguised.bin',
        ],
      ]),
    );

    const output = await preserveDocxSourcePackage(
      await generated.generateAsync({ type: 'arraybuffer' }),
      await source.generateAsync({ type: 'arraybuffer' }),
    );
    const archive = await JSZip.loadAsync(output);

    await expect(
      archive.file('word/document.xml')?.async('text'),
    ).resolves.toBe('<document>generated</document>');
    await expect(
      archive.file('customXml/item2.xml')?.async('text'),
    ).resolves.toBe('<customer-record id="42"/>');
    await expect(
      archive.file('word/embeddings/object1.bin')?.async('uint8array'),
    ).resolves.toEqual(new Uint8Array([0, 255, 12, 34, 128, 64]));
    expect(archive.file('word/vbaProject.bin')).toBeNull();
    expect(archive.file('word/vendorData/disguised.bin')).toBeNull();
    expect(archive.file('word/activeX/activeX1.bin')).toBeNull();
    expect(archive.file('customUI/customUI.xml')).toBeNull();
    expect(archive.file('_xmlsignatures/origin.sigs')).toBeNull();
    expect(archive.file('_xmlsignatures/sig1.xml')).toBeNull();

    const documentRelationships = await xmlEntry(
      archive,
      'word/_rels/document.xml.rels',
    );
    const documentItems = directChildren(
      documentRelationships.documentElement,
      'Relationship',
    );
    expect(documentItems.map((item) => attribute(item, 'Id'))).toEqual([
      'rId1',
      'rId2',
      'rId3',
    ]);
    expect(
      documentItems.map((item) => [
        attribute(item, 'Type')?.split('/').pop(),
        attribute(item, 'Target'),
      ]),
    ).toEqual([
      ['styles', 'styles.xml'],
      ['customXml', '../customXml/item2.xml'],
      ['oleObject', 'embeddings/object1.bin'],
    ]);

    const customXmlRelationships = await xmlEntry(
      archive,
      'customXml/_rels/item2.xml.rels',
    );
    expect(
      directChildren(
        customXmlRelationships.documentElement,
        'Relationship',
      ).map((item) => [
        attribute(item, 'Target'),
        attribute(item, 'TargetMode'),
      ]),
    ).toContainEqual([
      'https://schemas.a3s.dev/customer-record.xsd',
      'External',
    ]);

    const rootRelationships = await xmlEntry(archive, '_rels/.rels');
    expect(
      directChildren(rootRelationships.documentElement, 'Relationship').map(
        (item) => attribute(item, 'Target'),
      ),
    ).toEqual(['word/document.xml', 'docProps/custom.xml']);

    const types = await xmlEntry(archive, '[Content_Types].xml');
    const overrides = directChildren(types.documentElement, 'Override').map(
      (item) => [attribute(item, 'PartName'), attribute(item, 'ContentType')],
    );
    expect(overrides).toContainEqual([
      '/word/embeddings/object1.bin',
      'application/vnd.openxmlformats-officedocument.oleObject',
    ]);
    expect(overrides).toContainEqual([
      '/customXml/itemProps2.xml',
      'application/vnd.openxmlformats-officedocument.customXmlProperties+xml',
    ]);
    expect(overrides.some(([path]) => path?.includes('vbaProject'))).toBe(
      false,
    );
    expect(overrides.some(([path]) => path?.includes('_xmlsignatures'))).toBe(
      false,
    );
  });

  test('keeps registered DOCX source state through import, edit, and export', async () => {
    const original = await createDocxBlob({
      type: 'document',
      html: '<p>Original OOXML boundary</p>',
      pageSize: 'a4',
    });
    const sourceArchive = await JSZip.loadAsync(await original.arrayBuffer());
    const vendorBytes = new Uint8Array([222, 173, 190, 239, 0, 129]);
    sourceArchive.file('word/vendorData/payload.bin', vendorBytes);
    sourceArchive.file('word/vbaProject.bin', new Uint8Array([86, 66, 65]));
    sourceArchive.file('_xmlsignatures/sig1.xml', '<Signature/>');
    sourceArchive.file('customUI/customUI.xml', '<customUI/>');
    const documentEntry = sourceArchive.file('word/document.xml');
    const documentXml = await documentEntry?.async('text');
    if (!documentXml) throw new Error('Missing generated Word document.');
    sourceArchive.file(
      'word/document.xml',
      documentXml.replace(
        '</w:body>',
        [
          '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">',
          '<mc:Choice Requires="w14"><w:p><w:r><w:t>Choice branch</w:t></w:r></w:p></mc:Choice>',
          '<mc:Fallback><w:p><w:r><w:t>Fallback branch</w:t></w:r></w:p></mc:Fallback>',
          '</mc:AlternateContent></w:body>',
        ].join(''),
      ),
    );
    await appendXmlEntry(
      sourceArchive,
      '[Content_Types].xml',
      `<Override PartName="/word/vendorData/payload.bin" ContentType="application/vnd.a3s.vendor-data"/>`,
    );
    await appendXmlEntry(
      sourceArchive,
      'word/_rels/document.xml.rels',
      '<Relationship Id="rId900" Type="https://a3s.dev/relationships/vendor-data" Target="vendorData/payload.bin"/>',
    );
    const sourceBytes = await sourceArchive.generateAsync({
      type: 'arraybuffer',
      compression: 'DEFLATE',
    });
    const sourceFile = new File([sourceBytes], 'complex-boundary.docx', {
      type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      lastModified: 1_786_406_400_000,
    });
    const artifact = await importOfficeFile(sourceFile);

    try {
      expect(artifact.source).toMatchObject({
        name: 'complex-boundary.docx',
        contentType: sourceFile.type,
        size: sourceFile.size,
        updatedAt: sourceFile.lastModified,
      });
      expect(artifact.source?.fingerprint).toMatch(/^sha256:[0-9a-f]{64}$/);
      expect(artifact.compatibility?.issues.map((issue) => issue.code)).toEqual(
        expect.arrayContaining([
          'docx.package-state',
          'docx.signatures.removed',
          'docx.active-content.removed',
          'docx.markup-compatibility',
        ]),
      );
      if (artifact.content.type !== 'document') {
        throw new Error('Expected an imported document.');
      }
      artifact.content = {
        ...artifact.content,
        html: artifact.content.html.replace('Original', 'Edited'),
        model: undefined,
      };

      const exported = await createArtifactBlob(artifact);
      const output = await JSZip.loadAsync(await exported.arrayBuffer());
      await expect(
        output.file('word/vendorData/payload.bin')?.async('uint8array'),
      ).resolves.toEqual(vendorBytes);
      await expect(
        output.file('word/document.xml')?.async('text'),
      ).resolves.toContain('Edited OOXML boundary');
      await expect(
        output.file('word/_rels/document.xml.rels')?.async('text'),
      ).resolves.toContain('vendorData/payload.bin');
      expect(output.file('word/vbaProject.bin')).toBeNull();
      expect(output.file('_xmlsignatures/sig1.xml')).toBeNull();
      expect(output.file('customUI/customUI.xml')).toBeNull();
      const reopened = await importOfficeFile(
        new File([exported], 'reopened-boundary.docx', { type: exported.type }),
      );
      try {
        expect(reopened.content).toMatchObject({
          type: 'document',
          html: expect.stringContaining('Edited OOXML boundary'),
        });
      } finally {
        forgetSourceBlob(reopened.id);
      }

      forgetSourceBlob(artifact.id);
      await expect(createArtifactBlob(artifact)).rejects.toThrow(
        'original file is not available',
      );
      const wrongSource = await createDocxBlob({
        type: 'document',
        html: '<p>Different DOCX source</p>',
        pageSize: 'a4',
      });
      registerSourceBlob(artifact.id, wrongSource);
      await expect(createArtifactBlob(artifact)).rejects.toThrow(
        'does not match the imported source',
      );
      registerSourceBlob(artifact.id, sourceFile);
      const recovered = await createArtifactBlob(artifact);
      const recoveredArchive = await JSZip.loadAsync(
        await recovered.arrayBuffer(),
      );
      expect(
        recoveredArchive.file('word/vendorData/payload.bin'),
      ).not.toBeNull();
    } finally {
      forgetSourceBlob(artifact.id);
    }
  });

  test('rejects ambiguous or untyped source package parts', async () => {
    const generated = await createDocxBlob({
      type: 'document',
      html: '<p>Safe package identity</p>',
      pageSize: 'a4',
    });
    const generatedBytes = await generated.arrayBuffer();
    const ambiguous = await JSZip.loadAsync(generatedBytes);
    ambiguous.file('word/vendor/Case.xml', '<first/>');
    ambiguous.file('word/vendor/case.xml', '<second/>');

    await expect(
      preserveDocxSourcePackage(
        generatedBytes,
        await ambiguous.generateAsync({ type: 'arraybuffer' }),
      ),
    ).rejects.toThrow('case-ambiguous part names');

    const untyped = await JSZip.loadAsync(generatedBytes);
    untyped.file('word/vendor/untyped', new Uint8Array([1, 2, 3]));
    await expect(
      preserveDocxSourcePackage(
        generatedBytes,
        await untyped.generateAsync({ type: 'arraybuffer' }),
      ),
    ).rejects.toThrow('part has no content type');

    const wrongKind = new JSZip();
    wrongKind.file(
      '[Content_Types].xml',
      contentTypes(
        [
          ['rels', 'application/vnd.openxmlformats-package.relationships+xml'],
          ['xml', 'application/xml'],
        ],
        [
          [
            '/xl/workbook.xml',
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
          ],
        ],
      ),
    );
    wrongKind.file('xl/workbook.xml', '<workbook/>');
    await expect(
      preserveDocxSourcePackage(
        generatedBytes,
        await wrongKind.generateAsync({ type: 'arraybuffer' }),
      ),
    ).rejects.toThrow('not a WordprocessingML document');
  });
});

function contentTypes(
  defaults: readonly (readonly [string, string])[],
  overrides: readonly (readonly [string, string])[] = [],
): string {
  return [
    `<Types xmlns="${CONTENT_TYPES_NAMESPACE}">`,
    ...defaults.map(
      ([extension, contentType]) =>
        `<Default Extension="${extension}" ContentType="${contentType}"/>`,
    ),
    ...overrides.map(
      ([partName, contentType]) =>
        `<Override PartName="${partName}" ContentType="${contentType}"/>`,
    ),
    '</Types>',
  ].join('');
}

function relationships(
  items: readonly (readonly [string, string, string, string?])[],
): string {
  return [
    `<Relationships xmlns="${RELATIONSHIPS_NAMESPACE}">`,
    ...items.map(
      ([id, type, target, targetMode]) =>
        `<Relationship Id="${id}" Type="${type}" Target="${target}"${
          targetMode ? ` TargetMode="${targetMode}"` : ''
        }/>`,
    ),
    '</Relationships>',
  ].join('');
}

async function xmlEntry(archive: JSZip, path: string): Promise<Document> {
  const source = await archive.file(path)?.async('text');
  if (!source) throw new Error(`Missing OOXML part: ${path}`);
  return parseXml(source, path);
}

async function appendXmlEntry(
  archive: JSZip,
  path: string,
  fragment: string,
): Promise<void> {
  const entry = archive.file(path);
  const source = await entry?.async('text');
  if (!source) throw new Error(`Missing OOXML part: ${path}`);
  archive.file(path, source.replace(/<\/[^>]+>\s*$/, `${fragment}$&`));
}
