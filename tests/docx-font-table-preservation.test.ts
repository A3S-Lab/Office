import { describe, expect, test } from '@rstest/core';
import JSZip from 'jszip';
import { createDocxBlob } from '../src/internal/features/work/work-docx-export';
import { analyzeDocxCompatibility } from '../src/internal/features/work/work-office-diagnostics';
import {
  attribute,
  directChildren,
  parseXml,
} from '../src/internal/features/work/work-ooxml-package';
import { preserveDocxSourcePackage } from '../src/internal/features/work/work-ooxml-package-preservation';

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const STRICT_WORD_NAMESPACE =
  'http://purl.oclc.org/ooxml/wordprocessingml/main';
const STRICT_OFFICE_RELATIONSHIPS_NAMESPACE =
  'http://purl.oclc.org/ooxml/officeDocument/relationships';
const OFFICE_RELATIONSHIPS_NAMESPACE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PACKAGE_RELATIONSHIPS_NAMESPACE =
  'http://schemas.openxmlformats.org/package/2006/relationships';
const FONT_RELATIONSHIP = `${OFFICE_RELATIONSHIPS_NAMESPACE}/font`;
const IMAGE_RELATIONSHIP = `${OFFICE_RELATIONSHIPS_NAMESPACE}/image`;

describe('DOCX font-table preservation', () => {
  test('rewrites embedded-font references after relationship ID collisions', async () => {
    const base = await generatedDocx();
    const generated = await JSZip.loadAsync(base);
    const generatedFonts = await xmlEntry(generated, 'word/fontTable.xml');
    appendFont(generatedFonts, 'Generated Font', [
      ['family', [['val', 'swiss']]],
      [
        'embedRegular',
        [
          ['id', 'rId1', 'relationship'],
          ['fontKey', '{00000000-0000-0000-0000-000000000001}'],
        ],
      ],
    ]);
    generated.file(
      'word/fontTable.xml',
      new XMLSerializer().serializeToString(generatedFonts),
    );
    generated.file('word/fonts/generated.odttf', new Uint8Array([9, 9, 9, 9]));
    generated.file(
      'word/_rels/fontTable.xml.rels',
      relationships([['rId1', FONT_RELATIONSHIP, 'fonts/generated.odttf']]),
    );
    const generatedBytes = await generated.generateAsync({
      type: 'arraybuffer',
    });

    const source = await JSZip.loadAsync(base);
    const sourceFonts = await xmlEntry(source, 'word/fontTable.xml');
    appendFont(sourceFonts, 'A3S Embedded', [
      ['altName', [['val', 'A3S Embedded Regular']]],
      ['family', [['val', 'roman']]],
      [
        'embedRegular',
        [
          ['id', 'rId1', 'relationship'],
          ['fontKey', '{AAAAAAAA-AAAA-AAAA-AAAA-AAAAAAAAAAAA}'],
          ['subsetted', '1'],
        ],
      ],
      [
        'embedBold',
        [
          ['id', 'rId2', 'relationship'],
          ['fontKey', '{BBBBBBBB-BBBB-BBBB-BBBB-BBBBBBBBBBBB}'],
        ],
      ],
    ]);
    appendFont(sourceFonts, 'Metadata Only', [
      ['charset', [['val', '86']]],
      ['pitch', [['val', 'variable']]],
    ]);
    appendFont(sourceFonts, 'External Font', [
      ['family', [['val', 'swiss']]],
      [
        'embedRegular',
        [
          ['id', 'rIdExternal', 'relationship'],
          ['fontKey', '{CCCCCCCC-CCCC-CCCC-CCCC-CCCCCCCCCCCC}'],
        ],
      ],
    ]);
    appendFont(sourceFonts, 'Wrong Relationship Type', [
      [
        'embedRegular',
        [
          ['id', 'rIdImage', 'relationship'],
          ['fontKey', '{DDDDDDDD-DDDD-DDDD-DDDD-DDDDDDDDDDDD}'],
        ],
      ],
    ]);
    appendFont(sourceFonts, 'Wrong Content Type', [
      [
        'embedRegular',
        [
          ['id', 'rIdWrongContent', 'relationship'],
          ['fontKey', '{EEEEEEEE-EEEE-EEEE-EEEE-EEEEEEEEEEEE}'],
        ],
      ],
    ]);
    appendFont(sourceFonts, 'Colliding Payload', [
      [
        'embedRegular',
        [
          ['id', 'rIdCollision', 'relationship'],
          ['fontKey', '{FFFFFFFF-FFFF-FFFF-FFFF-FFFFFFFFFFFF}'],
        ],
      ],
    ]);
    source.file(
      'word/fontTable.xml',
      new XMLSerializer().serializeToString(sourceFonts),
    );
    const regularBytes = new Uint8Array([1, 2, 3, 4, 5]);
    const boldBytes = new Uint8Array([6, 7, 8, 9]);
    source.file('word/fonts/a3s-regular.odttf', regularBytes);
    source.file('word/fonts/a3s-bold.odttf', boldBytes);
    source.file('word/fonts/not-a-font.odttf', new Uint8Array([4, 3, 2, 1]));
    source.file('word/fonts/not-a-font.png', new Uint8Array([137, 80, 78, 71]));
    source.file('word/fonts/generated.odttf', new Uint8Array([1, 1, 1, 1]));
    source.file(
      'word/_rels/fontTable.xml.rels',
      relationships([
        ['rId1', FONT_RELATIONSHIP, 'fonts/a3s-regular.odttf'],
        ['rId2', FONT_RELATIONSHIP, 'fonts/a3s-bold.odttf', 'Internal'],
        [
          'rIdExternal',
          FONT_RELATIONSHIP,
          'https://fonts.a3s.dev/remote.odttf',
          'External',
        ],
        ['rIdImage', IMAGE_RELATIONSHIP, 'fonts/not-a-font.odttf'],
        ['rIdWrongContent', FONT_RELATIONSHIP, 'fonts/not-a-font.png'],
        ['rIdCollision', FONT_RELATIONSHIP, 'fonts/generated.odttf'],
      ]),
    );

    const sourceBytes = await source.generateAsync({ type: 'arraybuffer' });
    const compatibility = await analyzeDocxCompatibility(
      new File([sourceBytes], 'embedded-fonts.docx', {
        type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      }),
      [],
    );
    expect(compatibility.issues).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          code: 'docx.embedded-fonts',
          severity: 'warning',
        }),
      ]),
    );

    const output = await preserveDocxSourcePackage(generatedBytes, sourceBytes);
    const archive = await JSZip.loadAsync(output);
    await expect(
      archive.file('word/fonts/a3s-regular.odttf')?.async('uint8array'),
    ).resolves.toEqual(regularBytes);
    await expect(
      archive.file('word/fonts/a3s-bold.odttf')?.async('uint8array'),
    ).resolves.toEqual(boldBytes);
    await expect(
      archive.file('word/fonts/generated.odttf')?.async('uint8array'),
    ).resolves.toEqual(new Uint8Array([9, 9, 9, 9]));

    const relationshipDocument = await xmlEntry(
      archive,
      'word/_rels/fontTable.xml.rels',
    );
    expect(
      directChildren(relationshipDocument.documentElement, 'Relationship').map(
        (item) => [
          attribute(item, 'Id'),
          attribute(item, 'Type')?.split('/').pop(),
          attribute(item, 'Target'),
        ],
      ),
    ).toEqual([
      ['rId1', 'font', 'fonts/generated.odttf'],
      ['rId2', 'font', 'fonts/a3s-regular.odttf'],
      ['rId3', 'font', 'fonts/a3s-bold.odttf'],
      ['rIdImage', 'image', 'fonts/not-a-font.odttf'],
      ['rIdWrongContent', 'font', 'fonts/not-a-font.png'],
    ]);

    const fontTable = await xmlEntry(archive, 'word/fontTable.xml');
    expect(fontNames(fontTable)).toEqual([
      'Generated Font',
      'A3S Embedded',
      'Metadata Only',
      'External Font',
      'Wrong Relationship Type',
      'Wrong Content Type',
      'Colliding Payload',
    ]);
    expect(fontEmbeddings(fontTable, 'Generated Font')).toEqual([
      ['embedRegular', 'rId1'],
    ]);
    expect(fontEmbeddings(fontTable, 'A3S Embedded')).toEqual([
      ['embedRegular', 'rId2'],
      ['embedBold', 'rId3'],
    ]);
    expect(fontChildNames(fontTable, 'Metadata Only')).toEqual([
      'charset',
      'pitch',
    ]);
    expect(fontEmbeddings(fontTable, 'External Font')).toEqual([]);
    expect(fontEmbeddings(fontTable, 'Wrong Relationship Type')).toEqual([]);
    expect(fontEmbeddings(fontTable, 'Wrong Content Type')).toEqual([]);
    expect(fontEmbeddings(fontTable, 'Colliding Payload')).toEqual([]);

    const secondOutput = await preserveDocxSourcePackage(
      generatedBytes,
      output,
    );
    const secondArchive = await JSZip.loadAsync(secondOutput);
    const secondFontTable = await xmlEntry(secondArchive, 'word/fontTable.xml');
    expect(fontNames(secondFontTable)).toEqual(fontNames(fontTable));
    expect(fontEmbeddings(secondFontTable, 'A3S Embedded')).toEqual([
      ['embedRegular', 'rId2'],
      ['embedBold', 'rId3'],
    ]);
  });

  test('rejects ambiguous font names and relationship IDs', async () => {
    const generated = await generatedDocx();
    const duplicateFonts = await JSZip.loadAsync(generated);
    const fontTable = await xmlEntry(duplicateFonts, 'word/fontTable.xml');
    appendFont(fontTable, 'Duplicate Font', []);
    appendFont(fontTable, 'duplicate font', []);
    duplicateFonts.file(
      'word/fontTable.xml',
      new XMLSerializer().serializeToString(fontTable),
    );
    await expect(
      preserveDocxSourcePackage(
        generated,
        await duplicateFonts.generateAsync({ type: 'arraybuffer' }),
      ),
    ).rejects.toThrow('duplicate font names');

    const duplicateRelationships = await JSZip.loadAsync(generated);
    duplicateRelationships.file('word/fonts/first.odttf', new Uint8Array([1]));
    duplicateRelationships.file('word/fonts/second.odttf', new Uint8Array([2]));
    duplicateRelationships.file(
      'word/_rels/fontTable.xml.rels',
      relationships([
        ['rIdFont', FONT_RELATIONSHIP, 'fonts/first.odttf'],
        ['rIdFont', FONT_RELATIONSHIP, 'fonts/second.odttf'],
      ]),
    );
    await expect(
      preserveDocxSourcePackage(
        generated,
        await duplicateRelationships.generateAsync({ type: 'arraybuffer' }),
      ),
    ).rejects.toThrow('duplicate relationship IDs');
  });

  test('normalizes strict UTF-16 font metadata into the generated dialect', async () => {
    const generated = await generatedDocx();
    const source = await JSZip.loadAsync(generated);
    const sourceText = await source.file('word/fontTable.xml')?.async('text');
    if (!sourceText) throw new Error('Generated font table is missing.');
    const strictDocument = parseXml(
      sourceText.replaceAll(WORD_NAMESPACE, STRICT_WORD_NAMESPACE),
      'strict word/fontTable.xml',
    );
    const font = strictDocument.createElementNS(
      STRICT_WORD_NAMESPACE,
      'w:font',
    );
    font.setAttributeNS(STRICT_WORD_NAMESPACE, 'w:name', 'Strict Font');
    const family = strictDocument.createElementNS(
      STRICT_WORD_NAMESPACE,
      'w:family',
    );
    family.setAttributeNS(STRICT_WORD_NAMESPACE, 'w:val', 'modern');
    font.append(family);
    const embedding = strictDocument.createElementNS(
      STRICT_WORD_NAMESPACE,
      'w:embedRegular',
    );
    embedding.setAttributeNS(
      STRICT_OFFICE_RELATIONSHIPS_NAMESPACE,
      'r:id',
      'rIdStrictFont',
    );
    embedding.setAttributeNS(
      STRICT_WORD_NAMESPACE,
      'w:fontKey',
      '{11111111-2222-3333-4444-555555555555}',
    );
    font.append(embedding);
    strictDocument.documentElement.append(font);
    const strictXml = new XMLSerializer()
      .serializeToString(strictDocument)
      .replace(
        /^\s*<\?xml[^?]*\?>/i,
        '<?xml version="1.0" encoding="UTF-16" standalone="yes"?>',
      );
    source.file('word/fontTable.xml', utf16LittleEndian(strictXml));
    const strictFontBytes = new Uint8Array([10, 20, 30, 40]);
    source.file('word/fonts/strict.odttf', strictFontBytes);
    source.file(
      'word/_rels/fontTable.xml.rels',
      utf16LittleEndian(
        `<?xml version="1.0" encoding="UTF-16" standalone="yes"?>${relationships(
          [
            [
              'rIdStrictFont',
              `${STRICT_OFFICE_RELATIONSHIPS_NAMESPACE}/font`,
              'fonts/strict.odttf',
            ],
            [
              'rIdExternal',
              `${STRICT_OFFICE_RELATIONSHIPS_NAMESPACE}/font`,
              'https://fonts.a3s.dev/strict.odttf',
              'External',
            ],
          ],
        )}`,
      ),
    );
    const contentTypes = await source
      .file('[Content_Types].xml')
      ?.async('text');
    if (!contentTypes) throw new Error('Generated content types are missing.');
    source.file(
      '[Content_Types].xml',
      utf16LittleEndian(
        contentTypes.replace(
          /^\s*<\?xml[^?]*\?>/i,
          '<?xml version="1.0" encoding="UTF-16" standalone="yes"?>',
        ),
      ),
    );

    const sourceBytes = await source.generateAsync({ type: 'arraybuffer' });
    const compatibility = await analyzeDocxCompatibility(
      new File([sourceBytes], 'strict-utf16-font.docx'),
      [],
    );
    expect(compatibility.issues.map((item) => item.code)).toContain(
      'docx.embedded-fonts',
    );

    const output = await preserveDocxSourcePackage(generated, sourceBytes);
    const archive = await JSZip.loadAsync(output);
    const outputText = await archive.file('word/fontTable.xml')?.async('text');
    expect(outputText).toContain('encoding="UTF-8"');
    const outputFonts = await xmlEntry(archive, 'word/fontTable.xml');
    expect(outputFonts.documentElement.namespaceURI).toBe(WORD_NAMESPACE);
    expect(fontNames(outputFonts)).toEqual(['Strict Font']);
    expect(fontChildNames(outputFonts, 'Strict Font')).toEqual([
      'family',
      'embedRegular',
    ]);
    expect(fontEmbeddings(outputFonts, 'Strict Font')).toEqual([
      ['embedRegular', 'rIdStrictFont'],
    ]);
    expect(
      attribute(
        directChildren(fontByName(outputFonts, 'Strict Font'), 'family')[0],
        'val',
      ),
    ).toBe('modern');
    await expect(
      archive.file('word/fonts/strict.odttf')?.async('uint8array'),
    ).resolves.toEqual(strictFontBytes);
    await expect(
      archive.file('word/_rels/fontTable.xml.rels')?.async('text'),
    ).resolves.toContain('encoding="UTF-8"');
    await expect(
      archive.file('word/_rels/fontTable.xml.rels')?.async('text'),
    ).resolves.not.toContain('fonts.a3s.dev');
  });
});

type FontChild = readonly [
  string,
  readonly (readonly [string, string, ('relationship' | 'word')?])[],
];

async function generatedDocx(): Promise<ArrayBuffer> {
  return (
    await createDocxBlob({
      type: 'document',
      html: '<p>Embedded font relationship boundary</p>',
      pageSize: 'a4',
    })
  ).arrayBuffer();
}

function appendFont(
  document: Document,
  name: string,
  children: readonly FontChild[],
): void {
  const font = document.createElementNS(WORD_NAMESPACE, 'w:font');
  font.setAttributeNS(WORD_NAMESPACE, 'w:name', name);
  for (const [localName, attributes] of children) {
    const child = document.createElementNS(WORD_NAMESPACE, `w:${localName}`);
    for (const [attributeName, value, kind = 'word'] of attributes) {
      child.setAttributeNS(
        kind === 'relationship'
          ? OFFICE_RELATIONSHIPS_NAMESPACE
          : WORD_NAMESPACE,
        `${kind === 'relationship' ? 'r' : 'w'}:${attributeName}`,
        value,
      );
    }
    font.append(child);
  }
  document.documentElement.append(font);
}

function relationships(
  items: readonly (readonly [string, string, string, string?])[],
): string {
  return [
    `<Relationships xmlns="${PACKAGE_RELATIONSHIPS_NAMESPACE}">`,
    ...items.map(
      ([id, type, target, targetMode]) =>
        `<Relationship Id="${id}" Type="${type}" Target="${target}"${
          targetMode ? ` TargetMode="${targetMode}"` : ''
        }/>`,
    ),
    '</Relationships>',
  ].join('');
}

function fontNames(document: Document): string[] {
  return directChildren(document.documentElement, 'font').flatMap((font) => {
    const name = attribute(font, 'name');
    return name ? [name] : [];
  });
}

function fontEmbeddings(
  document: Document,
  name: string,
): Array<[string, string | null]> {
  const font = fontByName(document, name);
  return directChildren(font)
    .filter((child) => child.localName.startsWith('embed'))
    .map((child) => [child.localName, attribute(child, 'id')]);
}

function fontChildNames(document: Document, name: string): string[] {
  return directChildren(fontByName(document, name)).map(
    (child) => child.localName,
  );
}

function fontByName(document: Document, name: string): Element {
  const font = directChildren(document.documentElement, 'font').find(
    (item) => attribute(item, 'name') === name,
  );
  if (!font) throw new Error(`Missing font record: ${name}`);
  return font;
}

async function xmlEntry(archive: JSZip, path: string): Promise<Document> {
  const source = await archive.file(path)?.async('text');
  if (!source) throw new Error(`Missing OOXML part: ${path}`);
  return parseXml(source, path);
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
