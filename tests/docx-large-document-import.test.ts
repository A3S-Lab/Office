import { describe, expect, test } from '@rstest/core';
import { documentInitialSectionLayout } from '../src/internal/features/work/work-document-section';
import { documentModelUsesWindowing } from '../src/internal/features/work/work-document-windowing';
import {
  parseLargeSimpleDocxDocumentXml,
  simpleDocxXmlIsEligibleForLargeImport,
} from '../src/internal/features/work/work-docx-large-document-import';

describe('large simple DOCX import', () => {
  test('builds a windowed paragraph model directly from WordprocessingML', () => {
    const result = parseLargeSimpleDocxDocumentXml(
      wordDocument(
        [
          paragraphXml('Alpha &amp; beta'),
          paragraphXml('Gamma'),
          paragraphXml('Delta'),
        ].join(''),
      ),
      {
        minimumLogicalBlocks: 3,
        windowing: {
          blockSize: 2,
          blockThreshold: 2,
          tableRowSize: 2,
          tableRowThreshold: 2,
        },
      },
    );

    expect(result).not.toBeNull();
    expect(documentModelUsesWindowing(result?.root ?? { type: 'doc' })).toBe(
      true,
    );
    expect(result?.logicalBlockCount).toBe(3);
    expect(result?.html).toContain('<p>Alpha &amp; beta</p>');
    expect(
      result?.root.content?.[0]?.content?.map(
        (chunk) => chunk.attrs?.integrityFeatures,
      ),
    ).toEqual([0, 0]);
    expect(result?.root.content?.[0]?.content?.[0]?.content?.[0]).toEqual({
      type: 'paragraph',
      content: [{ type: 'text', text: 'Alpha & beta' }],
    });
  });

  test('slices table rows in the model while retaining one canonical HTML table', () => {
    const rows = Array.from({ length: 5 }, (_, index) =>
      tableRowXml(String(index + 1), `Record ${index + 1}`),
    ).join('');
    const result = parseLargeSimpleDocxDocumentXml(
      wordDocument(`<w:tbl>${rows}</w:tbl>`),
      {
        minimumLogicalBlocks: 5,
        windowing: {
          blockSize: 2,
          blockThreshold: 100,
          tableRowSize: 2,
          tableRowThreshold: 4,
        },
      },
    );

    expect(result?.tableRowCount).toBe(5);
    expect(result?.html.match(/<table/g)).toHaveLength(1);
    const tables =
      result?.root.content?.[0]?.content?.flatMap((chunk) =>
        (chunk.content ?? []).filter((node) => node.type === 'table'),
      ) ?? [];
    expect(tables.map((table) => table.content?.length)).toEqual([2, 2, 1]);
  });

  test('rejects rich documents so the complete compatibility importer owns them', () => {
    const rich = wordDocument(
      '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Title</w:t></w:r></w:p>',
    );
    expect(simpleDocxXmlIsEligibleForLargeImport(rich, 1)).toBe(false);
    expect(
      parseLargeSimpleDocxDocumentXml(rich, { minimumLogicalBlocks: 1 }),
    ).toBeNull();
  });

  test('rejects duplicate document envelope elements', () => {
    const duplicateSection = wordDocument(
      paragraphXml('Duplicate section'),
      '<w:sectPr/><w:sectPr/>',
    );
    expect(
      parseLargeSimpleDocxDocumentXml(duplicateSection, {
        minimumLogicalBlocks: 1,
      }),
    ).toBeNull();
  });

  test('keeps the lightweight parser default layout aligned with document sections', () => {
    const result = parseLargeSimpleDocxDocumentXml(
      wordDocument(paragraphXml('Default layout'), '<w:sectPr/>'),
      { minimumLogicalBlocks: 1 },
    );

    expect(result?.layout).toEqual(
      documentInitialSectionLayout({
        type: 'document',
        html: '<p></p>',
        pageSize: 'a4',
      }),
    );
  });
});

function paragraphXml(text: string): string {
  return `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
}

function tableRowXml(...cells: string[]): string {
  return `<w:tr>${cells
    .map(
      (value) =>
        `<w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr>${paragraphXml(value)}</w:tc>`,
    )
    .join('')}</w:tr>`;
}

function wordDocument(
  body: string,
  section = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>',
): string {
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
    '<w:body>',
    body,
    section,
    '</w:body>',
    '</w:document>',
  ].join('');
}
