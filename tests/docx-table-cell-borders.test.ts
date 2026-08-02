import { describe, expect, test } from '@rstest/core';
import JSZip from 'jszip';
import { createArtifact, createArtifactBlob } from '../src/core';
import {
  applyImportedDocxTableCellMarkers,
  markDocxTableCells,
} from '../src/internal/features/work/work-docx-table-cell-import';
import { prepareDocxImport } from '../src/internal/features/work/work-docx-import';
import { parseXml } from '../src/internal/features/work/work-ooxml-package';

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

describe('DOCX table cell borders', () => {
  test('maps table outer and inside borders onto independent cell edges', () => {
    const document = parseXml(`
      <w:document xmlns:w="${WORD_NAMESPACE}">
        <w:body>
          <w:tbl>
            <w:tblPr>
              <w:tblBorders>
                <w:top w:val="double" w:sz="12" w:color="4472C4"/>
                <w:right w:val="double" w:sz="12" w:color="4472C4"/>
                <w:bottom w:val="double" w:sz="12" w:color="4472C4"/>
                <w:left w:val="double" w:sz="12" w:color="4472C4"/>
                <w:insideH w:val="dashed" w:sz="6" w:color="70AD47"/>
                <w:insideV w:val="nil"/>
              </w:tblBorders>
            </w:tblPr>
            <w:tr>
              <w:tc>
                <w:tcPr>
                  <w:tcBorders>
                    <w:right w:val="dotted" w:sz="6" w:color="C00000"/>
                  </w:tcBorders>
                </w:tcPr>
                <w:p><w:r><w:t>A</w:t></w:r></w:p>
              </w:tc>
              <w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc>
            </w:tr>
            <w:tr>
              <w:tc><w:p><w:r><w:t>C</w:t></w:r></w:p></w:tc>
              <w:tc><w:p><w:r><w:t>D</w:t></w:r></w:p></w:tc>
            </w:tr>
          </w:tbl>
        </w:body>
      </w:document>
    `);

    const markers = markDocxTableCells(document);
    expect(markers.cells).toHaveLength(4);
    expect(markers.cells[0]?.borders).toEqual({
      top: { color: '#4472c4', style: 'double', width: 2 },
      right: { color: '#c00000', style: 'dotted', width: 1 },
      bottom: { color: '#70ad47', style: 'dashed', width: 1 },
      left: { color: '#4472c4', style: 'double', width: 2 },
    });
    expect(markers.cells[3]?.borders).toEqual({
      top: { color: '#70ad47', style: 'dashed', width: 1 },
      right: { color: '#4472c4', style: 'double', width: 2 },
      bottom: { color: '#4472c4', style: 'double', width: 2 },
      left: { color: '#000000', style: 'none', width: 0 },
    });

    const html = new DOMParser().parseFromString(
      `<table><tbody>
        <tr><td>${markers.cells[0]?.marker}</td><td>${markers.cells[1]?.marker}</td></tr>
        <tr><td>${markers.cells[2]?.marker}</td><td>${markers.cells[3]?.marker}</td></tr>
      </tbody></table>`,
      'text/html',
    );
    applyImportedDocxTableCellMarkers(html, markers);

    const cells = Array.from(html.querySelectorAll<HTMLTableCellElement>('td'));
    expect(cells[0]?.dataset).toMatchObject({
      officeCellBorderTopStyle: 'double',
      officeCellBorderRightStyle: 'dotted',
      officeCellBorderBottomStyle: 'dashed',
      officeCellBorderLeftStyle: 'double',
    });
    expect(cells[3]?.dataset).toMatchObject({
      officeCellBorderTopStyle: 'dashed',
      officeCellBorderRightStyle: 'double',
      officeCellBorderBottomStyle: 'double',
      officeCellBorderLeftStyle: 'none',
    });
    expect(cells[0]?.style.borderTopStyle).toBe('double');
    expect(cells[0]?.style.borderRightStyle).toBe('dotted');
    expect(cells[0]?.style.borderBottomStyle).toBe('dashed');
    expect(cells[0]?.style.borderLeftStyle).toBe('double');
    expect(html.body.textContent).not.toContain('__A3S_');
  });

  test('resolves and exports theme table presentation as stable RGB', async () => {
    const document = parseXml(`
      <w:document xmlns:w="${WORD_NAMESPACE}">
        <w:body>
          <w:tbl>
            <w:tblPr>
              <w:tblBorders>
                <w:top w:val="single" w:sz="9" w:color="4472C4"
                  w:themeColor="accent1" w:themeTint="80"/>
                <w:right w:val="double" w:sz="12" w:color="4472C4"
                  w:themeColor="accent1" w:themeShade="BF"/>
                <w:bottom w:val="dashed" w:sz="6" w:color="ED7D31"
                  w:themeColor="accent2"/>
                <w:left w:val="dotted" w:sz="6" w:color="A5A5A5"
                  w:themeColor="accent3"/>
              </w:tblBorders>
            </w:tblPr>
            <w:tr>
              <w:tc>
                <w:tcPr>
                  <w:shd w:val="clear" w:fill="ED7D31"
                    w:themeFill="accent2" w:themeFillTint="99"/>
                </w:tcPr>
                <w:p><w:r><w:t>Theme cell</w:t></w:r></w:p>
              </w:tc>
            </w:tr>
          </w:tbl>
        </w:body>
      </w:document>
    `);
    const theme = parseXml(`
      <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <a:themeElements>
          <a:clrScheme name="Office">
            <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
            <a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
            <a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>
          </a:clrScheme>
        </a:themeElements>
      </a:theme>
    `);

    const markers = markDocxTableCells(document, theme);

    expect(markers.cells).toEqual([
      {
        marker: '__A3S_WORK_TABLE_CELL_1__',
        backgroundColor: '#f8cbad',
        borders: {
          top: { color: '#a2b9e2', style: 'solid', width: 1.5 },
          right: { color: '#335593', style: 'double', width: 2 },
          bottom: { color: '#ed7d31', style: 'dashed', width: 1 },
          left: { color: '#a5a5a5', style: 'dotted', width: 1 },
        },
      },
    ]);

    const html = new DOMParser().parseFromString(
      `<table><tbody><tr><td>${markers.cells[0]?.marker}Theme cell</td></tr></tbody></table>`,
      'text/html',
    );
    applyImportedDocxTableCellMarkers(html, markers);
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    artifact.content.html = html.body.innerHTML;
    const blob = await createArtifactBlob(artifact);
    const exported = await JSZip.loadAsync(await blob.arrayBuffer());
    const exportedXml =
      (await exported.file('word/document.xml')?.async('string')) ?? '';

    expect(exportedXml).toMatch(/<w:shd\b[^>]*w:fill="F8CBAD"/);
    expect(exportedXml).toMatch(/<w:top\b[^>]*w:color="A2B9E2"/);
    expect(exportedXml).toMatch(/<w:right\b[^>]*w:color="335593"/);
    expect(exportedXml).toMatch(/<w:bottom\b[^>]*w:color="ED7D31"/);
    expect(exportedXml).toMatch(/<w:left\b[^>]*w:color="A5A5A5"/);
  });

  test('loads the package theme before marking table presentation', async () => {
    const archive = new JSZip();
    archive.file(
      'word/document.xml',
      `
        <w:document xmlns:w="${WORD_NAMESPACE}">
          <w:body>
            <w:tbl>
              <w:tblPr>
                <w:tblBorders>
                  <w:top w:val="single" w:sz="9" w:color="000000"
                    w:themeColor="accent1" w:themeTint="80"/>
                </w:tblBorders>
              </w:tblPr>
              <w:tr>
                <w:tc>
                  <w:tcPr>
                    <w:shd w:val="clear" w:fill="000000"
                      w:themeFill="accent2" w:themeFillTint="99"/>
                  </w:tcPr>
                  <w:p><w:r><w:t>Theme cell</w:t></w:r></w:p>
                </w:tc>
              </w:tr>
            </w:tbl>
            <w:sectPr/>
          </w:body>
        </w:document>
      `,
    );
    archive.file(
      'word/theme/theme1.xml',
      `
        <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <a:themeElements>
            <a:clrScheme name="Office">
              <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
              <a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
            </a:clrScheme>
          </a:themeElements>
        </a:theme>
      `,
    );

    const prepared = await prepareDocxImport(
      await archive.generateAsync({ type: 'arraybuffer' }),
    );

    expect(prepared.tableCellMarkers.cells[0]).toMatchObject({
      backgroundColor: '#f8cbad',
      borders: {
        top: { color: '#a2b9e2', style: 'solid', width: 1.5 },
      },
    });
  });
});
