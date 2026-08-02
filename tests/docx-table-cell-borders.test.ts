import { describe, expect, test } from '@rstest/core';
import {
  applyImportedDocxTableCellMarkers,
  markDocxTableCells,
} from '../src/internal/features/work/work-docx-table-cell-import';
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
});
