import { describe, expect, test } from '@rstest/core';
import {
  applyImportedDocxTableCellMarkers,
  markDocxTableCells,
} from '../src/internal/features/work/work-docx-table-cell-import';
import {
  applyImportedDocxTableSizingMarkers,
  markDocxTableSizing,
} from '../src/internal/features/work/work-docx-table-sizing-import';
import { parseXml } from '../src/internal/features/work/work-ooxml-package';

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

describe('DOCX table geometry', () => {
  test('keeps layout, percentage width, alignment, indent, margins, and grid widths independent', () => {
    const document = wordDocumentXml(`
      <w:tbl>
        <w:tblPr>
          <w:tblW w:type="pct" w:w="2500"/>
          <w:jc w:val="center"/>
          <w:tblInd w:type="dxa" w:w="120"/>
          <w:tblLayout w:type="autofit"/>
          <w:tblCellMar>
            <w:top w:type="dxa" w:w="0"/>
            <w:start w:type="dxa" w:w="144"/>
            <w:bottom w:type="dxa" w:w="72"/>
            <w:end w:type="dxa" w:w="216"/>
          </w:tblCellMar>
        </w:tblPr>
        <w:tblGrid>
          <w:gridCol w:w="1200"/>
          <w:gridCol w:w="1800"/>
        </w:tblGrid>
        <w:tr>
          <w:tc><w:p><w:r><w:t>Left</w:t></w:r></w:p></w:tc>
          <w:tc><w:p><w:r><w:t>Right</w:t></w:r></w:p></w:tc>
        </w:tr>
      </w:tbl>
    `);

    const markers = markDocxTableSizing(document);

    expect(markers.tables).toEqual([
      {
        marker: '__A3S_WORK_TABLE_SIZING_1__',
        geometry: {
          layout: 'autofit',
          width: { type: 'percent', value: 50 },
          alignment: 'center',
          indent: 8,
          cellMargins: {
            top: 0,
            right: 14.4,
            bottom: 4.8,
            left: 9.6,
          },
        },
        columnWidths: [80, 120],
      },
    ]);

    const html = new DOMParser().parseFromString(
      `<table><tbody><tr><td><p>${markers.tables[0]?.marker}Left</p></td><td><p>Right</p></td></tr></tbody></table>`,
      'text/html',
    );
    applyImportedDocxTableSizingMarkers(html, markers);

    const table = html.querySelector('table');
    expect(table?.dataset).toMatchObject({
      officeTableImported: 'true',
      officeTableLayout: 'autofit',
      officeTableWidthType: 'percent',
      officeTableWidth: '50',
      officeTableAlignment: 'center',
      officeTableIndent: '8',
      officeTableCellMarginTop: '0',
      officeTableCellMarginRight: '14.4',
      officeTableCellMarginBottom: '4.8',
      officeTableCellMarginLeft: '9.6',
    });
    expect(table?.style.width).toBe('50%');
    expect(table?.style.tableLayout).toBe('auto');
    expect(table?.style.marginLeft).toBe('auto');
    expect(table?.style.marginRight).toBe('auto');
    expect(table?.style.getPropertyValue('--work-table-cell-margin-left')).toBe(
      '9.6px',
    );
    expect(
      Array.from(html.querySelectorAll('td')).map((cell) =>
        cell.getAttribute('colwidth'),
      ),
    ).toEqual(['80', '120']);
    expect(html.body.textContent).not.toContain('__A3S_');
  });

  test('resolves inherited table defaults before direct table geometry', () => {
    const document = wordDocumentXml(`
      <w:tbl>
        <w:tblPr>
          <w:tblStyle w:val="CompactTable"/>
          <w:tblW w:type="pct" w:w="75%"/>
          <w:jc w:val="right"/>
        </w:tblPr>
        <w:tr><w:tc><w:p><w:r><w:t>Styled</w:t></w:r></w:p></w:tc></w:tr>
      </w:tbl>
    `);
    const styles = parseXml(`
      <w:styles xmlns:w="${WORD_NAMESPACE}">
        <w:style w:type="table" w:styleId="CompactTable">
          <w:tblPr>
            <w:tblLayout w:type="fixed"/>
            <w:tblW w:type="dxa" w:w="3600"/>
            <w:jc w:val="center"/>
            <w:tblCellMar>
              <w:top w:type="dxa" w:w="60"/>
              <w:left w:type="dxa" w:w="90"/>
              <w:bottom w:type="dxa" w:w="60"/>
              <w:right w:type="dxa" w:w="90"/>
            </w:tblCellMar>
          </w:tblPr>
        </w:style>
      </w:styles>
    `);

    const markers = markDocxTableSizing(document, styles);

    expect(markers.tables[0]?.geometry).toEqual({
      layout: 'fixed',
      width: { type: 'percent', value: 75 },
      alignment: 'right',
      indent: 0,
      cellMargins: { top: 4, right: 6, bottom: 4, left: 6 },
    });
  });

  test('keeps cell-level margins as partial overrides of table defaults', () => {
    const document = wordDocumentXml(`
      <w:tbl>
        <w:tr>
          <w:tc>
            <w:tcPr>
              <w:tcMar>
                <w:top w:type="dxa" w:w="120"/>
                <w:end w:type="dxa" w:w="240"/>
              </w:tcMar>
            </w:tcPr>
            <w:p><w:r><w:t>Override</w:t></w:r></w:p>
          </w:tc>
        </w:tr>
      </w:tbl>
    `);

    const markers = markDocxTableCells(document);
    expect(markers.cells[0]).toMatchObject({
      margins: { top: 8, right: 16 },
    });

    const html = new DOMParser().parseFromString(
      `<table><tbody><tr><td><p>${markers.cells[0]?.marker}Override</p></td></tr></tbody></table>`,
      'text/html',
    );
    applyImportedDocxTableCellMarkers(html, markers);

    const cell = html.querySelector('td');
    expect(cell?.dataset).toMatchObject({
      officeCellMarginTop: '8',
      officeCellMarginRight: '16',
    });
    expect(cell?.style.paddingTop).toBe('8px');
    expect(cell?.style.paddingRight).toBe('16px');
    expect(cell?.style.paddingBottom).toBe('');
    expect(cell?.style.paddingLeft).toBe('');
  });
});

function wordDocumentXml(body: string): Document {
  return parseXml(
    `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body>${body}</w:body></w:document>`,
  );
}
