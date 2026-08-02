import { describe, expect, test } from '@rstest/core';
import JSZip from 'jszip';
import { createArtifact, createArtifactBlob } from '../src/core';
import { prepareDocxImport } from '../src/internal/features/work/work-docx-import';
import {
  applyImportedDocxRunFormattingMarkers,
  markDocxRunFormatting,
} from '../src/internal/features/work/work-docx-run-formatting-import';
import {
  applyImportedDocxTableCellMarkers,
  markDocxTableCells,
} from '../src/internal/features/work/work-docx-table-cell-import';
import { parseXml } from '../src/internal/features/work/work-ooxml-package';

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const DRAWING_NAMESPACE =
  'http://schemas.openxmlformats.org/drawingml/2006/main';

describe('DOCX table style import', () => {
  test('resolves inherited whole-table and conditional cell presentation', () => {
    const document = wordDocumentXml(`
      <w:tbl>
        <w:tblPr>
          <w:tblStyle w:val="ReportTable"/>
          <w:tblLook w:firstRow="1" w:lastRow="1" w:noHBand="0" w:noVBand="1"/>
        </w:tblPr>
        <w:tr>
          <w:tc><w:p><w:r><w:t>Header A</w:t></w:r></w:p></w:tc>
          <w:tc><w:p><w:r><w:t>Header B</w:t></w:r></w:p></w:tc>
        </w:tr>
        <w:tr>
          <w:tc><w:p><w:r><w:t>Body A</w:t></w:r></w:p></w:tc>
          <w:tc>
            <w:tcPr>
              <w:shd w:val="clear" w:themeFill="accent4"/>
              <w:tcBorders>
                <w:top w:val="dotted" w:sz="9" w:themeColor="accent2"/>
              </w:tcBorders>
            </w:tcPr>
            <w:p><w:r><w:t>Direct override</w:t></w:r></w:p>
          </w:tc>
        </w:tr>
        <w:tr>
          <w:tc><w:p><w:r><w:t>Total A</w:t></w:r></w:p></w:tc>
          <w:tc><w:p><w:r><w:t>Total B</w:t></w:r></w:p></w:tc>
        </w:tr>
      </w:tbl>
    `);
    const styles = tableStylesXml();
    const theme = officeThemeXml();

    const markers = markDocxTableCells(document, theme, styles);

    expect(markers.cells).toHaveLength(6);
    expect(markers.cells[0]).toMatchObject({
      backgroundColor: '#4472c4',
      borders: {
        top: { color: '#4472c4', style: 'solid', width: 1 },
        bottom: { color: '#ffffff', style: 'double', width: 2 },
        left: { color: '#4472c4', style: 'solid', width: 1 },
      },
    });
    expect(markers.cells[2]).toMatchObject({
      backgroundColor: '#f8cbad',
      borders: {
        top: { color: '#a5a5a5', style: 'dashed', width: 1 },
      },
    });
    expect(markers.cells[3]).toMatchObject({
      backgroundColor: '#ffc000',
      borders: {
        top: { color: '#ed7d31', style: 'dotted', width: 1.5 },
      },
    });
    expect(markers.cells[4]).toMatchObject({
      backgroundColor: '#a5a5a5',
      borders: {
        bottom: { color: '#4472c4', style: 'solid', width: 1 },
      },
    });

    const html = new DOMParser().parseFromString(
      `<table><tbody>
        <tr><td>${markers.cells[0]?.marker}Header A</td><td>${markers.cells[1]?.marker}Header B</td></tr>
        <tr><td>${markers.cells[2]?.marker}Body A</td><td>${markers.cells[3]?.marker}Direct override</td></tr>
        <tr><td>${markers.cells[4]?.marker}Total A</td><td>${markers.cells[5]?.marker}Total B</td></tr>
      </tbody></table>`,
      'text/html',
    );
    applyImportedDocxTableCellMarkers(html, markers);

    const cells = html.querySelectorAll<HTMLTableCellElement>('td');
    expect(cells[0]?.dataset).toMatchObject({
      officeCellFill: '#4472c4',
      officeCellBorderBottomColor: '#ffffff',
      officeCellBorderBottomStyle: 'double',
    });
    expect(cells[2]?.dataset.officeCellFill).toBe('#f8cbad');
    expect(cells[3]?.dataset.officeCellFill).toBe('#ffc000');
    expect(cells[4]?.dataset.officeCellFill).toBe('#a5a5a5');
  });

  test('honors the default style, encoded table look, band sizes, and corner precedence', () => {
    const document = wordDocumentXml(`
      <w:tbl>
        <w:tblPr><w:tblLook w:val="01E0"/></w:tblPr>
        ${tableRowXml('row-1')}
        ${tableRowXml('row-2')}
        ${tableRowXml('row-3')}
        ${tableRowXml('row-4')}
        ${tableRowXml('row-5')}
        ${tableRowXml('row-6')}
      </w:tbl>
    `);

    const markers = markDocxTableCells(
      document,
      undefined,
      conditionalCoverageStylesXml(),
    );

    expect(markers.cells).toHaveLength(24);
    expect(markers.cells[0]?.backgroundColor).toBe('#888888');
    expect(markers.cells[3]?.backgroundColor).toBe('#999999');
    expect(markers.cells[20]?.backgroundColor).toBe('#aaaaaa');
    expect(markers.cells[23]?.backgroundColor).toBe('#bbbbbb');
    expect(markers.cells[1]?.backgroundColor).toBe('#444444');
    expect(markers.cells[21]?.backgroundColor).toBe('#555555');
    expect(markers.cells[4]?.backgroundColor).toBe('#666666');
    expect(markers.cells[7]?.backgroundColor).toBe('#777777');
    expect(markers.cells[5]).toMatchObject({
      backgroundColor: '#222222',
      borders: {
        top: { color: '#cc0000', style: 'solid', width: 1 },
      },
    });
    expect(markers.cells[6]).toMatchObject({
      backgroundColor: '#222222',
      borders: {
        top: { color: '#0000cc', style: 'solid', width: 1 },
      },
    });
    expect(markers.cells[9]?.backgroundColor).toBe('#222222');
    expect(markers.cells[13]?.backgroundColor).toBe('#333333');
  });

  test('uses grid offsets and spans when resolving first and last columns', () => {
    const document = wordDocumentXml(`
      <w:tbl>
        <w:tblPr>
          <w:tblLook w:firstRow="0" w:lastRow="0" w:firstColumn="1" w:lastColumn="1" w:noHBand="1" w:noVBand="1"/>
        </w:tblPr>
        <w:tr>
          <w:trPr><w:gridBefore w:val="1"/><w:gridAfter w:val="1"/></w:trPr>
          <w:tc>
            <w:tcPr><w:gridSpan w:val="2"/></w:tcPr>
            <w:p><w:r><w:t>Middle span</w:t></w:r></w:p>
          </w:tc>
        </w:tr>
      </w:tbl>
    `);

    const markers = markDocxTableCells(
      document,
      undefined,
      conditionalCoverageStylesXml(),
    );

    expect(markers.cells).toHaveLength(1);
    expect(markers.cells[0]?.backgroundColor).toBe('#111111');
  });

  test('applies conditional header run formatting without changing body text', () => {
    const document = wordDocumentXml(`
      <w:tbl>
        <w:tblPr>
          <w:tblStyle w:val="ReportTable"/>
          <w:tblLook w:firstRow="1" w:lastRow="1" w:noHBand="0" w:noVBand="1"/>
        </w:tblPr>
        <w:tr>
          <w:tc><w:p><w:r><w:t>Styled header</w:t></w:r></w:p></w:tc>
        </w:tr>
        <w:tr>
          <w:tc><w:p><w:r><w:t>Ordinary body</w:t></w:r></w:p></w:tc>
        </w:tr>
      </w:tbl>
    `);
    const styles = tableStylesXml();

    const markers = markDocxRunFormatting(
      document,
      styles,
      officeThemeXml(),
      styles,
    );

    expect(markers.runs).toHaveLength(2);
    expect(markers.runs[0]?.formatting).toMatchObject({
      bold: true,
      italic: true,
      strike: true,
      underline: true,
      color: '#ffffff',
    });
    expect(markers.runs[1]?.formatting).toMatchObject({
      color: '#000000',
    });
    expect(markers.runs[1]?.formatting.bold).toBeUndefined();

    const html = new DOMParser().parseFromString(
      `<table><tbody><tr><td><p>${markers.runs[0]?.startMarker}Styled header${markers.runs[0]?.endMarker}</p></td></tr><tr><td><p>${markers.runs[1]?.startMarker}Ordinary body${markers.runs[1]?.endMarker}</p></td></tr></tbody></table>`,
      'text/html',
    );
    applyImportedDocxRunFormattingMarkers(html, markers);

    const paragraphs = html.querySelectorAll('p');
    expect(paragraphs[0]?.querySelector('strong')).toHaveTextContent(
      'Styled header',
    );
    expect(paragraphs[0]?.querySelector('em')).toHaveTextContent(
      'Styled header',
    );
    expect(paragraphs[0]?.querySelector('u')).toHaveTextContent(
      'Styled header',
    );
    expect(paragraphs[0]?.querySelector('s')).toHaveTextContent(
      'Styled header',
    );
    expect(paragraphs[0]?.querySelector('span')?.style.color).toBe('#ffffff');
    expect(paragraphs[1]?.querySelector('strong')).toBeNull();
    expect(paragraphs[1]?.textContent).toBe('Ordinary body');
  });

  test('keeps direct run formatting above conditional table formatting', () => {
    const document = wordDocumentXml(`
      <w:tbl>
        <w:tblPr>
          <w:tblStyle w:val="ReportTable"/>
          <w:tblLook w:firstRow="1"/>
        </w:tblPr>
        <w:tr>
          <w:tc>
            <w:p>
              <w:r><w:t>Inherited header</w:t></w:r>
              <w:r>
                <w:rPr>
                  <w:b w:val="0"/>
                  <w:i w:val="0"/>
                  <w:u w:val="none"/>
                  <w:strike w:val="0"/>
                  <w:color w:val="123456"/>
                </w:rPr>
                <w:t>Direct override</w:t>
              </w:r>
            </w:p>
          </w:tc>
        </w:tr>
      </w:tbl>
    `);

    const markers = markDocxRunFormatting(
      document,
      tableStylesXml(),
      officeThemeXml(),
      tableStylesXml(),
    );

    expect(markers.runs).toHaveLength(2);
    expect(markers.runs[0]?.formatting.bold).toBe(true);
    expect(markers.runs[1]?.formatting).toMatchObject({
      bold: false,
      color: '#123456',
      italic: false,
      strike: false,
      underline: false,
    });
  });

  test('materializes inherited style presentation into stable DOCX output', async () => {
    const document = wordDocumentXml(`
      <w:tbl>
        <w:tblPr>
          <w:tblStyle w:val="ReportTable"/>
          <w:tblLook w:firstRow="1" w:lastRow="0" w:noHBand="0" w:noVBand="1"/>
        </w:tblPr>
        <w:tr>
          <w:tc><w:p><w:r><w:t>Styled header</w:t></w:r></w:p></w:tc>
        </w:tr>
        <w:tr>
          <w:tc><w:p><w:r><w:t>Styled body</w:t></w:r></w:p></w:tc>
        </w:tr>
      </w:tbl>
    `);
    const styles = tableStylesXml();
    const theme = officeThemeXml();
    const runs = markDocxRunFormatting(document, styles, theme, styles);
    const cells = markDocxTableCells(document, theme, styles);
    const html = new DOMParser().parseFromString(
      `<table><tbody><tr><td>${cells.cells[0]?.marker}<p>${runs.runs[0]?.startMarker}Styled header${runs.runs[0]?.endMarker}</p></td></tr><tr><td>${cells.cells[1]?.marker}<p>${runs.runs[1]?.startMarker}Styled body${runs.runs[1]?.endMarker}</p></td></tr></tbody></table>`,
      'text/html',
    );
    applyImportedDocxRunFormattingMarkers(html, runs);
    applyImportedDocxTableCellMarkers(html, cells);
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    artifact.content.html = html.body.innerHTML;

    const exported = await JSZip.loadAsync(
      await (await createArtifactBlob(artifact)).arrayBuffer(),
    );
    const documentXml =
      (await exported.file('word/document.xml')?.async('string')) ?? '';

    expect(documentXml).toMatch(/<w:shd\b[^>]*w:fill="4472C4"/);
    expect(documentXml).toMatch(/<w:shd\b[^>]*w:fill="F8CBAD"/);
    expect(documentXml).toMatch(/<w:b\/?/);
    expect(documentXml).toMatch(/<w:i\/?/);
    expect(documentXml).toMatch(/<w:u\b/);
    expect(documentXml).toMatch(/<w:strike\/?/);
    expect(documentXml).toMatch(/<w:color\b[^>]*w:val="FFFFFF"/);
  });

  test('loads table styles and the theme from the DOCX package', async () => {
    const archive = new JSZip();
    archive.file(
      'word/document.xml',
      serializeXml(
        wordDocumentXml(`
          <w:tbl>
            <w:tblPr>
              <w:tblStyle w:val="ReportTable"/>
              <w:tblLook w:firstRow="1" w:lastRow="0" w:noHBand="0" w:noVBand="1"/>
            </w:tblPr>
            <w:tr>
              <w:tc><w:p><w:r><w:t>Styled header</w:t></w:r></w:p></w:tc>
            </w:tr>
            <w:tr>
              <w:tc><w:p><w:r><w:t>Styled body</w:t></w:r></w:p></w:tc>
            </w:tr>
          </w:tbl>
          <w:sectPr/>
        `),
      ),
    );
    archive.file('word/styles.xml', serializeXml(tableStylesXml()));
    archive.file('word/theme/theme1.xml', serializeXml(officeThemeXml()));

    const prepared = await prepareDocxImport(
      await archive.generateAsync({ type: 'arraybuffer' }),
    );

    expect(prepared.tableCellMarkers.cells).toHaveLength(2);
    expect(prepared.tableCellMarkers.cells[0]).toMatchObject({
      backgroundColor: '#4472c4',
      borders: {
        bottom: { color: '#ffffff', style: 'double', width: 2 },
      },
    });
    expect(prepared.tableCellMarkers.cells[1]).toMatchObject({
      backgroundColor: '#f8cbad',
    });
    expect(prepared.runFormattingMarkers.runs[0]?.formatting).toMatchObject({
      bold: true,
      color: '#ffffff',
    });
    expect(prepared.runFormattingMarkers.runs[1]?.formatting).toMatchObject({
      color: '#000000',
    });
  });
});

function wordDocumentXml(body: string): Document {
  return parseXml(
    `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body>${body}</w:body></w:document>`,
  );
}

function tableStylesXml(): Document {
  return parseXml(`
    <w:styles xmlns:w="${WORD_NAMESPACE}">
      <w:style w:type="table" w:styleId="BaseTable">
        <w:tblPr>
          <w:tblBorders>
            <w:top w:val="single" w:sz="6" w:themeColor="accent1"/>
            <w:right w:val="single" w:sz="6" w:themeColor="accent1"/>
            <w:bottom w:val="single" w:sz="6" w:themeColor="accent1"/>
            <w:left w:val="single" w:sz="6" w:themeColor="accent1"/>
            <w:insideH w:val="dashed" w:sz="6" w:themeColor="accent3"/>
            <w:insideV w:val="nil"/>
          </w:tblBorders>
        </w:tblPr>
        <w:tcPr><w:shd w:val="clear" w:fill="FFFFFF"/></w:tcPr>
        <w:rPr><w:color w:themeColor="dk1"/></w:rPr>
        <w:tblStylePr w:type="band1Horz">
          <w:tcPr>
            <w:shd w:val="clear" w:themeFill="accent2" w:themeFillTint="99"/>
          </w:tcPr>
        </w:tblStylePr>
        <w:tblStylePr w:type="firstRow">
          <w:tcPr>
            <w:shd w:val="clear" w:themeFill="accent1"/>
            <w:tcBorders>
              <w:bottom w:val="double" w:sz="12" w:themeColor="lt1"/>
            </w:tcBorders>
          </w:tcPr>
          <w:rPr>
            <w:b/>
            <w:i/>
            <w:u w:val="single"/>
            <w:strike/>
            <w:color w:themeColor="lt1"/>
          </w:rPr>
        </w:tblStylePr>
      </w:style>
      <w:style w:type="table" w:styleId="ReportTable">
        <w:basedOn w:val="BaseTable"/>
        <w:tblStylePr w:type="lastRow">
          <w:tcPr><w:shd w:val="clear" w:themeFill="accent3"/></w:tcPr>
        </w:tblStylePr>
      </w:style>
    </w:styles>
  `);
}

function conditionalCoverageStylesXml(): Document {
  return parseXml(`
    <w:styles xmlns:w="${WORD_NAMESPACE}">
      <w:style w:type="table" w:default="1" w:styleId="CoverageTable">
        <w:tblPr>
          <w:tblStyleRowBandSize w:val="2"/>
          <w:tblStyleColBandSize w:val="1"/>
        </w:tblPr>
        <w:tblStylePr w:type="wholeTable">
          <w:tcPr><w:shd w:val="clear" w:fill="111111"/></w:tcPr>
        </w:tblStylePr>
        <w:tblStylePr w:type="band1Horz">
          <w:tcPr><w:shd w:val="clear" w:fill="222222"/></w:tcPr>
        </w:tblStylePr>
        <w:tblStylePr w:type="band2Horz">
          <w:tcPr><w:shd w:val="clear" w:fill="333333"/></w:tcPr>
        </w:tblStylePr>
        <w:tblStylePr w:type="band1Vert">
          <w:tcPr><w:tcBorders><w:top w:val="single" w:sz="6" w:color="CC0000"/></w:tcBorders></w:tcPr>
        </w:tblStylePr>
        <w:tblStylePr w:type="band2Vert">
          <w:tcPr><w:tcBorders><w:top w:val="single" w:sz="6" w:color="0000CC"/></w:tcBorders></w:tcPr>
        </w:tblStylePr>
        <w:tblStylePr w:type="firstRow">
          <w:tcPr><w:shd w:val="clear" w:fill="444444"/></w:tcPr>
        </w:tblStylePr>
        <w:tblStylePr w:type="lastRow">
          <w:tcPr><w:shd w:val="clear" w:fill="555555"/></w:tcPr>
        </w:tblStylePr>
        <w:tblStylePr w:type="firstCol">
          <w:tcPr><w:shd w:val="clear" w:fill="666666"/></w:tcPr>
        </w:tblStylePr>
        <w:tblStylePr w:type="lastCol">
          <w:tcPr><w:shd w:val="clear" w:fill="777777"/></w:tcPr>
        </w:tblStylePr>
        <w:tblStylePr w:type="nwCell">
          <w:tcPr><w:shd w:val="clear" w:fill="888888"/></w:tcPr>
        </w:tblStylePr>
        <w:tblStylePr w:type="neCell">
          <w:tcPr><w:shd w:val="clear" w:fill="999999"/></w:tcPr>
        </w:tblStylePr>
        <w:tblStylePr w:type="swCell">
          <w:tcPr><w:shd w:val="clear" w:fill="AAAAAA"/></w:tcPr>
        </w:tblStylePr>
        <w:tblStylePr w:type="seCell">
          <w:tcPr><w:shd w:val="clear" w:fill="BBBBBB"/></w:tcPr>
        </w:tblStylePr>
      </w:style>
    </w:styles>
  `);
}

function tableRowXml(label: string): string {
  return `<w:tr>
    <w:tc><w:p><w:r><w:t>${label}-1</w:t></w:r></w:p></w:tc>
    <w:tc><w:p><w:r><w:t>${label}-2</w:t></w:r></w:p></w:tc>
    <w:tc><w:p><w:r><w:t>${label}-3</w:t></w:r></w:p></w:tc>
    <w:tc><w:p><w:r><w:t>${label}-4</w:t></w:r></w:p></w:tc>
  </w:tr>`;
}

function officeThemeXml(): Document {
  return parseXml(`
    <a:theme xmlns:a="${DRAWING_NAMESPACE}">
      <a:themeElements>
        <a:clrScheme name="Office">
          <a:dk1><a:srgbClr val="000000"/></a:dk1>
          <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
          <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
          <a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
          <a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>
          <a:accent4><a:srgbClr val="FFC000"/></a:accent4>
        </a:clrScheme>
      </a:themeElements>
    </a:theme>
  `);
}

function serializeXml(document: Document): string {
  return new XMLSerializer().serializeToString(document);
}
