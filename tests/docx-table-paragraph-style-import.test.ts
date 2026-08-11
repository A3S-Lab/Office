import { describe, expect, test } from '@rstest/core';
import JSZip from 'jszip';
import { prepareDocxImport } from '../src/internal/features/work/work-docx-import';
import { createDocxBlob } from '../src/internal/features/work/work-docx-export';
import { createWorkDocumentModelFromContent } from '../src/internal/features/work/work-document-model-codec';
import { markDocxParagraphAlignments } from '../src/internal/features/work/work-docx-paragraph-alignment-import';
import { markDocxParagraphDirections } from '../src/internal/features/work/work-docx-paragraph-direction-import';
import { markDocxParagraphIndents } from '../src/internal/features/work/work-docx-paragraph-indent-import';
import {
  applyImportedDocxParagraphPaginationMarkers,
  markDocxParagraphPagination,
} from '../src/internal/features/work/work-docx-paragraph-pagination-import';
import { markDocxParagraphSpacing } from '../src/internal/features/work/work-docx-paragraph-spacing-import';
import { createDocxParagraphStyleResolver } from '../src/internal/features/work/work-docx-paragraph-styles';
import { markDocxParagraphTabStops } from '../src/internal/features/work/work-docx-tab-stop-import';
import { createDocxTableStyleResolver } from '../src/internal/features/work/work-docx-table-styles';
import { parseXml } from '../src/internal/features/work/work-ooxml-package';
import type { WorkDocumentNode } from '../src/internal/features/work/work-types';

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

describe('DOCX conditional table paragraph styles', () => {
  test('merges table paragraph layout after paragraph styles and before direct formatting', () => {
    const stylesDocument = tableParagraphStylesXml();
    const paragraphStyles = createDocxParagraphStyleResolver(stylesDocument);
    const tableStyles = createDocxTableStyleResolver(stylesDocument);

    expect(
      markDocxParagraphAlignments(
        tableDocumentXml(),
        paragraphStyles,
        tableStyles,
      ).paragraphs.map(({ alignment }) => alignment),
    ).toEqual(['left', 'right']);
    expect(
      markDocxParagraphDirections(
        tableDocumentXml(),
        paragraphStyles,
        tableStyles,
      ).paragraphs.map(({ direction }) => direction),
    ).toEqual(['ltr', 'rtl']);
    expect(
      markDocxParagraphIndents(
        tableDocumentXml(),
        paragraphStyles,
        tableStyles,
      ).paragraphs.map(({ indent }) => indent),
    ).toEqual([
      { left: 24, right: 0, firstLine: 6 },
      { left: 24, right: 12, firstLine: 0 },
    ]);
    expect(
      markDocxParagraphSpacing(
        tableDocumentXml(),
        paragraphStyles,
        tableStyles,
      ).paragraphs.map(({ spacing }) => spacing),
    ).toEqual([
      { before: 12, after: 0, lineHeight: '1.5', lineRule: 'auto' },
      { before: 12, after: 4, lineHeight: '1.25', lineRule: 'auto' },
    ]);
    expect(
      markDocxParagraphPagination(
        tableDocumentXml(),
        paragraphStyles,
        tableStyles,
      ).paragraphs.map(({ pagination }) => pagination),
    ).toEqual([
      {
        contextualSpacing: false,
        keepLines: true,
        keepWithNext: false,
        outlineLevel: 1,
        pageBreakBefore: true,
        widowControl: false,
      },
      {
        contextualSpacing: true,
        keepLines: true,
        keepWithNext: true,
        outlineLevel: 4,
        widowControl: true,
      },
    ]);
    expect(
      markDocxParagraphTabStops(
        tableDocumentXml(),
        paragraphStyles,
        tableStyles,
      ).paragraphs.map(({ tabStops }) => tabStops),
    ).toEqual([
      [
        { position: 72, alignment: 'decimal', leader: 'underscore' },
        { position: 96, alignment: 'right', leader: 'hyphen' },
      ],
      [
        { position: 24, alignment: 'left', leader: 'none' },
        { position: 48, alignment: 'center', leader: 'dot' },
      ],
    ]);
  });

  test('loads conditional table paragraph layout at the DOCX package boundary', async () => {
    const archive = new JSZip();
    archive.file('word/document.xml', serializeXml(tableDocumentXml(true)));
    archive.file('word/styles.xml', serializeXml(tableParagraphStylesXml()));

    const prepared = await prepareDocxImport(
      await archive.generateAsync({ type: 'arraybuffer' }),
    );

    expect(
      prepared.paragraphAlignmentMarkers.paragraphs.map(
        ({ alignment }) => alignment,
      ),
    ).toEqual(['left', 'right']);
    expect(
      prepared.paragraphDirectionMarkers.paragraphs.map(
        ({ direction }) => direction,
      ),
    ).toEqual(['ltr', 'rtl']);
    expect(prepared.paragraphIndentMarkers.paragraphs[0]?.indent).toEqual({
      left: 24,
      right: 0,
      firstLine: 6,
    });
    expect(prepared.paragraphSpacingMarkers.paragraphs[0]?.spacing).toEqual({
      before: 12,
      after: 0,
      lineHeight: '1.5',
      lineRule: 'auto',
    });
    expect(
      prepared.paragraphPaginationMarkers.paragraphs[0]?.pagination,
    ).toEqual({
      contextualSpacing: false,
      keepLines: true,
      keepWithNext: false,
      outlineLevel: 1,
      pageBreakBefore: true,
      widowControl: false,
    });
    expect(prepared.tabStopMarkers.paragraphs[0]?.tabStops).toEqual([
      { position: 72, alignment: 'decimal', leader: 'underscore' },
      { position: 96, alignment: 'right', leader: 'hyphen' },
    ]);

    const html = new DOMParser().parseFromString(
      [
        '<table><tbody><tr><td>',
        `<p>${prepared.paragraphPaginationMarkers.paragraphs[0]?.marker}Header</p>`,
        '</td></tr><tr><td>',
        `<p>${prepared.paragraphPaginationMarkers.paragraphs[1]?.marker}Body</p>`,
        '</td></tr></tbody></table>',
      ].join(''),
      'text/html',
    );
    applyImportedDocxParagraphPaginationMarkers(
      html,
      prepared.paragraphPaginationMarkers,
    );
    const paragraphs = Array.from(html.body.querySelectorAll('p'));
    expect(paragraphs[0]).toHaveAttribute(
      'data-office-contextual-spacing',
      'false',
    );
    expect(paragraphs[0]).toHaveAttribute('data-office-outline-level', '1');
    expect(paragraphs[1]).toHaveAttribute(
      'data-office-contextual-spacing',
      'true',
    );
    expect(paragraphs[1]).toHaveAttribute('data-office-outline-level', '4');

    const controlled = createWorkDocumentModelFromContent({
      type: 'document',
      pageSize: 'a4',
      html: html.body.innerHTML,
    });
    expect(
      collectNodes(controlled.model?.root, 'paragraph').map(({ attrs }) => ({
        contextualSpacing: attrs?.contextualSpacing,
        outlineLevel: attrs?.outlineLevel,
      })),
    ).toEqual([
      { contextualSpacing: false, outlineLevel: 1 },
      { contextualSpacing: true, outlineLevel: 4 },
    ]);

    const blob = await createDocxBlob({
      type: 'document',
      pageSize: 'a4',
      html: html.body.innerHTML,
    });
    const exported = await JSZip.loadAsync(await blob.arrayBuffer());
    const documentXml =
      (await exported.file('word/document.xml')?.async('text')) ?? '';
    expect(documentXml).toContain('<w:contextualSpacing w:val="false"/>');
    expect(documentXml).toContain('<w:outlineLvl w:val="1"/>');
    expect(documentXml).toContain('<w:contextualSpacing/>');
    expect(documentXml).toContain('<w:outlineLvl w:val="4"/>');
  });
});

function tableDocumentXml(withSection = false): Document {
  return parseXml(`
    <w:document xmlns:w="${WORD_NAMESPACE}">
      <w:body>
        <w:tbl>
          <w:tblPr>
            <w:tblStyle w:val="ReportTable"/>
            <w:tblLook w:firstRow="1" w:lastRow="0" w:noHBand="1" w:noVBand="1"/>
          </w:tblPr>
          <w:tr>
            <w:tc>
              <w:p>
                <w:pPr>
                  <w:pStyle w:val="Normal"/>
                  <w:bidi w:val="0"/>
                  <w:jc w:val="left"/>
                  <w:ind w:right="0"/>
                  <w:spacing w:line="360"/>
                  <w:keepNext w:val="0"/>
                  <w:outlineLvl w:val="1"/>
                  <w:tabs>
                    <w:tab w:val="clear" w:pos="720"/>
                    <w:tab w:val="right" w:pos="1440" w:leader="hyphen"/>
                  </w:tabs>
                </w:pPr>
                <w:r><w:t>Direct header</w:t><w:tab/><w:t>Value</w:t></w:r>
              </w:p>
            </w:tc>
          </w:tr>
          <w:tr>
            <w:tc>
              <w:p>
                <w:pPr><w:pStyle w:val="Normal"/></w:pPr>
                <w:r><w:t>Styled body</w:t><w:tab/><w:t>Value</w:t></w:r>
              </w:p>
            </w:tc>
          </w:tr>
        </w:tbl>
        ${withSection ? '<w:sectPr/>' : ''}
      </w:body>
    </w:document>
  `);
}

function tableParagraphStylesXml(): Document {
  return parseXml(`
    <w:styles xmlns:w="${WORD_NAMESPACE}">
      <w:docDefaults>
        <w:pPrDefault>
          <w:pPr>
            <w:tabs><w:tab w:val="left" w:pos="360"/></w:tabs>
          </w:pPr>
        </w:pPrDefault>
      </w:docDefaults>
      <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
        <w:pPr>
          <w:bidi w:val="0"/>
          <w:jc w:val="left"/>
          <w:ind w:left="120"/>
          <w:spacing w:after="80"/>
          <w:widowControl/>
          <w:tabs>
            <w:tab w:val="right" w:pos="720"/>
          </w:tabs>
        </w:pPr>
      </w:style>
      <w:style w:type="table" w:styleId="BaseTable">
        <w:tblStylePr w:type="wholeTable">
          <w:pPr>
            <w:bidi/>
            <w:jc w:val="start"/>
            <w:ind w:left="360" w:right="180"/>
            <w:spacing w:before="240" w:line="300"/>
            <w:keepLines/>
            <w:keepNext/>
            <w:contextualSpacing/>
            <w:outlineLvl w:val="4"/>
            <w:tabs>
              <w:tab w:val="center" w:pos="720" w:leader="dot"/>
            </w:tabs>
          </w:pPr>
        </w:tblStylePr>
      </w:style>
      <w:style w:type="table" w:styleId="ReportTable">
        <w:basedOn w:val="BaseTable"/>
        <w:tblStylePr w:type="firstRow">
          <w:pPr>
            <w:jc w:val="center"/>
            <w:ind w:firstLine="90"/>
            <w:spacing w:after="0"/>
            <w:pageBreakBefore/>
            <w:widowControl w:val="0"/>
            <w:contextualSpacing w:val="0"/>
            <w:outlineLvl w:val="2"/>
            <w:tabs>
              <w:tab w:val="clear" w:pos="360"/>
              <w:tab w:val="decimal" w:pos="1080" w:leader="underscore"/>
            </w:tabs>
          </w:pPr>
        </w:tblStylePr>
      </w:style>
    </w:styles>
  `);
}

function serializeXml(document: Document): string {
  return new XMLSerializer().serializeToString(document);
}

function collectNodes(
  root: WorkDocumentNode | undefined,
  type: string,
): WorkDocumentNode[] {
  if (!root) return [];
  const result: WorkDocumentNode[] = [];
  const pending = [root];
  while (pending.length) {
    const node = pending.shift();
    if (!node) continue;
    if (node.type === type) result.push(node);
    pending.unshift(...(node.content ?? []));
  }
  return result;
}
