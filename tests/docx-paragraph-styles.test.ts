import { describe, expect, test } from '@rstest/core';
import JSZip from 'jszip';
import { prepareDocxImport } from '../src/internal/features/work/work-docx-import';
import {
  applyImportedDocxParagraphDirectionMarkers,
  markDocxParagraphDirections,
} from '../src/internal/features/work/work-docx-paragraph-direction-import';
import { markDocxParagraphIndents } from '../src/internal/features/work/work-docx-paragraph-indent-import';
import { markDocxParagraphPagination } from '../src/internal/features/work/work-docx-paragraph-pagination-import';
import { markDocxParagraphSpacing } from '../src/internal/features/work/work-docx-paragraph-spacing-import';
import { markDocxParagraphTabStops } from '../src/internal/features/work/work-docx-tab-stop-import';
import { parseXml } from '../src/internal/features/work/work-ooxml-package';

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

function wordXml(body: string): Document {
  return parseXml(
    `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body>${body}</w:body></w:document>`,
  );
}

function stylesXml(body: string): Document {
  return parseXml(`<w:styles xmlns:w="${WORD_NAMESPACE}">${body}</w:styles>`);
}

describe('DOCX paragraph styles', () => {
  test('resolves pagination from defaults, based-on styles, and direct overrides', () => {
    const document = wordXml(`
      <w:p>
        <w:pPr>
          <w:pStyle w:val="Derived"/>
          <w:keepLines w:val="0"/>
        </w:pPr>
        <w:r><w:t>Styled paragraph</w:t></w:r>
      </w:p>
    `);
    const styles = stylesXml(`
      <w:docDefaults>
        <w:pPrDefault>
          <w:pPr><w:widowControl w:val="0"/></w:pPr>
        </w:pPrDefault>
      </w:docDefaults>
      <w:style w:type="paragraph" w:styleId="Base">
        <w:pPr>
          <w:keepLines/>
          <w:pageBreakBefore/>
        </w:pPr>
      </w:style>
      <w:style w:type="paragraph" w:styleId="Derived">
        <w:basedOn w:val="Base"/>
        <w:pPr><w:keepNext/></w:pPr>
      </w:style>
    `);

    const markers = markDocxParagraphPagination(document, styles);

    expect(markers.paragraphs).toHaveLength(1);
    expect(markers.paragraphs[0]?.pagination).toEqual({
      keepLines: false,
      keepWithNext: true,
      pageBreakBefore: true,
      widowControl: false,
    });
  });

  test('uses the default paragraph style when a paragraph has no direct properties', () => {
    const document = wordXml(
      '<w:p><w:r><w:t>Default styled paragraph</w:t></w:r></w:p>',
    );
    const styles = stylesXml(`
      <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
        <w:pPr><w:keepNext/></w:pPr>
      </w:style>
    `);

    const markers = markDocxParagraphPagination(document, styles);

    expect(markers.paragraphs).toHaveLength(1);
    expect(markers.paragraphs[0]?.pagination).toEqual({
      keepWithNext: true,
    });
  });

  test('loads paragraph styles from the DOCX package import boundary', async () => {
    const archive = new JSZip();
    archive.file(
      'word/document.xml',
      new XMLSerializer().serializeToString(
        wordXml(`
          <w:p>
            <w:pPr><w:pStyle w:val="Heading"/></w:pPr>
            <w:r><w:t>Styled heading</w:t></w:r>
          </w:p>
        `),
      ),
    );
    archive.file(
      'word/styles.xml',
      new XMLSerializer().serializeToString(
        stylesXml(`
          <w:style w:type="paragraph" w:styleId="Heading">
            <w:pPr><w:keepNext/></w:pPr>
            <w:rPr>
              <w:rFonts w:ascii="Arial"/>
              <w:sz w:val="24"/>
            </w:rPr>
          </w:style>
        `),
      ),
    );

    const prepared = await prepareDocxImport(
      await archive.generateAsync({ type: 'arraybuffer' }),
    );

    expect(prepared.paragraphPaginationMarkers.paragraphs).toHaveLength(1);
    expect(
      prepared.paragraphPaginationMarkers.paragraphs[0]?.pagination,
    ).toEqual({ keepWithNext: true });
    expect(prepared.runFormattingMarkers.runs[0]?.formatting).toEqual({
      fontFamily: 'Arial',
      fontSize: 12,
    });
  });

  test('merges inherited and direct indent attributes independently', () => {
    const document = wordXml(`
      <w:p>
        <w:pPr>
          <w:pStyle w:val="Derived"/>
          <w:ind w:right="0" w:firstLine="180"/>
        </w:pPr>
        <w:r><w:t>Indented paragraph</w:t></w:r>
      </w:p>
    `);
    const styles = stylesXml(`
      <w:docDefaults>
        <w:pPrDefault><w:pPr><w:ind w:left="360"/></w:pPr></w:pPrDefault>
      </w:docDefaults>
      <w:style w:type="paragraph" w:styleId="Base">
        <w:pPr><w:ind w:right="180"/></w:pPr>
      </w:style>
      <w:style w:type="paragraph" w:styleId="Derived">
        <w:basedOn w:val="Base"/>
      </w:style>
    `);

    const markers = markDocxParagraphIndents(document, styles);

    expect(markers.paragraphs).toHaveLength(1);
    expect(markers.paragraphs[0]?.indent).toEqual({
      left: 24,
      right: 0,
      firstLine: 12,
    });
  });

  test('resolves paragraph direction through styles and explicit overrides', () => {
    const document = wordXml(`
      <w:p>
        <w:pPr>
          <w:pStyle w:val="RightToLeft"/>
          <w:bidi w:val="0"/>
        </w:pPr>
        <w:r><w:t>Explicit left-to-right paragraph</w:t></w:r>
      </w:p>
      <w:p>
        <w:pPr><w:pStyle w:val="RightToLeft"/></w:pPr>
        <w:r><w:t>فقرة من اليمين إلى اليسار</w:t></w:r>
      </w:p>
    `);
    const styles = stylesXml(`
      <w:style w:type="paragraph" w:styleId="RightToLeft">
        <w:pPr><w:bidi/></w:pPr>
      </w:style>
    `);

    const markers = markDocxParagraphDirections(document, styles);

    expect(markers.paragraphs.map(({ direction }) => direction)).toEqual([
      'ltr',
      'rtl',
    ]);
  });

  test('keeps direction markers on numbered list items', () => {
    const document = wordXml(`
      <w:p>
        <w:pPr>
          <w:numPr>
            <w:ilvl w:val="1"/>
            <w:numId w:val="7"/>
          </w:numPr>
          <w:bidi/>
        </w:pPr>
        <w:r><w:t>عنصر قائمة</w:t></w:r>
      </w:p>
    `);

    const markers = markDocxParagraphDirections(document);
    expect(markers.paragraphs).toHaveLength(1);
    expect(markers.paragraphs[0]?.direction).toBe('rtl');

    const html = new DOMParser().parseFromString(
      `<ol><li>${markers.paragraphs[0]?.marker}عنصر قائمة</li></ol>`,
      'text/html',
    );
    applyImportedDocxParagraphDirectionMarkers(html, markers);

    expect(html.body.querySelector('li')).toHaveAttribute('dir', 'rtl');
    expect(html.body.textContent).toBe('عنصر قائمة');
  });

  test('resolves paragraph spacing and line rules through the style chain', () => {
    const document = wordXml(`
      <w:p>
        <w:pPr>
          <w:pStyle w:val="Derived"/>
          <w:spacing w:after="0" w:line="360"/>
        </w:pPr>
        <w:r><w:t>Spaced paragraph</w:t></w:r>
      </w:p>
    `);
    const styles = stylesXml(`
      <w:docDefaults>
        <w:pPrDefault>
          <w:pPr><w:spacing w:before="240"/></w:pPr>
        </w:pPrDefault>
      </w:docDefaults>
      <w:style w:type="paragraph" w:styleId="Base">
        <w:pPr>
          <w:spacing w:after="120" w:line="240" w:lineRule="exact"/>
        </w:pPr>
      </w:style>
      <w:style w:type="paragraph" w:styleId="Derived">
        <w:basedOn w:val="Base"/>
        <w:pPr><w:spacing w:before="120"/></w:pPr>
      </w:style>
    `);

    const markers = markDocxParagraphSpacing(document, styles);

    expect(markers.paragraphs).toHaveLength(1);
    expect(markers.paragraphs[0]?.spacing).toEqual({
      before: 6,
      after: 0,
      lineHeight: '1.5',
      lineRule: 'auto',
    });
  });

  test('merges inherited tab stops by position and honors direct clears', () => {
    const document = wordXml(`
      <w:p>
        <w:pPr>
          <w:pStyle w:val="Derived"/>
          <w:tabs>
            <w:tab w:val="clear" w:pos="720"/>
            <w:tab w:val="decimal" w:pos="2160" w:leader="underscore"/>
          </w:tabs>
        </w:pPr>
        <w:r><w:t>Item</w:t><w:tab/><w:t>12.50</w:t></w:r>
      </w:p>
    `);
    const styles = stylesXml(`
      <w:docDefaults>
        <w:pPrDefault>
          <w:pPr>
            <w:tabs><w:tab w:val="left" w:pos="720"/></w:tabs>
          </w:pPr>
        </w:pPrDefault>
      </w:docDefaults>
      <w:style w:type="paragraph" w:styleId="Base">
        <w:pPr>
          <w:tabs>
            <w:tab w:val="right" w:pos="1440" w:leader="dot"/>
          </w:tabs>
        </w:pPr>
      </w:style>
      <w:style w:type="paragraph" w:styleId="Derived">
        <w:basedOn w:val="Base"/>
        <w:pPr>
          <w:tabs><w:tab w:val="center" w:pos="1440"/></w:tabs>
        </w:pPr>
      </w:style>
    `);

    const markers = markDocxParagraphTabStops(document, styles);

    expect(markers.paragraphs).toHaveLength(1);
    expect(markers.paragraphs[0]?.tabStops).toEqual([
      { position: 96, alignment: 'center', leader: 'none' },
      { position: 144, alignment: 'decimal', leader: 'underscore' },
    ]);
    expect(markers.inlineTabs).toHaveLength(1);
  });
});
