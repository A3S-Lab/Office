import { describe, expect, test } from '@rstest/core';
import {
  applyImportedDocxRunFormattingMarkers,
  markDocxRunFormatting,
} from '../src/internal/features/work/work-docx-run-formatting-import';
import { parseXml } from '../src/internal/features/work/work-ooxml-package';

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const DRAWING_NAMESPACE =
  'http://schemas.openxmlformats.org/drawingml/2006/main';

function wordXml(body: string): Document {
  return parseXml(
    `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body>${body}</w:body></w:document>`,
  );
}

function stylesXml(body: string): Document {
  return parseXml(`<w:styles xmlns:w="${WORD_NAMESPACE}">${body}</w:styles>`);
}

function themeXml(body: string): Document {
  return parseXml(`<a:theme xmlns:a="${DRAWING_NAMESPACE}">${body}</a:theme>`);
}

describe('DOCX run formatting', () => {
  test('resolves document, paragraph, character, and direct run properties', () => {
    const document = wordXml(`
      <w:p>
        <w:pPr><w:pStyle w:val="Body"/></w:pPr>
        <w:r>
          <w:rPr>
            <w:rStyle w:val="Emphasis"/>
            <w:rFonts w:eastAsia="SimSun"/>
          </w:rPr>
          <w:t>Styled text</w:t>
        </w:r>
      </w:p>
    `);
    const styles = stylesXml(`
      <w:docDefaults>
        <w:rPrDefault>
          <w:rPr>
            <w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>
            <w:sz w:val="22"/>
          </w:rPr>
        </w:rPrDefault>
      </w:docDefaults>
      <w:style w:type="paragraph" w:styleId="Body">
        <w:rPr><w:color w:val="112233"/></w:rPr>
      </w:style>
      <w:style w:type="character" w:styleId="BaseEmphasis">
        <w:rPr><w:highlight w:val="yellow"/></w:rPr>
      </w:style>
      <w:style w:type="character" w:styleId="Emphasis">
        <w:basedOn w:val="BaseEmphasis"/>
        <w:rPr><w:sz w:val="28"/></w:rPr>
      </w:style>
    `);

    const markers = markDocxRunFormatting(document, styles);

    expect(markers.runs).toHaveLength(1);
    expect(markers.runs[0]?.formatting).toEqual({
      fontFamily: 'SimSun, Arial',
      fontSize: 14,
      color: '#112233',
      backgroundColor: '#ffff00',
    });
  });

  test('applies formatting markers as a structured inline span', () => {
    const document = wordXml(`
      <w:p>
        <w:r>
          <w:rPr>
            <w:rFonts w:ascii="Arial"/>
            <w:sz w:val="24"/>
            <w:color w:val="336699"/>
          </w:rPr>
          <w:t>A3S Office</w:t>
        </w:r>
      </w:p>
    `);
    const markers = markDocxRunFormatting(document);
    const marker = markers.runs[0];
    if (!marker) throw new Error('Expected a run marker.');
    const html = new DOMParser().parseFromString(
      `<p>${marker.startMarker}A3S Office${marker.endMarker}</p>`,
      'text/html',
    );

    applyImportedDocxRunFormattingMarkers(html, markers);

    const span = html.querySelector('p > span');
    expect(span?.textContent).toBe('A3S Office');
    expect(span?.getAttribute('style')).toContain('font-family: Arial');
    expect(span?.getAttribute('style')).toContain('font-size: 12pt');
    expect(span?.getAttribute('style')).toContain('color: #336699');
    expect(html.body.textContent).not.toContain('__A3S_');
  });

  test('resolves theme fonts and colors used by Word defaults', () => {
    const document = wordXml('<w:p><w:r><w:t>Theme text</w:t></w:r></w:p>');
    const styles = stylesXml(`
      <w:docDefaults>
        <w:rPrDefault>
          <w:rPr>
            <w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi"/>
            <w:color w:themeColor="accent1"/>
          </w:rPr>
        </w:rPrDefault>
      </w:docDefaults>
    `);
    const theme = themeXml(`
      <a:themeElements>
        <a:clrScheme name="Office">
          <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
        </a:clrScheme>
        <a:fontScheme name="Office">
          <a:majorFont><a:latin typeface="Aptos Display"/></a:majorFont>
          <a:minorFont><a:latin typeface="Aptos"/></a:minorFont>
        </a:fontScheme>
      </a:themeElements>
    `);

    const markers = markDocxRunFormatting(document, styles, theme);

    expect(markers.runs[0]?.formatting).toMatchObject({
      fontFamily: 'Aptos',
      color: '#4472c4',
    });
  });
});
