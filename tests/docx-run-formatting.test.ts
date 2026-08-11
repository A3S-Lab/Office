import { describe, expect, test } from '@rstest/core';
import { Editor } from '@tiptap/core';
import JSZip from 'jszip';
import { createArtifact, createArtifactBlob } from '../src/core';
import {
  applyImportedDocxRunFormattingMarkers,
  markDocxRunFormatting,
} from '../src/internal/features/work/work-docx-run-formatting-import';
import { parseXml } from '../src/internal/features/work/work-ooxml-package';
import { createWorkDocumentExtensions } from '../src/internal/features/work/work-document-extensions';

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
            <w:snapToGrid w:val="false"/>
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
      bold: false,
      italic: false,
      underline: false,
      strike: false,
      fontFamily: 'Arial, SimSun',
      wordLineHeightFactor: 1.15,
      wordSnapToGrid: false,
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
            <w:rFonts w:ascii="Segoe UI"/>
            <w:snapToGrid w:val="false"/>
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
    expect(span?.style.fontFamily).toBe('"Segoe UI"');
    expect(span?.dataset.officeWordLineHeightFactor).toBe('1.3301');
    expect(span?.getAttribute('style')).toContain(
      '--work-document-word-line-height-factor: 1.3301',
    );
    expect(span?.dataset.officeWordSnapToGrid).toBe('false');
    expect(span?.getAttribute('style')).toContain('font-size: 12pt');
    expect(span?.getAttribute('style')).toContain('color: #336699');
    expect(html.body.textContent).not.toContain('__A3S_');
  });

  test('preserves semantic run theme colors and drops them after an edit', async () => {
    const document = wordXml('<w:p><w:r><w:t>Theme text</w:t></w:r></w:p>');
    const styles = stylesXml(`
      <w:docDefaults>
        <w:rPrDefault>
          <w:rPr>
            <w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi"/>
            <w:color w:val="4472C4" w:themeColor="accent1" w:themeTint="80"/>
            <w:shd w:val="clear" w:fill="ED7D31" w:themeFill="accent2" w:themeFillShade="BF"/>
          </w:rPr>
        </w:rPrDefault>
      </w:docDefaults>
    `);
    const theme = themeXml(`
      <a:themeElements>
        <a:clrScheme name="Office">
          <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
          <a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
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
      color: '#a2b9e2',
      backgroundColor: '#b25e25',
      themeColor: {
        theme: 'accent1',
        resolved: '#a2b9e2',
        tint: '80',
      },
      themeFill: {
        theme: 'accent2',
        resolved: '#b25e25',
        shade: 'BF',
      },
    });
    const marker = markers.runs[0];
    if (!marker) throw new Error('Expected a run marker.');
    const html = new DOMParser().parseFromString(
      `<p>${marker.startMarker}Theme text${marker.endMarker}</p>`,
      'text/html',
    );
    applyImportedDocxRunFormattingMarkers(html, markers);
    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: html.body.innerHTML,
    });
    html.body.innerHTML = editor.getHTML();
    editor.destroy();
    expect(html.querySelector('span')?.dataset.officeThemeColor).toContain(
      'accent1',
    );
    expect(html.body.innerHTML).toContain('data-office-theme-fill');
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    artifact.content.html = html.body.innerHTML;
    const blob = await createArtifactBlob(artifact);
    const archive = await JSZip.loadAsync(await blob.arrayBuffer());
    const xml = (await archive.file('word/document.xml')?.async('text')) ?? '';
    expect(xml).toMatch(
      /<w:color\b[^>]*w:val="A2B9E2"[^>]*w:themeColor="accent1"[^>]*w:themeTint="80"/,
    );
    expect(xml).toMatch(
      /<w:shd\b[^>]*w:fill="B25E25"[^>]*w:themeFill="accent2"[^>]*w:themeFillShade="BF"/,
    );

    const span = html.querySelector('span');
    if (!span) throw new Error('Expected a themed span.');
    span.style.color = '#123456';
    artifact.content.html = html.body.innerHTML;
    const edited = await createArtifactBlob(artifact);
    const editedArchive = await JSZip.loadAsync(await edited.arrayBuffer());
    const editedXml =
      (await editedArchive.file('word/document.xml')?.async('text')) ?? '';
    expect(editedXml).toContain('<w:color w:val="123456"/>');
    expect(editedXml).not.toMatch(/<w:color\b[^>]*w:themeColor=/);
  });

  test('selects the Word font slot that matches each run script', () => {
    const run = (text: string) => `
      <w:p><w:r>
        <w:rPr>
          <w:rFonts
            w:ascii="Segoe UI"
            w:hAnsi="Calibri"
            w:eastAsia="Microsoft YaHei"
            w:cs="Arial"
          />
        </w:rPr>
        <w:t>${text}</w:t>
      </w:r></w:p>
    `;
    const document = wordXml(
      [
        run('A3S Office'),
        run('\u00e9'),
        run('\u4e2d\u6587\u6d4b\u8bd5'),
        run('\u0627\u062e\u062a\u0628\u0627\u0631'),
        run('(\u0627\u062e\u062a\u0628\u0627\u0631)'),
        run('\u05d1\u05d3\u05d9\u05e7\u05d4'),
      ].join(''),
    );

    const formatting = markDocxRunFormatting(document).runs.map(
      (marker) => marker.formatting,
    );

    expect(formatting).toEqual([
      expect.objectContaining({
        fontFamily: '"Segoe UI", Calibri, "Microsoft YaHei", Arial',
        wordLineHeightFactor: 1.3301,
      }),
      expect.objectContaining({
        fontFamily: 'Calibri, "Segoe UI", "Microsoft YaHei", Arial',
        wordLineHeightFactor: 1.2207,
      }),
      expect.objectContaining({
        fontFamily: '"Microsoft YaHei", Calibri, "Segoe UI", Arial',
        wordLineHeightFactor: 1.7143,
      }),
      expect.objectContaining({
        fontFamily: 'Arial, Calibri, "Segoe UI", "Microsoft YaHei"',
        wordLineHeightFactor: 1.15,
      }),
      expect.objectContaining({
        fontFamily: 'Arial, Calibri, "Segoe UI", "Microsoft YaHei"',
        wordLineHeightFactor: 1.15,
      }),
      expect.objectContaining({
        fontFamily: 'Arial, Calibri, "Segoe UI", "Microsoft YaHei"',
        wordLineHeightFactor: 1.15,
      }),
    ]);
  });

  test('uses complex-script emphasis and size for a complex run', () => {
    const document = wordXml(`
      <w:p><w:r>
        <w:rPr>
          <w:rFonts w:ascii="Segoe UI" w:cs="Arial"/>
          <w:b w:val="false"/><w:bCs/>
          <w:i w:val="false"/><w:iCs/>
          <w:sz w:val="22"/><w:szCs w:val="28"/>
        </w:rPr>
        <w:t>\u0627\u062e\u062a\u0628\u0627\u0631</w:t>
      </w:r></w:p>
    `);

    expect(markDocxRunFormatting(document).runs[0]?.formatting).toMatchObject({
      bold: true,
      italic: true,
      fontFamily: 'Arial, "Segoe UI"',
      wordLineHeightFactor: 1.15,
      fontSize: 14,
    });
  });

  test('honors an explicit complex-script run marker for neutral text', () => {
    const document = wordXml(`
      <w:p><w:r>
        <w:rPr>
          <w:rFonts w:ascii="Segoe UI" w:cs="Arial"/>
          <w:rtl/>
        </w:rPr>
        <w:t>123</w:t>
      </w:r></w:p>
    `);

    expect(markDocxRunFormatting(document).runs[0]?.formatting).toMatchObject({
      fontFamily: 'Arial, "Segoe UI"',
      wordLineHeightFactor: 1.15,
    });
  });

  test('uses the minor theme font and neutralizes omitted heading emphasis', () => {
    const document = wordXml(`
      <w:p>
        <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
        <w:r><w:t>Styled report table</w:t></w:r>
      </w:p>
    `);
    const styles = stylesXml(`
      <w:docDefaults><w:rPrDefault/><w:pPrDefault/></w:docDefaults>
      <w:style w:type="paragraph" w:styleId="Heading1">
        <w:rPr><w:color w:val="2E74B5"/><w:sz w:val="32"/></w:rPr>
      </w:style>
    `);
    const theme = themeXml(`
      <a:themeElements>
        <a:fontScheme name="Office">
          <a:majorFont><a:latin typeface="Aptos Display"/></a:majorFont>
          <a:minorFont><a:latin typeface="Aptos"/></a:minorFont>
        </a:fontScheme>
      </a:themeElements>
    `);

    const markers = markDocxRunFormatting(document, styles, theme);

    expect(markers.runs[0]?.formatting).toEqual({
      bold: false,
      italic: false,
      underline: false,
      strike: false,
      fontFamily: 'Aptos',
      wordLineHeightFactor: 1.15,
      fontSize: 16,
      color: '#2e74b5',
    });
    const marker = markers.runs[0];
    if (!marker) throw new Error('Expected a heading run marker.');
    const html = new DOMParser().parseFromString(
      `<h1>${marker.startMarker}Styled report table${marker.endMarker}</h1>`,
      'text/html',
    );

    applyImportedDocxRunFormattingMarkers(html, markers);

    const span = html.querySelector('h1 > span');
    expect(span?.style.fontFamily).toBe('Aptos');
    expect(span?.style.fontWeight).toBe('normal');
    expect(span?.style.fontStyle).toBe('normal');
  });
});
