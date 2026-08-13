import { describe, expect, test } from '@rstest/core';
import { Editor } from '@tiptap/core';
import JSZip from 'jszip';
import { importOfficeFile } from '../src/core';
import { createDocumentPageChromeEditorExtensions } from '../src/internal/features/work/editors/document-page-chrome-editor';
import { createWorkDocumentModelFromContent } from '../src/internal/features/work/work-document-model-codec';
import {
  DOCUMENT_PARAGRAPH_SHADING_PATTERNS,
  documentParagraphShadingDomAttributes,
  parseDocumentParagraphShading,
} from '../src/internal/features/work/work-document-paragraph-shading';
import { createDocxBlob } from '../src/internal/features/work/work-docx-export';
import { prepareDocxImport } from '../src/internal/features/work/work-docx-import';
import {
  applyImportedDocxParagraphShadingMarkers,
  markDocxParagraphShading,
  parseDocxParagraphShadingElement,
} from '../src/internal/features/work/work-docx-paragraph-shading-import';
import { createDocxParagraphStyleResolver } from '../src/internal/features/work/work-docx-paragraph-styles';
import { createDocxTableStyleResolver } from '../src/internal/features/work/work-docx-table-styles';
import { createDocxThemeResolver } from '../src/internal/features/work/work-docx-theme';
import { analyzeDocxCompatibility } from '../src/internal/features/work/work-office-diagnostics';
import {
  attribute,
  descendants,
  parseXml,
} from '../src/internal/features/work/work-ooxml-package';
import type { WorkDocumentNode } from '../src/internal/features/work/work-types';

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const STRICT_WORD_NAMESPACE =
  'http://purl.oclc.org/ooxml/wordprocessingml/main';
const DRAWING_NAMESPACE =
  'http://schemas.openxmlformats.org/drawingml/2006/main';
const RELATIONSHIP_NAMESPACE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

describe('DOCX paragraph shading', () => {
  test('resolves style and theme shading through edit and DOCX export', async () => {
    const document = wordXml(`
      <w:p>
        <w:pPr><w:pStyle w:val="Derived"/></w:pPr>
        <w:r><w:t>Themed pattern</w:t></w:r>
      </w:p>
    `);
    const styles = stylesXml(`
      <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
        <w:pPr><w:shd w:val="clear" w:fill="FFFFFF"/></w:pPr>
      </w:style>
      <w:style w:type="paragraph" w:styleId="Derived">
        <w:basedOn w:val="Normal"/>
        <w:pPr>
          <w:shd w:val="pct20"
            w:themeColor="accent6" w:themeShade="BF"
            w:themeFill="accent3" w:themeFillTint="80" w:themeFillShade="40"/>
        </w:pPr>
      </w:style>
    `);
    const theme = themeXml(`
      <a:themeElements>
        <a:clrScheme name="Office">
          <a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>
          <a:accent6><a:srgbClr val="70AD47"/></a:accent6>
        </a:clrScheme>
      </a:themeElements>
    `);

    const markers = markDocxParagraphShading(document, styles, theme);

    expect(markers).toMatchObject({ invalidCount: 0, spoofedCount: 0 });
    expect(markers.paragraphs).toHaveLength(1);
    expect(markers.paragraphs[0]?.shading).toEqual({
      pattern: 'pct20',
      color: {
        value: '#548235',
        theme: {
          theme: 'accent6',
          resolved: '#548235',
          shade: 'BF',
        },
      },
      fill: {
        value: '#d2d2d2',
        theme: {
          theme: 'accent3',
          resolved: '#d2d2d2',
          tint: '80',
          shade: '40',
        },
      },
    });

    const marker = markers.paragraphs[0];
    if (!marker) throw new Error('Expected a paragraph shading marker.');
    const html = new DOMParser().parseFromString(
      `<p>${marker.marker}Themed pattern</p>`,
      'text/html',
    );
    applyImportedDocxParagraphShadingMarkers(html, markers);
    const paragraph = html.body.querySelector<HTMLElement>('p');
    expect(paragraph?.style.backgroundColor).toBe('#d2d2d2');
    expect(paragraph?.style.backgroundImage).toContain('radial-gradient');
    expect(
      parseDocumentParagraphShading(paragraph?.dataset.officeParagraphShading),
    ).toEqual(marker.shading);
    expect(html.body.textContent).toBe('Themed pattern');

    const controlled = createWorkDocumentModelFromContent({
      type: 'document',
      pageSize: 'a4',
      html: html.body.innerHTML,
    });
    expect(
      collectNodes(controlled.model?.root, 'paragraph')[0]?.attrs
        ?.paragraphShading,
    ).toEqual(marker.shading);

    const blob = await createDocxBlob({
      type: 'document',
      pageSize: 'a4',
      html: html.body.innerHTML,
    });
    const archive = await JSZip.loadAsync(await blob.arrayBuffer());
    const compatibility = await analyzeDocxCompatibility(
      new File([blob], 'paragraph-shading.docx', { type: blob.type }),
      [],
    );
    expect(
      compatibility.issues.find(
        (issue) => issue.code === 'docx.paragraph-shading',
      ),
    ).toMatchObject({ severity: 'info' });
    const exported = parseXml(
      (await archive.file('word/document.xml')?.async('text')) ?? '',
    );
    const shading = descendants(exported, 'shd').find(
      (element) => element.parentElement?.localName === 'pPr',
    );
    if (!shading) throw new Error('Expected exported paragraph shading.');
    expect(attribute(shading, 'val')).toBe('pct20');
    expect(attribute(shading, 'color')).toBe('548235');
    expect(attribute(shading, 'themeColor')).toBe('accent6');
    expect(attribute(shading, 'themeShade')).toBe('BF');
    expect(attribute(shading, 'fill')).toBe('D2D2D2');
    expect(attribute(shading, 'themeFill')).toBe('accent3');
    expect(attribute(shading, 'themeFillTint')).toBe('80');
    expect(attribute(shading, 'themeFillShade')).toBe('40');

    if (!paragraph) throw new Error('Expected imported paragraph HTML.');
    paragraph.style.backgroundColor = '#123456';
    const editedBlob = await createDocxBlob({
      type: 'document',
      pageSize: 'a4',
      html: html.body.innerHTML,
    });
    const editedArchive = await JSZip.loadAsync(await editedBlob.arrayBuffer());
    const editedDocument = parseXml(
      (await editedArchive.file('word/document.xml')?.async('text')) ?? '',
    );
    const editedShading = descendants(editedDocument, 'shd').find(
      (element) => element.parentElement?.localName === 'pPr',
    );
    if (!editedShading) throw new Error('Expected edited paragraph shading.');
    expect(attribute(editedShading, 'val')).toBe('pct20');
    expect(attribute(editedShading, 'color')).toBe('548235');
    expect(attribute(editedShading, 'themeColor')).toBe('accent6');
    expect(attribute(editedShading, 'fill')).toBe('123456');
    expect(attribute(editedShading, 'themeFill')).toBeNull();
  });

  test('applies conditional table paragraph shading before direct formatting', () => {
    const stylesDocument = stylesXml(`
      <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
        <w:pPr><w:shd w:val="clear" w:fill="FFFFFF"/></w:pPr>
      </w:style>
      <w:style w:type="table" w:styleId="Report">
        <w:tblStylePr w:type="wholeTable">
          <w:pPr><w:shd w:val="clear" w:fill="DDEBF7"/></w:pPr>
        </w:tblStylePr>
        <w:tblStylePr w:type="firstRow">
          <w:pPr><w:shd w:val="pct25" w:color="FFFFFF" w:fill="4472C4"/></w:pPr>
        </w:tblStylePr>
      </w:style>
    `);
    const document = wordXml(`
      <w:tbl>
        <w:tblPr>
          <w:tblStyle w:val="Report"/>
          <w:tblLook w:firstRow="1" w:noHBand="1" w:noVBand="1"/>
        </w:tblPr>
        <w:tr><w:tc><w:p><w:r><w:t>Header</w:t></w:r></w:p></w:tc></w:tr>
        <w:tr><w:tc><w:p><w:pPr><w:shd w:val="nil"/></w:pPr><w:r><w:t>Reset</w:t></w:r></w:p></w:tc></w:tr>
      </w:tbl>
    `);
    const paragraphStyles = createDocxParagraphStyleResolver(stylesDocument);
    const tableStyles = createDocxTableStyleResolver(stylesDocument);

    const markers = markDocxParagraphShading(
      document,
      paragraphStyles,
      undefined,
      tableStyles,
    );

    expect(markers.paragraphs.map(({ shading }) => shading)).toEqual([
      {
        pattern: 'pct25',
        color: { value: '#ffffff' },
        fill: { value: '#4472c4' },
      },
      { pattern: 'nil' },
    ]);
  });

  test('resolves every core Word theme-color family from DrawingML keys', () => {
    const theme = themeXml(`
      <a:themeElements><a:clrScheme name="Office">
        <a:dk1><a:srgbClr val="010203"/></a:dk1>
        <a:lt1><a:srgbClr val="F1F2F3"/></a:lt1>
        <a:dk2><a:srgbClr val="112233"/></a:dk2>
        <a:lt2><a:srgbClr val="DDEEFF"/></a:lt2>
        <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
        <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
        <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
      </a:clrScheme></a:themeElements>
    `);
    const expected = new Map([
      ['dark1', '#010203'],
      ['text1', '#010203'],
      ['light1', '#f1f2f3'],
      ['background1', '#f1f2f3'],
      ['dark2', '#112233'],
      ['text2', '#112233'],
      ['light2', '#ddeeff'],
      ['background2', '#ddeeff'],
      ['accent1', '#4472c4'],
      ['hyperlink', '#0563c1'],
      ['followedHyperlink', '#954f72'],
    ]);
    for (const [themeColor, resolved] of expected) {
      const shading = descendants(
        wordXml(
          `<w:p><w:pPr><w:shd w:val="clear" w:themeFill="${themeColor}"/></w:pPr></w:p>`,
        ),
        'shd',
      )[0];
      if (!shading) throw new Error(`Expected ${themeColor} shading.`);
      expect(parseDocxParagraphShadingElement(shading, theme)).toEqual({
        pattern: 'clear',
        fill: {
          value: resolved,
          theme: { theme: themeColor, resolved },
        },
      });
    }
  });

  test('follows custom color-scheme mappings and bounds cyclic mappings', () => {
    const theme = themeXml(`
      <a:themeElements><a:clrScheme name="Custom">
        <a:dk2><a:srgbClr val="112233"/></a:dk2>
        <a:accent6><a:srgbClr val="70AD47"/></a:accent6>
      </a:clrScheme></a:themeElements>
    `);
    const shading = descendants(
      wordXml(
        '<w:p><w:pPr><w:shd w:val="clear" w:themeFill="background1"/></w:pPr></w:p>',
      ),
      'shd',
    )[0];
    if (!shading) throw new Error('Expected mapped paragraph shading.');

    const mapped = createDocxThemeResolver(
      theme,
      settingsXml('<w:clrSchemeMapping w:bg1="accent6" w:accent6="dark2"/>'),
    );
    expect(parseDocxParagraphShadingElement(shading, mapped)).toEqual({
      pattern: 'clear',
      fill: {
        value: '#112233',
        theme: {
          theme: 'background1',
          resolved: '#112233',
        },
      },
    });

    const cyclic = createDocxThemeResolver(
      theme,
      settingsXml(
        '<w:clrSchemeMapping w:bg1="accent6" w:accent6="accent5" w:accent5="accent6"/>',
      ),
    );
    expect(parseDocxParagraphShadingElement(shading, cyclic)).toBeNull();

    const malformed = createDocxThemeResolver(
      theme,
      settingsXml('<w:clrSchemeMapping w:bg1="accent6" unsafe="dark2"/>'),
    );
    expect(parseDocxParagraphShadingElement(shading, malformed)).toBeNull();
  });

  test('preserves source color-scheme mapping with exported theme shading', async () => {
    const sourceBlob = await createDocxBlob({
      type: 'document',
      pageSize: 'a4',
      html: '<p>Source settings</p>',
    });
    const sourceArchive = await JSZip.loadAsync(await sourceBlob.arrayBuffer());
    const sourceSettings = parseXml(
      (await sourceArchive.file('word/settings.xml')?.async('text')) ?? '',
    );
    const mapping = sourceSettings.createElementNS(
      WORD_NAMESPACE,
      'w:clrSchemeMapping',
    );
    mapping.setAttributeNS(WORD_NAMESPACE, 'w:bg1', 'accent6');
    sourceSettings.documentElement.append(mapping);
    sourceArchive.file('word/settings.xml', serializeXml(sourceSettings));
    const sourcePackage = await sourceArchive.generateAsync({
      type: 'arraybuffer',
    });

    const attributes = documentParagraphShadingDomAttributes({
      pattern: 'clear',
      fill: {
        value: '#70ad47',
        theme: {
          theme: 'background1',
          resolved: '#70ad47',
        },
      },
    });
    const blob = await createDocxBlob(
      {
        type: 'document',
        pageSize: 'a4',
        html: `<p data-office-paragraph-shading='${attributes['data-office-paragraph-shading']}' style="${attributes.style}">Mapped export</p>`,
      },
      sourcePackage,
    );
    const archive = await JSZip.loadAsync(await blob.arrayBuffer());
    const exportedSettings = parseXml(
      (await archive.file('word/settings.xml')?.async('text')) ?? '',
    );
    const exportedMapping = descendants(
      exportedSettings,
      'clrSchemeMapping',
    )[0];
    if (!exportedMapping)
      throw new Error('Expected preserved color-scheme mapping.');
    expect(attribute(exportedMapping, 'bg1')).toBe('accent6');

    const exportedDocument = parseXml(
      (await archive.file('word/document.xml')?.async('text')) ?? '',
    );
    const exportedShading = descendants(exportedDocument, 'shd').find(
      (element) => element.parentElement?.localName === 'pPr',
    );
    if (!exportedShading)
      throw new Error('Expected exported mapped paragraph shading.');
    expect(attribute(exportedShading, 'fill')).toBe('70AD47');
    expect(attribute(exportedShading, 'themeFill')).toBe('background1');
  });

  test('round-trips page-chrome shading through sanitization and its editor model', async () => {
    const attributes = documentParagraphShadingDomAttributes({
      pattern: 'diagCross',
      color: { value: '#112233' },
      fill: { value: '#ddeeff' },
    });
    const content = {
      type: 'document' as const,
      html: '<p>Body</p>',
      pageSize: 'a4' as const,
      pageChrome: {
        differentFirstPage: false,
        differentOddEvenPages: false,
        default: {
          headerHtml: `<p data-office-paragraph-shading='${attributes['data-office-paragraph-shading']}' style="${attributes.style}">Patterned header</p>`,
          footerHtml: '',
          showPageNumber: false,
        },
        first: { headerHtml: '', footerHtml: '', showPageNumber: false },
        even: { headerHtml: '', footerHtml: '', showPageNumber: false },
      },
    };

    const blob = await createDocxBlob(content);
    const archive = await JSZip.loadAsync(await blob.arrayBuffer());
    const headerPath = Object.keys(archive.files).find((path) =>
      /^word\/header\d*\.xml$/i.test(path),
    );
    if (!headerPath) throw new Error('Expected a generated header part.');
    const header = parseXml(
      (await archive.file(headerPath)?.async('text')) ?? '',
    );
    const exportedShading = descendants(header, 'shd').find(
      (element) => element.parentElement?.localName === 'pPr',
    );
    if (!exportedShading)
      throw new Error('Expected exported header paragraph shading.');
    expect(attribute(exportedShading, 'val')).toBe('diagCross');
    expect(attribute(exportedShading, 'color')).toBe('112233');
    expect(attribute(exportedShading, 'fill')).toBe('DDEEFF');

    const imported = await importOfficeFile(
      new File([blob], 'paragraph-shading.docx', { type: blob.type }),
    );
    if (imported.content.type !== 'document')
      throw new Error('Expected an imported document artifact.');
    const headerHtml = imported.content.pageChrome?.default.headerHtml ?? '';
    expect(headerHtml).toContain('data-office-paragraph-shading');
    expect(headerHtml).toContain('diagCross');
    expect(headerHtml).toContain('background-image');

    const editor = new Editor({
      extensions: createDocumentPageChromeEditorExtensions(),
      content: headerHtml,
    });
    expect(editor.getJSON().content?.[0]?.attrs?.paragraphShading).toEqual({
      pattern: 'diagCross',
      color: { value: '#112233' },
      fill: { value: '#ddeeff' },
    });
    expect(editor.getHTML()).toContain('data-office-paragraph-shading');
    editor.destroy();

    const roundTripped = await createDocxBlob(imported.content);
    const secondArchive = await JSZip.loadAsync(
      await roundTripped.arrayBuffer(),
    );
    const secondHeaderPath = Object.keys(secondArchive.files).find((path) =>
      /^word\/header\d*\.xml$/i.test(path),
    );
    if (!secondHeaderPath)
      throw new Error('Expected a round-tripped header part.');
    const secondHeader = parseXml(
      (await secondArchive.file(secondHeaderPath)?.async('text')) ?? '',
    );
    const secondShading = descendants(secondHeader, 'shd').find(
      (element) => element.parentElement?.localName === 'pPr',
    );
    if (!secondShading)
      throw new Error('Expected round-tripped header paragraph shading.');
    expect(attribute(secondShading, 'val')).toBe('diagCross');
    expect(attribute(secondShading, 'color')).toBe('112233');
    expect(attribute(secondShading, 'fill')).toBe('DDEEFF');
  });

  test('exports list-item paragraph shading without losing numbering', async () => {
    const attributes = documentParagraphShadingDomAttributes({
      pattern: 'clear',
      fill: { value: '#fff2cc' },
    });
    const blob = await createDocxBlob({
      type: 'document',
      pageSize: 'a4',
      html: `<ol><li><p data-office-paragraph-shading='${attributes['data-office-paragraph-shading']}' style="${attributes.style}">Shaded item</p></li></ol>`,
    });
    const archive = await JSZip.loadAsync(await blob.arrayBuffer());
    const document = parseXml(
      (await archive.file('word/document.xml')?.async('text')) ?? '',
    );
    const paragraph = descendants(document, 'p').find((element) =>
      element.textContent?.includes('Shaded item'),
    );
    if (!paragraph) throw new Error('Expected exported list paragraph.');
    const properties = descendants(paragraph, 'pPr')[0];
    if (!properties) throw new Error('Expected list paragraph properties.');
    expect(descendants(properties, 'numPr')).toHaveLength(1);
    const shading = descendants(properties, 'shd')[0];
    if (!shading) throw new Error('Expected list paragraph shading.');
    expect(attribute(shading, 'val')).toBe('clear');
    expect(attribute(shading, 'fill')).toBe('FFF2CC');
  });

  test('accepts every schema pattern and strict WordprocessingML', () => {
    for (const pattern of DOCUMENT_PARAGRAPH_SHADING_PATTERNS) {
      const element = descendants(
        strictWordXml(
          `<w:p><w:pPr><w:shd w:val="${pattern}" w:color="112233" w:fill="AABBCC"/></w:pPr></w:p>`,
        ),
        'shd',
      )[0];
      if (!element) throw new Error(`Expected strict ${pattern} shading.`);
      expect(parseDocxParagraphShadingElement(element)).toEqual({
        pattern,
        color: { value: '#112233' },
        fill: { value: '#aabbcc' },
      });
    }
  });

  test('normalizes malformed direct values to nil and ignores namespace spoofing', () => {
    const document = wordXml(
      `
        <w:p>
          <w:pPr><w:pStyle w:val="Shaded"/><w:shd w:fill="112233"/></w:pPr>
          <w:r><w:t>Missing pattern</w:t></w:r>
        </w:p>
        <w:p>
          <w:pPr><w:shd w:val="clear"/><w:shd w:val="solid"/></w:pPr>
          <w:r><w:t>Duplicate</w:t></w:r>
        </w:p>
        <w:p>
          <w:pPr><w:shd xmlns:r="${RELATIONSHIP_NAMESPACE}" w:val="clear" r:id="rIdUnsafe"/></w:pPr>
          <w:r><w:t>Relationship bound</w:t></w:r>
        </w:p>
        <w:p xmlns:x="urn:spoof">
          <w:pPr><w:pStyle w:val="Shaded"/><x:shd w:val="solid" w:color="FF0000"/></w:pPr>
          <w:r><w:t>Spoofed</w:t></w:r>
        </w:p>
      `,
    );
    const styles = stylesXml(`
      <w:style w:type="paragraph" w:styleId="Shaded">
        <w:pPr><w:shd w:val="clear" w:fill="DDEBF7"/></w:pPr>
      </w:style>
    `);

    const markers = markDocxParagraphShading(document, styles);

    expect(markers.invalidCount).toBe(4);
    expect(markers.spoofedCount).toBe(1);
    expect(markers.paragraphs.map(({ shading }) => shading)).toEqual([
      { pattern: 'nil' },
      { pattern: 'nil' },
      { pattern: 'nil' },
      { pattern: 'clear', fill: { value: '#ddebf7' } },
    ]);
  });

  test('loads paragraph shading at the DOCX package boundary', async () => {
    const archive = new JSZip();
    archive.file(
      'word/document.xml',
      serializeXml(
        wordXml(
          '<w:p><w:pPr><w:shd w:val="diagCross" w:color="112233" w:themeFill="background1"/></w:pPr><w:r><w:t>Package</w:t></w:r></w:p>',
        ),
      ),
    );
    archive.file(
      'word/theme/theme1.xml',
      serializeXml(
        themeXml(`
          <a:themeElements><a:clrScheme name="Custom">
            <a:dk2><a:srgbClr val="DDEEFF"/></a:dk2>
          </a:clrScheme></a:themeElements>
        `),
      ),
    );
    archive.file(
      'word/settings.xml',
      serializeXml(settingsXml('<w:clrSchemeMapping w:bg1="dark2"/>')),
    );

    const prepared = await prepareDocxImport(
      await archive.generateAsync({ type: 'arraybuffer' }),
    );

    expect(prepared.paragraphShadingMarkers.paragraphs[0]?.shading).toEqual({
      pattern: 'diagCross',
      color: { value: '#112233' },
      fill: {
        value: '#ddeeff',
        theme: {
          theme: 'background1',
          resolved: '#ddeeff',
        },
      },
    });
    expect(prepared.conversionBuffer.byteLength).toBeGreaterThan(0);
  });
});

function wordXml(body: string): Document {
  return parseXml(
    `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body>${body}</w:body></w:document>`,
  );
}

function strictWordXml(body: string): Document {
  return parseXml(
    `<w:document xmlns:w="${STRICT_WORD_NAMESPACE}"><w:body>${body}</w:body></w:document>`,
  );
}

function stylesXml(body: string): Document {
  return parseXml(`<w:styles xmlns:w="${WORD_NAMESPACE}">${body}</w:styles>`);
}

function themeXml(body: string): Document {
  return parseXml(`<a:theme xmlns:a="${DRAWING_NAMESPACE}">${body}</a:theme>`);
}

function settingsXml(body: string): Document {
  return parseXml(
    `<w:settings xmlns:w="${WORD_NAMESPACE}">${body}</w:settings>`,
  );
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
