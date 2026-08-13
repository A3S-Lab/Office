import { describe, expect, test } from '@rstest/core';
import { Editor } from '@tiptap/core';
import JSZip from 'jszip';
import { importOfficeFile } from '../src/core';
import { createDocumentPageChromeEditorExtensions } from '../src/internal/features/work/editors/document-page-chrome-editor';
import { createWorkDocumentModelFromContent } from '../src/internal/features/work/work-document-model-codec';
import {
  DOCUMENT_PARAGRAPH_BORDER_EDGES,
  DOCUMENT_PARAGRAPH_BORDER_STYLES,
  documentParagraphBordersDomAttributes,
  parseDocumentParagraphBorders,
} from '../src/internal/features/work/work-document-paragraph-borders';
import { createDocxBlob } from '../src/internal/features/work/work-docx-export';
import { prepareDocxImport } from '../src/internal/features/work/work-docx-import';
import {
  applyImportedDocxParagraphBorderMarkers,
  markDocxParagraphBorders,
  parseDirectDocxParagraphBorders,
} from '../src/internal/features/work/work-docx-paragraph-borders-import';
import { createDocxParagraphStyleResolver } from '../src/internal/features/work/work-docx-paragraph-styles';
import { createDocxTableStyleResolver } from '../src/internal/features/work/work-docx-table-styles';
import { createDocxThemeResolver } from '../src/internal/features/work/work-docx-theme';
import { analyzeDocxCompatibility } from '../src/internal/features/work/work-office-diagnostics';
import {
  attribute,
  descendants,
  directChildren,
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

describe('DOCX paragraph borders', () => {
  test('resolves style and theme borders through edit and exact DOCX export', async () => {
    const document = wordXml(`
      <w:p>
        <w:pPr><w:pStyle w:val="Derived"/></w:pPr>
        <w:r><w:t>Themed borders</w:t></w:r>
      </w:p>
    `);
    const styles = stylesXml(`
      <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
        <w:pPr><w:pBdr><w:top w:val="single" w:color="112233" w:sz="8" w:space="1"/></w:pBdr></w:pPr>
      </w:style>
      <w:style w:type="paragraph" w:styleId="Derived">
        <w:basedOn w:val="Normal"/>
        <w:pPr><w:pBdr>
          <w:left w:val="double" w:themeColor="accent6" w:themeTint="80" w:themeShade="40" w:sz="12" w:space="2" w:shadow="1" w:frame="0"/>
          <w:bottom w:val="dotted" w:color="auto" w:sz="6"/>
          <w:right w:val="apples" w:color="AABBCC" w:sz="20"/>
          <w:between w:val="wave" w:color="445566" w:sz="10" w:space="31"/>
          <w:bar w:val="threeDEmboss" w:color="778899" w:sz="14"/>
        </w:pBdr></w:pPr>
      </w:style>
    `);
    const theme = themeXml(`
      <a:themeElements><a:clrScheme name="Office">
        <a:accent6><a:srgbClr val="70AD47"/></a:accent6>
      </a:clrScheme></a:themeElements>
    `);

    const markers = markDocxParagraphBorders(document, styles, theme);

    expect(markers).toMatchObject({ invalidCount: 0, spoofedCount: 0 });
    expect(markers.paragraphs).toHaveLength(1);
    expect(markers.paragraphs[0]?.borders).toEqual({
      top: {
        style: 'single',
        color: { value: '#112233' },
        size: 8,
        space: 1,
      },
      left: {
        style: 'double',
        color: {
          value: '#b8d6a3',
          theme: {
            theme: 'accent6',
            resolved: '#b8d6a3',
            tint: '80',
            shade: '40',
          },
        },
        size: 12,
        space: 2,
        shadow: true,
        frame: false,
      },
      bottom: { style: 'dotted', color: { value: 'auto' }, size: 6 },
      right: { style: 'apples', color: { value: '#aabbcc' }, size: 20 },
      between: {
        style: 'wave',
        color: { value: '#445566' },
        size: 10,
        space: 31,
      },
      bar: {
        style: 'threeDEmboss',
        color: { value: '#778899' },
        size: 14,
      },
    });

    const marker = markers.paragraphs[0];
    if (!marker) throw new Error('Expected a paragraph border marker.');
    const html = new DOMParser().parseFromString(
      `<p>${marker.marker}Themed borders</p>`,
      'text/html',
    );
    applyImportedDocxParagraphBorderMarkers(html, markers);
    const paragraph = html.body.querySelector<HTMLElement>('p');
    expect(paragraph?.style.borderLeftStyle).toBe('double');
    expect(paragraph?.style.borderRightStyle).toBe('solid');
    expect(paragraph?.style.boxShadow).toContain('inset');
    expect(
      parseDocumentParagraphBorders(paragraph?.dataset.officeParagraphBorders),
    ).toEqual(marker.borders);
    expect(html.body.textContent).toBe('Themed borders');

    const controlled = createWorkDocumentModelFromContent({
      type: 'document',
      pageSize: 'a4',
      html: html.body.innerHTML,
    });
    expect(
      collectNodes(controlled.model?.root, 'paragraph')[0]?.attrs
        ?.paragraphBorders,
    ).toEqual(marker.borders);

    const blob = await createDocxBlob({
      type: 'document',
      pageSize: 'a4',
      html: html.body.innerHTML,
    });
    const archive = await JSZip.loadAsync(await blob.arrayBuffer());
    const compatibility = await analyzeDocxCompatibility(
      new File([blob], 'paragraph-borders.docx', { type: blob.type }),
      [],
    );
    expect(
      compatibility.issues.find(
        (issue) => issue.code === 'docx.paragraph-borders',
      ),
    ).toMatchObject({ severity: 'info' });
    const exported = parseXml(
      (await archive.file('word/document.xml')?.async('text')) ?? '',
    );
    const container = paragraphBorderContainers(exported)[0];
    if (!container) throw new Error('Expected exported paragraph borders.');
    expect(
      directChildren(container).map((element) => element.localName),
    ).toEqual(DOCUMENT_PARAGRAPH_BORDER_EDGES);
    const exportedLeft = directChildren(container, 'left')[0];
    if (!exportedLeft) throw new Error('Expected exported left border.');
    expect(borderAttributes(exportedLeft)).toEqual({
      val: 'double',
      color: 'B8D6A3',
      themeColor: 'accent6',
      themeTint: '80',
      themeShade: '40',
      sz: '12',
      space: '2',
      shadow: '1',
      frame: '0',
    });
    expect(attribute(requiredChild(container, 'right'), 'val')).toBe('apples');
    expect(attribute(requiredChild(container, 'between'), 'space')).toBe('31');
    expect(attribute(requiredChild(container, 'bar'), 'val')).toBe(
      'threeDEmboss',
    );

    if (!paragraph) throw new Error('Expected imported paragraph HTML.');
    paragraph.style.borderLeftColor = '#123456';
    const editedBlob = await createDocxBlob({
      type: 'document',
      pageSize: 'a4',
      html: html.body.innerHTML,
    });
    const editedArchive = await JSZip.loadAsync(await editedBlob.arrayBuffer());
    const edited = parseXml(
      (await editedArchive.file('word/document.xml')?.async('text')) ?? '',
    );
    const editedContainer = paragraphBorderContainers(edited)[0];
    if (!editedContainer) throw new Error('Expected edited paragraph borders.');
    const editedLeft = directChildren(editedContainer, 'left')[0];
    if (!editedLeft) throw new Error('Expected edited left border.');
    expect(attribute(editedLeft, 'color')).toBe('123456');
    expect(attribute(editedLeft, 'themeColor')).toBeNull();
    expect(attribute(editedLeft, 'val')).toBe('double');
  });

  test('applies conditional table borders before direct nil resets', () => {
    const stylesDocument = stylesXml(`
      <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
        <w:pPr><w:pBdr><w:top w:val="single" w:color="112233" w:sz="8"/></w:pBdr></w:pPr>
      </w:style>
      <w:style w:type="table" w:styleId="Report">
        <w:tblStylePr w:type="wholeTable"><w:pPr><w:pBdr><w:left w:val="dashed" w:color="4472C4" w:sz="6"/></w:pBdr></w:pPr></w:tblStylePr>
        <w:tblStylePr w:type="firstRow"><w:pPr><w:pBdr><w:bottom w:val="double" w:color="C00000" w:sz="12"/></w:pBdr></w:pPr></w:tblStylePr>
      </w:style>
    `);
    const document = wordXml(`
      <w:tbl>
        <w:tblPr><w:tblStyle w:val="Report"/><w:tblLook w:firstRow="1" w:noHBand="1" w:noVBand="1"/></w:tblPr>
        <w:tr><w:tc><w:p><w:r><w:t>Header</w:t></w:r></w:p></w:tc></w:tr>
        <w:tr><w:tc><w:p><w:pPr><w:pBdr><w:top w:val="nil"/><w:left w:val="nil"/></w:pBdr></w:pPr><w:r><w:t>Reset</w:t></w:r></w:p></w:tc></w:tr>
      </w:tbl>
    `);

    const markers = markDocxParagraphBorders(
      document,
      createDocxParagraphStyleResolver(stylesDocument),
      undefined,
      createDocxTableStyleResolver(stylesDocument),
    );

    expect(markers.paragraphs.map(({ borders }) => borders)).toEqual([
      {
        top: { style: 'single', color: { value: '#112233' }, size: 8 },
        left: { style: 'dashed', color: { value: '#4472c4' }, size: 6 },
        bottom: { style: 'double', color: { value: '#c00000' }, size: 12 },
      },
      { top: { style: 'nil' }, left: { style: 'nil' } },
    ]);
  });

  test('accepts every schema border style and strict WordprocessingML', () => {
    expect(DOCUMENT_PARAGRAPH_BORDER_STYLES).toHaveLength(197);
    for (const style of DOCUMENT_PARAGRAPH_BORDER_STYLES) {
      const size =
        style === 'nil' || style === 'none'
          ? 0
          : styleIndex(style) < 27
            ? 8
            : 20;
      const properties = descendants(
        strictWordXml(
          `<w:p><w:pPr><w:pBdr><w:top w:val="${style}" w:color="112233" w:sz="${size}" w:space="31" w:shadow="true" w:frame="false"/></w:pBdr></w:pPr></w:p>`,
        ),
        'pPr',
      )[0];
      if (!properties) throw new Error(`Expected strict ${style} properties.`);
      expect(parseDirectDocxParagraphBorders(properties)).toEqual({
        top: {
          style,
          color: { value: '#112233' },
          size,
          space: 31,
          shadow: true,
          frame: false,
        },
      });
    }
  });

  test('fails closed for malformed containers and ignores namespace spoofing', () => {
    const document = wordXml(`
      <w:p><w:pPr><w:pStyle w:val="Bordered"/><w:pBdr><w:left w:val="single"/><w:top w:val="single"/></w:pBdr></w:pPr><w:r><w:t>Order</w:t></w:r></w:p>
      <w:p><w:pPr><w:pBdr><w:top w:val="single"/><w:top w:val="double"/></w:pBdr></w:pPr><w:r><w:t>Duplicate edge</w:t></w:r></w:p>
      <w:p><w:pPr><w:pBdr/><w:pBdr/></w:pPr><w:r><w:t>Duplicate container</w:t></w:r></w:p>
      <w:p><w:pPr><w:pBdr><w:top w:val="unknown"/></w:pBdr></w:pPr><w:r><w:t>Unknown</w:t></w:r></w:p>
      <w:p><w:pPr><w:pBdr><w:top xmlns:r="${RELATIONSHIP_NAMESPACE}" w:val="single" r:id="rIdUnsafe"/></w:pBdr></w:pPr><w:r><w:t>Relationship</w:t></w:r></w:p>
      <w:p xmlns:x="urn:spoof"><w:pPr><w:pStyle w:val="Bordered"/><w:pBdr><x:top w:val="double"/></w:pBdr></w:pPr><w:r><w:t>Spoof</w:t></w:r></w:p>
    `);
    const styles = stylesXml(`
      <w:style w:type="paragraph" w:styleId="Bordered"><w:pPr><w:pBdr><w:top w:val="single" w:color="4472C4" w:sz="8"/></w:pBdr></w:pPr></w:style>
    `);

    const markers = markDocxParagraphBorders(document, styles);

    expect(markers.invalidCount).toBe(6);
    expect(markers.spoofedCount).toBe(1);
    expect(markers.paragraphs).toHaveLength(6);
    expect(
      markers.paragraphs
        .slice(0, 3)
        .every(({ borders }) =>
          DOCUMENT_PARAGRAPH_BORDER_EDGES.every(
            (edge) => borders[edge]?.style === 'nil',
          ),
        ),
    ).toBe(true);
    expect(
      markers.paragraphs.slice(3, 5).map(({ borders }) => borders),
    ).toEqual([{ top: { style: 'nil' } }, { top: { style: 'nil' } }]);
    expect(markers.paragraphs[5]?.borders.top).toEqual({
      style: 'single',
      color: { value: '#4472c4' },
      size: 8,
    });
  });

  test('follows custom theme mappings and gives tint precedence over shade', () => {
    const theme = themeXml(`
      <a:themeElements><a:clrScheme name="Custom">
        <a:dk2><a:srgbClr val="112233"/></a:dk2>
        <a:accent6><a:srgbClr val="70AD47"/></a:accent6>
      </a:clrScheme></a:themeElements>
    `);
    const resolver = createDocxThemeResolver(
      theme,
      settingsXml('<w:clrSchemeMapping w:bg1="accent6" w:accent6="dark2"/>'),
    );
    const properties = descendants(
      wordXml(
        '<w:p><w:pPr><w:pBdr><w:top w:val="single" w:themeColor="background1" w:themeTint="80" w:themeShade="40" w:sz="8"/></w:pBdr></w:pPr></w:p>',
      ),
      'pPr',
    )[0];
    if (!properties) throw new Error('Expected mapped border properties.');

    expect(parseDirectDocxParagraphBorders(properties, resolver)).toEqual({
      top: {
        style: 'single',
        color: {
          value: '#889199',
          theme: {
            theme: 'background1',
            resolved: '#889199',
            tint: '80',
            shade: '40',
          },
        },
        size: 8,
      },
    });

    const noThemeProperties = descendants(
      wordXml(
        '<w:p><w:pPr><w:pBdr><w:top w:val="single" w:themeColor="none" w:sz="8"/></w:pBdr></w:pPr></w:p>',
      ),
      'pPr',
    )[0];
    if (!noThemeProperties)
      throw new Error('Expected explicit no-theme border properties.');
    expect(
      parseDirectDocxParagraphBorders(noThemeProperties, resolver),
    ).toEqual({
      top: {
        style: 'single',
        color: {
          value: 'auto',
          theme: {
            theme: 'none',
            resolved: '#000000',
          },
        },
        size: 8,
      },
    });
  });

  test('round-trips page-chrome borders through sanitization and its editor', async () => {
    const attributes = documentParagraphBordersDomAttributes({
      top: { style: 'double', color: { value: '#112233' }, size: 12 },
      between: { style: 'wave', color: { value: '#445566' }, size: 8 },
      bar: { style: 'apples', color: { value: '#778899' }, size: 20 },
    });
    const content = {
      type: 'document' as const,
      html: '<p>Body</p>',
      pageSize: 'a4' as const,
      pageChrome: {
        differentFirstPage: false,
        differentOddEvenPages: false,
        default: {
          headerHtml: `<p data-office-paragraph-borders='${attributes['data-office-paragraph-borders']}' style="${attributes.style}">Bordered header</p>`,
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
    const exported = paragraphBorderContainers(header)[0];
    if (!exported) throw new Error('Expected exported header borders.');
    expect(
      directChildren(exported).map((element) => element.localName),
    ).toEqual(['top', 'between', 'bar']);

    const imported = await importOfficeFile(
      new File([blob], 'paragraph-borders.docx', { type: blob.type }),
    );
    if (imported.content.type !== 'document')
      throw new Error('Expected an imported document artifact.');
    const headerHtml = imported.content.pageChrome?.default.headerHtml ?? '';
    expect(headerHtml).toContain('data-office-paragraph-borders');
    expect(headerHtml).toContain('between');
    expect(headerHtml).toContain('box-shadow');

    const editor = new Editor({
      extensions: createDocumentPageChromeEditorExtensions(),
      content: headerHtml,
    });
    expect(editor.getJSON().content?.[0]?.attrs?.paragraphBorders).toEqual({
      top: { style: 'double', color: { value: '#112233' }, size: 12 },
      between: { style: 'wave', color: { value: '#445566' }, size: 8 },
      bar: { style: 'apples', color: { value: '#778899' }, size: 20 },
    });
    expect(editor.getHTML()).toContain('data-office-paragraph-borders');
    editor.destroy();

    const roundTripped = await createDocxBlob(imported.content);
    const secondArchive = await JSZip.loadAsync(
      await roundTripped.arrayBuffer(),
    );
    const secondHeaderPath = Object.keys(secondArchive.files).find((path) =>
      /^word\/header\d*\.xml$/i.test(path),
    );
    if (!secondHeaderPath) throw new Error('Expected a round-tripped header.');
    const secondHeader = parseXml(
      (await secondArchive.file(secondHeaderPath)?.async('text')) ?? '',
    );
    const secondBorders = paragraphBorderContainers(secondHeader)[0];
    if (!secondBorders)
      throw new Error('Expected round-tripped header borders.');
    expect(
      directChildren(secondBorders).map((element) => element.localName),
    ).toEqual(['top', 'between', 'bar']);
  });

  test('exports list-item paragraph borders without losing numbering', async () => {
    const attributes = documentParagraphBordersDomAttributes({
      bottom: { style: 'dashDotStroked', color: { value: '#c00000' }, size: 8 },
      bar: { style: 'doubleWave', color: { value: '#4472c4' }, size: 12 },
    });
    const blob = await createDocxBlob({
      type: 'document',
      pageSize: 'a4',
      html: `<ol><li><p data-office-paragraph-borders='${attributes['data-office-paragraph-borders']}' style="${attributes.style}">Bordered item</p></li></ol>`,
    });
    const archive = await JSZip.loadAsync(await blob.arrayBuffer());
    const document = parseXml(
      (await archive.file('word/document.xml')?.async('text')) ?? '',
    );
    const paragraph = descendants(document, 'p').find((element) =>
      element.textContent?.includes('Bordered item'),
    );
    if (!paragraph) throw new Error('Expected exported list paragraph.');
    const properties = directChildren(paragraph, 'pPr')[0];
    if (!properties) throw new Error('Expected list paragraph properties.');
    expect(descendants(properties, 'numPr')).toHaveLength(1);
    const borders = directChildren(properties, 'pBdr')[0];
    if (!borders) throw new Error('Expected list paragraph borders.');
    expect(directChildren(borders).map((element) => element.localName)).toEqual(
      ['bottom', 'bar'],
    );
    expect(attribute(requiredChild(borders, 'bar'), 'val')).toBe('doubleWave');
  });

  test('loads theme-mapped paragraph borders at the DOCX package boundary', async () => {
    const archive = new JSZip();
    archive.file(
      'word/document.xml',
      serializeXml(
        wordXml(
          '<w:p><w:pPr><w:pBdr><w:bar w:val="apples" w:themeColor="background1" w:sz="20"/></w:pBdr></w:pPr><w:r><w:t>Package</w:t></w:r></w:p>',
        ),
      ),
    );
    archive.file(
      'word/theme/theme1.xml',
      serializeXml(
        themeXml(
          '<a:themeElements><a:clrScheme name="Custom"><a:dk2><a:srgbClr val="DDEEFF"/></a:dk2></a:clrScheme></a:themeElements>',
        ),
      ),
    );
    archive.file(
      'word/settings.xml',
      serializeXml(settingsXml('<w:clrSchemeMapping w:bg1="dark2"/>')),
    );

    const prepared = await prepareDocxImport(
      await archive.generateAsync({ type: 'arraybuffer' }),
    );

    expect(prepared.paragraphBorderMarkers.paragraphs[0]?.borders).toEqual({
      bar: {
        style: 'apples',
        color: {
          value: '#ddeeff',
          theme: { theme: 'background1', resolved: '#ddeeff' },
        },
        size: 20,
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

function paragraphBorderContainers(document: Document): Element[] {
  return descendants(document, 'pBdr').filter(
    (element) => element.parentElement?.localName === 'pPr',
  );
}

function requiredChild(parent: ParentNode, name: string): Element {
  const child = directChildren(parent, name)[0];
  if (!child) throw new Error(`Expected ${name} child.`);
  return child;
}

function borderAttributes(element: Element): Record<string, string> {
  return Object.fromEntries(
    [
      'val',
      'color',
      'themeColor',
      'themeTint',
      'themeShade',
      'sz',
      'space',
      'shadow',
      'frame',
    ].flatMap((name) => {
      const value = attribute(element, name);
      return value === null ? [] : [[name, value]];
    }),
  );
}

function styleIndex(style: string): number {
  return DOCUMENT_PARAGRAPH_BORDER_STYLES.indexOf(
    style as (typeof DOCUMENT_PARAGRAPH_BORDER_STYLES)[number],
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
