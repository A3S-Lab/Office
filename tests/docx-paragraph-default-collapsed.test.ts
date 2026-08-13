import { Editor } from '@tiptap/core';
import StarterKit from '@tiptap/starter-kit';
import { describe, expect, test } from '@rstest/core';
import JSZip from 'jszip';
import { importOfficeFile } from '../src/core';
import { createDocumentPageChromeEditorExtensions } from '../src/internal/features/work/editors/document-page-chrome-editor';
import { DocumentParagraphFormatting } from '../src/internal/features/work/work-document-paragraph-formatting';
import {
  DocxParagraphDefaultCollapsedPatchCollector,
  parseDocxParagraphDefaultCollapsed,
} from '../src/internal/features/work/work-docx-paragraph-default-collapsed';
import { createDocxBlob } from '../src/internal/features/work/work-docx-export';
import { analyzeDocxCompatibility } from '../src/internal/features/work/work-office-diagnostics';
import {
  applyImportedDocxParagraphPaginationMarkers,
  markDocxParagraphPagination,
} from '../src/internal/features/work/work-docx-paragraph-pagination-import';
import {
  attribute,
  descendants,
  directChild,
  directChildren,
  parseXml,
} from '../src/internal/features/work/work-ooxml-package';
import {
  xmlAttributeLocalName,
  xmlAttributeNamespace,
  xmlNamespaceUri,
} from '../src/internal/features/work/work-docx-settings-xml';

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const STRICT_WORD_NAMESPACE =
  'http://purl.oclc.org/ooxml/wordprocessingml/main';
const WORD_2012_NAMESPACE =
  'http://schemas.microsoft.com/office/word/2012/wordml';
const MARKUP_COMPATIBILITY_NAMESPACE =
  'http://schemas.openxmlformats.org/markup-compatibility/2006';
const RELATIONSHIPS_NAMESPACE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const VENDOR_NAMESPACE = 'urn:a3s:test:default-collapsed';

describe('DOCX Office 2013 default-collapsed paragraphs', () => {
  test('accepts the exact core CT_OnOff lexical forms and style inheritance', () => {
    const cases = [
      ['<w15:collapsed/>', true],
      ['<w15:collapsed w:val="true"/>', true],
      ['<w15:collapsed w:val="on"/>', true],
      ['<w15:collapsed w:val="1"/>', true],
      ['<w15:collapsed w:val="false"/>', false],
      ['<w15:collapsed w:val="off"/>', false],
      ['<w15:collapsed w:val="0"/>', false],
      ['<w15:collapsed> \n </w15:collapsed>', true],
      ['<w15:collapsed ws:val="0"/>', false],
    ] as const;
    const document = wordXml(
      [
        ...cases.map(
          ([markup], index) =>
            `<w:p><w:pPr>${markup}</w:pPr><w:r><w:t>Case ${index}</w:t></w:r></w:p>`,
        ),
        '<w:p><w:pPr><w:pStyle w:val="Derived"/></w:pPr><w:r><w:t>Inherited</w:t></w:r></w:p>',
      ].join(''),
    );
    const styles = stylesXml(`
      <w:style w:type="paragraph" w:styleId="Base">
        <w:pPr><w15:collapsed w:val="on"/></w:pPr>
      </w:style>
      <w:style w:type="paragraph" w:styleId="Derived">
        <w:basedOn w:val="Base"/>
      </w:style>
    `);

    const markers = markDocxParagraphPagination(document, styles);

    expect(
      markers.paragraphs.map(({ pagination }) => pagination.defaultCollapsed),
    ).toEqual([...cases.map(([, value]) => value), true]);
    const html = new DOMParser().parseFromString(
      markers.paragraphs
        .map(
          ({ marker }, index) =>
            `<h1>${marker}${index === cases.length ? 'Inherited' : `Case ${index}`}</h1>`,
        )
        .join(''),
      'text/html',
    );
    applyImportedDocxParagraphPaginationMarkers(html, markers);
    expect(
      Array.from(html.body.children).map(
        (element) => (element as HTMLElement).dataset.officeDefaultCollapsed,
      ),
    ).toEqual([...cases.map(([, value]) => String(value)), 'true']);
    expect(html.body.textContent).not.toContain('__A3S_');

    const strictProperties = parseXml(
      `<ws:pPr xmlns:ws="${STRICT_WORD_NAMESPACE}" xmlns:w15="${WORD_2012_NAMESPACE}"><w15:collapsed ws:val="off"/></ws:pPr>`,
    ).documentElement;
    expect(parseDocxParagraphDefaultCollapsed(strictProperties)).toEqual({
      status: 'valid',
      value: false,
    });
  });

  test('fails closed for malformed direct values, duplicates, and namespace spoofing', () => {
    const invalid = [
      '<w15:collapsed/><w15:collapsed w:val="0"/>',
      '<w15:collapsed val="1"/>',
      '<w15:collapsed w15:val="1"/>',
      '<w15:collapsed w:val=""/>',
      '<w15:collapsed w:val="True"/>',
      '<w15:collapsed w:val="FALSE"/>',
      '<w15:collapsed w:val=" true"/>',
      '<w15:collapsed w:val="1 "/>',
      '<w15:collapsed w:val="yes"/>',
      '<w15:collapsed w:val="2"/>',
      '<w15:collapsed w:val="1" v:extra="unsafe"/>',
      '<w15:collapsed w:val="1" r:id="rIdUnsafe"/>',
      '<w15:collapsed>text</w15:collapsed>',
      '<w15:collapsed><w15:child/></w15:collapsed>',
    ];
    const document = wordXml(
      invalid
        .map(
          (markup, index) => `
            <w:p>
              <w:pPr><w:pStyle w:val="Base"/>${markup}</w:pPr>
              <w:r><w:t>Invalid ${index}</w:t></w:r>
            </w:p>`,
        )
        .join(''),
    );
    const styles = stylesXml(`
      <w:style w:type="paragraph" w:styleId="Base">
        <w:pPr><w15:collapsed w:val="1"/></w:pPr>
      </w:style>
    `);

    expect(markDocxParagraphPagination(document, styles).paragraphs).toEqual(
      [],
    );

    const spoofed = wordXml(`
      <w:p><w:pPr><v:collapsed v:val="1"/></w:pPr><w:r><w:t>Vendor</w:t></w:r></w:p>
      <w:p><w:pPr><w:collapsed w:val="1"/></w:pPr><w:r><w:t>Core</w:t></w:r></w:p>
      <w:p><w:pPr><w14:collapsed xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"/></w:pPr><w:r><w:t>Office 2010</w:t></w:r></w:p>
      <w:p><w:r><w:rPr><w15:collapsed/></w:rPr><w:t>Misplaced run property</w:t></w:r></w:p>
    `);
    expect(markDocxParagraphPagination(spoofed).paragraphs).toEqual([]);

    const collector = new DocxParagraphDefaultCollapsedPatchCollector();
    const html = new DOMParser().parseFromString(
      '<h1 data-office-default-collapsed="on">Malformed HTML state</h1>',
      'text/html',
    );
    collector.register(
      '__A3S_TEST_MARKER__',
      html.body.firstElementChild as HTMLElement,
    );
    expect(collector.patches).toEqual([]);
  });

  test('round-trips canonical body and page-chrome metadata without hiding browser content', async () => {
    const content = {
      type: 'document' as const,
      html: [
        '<h1 data-office-default-collapsed="true">Collapsed heading</h1>',
        '<h2 data-office-default-collapsed="false">Expanded heading</h2>',
        '<h6 data-office-default-collapsed="true">Deep collapsed heading</h6>',
        '<h3 data-office-default-collapsed="on">Malformed state</h3>',
        '<p>Visible body text</p>',
      ].join(''),
      pageSize: 'a4' as const,
      pageChrome: {
        differentFirstPage: false,
        differentOddEvenPages: false,
        default: {
          headerHtml:
            '<p data-office-default-collapsed="true">Header metadata</p>',
          footerHtml: '',
          showPageNumber: false,
        },
        first: { headerHtml: '', footerHtml: '', showPageNumber: false },
        even: { headerHtml: '', footerHtml: '', showPageNumber: false },
      },
    };

    const blob = await createDocxBlob(content);
    const compatibility = await analyzeDocxCompatibility(
      new File([blob], 'default-collapsed.docx', { type: blob.type }),
      [],
    );
    expect(
      compatibility.issues.find(
        (issue) => issue.code === 'docx.headings.default-collapsed',
      ),
    ).toMatchObject({ severity: 'info' });
    const archive = await JSZip.loadAsync(await blob.arrayBuffer());
    const document = await xmlEntry(archive, 'word/document.xml');
    assertIgnorableWord2012(document.documentElement);
    assertDefaultCollapsed(paragraphByText(document, 'Collapsed heading'), '1');
    assertDefaultCollapsed(paragraphByText(document, 'Expanded heading'), '0');
    const deepHeading = paragraphByText(document, 'Deep collapsed heading');
    assertDefaultCollapsed(deepHeading, '1');
    expect(
      attribute(
        directChild(directChild(deepHeading, 'pPr') ?? deepHeading, 'pStyle') ??
          deepHeading,
        'val',
      ),
    ).toBe('Heading6');
    expect(
      defaultCollapsedElements(paragraphByText(document, 'Malformed state')),
    ).toEqual([]);

    const headerPath = Object.keys(archive.files).find((path) =>
      /^word\/header\d*\.xml$/i.test(path),
    );
    expect(headerPath).toBeDefined();
    const header = await xmlEntry(archive, headerPath as string);
    assertIgnorableWord2012(header.documentElement);
    assertDefaultCollapsed(paragraphByText(header, 'Header metadata'), '1');

    const imported = await importOfficeFile(
      new File([blob], 'default-collapsed.docx', { type: blob.type }),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const importedBody = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    expect(
      htmlParagraphByText(importedBody, 'Collapsed heading')?.dataset
        .officeDefaultCollapsed,
    ).toBe('true');
    expect(
      htmlParagraphByText(importedBody, 'Expanded heading')?.dataset
        .officeDefaultCollapsed,
    ).toBe('false');
    expect(
      htmlParagraphByText(importedBody, 'Deep collapsed heading')?.dataset
        .officeDefaultCollapsed,
    ).toBe('true');
    expect(
      htmlParagraphByText(importedBody, 'Malformed state')?.dataset
        .officeDefaultCollapsed,
    ).toBeUndefined();
    expect(imported.content.html).toContain('Visible body text');
    expect(imported.content.pageChrome?.default.headerHtml).toContain(
      'data-office-default-collapsed="true"',
    );

    const pageChromeEditor = new Editor({
      extensions: createDocumentPageChromeEditorExtensions(),
      content: imported.content.pageChrome?.default.headerHtml,
    });
    expect(pageChromeEditor.getJSON().content?.[0]?.attrs).toMatchObject({
      defaultCollapsed: true,
    });
    expect(pageChromeEditor.getHTML()).toContain(
      'data-office-default-collapsed="true"',
    );
    pageChromeEditor.destroy();

    const editor = new Editor({
      extensions: [StarterKit, DocumentParagraphFormatting],
      content: imported.content.html,
    });
    expect(editor.getJSON().content?.[0]?.attrs).toMatchObject({
      defaultCollapsed: true,
    });
    expect(editor.getJSON().content?.[1]?.attrs).toMatchObject({
      defaultCollapsed: false,
    });
    expect(editor.getJSON().content?.[2]).toMatchObject({
      type: 'heading',
      attrs: { defaultCollapsed: true, level: 6 },
    });
    expect(editor.getHTML()).toContain('Collapsed heading');
    expect(editor.getHTML()).toContain('data-office-default-collapsed="true"');
    expect(editor.getHTML()).toContain('data-office-default-collapsed="false"');
    editor.destroy();

    const roundTripped = await createDocxBlob(imported.content);
    const secondArchive = await JSZip.loadAsync(
      await roundTripped.arrayBuffer(),
    );
    const secondDocument = await xmlEntry(secondArchive, 'word/document.xml');
    assertDefaultCollapsed(
      paragraphByText(secondDocument, 'Collapsed heading'),
      '1',
    );
    assertDefaultCollapsed(
      paragraphByText(secondDocument, 'Expanded heading'),
      '0',
    );
    assertDefaultCollapsed(
      paragraphByText(secondDocument, 'Deep collapsed heading'),
      '1',
    );
  });
});

function wordXml(body: string): Document {
  return parseXml(
    `<w:document xmlns:w="${WORD_NAMESPACE}" xmlns:ws="${STRICT_WORD_NAMESPACE}" xmlns:w15="${WORD_2012_NAMESPACE}" xmlns:r="${RELATIONSHIPS_NAMESPACE}" xmlns:v="${VENDOR_NAMESPACE}"><w:body>${body}</w:body></w:document>`,
  );
}

function stylesXml(body: string): Document {
  return parseXml(
    `<w:styles xmlns:w="${WORD_NAMESPACE}" xmlns:w15="${WORD_2012_NAMESPACE}">${body}</w:styles>`,
  );
}

async function xmlEntry(archive: JSZip, path: string): Promise<Document> {
  const source = await archive.file(path)?.async('text');
  if (!source) throw new Error(`Missing DOCX XML part: ${path}`);
  return parseXml(source, path);
}

function paragraphByText(document: Document, text: string): Element {
  const paragraph = descendants(document, 'p').find(
    (element) => element.textContent === text,
  );
  if (!paragraph) throw new Error(`Missing paragraph: ${text}`);
  return paragraph;
}

function htmlParagraphByText(
  document: Document,
  text: string,
): HTMLElement | undefined {
  return Array.from(
    document.body.querySelectorAll<HTMLElement>('p, h1, h2, h3, h4, h5, h6'),
  ).find((element) => element.textContent === text);
}

function defaultCollapsedElements(paragraph: Element): Element[] {
  const properties = directChildren(paragraph, 'pPr').find(
    (element) => element.namespaceURI === WORD_NAMESPACE,
  );
  return properties
    ? directChildren(properties, 'collapsed').filter(
        (element) => element.namespaceURI === WORD_2012_NAMESPACE,
      )
    : [];
}

function assertDefaultCollapsed(paragraph: Element, value: '0' | '1'): void {
  const collapsed = defaultCollapsedElements(paragraph);
  expect(collapsed).toHaveLength(1);
  expect(directChildren(collapsed[0])).toEqual([]);
  expect(
    Array.from(collapsed[0].attributes)
      .filter(
        (item) =>
          xmlAttributeNamespace(collapsed[0], item) !==
          'http://www.w3.org/2000/xmlns/',
      )
      .map((item) => ({
        localName: xmlAttributeLocalName(item),
        namespace: xmlAttributeNamespace(collapsed[0], item),
        value: item.value,
      })),
  ).toEqual([
    {
      localName: 'val',
      namespace: WORD_NAMESPACE,
      value,
    },
  ]);
  const properties = directChild(paragraph, 'pPr');
  expect(directChildren(properties as Element).at(-1)).toBe(collapsed[0]);
}

function assertIgnorableWord2012(root: Element): void {
  const ignorable = Array.from(root.attributes).find(
    (item) =>
      xmlAttributeLocalName(item) === 'Ignorable' &&
      xmlAttributeNamespace(root, item) === MARKUP_COMPATIBILITY_NAMESPACE,
  );
  expect(ignorable).toBeDefined();
  const namespaceBindings = (ignorable?.value ?? '')
    .trim()
    .split(/\s+/)
    .map((prefix) => [prefix, xmlNamespaceUri(root, prefix)]);
  if (
    !namespaceBindings.some(
      ([, namespace]) => namespace === WORD_2012_NAMESPACE,
    )
  ) {
    throw new Error(
      `Missing ignorable Word 2012 namespace: ${new XMLSerializer().serializeToString(root)}`,
    );
  }
}
