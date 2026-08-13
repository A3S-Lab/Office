import { describe, expect, test } from '@rstest/core';
import { render } from '@testing-library/react';
import JSZip from 'jszip';
import { WorkDocumentPdfPages } from '../src/internal/features/work/components/work-document-pages';
import { DocumentPageStack } from '../src/internal/features/work/editors/document-page-stack';
import {
  DOCUMENT_PAGE_BORDER_EDGES,
  documentPageBordersVisible,
  normalizeDocumentPageBorders,
  parseDocumentPageBorders,
  resolveDocumentPageBorders,
  serializeDocumentPageBorders,
  type WorkDocumentPageBorders,
} from '../src/internal/features/work/work-document-page-borders';
import {
  DOCUMENT_PARAGRAPH_BORDER_STYLES,
  type DocumentParagraphBorderStyle,
} from '../src/internal/features/work/work-document-paragraph-borders';
import {
  documentInitialSectionLayout,
  documentSectionDomAttributes,
} from '../src/internal/features/work/work-document-section';
import { createDocxBlob } from '../src/internal/features/work/work-docx-export';
import { prepareDocxImport } from '../src/internal/features/work/work-docx-import';
import {
  inspectDocxPageBorders,
  parseDocxPageBorders,
} from '../src/internal/features/work/work-docx-page-borders-import';
import { createDocxThemeResolver } from '../src/internal/features/work/work-docx-theme';
import { analyzeDocxCompatibility } from '../src/internal/features/work/work-office-diagnostics';
import {
  attribute,
  descendants,
  directChildren,
  parseXml,
} from '../src/internal/features/work/work-ooxml-package';
import type {
  WorkDocumentContent,
  WorkDocumentSectionLayout,
} from '../src/internal/features/work/work-types';

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const STRICT_WORD_NAMESPACE =
  'http://purl.oclc.org/ooxml/wordprocessingml/main';
const DRAWING_NAMESPACE =
  'http://schemas.openxmlformats.org/drawingml/2006/main';
const RELATIONSHIP_NAMESPACE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

describe('DOCX page borders', () => {
  test('normalizes native page-border semantics and resolves Word defaults', () => {
    const source: WorkDocumentPageBorders = {
      edges: {
        top: {
          style: 'double',
          color: { value: '#112233' },
          size: 12,
          space: 3,
          shadow: true,
        },
        right: { style: 'nil' },
      },
    };

    const serialized = serializeDocumentPageBorders(source);
    expect(serialized).toBe(
      '{"edges":{"top":{"style":"double","color":{"value":"#112233"},"size":12,"space":3,"shadow":true},"right":{"style":"nil"}}}',
    );
    expect(parseDocumentPageBorders(serialized)).toEqual(source);
    expect(
      normalizeDocumentPageBorders({ ...source, externalRelationship: 'rId1' }),
    ).toBeNull();
    expect(
      normalizeDocumentPageBorders({
        edges: { diagonal: { style: 'single' } },
      }),
    ).toBeNull();

    const resolved = resolveDocumentPageBorders(source, {
      top: 25,
      right: 23,
      bottom: 25,
      left: 23,
    });
    expect(resolved).toMatchObject({
      display: 'allPages',
      offsetFrom: 'text',
      zOrder: 'front',
      edges: {
        top: { color: '#112233', style: 'double', width: 2 },
        right: { color: 'transparent', style: 'none', width: 0 },
      },
    });
    expect(resolved?.insets.top).toBeCloseTo(88.488, 3);
    expect(documentPageBordersVisible(source, 0)).toBe(true);
    expect(
      documentPageBordersVisible({ ...source, display: 'firstPage' }, 1),
    ).toBe(true);
    expect(
      documentPageBordersVisible({ ...source, display: 'firstPage' }, 2),
    ).toBe(false);
    expect(
      documentPageBordersVisible({ ...source, display: 'notFirstPage' }, 2),
    ).toBe(true);
  });

  test('accepts every schema border style and strict WordprocessingML', () => {
    expect(DOCUMENT_PARAGRAPH_BORDER_STYLES).toHaveLength(197);
    for (const style of DOCUMENT_PARAGRAPH_BORDER_STYLES) {
      const size = pageBorderSize(style);
      const section = requiredDescendant(
        strictWordXml(
          `<w:sectPr><w:pgBorders w:display="notFirstPage" w:offsetFrom="page" w:zOrder="back"><w:top w:val="${style}" w:color="112233" w:sz="${size}" w:space="31" w:shadow="true" w:frame="false"/></w:pgBorders></w:sectPr>`,
        ),
        'sectPr',
      );
      expect(parseDocxPageBorders(section)).toEqual({
        display: 'notFirstPage',
        offsetFrom: 'page',
        zOrder: 'back',
        edges: {
          top: {
            style,
            color: { value: '#112233' },
            size,
            space: 31,
            shadow: true,
            frame: false,
          },
        },
      });
    }
  });

  test('imports theme-mapped borders per section without stale inheritance', async () => {
    const archive = new JSZip();
    archive.file(
      'word/document.xml',
      serializeXml(
        wordXml(`
          <w:p><w:pPr><w:sectPr>
            <w:pgBorders w:display="firstPage" w:offsetFrom="text">
              <w:top w:val="doubleWave" w:themeColor="background1" w:themeTint="80" w:themeShade="40" w:sz="20" w:space="7"/>
              <w:right w:val="dashed" w:color="auto" w:sz="8" w:shadow="1"/>
            </w:pgBorders>
          </w:sectPr></w:pPr><w:r><w:t>First</w:t></w:r></w:p>
          <w:p><w:r><w:t>Second</w:t></w:r></w:p>
          <w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>
        `),
      ),
    );
    archive.file(
      'word/theme/theme1.xml',
      serializeXml(
        themeXml(
          '<a:themeElements><a:clrScheme name="Custom"><a:dk2><a:srgbClr val="112233"/></a:dk2></a:clrScheme></a:themeElements>',
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

    expect(prepared.sections).toHaveLength(2);
    expect(prepared.sections[0]?.layout.pageBorders).toEqual({
      display: 'firstPage',
      offsetFrom: 'text',
      edges: {
        top: {
          style: 'doubleWave',
          color: {
            value: '#889199',
            theme: {
              theme: 'background1',
              resolved: '#889199',
              tint: '80',
              shade: '40',
            },
          },
          size: 20,
          space: 7,
        },
        right: {
          style: 'dashed',
          color: { value: 'auto' },
          size: 8,
          shadow: true,
        },
      },
    });
    expect(prepared.sections[1]?.layout.pageBorders).toBeUndefined();
  });

  test('fails closed for malformed containers and namespace spoofing', () => {
    const theme = createDocxThemeResolver();
    const cases = wordXml(`
      <w:sectPr><w:pgBorders><w:left w:val="single"/><w:top w:val="single"/></w:pgBorders></w:sectPr>
      <w:sectPr><w:pgBorders/><w:pgBorders/></w:sectPr>
      <w:sectPr><w:pgBorders w:display="sometimes"><w:top w:val="single"/></w:pgBorders></w:sectPr>
      <w:sectPr><w:pgBorders><w:top xmlns:r="${RELATIONSHIP_NAMESPACE}" w:val="apples" r:id="rIdUnsafe"/></w:pgBorders></w:sectPr>
      <w:sectPr xmlns:x="urn:spoof"><w:pgBorders><x:top w:val="single"/><w:right w:val="dashed" w:sz="8"/></w:pgBorders></w:sectPr>
    `);
    const sections = descendants(cases, 'sectPr');

    expect(inspectDocxPageBorders(requiredItem(sections, 0), theme)).toEqual({
      status: 'invalid',
      invalidCount: 1,
      spoofedCount: 0,
    });
    expect(inspectDocxPageBorders(requiredItem(sections, 1), theme)).toEqual({
      status: 'invalid',
      invalidCount: 2,
      spoofedCount: 0,
    });
    expect(inspectDocxPageBorders(requiredItem(sections, 2), theme)).toEqual({
      status: 'invalid',
      invalidCount: 1,
      spoofedCount: 0,
    });
    expect(inspectDocxPageBorders(requiredItem(sections, 3), theme)).toEqual({
      status: 'valid',
      pageBorders: { edges: { top: { style: 'nil' } } },
      invalidCount: 1,
      spoofedCount: 0,
    });
    expect(inspectDocxPageBorders(requiredItem(sections, 4), theme)).toEqual({
      status: 'valid',
      pageBorders: {
        edges: { right: { style: 'dashed', size: 8 } },
      },
      invalidCount: 0,
      spoofedCount: 1,
    });
  });

  test('exports exact ordered page-border XML for every section', async () => {
    const firstBorders: WorkDocumentPageBorders = {
      display: 'firstPage',
      offsetFrom: 'page',
      zOrder: 'back',
      edges: {
        top: {
          style: 'doubleWave',
          color: {
            value: '#b8d6a3',
            theme: {
              theme: 'accent6',
              resolved: '#b8d6a3',
              tint: '80',
              shade: '40',
            },
          },
          size: 20,
          space: 24,
          shadow: true,
          frame: false,
        },
        left: { style: 'single', color: { value: 'auto' }, size: 8 },
        bottom: { style: 'dotted', color: { value: '#c00000' }, size: 6 },
        right: { style: 'apples', color: { value: '#4472c4' }, size: 18 },
      },
    };
    const secondBorders: WorkDocumentPageBorders = {
      display: 'notFirstPage',
      edges: {
        bottom: { style: 'dashDotStroked', size: 8, space: 2 },
      },
    };
    const content = multiSectionContent(firstBorders, secondBorders);

    const blob = await createDocxBlob(content);
    const archive = await JSZip.loadAsync(await blob.arrayBuffer());
    const document = parseXml(
      (await archive.file('word/document.xml')?.async('text')) ?? '',
    );
    const sections = effectiveSections(document);
    expect(sections).toHaveLength(2);
    const first = requiredChild(requiredItem(sections, 0), 'pgBorders');
    const second = requiredChild(requiredItem(sections, 1), 'pgBorders');
    expect(directChildren(first).map((element) => element.localName)).toEqual(
      DOCUMENT_PAGE_BORDER_EDGES,
    );
    expect(pageBorderContainerAttributes(first)).toEqual({
      zOrder: 'back',
      display: 'firstPage',
      offsetFrom: 'page',
    });
    expect(pageBorderAttributes(requiredChild(first, 'top'))).toEqual({
      val: 'doubleWave',
      color: 'B8D6A3',
      themeColor: 'accent6',
      themeTint: '80',
      themeShade: '40',
      sz: '20',
      space: '24',
      shadow: '1',
      frame: '0',
    });
    expect(pageBorderContainerAttributes(second)).toEqual({
      display: 'notFirstPage',
    });
    expect(directChildren(second).map((element) => element.localName)).toEqual([
      'bottom',
    ]);
    for (const section of sections) {
      const names = directChildren(section).map((element) => element.localName);
      expect(names.indexOf('pgBorders')).toBeGreaterThan(
        names.indexOf('pgMar'),
      );
      expect(names.indexOf('pgBorders')).toBeLessThan(names.indexOf('cols'));
    }

    const imported = await prepareDocxImport(await blob.arrayBuffer());
    expect(
      imported.sections.map((section) => section.layout.pageBorders),
    ).toEqual([firstBorders, secondBorders]);
  });

  test('renders display, offset, and z-order semantics on PDF and live pages', () => {
    const base = baseLayout();
    const firstOnly: WorkDocumentPageBorders = {
      display: 'firstPage',
      offsetFrom: 'page',
      edges: {
        top: {
          style: 'dashed',
          color: { value: '#112233' },
          size: 16,
          space: 24,
        },
      },
    };
    const content: WorkDocumentContent = {
      type: 'document',
      pageSize: 'a4',
      html: '<p>First</p><div data-page-break="true"></div><p>Second</p>',
      pageBorders: firstOnly,
    };
    const rendered = render(<WorkDocumentPdfPages content={content} />);
    let pages = rendered.container.querySelectorAll('.work-pdf-export-page');
    expect(pages).toHaveLength(2);
    const firstBorder = pages[0]?.querySelector<HTMLElement>(
      '.work-document-page-border',
    );
    expect(firstBorder).not.toBeNull();
    expect(firstBorder?.dataset.documentPageBorderZOrder).toBe('front');
    expect(firstBorder?.style.top).toBe('32px');
    expect(firstBorder?.style.borderTopStyle).toBe('dashed');
    expect(pages[1]?.querySelector('.work-document-page-border')).toBeNull();

    const notFirst = {
      ...firstOnly,
      display: 'notFirstPage' as const,
      zOrder: 'back' as const,
    };
    rendered.rerender(
      <WorkDocumentPdfPages content={{ ...content, pageBorders: notFirst }} />,
    );
    pages = rendered.container.querySelectorAll('.work-pdf-export-page');
    expect(pages[0]?.querySelector('.work-document-page-border')).toBeNull();
    expect(
      pages[1]?.querySelector('.work-document-page-border.back'),
    ).not.toBeNull();

    rendered.rerender(
      <DocumentPageStack
        pageColor="#ffffff"
        pageCount={2}
        pageGap={28}
        pageHeight={1123}
        pages={[
          { layout: { ...base, pageBorders: firstOnly }, sectionPage: 1 },
          { layout: { ...base, pageBorders: notFirst }, sectionPage: 2 },
        ]}
      />,
    );
    expect(
      rendered.container.querySelector(
        '[data-document-page-border-stack="front"] [data-page-index="1"]',
      ),
    ).not.toBeNull();
    const backSurface = rendered.container.querySelector<HTMLElement>(
      '[data-document-page-border-stack="back"] [data-page-index="2"]',
    );
    expect(backSurface).not.toBeNull();
    expect(backSurface?.style.top).toBe('1151px');
  });

  test('reports valid, unsafe, and document-wide compatibility boundaries', async () => {
    const archive = new JSZip();
    archive.file(
      'word/document.xml',
      serializeXml(
        wordXml(`
          <w:p><w:pPr><w:sectPr><w:pgBorders><w:top w:val="apples" w:sz="20"/></w:pgBorders></w:sectPr></w:pPr><w:r><w:t>Valid</w:t></w:r></w:p>
          <w:sectPr><w:pgBorders><w:left w:val="single"/><w:top w:val="single"/></w:pgBorders></w:sectPr>
        `),
      ),
    );
    archive.file(
      'word/settings.xml',
      serializeXml(
        settingsXml(
          '<w:alignBordersAndEdges/><w:bordersDoNotSurroundHeader w:val="true"/>',
        ),
      ),
    );
    const bytes = await archive.generateAsync({ type: 'arraybuffer' });

    const compatibility = await analyzeDocxCompatibility(
      new File([bytes], 'page-borders.docx', {
        type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      }),
      [],
    );

    expect(compatibility.issues).toContainEqual(
      expect.objectContaining({
        code: 'docx.page-borders',
        severity: 'info',
      }),
    );
    expect(compatibility.issues).toContainEqual(
      expect.objectContaining({
        code: 'docx.page-borders.invalid',
        severity: 'warning',
      }),
    );
    expect(compatibility.issues).toContainEqual(
      expect.objectContaining({
        code: 'docx.page-borders.settings',
        severity: 'warning',
      }),
    );
  });
});

function multiSectionContent(
  firstBorders: WorkDocumentPageBorders,
  secondBorders: WorkDocumentPageBorders,
): WorkDocumentContent {
  const base = baseLayout();
  const document = new DOMParser().parseFromString('', 'text/html');
  for (const [index, pageBorders] of [firstBorders, secondBorders].entries()) {
    const section = document.createElement('section');
    const layout: WorkDocumentSectionLayout = {
      ...base,
      pageBorders,
      breakAfter: index === 0 ? 'nextPage' : 'continuous',
      pageNumberStart: index === 1 ? 7 : undefined,
    };
    for (const [name, value] of Object.entries(
      documentSectionDomAttributes(layout, `section-${index + 1}`),
    )) {
      section.setAttribute(name, value);
    }
    section.innerHTML = `<p>Section ${index + 1}</p>`;
    document.body.append(section);
  }
  return {
    type: 'document',
    pageSize: 'a4',
    html: document.body.innerHTML,
  };
}

function baseLayout(): WorkDocumentSectionLayout {
  return documentInitialSectionLayout({
    type: 'document',
    html: '<p></p>',
    pageSize: 'a4',
  });
}

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

function requiredDescendant(document: Document, localName: string): Element {
  const element = descendants(document, localName)[0];
  if (!element) throw new Error(`Expected ${localName} descendant.`);
  return element;
}

function requiredChild(parent: ParentNode, localName: string): Element {
  const element = directChildren(parent, localName)[0];
  if (!element) throw new Error(`Expected ${localName} child.`);
  return element;
}

function requiredItem<T>(items: readonly T[], index: number): T {
  const item = items[index];
  if (!item) throw new Error(`Expected item ${index}.`);
  return item;
}

function effectiveSections(document: Document): Element[] {
  return descendants(document, 'sectPr').filter(
    (element) => !element.closest('sectPrChange'),
  );
}

function pageBorderContainerAttributes(
  element: Element,
): Record<string, string> {
  return selectedAttributes(element, ['zOrder', 'display', 'offsetFrom']);
}

function pageBorderAttributes(element: Element): Record<string, string> {
  return selectedAttributes(element, [
    'val',
    'color',
    'themeColor',
    'themeTint',
    'themeShade',
    'sz',
    'space',
    'shadow',
    'frame',
  ]);
}

function selectedAttributes(
  element: Element,
  names: readonly string[],
): Record<string, string> {
  return Object.fromEntries(
    names.flatMap((name) => {
      const value = attribute(element, name);
      return value === null ? [] : [[name, value]];
    }),
  );
}

function pageBorderSize(style: DocumentParagraphBorderStyle): number {
  if (style === 'nil' || style === 'none') return 0;
  return DOCUMENT_PARAGRAPH_BORDER_STYLES.indexOf(style) < 27 ? 8 : 20;
}
