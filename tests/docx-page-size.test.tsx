import { describe, expect, test } from '@rstest/core';
import { render } from '@testing-library/react';
import JSZip from 'jszip';
import { WorkDocumentPdfPages } from '../src/internal/features/work/components/work-document-pages';
import {
  applyDocumentPageGeometry,
  normalizeDocumentPageGeometry,
  normalizeDocumentPaperSource,
  pageTwipsToMillimeters,
  resolveDocumentPageSize,
  serializeDocumentPageGeometry,
  serializeDocumentPaperSource,
  updateDocumentCustomPageMillimeters,
  updateDocumentPageOrientation,
  updateDocumentPaperSizePreset,
  type WorkDocumentPageGeometry,
  type WorkDocumentPaperSource,
} from '../src/internal/features/work/work-document-page-size';
import {
  documentInitialSectionLayout,
  documentSectionDomAttributes,
} from '../src/internal/features/work/work-document-section';
import { createDocxBlob } from '../src/internal/features/work/work-docx-export';
import { prepareDocxImport } from '../src/internal/features/work/work-docx-import';
import {
  inspectDocxPageSize,
  inspectDocxPaperSource,
} from '../src/internal/features/work/work-docx-page-size-import';
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
const RELATIONSHIP_NAMESPACE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

describe('DOCX page size and paper source', () => {
  test('keeps exact native geometry while bounding only the browser projection', () => {
    const geometry: WorkDocumentPageGeometry = {
      width: 1,
      height: 20_001,
      orientation: 'portrait',
      code: 118,
    };
    expect(serializeDocumentPageGeometry(geometry)).toBe(
      '{"width":1,"height":20001,"orientation":"portrait","code":118}',
    );
    expect(normalizeDocumentPageGeometry(geometry)).toEqual(geometry);
    expect(normalizeDocumentPageGeometry({ ...geometry, width: 0 })).toBeNull();
    expect(
      normalizeDocumentPageGeometry({ ...geometry, code: 119 }),
    ).toBeNull();
    expect(normalizeDocumentPageGeometry({ ...geometry, extra: 1 })).toBeNull();

    const paperSource: WorkDocumentPaperSource = { first: 0, other: 65_535 };
    expect(serializeDocumentPaperSource(paperSource)).toBe(
      '{"first":0,"other":65535}',
    );
    expect(normalizeDocumentPaperSource({})).toEqual({});
    expect(normalizeDocumentPaperSource({ first: 65_536 })).toBeNull();

    const layout = applyDocumentPageGeometry(baseLayout(), geometry);
    const resolved = resolveDocumentPageSize(layout);
    expect(layout.pageSize).toBe('custom');
    expect(layout.pageGeometry).toEqual(geometry);
    expect(resolved.nativeWidth).toBeCloseTo(pageTwipsToMillimeters(1), 4);
    expect(resolved.width).toBe(25.4);
    expect(resolved.bounded).toBe(true);

    const resized = updateDocumentCustomPageMillimeters(
      baseLayout(),
      'width',
      180,
    );
    expect(resized.pageSize).toBe('custom');
    expect(resized.pageGeometry?.width).toBe(10_205);
    const landscape = updateDocumentPageOrientation(resized, 'landscape');
    expect(landscape.pageGeometry).toMatchObject({
      width: resized.pageGeometry?.height,
      height: resized.pageGeometry?.width,
      orientation: 'landscape',
    });
    expect(updateDocumentPaperSizePreset(landscape, 'a3')).toMatchObject({
      pageSize: 'a3',
      orientation: 'landscape',
    });
    expect(updateDocumentPaperSizePreset(landscape, 'a3').pageGeometry).toBe(
      undefined,
    );
  });

  test('accepts exact strict universal measures and rejects ambiguous or unsafe setup', () => {
    const valid = requiredDescendant(
      strictWordXml(`
        <w:sectPr>
          <w:pgSz w:w="12.7cm" w:h="10in" w:orient="portrait" w:code="42"/>
          <w:paperSrc w:first="0" w:other="65535"/>
        </w:sectPr>
      `),
      'sectPr',
    );
    expect(inspectDocxPageSize(valid)).toEqual({
      status: 'valid',
      pageGeometry: {
        width: 7_200,
        height: 14_400,
        orientation: 'portrait',
        code: 42,
      },
      invalidCount: 0,
      spoofedCount: 0,
    });
    expect(inspectDocxPaperSource(valid)).toEqual({
      status: 'valid',
      paperSource: { first: 0, other: 65_535 },
      invalidCount: 0,
      spoofedCount: 0,
    });

    const invalidMarkup = [
      strictWordXml('<w:sectPr><w:pgSz w:w="1mm" w:h="14400"/></w:sectPr>'),
      wordXml('<w:sectPr><w:pgSz w:w="10in" w:h="14400"/></w:sectPr>'),
      wordXml('<w:sectPr><w:pgSz w:w="0" w:h="14400"/></w:sectPr>'),
      wordXml(
        '<w:sectPr><w:pgSz w:w="7200" w:h="14400" w:code="119"/></w:sectPr>',
      ),
      wordXml(
        `<w:sectPr xmlns:r="${RELATIONSHIP_NAMESPACE}"><w:pgSz w:w="7200" w:h="14400" r:id="rId1"/></w:sectPr>`,
      ),
      wordXml(
        '<w:sectPr><w:pgSz w:w="7200" w:h="14400"><w:proofErr/></w:pgSz></w:sectPr>',
      ),
      wordXml(
        '<w:sectPr><w:pgSz w:w="7200" w:h="14400"/><w:pgSz w:w="7200" w:h="14400"/></w:sectPr>',
      ),
    ];
    for (const document of invalidMarkup) {
      expect(
        inspectDocxPageSize(requiredDescendant(document, 'sectPr')).status,
      ).toBe('invalid');
    }

    const spoofed = requiredDescendant(
      wordXml(`
        <w:sectPr xmlns:x="urn:spoof">
          <x:pgSz w:w="1" w:h="1"/>
          <w:pgSz w:w="7200" w:h="14400"/>
        </w:sectPr>
      `),
      'sectPr',
    );
    expect(inspectDocxPageSize(spoofed)).toMatchObject({
      status: 'valid',
      spoofedCount: 1,
    });
  });

  test('imports exact multi-section geometry and inherited printer tray codes', async () => {
    const archive = new JSZip();
    archive.file(
      'word/document.xml',
      serializeXml(
        wordXml(`
          <w:p><w:pPr><w:sectPr>
            <w:pgSz w:w="10001" w:h="15003" w:orient="portrait" w:code="42"/>
            <w:paperSrc w:first="5" w:other="6"/>
          </w:sectPr></w:pPr><w:r><w:t>First</w:t></w:r></w:p>
          <w:p><w:r><w:t>Second</w:t></w:r></w:p>
          <w:sectPr/>
        `),
      ),
    );
    const imported = await prepareDocxImport(
      await archive.generateAsync({ type: 'arraybuffer' }),
    );
    expect(imported.sections).toHaveLength(2);
    expect(
      imported.sections.map((section) => section.layout.pageGeometry),
    ).toEqual([
      {
        width: 10_001,
        height: 15_003,
        orientation: 'portrait',
        code: 42,
      },
      {
        width: 10_001,
        height: 15_003,
        orientation: 'portrait',
        code: 42,
      },
    ]);
    expect(
      imported.sections.map((section) => section.layout.paperSource),
    ).toEqual([
      { first: 5, other: 6 },
      { first: 5, other: 6 },
    ]);
  });

  test('exports exact ordered pgSz and paperSrc values, then reimports them', async () => {
    const first: WorkDocumentPageGeometry = {
      width: 10_001,
      height: 15_003,
      orientation: 'portrait',
      code: 42,
    };
    const second: WorkDocumentPageGeometry = {
      width: 20_005,
      height: 12_007,
      orientation: 'landscape',
      code: 88,
    };
    const sources: WorkDocumentPaperSource[] = [{ first: 5, other: 6 }, {}];
    const content = multiSectionContent(first, second, sources);
    const blob = await createDocxBlob(content);
    const archive = await JSZip.loadAsync(await blob.arrayBuffer());
    const document = parseXml(
      (await archive.file('word/document.xml')?.async('text')) ?? '',
    );
    const sections = effectiveSections(document);
    expect(sections).toHaveLength(2);

    for (const [index, geometry] of [first, second].entries()) {
      const section = requiredItem(sections, index);
      const pageSize = requiredChild(section, 'pgSz');
      expect(Array.from(pageSize.attributes).map((item) => item.name)).toEqual([
        'w:w',
        'w:h',
        'w:orient',
        'w:code',
      ]);
      expect(
        selectedAttributes(pageSize, ['w', 'h', 'orient', 'code']),
      ).toEqual({
        w: String(geometry.width),
        h: String(geometry.height),
        orient: geometry.orientation,
        code: String(geometry.code),
      });
      const paperSource = requiredChild(section, 'paperSrc');
      expect(selectedAttributes(paperSource, ['first', 'other'])).toEqual(
        index === 0 ? { first: '5', other: '6' } : { first: '', other: '' },
      );
      const names = directChildren(section).map((element) => element.localName);
      expect(names.indexOf('pgSz')).toBeLessThan(names.indexOf('pgMar'));
      expect(names.indexOf('paperSrc')).toBeGreaterThan(names.indexOf('pgMar'));
      expect(names.indexOf('paperSrc')).toBeLessThan(names.indexOf('cols'));
    }

    const imported = await prepareDocxImport(await blob.arrayBuffer());
    expect(
      imported.sections.map((section) => section.layout.pageGeometry),
    ).toEqual([first, second]);
    expect(
      imported.sections.map((section) => section.layout.paperSource),
    ).toEqual(sources);
  });

  test('renders exact custom page dimensions and exposes exact PDF points', () => {
    const rendered = render(
      <WorkDocumentPdfPages
        content={{
          type: 'document',
          pageSize: 'custom',
          pageGeometry: { width: 10_001, height: 15_003 },
          html: '<p>Custom page</p>',
        }}
      />,
    );
    const page = rendered.container.querySelector<HTMLElement>(
      '.work-pdf-export-page',
    );
    expect(page).not.toBeNull();
    expect(Number.parseFloat(page?.style.width ?? '')).toBeCloseTo(
      (10_001 / 1_440) * 96,
      3,
    );
    expect(Number.parseFloat(page?.style.minHeight ?? '')).toBeCloseTo(
      (15_003 / 1_440) * 96,
      3,
    );
    expect(page?.dataset.pdfPageWidthPoints).toBe('500.05');
    expect(page?.dataset.pdfPageHeightPoints).toBe('750.15');
  });

  test('reports malformed, bounded, mixed-size, and printer-specific boundaries', async () => {
    const archive = new JSZip();
    archive.file(
      'word/document.xml',
      serializeXml(
        wordXml(`
          <w:p><w:pPr><w:sectPr>
            <w:pgSz w:w="1" w:h="15003" w:code="42"/>
            <w:paperSrc w:first="5" w:other="6"/>
          </w:sectPr></w:pPr><w:r><w:t>First</w:t></w:r></w:p>
          <w:p><w:pPr><w:sectPr>
            <w:pgSz w:w="20005" w:h="12007" w:orient="landscape"/>
          </w:sectPr></w:pPr><w:r><w:t>Second</w:t></w:r></w:p>
          <w:sectPr><w:pgSz w:w="0" w:h="100"/></w:sectPr>
        `),
      ),
    );
    const bytes = await archive.generateAsync({ type: 'arraybuffer' });
    const report = await analyzeDocxCompatibility(
      new File([bytes], 'page-size.docx', {
        type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      }),
      [],
    );
    const codes = report.issues.map((issue) => issue.code);
    expect(codes).toContain('docx.page-size');
    expect(codes).toContain('docx.paper-source');
    expect(codes).toContain('docx.page-size.invalid');
    expect(codes).toContain('docx.page-size.browser-bounds');
    expect(codes).toContain('docx.page-size.mixed-live-layout');
  });
});

function baseLayout(): WorkDocumentSectionLayout {
  return documentInitialSectionLayout({
    type: 'document',
    html: '<p></p>',
    pageSize: 'a4',
  });
}

function multiSectionContent(
  first: WorkDocumentPageGeometry,
  second: WorkDocumentPageGeometry,
  sources: readonly WorkDocumentPaperSource[],
): WorkDocumentContent {
  const document = new DOMParser().parseFromString('', 'text/html');
  for (const [index, pageGeometry] of [first, second].entries()) {
    const section = document.createElement('section');
    const layout = applyDocumentPageGeometry(baseLayout(), pageGeometry);
    layout.paperSource = sources[index];
    layout.breakAfter = index === 0 ? 'nextPage' : 'continuous';
    for (const [name, value] of Object.entries(
      documentSectionDomAttributes(layout, `section-${index + 1}`),
    )) {
      section.setAttribute(name, value);
    }
    section.innerHTML = `<p>Section ${index + 1}</p>`;
    document.body.append(section);
  }
  return { type: 'document', pageSize: 'a4', html: document.body.innerHTML };
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

function selectedAttributes(
  element: Element,
  names: readonly string[],
): Record<string, string | undefined> {
  return Object.fromEntries(
    names.map((name) => [name, attribute(element, name) ?? '']),
  );
}
