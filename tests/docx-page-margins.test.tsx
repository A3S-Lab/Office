import { describe, expect, test } from '@rstest/core';
import { render } from '@testing-library/react';
import JSZip from 'jszip';
import { WorkDocumentPdfPages } from '../src/internal/features/work/components/work-document-pages';
import {
  documentPageMarginBody,
  normalizeDocumentPageMargins,
  parseDocumentPageMargins,
  resolveDocumentPageMargins,
  serializeDocumentPageMargins,
  type WorkDocumentPageMargins,
  updateDocumentGutterPosition,
  updateDocumentMirrorMargins,
  updateDocumentPageMarginMode,
} from '../src/internal/features/work/work-document-page-margins';
import {
  documentInitialSectionLayout,
  documentSectionDomAttributes,
} from '../src/internal/features/work/work-document-section';
import { createDocxBlob } from '../src/internal/features/work/work-docx-export';
import { prepareDocxImport } from '../src/internal/features/work/work-docx-import';
import {
  inspectDocxPageMargins,
  inspectDocxPageMarginSettings,
} from '../src/internal/features/work/work-docx-page-margins-import';
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

describe('DOCX page margins', () => {
  test('keeps exact native geometry and resolves signed, mirrored, gutter, and bounded pages', () => {
    const source: WorkDocumentPageMargins = {
      top: -1_440,
      right: 720,
      bottom: 1_440,
      left: 2_160,
      header: 360,
      footer: 480,
      gutter: 240,
      mirrorMargins: true,
      gutterAtTop: false,
      gutterOnRight: false,
    };
    const serialized = serializeDocumentPageMargins(source);

    expect(serialized).toBe(
      '{"top":-1440,"right":720,"bottom":1440,"left":2160,"header":360,"footer":480,"gutter":240,"mirrorMargins":true,"gutterAtTop":false,"gutterOnRight":false}',
    );
    expect(parseDocumentPageMargins(serialized)).toEqual(source);
    expect(
      normalizeDocumentPageMargins({ ...source, unsafe: true }),
    ).toBeNull();
    expect(
      normalizeDocumentPageMargins({ ...source, right: 31_681 }),
    ).toBeNull();
    expect(normalizeDocumentPageMargins({ ...source, header: 1.5 })).toBeNull();

    const layout = { ...baseLayout(), pageMargins: source };
    const odd = resolveDocumentPageMargins(layout, 1);
    const even = resolveDocumentPageMargins(layout, 2);
    expect(odd).toMatchObject({
      gutterPosition: 'left',
      mirrorMargins: true,
      topMode: 'fromPageEdge',
      bottomMode: 'clearChrome',
      bounded: false,
    });
    expect(odd.body.left).toBeCloseTo(42.3333, 4);
    expect(odd.body.right).toBeCloseTo(12.7, 4);
    expect(even.gutterPosition).toBe('right');
    expect(even.body.left).toBeCloseTo(12.7, 4);
    expect(even.body.right).toBeCloseTo(42.3333, 4);
    expect(odd.headerDistance).toBeCloseTo(6.35, 4);
    expect(odd.footerDistance).toBeCloseTo(8.4667, 4);

    const boundedSource = {
      ...source,
      left: 31_680,
      right: 31_680,
      gutter: 31_680,
    };
    const bounded = resolveDocumentPageMargins(
      { ...layout, pageMargins: boundedSource },
      1,
    );
    expect(bounded.bounded).toBe(true);
    expect(bounded.body.left + bounded.body.right).toBeCloseTo(209, 4);
    expect(boundedSource.left).toBe(31_680);

    const topGutter = updateDocumentGutterPosition(layout, 'top');
    expect(topGutter.pageMargins).toMatchObject({
      gutterAtTop: true,
      gutterOnRight: false,
      mirrorMargins: false,
    });
    const mirrored = updateDocumentMirrorMargins(topGutter, true);
    expect(mirrored.pageMargins).toMatchObject({
      gutterAtTop: false,
      mirrorMargins: true,
    });
    expect(
      updateDocumentPageMarginMode(mirrored, 'bottom', 'fromPageEdge')
        .pageMargins?.bottom,
    ).toBe(-1_440);
  });

  test('parses exact strict universal measures and rejects ambiguous or unsafe markup', () => {
    const settings = inspectDocxPageMarginSettings(
      strictSettingsXml(
        '<w:mirrorMargins w:val="on"/><w:gutterAtTop w:val="0"/>',
      ),
    );
    const section = requiredDescendant(
      strictWordXml(`
        <w:sectPr>
          <w:pgMar w:top="-12.7mm" w:right="1in" w:bottom="36pt" w:left="6pc" w:header="0.5in" w:footer="1.27cm" w:gutter="0mm"/>
          <w:rtlGutter/>
        </w:sectPr>
      `),
      'sectPr',
    );

    expect(settings).toEqual({
      mirrorMargins: true,
      gutterAtTop: false,
      invalidCount: 0,
      spoofedCount: 0,
      incompatible: ['mirrorMargins'],
    });
    expect(inspectDocxPageMargins(section, settings)).toEqual({
      status: 'valid',
      pageMargins: {
        top: -720,
        right: 1_440,
        bottom: 720,
        left: 1_440,
        header: 720,
        footer: 720,
        gutter: 0,
        mirrorMargins: true,
        gutterAtTop: false,
        gutterOnRight: true,
      },
      invalidCount: 0,
      spoofedCount: 0,
    });

    const invalidSections = [
      strictWordXml(
        '<w:sectPr><w:pgMar w:top="1mm" w:right="720" w:bottom="720" w:left="720" w:header="360" w:footer="360" w:gutter="0"/></w:sectPr>',
      ),
      wordXml(
        '<w:sectPr><w:pgMar w:top="1in" w:right="720" w:bottom="720" w:left="720" w:header="360" w:footer="360" w:gutter="0"/></w:sectPr>',
      ),
      wordXml(
        `<w:sectPr xmlns:r="${RELATIONSHIP_NAMESPACE}"><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="360" w:footer="360" w:gutter="0" r:id="rIdUnsafe"/></w:sectPr>`,
      ),
      wordXml(
        '<w:sectPr><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="360" w:footer="360"><w:proofErr/></w:pgMar></w:sectPr>',
      ),
      wordXml(
        '<w:sectPr><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="360" w:footer="360" w:gutter="0"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="360" w:footer="360" w:gutter="0"/></w:sectPr>',
      ),
    ].map((document) => requiredDescendant(document, 'sectPr'));
    for (const invalid of invalidSections) {
      expect(inspectDocxPageMargins(invalid).status).toBe('invalid');
    }

    const spoofed = requiredDescendant(
      wordXml(`
        <w:sectPr xmlns:x="urn:spoof">
          <x:pgMar w:top="1" w:right="1" w:bottom="1" w:left="1" w:header="1" w:footer="1" w:gutter="1"/>
          <w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="360" w:footer="360" w:gutter="0"/>
        </w:sectPr>
      `),
      'sectPr',
    );
    expect(inspectDocxPageMargins(spoofed)).toMatchObject({
      status: 'valid',
      spoofedCount: 1,
    });
  });

  test('imports multi-section inheritance with global and per-section gutter state', async () => {
    const archive = new JSZip();
    archive.file(
      'word/document.xml',
      serializeXml(
        wordXml(`
          <w:p><w:pPr><w:sectPr>
            <w:pgMar w:top="-720" w:right="1440" w:bottom="1000" w:left="1800" w:header="360" w:footer="400" w:gutter="240"/>
            <w:rtlGutter/>
          </w:sectPr></w:pPr><w:r><w:t>First</w:t></w:r></w:p>
          <w:p><w:r><w:t>Second</w:t></w:r></w:p>
          <w:sectPr><w:rtlGutter w:val="0"/></w:sectPr>
        `),
      ),
    );
    archive.file(
      'word/settings.xml',
      serializeXml(
        settingsXml('<w:mirrorMargins w:val="1"/><w:gutterAtTop w:val="0"/>'),
      ),
    );

    const imported = await prepareDocxImport(
      await archive.generateAsync({ type: 'arraybuffer' }),
    );
    expect(imported.sections).toHaveLength(2);
    expect(imported.sections[0]?.layout.pageMargins).toEqual({
      top: -720,
      right: 1_440,
      bottom: 1_000,
      left: 1_800,
      header: 360,
      footer: 400,
      gutter: 240,
      mirrorMargins: true,
      gutterAtTop: false,
      gutterOnRight: true,
    });
    expect(imported.sections[1]?.layout.pageMargins).toEqual({
      top: -720,
      right: 1_440,
      bottom: 1_000,
      left: 1_800,
      header: 360,
      footer: 400,
      gutter: 240,
      mirrorMargins: true,
      gutterAtTop: false,
      gutterOnRight: false,
    });
    expect(imported.sections[0]?.layout.margins.top).toBeCloseTo(12.7, 4);
  });

  test('exports all seven values and ordered settings, then reimports them exactly', async () => {
    const first: WorkDocumentPageMargins = {
      top: -721,
      right: 1_441,
      bottom: 722,
      left: 1_442,
      header: 361,
      footer: 362,
      gutter: 363,
      mirrorMargins: true,
      gutterAtTop: false,
      gutterOnRight: true,
    };
    const second: WorkDocumentPageMargins = {
      top: 801,
      right: 802,
      bottom: -803,
      left: 804,
      header: 365,
      footer: 366,
      gutter: 367,
      mirrorMargins: true,
      gutterAtTop: false,
      gutterOnRight: false,
    };
    const content = multiSectionContent(first, second);
    const blob = await createDocxBlob(content);
    const archive = await JSZip.loadAsync(await blob.arrayBuffer());
    const document = parseXml(
      (await archive.file('word/document.xml')?.async('text')) ?? '',
    );
    const sections = effectiveSections(document);

    expect(sections).toHaveLength(2);
    for (const [index, expected] of [first, second].entries()) {
      const section = requiredItem(sections, index);
      const pageMargins = requiredChild(section, 'pgMar');
      expect(
        Array.from(pageMargins.attributes).map((item) => item.name),
      ).toEqual([
        'w:top',
        'w:right',
        'w:bottom',
        'w:left',
        'w:header',
        'w:footer',
        'w:gutter',
      ]);
      expect(selectedAttributes(pageMargins)).toEqual({
        top: String(expected.top),
        right: String(expected.right),
        bottom: String(expected.bottom),
        left: String(expected.left),
        header: String(expected.header),
        footer: String(expected.footer),
        gutter: String(expected.gutter),
      });
      expect(attribute(requiredChild(section, 'rtlGutter'), 'val')).toBe(
        expected.gutterOnRight ? '1' : '0',
      );
      const names = directChildren(section).map((element) => element.localName);
      expect(names.indexOf('pgMar')).toBeGreaterThan(names.indexOf('pgSz'));
      expect(names.indexOf('pgMar')).toBeLessThan(names.indexOf('cols'));
      expect(names.indexOf('rtlGutter')).toBeGreaterThan(names.indexOf('cols'));
      if (names.includes('docGrid')) {
        expect(names.indexOf('rtlGutter')).toBeLessThan(
          names.indexOf('docGrid'),
        );
      }
    }

    const settings = parseXml(
      (await archive.file('word/settings.xml')?.async('text')) ?? '',
    );
    const settingsNames = directChildren(settings.documentElement).map(
      (element) => element.localName,
    );
    expect(
      attribute(
        requiredChild(settings.documentElement, 'mirrorMargins'),
        'val',
      ),
    ).toBe('1');
    expect(
      attribute(requiredChild(settings.documentElement, 'gutterAtTop'), 'val'),
    ).toBe('0');
    expect(settingsNames.indexOf('mirrorMargins')).toBeLessThan(
      settingsNames.indexOf('gutterAtTop'),
    );
    expect(settingsNames.indexOf('gutterAtTop')).toBeLessThan(
      settingsNames.indexOf('compat'),
    );

    const imported = await prepareDocxImport(await blob.arrayBuffer());
    expect(
      imported.sections.map((section) => section.layout.pageMargins),
    ).toEqual([first, second]);

    const sourceArchive = await JSZip.loadAsync(await blob.arrayBuffer());
    const sourceSettings = parseXml(
      (await sourceArchive.file('word/settings.xml')?.async('text')) ?? '',
    );
    appendOnOffSetting(sourceSettings, 'mirrorMargins', '0');
    appendOnOffSetting(sourceSettings, 'gutterAtTop', '1');
    sourceArchive.file('word/settings.xml', serializeXml(sourceSettings));
    const sourceBytes = await sourceArchive.generateAsync({
      type: 'arraybuffer',
    });
    const sourceBacked = await createDocxBlob(content, sourceBytes);
    const sourceBackedArchive = await JSZip.loadAsync(
      await sourceBacked.arrayBuffer(),
    );
    const authoritativeSettings = parseXml(
      (await sourceBackedArchive.file('word/settings.xml')?.async('text')) ??
        '',
    );
    const mirrorSettings = directChildren(
      authoritativeSettings.documentElement,
      'mirrorMargins',
    );
    const topGutterSettings = directChildren(
      authoritativeSettings.documentElement,
      'gutterAtTop',
    );
    expect(mirrorSettings).toHaveLength(1);
    expect(attribute(requiredItem(mirrorSettings, 0), 'val')).toBe('1');
    expect(topGutterSettings).toHaveLength(1);
    expect(attribute(requiredItem(topGutterSettings, 0), 'val')).toBe('0');
  });

  test('renders mirrored gutters and header/footer distances per physical PDF page', () => {
    const pageMargins: WorkDocumentPageMargins = {
      top: -1_440,
      right: 720,
      bottom: 1_440,
      left: 1_440,
      header: 360,
      footer: 720,
      gutter: 720,
      mirrorMargins: true,
      gutterOnRight: false,
    };
    const rendered = render(
      <WorkDocumentPdfPages
        content={{
          type: 'document',
          pageSize: 'a4',
          html: '<p>First</p><div data-page-break="true"></div><p>Second</p>',
          pageMargins,
        }}
      />,
    );
    const pages = rendered.container.querySelectorAll<HTMLElement>(
      '.work-pdf-export-page',
    );
    expect(pages).toHaveLength(2);
    const first = requiredItem(Array.from(pages), 0);
    const second = requiredItem(Array.from(pages), 1);
    expect(Number.parseFloat(first.style.paddingLeft)).toBeCloseTo(144, 4);
    expect(Number.parseFloat(first.style.paddingRight)).toBeCloseTo(48, 4);
    expect(Number.parseFloat(second.style.paddingLeft)).toBeCloseTo(48, 4);
    expect(Number.parseFloat(second.style.paddingRight)).toBeCloseTo(144, 4);
    expect(first.dataset.documentPageGutterPosition).toBe('left');
    expect(second.dataset.documentPageGutterPosition).toBe('right');
    expect(first.dataset.documentPageTopMarginMode).toBe('fromPageEdge');
    expect(
      Number.parseFloat(
        first.style.getPropertyValue('--work-document-page-header-distance'),
      ),
    ).toBeCloseTo(24, 4);
    expect(
      Number.parseFloat(
        first.style.getPropertyValue('--work-document-page-header-height'),
      ),
    ).toBeCloseTo(72, 4);
    expect(
      Number.parseFloat(
        first.style.getPropertyValue('--work-document-page-footer-distance'),
      ),
    ).toBeCloseTo(48, 4);
  });

  test('reports malformed, conflicting, facing-page, and impossible-body boundaries', async () => {
    const archive = new JSZip();
    archive.file(
      'word/document.xml',
      serializeXml(
        wordXml(`
          <w:p><w:pPr><w:sectPr>
            <w:pgSz w:w="11906" w:h="16838"/>
            <w:pgMar w:top="-720" w:right="31680" w:bottom="720" w:left="31680" w:header="360" w:footer="360" w:gutter="31680"/>
          </w:sectPr></w:pPr><w:r><w:t>Valid</w:t></w:r></w:p>
          <w:sectPr>
            <w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="360" w:footer="360"/>
          </w:sectPr>
        `),
      ),
    );
    archive.file(
      'word/settings.xml',
      serializeXml(
        settingsXml('<w:mirrorMargins/><w:gutterAtTop/><w:bookFoldPrinting/>'),
      ),
    );
    const bytes = await archive.generateAsync({ type: 'arraybuffer' });
    const report = await analyzeDocxCompatibility(
      new File([bytes], 'page-margins.docx', {
        type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      }),
      [],
    );
    const codes = report.issues.map((issue) => issue.code);

    expect(codes).toContain('docx.page-margins');
    expect(codes).toContain('docx.page-margins.invalid');
    expect(codes).toContain('docx.page-margins.gutter-conflict');
    expect(codes).toContain('docx.page-margins.facing-pages');
    expect(codes).toContain('docx.page-margins.body-bounds');
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
  first: WorkDocumentPageMargins,
  second: WorkDocumentPageMargins,
): WorkDocumentContent {
  const document = new DOMParser().parseFromString('', 'text/html');
  for (const [index, pageMargins] of [first, second].entries()) {
    const section = document.createElement('section');
    const base = baseLayout();
    const layout: WorkDocumentSectionLayout = {
      ...base,
      margins: documentPageMarginBody(pageMargins, base.margins),
      pageMargins,
      breakAfter: index === 0 ? 'nextPage' : 'continuous',
    };
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

function settingsXml(body: string): Document {
  return parseXml(
    `<w:settings xmlns:w="${WORD_NAMESPACE}">${body}</w:settings>`,
  );
}

function strictSettingsXml(body: string): Document {
  return parseXml(
    `<w:settings xmlns:w="${STRICT_WORD_NAMESPACE}">${body}</w:settings>`,
  );
}

function serializeXml(document: Document): string {
  return new XMLSerializer().serializeToString(document);
}

function appendOnOffSetting(
  document: Document,
  localName: string,
  value: string,
): void {
  const element = document.createElementNS(WORD_NAMESPACE, `w:${localName}`);
  element.setAttributeNS(WORD_NAMESPACE, 'w:val', value);
  document.documentElement.append(element);
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

function selectedAttributes(element: Element): Record<string, string> {
  return Object.fromEntries(
    ['top', 'right', 'bottom', 'left', 'header', 'footer', 'gutter'].map(
      (name) => [name, attribute(element, name) ?? ''],
    ),
  );
}
