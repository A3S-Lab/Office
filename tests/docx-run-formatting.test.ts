import { describe, expect, test } from '@rstest/core';
import { Editor } from '@tiptap/core';
import JSZip from 'jszip';
import {
  createArtifact,
  createArtifactBlob,
  importOfficeFile,
} from '../src/core';
import {
  applyImportedDocxRunFormattingMarkers,
  markDocxRunFormatting,
} from '../src/internal/features/work/work-docx-run-formatting-import';
import { createDocxBlob } from '../src/internal/features/work/work-docx-export';
import { analyzeDocxCompatibility } from '../src/internal/features/work/work-office-diagnostics';
import {
  attribute,
  descendants,
  directChild,
  parseXml,
} from '../src/internal/features/work/work-ooxml-package';
import { createWorkDocumentExtensions } from '../src/internal/features/work/work-document-extensions';
import {
  documentUnderlineDomAttributes,
  workDocumentUnderlineStyles,
} from '../src/internal/features/work/work-document-underline';
import {
  documentStrikeDomAttributes,
  workDocumentStrikeStyles,
} from '../src/internal/features/work/work-document-strike';

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const STRICT_WORD_NAMESPACE =
  'http://purl.oclc.org/ooxml/wordprocessingml/main';
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
  test('resolves native single and double strike flags without flattening inheritance', () => {
    const document = wordXml(`
      <w:p>
        <w:r><w:rPr><w:strike/></w:rPr><w:t>single</w:t></w:r>
        <w:r><w:rPr><w:dstrike/></w:rPr><w:t>double</w:t></w:r>
        <w:r><w:rPr><w:strike/><w:dstrike/></w:rPr><w:t>double-wins</w:t></w:r>
        <w:r><w:rPr><w:strike w:val="0"/><w:dstrike w:val="false"/></w:rPr><w:t>none</w:t></w:r>
        <w:r><w:rPr><w:strike w:val="on"/><w:dstrike w:val="off"/></w:rPr><w:t>single-reset</w:t></w:r>
      </w:p>
    `);

    const markers = markDocxRunFormatting(document);
    expect(
      markers.runs.map(({ formatting }) => formatting.strike?.style),
    ).toEqual(['single', 'double', 'double', 'none', 'single']);

    const html = new DOMParser().parseFromString(
      `<p>${markers.runs
        .map(
          ({ startMarker, endMarker }, index) =>
            `${startMarker}strike-${index}${endMarker}`,
        )
        .join('')}</p>`,
      'text/html',
    );
    applyImportedDocxRunFormattingMarkers(html, markers);
    const strikes = [...html.querySelectorAll<HTMLElement>('s')];
    expect(strikes.map((element) => element.dataset.officeStrikeStyle)).toEqual(
      ['single', 'double', 'double', 'none', 'single'],
    );
    expect(strikes[1]?.style.textDecorationStyle).toBe('double');
    expect(strikes[3]?.style.textDecorationLine).toBe('none');
  });

  test('exports and reopens native strike styles in body and page chrome', async () => {
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    const strikeHtml = (
      style: (typeof workDocumentStrikeStyles)[number],
      text: string,
    ) => {
      const element = document.createElement('s');
      for (const [name, value] of Object.entries(
        documentStrikeDomAttributes({ style }),
      )) {
        element.setAttribute(name, value);
      }
      element.textContent = text;
      return element.outerHTML;
    };
    artifact.content.html = `<p>${workDocumentStrikeStyles
      .map((style) => strikeHtml(style, `body-${style}`))
      .join(' ')}</p>`;
    artifact.content.pageChrome = {
      differentFirstPage: false,
      differentOddEvenPages: false,
      default: {
        headerHtml: `<p>${strikeHtml('double', 'header-double')}</p>`,
        footerHtml: `<p>${strikeHtml('none', 'footer-none')}</p>`,
        showPageNumber: false,
      },
      first: { headerHtml: '', footerHtml: '', showPageNumber: false },
      even: { headerHtml: '', footerHtml: '', showPageNumber: false },
    };

    const blob = await createDocxBlob(artifact.content);
    const archive = await JSZip.loadAsync(await blob.arrayBuffer());
    const documentXml = parseXml(
      (await archive.file('word/document.xml')?.async('text')) ?? '',
    );
    for (const style of workDocumentStrikeStyles) {
      const run = descendants(documentXml, 'r').find(
        (candidate) =>
          directChild(candidate, 't')?.textContent === `body-${style}`,
      );
      if (!run) throw new Error(`Expected body ${style} strike run.`);
      expect(docxStrikeFlags(run)).toEqual({
        double: style === 'double',
        single: style === 'single',
      });
    }
    const headerPath = Object.keys(archive.files).find((path) =>
      /^word\/header\d*\.xml$/i.test(path),
    );
    const footerPath = Object.keys(archive.files).find((path) =>
      /^word\/footer\d*\.xml$/i.test(path),
    );
    if (!headerPath || !footerPath) {
      throw new Error('Expected generated header and footer parts.');
    }
    const headerXml = parseXml(
      (await archive.file(headerPath)?.async('text')) ?? '',
    );
    const footerXml = parseXml(
      (await archive.file(footerPath)?.async('text')) ?? '',
    );
    const headerRun = descendants(headerXml, 'r').find(
      (candidate) =>
        directChild(candidate, 't')?.textContent === 'header-double',
    );
    const footerRun = descendants(footerXml, 'r').find(
      (candidate) => directChild(candidate, 't')?.textContent === 'footer-none',
    );
    if (!headerRun || !footerRun) {
      throw new Error('Expected strike runs in page chrome.');
    }
    expect(docxStrikeFlags(headerRun)).toEqual({ double: true, single: false });
    expect(docxStrikeFlags(footerRun)).toEqual({
      double: false,
      single: false,
    });

    const reopened = await importOfficeFile(
      new File([blob], 'native-strikes.docx', { type: blob.type }),
    );
    if (reopened.content.type !== 'document') {
      throw new Error('Expected a reopened document artifact.');
    }
    for (const style of workDocumentStrikeStyles) {
      expect(reopened.content.html).toContain(
        `data-office-strike-style="${style}"`,
      );
    }
    expect(reopened.content.pageChrome?.default.headerHtml).toContain(
      'data-office-strike-style="double"',
    );
    expect(reopened.content.pageChrome?.default.footerHtml).toContain(
      'data-office-strike-style="none"',
    );
  });

  test('imports every native underline style and semantic underline color', () => {
    const document = wordXml(
      `<w:p>${workDocumentUnderlineStyles
        .map(
          (style) =>
            `<w:r><w:rPr><w:u w:val="${style}"/></w:rPr><w:t>${style}</w:t></w:r>`,
        )
        .join(
          '',
        )}<w:r><w:rPr><w:u w:val="wave" w:color="A2B9E2" w:themeColor="accent1" w:themeTint="80"/></w:rPr><w:t>Themed wave</w:t></w:r></w:p>`,
    );
    const theme = themeXml(`
      <a:themeElements>
        <a:clrScheme name="Office">
          <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
        </a:clrScheme>
      </a:themeElements>
    `);

    const markers = markDocxRunFormatting(document, undefined, theme);
    expect(
      markers.runs
        .slice(0, workDocumentUnderlineStyles.length)
        .map(({ formatting }) => formatting.underline?.style),
    ).toEqual(workDocumentUnderlineStyles);
    expect(markers.runs.at(-1)?.formatting.underline).toEqual({
      style: 'wave',
      color: '#a2b9e2',
      themeColor: {
        theme: 'accent1',
        resolved: '#a2b9e2',
        tint: '80',
      },
    });

    const html = new DOMParser().parseFromString(
      `<p>${markers.runs
        .map(
          ({ startMarker, endMarker }, index) =>
            `${startMarker}${index === markers.runs.length - 1 ? 'Themed wave' : workDocumentUnderlineStyles[index]}${endMarker}`,
        )
        .join('')}</p>`,
      'text/html',
    );
    applyImportedDocxRunFormattingMarkers(html, markers);
    const underlines = [...html.querySelectorAll<HTMLElement>('u')];
    expect(
      underlines
        .slice(0, workDocumentUnderlineStyles.length)
        .map((element) => element.dataset.officeUnderlineStyle),
    ).toEqual(workDocumentUnderlineStyles);
    expect(underlines[0]?.style.textDecorationLine).toBe('none');
    expect(underlines[3]?.style.textDecorationStyle).toBe('double');
    expect(underlines.at(-1)?.dataset.officeUnderlineThemeColor).toContain(
      'accent1',
    );
  });

  test('exports and reopens native underline styles in body and page chrome', async () => {
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    const body = workDocumentUnderlineStyles
      .map((style) => {
        const element = document.createElement('u');
        for (const [name, value] of Object.entries(
          documentUnderlineDomAttributes({
            style,
            ...(style === 'wave' ? { color: '#c00000' } : {}),
          }),
        )) {
          element.setAttribute(name, value);
        }
        element.textContent = style;
        return element.outerHTML;
      })
      .join(' ');
    const headerUnderline = document.createElement('u');
    for (const [name, value] of Object.entries(
      documentUnderlineDomAttributes({
        style: 'wavyDouble',
        color: '#4472c4',
      }),
    )) {
      headerUnderline.setAttribute(name, value);
    }
    headerUnderline.textContent = 'Header underline';
    artifact.content.html = `<p>${body}</p>`;
    artifact.content.pageChrome = {
      differentFirstPage: false,
      differentOddEvenPages: false,
      default: {
        headerHtml: `<p>${headerUnderline.outerHTML}</p>`,
        footerHtml: '',
        showPageNumber: false,
      },
      first: { headerHtml: '', footerHtml: '', showPageNumber: false },
      even: { headerHtml: '', footerHtml: '', showPageNumber: false },
    };

    const blob = await createDocxBlob(artifact.content);
    const archive = await JSZip.loadAsync(await blob.arrayBuffer());
    const documentXml = parseXml(
      (await archive.file('word/document.xml')?.async('text')) ?? '',
    );
    const values = Array.from(documentXml.getElementsByTagName('*'))
      .filter((element) => element.localName === 'u')
      .map((element) => attribute(element, 'val'));
    expect(values).toEqual(workDocumentUnderlineStyles);
    expect(
      attribute(
        Array.from(documentXml.getElementsByTagName('*')).find(
          (element) =>
            element.localName === 'u' && attribute(element, 'val') === 'wave',
        ) ?? documentXml.documentElement,
        'color',
      ),
    ).toBe('C00000');
    const headerPath = Object.keys(archive.files).find((path) =>
      /^word\/header\d*\.xml$/i.test(path),
    );
    if (!headerPath) {
      throw new Error('Expected a generated header part.');
    }
    const headerXml = headerPath
      ? ((await archive.file(headerPath)?.async('text')) ?? '')
      : '';
    expect(headerXml).toMatch(
      /<w:u\b[^>]*w:val="wavyDouble"[^>]*w:color="4472C4"/,
    );

    const reopened = await importOfficeFile(
      new File([blob], 'native-underlines.docx', { type: blob.type }),
    );
    if (reopened.content.type !== 'document') {
      throw new Error('Expected a reopened document artifact.');
    }
    for (const style of workDocumentUnderlineStyles) {
      expect(reopened.content.html).toContain(
        `data-office-underline-style="${style}"`,
      );
    }
    expect(reopened.content.pageChrome?.default.headerHtml).toContain(
      'data-office-underline-style="wavyDouble"',
    );
    expect(reopened.content.pageChrome?.default.headerHtml).toContain(
      'data-office-underline-color="#4472c4"',
    );
  });

  test('preserves semantic underline theme color until a direct color edit', async () => {
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    const underline = document.createElement('u');
    for (const [name, value] of Object.entries(
      documentUnderlineDomAttributes({
        style: 'wave',
        color: '#a2b9e2',
        themeColor: {
          theme: 'accent1',
          resolved: '#a2b9e2',
          tint: '80',
        },
      }),
    )) {
      underline.setAttribute(name, value);
    }
    underline.textContent = 'Themed underline';
    artifact.content.html = `<p>${underline.outerHTML}</p>`;

    const blob = await createDocxBlob(artifact.content);
    const archive = await JSZip.loadAsync(await blob.arrayBuffer());
    const xml = parseXml(
      (await archive.file('word/document.xml')?.async('text')) ?? '',
    );
    const exported = descendants(xml, 'u').find(
      (element) => attribute(element, 'val') === 'wave',
    );
    if (!exported) throw new Error('Expected an exported underline.');
    expect(attribute(exported, 'color')).toBe('A2B9E2');
    expect(attribute(exported, 'themeColor')).toBe('accent1');
    expect(attribute(exported, 'themeTint')).toBe('80');

    const reopened = await importOfficeFile(
      new File([blob], 'themed-underline.docx', { type: blob.type }),
    );
    if (reopened.content.type !== 'document') {
      throw new Error('Expected a reopened document artifact.');
    }
    expect(reopened.content.html).toContain(
      'data-office-underline-theme-color',
    );
    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: reopened.content.html,
    });
    editor.commands.selectAll();
    expect(editor.commands.setDocumentUnderlineColor('#123456')).toBe(true);
    expect(editor.getAttributes('underline')).toMatchObject({
      underlineColor: '#123456',
      underlineThemeColor: null,
    });
    artifact.content.html = editor.getHTML();
    editor.destroy();

    const edited = await createDocxBlob(artifact.content);
    const editedArchive = await JSZip.loadAsync(await edited.arrayBuffer());
    const editedXml = parseXml(
      (await editedArchive.file('word/document.xml')?.async('text')) ?? '',
    );
    const editedUnderline = descendants(editedXml, 'u').find(
      (element) => attribute(element, 'val') === 'wave',
    );
    if (!editedUnderline) throw new Error('Expected an edited underline.');
    expect(attribute(editedUnderline, 'color')).toBe('123456');
    expect(attribute(editedUnderline, 'themeColor')).toBeNull();
  });

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
      underline: { style: 'none' },
      strike: { style: 'none' },
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

  test('imports mutually exclusive native all-caps and small-caps effects', () => {
    const document = wordXml(`
      <w:p>
        <w:r><w:rPr><w:caps/></w:rPr><w:t>All caps</w:t></w:r>
        <w:r><w:rPr><w:smallCaps/></w:rPr><w:t>Small caps</w:t></w:r>
        <w:r><w:rPr><w:caps w:val="0"/><w:smallCaps w:val="0"/></w:rPr><w:t>Normal</w:t></w:r>
      </w:p>
    `);

    const markers = markDocxRunFormatting(document);
    expect(markers.runs.map(({ formatting }) => formatting)).toEqual([
      expect.objectContaining({ textCase: 'all-caps' }),
      expect.objectContaining({ textCase: 'small-caps' }),
      expect.objectContaining({ textCase: 'none' }),
    ]);
    const html = new DOMParser().parseFromString(
      `<p>${markers.runs
        .map(
          ({ startMarker, endMarker }, index) =>
            `${startMarker}${['All caps', 'Small caps', 'Normal'][index]}${endMarker}`,
        )
        .join('')}</p>`,
      'text/html',
    );
    applyImportedDocxRunFormattingMarkers(html, markers);

    const spans = [...html.querySelectorAll<HTMLElement>('p > span')];
    expect(spans.map((span) => span.dataset.officeTextCase)).toEqual([
      'all-caps',
      'small-caps',
      'none',
    ]);
    expect(spans[0]?.style.textTransform).toBe('uppercase');
    expect(spans[1]?.style.fontVariantCaps).toBe('small-caps');
    expect(spans[2]?.style.textTransform).toBe('none');
    expect(spans[2]?.style.fontVariantCaps).toBe('normal');
  });

  test('exports and reopens native text-case effects', async () => {
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    artifact.content.html =
      '<section data-document-section="true"><p><span data-office-text-case="all-caps" style="text-transform: uppercase;">All caps</span> <span data-office-text-case="small-caps" style="font-variant-caps: small-caps;">Small caps</span></p></section>';

    const blob = await createArtifactBlob(artifact);
    const archive = await JSZip.loadAsync(await blob.arrayBuffer());
    const xml = (await archive.file('word/document.xml')?.async('text')) ?? '';
    expect(xml).toMatch(/<w:caps\/?\s*>/);
    expect(xml).toMatch(/<w:smallCaps\/?\s*>/);

    const reopened = await importOfficeFile(
      new File([blob], 'text-case.docx', { type: blob.type }),
    );
    if (reopened.content.type !== 'document') {
      throw new Error('Expected a reopened document artifact.');
    }
    expect(reopened.content.html).toContain('data-office-text-case="all-caps"');
    expect(reopened.content.html).toContain(
      'data-office-text-case="small-caps"',
    );
  });

  test('exports nested text-case overrides without an inherited competing flag', async () => {
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    artifact.content.html = [
      '<section data-document-section="true"><p>',
      '<span data-office-text-case="small-caps" style="font-variant-caps: small-caps;">',
      'Outer ',
      '<span data-office-text-case="all-caps" style="text-transform: uppercase;">All override</span>',
      ' <span data-office-text-case="none" style="text-transform: none; font-variant-caps: normal;">Normal override</span>',
      '</span>',
      '</p></section>',
    ].join('');

    const blob = await createArtifactBlob(artifact);
    const archive = await JSZip.loadAsync(await blob.arrayBuffer());
    const xml = parseXml(
      (await archive.file('word/document.xml')?.async('text')) ?? '',
    );
    const propertiesFor = (text: string): Element | null => {
      const textElement = Array.from(xml.getElementsByTagName('*')).find(
        (element) => element.localName === 't' && element.textContent === text,
      );
      const run = textElement?.parentElement;
      return (
        Array.from(run?.children ?? []).find(
          (element) =>
            element.namespaceURI === WORD_NAMESPACE &&
            element.localName === 'rPr',
        ) ?? null
      );
    };

    const allCaps = propertiesFor('All override');
    expect(allCaps).not.toBeNull();
    expect(
      Array.from(allCaps?.getElementsByTagName('*') ?? []).filter(
        (element) => element.localName === 'caps',
      ),
    ).toHaveLength(1);
    expect(
      Array.from(allCaps?.getElementsByTagName('*') ?? []).filter(
        (element) => element.localName === 'smallCaps',
      ),
    ).toHaveLength(0);

    const normal = propertiesFor('Normal override');
    expect(normal).not.toBeNull();
    expect(
      Array.from(normal?.getElementsByTagName('*') ?? []).filter(
        (element) => element.localName === 'smallCaps',
      ),
    ).toHaveLength(1);
    expect(
      Array.from(normal?.getElementsByTagName('*') ?? []).filter(
        (element) => element.localName === 'caps',
      ),
    ).toHaveLength(0);
  });

  test('imports a bounded native run-property revision as reviewable formatting', () => {
    const document = wordXml(`
      <w:p><w:r>
        <w:rPr>
          <w:b/>
          <w:rPrChange w:id="7" w:author="Ada Reviewer" w:date="2026-08-17T14:30:00Z">
            <w:rPr>
              <w:i/>
              <w:smallCaps/>
              <w:sz w:val="24"/>
              <w:color w:val="336699"/>
            </w:rPr>
          </w:rPrChange>
        </w:rPr>
        <w:t>Changed format</w:t>
      </w:r></w:p>
    `);

    const markers = markDocxRunFormatting(document);
    const marker = markers.runs[0];
    if (!marker) throw new Error('Expected a formatting-change run marker.');
    expect(marker.change).toEqual({
      id: 'docx-format-change-7',
      author: 'Ada Reviewer',
      date: '2026-08-17T14:30:00.000Z',
      before:
        '[{"type":"italic"},{"type":"underline","attrs":{"underlineStyle":"none"}},{"type":"strike","attrs":{"strikeStyle":"none"}},{"type":"textStyle","attrs":{"color":"#336699","fontSize":"12pt","textCase":"small-caps"}}]',
    });
    const html = new DOMParser().parseFromString(
      `<p>${marker.startMarker}Changed format${marker.endMarker}</p>`,
      'text/html',
    );

    applyImportedDocxRunFormattingMarkers(html, markers);

    const change = html.querySelector<HTMLElement>(
      '[data-document-change][data-change-kind="formatting"]',
    );
    expect(change?.dataset.changeId).toBe('docx-format-change-7');
    expect(change?.dataset.changeAuthor).toBe('Ada Reviewer');
    expect(change?.dataset.changeBefore).toBe(marker.change.before);
    expect(change?.querySelector('strong')?.textContent).toBe('Changed format');
  });

  test('exports and reopens native underline formatting revisions', async () => {
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    const before =
      '[{"type":"underline","attrs":{"underlineColor":"#4472c4","underlineStyle":"double"}}]';
    const underline = document.createElement('u');
    for (const [name, value] of Object.entries(
      documentUnderlineDomAttributes({ style: 'wave', color: '#c00000' }),
    )) {
      underline.setAttribute(name, value);
    }
    underline.textContent = 'Changed underline';
    artifact.content.html = `<p><span data-document-change="true" data-change-kind="formatting" data-change-before='${before}' data-change-id="formatting-underline" data-change-author="Ada Reviewer" data-change-date="2026-08-22T12:00:00.000Z">${underline.outerHTML}</span></p>`;
    artifact.content.trackChanges = true;

    const blob = await createDocxBlob(artifact.content);
    const archive = await JSZip.loadAsync(await blob.arrayBuffer());
    const xml = parseXml(
      (await archive.file('word/document.xml')?.async('text')) ?? '',
    );
    const underlineProperties = descendants(xml, 'u').map((element) => ({
      color: attribute(element, 'color'),
      style: attribute(element, 'val'),
    }));
    expect(underlineProperties).toContainEqual({
      color: 'C00000',
      style: 'wave',
    });
    expect(underlineProperties).toContainEqual({
      color: '4472C4',
      style: 'double',
    });

    const reopened = await importOfficeFile(
      new File([blob], 'underline-revision.docx', { type: blob.type }),
    );
    if (reopened.content.type !== 'document') {
      throw new Error('Expected a reopened document artifact.');
    }
    expect(reopened.content.html).toContain('data-change-kind="formatting"');
    expect(reopened.content.html).toContain(
      'data-office-underline-style="wave"',
    );
    expect(reopened.content.html).toContain('underlineStyle');
    expect(reopened.content.html).toContain('double');
  });

  test('exports and reopens native strike formatting revisions', async () => {
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    const before = '[{"type":"strike","attrs":{"strikeStyle":"single"}}]';
    const strike = document.createElement('s');
    for (const [name, value] of Object.entries(
      documentStrikeDomAttributes({ style: 'double' }),
    )) {
      strike.setAttribute(name, value);
    }
    strike.textContent = 'Changed strike';
    artifact.content.html = `<p><span data-document-change="true" data-change-kind="formatting" data-change-before='${before}' data-change-id="formatting-strike" data-change-author="Ada Reviewer" data-change-date="2026-08-22T12:00:00.000Z">${strike.outerHTML}</span></p>`;
    artifact.content.trackChanges = true;

    const blob = await createDocxBlob(artifact.content);
    const archive = await JSZip.loadAsync(await blob.arrayBuffer());
    const xml = parseXml(
      (await archive.file('word/document.xml')?.async('text')) ?? '',
    );
    const run = descendants(xml, 'r').find(
      (candidate) =>
        directChild(candidate, 't')?.textContent === 'Changed strike',
    );
    if (!run) throw new Error('Expected changed strike run.');
    expect(docxStrikeFlags(run)).toEqual({ double: true, single: false });
    const change = descendants(run, 'rPrChange')[0];
    if (!change) throw new Error('Expected native run-property revision.');
    expect(docxStrikeFlags(change)).toEqual({ double: false, single: true });

    const reopened = await importOfficeFile(
      new File([blob], 'strike-revision.docx', { type: blob.type }),
    );
    if (reopened.content.type !== 'document') {
      throw new Error('Expected a reopened document artifact.');
    }
    expect(reopened.content.html).toContain('data-change-kind="formatting"');
    expect(reopened.content.html).toContain(
      'data-office-strike-style="double"',
    );
    expect(reopened.content.html).toContain('strikeStyle');
    expect(reopened.content.html).toContain('single');
  });

  test('reports supported character-formatting revisions without a structural warning', async () => {
    const compatibility = await formattingRevisionCompatibility(`
      <w:p><w:r><w:rPr>
        <w:b/>
        <w:rPrChange w:id="7" w:author="Ada Reviewer" w:date="2026-08-17T14:30:00Z">
          <w:rPr><w:i/><w:color w:val="336699"/></w:rPr>
        </w:rPrChange>
      </w:rPr><w:t>Supported change</w:t></w:r></w:p>
    `);

    expect(compatibility.issues).toContainEqual(
      expect.objectContaining({
        code: 'docx.revisions.formatting',
        severity: 'info',
      }),
    );
    expect(
      compatibility.issues.some(
        ({ code }) => code === 'docx.revisions.structural',
      ),
    ).toBe(false);
  });

  test('keeps malformed character-formatting revisions on the structural warning path', async () => {
    const compatibility = await formattingRevisionCompatibility(`
      <w:p><w:r><w:rPr>
        <w:rPrChange w:id="7">
          <w:rPr><w:i/></w:rPr>
        </w:rPrChange>
        <w:rPrChange xmlns:evil="https://example.test/evil" w:id="8" w:author="Spoofed Reviewer">
          <evil:rPr/>
        </w:rPrChange>
      </w:rPr><w:t>Malformed change</w:t></w:r></w:p>
    `);

    expect(
      compatibility.issues.some(
        ({ code }) => code === 'docx.revisions.formatting',
      ),
    ).toBe(false);
    expect(compatibility.issues).toContainEqual(
      expect.objectContaining({
        code: 'docx.revisions.structural',
        severity: 'warning',
      }),
    );
  });

  test('recognizes strict-namespace character-formatting revisions', async () => {
    const archive = new JSZip();
    archive.file(
      'word/document.xml',
      `<s:document xmlns:s="${STRICT_WORD_NAMESPACE}"><s:body>
        <s:p><s:r><s:rPr>
          <s:rPrChange s:id="12" s:author="Strict Reviewer">
            <s:rPr><s:u s:val="single"/></s:rPr>
          </s:rPrChange>
        </s:rPr><s:t>Strict change</s:t></s:r></s:p>
      </s:body></s:document>`,
    );
    const bytes = await archive.generateAsync({ type: 'arraybuffer' });

    const compatibility = await analyzeDocxCompatibility(
      new File([bytes], 'strict-formatting-revision.docx'),
      [],
    );

    expect(compatibility.issues).toContainEqual(
      expect.objectContaining({
        code: 'docx.revisions.formatting',
        severity: 'info',
      }),
    );
    expect(
      compatibility.issues.some(
        ({ code }) => code === 'docx.revisions.structural',
      ),
    ).toBe(false);
  });

  test('exports and reopens native run-property revisions without marker leakage', async () => {
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    const before =
      '[{"type":"italic"},{"type":"textStyle","attrs":{"color":"#336699","fontSize":"12pt","textCase":"small-caps"}}]';
    artifact.content.html = `<section data-document-section="true"><p><span data-document-change="true" data-change-kind="formatting" data-change-before='${before}' data-change-id="formatting-7" data-change-author="Ada Reviewer" data-change-date="2026-08-17T14:30:00.000Z"><strong>Changed format</strong></span></p></section>`;
    artifact.content.trackChanges = true;

    const blob = await createArtifactBlob(artifact);
    const archive = await JSZip.loadAsync(await blob.arrayBuffer());
    const xml = (await archive.file('word/document.xml')?.async('text')) ?? '';
    expect(xml).not.toContain('__A3S_WORK_FORMAT_CHANGE_');
    expect(xml).toMatch(
      /<w:rPrChange\b[^>]*w:id="1"[^>]*w:author="Ada Reviewer"[^>]*w:date="2026-08-17T14:30:00.000Z"/,
    );
    expect(xml).toMatch(
      /<w:rPrChange\b[^>]*>[\s\S]*?<w:rPr>[\s\S]*?<w:i\/>[\s\S]*?<w:smallCaps\/>[\s\S]*?<w:color\b[^>]*w:val="336699"[\s\S]*?<w:sz\b[^>]*w:val="24"[\s\S]*?<\/w:rPr>[\s\S]*?<\/w:rPrChange>/,
    );

    const reopened = await importOfficeFile(
      new File([blob], 'formatting-revision.docx', { type: blob.type }),
    );
    if (reopened.content.type !== 'document') {
      throw new Error('Expected a reopened document artifact.');
    }
    expect(reopened.content.html).toContain('data-change-kind="formatting"');
    expect(reopened.content.html).toContain(
      'data-change-author="Ada Reviewer"',
    );
    expect(reopened.content.html).toContain('Changed format');
  });

  test('exports imported formatting revisions that did not declare a date', async () => {
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    artifact.content.html =
      '<section data-document-section="true"><p><span data-document-change="true" data-change-kind="formatting" data-change-before="[]" data-change-id="formatting-without-date" data-change-author="Ada Reviewer" data-change-date=""><strong>Undated format</strong></span></p></section>';

    const blob = await createArtifactBlob(artifact);
    const archive = await JSZip.loadAsync(await blob.arrayBuffer());
    const xml = (await archive.file('word/document.xml')?.async('text')) ?? '';

    expect(xml).toMatch(
      /<w:rPrChange\b[^>]*w:author="Ada Reviewer"[^>]*w:date="[^"]+"/,
    );
    expect(xml).toContain('<w:b/>');
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
      underline: { style: 'none' },
      strike: { style: 'none' },
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

function docxStrikeFlags(container: Element): {
  double: boolean | undefined;
  single: boolean | undefined;
} {
  const properties =
    container.localName === 'rPr'
      ? container
      : descendants(container, 'rPr')[0];
  return {
    double: docxOnOff(directChild(properties ?? container, 'dstrike')),
    single: docxOnOff(directChild(properties ?? container, 'strike')),
  };
}

function docxOnOff(element: Element | undefined): boolean | undefined {
  if (!element) return undefined;
  const value = attribute(element, 'val')?.trim().toLowerCase();
  return value !== '0' && value !== 'false' && value !== 'off';
}

async function formattingRevisionCompatibility(
  body: string,
): ReturnType<typeof analyzeDocxCompatibility> {
  const archive = new JSZip();
  archive.file(
    'word/document.xml',
    `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body>${body}</w:body></w:document>`,
  );
  const bytes = await archive.generateAsync({ type: 'arraybuffer' });
  return analyzeDocxCompatibility(
    new File([bytes], 'formatting-revision.docx'),
    [],
  );
}
