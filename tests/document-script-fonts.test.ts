import { describe, expect, test } from '@rstest/core';
import { Editor } from '@tiptap/core';
import JSZip from 'jszip';
import { createArtifact, importOfficeFile } from '../src/core';
import { importedDocumentCharacterFormatting } from '../src/internal/features/work/work-document-format-changes';
import {
  documentFontNameFromCssFamily,
  documentScriptFontsDomAttributes,
  documentScriptFontFamily,
  documentScriptFontFamilyForRendering,
  documentScriptFontSegments,
  normalizeDocumentScriptFonts,
  parseDocumentScriptFonts,
  serializeDocumentScriptFonts,
  type WorkDocumentScriptFontSlot,
} from '../src/internal/features/work/work-document-script-fonts';
import {
  DocxRunFontsPatchCollector,
  patchDocxRunFonts,
} from '../src/internal/features/work/work-docx-run-fonts-export';
import {
  inspectDocxRunFonts,
  resolveDocxRunFonts,
} from '../src/internal/features/work/work-docx-run-fonts';
import { createWorkDocumentExtensions } from '../src/internal/features/work/work-document-extensions';
import { createDocxBlob } from '../src/internal/features/work/work-docx-export';
import {
  attribute,
  descendants,
  directChild,
  parseXml,
} from '../src/internal/features/work/work-ooxml-package';

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const STRICT_WORD_NAMESPACE =
  'http://purl.oclc.org/ooxml/wordprocessingml/main';

describe('Writer script-specific fonts', () => {
  test('normalizes and serializes the four native Word font slots', () => {
    const fonts = normalizeDocumentScriptFonts({
      ascii: { name: ' Segoe UI ', resolved: 'Segoe UI' },
      highAnsi: { name: 'Calibri', resolved: 'Calibri' },
      eastAsia: {
        name: 'SimSun',
        theme: 'minorEastAsia',
        resolved: 'Microsoft YaHei',
      },
      complexScript: { name: 'Arial', resolved: 'Arial' },
      hint: 'eastAsia',
    });

    expect(fonts).toEqual({
      ascii: { name: 'Segoe UI', resolved: 'Segoe UI' },
      highAnsi: { name: 'Calibri', resolved: 'Calibri' },
      eastAsia: {
        name: 'SimSun',
        theme: 'minorEastAsia',
        resolved: 'Microsoft YaHei',
      },
      complexScript: { name: 'Arial', resolved: 'Arial' },
      hint: 'eastAsia',
    });
    const serialized = serializeDocumentScriptFonts(fonts);
    expect(parseDocumentScriptFonts(serialized)).toEqual(fonts);
    expect(serialized).toBe(serializeDocumentScriptFonts(fonts));
  });

  test('rejects malformed, oversized, and unknown font metadata', () => {
    expect(normalizeDocumentScriptFonts({ unknown: 'Arial' })).toBeNull();
    expect(
      normalizeDocumentScriptFonts({ ascii: { name: 'x'.repeat(128) } }),
    ).toBeNull();
    expect(
      normalizeDocumentScriptFonts({
        ascii: { name: 'Arial\u0000Spoofed' },
      }),
    ).toBeNull();
    expect(
      normalizeDocumentScriptFonts({
        ascii: { theme: 'majorLatin' },
      }),
    ).toBeNull();
    expect(documentFontNameFromCssFamily('"unterminated')).toBeNull();
  });

  test('escapes CSS declaration syntax while retaining native identity', () => {
    const attributes = documentScriptFontsDomAttributes(
      {
        ascii: {
          name: 'Project; color: red',
          resolved: 'Project; color: red',
        },
      },
      'ascii',
    );
    const span = document.createElement('span');
    for (const [name, value] of Object.entries(attributes)) {
      span.setAttribute(name, value);
    }

    expect(attributes.style).toContain('\\3b ');
    expect(attributes.style).not.toContain('; color');
    const fonts = parseDocumentScriptFonts(
      attributes['data-office-script-fonts'],
    );
    expect(fonts?.ascii?.name).toBe('Project; color: red');
    expect(span.style.fontFamily).toContain('\\3b ');
    expect(span.style.color).toBe('');
    expect(documentFontNameFromCssFamily(span.style.fontFamily)).toBe(
      'Project; color: red',
    );
    expect(
      new DocxRunFontsPatchCollector('source').marker(
        fonts,
        'ascii',
        span.style.fontFamily,
      ),
    ).toBeTruthy();
  });

  test('deduplicates equivalent export patches and fails closed for missing markers', async () => {
    const fonts = normalizeDocumentScriptFonts({
      ascii: { name: 'Arial', resolved: 'Arial' },
    });
    if (!fonts) throw new Error('Expected valid script fonts.');
    const collector = new DocxRunFontsPatchCollector('source');
    const first = collector.marker(fonts, 'ascii', 'Arial');
    const second = collector.marker(fonts, 'ascii', 'Arial, sans-serif');

    expect(first).toBeTruthy();
    expect(second).toBe(first);
    expect(collector.patches).toHaveLength(1);

    const unresolvedTheme = normalizeDocumentScriptFonts({
      ascii: { theme: 'majorAscii' },
    });
    expect(
      new DocxRunFontsPatchCollector('source').marker(
        unresolvedTheme,
        'ascii',
        '',
      ),
    ).toBeTruthy();

    const archive = new JSZip();
    archive.file(
      'word/document.xml',
      `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body/></w:document>`,
    );
    const bytes = await archive.generateAsync({ type: 'arraybuffer' });
    await expect(patchDocxRunFonts(bytes, collector.patches)).rejects.toThrow(
      /run-font markers were not emitted/,
    );
  });

  test('segments one mixed run without splitting neutral punctuation from its script', () => {
    const text = '(A3S) café，中文 (اختبار) בדיקה';
    const segments = documentScriptFontSegments(text, 'default');

    expect(
      segments.map(({ from, to, slot }) => ({
        text: text.slice(from, to),
        slot,
      })),
    ).toEqual([
      { text: '(A3S) caf', slot: 'ascii' },
      { text: 'é，', slot: 'highAnsi' },
      { text: '中文 (', slot: 'eastAsia' },
      { text: 'اختبار) בדיקה', slot: 'complexScript' },
    ]);
  });

  test('routes common Indic and Southeast Asian shaping scripts through the complex slot', () => {
    for (const text of ['हिन्दी', 'বাংলা', 'ไทย', 'မြန်မာ', 'ខ្មែរ', 'ᠮᠣᠩᠭᠣᠯ']) {
      expect(documentScriptFontSegments(text)).toEqual([
        { from: 0, to: text.length, slot: 'complexScript' },
      ]);
    }
  });

  test('puts the requested slot first while retaining bounded font fallbacks', () => {
    const fonts = normalizeDocumentScriptFonts({
      ascii: { resolved: 'Segoe UI' },
      highAnsi: { resolved: 'Calibri' },
      eastAsia: { resolved: 'Microsoft YaHei' },
      complexScript: { resolved: 'Arial' },
    });
    expect(fonts).not.toBeNull();
    expect(documentScriptFontFamily(fonts, 'eastAsia')).toBe(
      '"Microsoft YaHei", Calibri, "Segoe UI", Arial',
    );
    expect(documentScriptFontFamily(fonts, 'complexScript')).toBe(
      'Arial, Calibri, "Segoe UI", "Microsoft YaHei"',
    );
  });

  test('retains a compatible CSS fallback list without trusting declaration syntax', () => {
    const fonts = normalizeDocumentScriptFonts({
      ascii: { name: 'Arial', resolved: 'Arial' },
    });
    expect(fonts).not.toBeNull();
    expect(
      documentScriptFontFamilyForRendering(
        fonts,
        'ascii',
        ' Arial, sans-serif ',
      ),
    ).toBe('Arial, sans-serif');
    expect(
      documentScriptFontFamilyForRendering(
        fonts,
        'ascii',
        'Calibri, sans-serif',
      ),
    ).toBe('Arial');
    for (const forged of [
      'Arial; color: red',
      'Arial { color: red }',
      'Arial } body { display: none',
      'Arial, "unterminated',
      'Arial, "fallback" "extra"',
      'Arial, inherit',
      'Arial, url(https://example.test/font)',
      'Arial,,serif',
    ]) {
      expect(documentScriptFontFamilyForRendering(fonts, 'ascii', forged)).toBe(
        'Arial',
      );
    }
  });

  test('resolves exact direct and theme slots across style inheritance', () => {
    const inherited = runProperties(
      `<w:rFonts w:ascii="Segoe UI" w:hAnsi="Calibri" w:eastAsia="SimSun" w:cs="Arial" w:hint="default"/>`,
    );
    const direct = runProperties(
      `<w:rFonts w:eastAsia="KaiTi" w:eastAsiaTheme="minorEastAsia" w:cstheme="majorBidi" w:hint="eastAsia"/>`,
      STRICT_WORD_NAMESPACE,
    );
    const fonts = resolveDocxRunFonts([inherited, direct], {
      colors: new Map(),
      fonts: new Map([
        ['minoreastasia', 'Microsoft YaHei'],
        ['majorbidi', 'Noto Naskh Arabic'],
      ]),
    });

    expect(fonts).toEqual({
      ascii: { name: 'Segoe UI', resolved: 'Segoe UI' },
      highAnsi: { name: 'Calibri', resolved: 'Calibri' },
      eastAsia: {
        name: 'KaiTi',
        theme: 'minorEastAsia',
        resolved: 'Microsoft YaHei',
      },
      complexScript: {
        theme: 'majorBidi',
        resolved: 'Noto Naskh Arabic',
      },
      hint: 'eastAsia',
    });
  });

  test('fails closed for duplicated, child-bearing, extra, and namespace-spoofed rFonts', () => {
    const invalid = [
      runProperties(
        '<w:rFonts w:ascii="Arial"/><w:rFonts w:eastAsia="SimSun"/>',
      ),
      runProperties('<w:rFonts w:ascii="Arial"><w:b/></w:rFonts>'),
      runProperties('<w:rFonts w:ascii="Arial" w:future="value"/>'),
      runProperties('<w:rFonts xmlns:evil="urn:spoofed" evil:ascii="Arial"/>'),
      runProperties('<evil:rFonts xmlns:evil="urn:spoofed" w:ascii="Arial"/>'),
    ];

    for (const properties of invalid) {
      expect(inspectDocxRunFonts(properties).status).toBe('invalid');
      expect(
        resolveDocxRunFonts([properties], {
          colors: new Map(),
          fonts: new Map(),
        }),
      ).toBeNull();
    }
  });

  test('keeps newly typed mixed-script text on the authoritative slot in one undo event', () => {
    const fonts = normalizeDocumentScriptFonts({
      ascii: { name: 'Segoe UI', resolved: 'Segoe UI' },
      highAnsi: { name: 'Calibri', resolved: 'Calibri' },
      eastAsia: { name: 'Microsoft YaHei', resolved: 'Microsoft YaHei' },
      complexScript: { name: 'Arial', resolved: 'Arial' },
    });
    const span = document.createElement('span');
    for (const [name, value] of Object.entries(
      documentScriptFontsDomAttributes(fonts, 'ascii'),
    )) {
      span.setAttribute(name, value);
    }
    span.textContent = 'A';
    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: `<p>${span.outerHTML}</p>`,
    });
    editor.commands.setTextSelection(2);

    expect(editor.commands.insertContent('中')).toBe(true);
    const html = new DOMParser().parseFromString(editor.getHTML(), 'text/html');
    expect(
      Array.from(
        html.querySelectorAll<HTMLElement>('[data-office-script-font-slot]'),
      ).map((element) => ({
        family: element.style.fontFamily,
        slot: element.dataset.officeScriptFontSlot,
        text: element.textContent,
      })),
    ).toEqual([
      {
        family: '"Segoe UI", Calibri, "Microsoft YaHei", Arial',
        slot: 'ascii',
        text: 'A',
      },
      {
        family: '"Microsoft YaHei", Calibri, "Segoe UI", Arial',
        slot: 'eastAsia',
        text: '中',
      },
    ]);

    expect(editor.commands.undo()).toBe(true);
    expect(editor.getText()).toBe('A');
    editor.destroy();
  });

  test('exports exact direct, theme, and hint attributes in every editable Word story', async () => {
    const fonts = normalizeDocumentScriptFonts({
      ascii: {
        name: 'Segoe UI',
        theme: 'minorAscii',
        resolved: 'Aptos',
      },
      highAnsi: {
        name: 'Calibri',
        theme: 'minorHAnsi',
        resolved: 'Aptos',
      },
      eastAsia: {
        name: 'SimSun',
        theme: 'minorEastAsia',
        resolved: 'Microsoft YaHei',
      },
      complexScript: {
        name: 'Arial',
        theme: 'minorBidi',
        resolved: 'Noto Naskh Arabic',
      },
      hint: 'eastAsia',
    });
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document' || !fonts) {
      throw new Error('Expected a document artifact and valid script fonts.');
    }
    const marked = (text: string, slot: WorkDocumentScriptFontSlot) => {
      const span = document.createElement('span');
      for (const [name, value] of Object.entries(
        documentScriptFontsDomAttributes(fonts, slot),
      )) {
        span.setAttribute(name, value);
      }
      span.textContent = text;
      return span.outerHTML;
    };
    artifact.content.html = [
      `<p>${marked('Body Latin', 'ascii')}</p>`,
      '<p>Notes',
      '<sup data-document-note-reference="true" data-note-kind="footnote" data-note-id="font-foot" data-note-number="1">1</sup>',
      '<sup data-document-note-reference="true" data-note-kind="endnote" data-note-id="font-end" data-note-number="1">1</sup></p>',
      `<aside data-document-note="true" data-note-kind="footnote" data-note-id="font-foot" data-note-number="1"><p>${marked('Footnote Asian', 'eastAsia')}</p></aside>`,
      `<aside data-document-note="true" data-note-kind="endnote" data-note-id="font-end" data-note-number="1"><p>${marked('Endnote complex', 'complexScript')}</p></aside>`,
    ].join('');
    artifact.content.pageChrome = {
      differentFirstPage: false,
      differentOddEvenPages: false,
      default: {
        headerHtml: `<p>${marked('Header high ANSI', 'highAnsi')}</p>`,
        footerHtml: `<p>${marked('Footer Asian', 'eastAsia')}</p>`,
        showPageNumber: false,
      },
      first: { headerHtml: '', footerHtml: '', showPageNumber: false },
      even: { headerHtml: '', footerHtml: '', showPageNumber: false },
    };

    const archive = await JSZip.loadAsync(
      await (await createDocxBlob(artifact.content)).arrayBuffer(),
    );
    for (const [path, text] of [
      ['word/document.xml', 'Body Latin'],
      ['word/header1.xml', 'Header high ANSI'],
      ['word/footer1.xml', 'Footer Asian'],
      ['word/footnotes.xml', 'Footnote Asian'],
      ['word/endnotes.xml', 'Endnote complex'],
    ] as const) {
      const xml = await archive.file(path)?.async('text');
      if (!xml) throw new Error(`Expected ${path}.`);
      const documentXml = parseXml(xml, path);
      const run = descendants(documentXml, 'r').find(
        (candidate) => directChild(candidate, 't')?.textContent === text,
      );
      const properties = run ? directChild(run, 'rPr') : null;
      const runFonts = properties ? directChild(properties, 'rFonts') : null;
      if (!runFonts) throw new Error(`Expected native fonts for ${text}.`);
      expect(wordFontAttributes(runFonts)).toEqual({
        ascii: 'Segoe UI',
        asciiTheme: 'minorAscii',
        cs: 'Arial',
        cstheme: 'minorBidi',
        eastAsia: 'SimSun',
        eastAsiaTheme: 'minorEastAsia',
        hAnsi: 'Calibri',
        hAnsiTheme: 'minorHAnsi',
        hint: 'eastAsia',
      });
    }
  });

  test('exports and reopens exact script fonts in formatting revisions', async () => {
    const beforeFonts = normalizeDocumentScriptFonts({
      ascii: {
        name: 'Before Latin',
        theme: 'majorAscii',
        resolved: 'Before Theme Latin',
      },
      highAnsi: {
        name: 'Before ANSI',
        theme: 'majorHAnsi',
        resolved: 'Before Theme ANSI',
      },
      eastAsia: {
        name: 'Before East',
        theme: 'majorEastAsia',
        resolved: 'Before Theme East',
      },
      complexScript: {
        name: 'Before Complex',
        theme: 'majorBidi',
        resolved: 'Before Theme Complex',
      },
      hint: 'cs',
    });
    const afterFonts = normalizeDocumentScriptFonts({
      ascii: {
        name: 'After Latin',
        theme: 'minorAscii',
        resolved: 'After Theme Latin',
      },
      highAnsi: {
        name: 'After ANSI',
        theme: 'minorHAnsi',
        resolved: 'After Theme ANSI',
      },
      eastAsia: {
        name: 'After East',
        theme: 'minorEastAsia',
        resolved: 'After Theme East',
      },
      complexScript: {
        name: 'After Complex',
        theme: 'minorBidi',
        resolved: 'After Theme Complex',
      },
      hint: 'eastAsia',
    });
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document' || !beforeFonts || !afterFonts) {
      throw new Error('Expected a document artifact and valid script fonts.');
    }
    const before = importedDocumentCharacterFormatting({
      fontFamily: documentScriptFontFamily(beforeFonts, 'ascii'),
      scriptFonts: beforeFonts,
      scriptFontSlot: 'ascii',
    });
    const current = document.createElement('span');
    for (const [name, value] of Object.entries(
      documentScriptFontsDomAttributes(afterFonts, 'ascii'),
    )) {
      current.setAttribute(name, value);
    }
    current.textContent = 'Changed fonts';
    const change = document.createElement('span');
    change.dataset.documentChange = 'true';
    change.dataset.changeKind = 'formatting';
    change.dataset.changeBefore = before;
    change.dataset.changeId = 'script-font-change';
    change.dataset.changeAuthor = 'Ada Reviewer';
    change.dataset.changeDate = '2026-08-23T12:00:00.000Z';
    change.append(current);
    artifact.content.html = `<p>${change.outerHTML}</p>`;
    artifact.content.trackChanges = true;

    const blob = await createDocxBlob(artifact.content);
    const archive = await JSZip.loadAsync(await blob.arrayBuffer());
    const xml = parseXml(
      (await archive.file('word/document.xml')?.async('text')) ?? '',
      'word/document.xml',
    );
    const run = descendants(xml, 'r').find(
      (candidate) =>
        directChild(candidate, 't')?.textContent === 'Changed fonts',
    );
    if (!run) throw new Error('Expected the changed-font run.');
    const currentProperties = directChild(run, 'rPr');
    const currentFonts = currentProperties
      ? directChild(currentProperties, 'rFonts')
      : null;
    const revision = descendants(run, 'rPrChange')[0];
    const previousProperties = revision ? directChild(revision, 'rPr') : null;
    const previousFonts = previousProperties
      ? directChild(previousProperties, 'rFonts')
      : null;
    if (!currentFonts || !previousFonts) {
      throw new Error('Expected current and previous native font slots.');
    }
    expect(wordFontAttributes(currentFonts)).toEqual({
      ascii: 'After Latin',
      asciiTheme: 'minorAscii',
      cs: 'After Complex',
      cstheme: 'minorBidi',
      eastAsia: 'After East',
      eastAsiaTheme: 'minorEastAsia',
      hAnsi: 'After ANSI',
      hAnsiTheme: 'minorHAnsi',
      hint: 'eastAsia',
    });
    expect(wordFontAttributes(previousFonts)).toEqual({
      ascii: 'Before Latin',
      asciiTheme: 'majorAscii',
      cs: 'Before Complex',
      cstheme: 'majorBidi',
      eastAsia: 'Before East',
      eastAsiaTheme: 'majorEastAsia',
      hAnsi: 'Before ANSI',
      hAnsiTheme: 'majorHAnsi',
      hint: 'cs',
    });

    const reopened = await importOfficeFile(
      new File([blob], 'script-font-revision.docx', { type: blob.type }),
    );
    if (reopened.content.type !== 'document') {
      throw new Error('Expected a reopened document artifact.');
    }
    expect(reopened.content.html).toContain('data-change-kind="formatting"');
    expect(reopened.content.html).toContain('After Latin');
    expect(reopened.content.html).toContain('Before Latin');
    expect(reopened.content.html).toContain('majorEastAsia');
  });
});

function runProperties(children: string, namespace = WORD_NAMESPACE): Element {
  return parseXml(`<w:rPr xmlns:w="${namespace}">${children}</w:rPr>`)
    .documentElement;
}

function wordFontAttributes(element: Element): Record<string, string> {
  return Object.fromEntries(
    [
      'ascii',
      'asciiTheme',
      'cs',
      'cstheme',
      'eastAsia',
      'eastAsiaTheme',
      'hAnsi',
      'hAnsiTheme',
      'hint',
    ].flatMap((name) => {
      const value = attribute(element, name);
      return value === null ? [] : [[name, value]];
    }),
  );
}
