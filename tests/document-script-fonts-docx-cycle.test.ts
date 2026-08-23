import { expect, test } from '@rstest/core';
import { Editor } from '@tiptap/core';
import JSZip from 'jszip';
import { createArtifact, importOfficeFile } from '../src/core';
import { applyDocumentFontDialogPatch } from '../src/internal/features/work/editors/document-font-dialog-model';
import {
  parseDocumentScriptFonts,
  type WorkDocumentScriptFonts,
} from '../src/internal/features/work/work-document-script-fonts';
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
const DRAWING_NAMESPACE =
  'http://schemas.openxmlformats.org/drawingml/2006/main';
const DOCX_MIME =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

test('imports inherited theme fonts, edits independent script slots, exports, and reopens them', async () => {
  const sourceBytes = await inheritedThemeFontFixture();
  const imported = await importOfficeFile(
    new File([sourceBytes], 'inherited-script-fonts.docx', {
      type: DOCX_MIME,
    }),
  );
  if (imported.content.type !== 'document') {
    throw new Error('Expected an imported document artifact.');
  }

  const importedSpans = scriptFontSpans(imported.content.html);
  expect(importedSpans.map(({ slot, text }) => ({ slot, text }))).toEqual([
    { slot: 'ascii', text: 'A' },
    { slot: 'highAnsi', text: 'é' },
    { slot: 'eastAsia', text: '中' },
    { slot: 'complexScript', text: 'ह' },
  ]);
  expect(importedSpans.map(({ firstFamily }) => firstFamily)).toEqual([
    'Aptos Display',
    'Aptos Display',
    'MS Mincho',
    'Noto Serif Arabic',
  ]);
  expect(importedSpans[0]?.fonts).toEqual(inheritedThemeFonts());

  const editor = new Editor({
    extensions: createWorkDocumentExtensions(),
    content: imported.content.html,
  });
  editor.commands.selectAll();
  const selection = {
    from: editor.state.selection.from,
    to: editor.state.selection.to,
  };
  expect(
    applyDocumentFontDialogPatch(editor, selection, {
      eastAsiaFont: 'Project East',
      complexScriptFont: 'Project Complex',
    }),
  ).toBe(true);
  imported.content.html = editor.getHTML();
  editor.destroy();

  const exported = await createDocxBlob(imported.content, sourceBytes);
  const archive = await JSZip.loadAsync(await exported.arrayBuffer());
  const documentXml = parseXml(
    (await archive.file('word/document.xml')?.async('text')) ?? '',
    'word/document.xml',
  );
  const editedRun = descendants(documentXml, 'r').find(
    (run) => directChild(run, 't')?.textContent === '中',
  );
  const properties = editedRun ? directChild(editedRun, 'rPr') : null;
  const fonts = properties ? directChild(properties, 'rFonts') : null;
  if (!fonts) throw new Error('Expected edited native run fonts.');
  expect(wordFontAttributes(fonts)).toEqual({
    asciiTheme: 'majorAscii',
    cstheme: undefined,
    cs: 'Project Complex',
    eastAsia: 'Project East',
    eastAsiaTheme: undefined,
    hAnsiTheme: 'majorHAnsi',
    hint: 'eastAsia',
  });

  const reopened = await importOfficeFile(
    new File([exported], 'reopened-script-fonts.docx', { type: DOCX_MIME }),
  );
  if (reopened.content.type !== 'document') {
    throw new Error('Expected a reopened document artifact.');
  }
  const reopenedSpans = scriptFontSpans(reopened.content.html);
  expect(reopenedSpans.map(({ firstFamily }) => firstFamily)).toEqual([
    'Aptos Display',
    'Aptos Display',
    'Project East',
    'Project Complex',
  ]);
  for (const { fonts: reopenedFonts } of reopenedSpans) {
    expect(reopenedFonts).toEqual({
      ascii: { theme: 'majorAscii', resolved: 'Aptos Display' },
      highAnsi: { theme: 'majorHAnsi', resolved: 'Aptos Display' },
      eastAsia: { name: 'Project East', resolved: 'Project East' },
      complexScript: {
        name: 'Project Complex',
        resolved: 'Project Complex',
      },
      hint: 'eastAsia',
    });
  }
});

async function inheritedThemeFontFixture(): Promise<ArrayBuffer> {
  const artifact = createArtifact('blank-document');
  if (artifact.content.type !== 'document') {
    throw new Error('Expected a document artifact.');
  }
  artifact.content.html = '<p>Font cycle source</p>';
  const archive = await JSZip.loadAsync(
    await (await createDocxBlob(artifact.content)).arrayBuffer(),
  );
  const documentSource = await archive.file('word/document.xml')?.async('text');
  if (!documentSource) throw new Error('Expected generated document XML.');
  const documentXml = parseXml(documentSource, 'word/document.xml');
  const paragraph = descendants(documentXml, 'p').find((candidate) =>
    candidate.textContent?.includes('Font cycle source'),
  );
  const run = paragraph
    ? descendants(paragraph, 'r').find((candidate) =>
        candidate.textContent?.includes('Font cycle source'),
      )
    : undefined;
  const text = run ? directChild(run, 't') : undefined;
  if (!paragraph || !run || !text) {
    throw new Error('Expected generated paragraph and run.');
  }
  text.textContent = 'Aé中ह';

  const paragraphProperties =
    directChild(paragraph, 'pPr') ??
    documentXml.createElementNS(WORD_NAMESPACE, 'w:pPr');
  if (!paragraphProperties.parentElement) {
    paragraph.prepend(paragraphProperties);
  }
  directChild(paragraphProperties, 'pStyle')?.remove();
  const paragraphStyle = wordElement(documentXml, 'pStyle');
  paragraphStyle.setAttributeNS(WORD_NAMESPACE, 'w:val', 'DerivedFonts');
  paragraphProperties.prepend(paragraphStyle);

  const runProperties =
    directChild(run, 'rPr') ??
    documentXml.createElementNS(WORD_NAMESPACE, 'w:rPr');
  if (!runProperties.parentElement) run.prepend(runProperties);
  directChild(runProperties, 'rStyle')?.remove();
  directChild(runProperties, 'rFonts')?.remove();
  const characterStyle = wordElement(documentXml, 'rStyle');
  characterStyle.setAttributeNS(WORD_NAMESPACE, 'w:val', 'AccentFonts');
  const directFonts = wordElement(documentXml, 'rFonts');
  directFonts.setAttributeNS(WORD_NAMESPACE, 'w:cstheme', 'majorBidi');
  directFonts.setAttributeNS(WORD_NAMESPACE, 'w:hint', 'eastAsia');
  runProperties.prepend(characterStyle);
  runProperties.append(directFonts);
  archive.file(
    'word/document.xml',
    new XMLSerializer().serializeToString(documentXml),
  );
  archive.file('word/styles.xml', stylesXml());
  archive.file('word/theme/theme1.xml', themeXml());
  return archive.generateAsync({ type: 'arraybuffer' });
}

function stylesXml(): string {
  return `<w:styles xmlns:w="${WORD_NAMESPACE}">
    <w:docDefaults><w:rPrDefault><w:rPr>
      <w:rFonts w:asciiTheme="minorAscii" w:hAnsiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia" w:cstheme="minorBidi"/>
    </w:rPr></w:rPrDefault></w:docDefaults>
    <w:style w:type="paragraph" w:styleId="BaseFonts">
      <w:rPr><w:rFonts w:asciiTheme="majorAscii"/></w:rPr>
    </w:style>
    <w:style w:type="paragraph" w:styleId="DerivedFonts">
      <w:basedOn w:val="BaseFonts"/>
      <w:rPr><w:rFonts w:hAnsiTheme="majorHAnsi"/></w:rPr>
    </w:style>
    <w:style w:type="character" w:styleId="BaseCharacters">
      <w:rPr><w:rFonts w:eastAsia="KaiTi"/></w:rPr>
    </w:style>
    <w:style w:type="character" w:styleId="AccentFonts">
      <w:basedOn w:val="BaseCharacters"/>
      <w:rPr><w:rFonts w:eastAsia="KaiTi" w:eastAsiaTheme="majorEastAsia"/></w:rPr>
    </w:style>
  </w:styles>`;
}

function themeXml(): string {
  return `<a:theme xmlns:a="${DRAWING_NAMESPACE}" name="Script fonts">
    <a:themeElements><a:fontScheme name="Script fonts">
      <a:majorFont>
        <a:latin typeface="Aptos Display"/><a:ea typeface="MS Mincho"/><a:cs typeface="Noto Serif Arabic"/>
      </a:majorFont>
      <a:minorFont>
        <a:latin typeface="Aptos"/><a:ea typeface="Microsoft YaHei"/><a:cs typeface="Noto Naskh Arabic"/>
      </a:minorFont>
    </a:fontScheme></a:themeElements>
  </a:theme>`;
}

function inheritedThemeFonts(): WorkDocumentScriptFonts {
  return {
    ascii: { theme: 'majorAscii', resolved: 'Aptos Display' },
    highAnsi: { theme: 'majorHAnsi', resolved: 'Aptos Display' },
    eastAsia: {
      name: 'KaiTi',
      theme: 'majorEastAsia',
      resolved: 'MS Mincho',
    },
    complexScript: {
      theme: 'majorBidi',
      resolved: 'Noto Serif Arabic',
    },
    hint: 'eastAsia',
  };
}

function scriptFontSpans(html: string): Array<{
  firstFamily: string | null;
  fonts: WorkDocumentScriptFonts | null;
  slot: string | undefined;
  text: string;
}> {
  const document = new DOMParser().parseFromString(html, 'text/html');
  return Array.from(
    document.querySelectorAll<HTMLElement>('[data-office-script-font-slot]'),
  )
    .filter((element) => Boolean(element.textContent))
    .map((element) => ({
      firstFamily: firstCssFamily(element.style.fontFamily),
      fonts: parseDocumentScriptFonts(element.dataset.officeScriptFonts),
      slot: element.dataset.officeScriptFontSlot,
      text: element.textContent ?? '',
    }));
}

function firstCssFamily(value: string): string | null {
  const family = value.split(',')[0]?.trim();
  return family ? family.replace(/^['"]|['"]$/g, '') : null;
}

function wordFontAttributes(
  element: Element,
): Record<string, string | undefined> {
  return {
    asciiTheme: attribute(element, 'asciiTheme') ?? undefined,
    cstheme: attribute(element, 'cstheme') ?? undefined,
    cs: attribute(element, 'cs') ?? undefined,
    eastAsia: attribute(element, 'eastAsia') ?? undefined,
    eastAsiaTheme: attribute(element, 'eastAsiaTheme') ?? undefined,
    hAnsiTheme: attribute(element, 'hAnsiTheme') ?? undefined,
    hint: attribute(element, 'hint') ?? undefined,
  };
}

function wordElement(document: Document, name: string): Element {
  return document.createElementNS(WORD_NAMESPACE, `w:${name}`);
}
