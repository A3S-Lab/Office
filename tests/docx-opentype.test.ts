import { describe, expect, test } from '@rstest/core';
import JSZip from 'jszip';
import {
  createArtifact,
  createArtifactBlob,
  importOfficeFile,
} from '../src/core';
import { importedDocumentCharacterFormatting } from '../src/internal/features/work/work-document-format-changes';
import {
  documentOpenTypeDomAttributes,
  parseDocumentOpenTypeFeatures,
} from '../src/internal/features/work/work-document-opentype';
import { resolveDocxOpenTypeFeatures } from '../src/internal/features/work/work-docx-opentype-import';
import { patchDocxOpenTypeFeatures } from '../src/internal/features/work/work-docx-opentype-export';
import {
  applyImportedDocxRunFormattingMarkers,
  markDocxRunFormatting,
} from '../src/internal/features/work/work-docx-run-formatting-import';
import {
  attribute,
  descendants,
  directChild,
  directChildren,
  parseXml,
} from '../src/internal/features/work/work-ooxml-package';

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const WORD_2010_NAMESPACE =
  'http://schemas.microsoft.com/office/word/2010/wordml';

describe('DOCX OpenType typography', () => {
  test('resolves independently inherited Office 2010 run properties', () => {
    const base = runProperties(`
      <w14:ligatures w14:val="standardContextual"/>
      <w14:numForm w14:val="oldStyle"/>
      <w14:numSpacing w14:val="proportional"/>
      <w14:stylisticSets>
        <w14:styleSet w14:id="4"/>
        <w14:styleSet w14:id="1" w14:val="1"/>
      </w14:stylisticSets>
      <w14:cntxtAlts w14:val="1"/>
    `);
    const direct = runProperties(`
      <w14:numForm w14:val="default"/>
      <w14:stylisticSets/>
      <w14:cntxtAlts w14:val="0"/>
    `);

    expect(resolveDocxOpenTypeFeatures([base, direct])).toEqual({
      features: {
        ligatures: 'standardContextual',
        numberForm: 'default',
        numberSpacing: 'proportional',
        stylisticSets: [],
        contextualAlternates: false,
      },
      invalidCount: 0,
      spoofedCount: 0,
    });
  });

  test('canonicalizes enabled stylistic sets and strict lexical booleans', () => {
    const properties = runProperties(`
      <w14:stylisticSets>
        <w14:styleSet w14:id=" +04 "/>
        <w14:styleSet w14:id="1" w14:val="true"/>
        <w14:styleSet w14:id="4" w14:val="0"/>
        <w14:styleSet w14:id="20" w14:val="false"/>
      </w14:stylisticSets>
      <w14:cntxtAlts/>
    `);

    expect(resolveDocxOpenTypeFeatures([properties])).toEqual({
      features: { stylisticSets: [4, 1], contextualAlternates: true },
      invalidCount: 0,
      spoofedCount: 0,
    });
  });

  test('fails closed for malformed, duplicated, and namespace-spoofed values', () => {
    const malformed = runProperties(`
      <w14:ligatures w14:val="standard"/>
      <w14:ligatures w14:val="contextual"/>
      <w14:cntxtAlts w14:val="on"/>
      <evil:numForm evil:val="lining"/>
    `);

    expect(resolveDocxOpenTypeFeatures([malformed])).toEqual({
      features: null,
      invalidCount: 2,
      spoofedCount: 1,
    });
  });

  test('projects imported body-run features into editable text style metadata', () => {
    const source = wordDocument(`
      <w:p><w:r><w:rPr>
        <w14:ligatures w14:val="all"/>
        <w14:numForm w14:val="oldStyle"/>
        <w14:numSpacing w14:val="tabular"/>
        <w14:stylisticSets><w14:styleSet w14:id="4"/></w14:stylisticSets>
        <w14:cntxtAlts w14:val="0"/>
      </w:rPr><w:t>Office 0123</w:t></w:r></w:p>
    `);
    const markers = markDocxRunFormatting(source);
    const marker = markers.runs[0];
    if (!marker) throw new Error('Missing body-run marker.');
    const html = new DOMParser().parseFromString(
      `<p>${marker.startMarker}Office 0123${marker.endMarker}</p>`,
      'text/html',
    );

    applyImportedDocxRunFormattingMarkers(html, markers);

    const span = html.querySelector('[data-office-opentype-features]');
    expect(span?.getAttribute('data-office-opentype-features')).toBe(
      '{"ligatures":"all","numberForm":"oldStyle","numberSpacing":"tabular","stylisticSets":[4],"contextualAlternates":false}',
    );
    expect(span?.getAttribute('style')).toContain('font-feature-settings');
    expect(span?.textContent).toBe('Office 0123');
  });

  test('exports and reopens exact native body-run OpenType properties', async () => {
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    const features =
      '{&quot;ligatures&quot;:&quot;standardContextual&quot;,&quot;numberForm&quot;:&quot;oldStyle&quot;,&quot;numberSpacing&quot;:&quot;tabular&quot;,&quot;stylisticSets&quot;:[4,1],&quot;contextualAlternates&quot;:false}';
    artifact.content.html = `<section data-document-section="true"><p><span data-office-opentype-features="${features}">Office 0123</span></p></section>`;

    const blob = await createArtifactBlob(artifact);
    const archive = await JSZip.loadAsync(await blob.arrayBuffer());
    const xml = (await archive.file('word/document.xml')?.async('text')) ?? '';

    expect(xml).toContain(`xmlns:w14="${WORD_2010_NAMESPACE}"`);
    expect(xml).toMatch(/mc:Ignorable="[^"]*\bw14\b[^"]*"/);
    expect(xml).toMatch(
      /<w14:ligatures w14:val="standardContextual"\/><w14:numForm w14:val="oldStyle"\/><w14:numSpacing w14:val="tabular"\/><w14:stylisticSets><w14:styleSet w14:id="4"\/><w14:styleSet w14:id="1"\/><\/w14:stylisticSets><w14:cntxtAlts w14:val="0"\/>/,
    );
    expect(xml).not.toContain('A3SOfficeOpenType');

    const reopened = await importOfficeFile(
      new File([blob], 'opentype-body.docx', { type: blob.type }),
    );
    if (reopened.content.type !== 'document') {
      throw new Error('Expected a reopened document artifact.');
    }
    expect(reopened.content.html).toContain(
      'data-office-opentype-features="{&quot;ligatures&quot;:&quot;standardContextual&quot;,&quot;numberForm&quot;:&quot;oldStyle&quot;,&quot;numberSpacing&quot;:&quot;tabular&quot;,&quot;stylisticSets&quot;:[4,1],&quot;contextualAlternates&quot;:false}"',
    );
  });

  test('exports and reopens native OpenType features in body, header, and footer', async () => {
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    artifact.content.html = `<p>${openTypeSpan(
      { ligatures: 'standardContextual' },
      'Body typography',
    )}</p>`;
    artifact.content.pageChrome = {
      differentFirstPage: false,
      differentOddEvenPages: false,
      default: {
        headerHtml: `<p>${openTypeSpan(
          { numberForm: 'oldStyle', stylisticSets: [4] },
          'Header typography',
        )}</p>`,
        footerHtml: `<p>${openTypeSpan(
          { numberSpacing: 'tabular', contextualAlternates: false },
          'Footer typography',
        )}</p>`,
        showPageNumber: false,
      },
      first: { headerHtml: '', footerHtml: '', showPageNumber: false },
      even: { headerHtml: '', footerHtml: '', showPageNumber: false },
    };

    const blob = await createArtifactBlob(artifact);
    const archive = await JSZip.loadAsync(await blob.arrayBuffer());
    const parts = await Promise.all(
      Object.keys(archive.files)
        .filter((path) =>
          /^word\/(?:document|header\d*|footer\d*)\.xml$/i.test(path),
        )
        .map(async (path) => (await archive.file(path)?.async('text')) ?? ''),
    );
    const xml = parts.join('\n');
    expect(xml).toContain('<w14:ligatures w14:val="standardContextual"/>');
    expect(xml).toContain('<w14:numForm w14:val="oldStyle"/>');
    expect(xml).toContain('<w14:styleSet w14:id="4"/>');
    expect(xml).toContain('<w14:numSpacing w14:val="tabular"/>');
    expect(xml).toContain('<w14:cntxtAlts w14:val="0"/>');

    const reopened = await importOfficeFile(
      new File([blob], 'opentype-stories.docx', { type: blob.type }),
    );
    if (reopened.content.type !== 'document') {
      throw new Error('Expected a reopened document artifact.');
    }
    expect(reopened.content.html).toContain('standardContextual');
    expect(reopened.content.pageChrome?.default.headerHtml).toContain(
      'oldStyle',
    );
    expect(reopened.content.pageChrome?.default.headerHtml).toContain(
      'stylisticSets',
    );
    expect(reopened.content.pageChrome?.default.footerHtml).toContain(
      'tabular',
    );
    expect(reopened.content.pageChrome?.default.footerHtml).toContain(
      'contextualAlternates',
    );
  });

  test('patches native OpenType properties across every supported Word story', async () => {
    const paths = [
      'word/document.xml',
      'word/header1.xml',
      'word/footer1.xml',
      'word/footnotes.xml',
      'word/endnotes.xml',
      'word/comments.xml',
    ];
    const source = new JSZip();
    for (const path of paths) {
      source.file(path, markerStory('OpenTypeMarker'));
    }
    const patched = await patchDocxOpenTypeFeatures(
      await source.generateAsync({ type: 'arraybuffer' }),
      [
        {
          marker: 'OpenTypeMarker',
          style: 'OriginalStyle',
          features: {
            ligatures: 'all',
            numberForm: 'lining',
            numberSpacing: 'proportional',
            stylisticSets: [1, 20],
            contextualAlternates: true,
          },
        },
      ],
    );
    const archive = await JSZip.loadAsync(patched);
    for (const path of paths) {
      const xml = (await archive.file(path)?.async('text')) ?? '';
      expect(xml).toContain('w:val="OriginalStyle"');
      expect(xml).toContain('<w14:ligatures w14:val="all"/>');
      expect(xml).toContain('<w14:styleSet w14:id="20"/>');
      expect(xml).toContain('<w14:cntxtAlts w14:val="1"/>');
      expect(xml).not.toContain('OpenTypeMarker');
    }
  });

  test('exports and reopens exact native OpenType formatting revisions', async () => {
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    const beforeFeatures = {
      ligatures: 'standard',
      numberForm: 'lining',
      numberSpacing: 'proportional',
      stylisticSets: [1],
      contextualAlternates: true,
    } as const;
    const currentFeatures = {
      ligatures: 'all',
      numberForm: 'oldStyle',
      numberSpacing: 'tabular',
      stylisticSets: [4, 20],
      contextualAlternates: false,
    } as const;
    const before = importedDocumentCharacterFormatting({
      openTypeFeatures: beforeFeatures,
    });
    const current = document.createElement('span');
    for (const [name, value] of Object.entries(
      documentOpenTypeDomAttributes(currentFeatures),
    )) {
      current.setAttribute(name, value);
    }
    current.textContent = 'Changed typography';
    const change = document.createElement('span');
    change.dataset.documentChange = 'true';
    change.dataset.changeKind = 'formatting';
    change.dataset.changeBefore = before;
    change.dataset.changeId = 'opentype-change';
    change.dataset.changeAuthor = 'Ada Reviewer';
    change.dataset.changeDate = '2026-09-01T01:30:00.000Z';
    change.append(current);
    artifact.content.html = `<p>${change.outerHTML}</p>`;
    artifact.content.trackChanges = true;

    const blob = await createArtifactBlob(artifact);
    const archive = await JSZip.loadAsync(await blob.arrayBuffer());
    const documentXml = parseXml(
      (await archive.file('word/document.xml')?.async('text')) ?? '',
    );
    const run = descendants(documentXml, 'r').find(
      (candidate) =>
        directChild(candidate, 't')?.textContent === 'Changed typography',
    );
    if (!run) throw new Error('Expected changed typography run.');
    const properties = directChild(run, 'rPr');
    const revision = properties ? directChild(properties, 'rPrChange') : null;
    const prior = revision ? directChild(revision, 'rPr') : null;
    if (!properties || !revision || !prior) {
      throw new Error('Expected current and prior OpenType properties.');
    }
    expect(openTypePropertyValues(properties)).toEqual({
      ligatures: 'all',
      numForm: 'oldStyle',
      numSpacing: 'tabular',
      stylisticSets: ['4', '20'],
      cntxtAlts: '0',
    });
    expect(openTypePropertyValues(prior)).toEqual({
      ligatures: 'standard',
      numForm: 'lining',
      numSpacing: 'proportional',
      stylisticSets: ['1'],
      cntxtAlts: '1',
    });

    const reopened = await importOfficeFile(
      new File([blob], 'opentype-revision.docx', { type: blob.type }),
    );
    if (reopened.content.type !== 'document') {
      throw new Error('Expected a reopened document artifact.');
    }
    const reopenedDocument = new DOMParser().parseFromString(
      reopened.content.html,
      'text/html',
    );
    const reopenedChange = reopenedDocument.querySelector<HTMLElement>(
      '[data-document-change][data-change-kind="formatting"]',
    );
    const reopenedCurrent = reopenedChange?.querySelector<HTMLElement>(
      '[data-office-opentype-features]',
    );
    if (!reopenedChange || !reopenedCurrent) {
      throw new Error('Expected reopened OpenType formatting revision.');
    }
    expect(
      parseDocumentOpenTypeFeatures(
        reopenedCurrent.dataset.officeOpentypeFeatures,
      ),
    ).toEqual(currentFeatures);
    const previousMarks = JSON.parse(
      reopenedChange.dataset.changeBefore ?? 'null',
    ) as Array<{
      type?: string;
      attrs?: { openTypeFeatures?: string };
    }>;
    const previousTextStyle = previousMarks.find(
      (mark) => mark.type === 'textStyle',
    );
    expect(
      parseDocumentOpenTypeFeatures(previousTextStyle?.attrs?.openTypeFeatures),
    ).toEqual({
      ligatures: 'standard',
      numberForm: 'lining',
      numberSpacing: 'proportional',
      stylisticSets: [1],
      contextualAlternates: true,
    });
  });
});

function openTypePropertyValues(properties: Element): {
  ligatures: string | null;
  numForm: string | null;
  numSpacing: string | null;
  stylisticSets: string[];
  cntxtAlts: string | null;
} {
  const ligatures = directChild(properties, 'ligatures');
  const numForm = directChild(properties, 'numForm');
  const numSpacing = directChild(properties, 'numSpacing');
  const stylisticSets = directChild(properties, 'stylisticSets');
  const cntxtAlts = directChild(properties, 'cntxtAlts');
  return {
    ligatures: ligatures ? attribute(ligatures, 'val') : null,
    numForm: numForm ? attribute(numForm, 'val') : null,
    numSpacing: numSpacing ? attribute(numSpacing, 'val') : null,
    stylisticSets: stylisticSets
      ? directChildren(stylisticSets, 'styleSet').map((styleSet) =>
          attribute(styleSet, 'id'),
        )
      : [],
    cntxtAlts: cntxtAlts ? attribute(cntxtAlts, 'val') : null,
  };
}

function openTypeSpan(
  features: Parameters<typeof documentOpenTypeDomAttributes>[0],
  text: string,
): string {
  const span = document.createElement('span');
  for (const [name, value] of Object.entries(
    documentOpenTypeDomAttributes(features),
  )) {
    span.setAttribute(name, value);
  }
  span.textContent = text;
  return span.outerHTML;
}

function markerStory(marker: string): string {
  return `<w:root xmlns:w="${WORD_NAMESPACE}"><w:p><w:r><w:rPr><w:rStyle w:val="${marker}"/></w:rPr><w:t>Story</w:t></w:r></w:p></w:root>`;
}

function runProperties(children: string): Element {
  const document = wordDocument(`
    <w:p><w:r><w:rPr>${children}</w:rPr></w:r></w:p>
  `);
  const properties = directChildren(document.documentElement, 'body')
    .flatMap((body) => directChildren(body, 'p'))
    .flatMap((paragraph) => directChildren(paragraph, 'r'))
    .flatMap((run) => directChildren(run, 'rPr'))[0];
  if (!properties) throw new Error('Missing run properties.');
  return properties;
}

function wordDocument(body: string): Document {
  return parseXml(`
    <w:document xmlns:w="${WORD_NAMESPACE}"
      xmlns:w14="${WORD_2010_NAMESPACE}"
      xmlns:evil="https://example.test/evil">
      <w:body>${body}</w:body>
    </w:document>
  `);
}
