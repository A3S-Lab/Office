import { describe, expect, test } from '@rstest/core';
import { Editor } from '@tiptap/core';
import JSZip from 'jszip';
import {
  createArtifact,
  createArtifactBlob,
  importOfficeFile,
} from '../src/core';
import { createDocumentPageChromeEditorExtensions } from '../src/internal/features/work/editors/document-page-chrome-editor';
import {
  createDocumentEquationElement,
  documentEquationFromElement,
  normalizeDocumentEquation,
  type WorkDocumentEquation,
} from '../src/internal/features/work/work-document-equations';
import { createWorkDocumentExtensions } from '../src/internal/features/work/work-document-extensions';
import { sanitizeDocumentPageChromeHtml } from '../src/internal/features/work/work-document-page-chrome';
import {
  inspectDocxEquation,
  isSupportedDocxEquationPlacement,
} from '../src/internal/features/work/work-docx-equation-import';
import {
  descendants,
  directChildren,
  parseXml,
} from '../src/internal/features/work/work-ooxml-package';

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const STRICT_WORD_NAMESPACE =
  'http://purl.oclc.org/ooxml/wordprocessingml/main';
const MATH_NAMESPACE =
  'http://schemas.openxmlformats.org/officeDocument/2006/math';
const STRICT_MATH_NAMESPACE = 'http://purl.oclc.org/ooxml/officeDocument/math';
const RELATIONSHIP_NAMESPACE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const XMLNS_NAMESPACE = 'http://www.w3.org/2000/xmlns/';
const VENDOR_NAMESPACE = 'urn:a3s:test:spoofed-equation';

describe('document equations', () => {
  test('keeps a bounded structured model editable and renders accessible MathML', () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: '<p>Equation: </p>',
    });
    try {
      editor.commands.setTextSelection(editor.state.doc.content.size - 1);
      expect(
        editor.commands.insertDocumentEquation(complexEquation('inline')),
      ).toBe(true);
      const document = new DOMParser().parseFromString(
        editor.getHTML(),
        'text/html',
      );
      const element = document.body.querySelector<HTMLElement>(
        '[data-document-equation]',
      );
      expect(element).not.toBeNull();
      expect(documentEquationFromElement(element as HTMLElement)).toEqual(
        complexEquation('inline'),
      );
      expect(element?.querySelector('math')).not.toBeNull();
      expect(element?.querySelector('mfrac')).not.toBeNull();
      expect(element?.querySelector('munderover')).not.toBeNull();
      const accent = element?.querySelector('mover[accent="true"]');
      expect(accent).not.toBeNull();
      expect(accent?.querySelector('mo')?.textContent).toBe('\u02dc');
      const overbar = element?.querySelector('mover[accent="false"]');
      expect(overbar).not.toBeNull();
      expect(overbar?.querySelector('mo')?.textContent).toBe('\u00af');
      const underbar = element?.querySelector('munder[accentunder="false"]');
      expect(underbar).not.toBeNull();
      expect(underbar?.querySelector('mo')?.textContent).toBe('\u00af');
      expect(element?.querySelector('mtable')).not.toBeNull();
      expect(element?.querySelectorAll('mtr')).toHaveLength(2);
      expect(element?.querySelectorAll('mtd')).toHaveLength(8);
      expect(element).toHaveAttribute('role', 'math');
      expect(element?.getAttribute('aria-label')).toContain('sqrt');
      expect(element?.getAttribute('aria-label')).toContain('accent(U+0303');
      expect(element?.getAttribute('aria-label')).toContain('overbar(x+y)');
      expect(element?.getAttribute('aria-label')).toContain('underbar(a-b)');
      expect(element?.getAttribute('aria-label')).toContain('matrix');

      const equationPosition = documentEquationPosition(editor);
      editor.commands.setNodeSelection(equationPosition);
      expect(
        editor.commands.updateDocumentEquation(simpleEquation('updated')),
      ).toBe(true);
      const updated = new DOMParser()
        .parseFromString(editor.getHTML(), 'text/html')
        .body.querySelector<HTMLElement>('[data-document-equation]');
      expect(documentEquationFromElement(updated as HTMLElement)).toEqual(
        simpleEquation('updated'),
      );

      expect(normalizeDocumentEquation(deepEquation(34))).toBeNull();
      expect(
        normalizeDocumentEquation({
          version: 1,
          display: 'inline',
          children: [{ type: 'run', text: 'x'.repeat(65_537) }],
        }),
      ).toBeNull();
      expect(
        normalizeDocumentEquation({
          version: 1,
          display: 'inline',
          children: [
            {
              type: 'bar',
              position: 'middle',
              children: [{ type: 'run', text: 'x' }],
            },
          ],
        }),
      ).toBeNull();
      expect(
        normalizeDocumentEquation({
          version: 1,
          display: 'inline',
          children: [
            {
              type: 'accent',
              character: '^',
              children: [{ type: 'run', text: 'x' }],
            },
          ],
        }),
      ).toBeNull();
      expect(
        normalizeDocumentEquation({
          version: 1,
          display: 'inline',
          children: [
            {
              type: 'matrix',
              baseAlignment: 'center',
              placeholdersHidden: false,
              columnAlignments: ['center', 'center'],
              rows: [
                [[{ type: 'run', text: 'a' }], [{ type: 'run', text: 'b' }]],
                [[{ type: 'run', text: 'c' }]],
              ],
            },
          ],
        }),
      ).toBeNull();
      expect(
        normalizeDocumentEquation({
          version: 1,
          display: 'inline',
          children: [
            {
              type: 'matrix',
              baseAlignment: 'center',
              placeholdersHidden: false,
              columnAlignments: ['center'],
              rows: Array.from({ length: 65 }, () => [[]]),
            },
          ],
        }),
      ).toBeNull();
      expect(
        normalizeDocumentEquation({
          version: 1,
          display: 'inline',
          children: [{ type: 'run', text: '\u0000' }],
        }),
      ).toBeNull();
      expect(
        normalizeDocumentEquation({
          version: 1,
          display: 'inline',
          children: Array.from({ length: 4_097 }, () => ({
            type: 'run',
            text: 'x',
          })),
        }),
      ).toBeNull();
    } finally {
      editor.destroy();
    }
  });

  test('keeps equation atoms structured while editing page chrome', () => {
    const editor = new Editor({
      extensions: createDocumentPageChromeEditorExtensions(),
      content: `<p>Header ${equationHtml(complexEquation('inline'))}</p>`,
    });
    try {
      let equations = 0;
      editor.state.doc.descendants((node) => {
        if (node.type.name === 'documentEquation') equations += 1;
      });
      expect(equations).toBe(1);
      expect(editor.getHTML()).toContain('data-document-equation="true"');
      expect(editor.getHTML()).toContain('<math');
      expect(editor.getHTML()).toContain('accent="true"');
      expect(editor.getHTML()).toContain('accentunder="false"');
      expect(editor.getHTML()).toContain('<mtable');

      const sanitized = new DOMParser().parseFromString(
        sanitizeDocumentPageChromeHtml(
          `${equationHtml(simpleEquation('Safe'))}<span data-document-equation="true" data-equation-model="{">Broken</span><math xmlns="http://www.w3.org/1998/Math/MathML"><mtext>Loose</mtext></math>`,
        ),
        'text/html',
      );
      expect(
        sanitized.body.querySelectorAll('span[data-document-equation]'),
      ).toHaveLength(1);
      expect(sanitized.body.querySelectorAll('math')).toHaveLength(1);
      expect(sanitized.body.textContent).toContain('Broken');
      expect(sanitized.body.textContent).toContain('Loose');
    } finally {
      editor.destroy();
    }
  });

  test('round-trips the structured OMML subset through every Word story twice', async () => {
    const artifact = equationArtifact();
    const first = await createArtifactBlob(artifact);
    await expectNativeEquations(first);

    const imported = await importOfficeFile(
      new File([first], 'structured-equations.docx', { type: first.type }),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const body = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    const bodyEquations = Array.from(
      body.body.querySelectorAll<HTMLElement>('[data-document-equation]'),
    );
    expect(
      bodyEquations.map(
        (element) => documentEquationFromElement(element)?.children.length,
      ),
    ).toEqual([18, 18, 1, 1]);
    expect(bodyEquations.map(documentEquationFromElement).every(Boolean)).toBe(
      true,
    );
    expect(documentEquationFromElement(bodyEquations[0])).toEqual(
      complexEquation('inline'),
    );
    expect(documentEquationFromElement(bodyEquations[1])).toEqual(
      complexEquation('block'),
    );
    expect(documentEquationFromElement(bodyEquations[2])?.children[0]).toEqual({
      type: 'bar',
      position: 'top',
      children: [{ type: 'run', text: 'F' }],
    });
    expect(documentEquationFromElement(bodyEquations[3])?.children[0]).toEqual({
      type: 'bar',
      position: 'bottom',
      children: [{ type: 'run', text: 'N' }],
    });
    expect(
      bodyEquations.some(
        (element) => documentEquationFromElement(element)?.display === 'block',
      ),
    ).toBe(true);
    expect(imported.content.pageChrome?.default.headerHtml).toContain(
      'data-document-equation="true"',
    );
    expect(imported.content.pageChrome?.default.footerHtml).toContain(
      'data-document-equation="true"',
    );
    expect(imported.compatibility.issues).toContainEqual(
      expect.objectContaining({ code: 'docx.equations', severity: 'info' }),
    );
    expect(imported.compatibility.issues).not.toContainEqual(
      expect.objectContaining({ code: 'docx.equations.unsupported' }),
    );

    await expectNativeEquations(await createArtifactBlob(imported));
  });

  test('imports strict UTF-16 body and endnote equations into transitional output', async () => {
    const artifact = equationArtifact();
    artifact.content.html += `<p>${STRICT_MATH_NAMESPACE}</p>`;
    const seed = await createArtifactBlob(artifact);
    const archive = await JSZip.loadAsync(await seed.arrayBuffer());
    for (const path of ['word/document.xml', 'word/endnotes.xml']) {
      const source = await archive.file(path)?.async('string');
      if (!source) throw new Error(`Expected ${path}.`);
      archive.file(
        path,
        strictUtf16(
          source
            .replaceAll(WORD_NAMESPACE, STRICT_WORD_NAMESPACE)
            .replaceAll(MATH_NAMESPACE, STRICT_MATH_NAMESPACE),
        ),
      );
    }
    const bytes = await archive.generateAsync({ type: 'arraybuffer' });
    const imported = await importOfficeFile(
      new File([bytes], 'strict-utf16-equations.docx'),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    expect(
      new DOMParser()
        .parseFromString(imported.content.html, 'text/html')
        .body.querySelectorAll('[data-document-equation]'),
    ).toHaveLength(4);
    expect(imported.content.html).toContain(STRICT_MATH_NAMESPACE);

    const output = await JSZip.loadAsync(
      await (await createArtifactBlob(imported)).arrayBuffer(),
    );
    const document = await xmlEntry(output, 'word/document.xml');
    expect(descendants(document, 'oMath')).not.toHaveLength(0);
    expect(
      descendants(document, 'oMath').every(
        (equation) => equation.namespaceURI === MATH_NAMESPACE,
      ),
    ).toBe(true);
  });

  test('fails closed for unsupported, relationship-bound, and spoofed equation markup', async () => {
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    artifact.content.html = [
      `<p>${equationHtml(simpleEquation('A'))}</p>`,
      `<p>${equationHtml(simpleEquation('B'))}</p>`,
      `<p>${equationHtml(simpleEquation('C'))}</p>`,
      `<p>${equationHtml(simpleEquation('D'))}</p>`,
    ].join('');
    artifact.content.pageChrome = {
      differentFirstPage: false,
      differentOddEvenPages: false,
      default: {
        headerHtml: `<p>${equationHtml(simpleEquation('H'))}</p>`,
        footerHtml: `<p>${equationHtml(simpleEquation('J'))}</p>`,
        showPageNumber: false,
      },
      first: { headerHtml: '', footerHtml: '', showPageNumber: false },
      even: { headerHtml: '', footerHtml: '', showPageNumber: false },
    };
    const seed = await createArtifactBlob(artifact);
    const archive = await JSZip.loadAsync(await seed.arrayBuffer());
    const document = await xmlEntry(archive, 'word/document.xml');
    const equations = descendants(document, 'oMath');
    expect(equations).toHaveLength(4);

    const spoofed = document.createElementNS(VENDOR_NAMESPACE, 'v:oMath');
    spoofed.append(
      ...Array.from(equations[0].childNodes, (node) => node.cloneNode(true)),
    );
    document.documentElement.setAttributeNS(
      XMLNS_NAMESPACE,
      'xmlns:v',
      VENDOR_NAMESPACE,
    );
    equations[0].replaceWith(spoofed);

    const matrix = document.createElementNS(MATH_NAMESPACE, 'm:m');
    const matrixProperties = document.createElementNS(MATH_NAMESPACE, 'm:mPr');
    const columnSpacing = document.createElementNS(MATH_NAMESPACE, 'm:cSp');
    columnSpacing.setAttributeNS(MATH_NAMESPACE, 'm:val', '120');
    matrixProperties.append(columnSpacing);
    const row = document.createElementNS(MATH_NAMESPACE, 'm:mr');
    const argument = document.createElementNS(MATH_NAMESPACE, 'm:e');
    argument.append(...Array.from(equations[1].childNodes));
    row.append(argument);
    matrix.append(matrixProperties, row);
    equations[1].replaceChildren(matrix);

    const run = directChildren(equations[2], 'r')[0];
    const properties = document.createElementNS(MATH_NAMESPACE, 'm:rPr');
    properties.setAttributeNS(RELATIONSHIP_NAMESPACE, 'r:embed', 'rIdUnsafe');
    run?.insertBefore(properties, run.firstChild);
    expect(inspectDocxEquation(equations[2]).status).toBe('unsupported');

    const paragraph = equations[3].parentElement;
    const misplacedRun = document.createElementNS(WORD_NAMESPACE, 'w:r');
    equations[3].replaceWith(misplacedRun);
    misplacedRun.append(equations[3]);
    expect(paragraph?.localName).toBe('p');

    const headerPath = matchingXmlPath(archive, /^word\/header\d*\.xml$/i);
    const header = await xmlEntry(archive, headerPath);
    const headerEquation = descendants(header, 'oMath')[0];
    const headerHyperlink = header.createElementNS(
      WORD_NAMESPACE,
      'w:hyperlink',
    );
    headerEquation.replaceWith(headerHyperlink);
    headerHyperlink.append(headerEquation);
    archive.file(headerPath, new XMLSerializer().serializeToString(header));

    const footerPath = matchingXmlPath(archive, /^word\/footer\d*\.xml$/i);
    const footer = await xmlEntry(archive, footerPath);
    const footerEquation = descendants(footer, 'oMath')[0];
    const spoofedFooter = footer.createElementNS(VENDOR_NAMESPACE, 'v:oMath');
    spoofedFooter.append(
      ...Array.from(footerEquation.childNodes, (node) => node.cloneNode(true)),
    );
    footer.documentElement.setAttributeNS(
      XMLNS_NAMESPACE,
      'xmlns:v',
      VENDOR_NAMESPACE,
    );
    footerEquation.replaceWith(spoofedFooter);
    archive.file(footerPath, new XMLSerializer().serializeToString(footer));
    expect(isSupportedDocxEquationPlacement(equations[3])).toBe(false);

    archive.file(
      'word/document.xml',
      new XMLSerializer().serializeToString(document),
    );
    const source = await archive.generateAsync({ type: 'arraybuffer' });
    const reparsed = await xmlEntry(
      await JSZip.loadAsync(source),
      'word/document.xml',
    );
    expect(
      descendants(reparsed, 'oMath').map(
        (equation) => inspectDocxEquation(equation).status,
      ),
    ).toEqual(['spoofed', 'unsupported', 'unsupported', 'supported']);
    expect(
      isSupportedDocxEquationPlacement(descendants(reparsed, 'oMath')[3]),
    ).toBe(false);
    const imported = await importOfficeFile(
      new File([source], 'unsafe-equations.docx'),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const html = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    expect(html.body.querySelector('[data-document-equation]')).toBeNull();
    expect(html.body.textContent).toContain('A');
    expect(html.body.textContent).toContain('B');
    expect(html.body.textContent).toContain('C');
    expect(html.body.textContent).toContain('D');
    expect(imported.content.pageChrome?.default.headerHtml).not.toContain(
      'data-document-equation',
    );
    expect(imported.content.pageChrome?.default.headerHtml).toContain('H');
    expect(imported.content.pageChrome?.default.footerHtml).not.toContain(
      'data-document-equation',
    );
    expect(imported.content.pageChrome?.default.footerHtml).toContain('J');
    expect(imported.compatibility.issues).toContainEqual(
      expect.objectContaining({ code: 'docx.equations.unsupported' }),
    );

    const output = await JSZip.loadAsync(
      await (await createArtifactBlob(imported)).arrayBuffer(),
    );
    for (const path of ['word/document.xml', headerPath, footerPath]) {
      expect(descendants(await xmlEntry(output, path), 'oMath')).toHaveLength(
        0,
      );
    }
  });

  test('rejects unmodeled and contradictory OMML properties', () => {
    const run = '<m:r><m:t>x</m:t></m:r>';
    const supported = [
      `<m:acc><m:e>${run}</m:e></m:acc>`,
      `<m:acc><m:accPr><m:chr m:val="&#x20D7;"/><m:ctrlPr/></m:accPr><m:e>${run}</m:e></m:acc>`,
      `<m:bar><m:e>${run}</m:e></m:bar>`,
      `<m:bar><m:barPr/><m:e>${run}</m:e></m:bar>`,
      `<m:bar><m:barPr><m:pos/><m:ctrlPr/></m:barPr><m:e>${run}</m:e></m:bar>`,
      `<m:bar><m:barPr><m:pos m:val="top"/><m:ctrlPr/></m:barPr><m:e>${run}</m:e></m:bar>`,
      `<m:f><m:fPr><m:type m:val="noBar"/><m:ctrlPr/></m:fPr><m:num>${run}</m:num><m:den>${run}</m:den></m:f>`,
      `<m:rad><m:radPr><m:degHide m:val="1"/><m:ctrlPr/></m:radPr><m:e>${run}</m:e></m:rad>`,
      `<m:nary><m:naryPr><m:chr m:val="&#x2211;"/><m:limLoc m:val="undOvr"/><m:subHide m:val="1"/><m:supHide m:val="true"/><m:ctrlPr/></m:naryPr><m:e>${run}</m:e></m:nary>`,
      `<m:m><m:mr><m:e>${run}</m:e></m:mr></m:m>`,
      `<m:m><m:mPr><m:baseJc m:val="bot"/><m:plcHide m:val="true"/><m:mcs><m:mc><m:mcPr><m:count m:val="2"/><m:mcJc m:val="left"/></m:mcPr></m:mc><m:mc><m:mcPr><m:count/><m:mcJc/></m:mcPr></m:mc></m:mcs><m:ctrlPr/></m:mPr><m:mr><m:e>${run}</m:e><m:e>${run}</m:e><m:e/></m:mr></m:m>`,
    ];
    expect(supported.map(inspectEquationBody)).toEqual(
      supported.map(() => 'supported'),
    );
    expect(inspectEquationModel(supported[0])?.children[0]).toEqual({
      type: 'accent',
      character: '\u0302',
      children: [{ type: 'run', text: 'x' }],
    });
    expect(inspectEquationModel(supported[1])?.children[0]).toEqual({
      type: 'accent',
      character: '\u20d7',
      children: [{ type: 'run', text: 'x' }],
    });
    expect(
      supported
        .slice(2, 6)
        .map((source) => inspectEquationModel(source)?.children[0]),
    ).toEqual([
      {
        type: 'bar',
        position: 'top',
        children: [{ type: 'run', text: 'x' }],
      },
      {
        type: 'bar',
        position: 'bottom',
        children: [{ type: 'run', text: 'x' }],
      },
      {
        type: 'bar',
        position: 'bottom',
        children: [{ type: 'run', text: 'x' }],
      },
      {
        type: 'bar',
        position: 'top',
        children: [{ type: 'run', text: 'x' }],
      },
    ]);

    const unsupported = [
      `<m:acc><m:accPr><m:chr/></m:accPr><m:e>${run}</m:e></m:acc>`,
      `<m:acc><m:accPr><m:chr m:val="^"/></m:accPr><m:e>${run}</m:e></m:acc>`,
      `<m:acc><m:accPr><m:chr m:val="&#x302;&#x303;"/></m:accPr><m:e>${run}</m:e></m:acc>`,
      `<m:acc><m:accPr><m:grow m:val="1"/></m:accPr><m:e>${run}</m:e></m:acc>`,
      `<m:acc><m:e>${run}</m:e><m:accPr><m:chr m:val="&#x302;"/></m:accPr></m:acc>`,
      `<m:acc><m:accPr><m:chr m:val="&#x302;"/><m:chr m:val="&#x303;"/></m:accPr><m:e>${run}</m:e></m:acc>`,
      `<m:acc><m:accPr><m:chr m:val="&#x302;"/></m:accPr><m:e/></m:acc>`,
      `<m:bar><m:barPr><m:pos m:val="left"/></m:barPr><m:e>${run}</m:e></m:bar>`,
      `<m:bar><m:barPr><m:pos m:val="bottom"/></m:barPr><m:e>${run}</m:e></m:bar>`,
      `<m:bar><m:barPr><m:grow m:val="1"/></m:barPr><m:e>${run}</m:e></m:bar>`,
      `<m:bar><m:e>${run}</m:e><m:barPr><m:pos m:val="top"/></m:barPr></m:bar>`,
      `<m:bar><m:barPr><m:ctrlPr/><m:pos m:val="top"/></m:barPr><m:e>${run}</m:e></m:bar>`,
      `<m:bar><m:barPr><m:pos m:val="top"/><m:pos m:val="bot"/></m:barPr><m:e>${run}</m:e></m:bar>`,
      `<m:bar><m:barPr><m:pos m:val="top" m:extra="semantic"/></m:barPr><m:e>${run}</m:e></m:bar>`,
      `<m:bar><v:barPr xmlns:v="${VENDOR_NAMESPACE}"><m:pos m:val="top"/></v:barPr><m:e>${run}</m:e></m:bar>`,
      `<m:bar><m:barPr><m:pos xmlns:r="${RELATIONSHIP_NAMESPACE}" m:val="top" r:id="rIdUnsafe"/></m:barPr><m:e>${run}</m:e></m:bar>`,
      `<m:bar><m:barPr><m:pos m:val="top"/></m:barPr><m:e/></m:bar>`,
      `<m:bar><m:e>${run}</m:e><m:e>${run}</m:e></m:bar>`,
      `<m:f><m:fPr><m:m><m:mr><m:e>${run}</m:e></m:mr></m:m></m:fPr><m:num>${run}</m:num><m:den>${run}</m:den></m:f>`,
      `<m:f><m:fPr><m:type m:val="bar" m:extra="semantic"/></m:fPr><m:num>${run}</m:num><m:den>${run}</m:den></m:f>`,
      `<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg>${run}</m:deg><m:e>${run}</m:e></m:rad>`,
      `<m:rad><m:radPr/><m:e>${run}</m:e></m:rad>`,
      `<m:nary><m:naryPr><m:chr m:val="&#x2211;"/><m:subHide m:val="on"/></m:naryPr><m:sub>${run}</m:sub><m:e>${run}</m:e></m:nary>`,
      `<m:nary><m:naryPr><m:chr m:val="&#x2211;"/></m:naryPr><m:e>${run}</m:e></m:nary>`,
      `<m:d><m:dPr><m:grow m:val="1"/></m:dPr><m:e>${run}</m:e></m:d>`,
      `<m:m><m:mPr><m:cSp m:val="120"/></m:mPr><m:mr><m:e>${run}</m:e></m:mr></m:m>`,
      `<m:m><m:mr><m:e>${run}</m:e><m:e>${run}</m:e></m:mr><m:mr><m:e>${run}</m:e></m:mr></m:m>`,
      `<m:m><m:mPr><m:mcs><m:mc><m:mcPr><m:count m:val="1"/><m:mcJc m:val="left"/></m:mcPr></m:mc></m:mcs></m:mPr><m:mr><m:e>${run}</m:e><m:e>${run}</m:e></m:mr></m:m>`,
      `<m:m><m:mPr><m:mcs><m:mc><m:mcPr><m:count m:val="0"/></m:mcPr></m:mc></m:mcs></m:mPr><m:mr><m:e>${run}</m:e></m:mr></m:m>`,
      `<m:m>${Array.from({ length: 65 }, () => `<m:mr><m:e>${run}</m:e></m:mr>`).join('')}</m:m>`,
      '<m:r><m:rPr><m:sty m:val="b"/></m:rPr><m:t>x</m:t></m:r>',
      `<m:rPr/>${run}`,
      deepOmml(34),
      `<m:r><m:t>${'x'.repeat(65_537)}</m:t></m:r>`,
      `<m:d>${Array.from({ length: 33 }, () => `<m:e>${run}</m:e>`).join('')}</m:d>`,
    ];
    expect(unsupported.map(inspectEquationBody)).toEqual(
      unsupported.map(() => 'unsupported'),
    );
  });
});

function equationArtifact() {
  const artifact = createArtifact('blank-document');
  if (artifact.content.type !== 'document') {
    throw new Error('Expected a document artifact.');
  }
  const inline = equationHtml(complexEquation('inline'));
  const block = equationHtml(complexEquation('block'));
  artifact.content.html = [
    `<p>Inline ${inline} equation</p>`,
    `<p>${block}</p>`,
    '<p>Notes',
    '<sup data-document-note-reference="true" data-note-kind="footnote" data-note-id="foot-equation" data-note-number="1">1</sup>',
    '<sup data-document-note-reference="true" data-note-kind="endnote" data-note-id="end-equation" data-note-number="1">1</sup>',
    '</p>',
    `<aside data-document-note="true" data-note-kind="footnote" data-note-id="foot-equation" data-note-number="1"><p>Foot ${equationHtml(barEquation('F', 'top'))}</p></aside>`,
    `<aside data-document-note="true" data-note-kind="endnote" data-note-id="end-equation" data-note-number="1"><p>End ${equationHtml(barEquation('N', 'bottom'))}</p></aside>`,
  ].join('');
  artifact.content.pageChrome = {
    differentFirstPage: false,
    differentOddEvenPages: false,
    default: {
      headerHtml: `<p>Header ${equationHtml(barEquation('H', 'top'))}</p>`,
      footerHtml: `<p>Footer ${equationHtml(barEquation('R', 'bottom'))}</p>`,
      showPageNumber: false,
    },
    first: { headerHtml: '', footerHtml: '', showPageNumber: false },
    even: { headerHtml: '', footerHtml: '', showPageNumber: false },
  };
  return artifact;
}

function complexEquation(
  display: WorkDocumentEquation['display'],
): WorkDocumentEquation {
  const run = (text: string) => ({ type: 'run' as const, text });
  return {
    version: 1,
    display,
    children: [
      run('x='),
      {
        type: 'fraction',
        fractionType: 'bar',
        numerator: [run('a+b')],
        denominator: [run('c')],
      },
      {
        type: 'fraction',
        fractionType: 'noBar',
        numerator: [run('n')],
        denominator: [run('k')],
      },
      {
        type: 'fraction',
        fractionType: 'skewed',
        numerator: [run('p')],
        denominator: [run('q')],
      },
      {
        type: 'fraction',
        fractionType: 'linear',
        numerator: [run('u')],
        denominator: [run('v')],
      },
      { type: 'superscript', base: [run('x')], superScript: [run('2')] },
      { type: 'subscript', base: [run('x')], subScript: [run('i')] },
      {
        type: 'subSuperScript',
        base: [run('x')],
        subScript: [run('i')],
        superScript: [run('n')],
      },
      { type: 'radical', children: [run('y')] },
      { type: 'radical', children: [run('z')], degree: [run('3')] },
      { type: 'function', name: [run('sin')], children: [run('θ')] },
      {
        type: 'nary',
        operator: '∑',
        limitLocation: 'underOver',
        children: [run('x_i')],
        subScript: [run('i=1')],
        superScript: [run('n')],
      },
      {
        type: 'nary',
        operator: '∫',
        limitLocation: 'subSup',
        children: [run('f(x)dx')],
        subScript: [run('0')],
        superScript: [run('1')],
      },
      {
        type: 'delimiter',
        opening: '[',
        closing: ']',
        separator: ';',
        arguments: [[run('a')], [run('b')]],
      },
      {
        type: 'accent',
        character: '\u0303',
        children: [run('x+y')],
      },
      {
        type: 'bar',
        position: 'top',
        children: [run('x+y')],
      },
      {
        type: 'bar',
        position: 'bottom',
        children: [run('a-b')],
      },
      {
        type: 'matrix',
        baseAlignment: 'top',
        placeholdersHidden: false,
        columnAlignments: ['left', 'center', 'center', 'right'],
        rows: [
          [[run('a')], [run('b')], [run('c')], [run('d')]],
          [[run('e')], [run('f')], [run('g')], []],
        ],
      },
    ],
  };
}

function simpleEquation(text: string): WorkDocumentEquation {
  return {
    version: 1,
    display: 'inline',
    children: [{ type: 'run', text }],
  };
}

function barEquation(
  text: string,
  position: 'top' | 'bottom',
): WorkDocumentEquation {
  return {
    version: 1,
    display: 'inline',
    children: [
      {
        type: 'bar',
        position,
        children: [{ type: 'run', text }],
      },
    ],
  };
}

function equationHtml(equation: WorkDocumentEquation): string {
  const document = new DOMParser().parseFromString('', 'text/html');
  return createDocumentEquationElement(document, equation).outerHTML;
}

function inspectEquationBody(body: string): string {
  return inspectEquation(body).status;
}

function inspectEquationModel(body: string): WorkDocumentEquation | null {
  return inspectEquation(body).equation;
}

function inspectEquation(body: string) {
  const document = parseXml(
    `<m:oMath xmlns:m="${MATH_NAMESPACE}">${body}</m:oMath>`,
    'equation.xml',
  );
  return inspectDocxEquation(document.documentElement);
}

function deepEquation(depth: number): unknown {
  let expression: unknown = { type: 'run', text: 'x' };
  for (let index = 0; index < depth; index += 1) {
    expression = {
      type: 'radical',
      children: [expression],
    };
  }
  return { version: 1, display: 'inline', children: [expression] };
}

function deepOmml(depth: number): string {
  let expression = '<m:r><m:t>x</m:t></m:r>';
  for (let index = 0; index < depth; index += 1) {
    expression = `<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:e>${expression}</m:e></m:rad>`;
  }
  return expression;
}

function documentEquationPosition(editor: Editor): number {
  let result = -1;
  editor.state.doc.descendants((node, position) => {
    if (result < 0 && node.type.name === 'documentEquation') result = position;
  });
  if (result < 0) throw new Error('Expected a document equation node.');
  return result;
}

async function expectNativeEquations(blob: Blob): Promise<void> {
  const archive = await JSZip.loadAsync(await blob.arrayBuffer());
  const document = await xmlEntry(archive, 'word/document.xml');
  expect(descendants(document, 'oMathPara')).toHaveLength(1);
  expect(descendants(document, 'oMath')).toHaveLength(2);
  expect(descendants(document, 'f')).toHaveLength(8);
  expect(descendants(document, 'sSup')).toHaveLength(2);
  expect(descendants(document, 'sSub')).toHaveLength(2);
  expect(descendants(document, 'sSubSup')).toHaveLength(2);
  expect(descendants(document, 'rad')).toHaveLength(4);
  expect(descendants(document, 'func')).toHaveLength(2);
  expect(descendants(document, 'nary')).toHaveLength(4);
  expect(descendants(document, 'd')).toHaveLength(2);
  const accents = descendants(document, 'acc');
  expect(accents).toHaveLength(2);
  for (const accent of accents) {
    const properties = directChildren(accent, 'accPr')[0];
    expect(mathValueAttribute(directChildren(properties, 'chr')[0])).toBe(
      '\u0303',
    );
    expect(directChildren(accent, 'e')).toHaveLength(1);
  }
  const bars = descendants(document, 'bar');
  expect(bars).toHaveLength(4);
  expect(
    bars.map((bar) =>
      mathValueAttribute(
        directChildren(directChildren(bar, 'barPr')[0], 'pos')[0],
      ),
    ),
  ).toEqual(['top', 'bot', 'top', 'bot']);
  expect(bars.every((bar) => directChildren(bar, 'e').length === 1)).toBe(true);
  const matrices = descendants(document, 'm');
  expect(matrices).toHaveLength(2);
  for (const matrix of matrices) {
    const rows = directChildren(matrix, 'mr');
    expect(rows).toHaveLength(2);
    expect(rows.map((row) => directChildren(row, 'e').length)).toEqual([4, 4]);
    expect(directChildren(rows[1], 'e')[3]?.childElementCount).toBe(0);
    const properties = directChildren(matrix, 'mPr')[0];
    expect(mathValueAttribute(directChildren(properties, 'baseJc')[0])).toBe(
      'top',
    );
    expect(mathValueAttribute(directChildren(properties, 'plcHide')[0])).toBe(
      '0',
    );
    const columns = descendants(properties, 'mc');
    expect(
      columns.map((column) =>
        mathValueAttribute(descendants(column, 'count')[0]),
      ),
    ).toEqual(['1', '2', '1']);
    expect(
      columns.map((column) =>
        mathValueAttribute(descendants(column, 'mcJc')[0]),
      ),
    ).toEqual(['left', 'center', 'right']);
  }
  const inline = descendants(document, 'oMath')[0];
  const componentStatuses = directChildren(inline).map((component) => {
    const equation = document.createElementNS(MATH_NAMESPACE, 'm:oMath');
    equation.append(component.cloneNode(true));
    return `${component.localName}:${inspectDocxEquation(equation).status}`;
  });
  expect(
    componentStatuses.every((status) => status.endsWith(':supported')),
  ).toBe(true);

  const stories = [
    {
      document: await xmlEntry(archive, 'word/footnotes.xml'),
      position: 'top',
    },
    {
      document: await xmlEntry(archive, 'word/endnotes.xml'),
      position: 'bot',
    },
    {
      document: await matchingXmlEntry(archive, /^word\/header\d*\.xml$/i),
      position: 'top',
    },
    {
      document: await matchingXmlEntry(archive, /^word\/footer\d*\.xml$/i),
      position: 'bot',
    },
  ];
  for (const story of stories) {
    expect(descendants(story.document, 'oMath')).toHaveLength(1);
    expect(descendants(story.document, 'oMath')[0]?.namespaceURI).toBe(
      MATH_NAMESPACE,
    );
    const bar = descendants(story.document, 'bar')[0];
    expect(bar).toBeDefined();
    expect(
      mathValueAttribute(
        directChildren(directChildren(bar, 'barPr')[0], 'pos')[0],
      ),
    ).toBe(story.position);
  }
}

function mathValueAttribute(element: Element | undefined): string | null {
  return (
    Array.from(element?.attributes ?? []).find(
      (attribute) =>
        attribute.localName === 'val' || /(?:^|:)val$/u.test(attribute.name),
    )?.value ?? null
  );
}

async function matchingXmlEntry(
  archive: JSZip,
  pattern: RegExp,
): Promise<Document> {
  return xmlEntry(archive, matchingXmlPath(archive, pattern));
}

function matchingXmlPath(archive: JSZip, pattern: RegExp): string {
  const path = Object.keys(archive.files).find((candidate) =>
    pattern.test(candidate),
  );
  if (!path) throw new Error(`Expected a package part matching ${pattern}.`);
  return path;
}

async function xmlEntry(archive: JSZip, path: string): Promise<Document> {
  const source = await archive.file(path)?.async('string');
  if (!source) throw new Error(`Expected ${path}.`);
  return parseXml(source, path);
}

function strictUtf16(source: string): Uint8Array {
  const normalized = source.replace(/^\s*<\?xml[^>]*\?>/i, '');
  const value = `<?xml version="1.0" encoding="UTF-16" standalone="yes"?>${normalized}`;
  const bytes = new Uint8Array(2 + value.length * 2);
  bytes[0] = 0xff;
  bytes[1] = 0xfe;
  const view = new DataView(bytes.buffer);
  for (let index = 0; index < value.length; index += 1) {
    view.setUint16(2 + index * 2, value.charCodeAt(index), true);
  }
  return bytes;
}
