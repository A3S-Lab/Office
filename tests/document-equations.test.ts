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
  xmlAttributeLocalName,
  xmlAttributeNamespace,
} from '../src/internal/features/work/work-docx-settings-xml';
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
      const upperGroupCharacter = Array.from(
        element?.querySelectorAll('mover[accent="false"]') ?? [],
      ).find(
        (candidate) => candidate.lastElementChild?.textContent === '\u23de',
      );
      expect(upperGroupCharacter?.textContent).toBe('x+y\u23de');
      const lowerGroupCharacter = Array.from(
        element?.querySelectorAll('munder[accentunder="false"]') ?? [],
      ).find(
        (candidate) => candidate.lastElementChild?.textContent === '\u23df',
      );
      expect(lowerGroupCharacter?.textContent).toBe('a-b\u23df');
      const hiddenPhantom = element?.querySelector('mphantom');
      expect(hiddenPhantom?.textContent).toBe('ghost');
      const hiddenPhantomPadding = hiddenPhantom?.querySelector('mpadded');
      expect(hiddenPhantomPadding).toHaveAttribute('width', '0in');
      expect(hiddenPhantomPadding).not.toHaveAttribute('height');
      expect(hiddenPhantomPadding).toHaveAttribute('depth', '0in');
      const visiblePhantomPadding = element?.querySelector(
        'mpadded[height="0in"]',
      );
      expect(visiblePhantomPadding?.textContent).toBe('visible');
      expect(visiblePhantomPadding?.closest('mphantom')).toBeNull();
      const preScripts = Array.from(
        element?.querySelectorAll('mmultiscripts') ?? [],
      );
      expect(preScripts).toHaveLength(3);
      expect(
        preScripts.map((script) =>
          Array.from(script.children, (child) => child.tagName.toLowerCase()),
        ),
      ).toEqual([
        ['mrow', 'mprescripts', 'mrow', 'mrow'],
        ['mrow', 'mprescripts', 'none', 'mrow'],
        ['mrow', 'mprescripts', 'mrow', 'none'],
      ]);
      expect(preScripts.map((script) => script.textContent)).toEqual([
        'Tij',
        'A2',
        'B1',
      ]);
      expect(element?.querySelectorAll('mprescripts')).toHaveLength(3);
      expect(element?.querySelectorAll('mmultiscripts none')).toHaveLength(2);
      expect(
        element?.querySelector('mtext[mathvariant="bold-fraktur"]')
          ?.textContent,
      ).toBe('styledF');
      expect(
        element?.querySelector('mtext[mathvariant="normal"]')?.textContent,
      ).toBe('normalRate');
      expect(
        element?.querySelector('mtext[mathvariant="double-struck"]')
          ?.textContent,
      ).toBe('doubleR');
      const borderBox = element?.querySelector('menclose');
      expect(borderBox).toHaveAttribute(
        'notation',
        'top left horizontalstrike updiagonalstrike',
      );
      expect(element?.querySelector('mpadded')).not.toBeNull();
      const equationArray = element?.querySelector('mtable[align="bottom"]');
      expect(equationArray).toHaveAttribute('rowspacing', '1.5em');
      expect(equationArray?.querySelectorAll('mtr')).toHaveLength(3);
      expect(equationArray?.querySelectorAll('mtd')).toHaveLength(3);
      expect(equationArray?.querySelectorAll('maligngroup')).toHaveLength(4);
      expect(equationArray?.querySelectorAll('malignmark')).toHaveLength(4);
      const lowerLimit = Array.from(
        element?.querySelectorAll('munder[accentunder="false"]') ?? [],
      ).find((candidate) => candidate.textContent === 'limx→0');
      expect(lowerLimit).toBeDefined();
      const upperLimit = Array.from(
        element?.querySelectorAll('mover[accent="false"]') ?? [],
      ).find((candidate) => candidate.textContent === '=def');
      expect(upperLimit).toBeDefined();
      const matrix = element?.querySelector('mtable[align="top"]');
      expect(matrix).not.toBeNull();
      expect(matrix?.querySelectorAll('mtr')).toHaveLength(2);
      expect(matrix?.querySelectorAll('mtd')).toHaveLength(8);
      expect(element).toHaveAttribute('role', 'math');
      expect(element?.getAttribute('aria-label')).toContain('sqrt');
      expect(element?.getAttribute('aria-label')).toContain('accent(U+0303');
      expect(element?.getAttribute('aria-label')).toContain('overbar(x+y)');
      expect(element?.getAttribute('aria-label')).toContain('underbar(a-b)');
      expect(element?.getAttribute('aria-label')).toContain(
        'group-character(U+23DE;position=top;baseline=bottom;x+y)',
      );
      expect(element?.getAttribute('aria-label')).toContain(
        'group-character(U+23DF;position=bottom;baseline=top;a-b)',
      );
      expect(element?.getAttribute('aria-label')).toContain(
        'phantom(hidden,zero-width,zero-descent,transparent;ghost)',
      );
      expect(element?.getAttribute('aria-label')).toContain(
        'phantom(visible,zero-ascent;visible)',
      );
      expect(element?.getAttribute('aria-label')).toContain(
        'pre-scripts(sub=i;sup=j;base=T)',
      );
      expect(element?.getAttribute('aria-label')).toContain(
        'pre-scripts(sub=none;sup=2;base=A)',
      );
      expect(element?.getAttribute('aria-label')).toContain(
        'pre-scripts(sub=1;sup=none;base=B)',
      );
      expect(element?.getAttribute('aria-label')).toContain(
        'run(literal,script=fraktur,style=bold,break@3,alignment;styledF)',
      );
      expect(element?.getAttribute('aria-label')).toContain(
        'run(normal-text,style=plain;normalRate)',
      );
      expect(element?.getAttribute('aria-label')).toContain(
        'run(script=doubleStruck,style=boldItalic;doubleR)',
      );
      expect(element?.getAttribute('aria-label')).toContain(
        'borderbox(top left horizontalstrike updiagonalstrike;boxed)',
      );
      expect(element?.getAttribute('aria-label')).toContain(
        'box(operator,no-break,differential,break@3,alignment;dx)',
      );
      expect(element?.getAttribute('aria-label')).toContain(
        'equation-array(bottom,max-distribution,spacing=multiple:3;x+y=1;2x+y=3;)',
      );
      expect(element?.getAttribute('aria-label')).toContain(
        'lower-limit(lim;x→0)',
      );
      expect(element?.getAttribute('aria-label')).toContain(
        'upper-limit(=;def)',
      );
      expect(element?.getAttribute('aria-label')).toContain('matrix');
      const rightAlignedBlock = createDocumentEquationElement(
        document,
        complexEquation('block'),
      );
      expect(rightAlignedBlock).toHaveClass('block', 'justification-right');

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
          display: 'block',
          justification: 'centerGroup',
          children: [{ type: 'run', text: 'x' }],
        }),
      ).toEqual({
        version: 1,
        display: 'block',
        children: [{ type: 'run', text: 'x' }],
      });
      expect(
        normalizeDocumentEquation({
          version: 1,
          display: 'block',
          justification: 'left',
          children: [{ type: 'run', text: 'x' }],
        }),
      ).toEqual({
        version: 1,
        display: 'block',
        justification: 'left',
        children: [{ type: 'run', text: 'x' }],
      });
      expect(
        normalizeDocumentEquation({
          version: 1,
          display: 'inline',
          justification: 'right',
          children: [{ type: 'run', text: 'x' }],
        }),
      ).toBeNull();
      expect(
        normalizeDocumentEquation({
          version: 1,
          display: 'block',
          justification: 'justify',
          children: [{ type: 'run', text: 'x' }],
        }),
      ).toBeNull();
      expect(
        normalizeDocumentEquation({
          version: 1,
          display: 'inline',
          children: [
            {
              type: 'subSuperScript',
              alignScripts: false,
              base: [{ type: 'run', text: 'x' }],
              subScript: [{ type: 'run', text: 'i' }],
              superScript: [{ type: 'run', text: 'n' }],
            },
          ],
        }),
      ).toEqual({
        version: 1,
        display: 'inline',
        children: [
          {
            type: 'subSuperScript',
            base: [{ type: 'run', text: 'x' }],
            subScript: [{ type: 'run', text: 'i' }],
            superScript: [{ type: 'run', text: 'n' }],
          },
        ],
      });
      expect(
        normalizeDocumentEquation({
          version: 1,
          display: 'inline',
          children: [
            {
              type: 'subSuperScript',
              alignScripts: true,
              base: [{ type: 'run', text: 'x' }],
              subScript: [{ type: 'run', text: 'i' }],
              superScript: [{ type: 'run', text: 'n' }],
            },
          ],
        })?.children[0],
      ).toEqual({
        type: 'subSuperScript',
        alignScripts: true,
        base: [{ type: 'run', text: 'x' }],
        subScript: [{ type: 'run', text: 'i' }],
        superScript: [{ type: 'run', text: 'n' }],
      });
      expect(
        normalizeDocumentEquation({
          version: 1,
          display: 'inline',
          children: [
            {
              type: 'subSuperScript',
              alignScripts: 'yes',
              base: [{ type: 'run', text: 'x' }],
              subScript: [{ type: 'run', text: 'i' }],
              superScript: [{ type: 'run', text: 'n' }],
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
              type: 'run',
              text: 'x',
              literal: false,
              normalText: false,
              script: 'roman',
              style: 'italic',
              alignment: false,
            },
          ],
        }),
      ).toEqual(simpleEquation('x'));
      expect(
        normalizeDocumentEquation({
          version: 1,
          display: 'inline',
          children: [{ type: 'run', text: 'x', script: 'blackboard' }],
        }),
      ).toBeNull();
      expect(
        normalizeDocumentEquation({
          version: 1,
          display: 'inline',
          children: [
            { type: 'run', text: 'x', manualBreak: { alignmentAt: 256 } },
          ],
        }),
      ).toBeNull();
      expect(
        normalizeDocumentEquation({
          version: 1,
          display: 'inline',
          children: [
            {
              type: 'lowerLimit',
              base: [],
              limit: [{ type: 'run', text: 'x→0' }],
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
              type: 'groupCharacter',
              character: 'xy',
              position: 'top',
              verticalJustification: 'bottom',
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
              type: 'groupCharacter',
              character: '\u23de',
              position: 'middle',
              verticalJustification: 'bottom',
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
              type: 'groupCharacter',
              character: '\u23df',
              position: 'bottom',
              verticalJustification: 'center',
              children: [],
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
              type: 'phantom',
              show: 'false',
              zeroWidth: false,
              zeroAscent: false,
              zeroDescent: false,
              transparent: false,
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
              type: 'phantom',
              show: true,
              zeroWidth: false,
              zeroAscent: false,
              zeroDescent: false,
              transparent: false,
              children: [],
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
              type: 'preSubSuperScript',
              base: [],
              subScript: [{ type: 'run', text: 'i' }],
              superScript: [{ type: 'run', text: 'j' }],
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
              type: 'preSubSuperScript',
              base: [{ type: 'run', text: 'T' }],
              subScript: 'i',
              superScript: [],
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
              type: 'upperLimit',
              base: [{ type: 'run', text: '=' }],
              limit: [],
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
              type: 'equationArray',
              baseAlignment: 'center',
              maximumDistribution: false,
              objectDistribution: false,
              rowSpacingRule: 'single',
              rowSpacing: 0,
              rows: [[{ type: 'run', text: '&'.repeat(4_096) }]],
            },
          ],
        }),
      ).not.toBeNull();
      expect(
        normalizeDocumentEquation({
          version: 1,
          display: 'inline',
          children: [
            {
              type: 'equationArray',
              baseAlignment: 'center',
              maximumDistribution: false,
              objectDistribution: false,
              rowSpacingRule: 'exact',
              rowSpacing: 65_536,
              rows: [[]],
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
              type: 'equationArray',
              baseAlignment: 'center',
              maximumDistribution: false,
              objectDistribution: false,
              rowSpacingRule: 'single',
              rowSpacing: 0,
              rows: [[{ type: 'run', text: '&'.repeat(4_097) }]],
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
              type: 'borderBox',
              hideTop: 'false',
              hideBottom: false,
              hideLeft: false,
              hideRight: false,
              strikeHorizontal: false,
              strikeVertical: false,
              strikeBottomLeftToTopRight: false,
              strikeTopLeftToBottomRight: false,
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
              type: 'box',
              operatorEmulator: true,
              noBreak: true,
              differential: false,
              alignment: true,
              manualBreak: { alignmentAt: 0 },
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
      expect(editor.getHTML()).toContain('<menclose');
      expect(editor.getHTML()).toContain('<mpadded');
      expect(editor.getHTML()).toContain('<mphantom');
      expect(editor.getHTML()).toContain('width="0in"');
      expect(editor.getHTML()).toContain('height="0in"');
      expect(editor.getHTML()).toContain('depth="0in"');
      expect(editor.getHTML()).toContain('<mmultiscripts');
      expect(editor.getHTML()).toContain('<mprescripts');
      expect(editor.getHTML()).toContain('<none');
      expect(editor.getHTML()).toContain('mathvariant="bold-fraktur"');
      expect(editor.getHTML()).toContain('<mtable');
      expect(editor.getHTML()).toContain('rowspacing="1.5em"');
      expect(editor.getHTML()).toContain('<maligngroup');
      expect(editor.getHTML()).toContain('<malignmark');

      const sanitized = new DOMParser().parseFromString(
        sanitizeDocumentPageChromeHtml(
          `${equationHtml(complexEquation('inline'))}<span data-document-equation="true" data-equation-model="{">Broken</span><math xmlns="http://www.w3.org/1998/Math/MathML"><mtext>Loose</mtext></math>`,
        ),
        'text/html',
      );
      expect(
        sanitized.body.querySelectorAll('span[data-document-equation]'),
      ).toHaveLength(1);
      expect(sanitized.body.querySelectorAll('math')).toHaveLength(1);
      expect(sanitized.body.querySelector('mphantom')).not.toBeNull();
      expect(sanitized.body.querySelector('mphantom mpadded')).toHaveAttribute(
        'width',
        '0in',
      );
      expect(sanitized.body.querySelector('mphantom mpadded')).toHaveAttribute(
        'depth',
        '0in',
      );
      expect(
        sanitized.body.querySelector('mpadded[height="0in"]'),
      ).not.toBeNull();
      expect(sanitized.body.querySelectorAll('mmultiscripts')).toHaveLength(3);
      expect(sanitized.body.querySelectorAll('mprescripts')).toHaveLength(3);
      expect(
        sanitized.body.querySelectorAll('mmultiscripts none'),
      ).toHaveLength(2);
      expect(
        sanitized.body.querySelector('mtext[mathvariant="bold-fraktur"]')
          ?.textContent,
      ).toBe('styledF');
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
    ).toEqual([34, 34, 1, 1]);
    expect(bodyEquations.map(documentEquationFromElement).every(Boolean)).toBe(
      true,
    );
    expect(documentEquationFromElement(bodyEquations[0])).toEqual(
      complexEquation('inline'),
    );
    expect(documentEquationFromElement(bodyEquations[1])).toEqual(
      complexEquation('block'),
    );
    expect(bodyEquations[1]).toHaveClass('block', 'justification-right');
    expect(documentEquationFromElement(bodyEquations[2])).toEqual(
      borderBoxEquation('F', 'top'),
    );
    expect(documentEquationFromElement(bodyEquations[3])).toEqual(
      boxEquation('N', 'bottom'),
    );
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
    const importedBody = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    const importedEquations = Array.from(
      importedBody.body.querySelectorAll<HTMLElement>(
        '[data-document-equation]',
      ),
    );
    expect(importedEquations).toHaveLength(4);
    expect(documentEquationFromElement(importedEquations[1])).toEqual(
      complexEquation('block'),
    );
    expect(importedEquations[1]).toHaveClass('block', 'justification-right');
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
      `<m:borderBox><m:e>${run}</m:e></m:borderBox>`,
      `<m:borderBox><m:borderBoxPr><m:hideTop/><m:hideBot m:val="0"/><m:hideLeft m:val="false"/><m:hideRight m:val="true"/><m:strikeH/><m:strikeV m:val="off"/><m:strikeBLTR m:val="on"/><m:strikeTLBR m:val="0"/><m:ctrlPr/></m:borderBoxPr><m:e>${run}</m:e></m:borderBox>`,
      `<m:box><m:e>${run}</m:e></m:box>`,
      `<m:box><m:boxPr><m:opEmu/><m:noBreak m:val="0"/><m:diff/><m:brk m:alnAt=" +003 "/><m:aln/><m:ctrlPr/></m:boxPr><m:e>${run}</m:e></m:box>`,
      `<m:box><m:boxPr><m:brk/></m:boxPr><m:e>${run}</m:e></m:box>`,
      `<m:f><m:fPr><m:type m:val="noBar"/><m:ctrlPr/></m:fPr><m:num>${run}</m:num><m:den>${run}</m:den></m:f>`,
      `<m:rad><m:radPr><m:degHide m:val="1"/><m:ctrlPr/></m:radPr><m:e>${run}</m:e></m:rad>`,
      `<m:nary><m:naryPr><m:chr m:val="&#x2211;"/><m:limLoc m:val="undOvr"/><m:subHide m:val="1"/><m:supHide m:val="true"/><m:ctrlPr/></m:naryPr><m:sub/><m:sup/><m:e>${run}</m:e></m:nary>`,
      `<m:m><m:mr><m:e>${run}</m:e></m:mr></m:m>`,
      `<m:m><m:mPr><m:baseJc m:val="bot"/><m:plcHide m:val="true"/><m:mcs><m:mc><m:mcPr><m:count m:val="2"/><m:mcJc m:val="left"/></m:mcPr></m:mc><m:mc><m:mcPr><m:count/><m:mcJc/></m:mcPr></m:mc></m:mcs><m:ctrlPr/></m:mPr><m:mr><m:e>${run}</m:e><m:e>${run}</m:e><m:e/></m:mr></m:m>`,
      `<m:eqArr><m:e>${run}</m:e></m:eqArr>`,
      `<m:eqArr><m:eqArrPr><m:baseJc m:val="bot"/><m:maxDist/><m:objDist m:val="0"/><m:rSpRule m:val="4"/><m:rSp m:val=" +0003 "/><m:ctrlPr/></m:eqArrPr><m:e><m:r><m:t>&amp;x+&amp;&amp;y</m:t></m:r></m:e><m:e/></m:eqArr>`,
      `<m:eqArr><m:eqArrPr><m:baseJc/><m:rSpRule/><m:rSp/></m:eqArrPr><m:e>${run}</m:e></m:eqArr>`,
      `<m:limLow><m:e>${run}</m:e><m:lim>${run}</m:lim></m:limLow>`,
      `<m:limLow><m:limLowPr><m:ctrlPr/></m:limLowPr><m:e>${run}</m:e><m:lim>${run}</m:lim></m:limLow>`,
      `<m:limUpp><m:e>${run}</m:e><m:lim>${run}</m:lim></m:limUpp>`,
      `<m:limUpp><m:limUppPr/><m:e>${run}</m:e><m:lim>${run}</m:lim></m:limUpp>`,
      `<m:groupChr><m:e>${run}</m:e></m:groupChr>`,
      `<m:groupChr><m:groupChrPr/><m:e>${run}</m:e></m:groupChr>`,
      `<m:groupChr><m:groupChrPr><m:chr/><m:pos/><m:vertJc/><m:ctrlPr/></m:groupChrPr><m:e>${run}</m:e></m:groupChr>`,
      `<m:groupChr><m:groupChrPr><m:chr m:val="&#x23DE;"/><m:pos m:val="top"/><m:vertJc m:val="bot"/><m:ctrlPr/></m:groupChrPr><m:e>${run}</m:e></m:groupChr>`,
      `<m:phant><m:e>${run}</m:e></m:phant>`,
      `<m:phant><m:phantPr/><m:e>${run}</m:e></m:phant>`,
      `<m:phant><m:phantPr><m:show/><m:zeroWid/><m:zeroAsc/><m:zeroDesc/><m:transp/><m:ctrlPr/></m:phantPr><m:e>${run}</m:e></m:phant>`,
      `<m:phant><m:phantPr><m:show m:val="false"/><m:zeroWid m:val="0"/><m:zeroAsc m:val="on"/><m:zeroDesc m:val="off"/><m:transp m:val="true"/><m:ctrlPr/></m:phantPr><m:e>${run}</m:e></m:phant>`,
      `<m:sPre><m:sub>${run}</m:sub><m:sup>${run}</m:sup><m:e>${run}</m:e></m:sPre>`,
      `<m:sPre><m:sPrePr/><m:sub/><m:sup>${run}</m:sup><m:e>${run}</m:e></m:sPre>`,
      `<m:sPre><m:sPrePr><m:ctrlPr/></m:sPrePr><m:sub>${run}</m:sub><m:sup/><m:e>${run}</m:e></m:sPre>`,
      `<m:sPre><m:sub/><m:sup/><m:e>${run}</m:e></m:sPre>`,
      '<m:r><m:rPr/><m:t>x</m:t></m:r>',
      '<m:r><m:rPr><m:lit/><m:nor m:val="0"/><m:scr m:val="fraktur"/><m:sty m:val="b"/><m:brk m:alnAt=" +003 "/><m:aln/></m:rPr><m:t>x</m:t></m:r>',
      '<m:r><m:rPr><m:lit m:val="false"/><m:nor m:val="on"/><m:scr/><m:sty/><m:brk/><m:aln m:val="off"/></m:rPr><m:t>x</m:t></m:r>',
      '<m:r><m:rPr><m:scr m:val="double-struck"/><m:sty m:val="p"/></m:rPr><m:t>x</m:t></m:r>',
      '<m:r><m:rPr><m:scr m:val="monospace"/><m:sty m:val="bi"/></m:rPr><m:t>x</m:t></m:r>',
      '<m:r><m:rPr><m:scr m:val="sans-serif"/><m:sty m:val="i"/></m:rPr><m:t>x</m:t></m:r>',
      '<m:r><m:rPr><m:scr m:val="script"/><m:sty m:val="bi"/></m:rPr><m:t>x</m:t></m:r>',
      '<m:r><m:rPr><m:scr m:val="roman"/><m:sty m:val="p"/></m:rPr><m:t>x</m:t></m:r>',
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
    expect(
      supported
        .slice(19, 23)
        .map((source) => inspectEquationModel(source)?.children[0]),
    ).toEqual([
      {
        type: 'lowerLimit',
        base: [{ type: 'run', text: 'x' }],
        limit: [{ type: 'run', text: 'x' }],
      },
      {
        type: 'lowerLimit',
        base: [{ type: 'run', text: 'x' }],
        limit: [{ type: 'run', text: 'x' }],
      },
      {
        type: 'upperLimit',
        base: [{ type: 'run', text: 'x' }],
        limit: [{ type: 'run', text: 'x' }],
      },
      {
        type: 'upperLimit',
        base: [{ type: 'run', text: 'x' }],
        limit: [{ type: 'run', text: 'x' }],
      },
    ]);
    expect(
      supported
        .slice(23, 27)
        .map((source) => inspectEquationModel(source)?.children[0]),
    ).toEqual([
      {
        type: 'groupCharacter',
        character: '\u23df',
        position: 'bottom',
        verticalJustification: 'top',
        children: [{ type: 'run', text: 'x' }],
      },
      {
        type: 'groupCharacter',
        character: '\u23df',
        position: 'bottom',
        verticalJustification: 'top',
        children: [{ type: 'run', text: 'x' }],
      },
      {
        type: 'groupCharacter',
        character: '',
        position: 'bottom',
        verticalJustification: 'bottom',
        children: [{ type: 'run', text: 'x' }],
      },
      {
        type: 'groupCharacter',
        character: '\u23de',
        position: 'top',
        verticalJustification: 'bottom',
        children: [{ type: 'run', text: 'x' }],
      },
    ]);
    expect(
      supported
        .slice(27, 31)
        .map((source) => inspectEquationModel(source)?.children[0]),
    ).toEqual([
      {
        type: 'phantom',
        show: true,
        zeroWidth: false,
        zeroAscent: false,
        zeroDescent: false,
        transparent: false,
        children: [{ type: 'run', text: 'x' }],
      },
      {
        type: 'phantom',
        show: true,
        zeroWidth: false,
        zeroAscent: false,
        zeroDescent: false,
        transparent: false,
        children: [{ type: 'run', text: 'x' }],
      },
      {
        type: 'phantom',
        show: true,
        zeroWidth: true,
        zeroAscent: true,
        zeroDescent: true,
        transparent: true,
        children: [{ type: 'run', text: 'x' }],
      },
      {
        type: 'phantom',
        show: false,
        zeroWidth: false,
        zeroAscent: true,
        zeroDescent: false,
        transparent: true,
        children: [{ type: 'run', text: 'x' }],
      },
    ]);
    expect(
      supported
        .slice(31, 35)
        .map((source) => inspectEquationModel(source)?.children[0]),
    ).toEqual([
      {
        type: 'preSubSuperScript',
        base: [{ type: 'run', text: 'x' }],
        subScript: [{ type: 'run', text: 'x' }],
        superScript: [{ type: 'run', text: 'x' }],
      },
      {
        type: 'preSubSuperScript',
        base: [{ type: 'run', text: 'x' }],
        subScript: [],
        superScript: [{ type: 'run', text: 'x' }],
      },
      {
        type: 'preSubSuperScript',
        base: [{ type: 'run', text: 'x' }],
        subScript: [{ type: 'run', text: 'x' }],
        superScript: [],
      },
      {
        type: 'preSubSuperScript',
        base: [{ type: 'run', text: 'x' }],
        subScript: [],
        superScript: [],
      },
    ]);
    expect(
      supported
        .slice(35, 43)
        .map((source) => inspectEquationModel(source)?.children[0]),
    ).toEqual([
      { type: 'run', text: 'x' },
      {
        type: 'run',
        text: 'x',
        literal: true,
        script: 'fraktur',
        style: 'bold',
        manualBreak: { alignmentAt: 3 },
        alignment: true,
      },
      {
        type: 'run',
        text: 'x',
        normalText: true,
        manualBreak: {},
      },
      {
        type: 'run',
        text: 'x',
        script: 'doubleStruck',
        style: 'plain',
      },
      {
        type: 'run',
        text: 'x',
        script: 'monospace',
        style: 'boldItalic',
      },
      { type: 'run', text: 'x', script: 'sansSerif' },
      {
        type: 'run',
        text: 'x',
        script: 'script',
        style: 'boldItalic',
      },
      { type: 'run', text: 'x', style: 'plain' },
    ]);
    expect(
      supported
        .slice(6, 11)
        .map((source) => inspectEquationModel(source)?.children[0]),
    ).toEqual([
      {
        type: 'borderBox',
        hideTop: false,
        hideBottom: false,
        hideLeft: false,
        hideRight: false,
        strikeHorizontal: false,
        strikeVertical: false,
        strikeBottomLeftToTopRight: false,
        strikeTopLeftToBottomRight: false,
        children: [{ type: 'run', text: 'x' }],
      },
      {
        type: 'borderBox',
        hideTop: true,
        hideBottom: false,
        hideLeft: false,
        hideRight: true,
        strikeHorizontal: true,
        strikeVertical: false,
        strikeBottomLeftToTopRight: true,
        strikeTopLeftToBottomRight: false,
        children: [{ type: 'run', text: 'x' }],
      },
      {
        type: 'box',
        operatorEmulator: false,
        noBreak: false,
        differential: false,
        alignment: false,
        children: [{ type: 'run', text: 'x' }],
      },
      {
        type: 'box',
        operatorEmulator: true,
        noBreak: false,
        differential: true,
        alignment: true,
        manualBreak: { alignmentAt: 3 },
        children: [{ type: 'run', text: 'x' }],
      },
      {
        type: 'box',
        operatorEmulator: false,
        noBreak: false,
        differential: false,
        alignment: false,
        manualBreak: {},
        children: [{ type: 'run', text: 'x' }],
      },
    ]);
    expect(
      supported
        .slice(16, 19)
        .map((source) => inspectEquationModel(source)?.children[0]),
    ).toEqual([
      {
        type: 'equationArray',
        baseAlignment: 'center',
        maximumDistribution: false,
        objectDistribution: false,
        rowSpacingRule: 'single',
        rowSpacing: 0,
        rows: [[{ type: 'run', text: 'x' }]],
      },
      {
        type: 'equationArray',
        baseAlignment: 'bottom',
        maximumDistribution: true,
        objectDistribution: false,
        rowSpacingRule: 'multiple',
        rowSpacing: 3,
        rows: [[{ type: 'run', text: '&x+&&y' }], []],
      },
      {
        type: 'equationArray',
        baseAlignment: 'center',
        maximumDistribution: false,
        objectDistribution: false,
        rowSpacingRule: 'single',
        rowSpacing: 0,
        rows: [[{ type: 'run', text: 'x' }]],
      },
    ]);
    expect(
      inspectEquationBody(
        `<m:eqArr>${Array.from({ length: 64 }, () => `<m:e>${run}</m:e>`).join('')}</m:eqArr>`,
      ),
    ).toBe('supported');
    expect(
      inspectEquationBody(
        `<m:eqArr><m:eqArrPr><m:rSpRule m:val="3"/><m:rSp m:val="65535"/></m:eqArrPr><m:e>${run}</m:e></m:eqArr>`,
      ),
    ).toBe('supported');

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
      `<m:borderBox><m:borderBoxPr><m:grow/></m:borderBoxPr><m:e>${run}</m:e></m:borderBox>`,
      `<m:borderBox><m:borderBoxPr><m:strikeH/><m:hideTop/></m:borderBoxPr><m:e>${run}</m:e></m:borderBox>`,
      `<m:borderBox><m:borderBoxPr><m:hideTop/><m:hideTop/></m:borderBoxPr><m:e>${run}</m:e></m:borderBox>`,
      `<m:borderBox><m:borderBoxPr><m:hideTop m:val="maybe"/></m:borderBoxPr><m:e>${run}</m:e></m:borderBox>`,
      `<m:borderBox><v:borderBoxPr xmlns:v="${VENDOR_NAMESPACE}"><m:hideTop/></v:borderBoxPr><m:e>${run}</m:e></m:borderBox>`,
      `<m:borderBox><m:borderBoxPr><m:hideTop xmlns:r="${RELATIONSHIP_NAMESPACE}" r:id="rIdUnsafe"/></m:borderBoxPr><m:e>${run}</m:e></m:borderBox>`,
      `<m:borderBox><m:e>${run}</m:e><m:borderBoxPr><m:hideTop/></m:borderBoxPr></m:borderBox>`,
      `<m:borderBox><m:borderBoxPr/><m:e/></m:borderBox>`,
      `<m:borderBox><m:e>${run}</m:e><m:e>${run}</m:e></m:borderBox>`,
      `<m:box><m:boxPr><m:grow/></m:boxPr><m:e>${run}</m:e></m:box>`,
      `<m:box><m:boxPr><m:aln/><m:opEmu/></m:boxPr><m:e>${run}</m:e></m:box>`,
      `<m:box><m:boxPr><m:opEmu/><m:opEmu/></m:boxPr><m:e>${run}</m:e></m:box>`,
      `<m:box><m:boxPr><m:brk m:alnAt="0"/></m:boxPr><m:e>${run}</m:e></m:box>`,
      `<m:box><m:boxPr><m:brk m:alnAt="256"/></m:boxPr><m:e>${run}</m:e></m:box>`,
      `<m:box><m:boxPr><m:brk m:alnAt="3.5"/></m:boxPr><m:e>${run}</m:e></m:box>`,
      `<m:box><m:boxPr><m:brk m:alnAt="3" m:extra="semantic"/></m:boxPr><m:e>${run}</m:e></m:box>`,
      `<m:box><m:boxPr><m:brk xmlns:r="${RELATIONSHIP_NAMESPACE}" r:id="rIdUnsafe"/></m:boxPr><m:e>${run}</m:e></m:box>`,
      `<m:box><m:boxPr><m:brk>${run}</m:brk></m:boxPr><m:e>${run}</m:e></m:box>`,
      `<m:box><m:e>${run}</m:e><m:boxPr><m:opEmu/></m:boxPr></m:box>`,
      `<m:box><m:boxPr/><m:e/></m:box>`,
      `<m:box><m:e>${run}</m:e><m:e>${run}</m:e></m:box>`,
      `<m:f><m:fPr><m:m><m:mr><m:e>${run}</m:e></m:mr></m:m></m:fPr><m:num>${run}</m:num><m:den>${run}</m:den></m:f>`,
      `<m:f><m:fPr><m:type m:val="bar" m:extra="semantic"/></m:fPr><m:num>${run}</m:num><m:den>${run}</m:den></m:f>`,
      `<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg>${run}</m:deg><m:e>${run}</m:e></m:rad>`,
      `<m:rad><m:radPr><m:ctrlPr/><m:degHide/></m:radPr><m:e>${run}</m:e></m:rad>`,
      `<m:nary><m:naryPr><m:chr m:val="&#x2211;"/><m:subHide m:val="on"/></m:naryPr><m:sub>${run}</m:sub><m:e>${run}</m:e></m:nary>`,
      `<m:nary><m:naryPr><m:chr m:val="&#x2211;"/></m:naryPr><m:e>${run}</m:e></m:nary>`,
      `<m:d><m:dPr><m:grow m:val="1"/></m:dPr><m:e>${run}</m:e></m:d>`,
      `<m:m><m:mPr><m:cSp m:val="120"/></m:mPr><m:mr><m:e>${run}</m:e></m:mr></m:m>`,
      `<m:m><m:mr><m:e>${run}</m:e><m:e>${run}</m:e></m:mr><m:mr><m:e>${run}</m:e></m:mr></m:m>`,
      `<m:m><m:mPr><m:mcs><m:mc><m:mcPr><m:count m:val="1"/><m:mcJc m:val="left"/></m:mcPr></m:mc></m:mcs></m:mPr><m:mr><m:e>${run}</m:e><m:e>${run}</m:e></m:mr></m:m>`,
      `<m:m><m:mPr><m:mcs><m:mc><m:mcPr><m:count m:val="0"/></m:mcPr></m:mc></m:mcs></m:mPr><m:mr><m:e>${run}</m:e></m:mr></m:m>`,
      `<m:m>${Array.from({ length: 65 }, () => `<m:mr><m:e>${run}</m:e></m:mr>`).join('')}</m:m>`,
      `<m:eqArr><m:eqArrPr><m:grow/></m:eqArrPr><m:e>${run}</m:e></m:eqArr>`,
      `<m:eqArr><m:eqArrPr><m:objDist/><m:maxDist/></m:eqArrPr><m:e>${run}</m:e></m:eqArr>`,
      `<m:eqArr><m:eqArrPr><m:maxDist/><m:maxDist/></m:eqArrPr><m:e>${run}</m:e></m:eqArr>`,
      `<m:eqArr><m:eqArrPr><m:baseJc m:val="bottom"/></m:eqArrPr><m:e>${run}</m:e></m:eqArr>`,
      `<m:eqArr><m:eqArrPr><m:rSpRule m:val="5"/></m:eqArrPr><m:e>${run}</m:e></m:eqArr>`,
      `<m:eqArr><m:eqArrPr><m:rSpRule m:val="-1"/></m:eqArrPr><m:e>${run}</m:e></m:eqArr>`,
      `<m:eqArr><m:eqArrPr><m:rSpRule m:val="1.5"/></m:eqArrPr><m:e>${run}</m:e></m:eqArr>`,
      `<m:eqArr><m:eqArrPr><m:rSp m:val="-1"/></m:eqArrPr><m:e>${run}</m:e></m:eqArr>`,
      `<m:eqArr><m:eqArrPr><m:rSp m:val="65536"/></m:eqArrPr><m:e>${run}</m:e></m:eqArr>`,
      `<m:eqArr><m:eqArrPr><m:rSp m:val="3.5"/></m:eqArrPr><m:e>${run}</m:e></m:eqArr>`,
      `<m:eqArr><m:eqArrPr><m:rSp m:val="3" m:extra="semantic"/></m:eqArrPr><m:e>${run}</m:e></m:eqArr>`,
      `<m:eqArr><m:eqArrPr><m:rSp xmlns:r="${RELATIONSHIP_NAMESPACE}" m:val="3" r:id="rIdUnsafe"/></m:eqArrPr><m:e>${run}</m:e></m:eqArr>`,
      `<m:eqArr><v:eqArrPr xmlns:v="${VENDOR_NAMESPACE}"><m:rSp m:val="3"/></v:eqArrPr><m:e>${run}</m:e></m:eqArr>`,
      `<m:eqArr><m:e>${run}</m:e><m:eqArrPr><m:rSp m:val="3"/></m:eqArrPr></m:eqArr>`,
      '<m:eqArr/>',
      `<m:eqArr>${Array.from({ length: 65 }, () => `<m:e>${run}</m:e>`).join('')}</m:eqArr>`,
      `<m:eqArr><m:e><m:r><m:t>${'&amp;'.repeat(4_097)}</m:t></m:r></m:e></m:eqArr>`,
      `<m:eqArr xmlns:r="${RELATIONSHIP_NAMESPACE}"><m:e r:id="rIdUnsafe">${run}</m:e></m:eqArr>`,
      `<m:eqArr><m:eqArrPr/><m:eqArrPr/><m:e>${run}</m:e></m:eqArr>`,
      `<m:limLow><m:limLowPr><m:grow/></m:limLowPr><m:e>${run}</m:e><m:lim>${run}</m:lim></m:limLow>`,
      `<m:limLow><m:e>${run}</m:e><m:limLowPr/><m:lim>${run}</m:lim></m:limLow>`,
      `<m:limLow><m:lim>${run}</m:lim><m:e>${run}</m:e></m:limLow>`,
      `<m:limLow><m:limLowPr/><m:limLowPr/><m:e>${run}</m:e><m:lim>${run}</m:lim></m:limLow>`,
      `<m:limLow><m:limLowPr><m:ctrlPr/><m:ctrlPr/></m:limLowPr><m:e>${run}</m:e><m:lim>${run}</m:lim></m:limLow>`,
      `<m:limLow><v:limLowPr xmlns:v="${VENDOR_NAMESPACE}"/><m:e>${run}</m:e><m:lim>${run}</m:lim></m:limLow>`,
      `<m:limLow><m:limLowPr><m:ctrlPr xmlns:r="${RELATIONSHIP_NAMESPACE}" r:id="rIdUnsafe"/></m:limLowPr><m:e>${run}</m:e><m:lim>${run}</m:lim></m:limLow>`,
      `<m:limLow><m:lim>${run}</m:lim></m:limLow>`,
      `<m:limLow><m:e>${run}</m:e></m:limLow>`,
      `<m:limLow><m:e/><m:lim>${run}</m:lim></m:limLow>`,
      `<m:limLow><m:e>${run}</m:e><m:lim/></m:limLow>`,
      `<m:limLow><m:e>${run}</m:e><m:e>${run}</m:e><m:lim>${run}</m:lim></m:limLow>`,
      `<m:limLow><m:e>${run}</m:e><m:lim>${run}</m:lim><m:lim>${run}</m:lim></m:limLow>`,
      `<m:limLow m:extra="semantic"><m:e>${run}</m:e><m:lim>${run}</m:lim></m:limLow>`,
      `<m:limUpp><m:limLowPr/><m:e>${run}</m:e><m:lim>${run}</m:lim></m:limUpp>`,
      `<m:limUpp><m:lim>${run}</m:lim><m:e>${run}</m:e></m:limUpp>`,
      `<m:limUpp><v:limUppPr xmlns:v="${VENDOR_NAMESPACE}"/><m:e>${run}</m:e><m:lim>${run}</m:lim></m:limUpp>`,
      `<m:limUpp xmlns:r="${RELATIONSHIP_NAMESPACE}"><m:e>${run}</m:e><m:lim r:id="rIdUnsafe">${run}</m:lim></m:limUpp>`,
      `<m:groupChr><m:groupChrPr><m:grow/></m:groupChrPr><m:e>${run}</m:e></m:groupChr>`,
      `<m:groupChr><m:e>${run}</m:e><m:groupChrPr/></m:groupChr>`,
      `<m:groupChr><m:groupChrPr><m:pos/><m:chr/></m:groupChrPr><m:e>${run}</m:e></m:groupChr>`,
      `<m:groupChr><m:groupChrPr><m:ctrlPr/><m:vertJc/></m:groupChrPr><m:e>${run}</m:e></m:groupChr>`,
      `<m:groupChr><m:groupChrPr/><m:groupChrPr/><m:e>${run}</m:e></m:groupChr>`,
      `<m:groupChr><m:groupChrPr><m:chr/><m:chr/></m:groupChrPr><m:e>${run}</m:e></m:groupChr>`,
      `<m:groupChr><m:groupChrPr><m:pos/><m:pos/></m:groupChrPr><m:e>${run}</m:e></m:groupChr>`,
      `<m:groupChr><m:groupChrPr><m:vertJc/><m:vertJc/></m:groupChrPr><m:e>${run}</m:e></m:groupChr>`,
      `<m:groupChr><m:e>${run}</m:e><m:e>${run}</m:e></m:groupChr>`,
      `<m:groupChr><m:groupChrPr><m:chr m:val="xy"/></m:groupChrPr><m:e>${run}</m:e></m:groupChr>`,
      `<m:groupChr><m:groupChrPr><m:pos m:val="bottom"/></m:groupChrPr><m:e>${run}</m:e></m:groupChr>`,
      `<m:groupChr><m:groupChrPr><m:vertJc m:val="center"/></m:groupChrPr><m:e>${run}</m:e></m:groupChr>`,
      `<m:groupChr><m:groupChrPr><m:chr m:val="&#x23DE;" m:extra="semantic"/></m:groupChrPr><m:e>${run}</m:e></m:groupChr>`,
      `<m:groupChr><m:groupChrPr><m:pos xmlns:r="${RELATIONSHIP_NAMESPACE}" m:val="top" r:id="rIdUnsafe"/></m:groupChrPr><m:e>${run}</m:e></m:groupChr>`,
      `<m:groupChr><v:groupChrPr xmlns:v="${VENDOR_NAMESPACE}"/><m:e>${run}</m:e></m:groupChr>`,
      `<m:groupChr><m:groupChrPr><m:ctrlPr xmlns:r="${RELATIONSHIP_NAMESPACE}" r:id="rIdUnsafe"/></m:groupChrPr><m:e>${run}</m:e></m:groupChr>`,
      `<m:groupChr><m:groupChrPr/><m:e/></m:groupChr>`,
      `<m:groupChr><m:groupChrPr/></m:groupChr>`,
      `<m:groupChr m:extra="semantic"><m:e>${run}</m:e></m:groupChr>`,
      `<m:groupChr xmlns:r="${RELATIONSHIP_NAMESPACE}"><m:e r:id="rIdUnsafe">${run}</m:e></m:groupChr>`,
      `<m:phant><m:phantPr><m:grow/></m:phantPr><m:e>${run}</m:e></m:phant>`,
      `<m:phant><m:e>${run}</m:e><m:phantPr/></m:phant>`,
      `<m:phant><m:phantPr><m:transp/><m:zeroDesc/></m:phantPr><m:e>${run}</m:e></m:phant>`,
      `<m:phant><m:phantPr><m:ctrlPr/><m:transp/></m:phantPr><m:e>${run}</m:e></m:phant>`,
      `<m:phant><m:phantPr/><m:phantPr/><m:e>${run}</m:e></m:phant>`,
      `<m:phant><m:phantPr><m:show/><m:show/></m:phantPr><m:e>${run}</m:e></m:phant>`,
      `<m:phant><m:phantPr><m:zeroWid/><m:zeroWid/></m:phantPr><m:e>${run}</m:e></m:phant>`,
      `<m:phant><m:phantPr><m:zeroAsc/><m:zeroAsc/></m:phantPr><m:e>${run}</m:e></m:phant>`,
      `<m:phant><m:phantPr><m:zeroDesc/><m:zeroDesc/></m:phantPr><m:e>${run}</m:e></m:phant>`,
      `<m:phant><m:phantPr><m:transp/><m:transp/></m:phantPr><m:e>${run}</m:e></m:phant>`,
      `<m:phant><m:phantPr><m:ctrlPr/><m:ctrlPr/></m:phantPr><m:e>${run}</m:e></m:phant>`,
      `<m:phant><m:phantPr><m:show m:val="maybe"/></m:phantPr><m:e>${run}</m:e></m:phant>`,
      `<m:phant><m:phantPr><m:zeroAsc m:val="2"/></m:phantPr><m:e>${run}</m:e></m:phant>`,
      `<m:phant><m:phantPr><m:transp m:val="maybe"/></m:phantPr><m:e>${run}</m:e></m:phant>`,
      `<m:phant><m:phantPr><m:transp m:val="1" m:extra="semantic"/></m:phantPr><m:e>${run}</m:e></m:phant>`,
      `<m:phant><v:phantPr xmlns:v="${VENDOR_NAMESPACE}"><m:show/></v:phantPr><m:e>${run}</m:e></m:phant>`,
      `<m:phant><m:phantPr><m:zeroDesc xmlns:r="${RELATIONSHIP_NAMESPACE}" r:id="rIdUnsafe"/></m:phantPr><m:e>${run}</m:e></m:phant>`,
      `<m:phant><m:phantPr><m:ctrlPr xmlns:r="${RELATIONSHIP_NAMESPACE}" r:id="rIdUnsafe"/></m:phantPr><m:e>${run}</m:e></m:phant>`,
      `<m:phant><m:phantPr m:extra="semantic"/><m:e>${run}</m:e></m:phant>`,
      `<m:phant><m:phantPr><m:show>${run}</m:show></m:phantPr><m:e>${run}</m:e></m:phant>`,
      `<m:phant><m:phantPr/><m:e/></m:phant>`,
      `<m:phant><m:phantPr/></m:phant>`,
      `<m:phant><m:e>${run}</m:e><m:e>${run}</m:e></m:phant>`,
      `<m:phant m:extra="semantic"><m:e>${run}</m:e></m:phant>`,
      `<m:phant xmlns:r="${RELATIONSHIP_NAMESPACE}"><m:e r:id="rIdUnsafe">${run}</m:e></m:phant>`,
      `<m:sPre><m:sPrePr><m:alnScr/></m:sPrePr><m:sub>${run}</m:sub><m:sup>${run}</m:sup><m:e>${run}</m:e></m:sPre>`,
      `<m:sPre><m:e>${run}</m:e><m:sub>${run}</m:sub><m:sup>${run}</m:sup></m:sPre>`,
      `<m:sPre><m:sup>${run}</m:sup><m:sub>${run}</m:sub><m:e>${run}</m:e></m:sPre>`,
      `<m:sPre><m:sub>${run}</m:sub><m:sPrePr/><m:sup>${run}</m:sup><m:e>${run}</m:e></m:sPre>`,
      `<m:sPre><m:sPrePr/><m:sPrePr/><m:sub>${run}</m:sub><m:sup>${run}</m:sup><m:e>${run}</m:e></m:sPre>`,
      `<m:sPre><m:sPrePr><m:ctrlPr/><m:ctrlPr/></m:sPrePr><m:sub>${run}</m:sub><m:sup>${run}</m:sup><m:e>${run}</m:e></m:sPre>`,
      `<m:sPre><m:sub>${run}</m:sub><m:sub>${run}</m:sub><m:sup>${run}</m:sup><m:e>${run}</m:e></m:sPre>`,
      `<m:sPre><m:sub>${run}</m:sub><m:sup>${run}</m:sup><m:sup>${run}</m:sup><m:e>${run}</m:e></m:sPre>`,
      `<m:sPre><m:sub>${run}</m:sub><m:sup>${run}</m:sup><m:e>${run}</m:e><m:e>${run}</m:e></m:sPre>`,
      `<m:sPre><m:sup>${run}</m:sup><m:e>${run}</m:e></m:sPre>`,
      `<m:sPre><m:sub>${run}</m:sub><m:e>${run}</m:e></m:sPre>`,
      `<m:sPre><m:sub>${run}</m:sub><m:sup>${run}</m:sup></m:sPre>`,
      `<m:sPre><m:sub>${run}</m:sub><m:sup>${run}</m:sup><m:e/></m:sPre>`,
      `<m:sPre><v:sPrePr xmlns:v="${VENDOR_NAMESPACE}"/><m:sub>${run}</m:sub><m:sup>${run}</m:sup><m:e>${run}</m:e></m:sPre>`,
      `<m:sPre><v:sub xmlns:v="${VENDOR_NAMESPACE}">${run}</v:sub><m:sup>${run}</m:sup><m:e>${run}</m:e></m:sPre>`,
      `<m:sPre><m:sPrePr m:extra="semantic"/><m:sub>${run}</m:sub><m:sup>${run}</m:sup><m:e>${run}</m:e></m:sPre>`,
      `<m:sPre><m:sPrePr><m:ctrlPr xmlns:r="${RELATIONSHIP_NAMESPACE}" r:id="rIdUnsafe"/></m:sPrePr><m:sub>${run}</m:sub><m:sup>${run}</m:sup><m:e>${run}</m:e></m:sPre>`,
      `<m:sPre xmlns:r="${RELATIONSHIP_NAMESPACE}"><m:sub r:id="rIdUnsafe">${run}</m:sub><m:sup>${run}</m:sup><m:e>${run}</m:e></m:sPre>`,
      `<m:sPre m:extra="semantic"><m:sub>${run}</m:sub><m:sup>${run}</m:sup><m:e>${run}</m:e></m:sPre>`,
      `<m:sPre>meaningful<m:sub>${run}</m:sub><m:sup>${run}</m:sup><m:e>${run}</m:e></m:sPre>`,
      `<m:sPre><m:sub>${run}<m:argPr/></m:sub><m:sup>${run}</m:sup><m:e>${run}</m:e></m:sPre>`,
      '<m:r><m:rPr><m:ctrlPr/></m:rPr><m:t>x</m:t></m:r>',
      '<m:r><m:rPr><m:sty m:val="b"/><m:scr m:val="fraktur"/></m:rPr><m:t>x</m:t></m:r>',
      '<m:r><m:t>x</m:t><m:rPr/></m:r>',
      '<m:r><m:rPr/><m:rPr/><m:t>x</m:t></m:r>',
      '<m:r><m:t>x</m:t><m:t>y</m:t></m:r>',
      '<m:r><m:rPr/></m:r>',
      '<m:r><m:rPr/><m:t/></m:r>',
      '<m:r m:extra="semantic"><m:t>x</m:t></m:r>',
      '<m:r>meaningful<m:t>x</m:t></m:r>',
      '<m:r><m:rPr m:extra="semantic"/><m:t>x</m:t></m:r>',
      '<m:r><m:rPr>meaningful</m:rPr><m:t>x</m:t></m:r>',
      '<m:r><m:rPr><m:lit/><m:lit/></m:rPr><m:t>x</m:t></m:r>',
      '<m:r><m:rPr><m:nor/><m:nor/></m:rPr><m:t>x</m:t></m:r>',
      '<m:r><m:rPr><m:scr/><m:scr/></m:rPr><m:t>x</m:t></m:r>',
      '<m:r><m:rPr><m:sty/><m:sty/></m:rPr><m:t>x</m:t></m:r>',
      '<m:r><m:rPr><m:brk/><m:brk/></m:rPr><m:t>x</m:t></m:r>',
      '<m:r><m:rPr><m:aln/><m:aln/></m:rPr><m:t>x</m:t></m:r>',
      '<m:r><m:rPr><m:lit m:val="maybe"/></m:rPr><m:t>x</m:t></m:r>',
      '<m:r><m:rPr><m:nor m:val="2"/></m:rPr><m:t>x</m:t></m:r>',
      '<m:r><m:rPr><m:aln m:val=""/></m:rPr><m:t>x</m:t></m:r>',
      '<m:r><m:rPr><m:scr m:val="blackboard"/></m:rPr><m:t>x</m:t></m:r>',
      '<m:r><m:rPr><m:sty m:val="bold"/></m:rPr><m:t>x</m:t></m:r>',
      '<m:r><m:rPr><m:scr m:val="fraktur" m:extra="semantic"/></m:rPr><m:t>x</m:t></m:r>',
      `<m:r><m:rPr><m:sty xmlns:r="${RELATIONSHIP_NAMESPACE}" m:val="b" r:id="rIdUnsafe"/></m:rPr><m:t>x</m:t></m:r>`,
      '<m:r><m:rPr><m:brk m:alnAt="0"/></m:rPr><m:t>x</m:t></m:r>',
      '<m:r><m:rPr><m:brk m:alnAt="256"/></m:rPr><m:t>x</m:t></m:r>',
      '<m:r><m:rPr><m:brk m:alnAt="3.5"/></m:rPr><m:t>x</m:t></m:r>',
      '<m:r><m:rPr><m:brk m:alnAt="3" m:extra="semantic"/></m:rPr><m:t>x</m:t></m:r>',
      `<m:r><m:rPr><m:brk xmlns:r="${RELATIONSHIP_NAMESPACE}" r:id="rIdUnsafe"/></m:rPr><m:t>x</m:t></m:r>`,
      `<m:r><m:rPr><m:brk>${run}</m:brk></m:rPr><m:t>x</m:t></m:r>`,
      `<m:r><v:rPr xmlns:v="${VENDOR_NAMESPACE}"/><m:t>x</m:t></m:r>`,
      `<m:r><m:rPr><v:scr xmlns:v="${VENDOR_NAMESPACE}" v:val="fraktur"/></m:rPr><m:t>x</m:t></m:r>`,
      `<m:r><v:t xmlns:v="${VENDOR_NAMESPACE}">x</v:t></m:r>`,
      `<m:r><w:rPr xmlns:w="${WORD_NAMESPACE}"/><m:t>x</m:t></m:r>`,
      `<m:r><m:rPr xmlns:r="${RELATIONSHIP_NAMESPACE}" r:id="rIdUnsafe"/><m:t>x</m:t></m:r>`,
      `<m:r><m:t xmlns:r="${RELATIONSHIP_NAMESPACE}" r:id="rIdUnsafe">x</m:t></m:r>`,
      '<m:r><m:t xml:space="invalid">x</m:t></m:r>',
      `<m:rPr/>${run}`,
      deepOmml(34),
      `<m:r><m:t>${'x'.repeat(65_537)}</m:t></m:r>`,
      `<m:d>${Array.from({ length: 33 }, () => `<m:e>${run}</m:e>`).join('')}</m:d>`,
    ];
    expect(unsupported.map(inspectEquationBody)).toEqual(
      unsupported.map(() => 'unsupported'),
    );
  });

  test('preserves display-math paragraph justification and rejects malformed roots', () => {
    const run = '<m:r><m:t>x</m:t></m:r>';
    const paragraph = (properties = '', namespace = MATH_NAMESPACE) =>
      `<m:oMathPara xmlns:m="${namespace}">${properties}<m:oMath>${run}</m:oMath></m:oMathPara>`;
    const supported = [
      paragraph(),
      paragraph('<m:oMathParaPr/>'),
      paragraph('<m:oMathParaPr><m:jc/></m:oMathParaPr>'),
      paragraph('<m:oMathParaPr><m:jc m:val="left"/></m:oMathParaPr>'),
      paragraph('<m:oMathParaPr><m:jc m:val="right"/></m:oMathParaPr>'),
      paragraph('<m:oMathParaPr><m:jc m:val="center"/></m:oMathParaPr>'),
      paragraph('<m:oMathParaPr><m:jc m:val="centerGroup"/></m:oMathParaPr>'),
      paragraph(
        '<m:oMathParaPr><m:jc m:val="left"/></m:oMathParaPr>',
        STRICT_MATH_NAMESPACE,
      ),
    ];
    const inspections = supported.map(inspectEquationRoot);
    expect(inspections.map(({ status }) => status)).toEqual(
      supported.map(() => 'supported'),
    );
    expect(inspections.map(({ equation }) => equation?.justification)).toEqual([
      undefined,
      undefined,
      undefined,
      'left',
      'right',
      'center',
      undefined,
      'left',
    ]);

    const unsupported = [
      paragraph('<m:oMathParaPr><m:dispDef/></m:oMathParaPr>'),
      `<m:oMathPara xmlns:m="${MATH_NAMESPACE}"><m:oMath>${run}</m:oMath><m:oMathParaPr/></m:oMathPara>`,
      paragraph('<m:oMathParaPr/><m:oMathParaPr/>'),
      paragraph('<m:oMathParaPr><m:jc/><m:jc/></m:oMathParaPr>'),
      `<m:oMathPara xmlns:m="${MATH_NAMESPACE}"><m:oMath>${run}</m:oMath><m:oMath>${run}</m:oMath></m:oMathPara>`,
      `<m:oMathPara xmlns:m="${MATH_NAMESPACE}"><m:oMathParaPr/></m:oMathPara>`,
      paragraph('<m:oMathParaPr><m:jc m:val="justify"/></m:oMathParaPr>'),
      paragraph(
        '<m:oMathParaPr><m:jc m:val="left" m:extra="semantic"/></m:oMathParaPr>',
      ),
      paragraph(
        `<m:oMathParaPr xmlns:r="${RELATIONSHIP_NAMESPACE}"><m:jc m:val="left" r:id="rIdUnsafe"/></m:oMathParaPr>`,
      ),
      paragraph(
        `<v:oMathParaPr xmlns:v="${VENDOR_NAMESPACE}"><m:jc m:val="left"/></v:oMathParaPr>`,
      ),
      paragraph(
        `<m:oMathParaPr><v:jc xmlns:v="${VENDOR_NAMESPACE}" v:val="left"/></m:oMathParaPr>`,
      ),
      paragraph('<m:oMathParaPr>meaningful</m:oMathParaPr>'),
      paragraph('<m:oMathParaPr m:extra="semantic"/>'),
      paragraph('<m:oMathParaPr><m:jc>meaningful</m:jc></m:oMathParaPr>'),
      `<m:oMathPara xmlns:m="${MATH_NAMESPACE}" m:extra="semantic"><m:oMath>${run}</m:oMath></m:oMathPara>`,
      `<m:oMathPara xmlns:m="${MATH_NAMESPACE}">meaningful<m:oMath>${run}</m:oMath></m:oMathPara>`,
    ];
    expect(
      unsupported.map(inspectEquationRoot).map(({ status }) => status),
    ).toEqual(unsupported.map(() => 'unsupported'));
    expect(
      inspectEquationRoot(
        `<v:oMathPara xmlns:v="${VENDOR_NAMESPACE}" xmlns:m="${MATH_NAMESPACE}"><m:oMath>${run}</m:oMath></v:oMathPara>`,
      ).status,
    ).toBe('spoofed');
  });

  test('preserves aligned right scripts and strictly validates script properties', () => {
    const run = '<m:r><m:t>x</m:t></m:r>';
    const supported = [
      `<m:sSup><m:sSupPr><m:ctrlPr/></m:sSupPr><m:e>${run}</m:e><m:sup>${run}</m:sup></m:sSup>`,
      `<m:sSub><m:sSubPr><m:ctrlPr/></m:sSubPr><m:e>${run}</m:e><m:sub>${run}</m:sub></m:sSub>`,
      `<m:sSubSup><m:e>${run}</m:e><m:sub>${run}</m:sub><m:sup>${run}</m:sup></m:sSubSup>`,
      `<m:sSubSup><m:sSubSupPr/><m:e>${run}</m:e><m:sub>${run}</m:sub><m:sup>${run}</m:sup></m:sSubSup>`,
      `<m:sSubSup><m:sSubSupPr><m:alnScr/></m:sSubSupPr><m:e>${run}</m:e><m:sub>${run}</m:sub><m:sup>${run}</m:sup></m:sSubSup>`,
      `<m:sSubSup><m:sSubSupPr><m:alnScr m:val="0"/><m:ctrlPr/></m:sSubSupPr><m:e>${run}</m:e><m:sub>${run}</m:sub><m:sup>${run}</m:sup></m:sSubSup>`,
      `<m:sSubSup><m:sSubSupPr><m:alnScr m:val="on"/></m:sSubSupPr><m:e>${run}</m:e><m:sub>${run}</m:sub><m:sup>${run}</m:sup></m:sSubSup>`,
      `<m:func><m:funcPr><m:ctrlPr/></m:funcPr><m:fName>${run}</m:fName><m:e>${run}</m:e></m:func>`,
    ];
    expect(supported.map(inspectEquationBody)).toEqual(
      supported.map(() => 'supported'),
    );
    expect(
      supported.slice(2, 7).map((source) => {
        const expression = inspectEquationModel(source)?.children[0];
        return expression?.type === 'subSuperScript'
          ? expression.alignScripts
          : false;
      }),
    ).toEqual([undefined, undefined, true, undefined, true]);

    const unsupported = [
      `<m:sSup><m:e>${run}</m:e><m:sSupPr/><m:sup>${run}</m:sup></m:sSup>`,
      `<m:sSup><m:sSupPr/><m:sSupPr/><m:e>${run}</m:e><m:sup>${run}</m:sup></m:sSup>`,
      `<m:sSup><m:sSupPr><m:ctrlPr/><m:ctrlPr/></m:sSupPr><m:e>${run}</m:e><m:sup>${run}</m:sup></m:sSup>`,
      `<m:sSup><m:sSupPr><m:ctrlPr><w:rPr xmlns:w="${WORD_NAMESPACE}"/></m:ctrlPr></m:sSupPr><m:e>${run}</m:e><m:sup>${run}</m:sup></m:sSup>`,
      `<m:sSup><v:sSupPr xmlns:v="${VENDOR_NAMESPACE}"/><m:e>${run}</m:e><m:sup>${run}</m:sup></m:sSup>`,
      `<m:sSub><m:sub>${run}</m:sub><m:e>${run}</m:e></m:sSub>`,
      `<m:sSubSup><m:e>${run}</m:e><m:sSubSupPr/><m:sub>${run}</m:sub><m:sup>${run}</m:sup></m:sSubSup>`,
      `<m:sSubSup><m:sSubSupPr><m:ctrlPr/><m:alnScr/></m:sSubSupPr><m:e>${run}</m:e><m:sub>${run}</m:sub><m:sup>${run}</m:sup></m:sSubSup>`,
      `<m:sSubSup><m:sSubSupPr><m:alnScr/><m:alnScr/></m:sSubSupPr><m:e>${run}</m:e><m:sub>${run}</m:sub><m:sup>${run}</m:sup></m:sSubSup>`,
      `<m:sSubSup><m:sSubSupPr><m:alnScr m:val="maybe"/></m:sSubSupPr><m:e>${run}</m:e><m:sub>${run}</m:sub><m:sup>${run}</m:sup></m:sSubSup>`,
      `<m:sSubSup><m:sSubSupPr><m:alnScr m:val="1" m:extra="semantic"/></m:sSubSupPr><m:e>${run}</m:e><m:sub>${run}</m:sub><m:sup>${run}</m:sup></m:sSubSup>`,
      `<m:sSubSup xmlns:r="${RELATIONSHIP_NAMESPACE}"><m:sSubSupPr><m:alnScr r:id="rIdUnsafe"/></m:sSubSupPr><m:e>${run}</m:e><m:sub>${run}</m:sub><m:sup>${run}</m:sup></m:sSubSup>`,
      `<m:sSubSup><m:sSubSupPr><v:alnScr xmlns:v="${VENDOR_NAMESPACE}"/></m:sSubSupPr><m:e>${run}</m:e><m:sub>${run}</m:sub><m:sup>${run}</m:sup></m:sSubSup>`,
      `<m:sSubSup><m:sSubSupPr>meaningful<m:alnScr/></m:sSubSupPr><m:e>${run}</m:e><m:sub>${run}</m:sub><m:sup>${run}</m:sup></m:sSubSup>`,
      `<m:sSubSup><m:sSubSupPr/><m:e>${run}</m:e><m:sub>${run}</m:sub></m:sSubSup>`,
      `<m:sSubSup><m:e>${run}</m:e><m:sub>${run}</m:sub><m:sub>${run}</m:sub><m:sup>${run}</m:sup></m:sSubSup>`,
      `<m:func><m:fName>${run}</m:fName><m:funcPr/><m:e>${run}</m:e></m:func>`,
      `<m:func><m:funcPr><m:ctrlPr xmlns:r="${RELATIONSHIP_NAMESPACE}" r:id="rIdUnsafe"/></m:funcPr><m:fName>${run}</m:fName><m:e>${run}</m:e></m:func>`,
    ];
    expect(unsupported.map(inspectEquationBody)).toEqual(
      unsupported.map(() => 'unsupported'),
    );
  });

  test('normalizes fraction type defaults and strictly validates fraction structure', () => {
    const run = '<m:r><m:t>x</m:t></m:r>';
    const supported = [
      `<m:f><m:num>${run}</m:num><m:den>${run}</m:den></m:f>`,
      `<m:f><m:fPr/><m:num>${run}</m:num><m:den>${run}</m:den></m:f>`,
      `<m:f><m:fPr><m:type/></m:fPr><m:num>${run}</m:num><m:den>${run}</m:den></m:f>`,
      `<m:f><m:fPr><m:type m:val="bar"/><m:ctrlPr/></m:fPr><m:num>${run}</m:num><m:den>${run}</m:den></m:f>`,
      `<m:f><m:fPr><m:type m:val="noBar"/></m:fPr><m:num>${run}</m:num><m:den>${run}</m:den></m:f>`,
      `<m:f><m:fPr><m:type m:val="skw"/></m:fPr><m:num>${run}</m:num><m:den>${run}</m:den></m:f>`,
      `<m:f><m:fPr><m:type m:val="lin"/></m:fPr><m:num>${run}</m:num><m:den>${run}</m:den></m:f>`,
    ];
    expect(supported.map(inspectEquationBody)).toEqual(
      supported.map(() => 'supported'),
    );
    expect(
      supported.map((source) => {
        const expression = inspectEquationModel(source)?.children[0];
        return expression?.type === 'fraction' ? expression.fractionType : null;
      }),
    ).toEqual(['bar', 'bar', 'bar', 'bar', 'noBar', 'skewed', 'linear']);

    const unsupported = [
      `<m:f><m:num>${run}</m:num><m:fPr/><m:den>${run}</m:den></m:f>`,
      `<m:f><m:den>${run}</m:den><m:num>${run}</m:num></m:f>`,
      `<m:f><m:fPr/><m:fPr/><m:num>${run}</m:num><m:den>${run}</m:den></m:f>`,
      `<m:f><m:fPr><m:ctrlPr/><m:type/></m:fPr><m:num>${run}</m:num><m:den>${run}</m:den></m:f>`,
      `<m:f><m:fPr><m:type/><m:type/></m:fPr><m:num>${run}</m:num><m:den>${run}</m:den></m:f>`,
      `<m:f><m:fPr><m:type m:val="diagonal"/></m:fPr><m:num>${run}</m:num><m:den>${run}</m:den></m:f>`,
      `<m:f><m:fPr><m:type m:val="bar" m:extra="semantic"/></m:fPr><m:num>${run}</m:num><m:den>${run}</m:den></m:f>`,
      `<m:f xmlns:r="${RELATIONSHIP_NAMESPACE}"><m:fPr><m:type m:val="bar" r:id="rIdUnsafe"/></m:fPr><m:num>${run}</m:num><m:den>${run}</m:den></m:f>`,
      `<m:f><v:fPr xmlns:v="${VENDOR_NAMESPACE}"><m:type m:val="bar"/></v:fPr><m:num>${run}</m:num><m:den>${run}</m:den></m:f>`,
      `<m:f><m:fPr><v:type xmlns:v="${VENDOR_NAMESPACE}" v:val="bar"/></m:fPr><m:num>${run}</m:num><m:den>${run}</m:den></m:f>`,
      `<m:f><m:fPr><m:ctrlPr><w:rPr xmlns:w="${WORD_NAMESPACE}"/></m:ctrlPr></m:fPr><m:num>${run}</m:num><m:den>${run}</m:den></m:f>`,
      `<m:f><m:fPr>meaningful<m:type/></m:fPr><m:num>${run}</m:num><m:den>${run}</m:den></m:f>`,
      `<m:f><m:num>meaningful</m:num><m:den>${run}</m:den></m:f>`,
      `<m:f><m:num/><m:den>${run}</m:den></m:f>`,
      `<m:f><m:num>${run}</m:num><m:den/></m:f>`,
      `<m:f><m:num>${run}</m:num></m:f>`,
      `<m:f><m:num>${run}</m:num><m:num>${run}</m:num><m:den>${run}</m:den></m:f>`,
      `<m:f><m:num>${run}</m:num><m:den>${run}</m:den><m:den>${run}</m:den></m:f>`,
    ];
    expect(unsupported.map(inspectEquationBody)).toEqual(
      unsupported.map(() => 'unsupported'),
    );
  });

  test('normalizes radical degree defaults and strictly validates radical structure', () => {
    const run = '<m:r><m:t>x</m:t></m:r>';
    const supported = [
      `<m:rad><m:e>${run}</m:e></m:rad>`,
      `<m:rad><m:radPr/><m:e>${run}</m:e></m:rad>`,
      `<m:rad><m:radPr><m:degHide/></m:radPr><m:deg/><m:e>${run}</m:e></m:rad>`,
      `<m:rad><m:radPr><m:degHide m:val="0"/><m:ctrlPr/></m:radPr><m:deg/><m:e>${run}</m:e></m:rad>`,
      `<m:rad><m:deg>${run}</m:deg><m:e>${run}</m:e></m:rad>`,
      `<m:rad><m:radPr><m:degHide m:val="false"/><m:ctrlPr/></m:radPr><m:deg>${run}</m:deg><m:e>${run}</m:e></m:rad>`,
    ];
    expect(supported.map(inspectEquationBody)).toEqual(
      supported.map(() => 'supported'),
    );
    expect(
      supported.map((source) => inspectEquationModel(source)?.children[0]),
    ).toEqual([
      { type: 'radical', children: [{ type: 'run', text: 'x' }] },
      { type: 'radical', children: [{ type: 'run', text: 'x' }] },
      { type: 'radical', children: [{ type: 'run', text: 'x' }] },
      { type: 'radical', children: [{ type: 'run', text: 'x' }] },
      {
        type: 'radical',
        children: [{ type: 'run', text: 'x' }],
        degree: [{ type: 'run', text: 'x' }],
      },
      {
        type: 'radical',
        children: [{ type: 'run', text: 'x' }],
        degree: [{ type: 'run', text: 'x' }],
      },
    ]);

    const unsupported = [
      `<m:rad><m:e>${run}</m:e><m:deg>${run}</m:deg></m:rad>`,
      `<m:rad><m:deg>${run}</m:deg><m:radPr/><m:e>${run}</m:e></m:rad>`,
      `<m:rad><m:radPr/><m:radPr/><m:e>${run}</m:e></m:rad>`,
      `<m:rad><m:deg>${run}</m:deg><m:deg>${run}</m:deg><m:e>${run}</m:e></m:rad>`,
      `<m:rad><m:radPr><m:ctrlPr/><m:degHide/></m:radPr><m:e>${run}</m:e></m:rad>`,
      `<m:rad><m:radPr><m:degHide/><m:degHide/></m:radPr><m:e>${run}</m:e></m:rad>`,
      `<m:rad><m:radPr><m:degHide m:val="maybe"/></m:radPr><m:e>${run}</m:e></m:rad>`,
      `<m:rad><m:radPr><m:degHide m:val="1" m:extra="semantic"/></m:radPr><m:e>${run}</m:e></m:rad>`,
      `<m:rad xmlns:r="${RELATIONSHIP_NAMESPACE}"><m:radPr><m:degHide r:id="rIdUnsafe"/></m:radPr><m:e>${run}</m:e></m:rad>`,
      `<m:rad><v:radPr xmlns:v="${VENDOR_NAMESPACE}"><m:degHide/></v:radPr><m:e>${run}</m:e></m:rad>`,
      `<m:rad><m:radPr><v:degHide xmlns:v="${VENDOR_NAMESPACE}"/></m:radPr><m:e>${run}</m:e></m:rad>`,
      `<m:rad><m:radPr><m:ctrlPr><w:rPr xmlns:w="${WORD_NAMESPACE}"/></m:ctrlPr></m:radPr><m:e>${run}</m:e></m:rad>`,
      `<m:rad><m:radPr>meaningful<m:degHide/></m:radPr><m:e>${run}</m:e></m:rad>`,
      `<m:rad><m:radPr><m:degHide/></m:radPr><m:deg>${run}</m:deg><m:e>${run}</m:e></m:rad>`,
      `<m:rad><m:deg>meaningful</m:deg><m:e>${run}</m:e></m:rad>`,
      `<m:rad><m:deg>${run}</m:deg></m:rad>`,
      `<m:rad><m:e/></m:rad>`,
      `<m:rad><m:e>${run}</m:e><m:e>${run}</m:e></m:rad>`,
    ];
    expect(unsupported.map(inspectEquationBody)).toEqual(
      unsupported.map(() => 'unsupported'),
    );
  });

  test('normalizes n-ary defaults and strictly validates operator structure', () => {
    const run = '<m:r><m:t>x</m:t></m:r>';
    const supported = [
      `<m:nary><m:sub>${run}</m:sub><m:sup>${run}</m:sup><m:e>${run}</m:e></m:nary>`,
      `<m:nary><m:naryPr/><m:sub>${run}</m:sub><m:sup>${run}</m:sup><m:e>${run}</m:e></m:nary>`,
      `<m:nary><m:naryPr><m:chr m:val="&#x2211;"/></m:naryPr><m:sub>${run}</m:sub><m:sup>${run}</m:sup><m:e>${run}</m:e></m:nary>`,
      `<m:nary><m:naryPr><m:chr m:val="&#x2211;"/><m:limLoc/><m:grow m:val="0"/><m:subHide m:val="false"/><m:supHide m:val="off"/><m:ctrlPr/></m:naryPr><m:sub>${run}</m:sub><m:sup>${run}</m:sup><m:e>${run}</m:e></m:nary>`,
      `<m:nary><m:naryPr><m:chr m:val="&#x2211;"/><m:limLoc m:val="subSup"/><m:subHide/><m:supHide m:val="true"/></m:naryPr><m:sub/><m:sup/><m:e>${run}</m:e></m:nary>`,
      `<m:nary><m:naryPr><m:chr m:val="&#x220F;"/><m:grow m:val="off"/><m:subHide m:val="on"/><m:supHide m:val="0"/></m:naryPr><m:sub/><m:sup>${run}</m:sup><m:e>${run}</m:e></m:nary>`,
      `<m:nary><m:naryPr><m:limLoc/></m:naryPr><m:sub>${run}</m:sub><m:sup>${run}</m:sup><m:e>${run}</m:e></m:nary>`,
    ];
    expect(supported.map(inspectEquationBody)).toEqual(
      supported.map(() => 'supported'),
    );
    expect(
      supported.map((source) => {
        const expression = inspectEquationModel(source)?.children[0];
        return expression?.type === 'nary'
          ? [
              expression.operator,
              expression.limitLocation,
              expression.subScript?.[0]?.type ?? null,
              expression.superScript?.[0]?.type ?? null,
            ]
          : null;
      }),
    ).toEqual([
      ['\u222b', 'subSup', 'run', 'run'],
      ['\u222b', 'subSup', 'run', 'run'],
      ['\u2211', 'underOver', 'run', 'run'],
      ['\u2211', 'underOver', 'run', 'run'],
      ['\u2211', 'subSup', null, null],
      ['\u220f', 'underOver', null, 'run'],
      ['\u222b', 'underOver', 'run', 'run'],
    ]);

    const unsupported = [
      `<m:nary><m:sub>${run}</m:sub><m:naryPr/><m:sup>${run}</m:sup><m:e>${run}</m:e></m:nary>`,
      `<m:nary><m:sup>${run}</m:sup><m:sub>${run}</m:sub><m:e>${run}</m:e></m:nary>`,
      `<m:nary><m:sub>${run}</m:sub><m:e>${run}</m:e><m:sup>${run}</m:sup></m:nary>`,
      `<m:nary><m:naryPr/><m:naryPr/><m:sub>${run}</m:sub><m:sup>${run}</m:sup><m:e>${run}</m:e></m:nary>`,
      `<m:nary><m:sub>${run}</m:sub><m:sub>${run}</m:sub><m:sup>${run}</m:sup><m:e>${run}</m:e></m:nary>`,
      `<m:nary><m:sub>${run}</m:sub><m:sup>${run}</m:sup><m:sup>${run}</m:sup><m:e>${run}</m:e></m:nary>`,
      `<m:nary><m:sub>${run}</m:sub><m:sup>${run}</m:sup><m:e>${run}</m:e><m:e>${run}</m:e></m:nary>`,
      `<m:nary><m:sup>${run}</m:sup><m:e>${run}</m:e></m:nary>`,
      `<m:nary><m:sub>${run}</m:sub><m:e>${run}</m:e></m:nary>`,
      `<m:nary><m:sub>${run}</m:sub><m:sup>${run}</m:sup></m:nary>`,
      `<m:nary><m:sub/><m:sup>${run}</m:sup><m:e>${run}</m:e></m:nary>`,
      `<m:nary><m:sub>${run}</m:sub><m:sup/><m:e>${run}</m:e></m:nary>`,
      `<m:nary><m:sub>${run}</m:sub><m:sup>${run}</m:sup><m:e/></m:nary>`,
      `<m:nary><m:naryPr><m:chr/></m:naryPr><m:sub>${run}</m:sub><m:sup>${run}</m:sup><m:e>${run}</m:e></m:nary>`,
      `<m:nary><m:naryPr><m:chr m:val="+"/></m:naryPr><m:sub>${run}</m:sub><m:sup>${run}</m:sup><m:e>${run}</m:e></m:nary>`,
      `<m:nary><m:naryPr><m:limLoc/><m:chr m:val="&#x2211;"/></m:naryPr><m:sub>${run}</m:sub><m:sup>${run}</m:sup><m:e>${run}</m:e></m:nary>`,
      `<m:nary><m:naryPr><m:supHide/><m:subHide/></m:naryPr><m:sub/><m:sup/><m:e>${run}</m:e></m:nary>`,
      `<m:nary><m:naryPr><m:ctrlPr/><m:subHide/></m:naryPr><m:sub/><m:sup>${run}</m:sup><m:e>${run}</m:e></m:nary>`,
      `<m:nary><m:naryPr><m:chr m:val="&#x2211;"/><m:chr m:val="&#x220F;"/></m:naryPr><m:sub>${run}</m:sub><m:sup>${run}</m:sup><m:e>${run}</m:e></m:nary>`,
      `<m:nary><m:naryPr><m:limLoc m:val="beside"/></m:naryPr><m:sub>${run}</m:sub><m:sup>${run}</m:sup><m:e>${run}</m:e></m:nary>`,
      `<m:nary><m:naryPr><m:grow/></m:naryPr><m:sub>${run}</m:sub><m:sup>${run}</m:sup><m:e>${run}</m:e></m:nary>`,
      `<m:nary><m:naryPr><m:grow m:val="true"/></m:naryPr><m:sub>${run}</m:sub><m:sup>${run}</m:sup><m:e>${run}</m:e></m:nary>`,
      `<m:nary><m:naryPr><m:grow m:val="maybe"/></m:naryPr><m:sub>${run}</m:sub><m:sup>${run}</m:sup><m:e>${run}</m:e></m:nary>`,
      `<m:nary><m:naryPr><m:subHide/></m:naryPr><m:sub>${run}</m:sub><m:sup>${run}</m:sup><m:e>${run}</m:e></m:nary>`,
      `<m:nary><m:naryPr><m:supHide/></m:naryPr><m:sub>${run}</m:sub><m:sup>${run}</m:sup><m:e>${run}</m:e></m:nary>`,
      `<m:nary><m:naryPr><m:subHide m:val="maybe"/></m:naryPr><m:sub/><m:sup>${run}</m:sup><m:e>${run}</m:e></m:nary>`,
      `<m:nary><m:naryPr><m:chr m:val="&#x2211;" m:extra="semantic"/></m:naryPr><m:sub>${run}</m:sub><m:sup>${run}</m:sup><m:e>${run}</m:e></m:nary>`,
      `<m:nary xmlns:r="${RELATIONSHIP_NAMESPACE}"><m:naryPr><m:chr m:val="&#x2211;" r:id="rIdUnsafe"/></m:naryPr><m:sub>${run}</m:sub><m:sup>${run}</m:sup><m:e>${run}</m:e></m:nary>`,
      `<m:nary><v:naryPr xmlns:v="${VENDOR_NAMESPACE}"><m:chr m:val="&#x2211;"/></v:naryPr><m:sub>${run}</m:sub><m:sup>${run}</m:sup><m:e>${run}</m:e></m:nary>`,
      `<m:nary><m:naryPr><v:chr xmlns:v="${VENDOR_NAMESPACE}" v:val="&#x2211;"/></m:naryPr><m:sub>${run}</m:sub><m:sup>${run}</m:sup><m:e>${run}</m:e></m:nary>`,
      `<m:nary><m:naryPr><m:ctrlPr><w:rPr xmlns:w="${WORD_NAMESPACE}"/></m:ctrlPr></m:naryPr><m:sub>${run}</m:sub><m:sup>${run}</m:sup><m:e>${run}</m:e></m:nary>`,
      `<m:nary><m:naryPr>meaningful<m:chr m:val="&#x2211;"/></m:naryPr><m:sub>${run}</m:sub><m:sup>${run}</m:sup><m:e>${run}</m:e></m:nary>`,
      `<m:nary m:extra="semantic"><m:sub>${run}</m:sub><m:sup>${run}</m:sup><m:e>${run}</m:e></m:nary>`,
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
    `<aside data-document-note="true" data-note-kind="footnote" data-note-id="foot-equation" data-note-number="1"><p>Foot ${equationHtml(borderBoxEquation('F', 'top'))}</p></aside>`,
    `<aside data-document-note="true" data-note-kind="endnote" data-note-id="end-equation" data-note-number="1"><p>End ${equationHtml(boxEquation('N', 'bottom'))}</p></aside>`,
  ].join('');
  artifact.content.pageChrome = {
    differentFirstPage: false,
    differentOddEvenPages: false,
    default: {
      headerHtml: `<p>Header ${equationHtml(borderBoxEquation('H', 'top'))}</p>`,
      footerHtml: `<p>Footer ${equationHtml(boxEquation('R', 'bottom'))}</p>`,
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
    ...(display === 'block' ? { justification: 'right' as const } : {}),
    children: [
      run('x='),
      {
        type: 'run',
        text: 'styledF',
        literal: true,
        script: 'fraktur',
        style: 'bold',
        manualBreak: { alignmentAt: 3 },
        alignment: true,
      },
      {
        type: 'run',
        text: 'normalRate',
        normalText: true,
        style: 'plain',
      },
      {
        type: 'run',
        text: 'doubleR',
        script: 'doubleStruck',
        style: 'boldItalic',
      },
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
        alignScripts: true,
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
        type: 'nary',
        operator: '∏',
        limitLocation: 'underOver',
        children: [run('a_i')],
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
        type: 'groupCharacter',
        character: '\u23de',
        position: 'top',
        verticalJustification: 'bottom',
        children: [run('x+y')],
      },
      {
        type: 'groupCharacter',
        character: '\u23df',
        position: 'bottom',
        verticalJustification: 'top',
        children: [run('a-b')],
      },
      {
        type: 'phantom',
        show: false,
        zeroWidth: true,
        zeroAscent: false,
        zeroDescent: true,
        transparent: true,
        children: [run('ghost')],
      },
      {
        type: 'phantom',
        show: true,
        zeroWidth: false,
        zeroAscent: true,
        zeroDescent: false,
        transparent: false,
        children: [run('visible')],
      },
      {
        type: 'preSubSuperScript',
        base: [run('T')],
        subScript: [run('i')],
        superScript: [run('j')],
      },
      {
        type: 'preSubSuperScript',
        base: [run('A')],
        subScript: [],
        superScript: [run('2')],
      },
      {
        type: 'preSubSuperScript',
        base: [run('B')],
        subScript: [run('1')],
        superScript: [],
      },
      {
        type: 'borderBox',
        hideTop: false,
        hideBottom: true,
        hideLeft: false,
        hideRight: true,
        strikeHorizontal: true,
        strikeVertical: false,
        strikeBottomLeftToTopRight: true,
        strikeTopLeftToBottomRight: false,
        children: [run('boxed')],
      },
      {
        type: 'box',
        operatorEmulator: true,
        noBreak: true,
        differential: true,
        alignment: true,
        manualBreak: { alignmentAt: 3 },
        children: [run('dx')],
      },
      {
        type: 'equationArray',
        baseAlignment: 'bottom',
        maximumDistribution: true,
        objectDistribution: false,
        rowSpacingRule: 'multiple',
        rowSpacing: 3,
        rows: [[run('&x+&&y=1')], [run('2&x+&&y=3')], []],
      },
      {
        type: 'lowerLimit',
        base: [run('lim')],
        limit: [run('x→0')],
      },
      {
        type: 'upperLimit',
        base: [run('=')],
        limit: [run('def')],
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

function borderBoxEquation(
  text: string,
  barPosition: 'top' | 'bottom',
): WorkDocumentEquation {
  return {
    version: 1,
    display: 'inline',
    children: [
      {
        type: 'borderBox',
        hideTop: false,
        hideBottom: false,
        hideLeft: false,
        hideRight: false,
        strikeHorizontal: false,
        strikeVertical: false,
        strikeBottomLeftToTopRight: false,
        strikeTopLeftToBottomRight: false,
        children: [
          {
            type: 'equationArray',
            baseAlignment: 'center',
            maximumDistribution: false,
            objectDistribution: false,
            rowSpacingRule: 'single',
            rowSpacing: 0,
            rows: [
              [
                {
                  type: 'lowerLimit',
                  base: [
                    {
                      type: 'preSubSuperScript',
                      base: [
                        {
                          type: 'phantom',
                          show: true,
                          zeroWidth: false,
                          zeroAscent: true,
                          zeroDescent: false,
                          transparent: true,
                          children: [
                            {
                              type: 'groupCharacter',
                              character:
                                barPosition === 'top' ? '\u23de' : '\u23df',
                              position: barPosition,
                              verticalJustification:
                                barPosition === 'top' ? 'bottom' : 'top',
                              children: [
                                {
                                  type: 'bar',
                                  position: barPosition,
                                  children: [
                                    {
                                      type: 'subSuperScript',
                                      alignScripts: true,
                                      base: [
                                        {
                                          type: 'radical',
                                          children: [
                                            {
                                              type: 'nary',
                                              operator: '\u222b',
                                              limitLocation: 'subSup',
                                              children: [
                                                {
                                                  type: 'fraction',
                                                  fractionType: 'bar',
                                                  numerator: [
                                                    {
                                                      type: 'run',
                                                      text,
                                                      literal: true,
                                                      script: 'doubleStruck',
                                                      style: 'plain',
                                                      manualBreak: {
                                                        alignmentAt: 2,
                                                      },
                                                    },
                                                  ],
                                                  denominator: [
                                                    { type: 'run', text: '1' },
                                                  ],
                                                },
                                              ],
                                              superScript: [
                                                { type: 'run', text: 'n' },
                                              ],
                                            },
                                          ],
                                        },
                                      ],
                                      subScript: [{ type: 'run', text: '1' }],
                                      superScript: [{ type: 'run', text: '2' }],
                                    },
                                  ],
                                },
                              ],
                            },
                          ],
                        },
                      ],
                      subScript: [{ type: 'run', text: 'p' }],
                      superScript: [],
                    },
                  ],
                  limit: [{ type: 'run', text: 'i' }],
                },
              ],
            ],
          },
        ],
      },
    ],
  };
}

function boxEquation(
  text: string,
  barPosition: 'top' | 'bottom',
): WorkDocumentEquation {
  return {
    version: 1,
    display: 'inline',
    children: [
      {
        type: 'box',
        operatorEmulator: true,
        noBreak: true,
        differential: false,
        alignment: true,
        manualBreak: {},
        children: [
          {
            type: 'equationArray',
            baseAlignment: 'center',
            maximumDistribution: false,
            objectDistribution: false,
            rowSpacingRule: 'single',
            rowSpacing: 0,
            rows: [
              [
                {
                  type: 'upperLimit',
                  base: [
                    {
                      type: 'preSubSuperScript',
                      base: [
                        {
                          type: 'phantom',
                          show: false,
                          zeroWidth: true,
                          zeroAscent: false,
                          zeroDescent: true,
                          transparent: false,
                          children: [
                            {
                              type: 'groupCharacter',
                              character:
                                barPosition === 'top' ? '\u23de' : '\u23df',
                              position: barPosition,
                              verticalJustification:
                                barPosition === 'top' ? 'bottom' : 'top',
                              children: [
                                {
                                  type: 'bar',
                                  position: barPosition,
                                  children: [
                                    {
                                      type: 'subSuperScript',
                                      alignScripts: true,
                                      base: [
                                        {
                                          type: 'radical',
                                          children: [
                                            {
                                              type: 'nary',
                                              operator: '\u2211',
                                              limitLocation: 'underOver',
                                              children: [
                                                {
                                                  type: 'fraction',
                                                  fractionType: 'linear',
                                                  numerator: [
                                                    {
                                                      type: 'run',
                                                      text,
                                                      normalText: true,
                                                      script: 'sansSerif',
                                                      style: 'boldItalic',
                                                      manualBreak: {},
                                                      alignment: true,
                                                    },
                                                  ],
                                                  denominator: [
                                                    { type: 'run', text: '1' },
                                                  ],
                                                },
                                              ],
                                              subScript: [
                                                { type: 'run', text: 'k' },
                                              ],
                                            },
                                          ],
                                        },
                                      ],
                                      subScript: [{ type: 'run', text: '1' }],
                                      superScript: [{ type: 'run', text: '2' }],
                                    },
                                  ],
                                },
                              ],
                            },
                          ],
                        },
                      ],
                      subScript: [],
                      superScript: [{ type: 'run', text: 'q' }],
                    },
                  ],
                  limit: [{ type: 'run', text: 'j' }],
                },
              ],
            ],
          },
        ],
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

function inspectEquationRoot(root: string) {
  return inspectDocxEquation(parseXml(root, 'equation.xml').documentElement);
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
  const mathParagraphs = descendants(document, 'oMathPara');
  expect(mathParagraphs).toHaveLength(1);
  expect(
    directChildren(mathParagraphs[0]).map((child) => child.localName),
  ).toEqual(['oMathParaPr', 'oMath']);
  const mathParagraphProperties = directChildren(
    mathParagraphs[0],
    'oMathParaPr',
  )[0];
  expect(
    directChildren(mathParagraphProperties).map((child) => child.localName),
  ).toEqual(['jc']);
  expect(
    mathValueAttribute(directChildren(mathParagraphProperties, 'jc')[0]),
  ).toBe('right');
  expect(descendants(document, 'oMath')).toHaveLength(2);
  const fractions = descendants(document, 'f');
  expect(fractions).toHaveLength(8);
  expect(
    fractions.map((fraction) =>
      directChildren(fraction).map((child) => child.localName),
    ),
  ).toEqual([
    ['num', 'den'],
    ['fPr', 'num', 'den'],
    ['fPr', 'num', 'den'],
    ['fPr', 'num', 'den'],
    ['num', 'den'],
    ['fPr', 'num', 'den'],
    ['fPr', 'num', 'den'],
    ['fPr', 'num', 'den'],
  ]);
  expect(
    fractions.map((fraction) => {
      const properties = directChildren(fraction, 'fPr')[0];
      return properties
        ? mathValueAttribute(directChildren(properties, 'type')[0])
        : null;
    }),
  ).toEqual([null, 'noBar', 'skw', 'lin', null, 'noBar', 'skw', 'lin']);
  expect(descendants(document, 'sSup')).toHaveLength(2);
  expect(descendants(document, 'sSub')).toHaveLength(2);
  const alignedScripts = descendants(document, 'sSubSup');
  expect(alignedScripts).toHaveLength(2);
  for (const script of alignedScripts) {
    expect(directChildren(script).map((child) => child.localName)).toEqual([
      'sSubSupPr',
      'e',
      'sub',
      'sup',
    ]);
    const properties = directChildren(script, 'sSubSupPr')[0];
    expect(directChildren(properties).map((child) => child.localName)).toEqual([
      'alnScr',
    ]);
    expect(mathValueAttribute(directChildren(properties, 'alnScr')[0])).toBe(
      '1',
    );
  }
  const radicals = descendants(document, 'rad');
  expect(radicals).toHaveLength(4);
  for (const radical of radicals) {
    expect(directChildren(radical).map((child) => child.localName)).toEqual([
      'radPr',
      'deg',
      'e',
    ]);
  }
  expect(
    radicals.map((radical) =>
      directChildren(directChildren(radical, 'radPr')[0]).map(
        (child) => child.localName,
      ),
    ),
  ).toEqual([['degHide'], [], ['degHide'], []]);
  expect(
    radicals.map(
      (radical) => directChildren(radical, 'deg')[0]?.textContent ?? '',
    ),
  ).toEqual(['', '3', '', '3']);
  expect(
    radicals
      .filter((_, index) => index % 2 === 0)
      .map((radical) =>
        mathValueAttribute(
          directChildren(directChildren(radical, 'radPr')[0], 'degHide')[0],
        ),
      ),
  ).toEqual(['1', '1']);
  expect(descendants(document, 'func')).toHaveLength(2);
  const naries = descendants(document, 'nary');
  expect(naries).toHaveLength(6);
  for (const nary of naries) {
    expect(directChildren(nary).map((child) => child.localName)).toEqual([
      'naryPr',
      'sub',
      'sup',
      'e',
    ]);
  }
  expect(
    naries.map((nary) =>
      directChildren(directChildren(nary, 'naryPr')[0]).map(
        (child) => child.localName,
      ),
    ),
  ).toEqual([
    ['chr', 'limLoc'],
    ['chr', 'limLoc'],
    ['chr', 'limLoc', 'subHide', 'supHide'],
    ['chr', 'limLoc'],
    ['chr', 'limLoc'],
    ['chr', 'limLoc', 'subHide', 'supHide'],
  ]);
  expect(
    naries.map((nary) => {
      const properties = directChildren(nary, 'naryPr')[0];
      return [
        mathValueAttribute(directChildren(properties, 'chr')[0]),
        mathValueAttribute(directChildren(properties, 'limLoc')[0]),
        directChildren(nary, 'sub')[0]?.textContent ?? '',
        directChildren(nary, 'sup')[0]?.textContent ?? '',
      ];
    }),
  ).toEqual([
    ['\u2211', 'undOvr', 'i=1', 'n'],
    ['\u222b', 'subSup', '0', '1'],
    ['\u220f', 'undOvr', '', ''],
    ['\u2211', 'undOvr', 'i=1', 'n'],
    ['\u222b', 'subSup', '0', '1'],
    ['\u220f', 'undOvr', '', ''],
  ]);
  expect(
    naries
      .filter((_, index) => index % 3 === 2)
      .flatMap((nary) => {
        const properties = directChildren(nary, 'naryPr')[0];
        return ['subHide', 'supHide'].map((name) =>
          mathValueAttribute(directChildren(properties, name)[0]),
        );
      }),
  ).toEqual(['1', '1', '1', '1']);
  expect(descendants(document, 'd')).toHaveLength(2);
  const preScripts = descendants(document, 'sPre');
  expect(preScripts).toHaveLength(6);
  for (const preScript of preScripts) {
    expect(directChildren(preScript).map((child) => child.localName)).toEqual([
      'sub',
      'sup',
      'e',
    ]);
  }
  expect(
    preScripts.map((preScript) =>
      ['sub', 'sup', 'e'].map(
        (name) => directChildren(preScript, name)[0]?.textContent ?? '',
      ),
    ),
  ).toEqual([
    ['i', 'j', 'T'],
    ['', '2', 'A'],
    ['1', '', 'B'],
    ['i', 'j', 'T'],
    ['', '2', 'A'],
    ['1', '', 'B'],
  ]);
  const styledRuns = descendants(document, 'r').filter(
    (run) => run.namespaceURI === MATH_NAMESPACE && mathRunProperties(run),
  );
  expect(
    styledRuns.map((run) => [run.textContent, mathRunProperties(run)]),
  ).toEqual([
    [
      'styledF',
      {
        children: ['lit', 'scr', 'sty', 'brk', 'aln'],
        values: ['1', 'fraktur', 'b', '3', '1'],
      },
    ],
    ['normalRate', { children: ['nor', 'sty'], values: ['1', 'p'] }],
    ['doubleR', { children: ['scr', 'sty'], values: ['double-struck', 'bi'] }],
    [
      'styledF',
      {
        children: ['lit', 'scr', 'sty', 'brk', 'aln'],
        values: ['1', 'fraktur', 'b', '3', '1'],
      },
    ],
    ['normalRate', { children: ['nor', 'sty'], values: ['1', 'p'] }],
    ['doubleR', { children: ['scr', 'sty'], values: ['double-struck', 'bi'] }],
  ]);
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
  const groupCharacters = descendants(document, 'groupChr');
  expect(groupCharacters).toHaveLength(4);
  for (const groupCharacter of groupCharacters) {
    expect(
      directChildren(groupCharacter).map((child) => child.localName),
    ).toEqual(['groupChrPr', 'e']);
    expect(
      directChildren(directChildren(groupCharacter, 'groupChrPr')[0]).map(
        (child) => child.localName,
      ),
    ).toEqual(['chr', 'pos', 'vertJc']);
  }
  expect(
    groupCharacters.map((groupCharacter) => {
      const properties = directChildren(groupCharacter, 'groupChrPr')[0];
      return ['chr', 'pos', 'vertJc'].map((name) =>
        mathValueAttribute(directChildren(properties, name)[0]),
      );
    }),
  ).toEqual([
    ['\u23de', 'top', 'bot'],
    ['\u23df', 'bot', 'top'],
    ['\u23de', 'top', 'bot'],
    ['\u23df', 'bot', 'top'],
  ]);
  const phantoms = descendants(document, 'phant');
  expect(phantoms).toHaveLength(4);
  for (const phantom of phantoms) {
    expect(directChildren(phantom).map((child) => child.localName)).toEqual([
      'phantPr',
      'e',
    ]);
    expect(
      directChildren(directChildren(phantom, 'phantPr')[0]).map(
        (child) => child.localName,
      ),
    ).toEqual(['show', 'zeroWid', 'zeroAsc', 'zeroDesc', 'transp']);
  }
  expect(
    phantoms.map((phantom) =>
      directChildren(directChildren(phantom, 'phantPr')[0]).map(
        mathValueAttribute,
      ),
    ),
  ).toEqual([
    ['0', '1', '0', '1', '1'],
    ['1', '0', '1', '0', '0'],
    ['0', '1', '0', '1', '1'],
    ['1', '0', '1', '0', '0'],
  ]);
  const borderBoxes = descendants(document, 'borderBox');
  expect(borderBoxes).toHaveLength(2);
  for (const borderBox of borderBoxes) {
    const properties = directChildren(borderBox, 'borderBoxPr')[0];
    expect(directChildren(properties).map((child) => child.localName)).toEqual([
      'hideTop',
      'hideBot',
      'hideLeft',
      'hideRight',
      'strikeH',
      'strikeV',
      'strikeBLTR',
      'strikeTLBR',
    ]);
    expect(directChildren(properties).map(mathValueAttribute)).toEqual([
      '0',
      '1',
      '0',
      '1',
      '1',
      '0',
      '1',
      '0',
    ]);
    expect(directChildren(borderBox, 'e')).toHaveLength(1);
  }
  const boxes = descendants(document, 'box');
  expect(boxes).toHaveLength(2);
  for (const box of boxes) {
    const properties = directChildren(box, 'boxPr')[0];
    expect(directChildren(properties).map((child) => child.localName)).toEqual([
      'opEmu',
      'noBreak',
      'diff',
      'brk',
      'aln',
    ]);
    expect(
      ['opEmu', 'noBreak', 'diff', 'aln'].map((name) =>
        mathValueAttribute(directChildren(properties, name)[0]),
      ),
    ).toEqual(['1', '1', '1', '1']);
    expect(
      mathNamedAttribute(directChildren(properties, 'brk')[0], 'alnAt'),
    ).toBe('3');
    expect(directChildren(box, 'e')).toHaveLength(1);
  }
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
  const equationArrays = descendants(document, 'eqArr');
  expect(equationArrays).toHaveLength(2);
  for (const equationArray of equationArrays) {
    const properties = directChildren(equationArray, 'eqArrPr')[0];
    expect(directChildren(properties).map((child) => child.localName)).toEqual([
      'baseJc',
      'maxDist',
      'objDist',
      'rSpRule',
      'rSp',
    ]);
    expect(directChildren(properties).map(mathValueAttribute)).toEqual([
      'bot',
      '1',
      '0',
      '4',
      '3',
    ]);
    const rows = directChildren(equationArray, 'e');
    expect(rows).toHaveLength(3);
    expect(rows[0]?.textContent).toBe('&x+&&y=1');
    expect(rows[2]?.childElementCount).toBe(0);
  }
  const lowerLimits = descendants(document, 'limLow');
  expect(lowerLimits).toHaveLength(2);
  for (const lowerLimit of lowerLimits) {
    expect(directChildren(lowerLimit).map((child) => child.localName)).toEqual([
      'e',
      'lim',
    ]);
    expect(directChildren(lowerLimit, 'e')[0]?.textContent).toBe('lim');
    expect(directChildren(lowerLimit, 'lim')[0]?.textContent).toBe('x→0');
  }
  const upperLimits = descendants(document, 'limUpp');
  expect(upperLimits).toHaveLength(2);
  for (const upperLimit of upperLimits) {
    expect(directChildren(upperLimit).map((child) => child.localName)).toEqual([
      'e',
      'lim',
    ]);
    expect(directChildren(upperLimit, 'e')[0]?.textContent).toBe('=');
    expect(directChildren(upperLimit, 'lim')[0]?.textContent).toBe('def');
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
      groupCharacter: '\u23de',
      verticalJustification: 'bot',
      container: 'borderBox',
      fractionType: null,
      limit: 'limLow',
      phantomValues: ['1', '0', '1', '0', '1'],
      preSubScript: 'p',
      preSuperScript: '',
      runProperties: {
        children: ['lit', 'scr', 'sty', 'brk'],
        values: ['1', 'double-struck', 'p', '2'],
      },
    },
    {
      document: await xmlEntry(archive, 'word/endnotes.xml'),
      position: 'bot',
      groupCharacter: '\u23df',
      verticalJustification: 'top',
      container: 'box',
      fractionType: 'lin',
      limit: 'limUpp',
      phantomValues: ['0', '1', '0', '1', '0'],
      preSubScript: '',
      preSuperScript: 'q',
      runProperties: {
        children: ['nor', 'scr', 'sty', 'brk', 'aln'],
        values: ['1', 'sans-serif', 'bi', '', '1'],
      },
    },
    {
      document: await matchingXmlEntry(archive, /^word\/header\d*\.xml$/i),
      position: 'top',
      groupCharacter: '\u23de',
      verticalJustification: 'bot',
      container: 'borderBox',
      fractionType: null,
      limit: 'limLow',
      phantomValues: ['1', '0', '1', '0', '1'],
      preSubScript: 'p',
      preSuperScript: '',
      runProperties: {
        children: ['lit', 'scr', 'sty', 'brk'],
        values: ['1', 'double-struck', 'p', '2'],
      },
    },
    {
      document: await matchingXmlEntry(archive, /^word\/footer\d*\.xml$/i),
      position: 'bot',
      groupCharacter: '\u23df',
      verticalJustification: 'top',
      container: 'box',
      fractionType: 'lin',
      limit: 'limUpp',
      phantomValues: ['0', '1', '0', '1', '0'],
      preSubScript: '',
      preSuperScript: 'q',
      runProperties: {
        children: ['nor', 'scr', 'sty', 'brk', 'aln'],
        values: ['1', 'sans-serif', 'bi', '', '1'],
      },
    },
  ];
  for (const story of stories) {
    expect(descendants(story.document, 'oMath')).toHaveLength(1);
    expect(descendants(story.document, 'oMath')[0]?.namespaceURI).toBe(
      MATH_NAMESPACE,
    );
    expect(descendants(story.document, story.container)).toHaveLength(1);
    const alignedScripts = descendants(story.document, 'sSubSup');
    expect(alignedScripts).toHaveLength(1);
    const scriptProperties = directChildren(alignedScripts[0], 'sSubSupPr')[0];
    expect(
      directChildren(alignedScripts[0]).map((child) => child.localName),
    ).toEqual(['sSubSupPr', 'e', 'sub', 'sup']);
    expect(
      mathValueAttribute(directChildren(scriptProperties, 'alnScr')[0]),
    ).toBe('1');
    const radicals = descendants(story.document, 'rad');
    expect(radicals).toHaveLength(1);
    expect(directChildren(radicals[0]).map((child) => child.localName)).toEqual(
      ['radPr', 'deg', 'e'],
    );
    expect(directChildren(radicals[0], 'deg')[0]?.childElementCount).toBe(0);
    expect(
      mathValueAttribute(
        directChildren(directChildren(radicals[0], 'radPr')[0], 'degHide')[0],
      ),
    ).toBe('1');
    const storyNaries = descendants(story.document, 'nary');
    expect(storyNaries).toHaveLength(1);
    const storyNary = storyNaries[0];
    expect(directChildren(storyNary).map((child) => child.localName)).toEqual([
      'naryPr',
      'sub',
      'sup',
      'e',
    ]);
    const storyNaryProperties = directChildren(storyNary, 'naryPr')[0];
    const borderStory = story.container === 'borderBox';
    expect(
      directChildren(storyNaryProperties).map((child) => child.localName),
    ).toEqual(
      borderStory ? ['chr', 'limLoc', 'subHide'] : ['chr', 'limLoc', 'supHide'],
    );
    expect(
      ['chr', 'limLoc'].map((name) =>
        mathValueAttribute(directChildren(storyNaryProperties, name)[0]),
      ),
    ).toEqual(borderStory ? ['\u222b', 'subSup'] : ['\u2211', 'undOvr']);
    expect(directChildren(storyNary, 'sub')[0]?.textContent ?? '').toBe(
      borderStory ? '' : 'k',
    );
    expect(directChildren(storyNary, 'sup')[0]?.textContent ?? '').toBe(
      borderStory ? 'n' : '',
    );
    const fractions = descendants(story.document, 'f');
    expect(fractions).toHaveLength(1);
    expect(
      directChildren(fractions[0]).map((child) => child.localName),
    ).toEqual(
      story.fractionType === null ? ['num', 'den'] : ['fPr', 'num', 'den'],
    );
    expect(
      story.fractionType === null
        ? null
        : mathValueAttribute(
            directChildren(directChildren(fractions[0], 'fPr')[0], 'type')[0],
          ),
    ).toBe(story.fractionType);
    const limit = descendants(story.document, story.limit)[0];
    expect(limit).toBeDefined();
    expect(directChildren(limit).map((child) => child.localName)).toEqual([
      'e',
      'lim',
    ]);
    const equationArray = descendants(story.document, 'eqArr')[0];
    expect(equationArray).toBeDefined();
    const equationArrayProperties = directChildren(equationArray, 'eqArrPr')[0];
    expect(
      directChildren(equationArrayProperties).map(mathValueAttribute),
    ).toEqual(['center', '0', '0', '0', '0']);
    expect(directChildren(equationArray, 'e')).toHaveLength(1);
    const storyPreScripts = descendants(story.document, 'sPre');
    expect(storyPreScripts).toHaveLength(1);
    const preScript = storyPreScripts[0];
    expect(directChildren(preScript).map((child) => child.localName)).toEqual([
      'sub',
      'sup',
      'e',
    ]);
    expect(directChildren(preScript, 'sub')[0]?.textContent ?? '').toBe(
      story.preSubScript,
    );
    expect(directChildren(preScript, 'sup')[0]?.textContent ?? '').toBe(
      story.preSuperScript,
    );
    const styledStoryRuns = descendants(story.document, 'r').filter(
      (run) => run.namespaceURI === MATH_NAMESPACE && mathRunProperties(run),
    );
    expect(styledStoryRuns).toHaveLength(1);
    expect(mathRunProperties(styledStoryRuns[0])).toEqual(story.runProperties);
    const storyPhantoms = descendants(story.document, 'phant');
    expect(storyPhantoms).toHaveLength(1);
    const phantom = storyPhantoms[0];
    expect(phantom).toBeDefined();
    expect(directChildren(phantom).map((child) => child.localName)).toEqual([
      'phantPr',
      'e',
    ]);
    const phantomProperties = directChildren(phantom, 'phantPr')[0];
    expect(
      directChildren(phantomProperties).map((child) => child.localName),
    ).toEqual(['show', 'zeroWid', 'zeroAsc', 'zeroDesc', 'transp']);
    expect(directChildren(phantomProperties).map(mathValueAttribute)).toEqual(
      story.phantomValues,
    );
    const groupCharacter = descendants(story.document, 'groupChr')[0];
    expect(groupCharacter).toBeDefined();
    expect(
      directChildren(groupCharacter).map((child) => child.localName),
    ).toEqual(['groupChrPr', 'e']);
    const groupCharacterProperties = directChildren(
      groupCharacter,
      'groupChrPr',
    )[0];
    expect(
      ['chr', 'pos', 'vertJc'].map((name) =>
        mathValueAttribute(directChildren(groupCharacterProperties, name)[0]),
      ),
    ).toEqual([
      story.groupCharacter,
      story.position,
      story.verticalJustification,
    ]);
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

function mathRunProperties(
  run: Element | undefined,
): { children: string[]; values: Array<string | null> } | null {
  if (!run) return null;
  const properties = directChildren(run, 'rPr').find(
    (candidate) => candidate.namespaceURI === MATH_NAMESPACE,
  );
  if (!properties) return null;
  const children = directChildren(properties);
  return {
    children: children.map((child) => child.localName),
    values: children.map((child) =>
      child.localName === 'brk'
        ? (mathNamedAttribute(child, 'alnAt') ?? '')
        : mathValueAttribute(child),
    ),
  };
}

function mathNamedAttribute(
  element: Element | undefined,
  name: string,
): string | null {
  return (
    Array.from(element?.attributes ?? []).find(
      (attribute) =>
        xmlAttributeNamespace(element as Element, attribute) ===
          MATH_NAMESPACE && xmlAttributeLocalName(attribute) === name,
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
