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
import { patchDocxEquations } from '../src/internal/features/work/work-docx-equation-export';
import {
  xmlAttributeLocalName,
  xmlAttributeNamespace,
  xmlNamespaceUri,
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
const WORD_DATE_UTC_NAMESPACE =
  'http://schemas.microsoft.com/office/word/2023/wordml/word16du';
const WORD_2010_NAMESPACE =
  'http://schemas.microsoft.com/office/word/2010/wordml';
const WORD_2010_GLOW = `<w14:glow xmlns:w14="${WORD_2010_NAMESPACE}" w14:rad="63500"><w14:srgbClr w14:val="FFFF00"/></w14:glow>`;
const WORD_2010_SHADOW = `<w14:shadow xmlns:w14="${WORD_2010_NAMESPACE}"><w14:srgbClr w14:val="000000"/></w14:shadow>`;
const WORD_2010_REFLECTION = `<w14:reflection xmlns:w14="${WORD_2010_NAMESPACE}"/>`;
const WORD_2010_TEXT_OUTLINE = `<w14:textOutline xmlns:w14="${WORD_2010_NAMESPACE}"/>`;
const WORD_2010_TEXT_FILL = `<w14:textFill xmlns:w14="${WORD_2010_NAMESPACE}"/>`;
const WORD_2010_SCENE_3D = `<w14:scene3d xmlns:w14="${WORD_2010_NAMESPACE}"/>`;
const MATH_NAMESPACE =
  'http://schemas.openxmlformats.org/officeDocument/2006/math';
const STRICT_MATH_NAMESPACE = 'http://purl.oclc.org/ooxml/officeDocument/math';
const RELATIONSHIP_NAMESPACE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const MARKUP_COMPATIBILITY_NAMESPACE =
  'http://schemas.openxmlformats.org/markup-compatibility/2006';
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
      const wordStyledRun = element?.querySelector(
        'mtext[mathvariant="bold-fraktur"]',
      );
      expect(wordStyledRun).toHaveAttribute('mathcolor', '#1a2b3c');
      expect(wordStyledRun).toHaveAttribute('mathsize', '12.5pt');
      expect(wordStyledRun).toHaveAttribute('dir', 'ltr');
      expect(wordStyledRun).toHaveAttribute('lang', 'en-US');
      expect(wordStyledRun?.getAttribute('style')).toContain(
        'font-family: "Cambria Math"',
      );
      expect(wordStyledRun?.getAttribute('style')).toContain(
        'font-weight: bold',
      );
      expect(wordStyledRun?.getAttribute('style')).toContain(
        'font-style: normal',
      );
      expect(wordStyledRun?.getAttribute('style')).toContain(
        'text-decoration-line: underline line-through',
      );
      expect(wordStyledRun?.getAttribute('style')).toContain(
        'text-decoration-style: double',
      );
      expect(wordStyledRun?.getAttribute('style')).toContain(
        'text-decoration-color: #abcdef',
      );
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
      expect(matrix?.getAttribute('rowspacing')).toBe('12pt');
      expect(matrix?.getAttribute('columnspacing')).toBe('1.5em');
      expect(matrix?.hasAttribute('columnwidth')).toBe(false);
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
      expect(element?.getAttribute('aria-label')).toContain(
        'matrix(row-spacing=exact:12,column-gap=multiple:3,minimum-column-width=120twip;',
      );
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
        })?.children[0],
      ).toEqual({
        type: 'lowerLimit',
        base: [],
        limit: [{ type: 'run', text: 'x→0' }],
      });
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
        })?.children[0],
      ).toEqual({
        type: 'phantom',
        show: true,
        zeroWidth: false,
        zeroAscent: false,
        zeroDescent: false,
        transparent: false,
        children: [],
      });
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
        })?.children[0],
      ).toEqual({
        type: 'preSubSuperScript',
        base: [],
        subScript: [{ type: 'run', text: 'i' }],
        superScript: [{ type: 'run', text: 'j' }],
      });
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
        })?.children[0],
      ).toEqual({
        type: 'upperLimit',
        base: [{ type: 'run', text: '=' }],
        limit: [],
      });
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
    ).toEqual([37, 37, 1, 1]);
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
    expect(
      descendants(document, 'rPr').some(
        (properties) => properties.namespaceURI === WORD_NAMESPACE,
      ),
    ).toBe(true);
    expect(
      descendants(document, 'rPr').some(
        (properties) => properties.namespaceURI === STRICT_WORD_NAMESPACE,
      ),
    ).toBe(false);
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
    columnSpacing.setAttributeNS(MATH_NAMESPACE, 'm:val', '31681');
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
      `<m:d><m:dPr><m:grow m:val="1"/><m:shp/><m:ctrlPr/></m:dPr><m:e>${run}</m:e></m:d>`,
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
      `<m:bar><m:barPr><m:pos m:val="left"/></m:barPr><m:e>${run}</m:e></m:bar>`,
      `<m:bar><m:barPr><m:pos m:val="bottom"/></m:barPr><m:e>${run}</m:e></m:bar>`,
      `<m:bar><m:barPr><m:grow m:val="1"/></m:barPr><m:e>${run}</m:e></m:bar>`,
      `<m:bar><m:e>${run}</m:e><m:barPr><m:pos m:val="top"/></m:barPr></m:bar>`,
      `<m:bar><m:barPr><m:ctrlPr/><m:pos m:val="top"/></m:barPr><m:e>${run}</m:e></m:bar>`,
      `<m:bar><m:barPr><m:pos m:val="top"/><m:pos m:val="bot"/></m:barPr><m:e>${run}</m:e></m:bar>`,
      `<m:bar><m:barPr><m:pos m:val="top" m:extra="semantic"/></m:barPr><m:e>${run}</m:e></m:bar>`,
      `<m:bar><v:barPr xmlns:v="${VENDOR_NAMESPACE}"><m:pos m:val="top"/></v:barPr><m:e>${run}</m:e></m:bar>`,
      `<m:bar><m:barPr><m:pos xmlns:r="${RELATIONSHIP_NAMESPACE}" m:val="top" r:id="rIdUnsafe"/></m:barPr><m:e>${run}</m:e></m:bar>`,
      `<m:bar><m:e>${run}</m:e><m:e>${run}</m:e></m:bar>`,
      `<m:borderBox><m:borderBoxPr><m:grow/></m:borderBoxPr><m:e>${run}</m:e></m:borderBox>`,
      `<m:borderBox><m:borderBoxPr><m:strikeH/><m:hideTop/></m:borderBoxPr><m:e>${run}</m:e></m:borderBox>`,
      `<m:borderBox><m:borderBoxPr><m:hideTop/><m:hideTop/></m:borderBoxPr><m:e>${run}</m:e></m:borderBox>`,
      `<m:borderBox><m:borderBoxPr><m:hideTop m:val="maybe"/></m:borderBoxPr><m:e>${run}</m:e></m:borderBox>`,
      `<m:borderBox><v:borderBoxPr xmlns:v="${VENDOR_NAMESPACE}"><m:hideTop/></v:borderBoxPr><m:e>${run}</m:e></m:borderBox>`,
      `<m:borderBox><m:borderBoxPr><m:hideTop xmlns:r="${RELATIONSHIP_NAMESPACE}" r:id="rIdUnsafe"/></m:borderBoxPr><m:e>${run}</m:e></m:borderBox>`,
      `<m:borderBox><m:e>${run}</m:e><m:borderBoxPr><m:hideTop/></m:borderBoxPr></m:borderBox>`,
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
      `<m:box><m:e>${run}</m:e><m:e>${run}</m:e></m:box>`,
      `<m:f><m:fPr><m:m><m:mr><m:e>${run}</m:e></m:mr></m:m></m:fPr><m:num>${run}</m:num><m:den>${run}</m:den></m:f>`,
      `<m:f><m:fPr><m:type m:val="bar" m:extra="semantic"/></m:fPr><m:num>${run}</m:num><m:den>${run}</m:den></m:f>`,
      `<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg>${run}</m:deg><m:e>${run}</m:e></m:rad>`,
      `<m:rad><m:radPr><m:ctrlPr/><m:degHide/></m:radPr><m:e>${run}</m:e></m:rad>`,
      `<m:nary><m:naryPr><m:chr m:val="&#x2211;"/><m:subHide m:val="on"/></m:naryPr><m:sub>${run}</m:sub><m:e>${run}</m:e></m:nary>`,
      `<m:nary><m:naryPr><m:chr m:val="&#x2211;"/></m:naryPr><m:e>${run}</m:e></m:nary>`,
      `<m:m><m:mPr><m:cSp m:val="31681"/></m:mPr><m:mr><m:e>${run}</m:e></m:mr></m:m>`,
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

  test('preserves bounded OMML matrix spacing and projects safe gaps', async () => {
    const equation: WorkDocumentEquation = {
      version: 1,
      display: 'inline',
      children: [
        {
          type: 'matrix',
          baseAlignment: 'top',
          placeholdersHidden: false,
          columnAlignments: ['left', 'right'],
          spacing: {
            rowSpacingRule: 'exact',
            rowSpacing: 12,
            columnGapRule: 'multiple',
            columnGap: 3,
            minimumColumnWidthTwips: 120,
          },
          rows: [
            [[{ type: 'run', text: 'a' }], [{ type: 'run', text: 'b' }]],
            [[{ type: 'run', text: 'c' }], [{ type: 'run', text: 'd' }]],
          ],
        },
      ],
    };
    expect(normalizeDocumentEquation(equation)).toEqual(equation);

    const matrixModel = (spacing: unknown) =>
      ({
        version: 1,
        display: 'inline',
        children: [
          {
            type: 'matrix',
            baseAlignment: 'center',
            placeholdersHidden: false,
            columnAlignments: ['center'],
            spacing,
            rows: [[[{ type: 'run', text: 'x' }]]],
          },
        ],
      }) as unknown as WorkDocumentEquation;
    const maximumSpacing = {
      rowSpacingRule: 'multiple',
      rowSpacing: 65_535,
      columnGapRule: 'exact',
      columnGap: 65_535,
      minimumColumnWidthTwips: 31_680,
    };
    expect(normalizeDocumentEquation(matrixModel(maximumSpacing))).toEqual(
      matrixModel(maximumSpacing),
    );
    const invalidSpacing = [
      null,
      {},
      { ...maximumSpacing, extra: true },
      { ...maximumSpacing, rowSpacingRule: 'atLeast' },
      { ...maximumSpacing, columnGapRule: 'atLeast' },
      { ...maximumSpacing, rowSpacing: -1 },
      { ...maximumSpacing, rowSpacing: 65_536 },
      { ...maximumSpacing, rowSpacing: 1.5 },
      { ...maximumSpacing, rowSpacing: '12' },
      { ...maximumSpacing, columnGap: -1 },
      { ...maximumSpacing, columnGap: 65_536 },
      { ...maximumSpacing, columnGap: Number.NaN },
      { ...maximumSpacing, minimumColumnWidthTwips: -1 },
      { ...maximumSpacing, minimumColumnWidthTwips: 31_681 },
      { ...maximumSpacing, minimumColumnWidthTwips: 1.5 },
    ];
    expect(
      invalidSpacing.map((spacing) =>
        normalizeDocumentEquation(matrixModel(spacing)),
      ),
    ).toEqual(invalidSpacing.map(() => null));

    const run = '<m:r><m:t>x</m:t></m:r>';
    const matrixBody = (properties: string) =>
      `<m:m><m:mPr xmlns:r="${RELATIONSHIP_NAMESPACE}" xmlns:v="${VENDOR_NAMESPACE}">${properties}</m:mPr><m:mr><m:e>${run}</m:e><m:e>${run}</m:e></m:mr></m:m>`;
    const completeProperties =
      '<m:rSpRule m:val="+0003"/><m:cGpRule m:val="4"/><m:rSp m:val="+00012"/><m:cSp m:val="120"/><m:cGp m:val="3"/>';
    expect(
      inspectEquationModel(matrixBody(completeProperties))?.children[0],
    ).toEqual({
      type: 'matrix',
      baseAlignment: 'center',
      placeholdersHidden: false,
      columnAlignments: ['center', 'center'],
      spacing: {
        rowSpacingRule: 'exact',
        rowSpacing: 12,
        columnGapRule: 'multiple',
        columnGap: 3,
        minimumColumnWidthTwips: 120,
      },
      rows: [[[{ type: 'run', text: 'x' }], [{ type: 'run', text: 'x' }]]],
    });
    expect(
      inspectEquationModel(matrixBody('<m:cSp/>'))?.children[0],
    ).toMatchObject({
      spacing: {
        rowSpacingRule: 'single',
        rowSpacing: 0,
        columnGapRule: 'single',
        columnGap: 0,
        minimumColumnWidthTwips: 0,
      },
    });
    const spacingRules = [
      'single',
      'oneAndHalf',
      'double',
      'exact',
      'multiple',
    ] as const;
    for (const [index, rowSpacingRule] of spacingRules.entries()) {
      const columnGapRule = spacingRules[4 - index];
      expect(
        inspectEquationModel(
          matrixBody(
            `<m:rSpRule m:val="${index}"/><m:cGpRule m:val="${4 - index}"/>`,
          ),
        )?.children[0],
      ).toMatchObject({
        spacing: { rowSpacingRule, columnGapRule },
      });
    }

    const malformedProperties = [
      '<m:rSpRule m:val="5"/>',
      '<m:cGpRule m:val="-1"/>',
      '<m:rSp m:val="65536"/>',
      '<m:rSp m:val="1.5"/>',
      '<m:cSp m:val="31681"/>',
      '<m:cGp m:val="65536"/>',
      '<m:cGpRule m:val="4"/><m:rSpRule m:val="3"/>',
      '<m:cSp m:val="120"/><m:rSp m:val="12"/>',
      '<m:rSpRule/><m:rSpRule/>',
      '<m:cGp/><m:cGp/>',
      '<m:cSp val="120"/>',
      '<m:cSp v:val="120"/>',
      '<m:cSp m:val="120" m:extra="semantic"/>',
      '<m:cSp m:val="120">meaningful</m:cSp>',
      `<m:cSp r:id="rIdUnsafe"/>`,
      '<v:cSp v:val="120"/>',
    ];
    expect(
      malformedProperties.map(matrixBody).map(inspectEquationBody),
    ).toEqual(malformedProperties.map(() => 'unsupported'));

    const strictSource = `<m:oMath xmlns:m="${STRICT_MATH_NAMESPACE}"><m:m><m:mPr><m:rSpRule m:val="2"/><m:cGpRule m:val="3"/><m:rSp m:val="65535"/><m:cSp m:val="31680"/><m:cGp m:val="400"/></m:mPr><m:mr><m:e><m:r><m:t>x</m:t></m:r></m:e></m:mr></m:m></m:oMath>`;
    expect(inspectEquationRoot(strictSource)).toMatchObject({
      status: 'supported',
      equation: {
        children: [
          {
            spacing: {
              rowSpacingRule: 'double',
              rowSpacing: 65_535,
              columnGapRule: 'exact',
              columnGap: 400,
              minimumColumnWidthTwips: 31_680,
            },
          },
        ],
      },
    });

    const previewDocument = new DOMParser().parseFromString('', 'text/html');
    const preview = createDocumentEquationElement(previewDocument, equation);
    const table = preview.querySelector('mtable');
    expect(table?.getAttribute('rowspacing')).toBe('12pt');
    expect(table?.getAttribute('columnspacing')).toBe('1.5em');
    expect(table?.hasAttribute('columnwidth')).toBe(false);
    expect(preview.getAttribute('aria-label')).toContain(
      'minimum-column-width=120twip',
    );

    const expectNativeMatrixSpacing = async (blob: Blob) => {
      const archive = await JSZip.loadAsync(await blob.arrayBuffer());
      const document = await xmlEntry(archive, 'word/document.xml');
      const matrix = descendants(document, 'm').find(
        (candidate) => candidate.namespaceURI === MATH_NAMESPACE,
      );
      expect(matrix).toBeDefined();
      const properties = directChildren(matrix as Element, 'mPr')[0];
      expect(
        directChildren(properties).map((child) => child.localName),
      ).toEqual([
        'baseJc',
        'plcHide',
        'rSpRule',
        'cGpRule',
        'rSp',
        'cSp',
        'cGp',
        'mcs',
      ]);
      expect(
        directChildren(properties).slice(0, 7).map(mathValueAttribute),
      ).toEqual(['top', '0', '3', '4', '12', '120', '3']);
    };

    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    artifact.content.html = `<p>${equationHtml(equation)}</p>`;
    const first = await createArtifactBlob(artifact);
    await expectNativeMatrixSpacing(first);
    const imported = await importOfficeFile(
      new File([first], 'matrix-spacing.docx', { type: first.type }),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const importedDocument = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    const importedEquation = importedDocument.body.querySelector<HTMLElement>(
      '[data-document-equation]',
    );
    expect(
      documentEquationFromElement(importedEquation as HTMLElement),
    ).toEqual(equation);
    expect(imported.compatibility.issues).not.toContainEqual(
      expect.objectContaining({ code: 'docx.equations.unsupported' }),
    );
    await expectNativeMatrixSpacing(await createArtifactBlob(imported));
  });

  test('preserves OMML n-ary growth and delimiter sizing semantics', async () => {
    const run = (text: string) => ({ type: 'run' as const, text });
    const equation: WorkDocumentEquation = {
      version: 1,
      display: 'inline',
      children: [
        {
          type: 'nary',
          operator: '\u2211',
          limitLocation: 'underOver',
          grow: true,
          subScript: [run('i=1')],
          superScript: [run('n')],
          children: [run('x')],
        },
        {
          type: 'nary',
          operator: '\u222b',
          limitLocation: 'subSup',
          children: [run('f(x)')],
        },
        {
          type: 'delimiter',
          opening: '(',
          closing: ')',
          separator: '|',
          grow: false,
          shape: 'match',
          arguments: [[run('a')], [run('b')]],
        },
        {
          type: 'delimiter',
          opening: '{',
          closing: '}',
          separator: ';',
          shape: 'match',
          arguments: [[run('c')], [run('d')]],
        },
      ],
    };
    expect(normalizeDocumentEquation(equation)).toEqual(equation);

    const canonicalDefaults = normalizeDocumentEquation({
      version: 1,
      display: 'inline',
      children: [
        { ...equation.children[0], grow: false },
        {
          ...equation.children[3],
          grow: true,
          shape: 'centered',
        },
      ],
    });
    expect(canonicalDefaults?.children[0]).not.toHaveProperty('grow');
    expect(canonicalDefaults?.children[1]).not.toHaveProperty('grow');
    expect(canonicalDefaults?.children[1]).not.toHaveProperty('shape');
    const invalidModels = [
      {
        ...equation,
        children: [{ ...equation.children[0], grow: 'true' }],
      },
      {
        ...equation,
        children: [{ ...equation.children[2], grow: 0 }],
      },
      {
        ...equation,
        children: [{ ...equation.children[3], shape: 'round' }],
      },
    ] as unknown as WorkDocumentEquation[];
    expect(invalidModels.map(normalizeDocumentEquation)).toEqual([
      null,
      null,
      null,
    ]);

    const strictSource = `<m:oMath xmlns:m="${STRICT_MATH_NAMESPACE}"><m:nary><m:naryPr><m:chr m:val="&#x2211;"/><m:limLoc m:val="undOvr"/><m:grow/></m:naryPr><m:sub><m:r><m:t>i</m:t></m:r></m:sub><m:sup><m:r><m:t>n</m:t></m:r></m:sup><m:e><m:r><m:t>x</m:t></m:r></m:e></m:nary><m:d><m:dPr><m:begChr m:val="["/><m:sepChr m:val=";"/><m:endChr m:val="]"/><m:grow m:val="0"/><m:shp m:val="match"/></m:dPr><m:e><m:r><m:t>y</m:t></m:r></m:e></m:d></m:oMath>`;
    expect(inspectEquationRoot(strictSource)).toMatchObject({
      status: 'supported',
      equation: {
        children: [
          { type: 'nary', grow: true },
          { type: 'delimiter', grow: false, shape: 'match' },
        ],
      },
    });

    const previewDocument = new DOMParser().parseFromString('', 'text/html');
    const preview = createDocumentEquationElement(previewDocument, equation);
    const operator = Array.from(preview.querySelectorAll('mo')).find(
      (candidate) => candidate.textContent === '\u2211',
    );
    const integral = Array.from(preview.querySelectorAll('mo')).find(
      (candidate) => candidate.textContent === '\u222b',
    );
    const fixedOpening = Array.from(preview.querySelectorAll('mo')).find(
      (candidate) => candidate.textContent === '(',
    );
    const matchedOpening = Array.from(preview.querySelectorAll('mo')).find(
      (candidate) => candidate.textContent === '{',
    );
    expect(operator?.getAttribute('stretchy')).toBe('true');
    expect(integral?.getAttribute('stretchy')).toBe('false');
    expect(fixedOpening?.getAttribute('stretchy')).toBe('false');
    expect(fixedOpening?.hasAttribute('symmetric')).toBe(false);
    expect(matchedOpening?.getAttribute('stretchy')).toBe('true');
    expect(matchedOpening?.getAttribute('symmetric')).toBe('false');
    expect(preview.getAttribute('aria-label')).toContain('nary(grow;');
    expect(preview.getAttribute('aria-label')).toContain(
      'delimiter(grow=false,shape=match;',
    );

    const expectNativeSizing = async (blob: Blob) => {
      const archive = await JSZip.loadAsync(await blob.arrayBuffer());
      const document = await xmlEntry(archive, 'word/document.xml');
      const naries = descendants(document, 'nary');
      expect(naries).toHaveLength(2);
      expect(
        directChildren(directChildren(naries[0], 'naryPr')[0]).map(
          (child) => child.localName,
        ),
      ).toEqual(['chr', 'limLoc', 'grow']);
      expect(
        directChildren(directChildren(naries[0], 'naryPr')[0]).map(
          mathValueAttribute,
        ),
      ).toEqual(['\u2211', 'undOvr', '1']);
      expect(
        directChildren(directChildren(naries[1], 'naryPr')[0]).map(
          (child) => child.localName,
        ),
      ).toEqual(['chr', 'limLoc', 'subHide', 'supHide']);

      const delimiters = descendants(document, 'd');
      expect(delimiters).toHaveLength(2);
      expect(
        directChildren(directChildren(delimiters[0], 'dPr')[0]).map(
          (child) => child.localName,
        ),
      ).toEqual(['begChr', 'sepChr', 'endChr', 'grow', 'shp']);
      expect(
        directChildren(directChildren(delimiters[0], 'dPr')[0]).map(
          mathValueAttribute,
        ),
      ).toEqual(['(', '|', ')', '0', 'match']);
      expect(
        directChildren(directChildren(delimiters[1], 'dPr')[0]).map(
          (child) => child.localName,
        ),
      ).toEqual(['begChr', 'sepChr', 'endChr', 'shp']);
      expect(
        directChildren(directChildren(delimiters[1], 'dPr')[0]).map(
          mathValueAttribute,
        ),
      ).toEqual(['{', ';', '}', 'match']);
    };

    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    artifact.content.html = `<p>${equationHtml(equation)}</p>`;
    const first = await createArtifactBlob(artifact);
    await expectNativeSizing(first);
    const imported = await importOfficeFile(
      new File([first], 'operator-sizing.docx', { type: first.type }),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const importedDocument = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    const importedEquation = importedDocument.body.querySelector<HTMLElement>(
      '[data-document-equation]',
    );
    expect(
      documentEquationFromElement(importedEquation as HTMLElement),
    ).toEqual(equation);
    expect(imported.compatibility.issues).not.toContainEqual(
      expect.objectContaining({ code: 'docx.equations.unsupported' }),
    );
    await expectNativeSizing(await createArtifactBlob(imported));
  });

  test('preserves bounded Word run properties inside math runs', () => {
    const wordProperties = [
      '<w:rFonts w:ascii=" Cambria Math " w:hAnsi="Cambria Math" w:eastAsia="等线" w:cs="Arial" w:asciiTheme="majorAscii" w:hAnsiTheme="majorHAnsi" w:eastAsiaTheme="minorEastAsia" w:cstheme="majorBidi" w:hint="eastAsia"/>',
      '<w:b w:val="on"/>',
      '<w:bCs w:val="0"/>',
      '<w:i w:val="false"/>',
      '<w:iCs/>',
      '<w:caps w:val="0"/>',
      '<w:smallCaps/>',
      '<w:strike w:val="off"/>',
      '<w:dstrike w:val="true"/>',
      '<w:outline/>',
      '<w:shadow/>',
      '<w:emboss w:val="0"/>',
      '<w:imprint w:val="false"/>',
      '<w:noProof/>',
      '<w:snapToGrid w:val="0"/>',
      '<w:vanish w:val="0"/>',
      '<w:webHidden/>',
      '<w:color w:val="1A2B3C" w:themeColor="accent2" w:themeTint="80" w:themeShade="40"/>',
      '<w:spacing w:val="20"/>',
      '<w:w w:val="90"/>',
      '<w:kern w:val="22"/>',
      '<w:position w:val="2"/>',
      '<w:sz w:val="25"/>',
      '<w:szCs w:val="28"/>',
      '<w:u w:val="wavyDouble" w:color="ABCDEF" w:themeColor="accent3" w:themeTint="20"/>',
      '<w:effect w:val="shimmer"/>',
      '<w:bdr w:val="double" w:color="112233" w:themeColor="accent1" w:themeTint="66" w:themeShade="33" w:sz="24" w:space="2" w:shadow="1" w:frame="0"/>',
      '<w:fitText w:val="720" w:id="50"/>',
      '<w:vertAlign w:val="baseline"/>',
      '<w:rtl w:val="false"/>',
      '<w:cs w:val="0"/>',
      '<w:em w:val="circle"/>',
      '<w:lang w:val="en-US" w:eastAsia="zh-CN" w:bidi="ar-SA"/>',
      '<w:eastAsianLayout w:id="9" w:combine="1" w:combineBrackets="curly" w:vert="0" w:vertCompress="1"/>',
      '<w:specVanish w:val="0"/>',
      '<w14:glow w14:rad="63500"><w14:schemeClr w14:val="accent4"><w14:alpha w14:val="50000"/></w14:schemeClr></w14:glow>',
      '<w14:shadow w14:blurRad="25400" w14:dist="12700" w14:dir="5400000" w14:sx="90000" w14:sy="-100000" w14:kx="-60000" w14:ky="120000" w14:algn="tr"><w14:schemeClr w14:val="accent5"><w14:lumMod w14:val="75000"/></w14:schemeClr></w14:shadow>',
      '<w14:reflection w14:blurRad="12700" w14:stA="60000" w14:stPos="0" w14:endA="10000" w14:endPos="100000" w14:dist="6350" w14:dir="5400000" w14:fadeDir="16200000" w14:sx="100000" w14:sy="-50000" w14:kx="-60000" w14:ky="120000" w14:algn="b"/>',
      '<w14:textOutline w14:w="12700" w14:cap="sq" w14:cmpd="dbl" w14:algn="in"><w14:solidFill><w14:schemeClr w14:val="accent6"><w14:alpha w14:val="80000"/></w14:schemeClr></w14:solidFill><w14:prstDash w14:val="dashDot"/><w14:miter w14:lim="400000"/></w14:textOutline>',
      '<w14:textFill><w14:solidFill><w14:srgbClr w14:val="2468AC"><w14:shade w14:val="25000"/></w14:srgbClr></w14:solidFill></w14:textFill>',
    ].join('');
    const richRun = (mathNamespace: string, wordNamespace: string) =>
      `<m:oMath xmlns:m="${mathNamespace}" xmlns:w="${wordNamespace}" xmlns:w14="${WORD_2010_NAMESPACE}"><m:r><m:rPr><m:lit m:val="1"/><m:scr m:val="fraktur"/><m:sty m:val="b"/><m:brk m:alnAt="3"/><m:aln m:val="1"/></m:rPr><w:rPr>${wordProperties}</w:rPr><m:t>styledF</m:t></m:r></m:oMath>`;
    for (const [mathNamespace, wordNamespace] of [
      [MATH_NAMESPACE, WORD_NAMESPACE],
      [STRICT_MATH_NAMESPACE, STRICT_WORD_NAMESPACE],
    ]) {
      const inspection = inspectEquationRoot(
        richRun(mathNamespace, wordNamespace),
      );
      expect(inspection.status).toBe('supported');
      expect(inspection.text).toBe('styledF');
      expect(inspection.equation?.children[0]).toEqual({
        type: 'run',
        text: 'styledF',
        literal: true,
        script: 'fraktur',
        style: 'bold',
        manualBreak: { alignmentAt: 3 },
        alignment: true,
        wordRunProperties: richWordRunProperties(),
      });
    }

    expect(
      inspectEquationModel(
        `<m:r xmlns:w="${WORD_NAMESPACE}"><w:rPr/><w:t>x</w:t></m:r>`,
      ),
    ).toEqual(simpleEquation('x'));
    expect(
      inspectEquation(
        `<m:r xmlns:w="${WORD_NAMESPACE}"><m:t xml:space="preserve"> x </m:t></m:r>`,
      ).text,
    ).toBe(' x ');

    const equationWithWordRunProperties = (wordRunProperties: unknown) => ({
      version: 1,
      display: 'inline',
      children: [{ type: 'run', text: 'x', wordRunProperties }],
    });
    expect(
      normalizeDocumentEquation(
        equationWithWordRunProperties({ fonts: {}, languages: {} }),
      ),
    ).toEqual(simpleEquation('x'));
    expect(
      normalizeDocumentEquation(
        equationWithWordRunProperties({
          ...richWordRunProperties(),
          fonts: {
            ...richWordRunProperties().fonts,
            ascii: ' Cambria Math ',
          },
          color: {
            ...richWordRunProperties().color,
            value: '#1A2B3C',
          },
          underline: {
            ...richWordRunProperties().underline,
            color: {
              ...richWordRunProperties().underline.color,
              value: '#ABCDEF',
            },
          },
        }),
      )?.children[0],
    ).toEqual({
      type: 'run',
      text: 'x',
      wordRunProperties: richWordRunProperties(),
    });
    const invalidModels = [
      { extra: true },
      { fonts: { extra: 'Cambria Math' } },
      { fonts: { ascii: '' } },
      { fonts: { ascii: 'x'.repeat(128) } },
      { fonts: { ascii: 'Cambria\nMath' } },
      { fonts: { asciiTheme: 'majorLatin' } },
      { fonts: { hint: 'latin' } },
      { bold: 'true' },
      { color: {} },
      { color: { value: '#12345' } },
      { color: { theme: 'none' } },
      { color: { theme: 'none', tint: '80' } },
      { color: { theme: 'accent1', tint: '0' } },
      { fontSize: 0 },
      { fontSize: 1.25 },
      { fontSize: 513 },
      { underline: { style: 'triple' } },
      { underline: { style: 'single', color: {} } },
      { languages: { latin: 'en_US' } },
      { languages: { latin: 'a'.repeat(86) } },
      { languages: { extra: 'en-US' } },
    ];
    expect(
      invalidModels.map((properties) =>
        normalizeDocumentEquation(equationWithWordRunProperties(properties)),
      ),
    ).toEqual(invalidModels.map(() => null));

    const wordRun = (properties: string, text = '<m:t>x</m:t>') =>
      `<m:r xmlns:w="${WORD_NAMESPACE}">${properties}${text}</m:r>`;
    const unsupported = [
      wordRun('<m:t>x</m:t><w:rPr/>', ''),
      wordRun('<w:rPr/><m:rPr/>'),
      wordRun('<w:rPr/><w:rPr/>'),
      wordRun('<w:rPr w:val="semantic"/>'),
      wordRun('<w:rPr>meaningful</w:rPr>'),
      wordRun(`<w:rPr>${WORD_2010_SCENE_3D}</w:rPr>`),
      wordRun('<w:rPr><w:b/><w:rFonts w:ascii="Cambria Math"/></w:rPr>'),
      wordRun('<w:rPr><w:b/><w:b/></w:rPr>'),
      wordRun('<w:rPr><w:b w:val="maybe"/></w:rPr>'),
      wordRun('<w:rPr><w:b><w:i/></w:b></w:rPr>'),
      wordRun(
        `<w:rPr xmlns:r="${RELATIONSHIP_NAMESPACE}"><w:b r:id="rIdUnsafe"/></w:rPr>`,
      ),
      wordRun('<w:rPr><w:rFonts ascii="Cambria Math"/></w:rPr>'),
      wordRun('<w:rPr><w:rFonts w:ascii=""/></w:rPr>'),
      wordRun(`<w:rPr><w:rFonts w:ascii="${'x'.repeat(128)}"/></w:rPr>`),
      wordRun('<w:rPr><w:rFonts w:asciiTheme="majorLatin"/></w:rPr>'),
      wordRun('<w:rPr><w:color/></w:rPr>'),
      wordRun('<w:rPr><w:color w:val="12345"/></w:rPr>'),
      wordRun('<w:rPr><w:color w:themeColor="none"/></w:rPr>'),
      wordRun('<w:rPr><w:color w:themeTint="80"/></w:rPr>'),
      wordRun('<w:rPr><w:color w:themeColor="none" w:themeTint="80"/></w:rPr>'),
      wordRun('<w:rPr><w:sz w:val="0"/></w:rPr>'),
      wordRun('<w:rPr><w:sz w:val="1025"/></w:rPr>'),
      wordRun('<w:rPr><w:sz w:val="25.5"/></w:rPr>'),
      wordRun('<w:rPr><w:u/></w:rPr>'),
      wordRun('<w:rPr><w:u w:val="triple"/></w:rPr>'),
      wordRun('<w:rPr><w:u w:val="single" w:color="12345"/></w:rPr>'),
      wordRun('<w:rPr><w:lang w:val="en_US"/></w:rPr>'),
      wordRun(`<w:rPr><w:lang w:val="${'a'.repeat(86)}"/></w:rPr>`),
      wordRun(`<v:rPr xmlns:v="${VENDOR_NAMESPACE}"><w:b/></v:rPr>`),
      wordRun('<w:rPr/>', `<w:t w:extra="semantic">x</w:t>`),
      wordRun('<w:rPr/>', '<w:t>x</w:t><m:t>y</m:t>'),
    ];
    expect(unsupported.map(inspectEquationBody)).toEqual(
      unsupported.map(() => 'unsupported'),
    );
    expect(
      inspectEquation(
        wordRun(`<w:rPr>${WORD_2010_SCENE_3D}</w:rPr>`, '<w:t>fallback</w:t>'),
      ),
    ).toMatchObject({ status: 'unsupported', text: 'fallback' });
  });

  test('preserves Word casing, relief, and visibility inside OMML', async () => {
    const equation = {
      version: 1,
      display: 'inline',
      children: [
        {
          type: 'run',
          text: 'all-caps-outline',
          wordRunProperties: {
            allCaps: true,
            outline: true,
            shadow: true,
            hidden: false,
            webHidden: false,
          },
        },
        {
          type: 'run',
          text: 'small-caps-hidden',
          wordRunProperties: {
            allCaps: false,
            smallCaps: true,
            outline: false,
            shadow: false,
            emboss: true,
            imprint: false,
            hidden: true,
            webHidden: true,
          },
        },
        {
          type: 'run',
          text: 'explicit-effect-resets',
          wordRunProperties: {
            allCaps: false,
            smallCaps: false,
            outline: false,
            shadow: false,
            emboss: false,
            imprint: false,
            hidden: false,
            webHidden: false,
          },
        },
        {
          type: 'nary',
          operator: '\u2211',
          limitLocation: 'underOver',
          controlProperties: { smallCaps: true, shadow: true },
          children: [{ type: 'run', text: 'operator-effects' }],
        },
      ],
    } as unknown as WorkDocumentEquation;
    expect(normalizeDocumentEquation(equation)).toEqual(equation);

    const equationWithWordRunProperties = (wordRunProperties: unknown) =>
      ({
        version: 1,
        display: 'inline',
        children: [{ type: 'run', text: 'x', wordRunProperties }],
      }) as unknown as WorkDocumentEquation;
    for (const properties of [
      { allCaps: false, smallCaps: true },
      { outline: true, shadow: true, emboss: false, imprint: false },
      { hidden: false, webHidden: true },
    ]) {
      expect(
        normalizeDocumentEquation(equationWithWordRunProperties(properties))
          ?.children[0],
      ).toMatchObject({ wordRunProperties: properties });
    }
    const invalidModels = [
      { allCaps: 'true' },
      { smallCaps: 1 },
      { outline: null },
      { shadow: 'false' },
      { emboss: 0 },
      { imprint: 'off' },
      { hidden: 'true' },
      { webHidden: 1 },
      { allCaps: true, smallCaps: true },
      { strike: true, doubleStrike: true },
      { outline: true, emboss: true },
      { outline: true, imprint: true },
      { shadow: true, emboss: true },
      { shadow: true, imprint: true },
      { emboss: true, imprint: true },
    ];
    expect(
      invalidModels.map((properties) =>
        normalizeDocumentEquation(equationWithWordRunProperties(properties)),
      ),
    ).toEqual(invalidModels.map(() => null));

    const wordRun = (properties: string, namespace = WORD_NAMESPACE) =>
      `<m:r xmlns:w="${namespace}"><w:rPr>${properties}</w:rPr><m:t>x</m:t></m:r>`;
    const effectProperties =
      '<w:caps w:val="0"/><w:smallCaps/><w:outline w:val="false"/><w:shadow/><w:emboss w:val="0"/><w:imprint w:val="off"/><w:vanish/><w:webHidden w:val="0"/>';
    expect(
      inspectEquationModel(wordRun(effectProperties))?.children[0],
    ).toEqual({
      type: 'run',
      text: 'x',
      wordRunProperties: {
        allCaps: false,
        smallCaps: true,
        outline: false,
        shadow: true,
        emboss: false,
        imprint: false,
        hidden: true,
        webHidden: false,
      },
    });
    expect(
      inspectEquationRoot(
        `<m:oMath xmlns:m="${STRICT_MATH_NAMESPACE}" xmlns:w="${STRICT_WORD_NAMESPACE}"><m:r><w:rPr>${effectProperties}</w:rPr><m:t>strict-effects</m:t></m:r></m:oMath>`,
      ),
    ).toMatchObject({
      status: 'supported',
      equation: {
        children: [
          {
            wordRunProperties: {
              allCaps: false,
              smallCaps: true,
              outline: false,
              shadow: true,
              emboss: false,
              imprint: false,
              hidden: true,
              webHidden: false,
            },
          },
        ],
      },
    });

    const unsupported = [
      wordRun('<w:caps w:val="maybe"/>'),
      wordRun('<w:outline w:val="2"/>'),
      wordRun('<w:vanish w:val="hidden"/>'),
      wordRun('<w:caps/><w:smallCaps/>'),
      wordRun('<w:strike/><w:dstrike/>'),
      wordRun('<w:outline/><w:emboss/>'),
      wordRun('<w:outline/><w:imprint/>'),
      wordRun('<w:shadow/><w:emboss/>'),
      wordRun('<w:shadow/><w:imprint/>'),
      wordRun('<w:emboss/><w:imprint/>'),
      wordRun('<w:smallCaps/><w:caps w:val="0"/>'),
      wordRun('<w:emboss w:val="0"/><w:shadow/>'),
      wordRun('<w:webHidden/><w:vanish w:val="0"/>'),
      wordRun('<w:caps/><w:caps w:val="0"/>'),
      wordRun('<w:caps val="1"/>'),
      wordRun('<w:outline w:val="1" w:extra="semantic"/>'),
      wordRun(
        `<w:shadow xmlns:r="${RELATIONSHIP_NAMESPACE}" r:id="rIdUnsafe"/>`,
      ),
      wordRun('<w:emboss><w:b/></w:emboss>'),
      wordRun(`<v:caps xmlns:v="${VENDOR_NAMESPACE}" v:val="1"/>`),
      wordRun('<m:vanish m:val="1"/>'),
    ];
    expect(unsupported.map(inspectEquationBody)).toEqual(
      unsupported.map(() => 'unsupported'),
    );

    const document = new DOMParser().parseFromString('', 'text/html');
    const preview = createDocumentEquationElement(document, equation);
    const styleFor = (text: string) =>
      Array.from(preview.querySelectorAll('mtext, mo'))
        .find((element) => element.textContent === text)
        ?.getAttribute('style');
    expect(styleFor('all-caps-outline')).toContain('text-transform:uppercase');
    expect(styleFor('all-caps-outline')).not.toMatch(
      /text-shadow|text-stroke|visibility|display/iu,
    );
    expect(styleFor('small-caps-hidden')).toContain('text-transform:none');
    expect(styleFor('small-caps-hidden')).toContain(
      'font-variant-caps:small-caps',
    );
    expect(styleFor('small-caps-hidden')).not.toMatch(/visibility|display/iu);
    expect(styleFor('explicit-effect-resets')).toContain('text-transform:none');
    expect(styleFor('explicit-effect-resets')).toContain(
      'font-variant-caps:normal',
    );
    expect(styleFor('\u2211')).toContain('font-variant-caps:small-caps');
    expect(preview.textContent).toContain('small-caps-hidden');
    const sanitized = new DOMParser().parseFromString(
      sanitizeDocumentPageChromeHtml(preview.outerHTML),
      'text/html',
    );
    expect(
      Array.from(sanitized.querySelectorAll('mtext'))
        .find((element) => element.textContent === 'all-caps-outline')
        ?.getAttribute('style'),
    ).toContain('text-transform:uppercase');
    expect(
      Array.from(sanitized.querySelectorAll('mtext'))
        .find((element) => element.textContent === 'small-caps-hidden')
        ?.getAttribute('style'),
    ).toContain('font-variant-caps:small-caps');

    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    artifact.content.html = `<p>${preview.outerHTML}</p>`;
    const first = await createArtifactBlob(artifact);
    await expectNativeWordRunEffects(first);
    const imported = await importOfficeFile(
      new File([first], 'word-run-effects.docx', { type: first.type }),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const importedDocument = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    const importedEquation = importedDocument.body.querySelector<HTMLElement>(
      '[data-document-equation]',
    );
    expect(
      documentEquationFromElement(importedEquation as HTMLElement),
    ).toEqual(equation);
    expect(imported.compatibility.issues).not.toContainEqual(
      expect.objectContaining({ code: 'docx.equations.unsupported' }),
    );
    await expectNativeWordRunEffects(await createArtifactBlob(imported));
  });

  test('preserves Word text animations and line borders inside OMML', async () => {
    const textEffects = [
      'blinkBackground',
      'lights',
      'antsBlack',
      'antsRed',
      'shimmer',
      'sparkle',
      'none',
    ] as const;
    const lineBorderStyles = [
      'nil',
      'none',
      'single',
      'thick',
      'double',
      'dotted',
      'dashed',
      'dotDash',
      'dotDotDash',
      'triple',
      'thinThickSmallGap',
      'thickThinSmallGap',
      'thinThickThinSmallGap',
      'thinThickMediumGap',
      'thickThinMediumGap',
      'thinThickThinMediumGap',
      'thinThickLargeGap',
      'thickThinLargeGap',
      'thinThickThinLargeGap',
      'wave',
      'doubleWave',
      'dashSmallGap',
      'dashDotStroked',
      'threeDEmboss',
      'threeDEngrave',
      'outset',
      'inset',
    ] as const;
    const equation = {
      version: 1,
      display: 'inline',
      children: [
        {
          type: 'run',
          text: 'animated-bordered',
          wordRunProperties: {
            textEffect: 'shimmer',
            border: {
              style: 'double',
              color: {
                value: '#1a2b3c',
                theme: 'accent1',
                tint: '66',
                shade: '33',
              },
              sizeEighthPoints: 24,
              spacingPoints: 2,
              shadow: true,
              frame: false,
            },
          },
        },
        {
          type: 'run',
          text: 'explicit-border-removal',
          wordRunProperties: {
            textEffect: 'none',
            border: { style: 'none', shadow: false, frame: false },
          },
        },
        {
          type: 'run',
          text: 'native-wave-border',
          wordRunProperties: {
            textEffect: 'antsBlack',
            border: {
              style: 'wave',
              color: { theme: 'accent2' },
              sizeEighthPoints: 16,
              spacingPoints: 3,
            },
          },
        },
        {
          type: 'run',
          text: 'theme-only-border',
          wordRunProperties: {
            textEffect: 'lights',
            border: {
              style: 'single',
              color: { theme: 'accent4' },
              sizeEighthPoints: 8,
            },
          },
        },
        {
          type: 'nary',
          operator: '\u2211',
          limitLocation: 'underOver',
          controlProperties: {
            textEffect: 'sparkle',
            border: {
              style: 'dotted',
              color: { value: 'auto' },
              sizeEighthPoints: 8,
              spacingPoints: 0,
            },
          },
          children: [{ type: 'run', text: 'operator-adornments' }],
        },
      ],
    } as unknown as WorkDocumentEquation;
    expect(normalizeDocumentEquation(equation)).toEqual(equation);

    const equationWithWordRunProperties = (wordRunProperties: unknown) =>
      ({
        version: 1,
        display: 'inline',
        children: [{ type: 'run', text: 'x', wordRunProperties }],
      }) as unknown as WorkDocumentEquation;
    for (const textEffect of textEffects) {
      expect(
        normalizeDocumentEquation(equationWithWordRunProperties({ textEffect }))
          ?.children[0],
      ).toMatchObject({ wordRunProperties: { textEffect } });
    }
    for (const style of lineBorderStyles) {
      expect(
        normalizeDocumentEquation(
          equationWithWordRunProperties({ border: { style } }),
        )?.children[0],
      ).toMatchObject({ wordRunProperties: { border: { style } } });
    }
    for (const border of [
      { style: 'single', sizeEighthPoints: 2, spacingPoints: 0 },
      { style: 'single', sizeEighthPoints: 96, spacingPoints: 31 },
      {
        style: 'double',
        color: { theme: 'accent1', tint: '66', shade: '33' },
        shadow: false,
        frame: true,
      },
    ]) {
      expect(
        normalizeDocumentEquation(equationWithWordRunProperties({ border }))
          ?.children[0],
      ).toMatchObject({ wordRunProperties: { border } });
    }
    const invalidModels = [
      { textEffect: 'pulse' },
      { textEffect: null },
      { border: null },
      { border: { style: 'apples' } },
      { border: { style: 'single', extra: true } },
      { border: { style: 'single', sizeEighthPoints: 1 } },
      { border: { style: 'single', sizeEighthPoints: 97 } },
      { border: { style: 'single', sizeEighthPoints: 2.5 } },
      { border: { style: 'single', spacingPoints: -1 } },
      { border: { style: 'single', spacingPoints: 32 } },
      { border: { style: 'single', spacingPoints: 0.5 } },
      { border: { style: 'single', shadow: 'true' } },
      { border: { style: 'single', frame: 1 } },
      { border: { style: 'single', color: { value: '#xyzxyz' } } },
    ];
    expect(
      invalidModels.map((properties) =>
        normalizeDocumentEquation(equationWithWordRunProperties(properties)),
      ),
    ).toEqual(invalidModels.map(() => null));

    const wordRun = (properties: string, namespace = WORD_NAMESPACE) =>
      `<m:r xmlns:w="${namespace}"><w:rPr>${properties}</w:rPr><m:t>x</m:t></m:r>`;
    const adornmentProperties =
      '<w:effect w:val="shimmer"/><w:bdr w:val="double" w:color="1A2B3C" w:themeColor="accent1" w:themeTint="66" w:themeShade="33" w:sz="24" w:space="2" w:shadow="on" w:frame="0"/>';
    expect(
      inspectEquationModel(wordRun(adornmentProperties))?.children[0],
    ).toEqual({
      type: 'run',
      text: 'x',
      wordRunProperties: {
        textEffect: 'shimmer',
        border: {
          style: 'double',
          color: {
            value: '#1a2b3c',
            theme: 'accent1',
            tint: '66',
            shade: '33',
          },
          sizeEighthPoints: 24,
          spacingPoints: 2,
          shadow: true,
          frame: false,
        },
      },
    });
    expect(
      inspectEquationRoot(
        `<m:oMath xmlns:m="${STRICT_MATH_NAMESPACE}" xmlns:w="${STRICT_WORD_NAMESPACE}"><m:r><w:rPr>${adornmentProperties}</w:rPr><m:t>strict-adornments</m:t></m:r></m:oMath>`,
      ),
    ).toMatchObject({
      status: 'supported',
      equation: {
        children: [
          {
            wordRunProperties: {
              textEffect: 'shimmer',
              border: {
                style: 'double',
                sizeEighthPoints: 24,
                spacingPoints: 2,
                shadow: true,
                frame: false,
              },
            },
          },
        ],
      },
    });
    expect(
      textEffects.map((value) =>
        inspectEquationBody(wordRun(`<w:effect w:val="${value}"/>`)),
      ),
    ).toEqual(textEffects.map(() => 'supported'));
    expect(
      lineBorderStyles.map((value) =>
        inspectEquationBody(wordRun(`<w:bdr w:val="${value}"/>`)),
      ),
    ).toEqual(lineBorderStyles.map(() => 'supported'));

    const unsupported = [
      wordRun('<w:effect/>'),
      wordRun('<w:effect w:val="pulse"/>'),
      wordRun('<w:bdr/>'),
      wordRun('<w:bdr w:val="apples"/>'),
      wordRun('<w:bdr w:val="single" w:sz="1"/>'),
      wordRun('<w:bdr w:val="single" w:sz="97"/>'),
      wordRun('<w:bdr w:val="single" w:sz="2.5"/>'),
      wordRun('<w:bdr w:val="single" w:space="32"/>'),
      wordRun('<w:bdr w:val="single" w:shadow="maybe"/>'),
      wordRun('<w:bdr w:val="single" w:frame="2"/>'),
      wordRun('<w:bdr w:val="single" w:color="XYZXYZ"/>'),
      wordRun('<w:bdr w:val="single" w:themeColor="none" w:themeTint="80"/>'),
      wordRun('<w:bdr w:val="single"/><w:effect w:val="none"/>'),
      wordRun('<w:effect w:val="none"/><w:effect w:val="shimmer"/>'),
      wordRun('<w:bdr w:val="single"/><w:bdr w:val="double"/>'),
      wordRun('<w:effect val="shimmer"/>'),
      wordRun('<w:bdr val="single"/>'),
      wordRun('<w:bdr w:val="single" w:extra="semantic"/>'),
      wordRun(
        `<w:bdr xmlns:r="${RELATIONSHIP_NAMESPACE}" w:val="single" r:id="rIdUnsafe"/>`,
      ),
      wordRun('<w:bdr w:val="single"><w:b/></w:bdr>'),
      wordRun(`<v:effect xmlns:v="${VENDOR_NAMESPACE}" v:val="shimmer"/>`),
      wordRun(`<v:bdr xmlns:v="${VENDOR_NAMESPACE}" v:val="single"/>`),
      wordRun('<m:effect m:val="shimmer"/>'),
      wordRun('<m:bdr m:val="single"/>'),
    ];
    expect(unsupported.map(inspectEquationBody)).toEqual(
      unsupported.map(() => 'unsupported'),
    );

    const document = new DOMParser().parseFromString('', 'text/html');
    const preview = createDocumentEquationElement(document, equation);
    const styleFor = (text: string) =>
      Array.from(preview.querySelectorAll('mtext, mo'))
        .find((element) => element.textContent === text)
        ?.getAttribute('style');
    expect(styleFor('animated-bordered')).toContain(
      'border:3pt double #1a2b3c',
    );
    expect(styleFor('animated-bordered')).toContain('padding:2pt');
    expect(styleFor('animated-bordered')).not.toMatch(/animation|box-shadow/iu);
    expect(styleFor('explicit-border-removal')).toContain('border-style:none');
    expect(styleFor('native-wave-border') ?? '').not.toMatch(
      /border|padding/iu,
    );
    expect(styleFor('theme-only-border') ?? '').not.toMatch(/border|padding/iu);
    expect(styleFor('\u2211')).toContain('border:1pt dotted currentColor');
    expect(styleFor('\u2211')).toContain('padding:0pt');
    const sanitized = new DOMParser().parseFromString(
      sanitizeDocumentPageChromeHtml(preview.outerHTML),
      'text/html',
    );
    expect(
      Array.from(sanitized.querySelectorAll('mtext'))
        .find((element) => element.textContent === 'animated-bordered')
        ?.getAttribute('style'),
    ).toContain('border:3pt double #1a2b3c');

    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    artifact.content.html = `<p>${preview.outerHTML}</p>`;
    const first = await createArtifactBlob(artifact);
    await expectNativeWordRunAdornments(first);
    const imported = await importOfficeFile(
      new File([first], 'word-run-adornments.docx', { type: first.type }),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const importedDocument = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    const importedEquation = importedDocument.body.querySelector<HTMLElement>(
      '[data-document-equation]',
    );
    expect(
      documentEquationFromElement(importedEquation as HTMLElement),
    ).toEqual(equation);
    expect(imported.compatibility.issues).not.toContainEqual(
      expect.objectContaining({ code: 'docx.equations.unsupported' }),
    );
    await expectNativeWordRunAdornments(await createArtifactBlob(imported));
  });

  test('preserves Word character geometry inside OMML', async () => {
    const equation = {
      version: 1,
      display: 'inline',
      children: [
        {
          type: 'run',
          text: 'expanded-raised',
          wordRunProperties: {
            characterSpacingTwips: 200,
            characterScalePercent: 75,
            kerningThresholdHalfPoints: 24,
            positionHalfPoints: 6,
            fontSize: 14,
          },
        },
        {
          type: 'run',
          text: 'explicit-resets',
          wordRunProperties: {
            characterSpacingTwips: 0,
            characterScalePercent: 100,
            kerningThresholdHalfPoints: 0,
            positionHalfPoints: 0,
            fontSize: 12,
          },
        },
        {
          type: 'run',
          text: 'kerning-below-threshold',
          wordRunProperties: {
            kerningThresholdHalfPoints: 25,
            fontSize: 12,
          },
        },
        {
          type: 'run',
          text: 'kerning-inherited-size',
          wordRunProperties: { kerningThresholdHalfPoints: 24 },
        },
        {
          type: 'nary',
          operator: '\u2211',
          limitLocation: 'underOver',
          controlProperties: {
            characterSpacingTwips: -40,
            characterScalePercent: 125,
            kerningThresholdHalfPoints: 30,
            positionHalfPoints: -4,
            fontSize: 15,
          },
          children: [{ type: 'run', text: 'operator-geometry' }],
        },
      ],
    } as unknown as WorkDocumentEquation;
    expect(normalizeDocumentEquation(equation)).toEqual(equation);

    const equationWithWordRunProperties = (wordRunProperties: unknown) =>
      ({
        version: 1,
        display: 'inline',
        children: [{ type: 'run', text: 'x', wordRunProperties }],
      }) as unknown as WorkDocumentEquation;
    const validBoundaries = [
      { characterSpacingTwips: -31_680 },
      { characterSpacingTwips: 31_680 },
      { characterScalePercent: 1 },
      { characterScalePercent: 600 },
      { kerningThresholdHalfPoints: 0 },
      { kerningThresholdHalfPoints: 3_277 },
      { positionHalfPoints: -2_147_483_648 },
      { positionHalfPoints: 2_147_483_647 },
    ];
    for (const properties of validBoundaries) {
      expect(
        normalizeDocumentEquation(equationWithWordRunProperties(properties))
          ?.children[0],
      ).toMatchObject({ wordRunProperties: properties });
    }
    const invalidModels = [
      { characterSpacingTwips: -31_681 },
      { characterSpacingTwips: 31_681 },
      { characterSpacingTwips: 0.5 },
      { characterSpacingTwips: '200' },
      { characterScalePercent: 0 },
      { characterScalePercent: 601 },
      { characterScalePercent: 75.5 },
      { characterScalePercent: '75' },
      { kerningThresholdHalfPoints: -1 },
      { kerningThresholdHalfPoints: 3_278 },
      { kerningThresholdHalfPoints: 0.5 },
      { positionHalfPoints: -2_147_483_649 },
      { positionHalfPoints: 2_147_483_648 },
      { positionHalfPoints: 0.5 },
    ];
    expect(
      invalidModels.map((properties) =>
        normalizeDocumentEquation(equationWithWordRunProperties(properties)),
      ),
    ).toEqual(invalidModels.map(() => null));

    const wordRun = (properties: string, namespace = WORD_NAMESPACE) =>
      `<m:r xmlns:w="${namespace}"><w:rPr>${properties}</w:rPr><m:t>x</m:t></m:r>`;
    expect(
      inspectEquationModel(
        wordRun(
          '<w:spacing w:val="+00200"/><w:w w:val="+0075"/><w:kern w:val="+0024"/><w:position w:val="-0006"/><w:sz w:val="28"/>',
        ),
      )?.children[0],
    ).toEqual({
      type: 'run',
      text: 'x',
      wordRunProperties: {
        characterSpacingTwips: 200,
        characterScalePercent: 75,
        kerningThresholdHalfPoints: 24,
        positionHalfPoints: -6,
        fontSize: 14,
      },
    });
    expect(inspectEquationModel(wordRun('<w:w/>'))?.children[0]).toMatchObject({
      wordRunProperties: { characterScalePercent: 100 },
    });
    expect(
      inspectEquationRoot(
        `<m:oMath xmlns:m="${STRICT_MATH_NAMESPACE}" xmlns:w="${STRICT_WORD_NAMESPACE}"><m:r><w:rPr><w:spacing w:val="-40"/><w:w w:val="125"/><w:kern w:val="24"/><w:position w:val="1.27cm"/><w:sz w:val="12pt"/></w:rPr><m:t>strict-geometry</m:t></m:r></m:oMath>`,
      ),
    ).toMatchObject({
      status: 'supported',
      equation: {
        children: [
          {
            wordRunProperties: {
              characterSpacingTwips: -40,
              characterScalePercent: 125,
              kerningThresholdHalfPoints: 24,
              positionHalfPoints: 72,
              fontSize: 12,
            },
          },
        ],
      },
    });
    for (const [source, positionHalfPoints] of [
      ['6.35mm', 36],
      ['1.27cm', 72],
      ['0.125in', 18],
      ['9pt', 18],
      ['0.75pc', 18],
      ['0.75pi', 18],
      ['-0.5pt', -1],
      ['0pt', 0],
    ] as const) {
      expect(
        inspectEquationModel(
          wordRun(`<w:position w:val="${source}"/>`, STRICT_WORD_NAMESPACE),
        )?.children[0],
        source,
      ).toMatchObject({ wordRunProperties: { positionHalfPoints } });
    }

    const unsupported = [
      wordRun('<w:spacing/>'),
      wordRun('<w:spacing w:val="-31681"/>'),
      wordRun('<w:spacing w:val="31681"/>'),
      wordRun('<w:spacing w:val="1.5"/>'),
      wordRun('<w:spacing w:val="1pt"/>', STRICT_WORD_NAMESPACE),
      wordRun('<w:w w:val="0"/>'),
      wordRun('<w:w w:val="601"/>'),
      wordRun('<w:w w:val="75.5"/>'),
      wordRun('<w:kern/>'),
      wordRun('<w:kern w:val="-1"/>'),
      wordRun('<w:kern w:val="3278"/>'),
      wordRun('<w:kern w:val="1pt"/>', STRICT_WORD_NAMESPACE),
      wordRun('<w:position/>'),
      wordRun('<w:position w:val="-2147483649"/>'),
      wordRun('<w:position w:val="2147483648"/>'),
      wordRun('<w:position w:val="1.5"/>'),
      wordRun('<w:position w:val="1pt"/>'),
      wordRun('<w:position w:val="0.1pt"/>', STRICT_WORD_NAMESPACE),
      wordRun('<w:position w:val="+1pt"/>', STRICT_WORD_NAMESPACE),
      wordRun('<w:position val="2"/>'),
      wordRun('<w:position w:val="2" w:extra="semantic"/>'),
      wordRun(
        `<w:position xmlns:r="${RELATIONSHIP_NAMESPACE}" w:val="2" r:id="rIdUnsafe"/>`,
      ),
      wordRun('<w:position w:val="2"><w:b/></w:position>'),
      wordRun('<w:position w:val="2"/><w:position w:val="4"/>'),
      wordRun('<w:kern w:val="24"/><w:w w:val="75"/>'),
      wordRun('<w:sz w:val="24"/><w:position w:val="2"/>'),
      wordRun(`<v:spacing xmlns:v="${VENDOR_NAMESPACE}" w:val="20"/>`),
      wordRun('<m:position m:val="2"/>'),
    ];
    expect(unsupported.map(inspectEquationBody)).toEqual(
      unsupported.map(() => 'unsupported'),
    );

    const document = new DOMParser().parseFromString('', 'text/html');
    const preview = createDocumentEquationElement(document, equation);
    const styleFor = (text: string) =>
      Array.from(preview.querySelectorAll('mtext, mo'))
        .find((element) => element.textContent === text)
        ?.getAttribute('style');
    expect(styleFor('expanded-raised')).toContain('letter-spacing:10pt');
    expect(styleFor('expanded-raised')).toContain('font-stretch:75%');
    expect(styleFor('expanded-raised')).toContain('font-kerning:normal');
    expect(styleFor('expanded-raised')).toContain('vertical-align:3pt');
    expect(styleFor('explicit-resets')).toContain('letter-spacing:0pt');
    expect(styleFor('explicit-resets')).toContain('font-stretch:100%');
    expect(styleFor('explicit-resets')).toContain('font-kerning:normal');
    expect(styleFor('explicit-resets')).toContain('vertical-align:0pt');
    expect(styleFor('kerning-below-threshold')).toContain('font-kerning:none');
    expect(styleFor('kerning-inherited-size')).toBeNull();
    expect(styleFor('\u2211')).toContain('letter-spacing:-2pt');
    expect(styleFor('\u2211')).toContain('font-stretch:125%');
    expect(styleFor('\u2211')).toContain('font-kerning:normal');
    expect(styleFor('\u2211')).toContain('vertical-align:-2pt');
    const sanitized = new DOMParser().parseFromString(
      sanitizeDocumentPageChromeHtml(preview.outerHTML),
      'text/html',
    );
    expect(
      Array.from(sanitized.querySelectorAll('mtext, mo'))
        .find((element) => element.textContent === 'expanded-raised')
        ?.getAttribute('style'),
    ).toBe(styleFor('expanded-raised'));

    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    artifact.content.html = `<p>${preview.outerHTML}</p>`;
    const first = await createArtifactBlob(artifact);
    await expectNativeWordRunGeometry(first);
    const imported = await importOfficeFile(
      new File([first], 'word-run-geometry.docx', { type: first.type }),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const importedDocument = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    const importedEquation = importedDocument.body.querySelector<HTMLElement>(
      '[data-document-equation]',
    );
    expect(
      documentEquationFromElement(importedEquation as HTMLElement),
    ).toEqual(equation);
    expect(imported.compatibility.issues).not.toContainEqual(
      expect.objectContaining({ code: 'docx.equations.unsupported' }),
    );
    await expectNativeWordRunGeometry(await createArtifactBlob(imported));
  });

  test('preserves Word highlight and patterned shading inside OMML', async () => {
    const equation: WorkDocumentEquation = {
      version: 1,
      display: 'inline',
      children: [
        {
          type: 'run',
          text: 'highlighted',
          wordRunProperties: {
            highlight: 'yellow',
            shading: {
              pattern: 'clear',
              fill: { value: '#112233' },
            },
          },
        },
        {
          type: 'run',
          text: 'clear-shading',
          wordRunProperties: {
            shading: {
              pattern: 'clear',
              fill: {
                value: '#112233',
                theme: 'accent4',
                tint: '80',
              },
            },
          },
        },
        {
          type: 'run',
          text: 'solid-shading',
          wordRunProperties: {
            shading: {
              pattern: 'solid',
              color: {
                value: '#abcdef',
                theme: 'text2',
                shade: '40',
              },
              fill: { value: '#445566' },
            },
          },
        },
        {
          type: 'run',
          text: 'patterned-shading',
          wordRunProperties: {
            shading: {
              pattern: 'pct20',
              color: { value: '#ff0000' },
              fill: { theme: 'accent3', tint: '20' },
            },
          },
        },
        {
          type: 'run',
          text: 'theme-only-shading',
          wordRunProperties: {
            shading: {
              pattern: 'clear',
              fill: { theme: 'accent2', tint: '99' },
            },
          },
        },
        {
          type: 'run',
          text: 'unhighlighted',
          wordRunProperties: {
            highlight: 'none',
            shading: {
              pattern: 'clear',
              fill: { value: '#112233' },
            },
          },
        },
        {
          type: 'run',
          text: 'nil-shading',
          wordRunProperties: {
            shading: {
              pattern: 'nil',
              fill: { value: '#112233' },
            },
          },
        },
        {
          type: 'nary',
          operator: '\u2211',
          limitLocation: 'underOver',
          controlProperties: { highlight: 'darkCyan' },
          children: [{ type: 'run', text: 'operator-background' }],
        },
      ],
    };
    expect(normalizeDocumentEquation(equation)).toEqual(equation);

    const equationWithWordRunProperties = (wordRunProperties: unknown) =>
      ({
        version: 1,
        display: 'inline',
        children: [{ type: 'run', text: 'x', wordRunProperties }],
      }) as unknown as WorkDocumentEquation;
    const invalidModels = [
      { highlight: 'orange' },
      { shading: {} },
      { shading: { pattern: 'pct33' } },
      { shading: { pattern: 'clear', extra: true } },
      { shading: { pattern: 'clear', color: {} } },
      { shading: { pattern: 'clear', fill: { value: '#12345' } } },
      {
        shading: {
          pattern: 'clear',
          fill: { theme: 'none', tint: '80' },
        },
      },
    ];
    expect(
      invalidModels.map((properties) =>
        normalizeDocumentEquation(equationWithWordRunProperties(properties)),
      ),
    ).toEqual(invalidModels.map(() => null));

    const wordRun = (properties: string, namespace = WORD_NAMESPACE) =>
      `<m:r xmlns:w="${namespace}"><w:rPr>${properties}</w:rPr><m:t>x</m:t></m:r>`;
    const highlightValues = [
      'black',
      'blue',
      'cyan',
      'green',
      'magenta',
      'red',
      'yellow',
      'white',
      'darkBlue',
      'darkCyan',
      'darkGreen',
      'darkMagenta',
      'darkRed',
      'darkYellow',
      'darkGray',
      'lightGray',
      'none',
    ];
    expect(
      highlightValues.map((value) =>
        inspectEquationBody(wordRun(`<w:highlight w:val="${value}"/>`)),
      ),
    ).toEqual(highlightValues.map(() => 'supported'));
    const shadingPatterns = [
      'nil',
      'clear',
      'solid',
      'horzStripe',
      'vertStripe',
      'reverseDiagStripe',
      'diagStripe',
      'horzCross',
      'diagCross',
      'thinHorzStripe',
      'thinVertStripe',
      'thinReverseDiagStripe',
      'thinDiagStripe',
      'thinHorzCross',
      'thinDiagCross',
      'pct5',
      'pct10',
      'pct12',
      'pct15',
      'pct20',
      'pct25',
      'pct30',
      'pct35',
      'pct37',
      'pct40',
      'pct45',
      'pct50',
      'pct55',
      'pct60',
      'pct62',
      'pct65',
      'pct70',
      'pct75',
      'pct80',
      'pct85',
      'pct87',
      'pct90',
      'pct95',
    ];
    expect(
      shadingPatterns.map((value) =>
        inspectEquationBody(wordRun(`<w:shd w:val="${value}"/>`)),
      ),
    ).toEqual(shadingPatterns.map(() => 'supported'));

    const invalidMarkup = [
      wordRun('<w:highlight/>'),
      wordRun('<w:highlight w:val="orange"/>'),
      wordRun('<w:highlight val="yellow"/>'),
      wordRun('<w:highlight w:val="yellow" w:extra="semantic"/>'),
      wordRun(
        `<w:highlight xmlns:r="${RELATIONSHIP_NAMESPACE}" w:val="yellow" r:id="rIdUnsafe"/>`,
      ),
      wordRun('<w:highlight w:val="yellow"><w:b/></w:highlight>'),
      wordRun('<w:highlight w:val="yellow"/><w:highlight w:val="red"/>'),
      wordRun('<w:shd/>'),
      wordRun('<w:shd w:val="pct33"/>'),
      wordRun('<w:shd w:val="clear" w:fill="12345"/>'),
      wordRun('<w:shd w:val="clear" w:themeFill="none"/>'),
      wordRun('<w:shd w:val="clear" w:themeFillTint="80"/>'),
      wordRun('<w:shd w:val="clear" fill="112233"/>'),
      wordRun('<w:shd w:val="clear" w:extra="semantic"/>'),
      wordRun(
        `<w:shd xmlns:r="${RELATIONSHIP_NAMESPACE}" w:val="clear" r:id="rIdUnsafe"/>`,
      ),
      wordRun('<w:shd w:val="clear"><w:b/></w:shd>'),
      wordRun('<w:shd w:val="clear"/><w:shd w:val="solid"/>'),
      wordRun('<w:shd w:val="clear"/><w:highlight w:val="yellow"/>'),
      wordRun('<w:shd w:val="clear"/><w:u w:val="single"/>'),
      wordRun(`<v:highlight xmlns:v="${VENDOR_NAMESPACE}" w:val="yellow"/>`),
      wordRun(`<v:shd xmlns:v="${VENDOR_NAMESPACE}" w:val="clear"/>`),
      wordRun('<m:highlight m:val="yellow"/>'),
      wordRun('<m:shd m:val="clear"/>'),
    ];
    expect(invalidMarkup.map(inspectEquationBody)).toEqual(
      invalidMarkup.map(() => 'unsupported'),
    );

    const strictInspection = inspectEquationRoot(
      `<m:oMath xmlns:m="${STRICT_MATH_NAMESPACE}" xmlns:w="${STRICT_WORD_NAMESPACE}"><m:r><w:rPr><w:highlight w:val="darkBlue"/><w:shd w:val="pct37" w:color="ABCDEF" w:themeColor="text2" w:themeShade="40" w:fill="112233" w:themeFill="accent4" w:themeFillTint="80"/></w:rPr><m:t>strict-background</m:t></m:r></m:oMath>`,
    );
    expect(strictInspection).toMatchObject({
      status: 'supported',
      equation: {
        children: [
          {
            wordRunProperties: {
              highlight: 'darkBlue',
              shading: {
                pattern: 'pct37',
                color: {
                  value: '#abcdef',
                  theme: 'text2',
                  shade: '40',
                },
                fill: {
                  value: '#112233',
                  theme: 'accent4',
                  tint: '80',
                },
              },
            },
          },
        ],
      },
    });

    const document = new DOMParser().parseFromString('', 'text/html');
    const preview = createDocumentEquationElement(document, equation);
    const backgroundFor = (text: string) =>
      Array.from(preview.querySelectorAll('mtext, mo'))
        .find((element) => element.textContent === text)
        ?.getAttribute('mathbackground');
    expect(backgroundFor('highlighted')).toBe('#ffff00');
    expect(backgroundFor('clear-shading')).toBe('#112233');
    expect(backgroundFor('solid-shading')).toBe('#abcdef');
    expect(backgroundFor('patterned-shading')).toBeNull();
    expect(backgroundFor('theme-only-shading')).toBeNull();
    expect(backgroundFor('unhighlighted')).toBe('transparent');
    expect(backgroundFor('nil-shading')).toBe('transparent');
    expect(backgroundFor('\u2211')).toBe('#008080');
    const sanitized = new DOMParser().parseFromString(
      sanitizeDocumentPageChromeHtml(preview.outerHTML),
      'text/html',
    );
    expect(
      Array.from(sanitized.querySelectorAll('[mathbackground]')).map(
        (element) => element.getAttribute('mathbackground'),
      ),
    ).toEqual([
      '#ffff00',
      '#112233',
      '#abcdef',
      'transparent',
      'transparent',
      '#008080',
    ]);

    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    artifact.content.html = `<p>${preview.outerHTML}</p>`;
    const first = await createArtifactBlob(artifact);
    await expectNativeWordRunBackgrounds(first);
    const imported = await importOfficeFile(
      new File([first], 'word-run-backgrounds.docx', { type: first.type }),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const importedDocument = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    const importedEquation = importedDocument.body.querySelector<HTMLElement>(
      '[data-document-equation]',
    );
    expect(
      documentEquationFromElement(importedEquation as HTMLElement),
    ).toEqual(equation);
    expect(imported.compatibility.issues).not.toContainEqual(
      expect.objectContaining({ code: 'docx.equations.unsupported' }),
    );
    await expectNativeWordRunBackgrounds(await createArtifactBlob(imported));
  });

  test('preserves bounded manual run widths as native-only OMML metadata', async () => {
    const equation = {
      version: 1,
      display: 'inline',
      children: [
        {
          type: 'run',
          text: 'fit-width',
          wordRunProperties: {
            fitText: { widthTwips: 720, id: 50 },
          },
        },
        {
          type: 'run',
          text: 'fit-zero',
          wordRunProperties: {
            fitText: { widthTwips: 0, id: 0 },
          },
        },
        {
          type: 'run',
          text: 'fit-without-id',
          wordRunProperties: {
            fitText: { widthTwips: 31_680 },
          },
        },
        {
          type: 'nary',
          operator: '\u2211',
          limitLocation: 'underOver',
          controlProperties: {
            fitText: { widthTwips: 1_440, id: -7 },
          },
          children: [{ type: 'run', text: 'operator-fit-width' }],
        },
      ],
    } as unknown as WorkDocumentEquation;
    expect(normalizeDocumentEquation(equation)).toEqual(equation);

    const equationWithWordRunProperties = (wordRunProperties: unknown) =>
      ({
        version: 1,
        display: 'inline',
        children: [{ type: 'run', text: 'x', wordRunProperties }],
      }) as unknown as WorkDocumentEquation;
    for (const fitText of [
      { widthTwips: 0 },
      { widthTwips: 31_680 },
      { widthTwips: 720, id: -2_147_483_648 },
      { widthTwips: 720, id: 2_147_483_647 },
    ]) {
      expect(
        normalizeDocumentEquation(equationWithWordRunProperties({ fitText }))
          ?.children[0],
      ).toMatchObject({ wordRunProperties: { fitText } });
    }
    const invalidModels = [
      { fitText: null },
      { fitText: {} },
      { fitText: { widthTwips: 720, extra: true } },
      { fitText: { widthTwips: -1 } },
      { fitText: { widthTwips: 31_681 } },
      { fitText: { widthTwips: 0.5 } },
      { fitText: { widthTwips: '720' } },
      { fitText: { widthTwips: 720, id: -2_147_483_649 } },
      { fitText: { widthTwips: 720, id: 2_147_483_648 } },
      { fitText: { widthTwips: 720, id: 0.5 } },
      { fitText: { widthTwips: 720, id: '50' } },
    ];
    expect(
      invalidModels.map((properties) =>
        normalizeDocumentEquation(equationWithWordRunProperties(properties)),
      ),
    ).toEqual(invalidModels.map(() => null));

    const wordRun = (properties: string, namespace = WORD_NAMESPACE) =>
      `<m:r xmlns:w="${namespace}"><w:rPr>${properties}</w:rPr><m:t>x</m:t></m:r>`;
    expect(
      inspectEquationModel(wordRun('<w:fitText w:val="720" w:id="50"/>'))
        ?.children[0],
    ).toEqual({
      type: 'run',
      text: 'x',
      wordRunProperties: { fitText: { widthTwips: 720, id: 50 } },
    });
    expect(
      inspectEquationModel(wordRun('<w:fitText w:val="0" w:id="0"/>'))
        ?.children[0],
    ).toMatchObject({
      wordRunProperties: { fitText: { widthTwips: 0, id: 0 } },
    });
    expect(
      inspectEquationModel(wordRun('<w:fitText w:val="31680"/>'))?.children[0],
    ).toMatchObject({
      wordRunProperties: { fitText: { widthTwips: 31_680 } },
    });
    expect(
      inspectEquationModel(
        wordRun('<w:fitText w:val="720" w:id="-2147483648"/>'),
      )?.children[0],
    ).toMatchObject({
      wordRunProperties: {
        fitText: { widthTwips: 720, id: -2_147_483_648 },
      },
    });
    expect(
      inspectEquationModel(
        wordRun('<w:fitText w:val="720" w:id="2147483647"/>'),
      )?.children[0],
    ).toMatchObject({
      wordRunProperties: {
        fitText: { widthTwips: 720, id: 2_147_483_647 },
      },
    });

    const strictWidths = [
      ['0pt', 0],
      ['0.05pt', 1],
      ['0.5in', 720],
      ['12.7mm', 720],
      ['1.27cm', 720],
      ['3pc', 720],
      ['3pi', 720],
      ['22in', 31_680],
    ] as const;
    for (const [source, widthTwips] of strictWidths) {
      expect(
        inspectEquationModel(
          wordRun(
            `<w:fitText w:val="${source}" w:id="+50"/>`,
            STRICT_WORD_NAMESPACE,
          ),
        )?.children[0],
        source,
      ).toMatchObject({
        wordRunProperties: { fitText: { widthTwips, id: 50 } },
      });
    }

    const invalidMarkup = [
      wordRun('<w:fitText/>'),
      wordRun('<w:fitText w:id="50"/>'),
      wordRun('<w:fitText w:val="-1"/>'),
      wordRun('<w:fitText w:val="31681"/>'),
      wordRun('<w:fitText w:val="720.5"/>'),
      wordRun('<w:fitText w:val="720" w:id="-2147483649"/>'),
      wordRun('<w:fitText w:val="720" w:id="2147483648"/>'),
      wordRun('<w:fitText w:val="720" w:id="1.5"/>'),
      wordRun('<w:fitText w:val="720" w:id=""/>'),
      wordRun('<w:fitText val="720"/>'),
      wordRun('<w:fitText w:val="720" id="50"/>'),
      wordRun('<w:fitText w:val="720" w:extra="semantic"/>'),
      wordRun(
        `<w:fitText xmlns:r="${RELATIONSHIP_NAMESPACE}" w:val="720" r:id="rIdUnsafe"/>`,
      ),
      wordRun('<w:fitText w:val="720"><w:b/></w:fitText>'),
      wordRun('<w:fitText w:val="720"/><w:fitText w:val="1440"/>'),
      wordRun('<w:fitText w:val="720"/><w:shd w:val="clear"/>'),
      wordRun('<w:rtl/><w:fitText w:val="720"/>'),
      wordRun(`<v:fitText xmlns:v="${VENDOR_NAMESPACE}" v:val="720"/>`),
      wordRun('<m:fitText m:val="720"/>'),
      wordRun('<w:fitText w:val="0.5in"/>'),
      wordRun('<w:fitText w:val="0.01pt"/>', STRICT_WORD_NAMESPACE),
      wordRun('<w:fitText w:val="22.01in"/>', STRICT_WORD_NAMESPACE),
      wordRun('<w:fitText w:val="-1pt"/>', STRICT_WORD_NAMESPACE),
      wordRun('<w:fitText w:val="+0.5in"/>', STRICT_WORD_NAMESPACE),
      wordRun('<w:fitText w:val="1e2pt"/>', STRICT_WORD_NAMESPACE),
    ];
    expect(invalidMarkup.map(inspectEquationBody)).toEqual(
      invalidMarkup.map(() => 'unsupported'),
    );

    const document = new DOMParser().parseFromString('', 'text/html');
    const preview = createDocumentEquationElement(document, equation);
    for (const text of ['fit-width', 'fit-zero', 'fit-without-id', '\u2211']) {
      const element = Array.from(preview.querySelectorAll('mtext, mo')).find(
        (candidate) => candidate.textContent === text,
      );
      expect(element, text).toBeDefined();
      expect(element?.getAttribute('style'), text).toBeNull();
      expect(element?.getAttribute('width'), text).toBeNull();
    }
    const sanitized = new DOMParser().parseFromString(
      sanitizeDocumentPageChromeHtml(preview.outerHTML),
      'text/html',
    );
    expect(
      documentEquationFromElement(
        sanitized.body.querySelector<HTMLElement>(
          '[data-document-equation]',
        ) as HTMLElement,
      ),
    ).toEqual(equation);

    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    artifact.content.html = `<p>${preview.outerHTML}</p>`;
    const first = await createArtifactBlob(artifact);
    await expectNativeWordRunFitText(first);
    const imported = await importOfficeFile(
      new File([first], 'word-run-fit-text.docx', { type: first.type }),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const importedDocument = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    expect(
      documentEquationFromElement(
        importedDocument.body.querySelector<HTMLElement>(
          '[data-document-equation]',
        ) as HTMLElement,
      ),
    ).toEqual(equation);
    expect(imported.compatibility.issues).not.toContainEqual(
      expect.objectContaining({ code: 'docx.equations.unsupported' }),
    );
    await expectNativeWordRunFitText(await createArtifactBlob(imported));
  });

  test('preserves explicit Word baseline, superscript, and subscript alignment inside OMML', async () => {
    const equation = {
      version: 1,
      display: 'inline',
      children: [
        {
          type: 'run',
          text: 'baseline-aligned',
          wordRunProperties: { verticalAlignment: 'baseline' },
        },
        {
          type: 'run',
          text: 'superscript-aligned',
          wordRunProperties: {
            fontSize: 12,
            verticalAlignment: 'superscript',
          },
        },
        {
          type: 'run',
          text: 'subscript-aligned',
          wordRunProperties: { verticalAlignment: 'subscript' },
        },
        {
          type: 'run',
          text: 'positioned-superscript',
          wordRunProperties: {
            positionHalfPoints: 4,
            verticalAlignment: 'superscript',
          },
        },
        {
          type: 'nary',
          operator: '\u2211',
          limitLocation: 'underOver',
          controlProperties: { verticalAlignment: 'subscript' },
          children: [{ type: 'run', text: 'operator-vertical-alignment' }],
        },
      ],
    } as unknown as WorkDocumentEquation;
    expect(normalizeDocumentEquation(equation)).toEqual(equation);

    const equationWithWordRunProperties = (wordRunProperties: unknown) =>
      ({
        version: 1,
        display: 'inline',
        children: [{ type: 'run', text: 'x', wordRunProperties }],
      }) as unknown as WorkDocumentEquation;
    for (const verticalAlignment of [
      'baseline',
      'superscript',
      'subscript',
    ] as const) {
      expect(
        normalizeDocumentEquation(
          equationWithWordRunProperties({ verticalAlignment }),
        )?.children[0],
      ).toMatchObject({ wordRunProperties: { verticalAlignment } });
    }
    const invalidModels = [
      { verticalAlignment: null },
      { verticalAlignment: '' },
      { verticalAlignment: 'super' },
      { verticalAlignment: 'Superscript' },
      { verticalAlignment: 1 },
    ];
    expect(
      invalidModels.map((properties) =>
        normalizeDocumentEquation(equationWithWordRunProperties(properties)),
      ),
    ).toEqual(invalidModels.map(() => null));

    const wordRun = (properties: string, namespace = WORD_NAMESPACE) =>
      `<m:r xmlns:w="${namespace}"><w:rPr>${properties}</w:rPr><m:t>x</m:t></m:r>`;
    for (const [source, verticalAlignment] of [
      ['baseline', 'baseline'],
      [' superscript ', 'superscript'],
      ['subscript', 'subscript'],
    ] as const) {
      expect(
        inspectEquationModel(wordRun(`<w:vertAlign w:val="${source}"/>`))
          ?.children[0],
      ).toEqual({
        type: 'run',
        text: 'x',
        wordRunProperties: { verticalAlignment },
      });
      expect(
        inspectEquationModel(
          wordRun(`<w:vertAlign w:val="${source}"/>`, STRICT_WORD_NAMESPACE),
        )?.children[0],
      ).toEqual({
        type: 'run',
        text: 'x',
        wordRunProperties: { verticalAlignment },
      });
    }
    expect(
      inspectEquationModel(
        wordRun('<w:position w:val="4"/><w:vertAlign w:val="superscript"/>'),
      )?.children[0],
    ).toMatchObject({
      wordRunProperties: {
        positionHalfPoints: 4,
        verticalAlignment: 'superscript',
      },
    });

    const invalidMarkup = [
      wordRun('<w:vertAlign/>'),
      wordRun('<w:vertAlign w:val=""/>'),
      wordRun('<w:vertAlign w:val="super"/>'),
      wordRun('<w:vertAlign w:val="Superscript"/>'),
      wordRun('<w:vertAlign w:val="1"/>'),
      wordRun('<w:vertAlign val="superscript"/>'),
      wordRun('<w:vertAlign w:val="superscript" w:extra="semantic"/>'),
      wordRun(
        `<w:vertAlign xmlns:r="${RELATIONSHIP_NAMESPACE}" w:val="superscript" r:id="rIdUnsafe"/>`,
      ),
      wordRun('<w:vertAlign w:val="superscript"><w:b/></w:vertAlign>'),
      wordRun(
        '<w:vertAlign w:val="superscript"/><w:vertAlign w:val="subscript"/>',
      ),
      wordRun('<w:vertAlign w:val="superscript"/><w:fitText w:val="720"/>'),
      wordRun('<w:rtl/><w:vertAlign w:val="superscript"/>'),
      wordRun(
        `<v:vertAlign xmlns:v="${VENDOR_NAMESPACE}" v:val="superscript"/>`,
      ),
      wordRun('<m:vertAlign m:val="superscript"/>'),
    ];
    expect(invalidMarkup.map(inspectEquationBody)).toEqual(
      invalidMarkup.map(() => 'unsupported'),
    );

    const document = new DOMParser().parseFromString('', 'text/html');
    const preview = createDocumentEquationElement(document, equation);
    const previewFor = (text: string) =>
      Array.from(preview.querySelectorAll('mtext, mo')).find(
        (candidate) => candidate.textContent === text,
      );
    const baseline = previewFor('baseline-aligned');
    expect(baseline?.getAttribute('style')).toContain(
      'vertical-align:baseline',
    );
    expect(baseline?.getAttribute('style')).not.toContain('font-size:smaller');
    const superscript = previewFor('superscript-aligned');
    expect(superscript?.getAttribute('mathsize')).toBe('12pt');
    expect(superscript?.getAttribute('style')).toContain(
      'vertical-align:super',
    );
    expect(superscript?.getAttribute('style')).toContain('font-size:smaller');
    const subscript = previewFor('subscript-aligned');
    expect(subscript?.getAttribute('style')).toContain('vertical-align:sub');
    expect(subscript?.getAttribute('style')).toContain('font-size:smaller');
    const positioned = previewFor('positioned-superscript');
    expect(positioned?.getAttribute('style')).toContain('vertical-align:2pt');
    expect(positioned?.getAttribute('style')).toContain('vertical-align:super');
    expect(positioned?.getAttribute('style')).toContain('font-size:smaller');
    expect(previewFor('\u2211')?.getAttribute('style')).toContain(
      'vertical-align:sub',
    );

    const sanitized = new DOMParser().parseFromString(
      sanitizeDocumentPageChromeHtml(preview.outerHTML),
      'text/html',
    );
    expect(
      documentEquationFromElement(
        sanitized.body.querySelector<HTMLElement>(
          '[data-document-equation]',
        ) as HTMLElement,
      ),
    ).toEqual(equation);
    expect(
      Array.from(sanitized.querySelectorAll('mtext, mo'))
        .find((candidate) => candidate.textContent === 'superscript-aligned')
        ?.getAttribute('style'),
    ).toContain('font-size:smaller');

    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    artifact.content.html = `<p>${preview.outerHTML}</p>`;
    const first = await createArtifactBlob(artifact);
    await expectNativeWordRunVerticalAlignments(first);
    const imported = await importOfficeFile(
      new File([first], 'word-run-vertical-alignment.docx', {
        type: first.type,
      }),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const importedDocument = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    expect(
      documentEquationFromElement(
        importedDocument.body.querySelector<HTMLElement>(
          '[data-document-equation]',
        ) as HTMLElement,
      ),
    ).toEqual(equation);
    expect(imported.compatibility.issues).not.toContainEqual(
      expect.objectContaining({ code: 'docx.equations.unsupported' }),
    );
    await expectNativeWordRunVerticalAlignments(
      await createArtifactBlob(imported),
    );
  });

  test('preserves Word emphasis marks inside OMML', async () => {
    const equation = {
      version: 1,
      display: 'inline',
      children: [
        {
          type: 'run',
          text: 'no-emphasis',
          wordRunProperties: { emphasisMark: 'none' },
        },
        {
          type: 'run',
          text: 'dot-emphasis',
          wordRunProperties: { emphasisMark: 'dot' },
        },
        {
          type: 'run',
          text: 'comma-emphasis',
          wordRunProperties: { emphasisMark: 'comma' },
        },
        {
          type: 'run',
          text: 'circle-emphasis',
          wordRunProperties: { emphasisMark: 'circle' },
        },
        {
          type: 'run',
          text: 'under-dot-emphasis',
          wordRunProperties: { emphasisMark: 'underDot' },
        },
        {
          type: 'run',
          text: 'ordered-emphasis',
          wordRunProperties: {
            verticalAlignment: 'baseline',
            rightToLeft: false,
            complexScript: true,
            emphasisMark: 'dot',
            languages: { latin: 'en-US' },
          },
        },
        {
          type: 'nary',
          operator: '\u2211',
          limitLocation: 'underOver',
          controlProperties: { emphasisMark: 'underDot' },
          children: [{ type: 'run', text: 'operator-emphasis-mark' }],
        },
      ],
    } as unknown as WorkDocumentEquation;
    expect(normalizeDocumentEquation(equation)).toEqual(equation);

    const equationWithWordRunProperties = (wordRunProperties: unknown) =>
      ({
        version: 1,
        display: 'inline',
        children: [{ type: 'run', text: 'x', wordRunProperties }],
      }) as unknown as WorkDocumentEquation;
    for (const emphasisMark of [
      'none',
      'dot',
      'comma',
      'circle',
      'underDot',
    ] as const) {
      expect(
        normalizeDocumentEquation(
          equationWithWordRunProperties({ emphasisMark }),
        )?.children[0],
      ).toMatchObject({ wordRunProperties: { emphasisMark } });
    }
    const invalidModels = [
      { emphasisMark: null },
      { emphasisMark: '' },
      { emphasisMark: 'under-dot' },
      { emphasisMark: 'underDot ' },
      { emphasisMark: 'Dot' },
      { emphasisMark: 1 },
    ];
    expect(
      invalidModels.map((properties) =>
        normalizeDocumentEquation(equationWithWordRunProperties(properties)),
      ),
    ).toEqual(invalidModels.map(() => null));

    const wordRun = (properties: string, namespace = WORD_NAMESPACE) =>
      `<m:r xmlns:w="${namespace}"><w:rPr>${properties}</w:rPr><m:t>x</m:t></m:r>`;
    for (const [source, emphasisMark] of [
      ['none', 'none'],
      [' dot ', 'dot'],
      ['comma', 'comma'],
      ['circle', 'circle'],
      [' underDot ', 'underDot'],
    ] as const) {
      expect(
        inspectEquationModel(wordRun(`<w:em w:val="${source}"/>`))?.children[0],
      ).toEqual({
        type: 'run',
        text: 'x',
        wordRunProperties: { emphasisMark },
      });
      expect(
        inspectEquationModel(
          wordRun(`<w:em w:val="${source}"/>`, STRICT_WORD_NAMESPACE),
        )?.children[0],
      ).toEqual({
        type: 'run',
        text: 'x',
        wordRunProperties: { emphasisMark },
      });
    }
    expect(
      inspectEquationModel(
        wordRun(
          '<w:vertAlign w:val="baseline"/><w:rtl w:val="0"/><w:cs/><w:em w:val="dot"/><w:lang w:val="en-US"/>',
        ),
      )?.children[0],
    ).toEqual({
      type: 'run',
      text: 'x',
      wordRunProperties: {
        verticalAlignment: 'baseline',
        rightToLeft: false,
        complexScript: true,
        emphasisMark: 'dot',
        languages: { latin: 'en-US' },
      },
    });

    const invalidMarkup = [
      wordRun('<w:em/>'),
      wordRun('<w:em w:val=""/>'),
      wordRun('<w:em w:val="under-dot"/>'),
      wordRun('<w:em w:val="Dot"/>'),
      wordRun('<w:em w:val="1"/>'),
      wordRun('<w:em val="dot"/>'),
      wordRun('<w:em w:val="dot" w:extra="semantic"/>'),
      wordRun(
        `<w:em xmlns:r="${RELATIONSHIP_NAMESPACE}" w:val="dot" r:id="rIdUnsafe"/>`,
      ),
      wordRun('<w:em w:val="dot"><w:b/></w:em>'),
      wordRun('<w:em w:val="dot"/><w:em w:val="comma"/>'),
      wordRun('<w:em w:val="dot"/><w:cs/>'),
      wordRun('<w:lang w:val="en-US"/><w:em w:val="dot"/>'),
      wordRun('<w:em w:val="dot"/><w:vertAlign w:val="baseline"/>'),
      wordRun(`<v:em xmlns:v="${VENDOR_NAMESPACE}" v:val="dot"/>`),
      wordRun('<m:em m:val="dot"/>'),
    ];
    expect(invalidMarkup.map(inspectEquationBody)).toEqual(
      invalidMarkup.map(() => 'unsupported'),
    );

    const document = new DOMParser().parseFromString('', 'text/html');
    const preview = createDocumentEquationElement(document, equation);
    const styleFor = (text: string) =>
      Array.from(preview.querySelectorAll('mtext, mo'))
        .find((candidate) => candidate.textContent === text)
        ?.getAttribute('style') ?? '';
    expect(styleFor('no-emphasis')).toContain('text-emphasis-style:none');
    expect(styleFor('no-emphasis')).not.toContain('text-emphasis-position');
    expect(styleFor('dot-emphasis')).toContain(
      'text-emphasis-style:filled dot',
    );
    expect(styleFor('dot-emphasis')).toContain(
      'text-emphasis-position:over right',
    );
    expect(styleFor('comma-emphasis')).toContain('text-emphasis-style:","');
    expect(styleFor('comma-emphasis')).toContain(
      'text-emphasis-position:over right',
    );
    expect(styleFor('circle-emphasis')).toContain(
      'text-emphasis-style:open circle',
    );
    expect(styleFor('circle-emphasis')).toContain(
      'text-emphasis-position:over right',
    );
    expect(styleFor('under-dot-emphasis')).toContain(
      'text-emphasis-style:filled dot',
    );
    expect(styleFor('under-dot-emphasis')).toContain(
      'text-emphasis-position:under right',
    );
    expect(styleFor('\u2211')).toContain('text-emphasis-position:under right');

    const sanitized = new DOMParser().parseFromString(
      sanitizeDocumentPageChromeHtml(preview.outerHTML),
      'text/html',
    );
    expect(
      documentEquationFromElement(
        sanitized.body.querySelector<HTMLElement>(
          '[data-document-equation]',
        ) as HTMLElement,
      ),
    ).toEqual(equation);
    expect(
      Array.from(sanitized.querySelectorAll('mtext, mo'))
        .find((candidate) => candidate.textContent === 'comma-emphasis')
        ?.getAttribute('style'),
    ).toContain('text-emphasis-style:","');

    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    artifact.content.html = `<p>${preview.outerHTML}</p>`;
    const first = await createArtifactBlob(artifact);
    await expectNativeWordRunEmphasisMarks(first);
    const imported = await importOfficeFile(
      new File([first], 'word-run-emphasis-marks.docx', { type: first.type }),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const importedDocument = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    expect(
      documentEquationFromElement(
        importedDocument.body.querySelector<HTMLElement>(
          '[data-document-equation]',
        ) as HTMLElement,
      ),
    ).toEqual(equation);
    expect(imported.compatibility.issues).not.toContainEqual(
      expect.objectContaining({ code: 'docx.equations.unsupported' }),
    );
    await expectNativeWordRunEmphasisMarks(await createArtifactBlob(imported));
  });

  test('preserves East Asian typography settings as native-only OMML metadata', async () => {
    const equation = {
      version: 1,
      display: 'inline',
      children: [
        {
          type: 'run',
          text: 'layout-id-zero',
          wordRunProperties: { eastAsianLayout: { id: 0 } },
        },
        {
          type: 'run',
          text: 'two-lines-round',
          wordRunProperties: {
            eastAsianLayout: {
              combine: true,
              combineBrackets: 'round',
            },
          },
        },
        {
          type: 'run',
          text: 'explicit-combine-reset',
          wordRunProperties: {
            eastAsianLayout: {
              combine: false,
              combineBrackets: 'none',
            },
          },
        },
        {
          type: 'run',
          text: 'horizontal-in-vertical',
          wordRunProperties: {
            eastAsianLayout: { vertical: true },
          },
        },
        {
          type: 'run',
          text: 'explicit-vertical-reset',
          wordRunProperties: {
            eastAsianLayout: {
              vertical: false,
              verticalCompress: true,
            },
          },
        },
        {
          type: 'run',
          text: 'full-east-asian-layout',
          wordRunProperties: {
            languages: { eastAsia: 'ja-JP' },
            eastAsianLayout: {
              id: -2_147_483_648,
              combine: true,
              combineBrackets: 'curly',
              vertical: true,
              verticalCompress: false,
            },
          },
        },
        {
          type: 'run',
          text: 'layout-id-maximum',
          wordRunProperties: {
            eastAsianLayout: { id: 2_147_483_647 },
          },
        },
        {
          type: 'nary',
          operator: '\u2211',
          limitLocation: 'underOver',
          controlProperties: {
            eastAsianLayout: {
              id: -7,
              combine: false,
              combineBrackets: 'angle',
              vertical: true,
              verticalCompress: false,
            },
          },
          children: [{ type: 'run', text: 'operator-east-asian-layout' }],
        },
      ],
    } as unknown as WorkDocumentEquation;
    expect(normalizeDocumentEquation(equation)).toEqual(equation);

    const equationWithWordRunProperties = (wordRunProperties: unknown) =>
      ({
        version: 1,
        display: 'inline',
        children: [{ type: 'run', text: 'x', wordRunProperties }],
      }) as unknown as WorkDocumentEquation;
    expect(
      normalizeDocumentEquation(
        equationWithWordRunProperties({ eastAsianLayout: {} }),
      ),
    ).toEqual(simpleEquation('x'));
    for (const combineBrackets of [
      'none',
      'round',
      'square',
      'angle',
      'curly',
    ] as const) {
      expect(
        normalizeDocumentEquation(
          equationWithWordRunProperties({
            eastAsianLayout: { combineBrackets },
          }),
        )?.children[0],
      ).toMatchObject({
        wordRunProperties: { eastAsianLayout: { combineBrackets } },
      });
    }
    const invalidModels = [
      { eastAsianLayout: null },
      { eastAsianLayout: [] },
      { eastAsianLayout: { extra: true } },
      { eastAsianLayout: { id: '1' } },
      { eastAsianLayout: { id: 1.5 } },
      { eastAsianLayout: { id: -2_147_483_649 } },
      { eastAsianLayout: { id: 2_147_483_648 } },
      { eastAsianLayout: { combine: 1 } },
      { eastAsianLayout: { combineBrackets: '' } },
      { eastAsianLayout: { combineBrackets: 'Round' } },
      { eastAsianLayout: { combineBrackets: 'round ' } },
      { eastAsianLayout: { vertical: 'true' } },
      { eastAsianLayout: { verticalCompress: null } },
    ];
    expect(
      invalidModels.map((properties) =>
        normalizeDocumentEquation(equationWithWordRunProperties(properties)),
      ),
    ).toEqual(invalidModels.map(() => null));

    const wordRun = (properties: string, namespace = WORD_NAMESPACE) =>
      `<m:r xmlns:w="${namespace}"><w:rPr>${properties}</w:rPr><m:t>x</m:t></m:r>`;
    for (const namespace of [WORD_NAMESPACE, STRICT_WORD_NAMESPACE]) {
      expect(
        inspectEquationModel(
          wordRun(
            '<w:eastAsianLayout w:vertCompress=" off " w:combineBrackets="curly" w:vert=" ON " w:id=" +42 " w:combine=" false "/>',
            namespace,
          ),
        )?.children[0],
      ).toEqual({
        type: 'run',
        text: 'x',
        wordRunProperties: {
          eastAsianLayout: {
            id: 42,
            combine: false,
            combineBrackets: 'curly',
            vertical: true,
            verticalCompress: false,
          },
        },
      });
    }
    for (const combineBrackets of [
      'none',
      'round',
      'square',
      'angle',
      'curly',
    ] as const) {
      expect(
        inspectEquationModel(
          wordRun(
            `<w:eastAsianLayout w:combineBrackets="${combineBrackets}"/>`,
          ),
        )?.children[0],
      ).toMatchObject({
        wordRunProperties: { eastAsianLayout: { combineBrackets } },
      });
    }
    expect(inspectEquationModel(wordRun('<w:eastAsianLayout/>'))).toEqual(
      simpleEquation('x'),
    );
    expect(
      inspectEquationModel(
        wordRun(
          '<w:lang w:eastAsia="ja-JP"/><w:eastAsianLayout w:id="-2147483648" w:combine="1" w:combineBrackets="square" w:vert="0" w:vertCompress="true"/>',
        ),
      )?.children[0],
    ).toEqual({
      type: 'run',
      text: 'x',
      wordRunProperties: {
        languages: { eastAsia: 'ja-JP' },
        eastAsianLayout: {
          id: -2_147_483_648,
          combine: true,
          combineBrackets: 'square',
          vertical: false,
          verticalCompress: true,
        },
      },
    });

    const invalidMarkup = [
      wordRun('<w:eastAsianLayout w:id=""/>'),
      wordRun('<w:eastAsianLayout w:id="1.5"/>'),
      wordRun('<w:eastAsianLayout w:id="0x1"/>'),
      wordRun('<w:eastAsianLayout w:id="-2147483649"/>'),
      wordRun('<w:eastAsianLayout w:id="2147483648"/>'),
      wordRun('<w:eastAsianLayout w:combine="maybe"/>'),
      wordRun('<w:eastAsianLayout w:combineBrackets=""/>'),
      wordRun('<w:eastAsianLayout w:combineBrackets="Round"/>'),
      wordRun('<w:eastAsianLayout w:combineBrackets="double"/>'),
      wordRun('<w:eastAsianLayout w:vert=""/>'),
      wordRun('<w:eastAsianLayout w:vertCompress="2"/>'),
      wordRun('<w:eastAsianLayout id="1"/>'),
      wordRun('<w:eastAsianLayout w:id="1" w:extra="semantic"/>'),
      wordRun(
        `<w:eastAsianLayout xmlns:r="${RELATIONSHIP_NAMESPACE}" w:id="1" r:id="rIdUnsafe"/>`,
      ),
      wordRun('<w:eastAsianLayout w:id="1"><w:b/></w:eastAsianLayout>'),
      wordRun('<w:eastAsianLayout w:id="1">meaningful</w:eastAsianLayout>'),
      wordRun('<w:eastAsianLayout w:id="1"/><w:eastAsianLayout w:id="2"/>'),
      wordRun('<w:eastAsianLayout w:id="1"/><w:lang w:eastAsia="ja-JP"/>'),
      wordRun('<w:specVanish/><w:eastAsianLayout w:id="1"/>'),
      wordRun(`<v:eastAsianLayout xmlns:v="${VENDOR_NAMESPACE}" v:id="1"/>`),
      wordRun('<m:eastAsianLayout m:id="1"/>'),
    ];
    expect(invalidMarkup.map(inspectEquationBody)).toEqual(
      invalidMarkup.map(() => 'unsupported'),
    );

    const document = new DOMParser().parseFromString('', 'text/html');
    const preview = createDocumentEquationElement(document, equation);
    expect(preview.outerHTML).not.toMatch(
      /(?:text-combine-upright|writing-mode|text-orientation|transform\s*:)/u,
    );
    const sanitized = new DOMParser().parseFromString(
      sanitizeDocumentPageChromeHtml(preview.outerHTML),
      'text/html',
    );
    expect(
      documentEquationFromElement(
        sanitized.body.querySelector<HTMLElement>(
          '[data-document-equation]',
        ) as HTMLElement,
      ),
    ).toEqual(equation);

    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    artifact.content.html = `<p>${preview.outerHTML}</p>`;
    const first = await createArtifactBlob(artifact);
    await expectNativeWordRunEastAsianLayouts(first);
    const imported = await importOfficeFile(
      new File([first], 'word-run-east-asian-layout.docx', {
        type: first.type,
      }),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const importedDocument = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    expect(
      documentEquationFromElement(
        importedDocument.body.querySelector<HTMLElement>(
          '[data-document-equation]',
        ) as HTMLElement,
      ),
    ).toEqual(equation);
    expect(imported.compatibility.issues).not.toContainEqual(
      expect.objectContaining({ code: 'docx.equations.unsupported' }),
    );
    await expectNativeWordRunEastAsianLayouts(
      await createArtifactBlob(imported),
    );
  });

  test('preserves paragraph-mark visibility flags as inert OMML metadata', async () => {
    const equation = {
      version: 1,
      display: 'inline',
      children: [
        {
          type: 'run',
          text: 'always-hidden-paragraph-mark',
          wordRunProperties: { paragraphMarkAlwaysHidden: true },
        },
        {
          type: 'run',
          text: 'explicit-paragraph-mark-reset',
          wordRunProperties: { paragraphMarkAlwaysHidden: false },
        },
        {
          type: 'run',
          text: 'spec-vanish-with-hidden-run',
          wordRunProperties: {
            hidden: true,
            paragraphMarkAlwaysHidden: true,
          },
        },
        {
          type: 'run',
          text: 'ordered-spec-vanish',
          wordRunProperties: {
            languages: { eastAsia: 'zh-CN' },
            eastAsianLayout: {
              id: 5,
              combine: false,
              vertical: true,
            },
            paragraphMarkAlwaysHidden: false,
          },
        },
        {
          type: 'nary',
          operator: '\u2211',
          limitLocation: 'underOver',
          controlProperties: { paragraphMarkAlwaysHidden: true },
          children: [{ type: 'run', text: 'operator-spec-vanish' }],
        },
      ],
    } as unknown as WorkDocumentEquation;
    expect(normalizeDocumentEquation(equation)).toEqual(equation);

    const equationWithWordRunProperties = (wordRunProperties: unknown) =>
      ({
        version: 1,
        display: 'inline',
        children: [{ type: 'run', text: 'x', wordRunProperties }],
      }) as unknown as WorkDocumentEquation;
    for (const paragraphMarkAlwaysHidden of [true, false]) {
      expect(
        normalizeDocumentEquation(
          equationWithWordRunProperties({ paragraphMarkAlwaysHidden }),
        )?.children[0],
      ).toEqual({
        type: 'run',
        text: 'x',
        wordRunProperties: { paragraphMarkAlwaysHidden },
      });
    }
    expect(
      normalizeDocumentEquation(
        equationWithWordRunProperties({
          paragraphMarkAlwaysHidden: undefined,
        }),
      ),
    ).toEqual(simpleEquation('x'));
    const invalidModels = [
      { paragraphMarkAlwaysHidden: null },
      { paragraphMarkAlwaysHidden: 0 },
      { paragraphMarkAlwaysHidden: 1 },
      { paragraphMarkAlwaysHidden: 'false' },
      { paragraphMarkAlwaysHidden: {} },
    ];
    expect(
      invalidModels.map((properties) =>
        normalizeDocumentEquation(equationWithWordRunProperties(properties)),
      ),
    ).toEqual(invalidModels.map(() => null));

    const wordRun = (properties: string, namespace = WORD_NAMESPACE) =>
      `<m:r xmlns:w="${namespace}"><w:rPr>${properties}</w:rPr><m:t>x</m:t></m:r>`;
    for (const namespace of [WORD_NAMESPACE, STRICT_WORD_NAMESPACE]) {
      for (const [markup, paragraphMarkAlwaysHidden] of [
        ['<w:specVanish/>', true],
        ['<w:specVanish w:val="1"/>', true],
        ['<w:specVanish w:val=" true "/>', true],
        ['<w:specVanish w:val="ON"/>', true],
        ['<w:specVanish w:val="0"/>', false],
        ['<w:specVanish w:val=" false "/>', false],
        ['<w:specVanish w:val="OFF"/>', false],
      ] as const) {
        expect(
          inspectEquationModel(wordRun(markup, namespace))?.children[0],
        ).toEqual({
          type: 'run',
          text: 'x',
          wordRunProperties: { paragraphMarkAlwaysHidden },
        });
      }
    }
    expect(
      inspectEquationModel(
        wordRun(
          '<w:lang w:eastAsia="zh-CN"/><w:eastAsianLayout w:id="5" w:combine="0" w:vert="1"/><w:specVanish w:val="0"/>',
        ),
      )?.children[0],
    ).toEqual({
      type: 'run',
      text: 'x',
      wordRunProperties: {
        languages: { eastAsia: 'zh-CN' },
        eastAsianLayout: { id: 5, combine: false, vertical: true },
        paragraphMarkAlwaysHidden: false,
      },
    });

    const invalidMarkup = [
      wordRun('<w:specVanish w:val=""/>'),
      wordRun('<w:specVanish w:val="maybe"/>'),
      wordRun('<w:specVanish w:val="2"/>'),
      wordRun('<w:specVanish val="1"/>'),
      wordRun('<w:specVanish w:val="1" w:extra="semantic"/>'),
      wordRun(
        `<w:specVanish xmlns:r="${RELATIONSHIP_NAMESPACE}" w:val="1" r:id="rIdUnsafe"/>`,
      ),
      wordRun('<w:specVanish><w:b/></w:specVanish>'),
      wordRun('<w:specVanish>meaningful</w:specVanish>'),
      wordRun('<w:specVanish/><w:specVanish w:val="0"/>'),
      wordRun('<w:specVanish/><w:eastAsianLayout w:id="1"/>'),
      wordRun('<w:specVanish/><w:lang w:eastAsia="zh-CN"/>'),
      wordRun(`<v:specVanish xmlns:v="${VENDOR_NAMESPACE}" v:val="1"/>`),
      wordRun('<m:specVanish m:val="1"/>'),
      wordRun(WORD_2010_SCENE_3D),
    ];
    expect(invalidMarkup.map(inspectEquationBody)).toEqual(
      invalidMarkup.map(() => 'unsupported'),
    );

    const document = new DOMParser().parseFromString('', 'text/html');
    const preview = createDocumentEquationElement(document, equation);
    for (const text of [
      'always-hidden-paragraph-mark',
      'explicit-paragraph-mark-reset',
      'spec-vanish-with-hidden-run',
      'ordered-spec-vanish',
      'operator-spec-vanish',
    ]) {
      expect(preview.textContent).toContain(text);
    }
    expect(preview.outerHTML).not.toMatch(
      /(?:display\s*:\s*none|visibility\s*:\s*hidden|opacity\s*:\s*0(?:\D|$))/u,
    );
    const sanitized = new DOMParser().parseFromString(
      sanitizeDocumentPageChromeHtml(preview.outerHTML),
      'text/html',
    );
    expect(
      documentEquationFromElement(
        sanitized.body.querySelector<HTMLElement>(
          '[data-document-equation]',
        ) as HTMLElement,
      ),
    ).toEqual(equation);

    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    artifact.content.html = `<p>${preview.outerHTML}</p>`;
    const first = await createArtifactBlob(artifact);
    await expectNativeWordRunSpecVanish(first);
    const imported = await importOfficeFile(
      new File([first], 'word-run-spec-vanish.docx', { type: first.type }),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const importedDocument = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    expect(
      documentEquationFromElement(
        importedDocument.body.querySelector<HTMLElement>(
          '[data-document-equation]',
        ) as HTMLElement,
      ),
    ).toEqual(equation);
    expect(imported.compatibility.issues).not.toContainEqual(
      expect.objectContaining({ code: 'docx.equations.unsupported' }),
    );
    await expectNativeWordRunSpecVanish(await createArtifactBlob(imported));
  });

  test('preserves bounded Office 2010 text glow effects as native OMML metadata', async () => {
    const transforms = [
      { type: 'tint', value: 0 },
      { type: 'shade', value: 100_000 },
      { type: 'alpha', value: 50_000 },
      { type: 'hueMod', value: 200_000 },
      { type: 'saturation', value: -25_000 },
      { type: 'saturationOffset', value: 25_000 },
      { type: 'saturationModulation', value: 50_000 },
      { type: 'luminance', value: 40_000 },
      { type: 'luminanceOffset', value: -10_000 },
      { type: 'luminanceModulation', value: 75_000 },
      { type: 'alpha', value: 100_000 },
    ];
    const equation = {
      version: 1,
      display: 'inline',
      children: [
        {
          type: 'run',
          text: 'rgb-glow',
          wordRunProperties: {
            glow: {
              radiusEmus: 63_500,
              color: { type: 'rgb', value: '#a1b2c3', transforms },
            },
          },
        },
        {
          type: 'run',
          text: 'scheme-glow',
          wordRunProperties: {
            glow: {
              radiusEmus: 0,
              color: { type: 'scheme', value: 'placeholder' },
            },
          },
        },
        {
          type: 'run',
          text: 'default-radius-glow',
          wordRunProperties: {
            glow: { color: { type: 'scheme', value: 'accent2' } },
          },
        },
        {
          type: 'run',
          text: 'ordered-glow',
          wordRunProperties: {
            shadow: false,
            paragraphMarkAlwaysHidden: false,
            glow: { color: { type: 'rgb', value: '#112233' } },
          },
        },
        {
          type: 'nary',
          operator: '\u2211',
          limitLocation: 'underOver',
          controlProperties: {
            glow: {
              radiusEmus: 12_700,
              color: { type: 'scheme', value: 'hyperlink' },
            },
          },
          children: [{ type: 'run', text: 'operator-glow' }],
        },
      ],
    } as unknown as WorkDocumentEquation;
    expect(normalizeDocumentEquation(equation)).toEqual(equation);

    const equationWithGlow = (glow: unknown) =>
      ({
        version: 1,
        display: 'inline',
        children: [{ type: 'run', text: 'x', wordRunProperties: { glow } }],
      }) as unknown as WorkDocumentEquation;
    expect(
      normalizeDocumentEquation(
        equationWithGlow({
          color: { type: 'rgb', value: '#ABCDEF', transforms: [] },
        }),
      )?.children[0],
    ).toEqual({
      type: 'run',
      text: 'x',
      wordRunProperties: {
        glow: { color: { type: 'rgb', value: '#abcdef' } },
      },
    });
    expect(
      normalizeDocumentEquation({
        version: 1,
        display: 'inline',
        children: [
          {
            type: 'run',
            text: 'x',
            wordRunProperties: { glow: undefined },
          },
        ],
      } as unknown as WorkDocumentEquation),
    ).toEqual(simpleEquation('x'));
    const maximumTransforms = [
      { type: 'hueMod', value: 2_147_483_647 },
      { type: 'saturation', value: -2_147_483_648 },
      ...Array.from({ length: 61 }, () => ({
        type: 'alpha' as const,
        value: 100_000,
      })),
      { type: 'luminanceModulation', value: 2_147_483_647 },
    ];
    expect(
      normalizeDocumentEquation(
        equationWithGlow({
          radiusEmus: 2_147_483_647,
          color: {
            type: 'rgb',
            value: '#abcdef',
            transforms: maximumTransforms,
          },
        }),
      )?.children[0],
    ).toEqual({
      type: 'run',
      text: 'x',
      wordRunProperties: {
        glow: {
          radiusEmus: 2_147_483_647,
          color: {
            type: 'rgb',
            value: '#abcdef',
            transforms: maximumTransforms,
          },
        },
      },
    });
    const validRgb = { color: { type: 'rgb', value: '#abcdef' } };
    const invalidGlows = [
      null,
      false,
      {},
      { radiusEmus: 1 },
      { ...validRgb, extra: true },
      { ...validRgb, radiusEmus: -1 },
      { ...validRgb, radiusEmus: 1.5 },
      { ...validRgb, radiusEmus: 2_147_483_648 },
      { color: null },
      { color: {} },
      { color: { type: 'rgb' } },
      { color: { type: 'rgb', value: '#abc' } },
      { color: { type: 'rgb', value: 'ABCDEF' } },
      { color: { type: 'rgb', value: '#abcdef', extra: true } },
      { color: { type: 'scheme', value: 'none' } },
      { color: { type: 'scheme', value: 'bg1' } },
      { color: { type: 'unknown', value: '#abcdef' } },
      { color: { type: 'rgb', value: '#abcdef', transforms: null } },
      {
        color: {
          type: 'rgb',
          value: '#abcdef',
          transforms: Array.from({ length: 65 }, () => ({
            type: 'alpha',
            value: 50_000,
          })),
        },
      },
      ...[
        null,
        {},
        { type: 'unknown', value: 1 },
        { type: 'tint', value: -1 },
        { type: 'shade', value: 100_001 },
        { type: 'alpha', value: 1.5 },
        { type: 'alpha', value: '50000' },
        { type: 'hueMod', value: -1 },
        { type: 'hueMod', value: 2_147_483_648 },
        { type: 'saturation', value: -2_147_483_649 },
        { type: 'luminanceModulation', value: 2_147_483_648 },
        { type: 'tint', value: 50_000, extra: true },
      ].map((transform) => ({
        color: {
          type: 'rgb',
          value: '#abcdef',
          transforms: [transform],
        },
      })),
    ];
    expect(
      invalidGlows.map((glow) =>
        normalizeDocumentEquation(equationWithGlow(glow)),
      ),
    ).toEqual(invalidGlows.map(() => null));

    const wordRun = (properties: string, namespace = WORD_NAMESPACE) =>
      `<m:r xmlns:w="${namespace}" xmlns:w14="${WORD_2010_NAMESPACE}"><w:rPr>${properties}</w:rPr><m:t>x</m:t></m:r>`;
    const transformedRgb =
      '<w14:glow w14:rad=" +63500 "><w14:srgbClr w14:val="a1B2c3">' +
      '<w14:tint w14:val="0"/><w14:shade w14:val="100000"/>' +
      '<w14:alpha w14:val="50000"/><w14:hueMod w14:val="200000"/>' +
      '<w14:sat w14:val="-25000"/><w14:satOff w14:val="+25000"/>' +
      '<w14:satMod w14:val="50000"/><w14:lum w14:val="40000"/>' +
      '<w14:lumOff w14:val="-10000"/><w14:lumMod w14:val="75000"/>' +
      '<w14:alpha w14:val="100000"/></w14:srgbClr></w14:glow>';
    for (const namespace of [WORD_NAMESPACE, STRICT_WORD_NAMESPACE]) {
      expect(
        inspectEquationModel(wordRun(transformedRgb, namespace))?.children[0],
      ).toEqual({
        type: 'run',
        text: 'x',
        wordRunProperties: {
          glow: {
            radiusEmus: 63_500,
            color: { type: 'rgb', value: '#a1b2c3', transforms },
          },
        },
      });
    }
    expect(inspectEquationModel(wordRun(WORD_2010_GLOW))?.children[0]).toEqual({
      type: 'run',
      text: 'x',
      wordRunProperties: {
        glow: {
          radiusEmus: 63_500,
          color: { type: 'rgb', value: '#ffff00' },
        },
      },
    });
    const maximumTransformMarkup =
      '<w14:glow w14:rad="2147483647"><w14:srgbClr w14:val="ABCDEF">' +
      '<w14:hueMod w14:val="2147483647"/><w14:sat w14:val="-2147483648"/>' +
      Array.from({ length: 61 }, () => '<w14:alpha w14:val="100000"/>').join(
        '',
      ) +
      '<w14:lumMod w14:val="2147483647"/></w14:srgbClr></w14:glow>';
    expect(
      inspectEquationModel(wordRun(maximumTransformMarkup))?.children[0],
    ).toEqual({
      type: 'run',
      text: 'x',
      wordRunProperties: {
        glow: {
          radiusEmus: 2_147_483_647,
          color: {
            type: 'rgb',
            value: '#abcdef',
            transforms: maximumTransforms,
          },
        },
      },
    });
    const schemeColors = [
      ['bg1', 'background1'],
      ['tx1', 'text1'],
      ['bg2', 'background2'],
      ['tx2', 'text2'],
      ['accent1', 'accent1'],
      ['accent2', 'accent2'],
      ['accent3', 'accent3'],
      ['accent4', 'accent4'],
      ['accent5', 'accent5'],
      ['accent6', 'accent6'],
      ['hlink', 'hyperlink'],
      ['folHlink', 'followedHyperlink'],
      ['dk1', 'dark1'],
      ['lt1', 'light1'],
      ['dk2', 'dark2'],
      ['lt2', 'light2'],
      ['phClr', 'placeholder'],
    ] as const;
    for (const [source, value] of schemeColors) {
      expect(
        inspectEquationModel(
          wordRun(
            `<w14:glow w14:rad="0"><w14:schemeClr w14:val="${source}"/></w14:glow>`,
          ),
        )?.children[0],
      ).toEqual({
        type: 'run',
        text: 'x',
        wordRunProperties: {
          glow: { radiusEmus: 0, color: { type: 'scheme', value } },
        },
      });
    }
    expect(
      inspectEquationModel(
        wordRun(
          '<w:shadow w:val="0"/><w:specVanish w:val="0"/><w14:glow><w14:srgbClr w14:val="112233"/></w14:glow>',
        ),
      )?.children[0],
    ).toEqual({
      type: 'run',
      text: 'x',
      wordRunProperties: {
        shadow: false,
        paragraphMarkAlwaysHidden: false,
        glow: { color: { type: 'rgb', value: '#112233' } },
      },
    });

    const validColor = '<w14:srgbClr w14:val="A1B2C3"/>';
    const invalidMarkup = [
      wordRun('<w14:glow/>'),
      wordRun(
        `<w14:glow>${validColor}<w14:schemeClr w14:val="accent1"/></w14:glow>`,
      ),
      wordRun(
        `<w14:glow>${validColor}</w14:glow><w14:glow>${validColor}</w14:glow>`,
      ),
      wordRun(`<w14:glow>${validColor}</w14:glow><w:specVanish/>`),
      wordRun(`<w14:glow>${validColor}</w14:glow><w:lang w:val="en-US"/>`),
      wordRun(`<w:glow>${validColor}</w:glow>`),
      wordRun(`<v:glow xmlns:v="${VENDOR_NAMESPACE}">${validColor}</v:glow>`),
      wordRun(
        `<w14:glow xmlns:w14="${VENDOR_NAMESPACE}"><w14:srgbClr w14:val="A1B2C3"/></w14:glow>`,
      ),
      wordRun(`<w14:glow rad="1">${validColor}</w14:glow>`),
      wordRun(`<w14:glow w14:rad="-1">${validColor}</w14:glow>`),
      wordRun(`<w14:glow w14:rad="1.5">${validColor}</w14:glow>`),
      wordRun(`<w14:glow w14:rad="2147483648">${validColor}</w14:glow>`),
      wordRun(`<w14:glow w14:extra="semantic">${validColor}</w14:glow>`),
      wordRun(
        `<w14:glow xmlns:r="${RELATIONSHIP_NAMESPACE}" r:id="rIdUnsafe">${validColor}</w14:glow>`,
      ),
      wordRun('<w14:glow><w14:srgbClr/></w14:glow>'),
      wordRun('<w14:glow><w14:srgbClr val="A1B2C3"/></w14:glow>'),
      wordRun('<w14:glow><w14:srgbClr w14:val="ABC"/></w14:glow>'),
      wordRun(
        '<w14:glow><w14:srgbClr w14:val="A1B2C3" w14:extra="semantic"/></w14:glow>',
      ),
      wordRun('<w14:glow><w14:schemeClr w14:val="none"/></w14:glow>'),
      wordRun('<w14:glow><w14:schemeClr w14:val="ACCENT1"/></w14:glow>'),
      wordRun(
        '<w14:glow><w14:srgbClr w14:val="A1B2C3">meaningful</w14:srgbClr></w14:glow>',
      ),
      wordRun(
        '<w14:glow><w14:srgbClr w14:val="A1B2C3"><w14:alpha/></w14:srgbClr></w14:glow>',
      ),
      wordRun(
        '<w14:glow><w14:srgbClr w14:val="A1B2C3"><w14:alpha val="50000"/></w14:srgbClr></w14:glow>',
      ),
      wordRun(
        '<w14:glow><w14:srgbClr w14:val="A1B2C3"><w14:tint w14:val="-1"/></w14:srgbClr></w14:glow>',
      ),
      wordRun(
        '<w14:glow><w14:srgbClr w14:val="A1B2C3"><w14:shade w14:val="100001"/></w14:srgbClr></w14:glow>',
      ),
      wordRun(
        '<w14:glow><w14:srgbClr w14:val="A1B2C3"><w14:hueMod w14:val="-1"/></w14:srgbClr></w14:glow>',
      ),
      wordRun(
        '<w14:glow><w14:srgbClr w14:val="A1B2C3"><w14:lum w14:val="2147483648"/></w14:srgbClr></w14:glow>',
      ),
      wordRun(
        '<w14:glow><w14:srgbClr w14:val="A1B2C3">' +
          Array.from({ length: 65 }, () => '<w14:alpha w14:val="50000"/>').join(
            '',
          ) +
          '</w14:srgbClr></w14:glow>',
      ),
      wordRun(
        '<w14:glow><w14:srgbClr w14:val="A1B2C3"><w14:alpha w14:val="50000">meaningful</w14:alpha></w14:srgbClr></w14:glow>',
      ),
      wordRun(
        '<w14:glow><w14:srgbClr w14:val="A1B2C3"><w14:alpha w14:val="50000"><w14:tint w14:val="1"/></w14:alpha></w14:srgbClr></w14:glow>',
      ),
      wordRun(
        '<w14:glow><w14:srgbClr w14:val="A1B2C3"><v:alpha xmlns:v="urn:a3s:test" v:val="50000"/></w14:srgbClr></w14:glow>',
      ),
      wordRun(WORD_2010_SCENE_3D),
    ];
    expect(invalidMarkup.map(inspectEquationBody)).toEqual(
      invalidMarkup.map(() => 'unsupported'),
    );

    const document = new DOMParser().parseFromString('', 'text/html');
    const preview = createDocumentEquationElement(document, equation);
    for (const text of [
      'rgb-glow',
      'scheme-glow',
      'default-radius-glow',
      'ordered-glow',
      'operator-glow',
    ]) {
      expect(preview.textContent).toContain(text);
    }
    expect(preview.outerHTML).not.toMatch(/text-shadow/iu);

    const collisionMarker = '__A3S_GLOW_NAMESPACE_COLLISION__';
    const collisionArchive = new JSZip();
    collisionArchive.file(
      'word/document.xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${WORD_NAMESPACE}" xmlns:w14="urn:a3s:occupied-w14" xmlns:keep="urn:a3s:keep" xmlns:mc="${MARKUP_COMPATIBILITY_NAMESPACE}" mc:Ignorable="keep"><w:body><w:p><w:r><w:t>${collisionMarker}</w:t></w:r></w:p></w:body></w:document>`,
    );
    const collisionPatched = await patchDocxEquations(
      await collisionArchive.generateAsync({ type: 'arraybuffer' }),
      [
        {
          marker: collisionMarker,
          equation: equationWithGlow({
            radiusEmus: 1,
            color: { type: 'rgb', value: '#010203' },
          }),
        },
      ],
    );
    const collisionOutput = await xmlEntry(
      await JSZip.loadAsync(collisionPatched),
      'word/document.xml',
    );
    const collisionRoot = collisionOutput.documentElement;
    expect(xmlNamespaceUri(collisionRoot, 'w14')).toBe('urn:a3s:occupied-w14');
    const generatedWord2010Prefix = Array.from(collisionRoot.attributes)
      .find(
        (attribute) =>
          attribute.value === WORD_2010_NAMESPACE &&
          attribute.name.startsWith('xmlns:'),
      )
      ?.name.slice('xmlns:'.length);
    expect(generatedWord2010Prefix).toBeDefined();
    expect(generatedWord2010Prefix).not.toBe('w14');
    const ignorableTokens = Array.from(collisionRoot.attributes)
      .find(
        (attribute) =>
          xmlAttributeNamespace(collisionRoot, attribute) ===
            MARKUP_COMPATIBILITY_NAMESPACE &&
          xmlAttributeLocalName(attribute) === 'Ignorable',
      )
      ?.value.trim()
      .split(/\s+/u);
    expect(ignorableTokens).toEqual(['keep', generatedWord2010Prefix]);
    expect(
      inspectDocxEquation(descendants(collisionOutput, 'oMath')[0]),
    ).toMatchObject({
      status: 'supported',
      equation: equationWithGlow({
        radiusEmus: 1,
        color: { type: 'rgb', value: '#010203' },
      }),
    });
    const sanitized = new DOMParser().parseFromString(
      sanitizeDocumentPageChromeHtml(preview.outerHTML),
      'text/html',
    );
    expect(
      documentEquationFromElement(
        sanitized.body.querySelector<HTMLElement>(
          '[data-document-equation]',
        ) as HTMLElement,
      ),
    ).toEqual(equation);

    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    artifact.content.html = `<p>${preview.outerHTML}</p>`;
    const first = await createArtifactBlob(artifact);
    await expectNativeWordRunGlows(first);
    const imported = await importOfficeFile(
      new File([first], 'word-run-glow.docx', { type: first.type }),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const importedDocument = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    expect(
      documentEquationFromElement(
        importedDocument.body.querySelector<HTMLElement>(
          '[data-document-equation]',
        ) as HTMLElement,
      ),
    ).toEqual(equation);
    expect(imported.compatibility.issues).not.toContainEqual(
      expect.objectContaining({ code: 'docx.equations.unsupported' }),
    );
    await expectNativeWordRunGlows(await createArtifactBlob(imported));
  });

  test('preserves bounded Office 2010 text shadow effects as native OMML metadata', async () => {
    const transforms = [
      { type: 'alpha', value: 50_000 },
      { type: 'saturationOffset', value: -25_000 },
      { type: 'alpha', value: 100_000 },
    ];
    const maximumDirection = 21_599_999 / 60_000;
    const minimumScale = -2_147_483_648 / 1_000;
    const maximumScale = 2_147_483_647 / 1_000;
    const minimumSkew = -5_399_999 / 60_000;
    const maximumSkew = 5_399_999 / 60_000;
    const boundaryEffect = {
      blurRadiusEmus: 2_147_483_647,
      distanceEmus: 2_147_483_647,
      directionDegrees: maximumDirection,
      horizontalScalePercent: minimumScale,
      verticalScalePercent: maximumScale,
      horizontalSkewDegrees: minimumSkew,
      verticalSkewDegrees: maximumSkew,
      alignment: 'topLeft',
      color: { type: 'rgb', value: '#abcdef' },
    };
    const equation = {
      version: 1,
      display: 'inline',
      children: [
        {
          type: 'run',
          text: 'rgb-shadow-effect',
          wordRunProperties: {
            shadowEffect: {
              blurRadiusEmus: 63_500,
              distanceEmus: 12_700,
              directionDegrees: 45,
              horizontalScalePercent: 100,
              verticalScalePercent: -50,
              horizontalSkewDegrees: -1,
              verticalSkewDegrees: 2,
              alignment: 'bottomRight',
              color: { type: 'rgb', value: '#a1b2c3', transforms },
            },
          },
        },
        {
          type: 'run',
          text: 'scheme-shadow-effect',
          wordRunProperties: {
            shadowEffect: {
              blurRadiusEmus: 0,
              distanceEmus: 0,
              directionDegrees: 0,
              horizontalScalePercent: 0,
              verticalScalePercent: 0,
              horizontalSkewDegrees: 0,
              verticalSkewDegrees: 0,
              alignment: 'none',
              color: { type: 'scheme', value: 'placeholder' },
            },
          },
        },
        {
          type: 'run',
          text: 'default-shadow-effect',
          wordRunProperties: {
            shadowEffect: {
              color: { type: 'scheme', value: 'accent3' },
            },
          },
        },
        {
          type: 'run',
          text: 'ordered-shadow-effect',
          wordRunProperties: {
            shadow: false,
            paragraphMarkAlwaysHidden: false,
            glow: { color: { type: 'rgb', value: '#112233' } },
            shadowEffect: { color: { type: 'rgb', value: '#445566' } },
          },
        },
        {
          type: 'run',
          text: 'boundary-shadow-effect',
          wordRunProperties: { shadowEffect: boundaryEffect },
        },
        {
          type: 'nary',
          operator: '\u2211',
          limitLocation: 'underOver',
          controlProperties: {
            shadowEffect: {
              blurRadiusEmus: 25_400,
              alignment: 'center',
              color: { type: 'scheme', value: 'hyperlink' },
            },
          },
          children: [{ type: 'run', text: 'operator-shadow-effect' }],
        },
      ],
    } as unknown as WorkDocumentEquation;
    expect(normalizeDocumentEquation(equation)).toEqual(equation);

    const equationWithShadowEffect = (shadowEffect: unknown) =>
      ({
        version: 1,
        display: 'inline',
        children: [
          { type: 'run', text: 'x', wordRunProperties: { shadowEffect } },
        ],
      }) as unknown as WorkDocumentEquation;
    expect(
      normalizeDocumentEquation({
        version: 1,
        display: 'inline',
        children: [
          {
            type: 'run',
            text: 'x',
            wordRunProperties: { shadowEffect: undefined },
          },
        ],
      } as unknown as WorkDocumentEquation),
    ).toEqual(simpleEquation('x'));
    expect(
      normalizeDocumentEquation(equationWithShadowEffect(boundaryEffect))
        ?.children[0],
    ).toEqual({
      type: 'run',
      text: 'x',
      wordRunProperties: { shadowEffect: boundaryEffect },
    });
    const validEffect = { color: { type: 'rgb', value: '#abcdef' } };
    const invalidEffects = [
      null,
      false,
      {},
      { blurRadiusEmus: 1 },
      { ...validEffect, extra: true },
      { ...validEffect, blurRadiusEmus: -1 },
      { ...validEffect, blurRadiusEmus: 1.5 },
      { ...validEffect, blurRadiusEmus: 2_147_483_648 },
      { ...validEffect, distanceEmus: -1 },
      { ...validEffect, distanceEmus: 1.5 },
      { ...validEffect, distanceEmus: 2_147_483_648 },
      { ...validEffect, directionDegrees: -1 / 60_000 },
      { ...validEffect, directionDegrees: 360 },
      { ...validEffect, directionDegrees: 1 / 120_000 },
      { ...validEffect, horizontalScalePercent: minimumScale - 0.001 },
      { ...validEffect, verticalScalePercent: maximumScale + 0.001 },
      { ...validEffect, horizontalScalePercent: 0.0005 },
      { ...validEffect, horizontalSkewDegrees: -90 },
      { ...validEffect, verticalSkewDegrees: 90 },
      { ...validEffect, horizontalSkewDegrees: 1 / 120_000 },
      { ...validEffect, alignment: 'middle' },
      { color: { type: 'scheme', value: 'none' } },
      {
        color: {
          type: 'rgb',
          value: '#abcdef',
          transforms: Array.from({ length: 65 }, () => ({
            type: 'alpha',
            value: 50_000,
          })),
        },
      },
    ];
    expect(
      invalidEffects.map((effect) =>
        normalizeDocumentEquation(equationWithShadowEffect(effect)),
      ),
    ).toEqual(invalidEffects.map(() => null));

    const wordRun = (properties: string, namespace = WORD_NAMESPACE) =>
      `<m:r xmlns:w="${namespace}" xmlns:w14="${WORD_2010_NAMESPACE}"><w:rPr>${properties}</w:rPr><m:t>x</m:t></m:r>`;
    const transformedRgb =
      '<w14:shadow w14:blurRad=" +63500 " w14:dist="12700" ' +
      'w14:dir="2700000" w14:sx="100000" w14:sy="-50000" ' +
      'w14:kx="-60000" w14:ky="120000" w14:algn="br">' +
      '<w14:srgbClr w14:val="a1B2c3"><w14:alpha w14:val="50000"/>' +
      '<w14:satOff w14:val="-25000"/><w14:alpha w14:val="100000"/>' +
      '</w14:srgbClr></w14:shadow>';
    for (const namespace of [WORD_NAMESPACE, STRICT_WORD_NAMESPACE]) {
      expect(
        inspectEquationModel(wordRun(transformedRgb, namespace))?.children[0],
      ).toEqual({
        type: 'run',
        text: 'x',
        wordRunProperties: {
          shadowEffect: {
            blurRadiusEmus: 63_500,
            distanceEmus: 12_700,
            directionDegrees: 45,
            horizontalScalePercent: 100,
            verticalScalePercent: -50,
            horizontalSkewDegrees: -1,
            verticalSkewDegrees: 2,
            alignment: 'bottomRight',
            color: { type: 'rgb', value: '#a1b2c3', transforms },
          },
        },
      });
    }
    expect(
      inspectEquationModel(wordRun(WORD_2010_SHADOW))?.children[0],
    ).toEqual({
      type: 'run',
      text: 'x',
      wordRunProperties: {
        shadowEffect: { color: { type: 'rgb', value: '#000000' } },
      },
    });
    const boundaryMarkup =
      '<w14:shadow w14:blurRad="2147483647" w14:dist="2147483647" ' +
      'w14:dir="21599999" w14:sx="-2147483648" w14:sy="2147483647" ' +
      'w14:kx="-5399999" w14:ky="5399999" w14:algn="tl">' +
      '<w14:srgbClr w14:val="ABCDEF"/></w14:shadow>';
    expect(inspectEquationModel(wordRun(boundaryMarkup))?.children[0]).toEqual({
      type: 'run',
      text: 'x',
      wordRunProperties: { shadowEffect: boundaryEffect },
    });
    const alignments = [
      ['none', 'none'],
      ['tl', 'topLeft'],
      ['t', 'top'],
      ['tr', 'topRight'],
      ['l', 'left'],
      ['ctr', 'center'],
      ['r', 'right'],
      ['bl', 'bottomLeft'],
      ['b', 'bottom'],
      ['br', 'bottomRight'],
    ] as const;
    for (const [source, alignment] of alignments) {
      expect(
        inspectEquationModel(
          wordRun(
            `<w14:shadow w14:algn="${source}"><w14:schemeClr w14:val="accent1"/></w14:shadow>`,
          ),
        )?.children[0],
      ).toEqual({
        type: 'run',
        text: 'x',
        wordRunProperties: {
          shadowEffect: {
            alignment,
            color: { type: 'scheme', value: 'accent1' },
          },
        },
      });
    }
    expect(
      inspectEquationModel(
        wordRun(
          '<w:shadow w:val="0"/><w:specVanish w:val="0"/>' +
            '<w14:glow><w14:srgbClr w14:val="112233"/></w14:glow>' +
            '<w14:shadow><w14:srgbClr w14:val="445566"/></w14:shadow>',
        ),
      )?.children[0],
    ).toEqual({
      type: 'run',
      text: 'x',
      wordRunProperties: {
        shadow: false,
        paragraphMarkAlwaysHidden: false,
        glow: { color: { type: 'rgb', value: '#112233' } },
        shadowEffect: { color: { type: 'rgb', value: '#445566' } },
      },
    });

    const validColor = '<w14:srgbClr w14:val="A1B2C3"/>';
    const invalidMarkup = [
      wordRun('<w14:shadow/>'),
      wordRun(
        `<w14:shadow>${validColor}<w14:schemeClr w14:val="accent1"/></w14:shadow>`,
      ),
      wordRun(
        `<w14:shadow>${validColor}</w14:shadow><w14:shadow>${validColor}</w14:shadow>`,
      ),
      wordRun(
        `<w14:shadow>${validColor}</w14:shadow><w14:glow>${validColor}</w14:glow>`,
      ),
      wordRun(`<w14:shadow>${validColor}</w14:shadow><w:specVanish/>`),
      wordRun(`<w:shadow>${validColor}</w:shadow>`),
      wordRun(
        `<v:shadow xmlns:v="${VENDOR_NAMESPACE}">${validColor}</v:shadow>`,
      ),
      wordRun(
        `<w14:shadow xmlns:w14="${VENDOR_NAMESPACE}"><w14:srgbClr w14:val="A1B2C3"/></w14:shadow>`,
      ),
      wordRun(`<w14:shadow blurRad="1">${validColor}</w14:shadow>`),
      wordRun(`<w14:shadow w14:blurRad="-1">${validColor}</w14:shadow>`),
      wordRun(`<w14:shadow w14:blurRad="1.5">${validColor}</w14:shadow>`),
      wordRun(
        `<w14:shadow w14:blurRad="2147483648">${validColor}</w14:shadow>`,
      ),
      wordRun(`<w14:shadow w14:dist="-1">${validColor}</w14:shadow>`),
      wordRun(`<w14:shadow w14:dist="2147483648">${validColor}</w14:shadow>`),
      wordRun(`<w14:shadow w14:dir="-1">${validColor}</w14:shadow>`),
      wordRun(`<w14:shadow w14:dir="21600000">${validColor}</w14:shadow>`),
      wordRun(`<w14:shadow w14:sx="-2147483649">${validColor}</w14:shadow>`),
      wordRun(`<w14:shadow w14:sy="2147483648">${validColor}</w14:shadow>`),
      wordRun(`<w14:shadow w14:sx="1.5">${validColor}</w14:shadow>`),
      wordRun(`<w14:shadow w14:kx="-5400000">${validColor}</w14:shadow>`),
      wordRun(`<w14:shadow w14:ky="5400000">${validColor}</w14:shadow>`),
      wordRun(`<w14:shadow w14:algn="middle">${validColor}</w14:shadow>`),
      wordRun(`<w14:shadow w14:algn="CTR">${validColor}</w14:shadow>`),
      wordRun(`<w14:shadow w14:extra="semantic">${validColor}</w14:shadow>`),
      wordRun(
        `<w14:shadow xmlns:r="${RELATIONSHIP_NAMESPACE}" r:id="rIdUnsafe">${validColor}</w14:shadow>`,
      ),
      wordRun(`<w14:shadow>${validColor}meaningful</w14:shadow>`),
      wordRun(
        '<w14:shadow><v:srgbClr xmlns:v="urn:a3s:test" v:val="A1B2C3"/></w14:shadow>',
      ),
      wordRun(WORD_2010_SCENE_3D),
    ];
    expect(invalidMarkup.map(inspectEquationBody)).toEqual(
      invalidMarkup.map(() => 'unsupported'),
    );

    const document = new DOMParser().parseFromString('', 'text/html');
    const preview = createDocumentEquationElement(document, equation);
    for (const text of [
      'rgb-shadow-effect',
      'scheme-shadow-effect',
      'default-shadow-effect',
      'ordered-shadow-effect',
      'boundary-shadow-effect',
      'operator-shadow-effect',
    ]) {
      expect(preview.textContent).toContain(text);
    }
    expect(preview.outerHTML).not.toMatch(/text-shadow/iu);
    const sanitized = new DOMParser().parseFromString(
      sanitizeDocumentPageChromeHtml(preview.outerHTML),
      'text/html',
    );
    expect(
      documentEquationFromElement(
        sanitized.body.querySelector<HTMLElement>(
          '[data-document-equation]',
        ) as HTMLElement,
      ),
    ).toEqual(equation);

    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    artifact.content.html = `<p>${preview.outerHTML}</p>`;
    const first = await createArtifactBlob(artifact);
    await expectNativeWordRunShadowEffects(first);
    const imported = await importOfficeFile(
      new File([first], 'word-run-shadow-effect.docx', { type: first.type }),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const importedDocument = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    expect(
      documentEquationFromElement(
        importedDocument.body.querySelector<HTMLElement>(
          '[data-document-equation]',
        ) as HTMLElement,
      ),
    ).toEqual(equation);
    expect(imported.compatibility.issues).not.toContainEqual(
      expect.objectContaining({ code: 'docx.equations.unsupported' }),
    );
    await expectNativeWordRunShadowEffects(await createArtifactBlob(imported));
  });

  test('preserves bounded Office 2010 text reflection effects as native OMML metadata', async () => {
    const maximumDirection = 21_599_999 / 60_000;
    const minimumScale = -2_147_483_648 / 1_000;
    const maximumScale = 2_147_483_647 / 1_000;
    const minimumSkew = -5_399_999 / 60_000;
    const maximumSkew = 5_399_999 / 60_000;
    const boundaryEffect = {
      blurRadiusEmus: 2_147_483_647,
      startOpacityPercent: 0,
      startPositionPercent: 100,
      endOpacityPercent: 100,
      endPositionPercent: 0,
      distanceEmus: 2_147_483_647,
      directionDegrees: maximumDirection,
      fadeDirectionDegrees: maximumDirection,
      horizontalScalePercent: minimumScale,
      verticalScalePercent: maximumScale,
      horizontalSkewDegrees: minimumSkew,
      verticalSkewDegrees: maximumSkew,
      alignment: 'topLeft',
    };
    const equation = {
      version: 1,
      display: 'inline',
      children: [
        {
          type: 'run',
          text: 'full-reflection-effect',
          wordRunProperties: {
            reflectionEffect: {
              blurRadiusEmus: 63_500,
              startOpacityPercent: 50,
              startPositionPercent: 0,
              endOpacityPercent: 10,
              endPositionPercent: 100,
              distanceEmus: 12_700,
              directionDegrees: 90,
              fadeDirectionDegrees: 270,
              horizontalScalePercent: 100,
              verticalScalePercent: -50,
              horizontalSkewDegrees: -1,
              verticalSkewDegrees: 2,
              alignment: 'bottom',
            },
          },
        },
        {
          type: 'run',
          text: 'zero-reflection-effect',
          wordRunProperties: {
            reflectionEffect: {
              blurRadiusEmus: 0,
              startOpacityPercent: 0,
              startPositionPercent: 0,
              endOpacityPercent: 0,
              endPositionPercent: 0,
              distanceEmus: 0,
              directionDegrees: 0,
              fadeDirectionDegrees: 0,
              horizontalScalePercent: 0,
              verticalScalePercent: 0,
              horizontalSkewDegrees: 0,
              verticalSkewDegrees: 0,
              alignment: 'none',
            },
          },
        },
        {
          type: 'run',
          text: 'default-reflection-effect',
          wordRunProperties: { reflectionEffect: {} },
        },
        {
          type: 'run',
          text: 'ordered-reflection-effect',
          wordRunProperties: {
            shadow: false,
            paragraphMarkAlwaysHidden: false,
            glow: { color: { type: 'rgb', value: '#112233' } },
            shadowEffect: { color: { type: 'rgb', value: '#445566' } },
            reflectionEffect: { startOpacityPercent: 25 },
          },
        },
        {
          type: 'run',
          text: 'boundary-reflection-effect',
          wordRunProperties: { reflectionEffect: boundaryEffect },
        },
        {
          type: 'nary',
          operator: '\u2211',
          limitLocation: 'underOver',
          controlProperties: {
            reflectionEffect: {
              blurRadiusEmus: 25_400,
              endOpacityPercent: 75,
              alignment: 'center',
            },
          },
          children: [{ type: 'run', text: 'operator-reflection-effect' }],
        },
      ],
    } as unknown as WorkDocumentEquation;
    expect(normalizeDocumentEquation(equation)).toEqual(equation);

    const equationWithReflectionEffect = (reflectionEffect: unknown) =>
      ({
        version: 1,
        display: 'inline',
        children: [
          { type: 'run', text: 'x', wordRunProperties: { reflectionEffect } },
        ],
      }) as unknown as WorkDocumentEquation;
    expect(
      normalizeDocumentEquation({
        version: 1,
        display: 'inline',
        children: [
          {
            type: 'run',
            text: 'x',
            wordRunProperties: { reflectionEffect: undefined },
          },
        ],
      } as unknown as WorkDocumentEquation),
    ).toEqual(simpleEquation('x'));
    expect(
      normalizeDocumentEquation(equationWithReflectionEffect(boundaryEffect))
        ?.children[0],
    ).toEqual({
      type: 'run',
      text: 'x',
      wordRunProperties: { reflectionEffect: boundaryEffect },
    });
    const invalidEffects = [
      null,
      false,
      { extra: true },
      { blurRadiusEmus: -1 },
      { blurRadiusEmus: 1.5 },
      { blurRadiusEmus: 2_147_483_648 },
      { distanceEmus: -1 },
      { distanceEmus: 1.5 },
      { distanceEmus: 2_147_483_648 },
      { startOpacityPercent: -0.001 },
      { startOpacityPercent: 100.001 },
      { startOpacityPercent: 0.0005 },
      { startPositionPercent: Number.NaN },
      { endOpacityPercent: Number.POSITIVE_INFINITY },
      { endPositionPercent: '50' },
      { directionDegrees: -1 / 60_000 },
      { directionDegrees: 360 },
      { directionDegrees: 1 / 120_000 },
      { fadeDirectionDegrees: -1 / 60_000 },
      { fadeDirectionDegrees: 360 },
      { horizontalScalePercent: minimumScale - 0.001 },
      { verticalScalePercent: maximumScale + 0.001 },
      { horizontalScalePercent: 0.0005 },
      { horizontalSkewDegrees: -90 },
      { verticalSkewDegrees: 90 },
      { horizontalSkewDegrees: 1 / 120_000 },
      { alignment: 'middle' },
    ];
    expect(
      invalidEffects.map((effect) =>
        normalizeDocumentEquation(equationWithReflectionEffect(effect)),
      ),
    ).toEqual(invalidEffects.map(() => null));

    const wordRun = (properties: string, namespace = WORD_NAMESPACE) =>
      `<m:r xmlns:w="${namespace}" xmlns:w14="${WORD_2010_NAMESPACE}"><w:rPr>${properties}</w:rPr><m:t>x</m:t></m:r>`;
    const fullMarkup =
      '<w14:reflection w14:blurRad=" +63500 " w14:stA="50000" ' +
      'w14:stPos="0" w14:endA="10000" w14:endPos="100000" ' +
      'w14:dist="12700" w14:dir="5400000" w14:fadeDir="16200000" ' +
      'w14:sx="100000" w14:sy="-50000" w14:kx="-60000" ' +
      'w14:ky="120000" w14:algn="b"/>';
    for (const namespace of [WORD_NAMESPACE, STRICT_WORD_NAMESPACE]) {
      expect(
        inspectEquationModel(wordRun(fullMarkup, namespace))?.children[0],
      ).toEqual({
        type: 'run',
        text: 'x',
        wordRunProperties: {
          reflectionEffect: {
            blurRadiusEmus: 63_500,
            startOpacityPercent: 50,
            startPositionPercent: 0,
            endOpacityPercent: 10,
            endPositionPercent: 100,
            distanceEmus: 12_700,
            directionDegrees: 90,
            fadeDirectionDegrees: 270,
            horizontalScalePercent: 100,
            verticalScalePercent: -50,
            horizontalSkewDegrees: -1,
            verticalSkewDegrees: 2,
            alignment: 'bottom',
          },
        },
      });
    }
    expect(
      inspectEquationModel(wordRun(WORD_2010_REFLECTION))?.children[0],
    ).toEqual({
      type: 'run',
      text: 'x',
      wordRunProperties: { reflectionEffect: {} },
    });
    const boundaryMarkup =
      '<w14:reflection w14:blurRad="2147483647" w14:stA="0" ' +
      'w14:stPos="100000" w14:endA="100000" w14:endPos="0" ' +
      'w14:dist="2147483647" w14:dir="21599999" ' +
      'w14:fadeDir="21599999" w14:sx="-2147483648" ' +
      'w14:sy="2147483647" w14:kx="-5399999" w14:ky="5399999" ' +
      'w14:algn="tl"/>';
    expect(inspectEquationModel(wordRun(boundaryMarkup))?.children[0]).toEqual({
      type: 'run',
      text: 'x',
      wordRunProperties: { reflectionEffect: boundaryEffect },
    });
    const alignments = [
      ['none', 'none'],
      ['tl', 'topLeft'],
      ['t', 'top'],
      ['tr', 'topRight'],
      ['l', 'left'],
      ['ctr', 'center'],
      ['r', 'right'],
      ['bl', 'bottomLeft'],
      ['b', 'bottom'],
      ['br', 'bottomRight'],
    ] as const;
    for (const [source, alignment] of alignments) {
      expect(
        inspectEquationModel(wordRun(`<w14:reflection w14:algn="${source}"/>`))
          ?.children[0],
      ).toEqual({
        type: 'run',
        text: 'x',
        wordRunProperties: { reflectionEffect: { alignment } },
      });
    }
    expect(
      inspectEquationModel(
        wordRun(
          '<w:shadow w:val="0"/><w:specVanish w:val="0"/>' +
            '<w14:glow><w14:srgbClr w14:val="112233"/></w14:glow>' +
            '<w14:shadow><w14:srgbClr w14:val="445566"/></w14:shadow>' +
            '<w14:reflection w14:stA="25000"/>',
        ),
      )?.children[0],
    ).toEqual({
      type: 'run',
      text: 'x',
      wordRunProperties: {
        shadow: false,
        paragraphMarkAlwaysHidden: false,
        glow: { color: { type: 'rgb', value: '#112233' } },
        shadowEffect: { color: { type: 'rgb', value: '#445566' } },
        reflectionEffect: { startOpacityPercent: 25 },
      },
    });

    const invalidMarkup = [
      wordRun('<w14:reflection/><w14:reflection/>'),
      wordRun(
        `<w14:reflection/><w14:shadow><w14:srgbClr w14:val="000000"/></w14:shadow>`,
      ),
      wordRun(
        '<w14:reflection/><w14:glow><w14:srgbClr w14:val="000000"/></w14:glow>',
      ),
      wordRun('<w14:reflection/><w:specVanish/>'),
      wordRun('<w:reflection/>'),
      wordRun(`<v:reflection xmlns:v="${VENDOR_NAMESPACE}"/>`),
      wordRun(
        `<w14:reflection xmlns:w14="${VENDOR_NAMESPACE}" w14:blurRad="1"/>`,
      ),
      wordRun('<w14:reflection blurRad="1"/>'),
      wordRun('<w14:reflection w14:blurRad="-1"/>'),
      wordRun('<w14:reflection w14:blurRad="1.5"/>'),
      wordRun('<w14:reflection w14:blurRad="2147483648"/>'),
      wordRun('<w14:reflection w14:stA="-1"/>'),
      wordRun('<w14:reflection w14:stA="100001"/>'),
      wordRun('<w14:reflection w14:stPos="1.5"/>'),
      wordRun('<w14:reflection w14:endA="-1"/>'),
      wordRun('<w14:reflection w14:endPos="100001"/>'),
      wordRun('<w14:reflection w14:dist="-1"/>'),
      wordRun('<w14:reflection w14:dist="2147483648"/>'),
      wordRun('<w14:reflection w14:dir="-1"/>'),
      wordRun('<w14:reflection w14:dir="21600000"/>'),
      wordRun('<w14:reflection w14:fadeDir="-1"/>'),
      wordRun('<w14:reflection w14:fadeDir="21600000"/>'),
      wordRun('<w14:reflection w14:sx="-2147483649"/>'),
      wordRun('<w14:reflection w14:sy="2147483648"/>'),
      wordRun('<w14:reflection w14:sx="1.5"/>'),
      wordRun('<w14:reflection w14:kx="-5400000"/>'),
      wordRun('<w14:reflection w14:ky="5400000"/>'),
      wordRun('<w14:reflection w14:algn="middle"/>'),
      wordRun('<w14:reflection w14:algn="CTR"/>'),
      wordRun('<w14:reflection w14:extra="semantic"/>'),
      wordRun(
        `<w14:reflection xmlns:r="${RELATIONSHIP_NAMESPACE}" r:id="rIdUnsafe"/>`,
      ),
      wordRun('<w14:reflection>meaningful</w14:reflection>'),
      wordRun('<w14:reflection><w14:alpha w14:val="50000"/></w14:reflection>'),
      wordRun(WORD_2010_SCENE_3D),
    ];
    expect(invalidMarkup.map(inspectEquationBody)).toEqual(
      invalidMarkup.map(() => 'unsupported'),
    );

    const document = new DOMParser().parseFromString('', 'text/html');
    const preview = createDocumentEquationElement(document, equation);
    for (const text of [
      'full-reflection-effect',
      'zero-reflection-effect',
      'default-reflection-effect',
      'ordered-reflection-effect',
      'boundary-reflection-effect',
      'operator-reflection-effect',
    ]) {
      expect(preview.textContent).toContain(text);
    }
    expect(preview.querySelector('math')?.outerHTML).not.toMatch(
      /box-reflect|opacity|transform/iu,
    );
    const sanitized = new DOMParser().parseFromString(
      sanitizeDocumentPageChromeHtml(preview.outerHTML),
      'text/html',
    );
    expect(
      documentEquationFromElement(
        sanitized.body.querySelector<HTMLElement>(
          '[data-document-equation]',
        ) as HTMLElement,
      ),
    ).toEqual(equation);

    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    artifact.content.html = `<p>${preview.outerHTML}</p>`;
    const first = await createArtifactBlob(artifact);
    await expectNativeWordRunReflectionEffects(first);
    const imported = await importOfficeFile(
      new File([first], 'word-run-reflection-effect.docx', {
        type: first.type,
      }),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const importedDocument = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    expect(
      documentEquationFromElement(
        importedDocument.body.querySelector<HTMLElement>(
          '[data-document-equation]',
        ) as HTMLElement,
      ),
    ).toEqual(equation);
    expect(imported.compatibility.issues).not.toContainEqual(
      expect.objectContaining({ code: 'docx.equations.unsupported' }),
    );
    await expectNativeWordRunReflectionEffects(
      await createArtifactBlob(imported),
    );
  });

  test('preserves bounded Office 2010 text outline effects as native OMML metadata', async () => {
    const transforms = [
      { type: 'tint', value: 0 },
      { type: 'alpha', value: 50_000 },
      { type: 'luminanceModulation', value: 75_000 },
    ];
    const fullEffect = {
      widthEmus: 12_700,
      cap: 'square',
      compound: 'double',
      alignment: 'inset',
      fill: {
        type: 'gradient',
        stops: [
          {
            positionPercent: 0,
            color: { type: 'rgb', value: '#a1b2c3', transforms },
          },
          {
            positionPercent: 100,
            color: { type: 'scheme', value: 'accent3' },
          },
        ],
        shade: { type: 'linear', angleDegrees: 90, scaled: true },
      },
      dash: { preset: 'longDashDotDot' },
      join: { type: 'miter', limitPercent: 250 },
    };
    const boundaryPositions = [0, 0.001, 10, 20, 30, 40, 50, 60, 99.999, 100];
    const boundaryStops = boundaryPositions.map((positionPercent, index) => ({
      positionPercent,
      color: {
        type: 'rgb',
        value: `#${index.toString(16).padStart(6, '0')}`,
      },
    }));
    const minimumPercentage = -2_147_483_648 / 1_000;
    const maximumPercentage = 2_147_483_647 / 1_000;
    const boundaryEffect = {
      widthEmus: 20_116_800,
      cap: 'round',
      compound: 'triple',
      alignment: 'inset',
      fill: {
        type: 'gradient',
        stops: boundaryStops,
        shade: {
          type: 'path',
          path: 'rectangle',
          fillToRectangle: {
            leftPercent: minimumPercentage,
            topPercent: maximumPercentage,
            rightPercent: 0,
            bottomPercent: 0.001,
          },
        },
      },
      dash: {},
      join: { type: 'miter', limitPercent: maximumPercentage },
    };
    const equation = {
      version: 1,
      display: 'inline',
      children: [
        {
          type: 'run',
          text: 'full-text-outline-effect',
          wordRunProperties: { textOutlineEffect: fullEffect },
        },
        {
          type: 'run',
          text: 'explicit-default-text-outline-effect',
          wordRunProperties: {
            textOutlineEffect: {
              widthEmus: 0,
              cap: 'flat',
              compound: 'single',
              alignment: 'center',
              fill: { type: 'none' },
              dash: { preset: 'solid' },
              join: { type: 'bevel' },
            },
          },
        },
        {
          type: 'run',
          text: 'empty-text-outline-effect',
          wordRunProperties: { textOutlineEffect: {} },
        },
        {
          type: 'run',
          text: 'empty-components-text-outline-effect',
          wordRunProperties: {
            textOutlineEffect: {
              fill: { type: 'solid' },
              dash: {},
              join: { type: 'miter' },
            },
          },
        },
        {
          type: 'run',
          text: 'ordered-text-outline-effect',
          wordRunProperties: {
            shadow: false,
            paragraphMarkAlwaysHidden: false,
            glow: { color: { type: 'rgb', value: '#112233' } },
            shadowEffect: { color: { type: 'rgb', value: '#445566' } },
            reflectionEffect: { startOpacityPercent: 25 },
            textOutlineEffect: { fill: { type: 'none' } },
          },
        },
        {
          type: 'run',
          text: 'boundary-text-outline-effect',
          wordRunProperties: { textOutlineEffect: boundaryEffect },
        },
        {
          type: 'nary',
          operator: '\u2211',
          limitLocation: 'underOver',
          controlProperties: {
            textOutlineEffect: {
              widthEmus: 25_400,
              fill: {
                type: 'solid',
                color: { type: 'scheme', value: 'hyperlink' },
              },
              join: { type: 'round' },
            },
          },
          children: [{ type: 'run', text: 'operator-text-outline-effect' }],
        },
      ],
    } as unknown as WorkDocumentEquation;
    expect(normalizeDocumentEquation(equation)).toEqual(equation);

    const equationWithTextOutlineEffect = (textOutlineEffect: unknown) =>
      ({
        version: 1,
        display: 'inline',
        children: [
          { type: 'run', text: 'x', wordRunProperties: { textOutlineEffect } },
        ],
      }) as unknown as WorkDocumentEquation;
    expect(
      normalizeDocumentEquation({
        version: 1,
        display: 'inline',
        children: [
          {
            type: 'run',
            text: 'x',
            wordRunProperties: { textOutlineEffect: undefined },
          },
        ],
      } as unknown as WorkDocumentEquation),
    ).toEqual(simpleEquation('x'));
    expect(
      normalizeDocumentEquation(equationWithTextOutlineEffect(boundaryEffect))
        ?.children[0],
    ).toEqual({
      type: 'run',
      text: 'x',
      wordRunProperties: { textOutlineEffect: boundaryEffect },
    });
    const validStop = {
      positionPercent: 50,
      color: { type: 'rgb', value: '#abcdef' },
    };
    const invalidEffects = [
      null,
      false,
      { extra: true },
      { widthEmus: -1 },
      { widthEmus: 1.5 },
      { widthEmus: 20_116_801 },
      { cap: 'butt' },
      { compound: 'quadruple' },
      { alignment: 'outside' },
      { fill: null },
      { fill: { type: 'none', color: validStop.color } },
      { fill: { type: 'solid', extra: true } },
      { fill: { type: 'solid', color: null } },
      { fill: { type: 'gradient', extra: true } },
      { fill: { type: 'gradient', stops: [] } },
      { fill: { type: 'gradient', stops: [validStop] } },
      {
        fill: {
          type: 'gradient',
          stops: Array.from({ length: 11 }, () => validStop),
        },
      },
      { fill: { type: 'gradient', stops: 'two' } },
      {
        fill: {
          type: 'gradient',
          stops: [validStop, { ...validStop, extra: true }],
        },
      },
      {
        fill: {
          type: 'gradient',
          stops: [validStop, { ...validStop, positionPercent: -0.001 }],
        },
      },
      {
        fill: {
          type: 'gradient',
          stops: [validStop, { ...validStop, positionPercent: 100.001 }],
        },
      },
      {
        fill: {
          type: 'gradient',
          stops: [validStop, { ...validStop, positionPercent: 0.0005 }],
        },
      },
      {
        fill: {
          type: 'gradient',
          stops: [validStop, { positionPercent: 100 }],
        },
      },
      { fill: { type: 'gradient', shade: null } },
      { fill: { type: 'gradient', shade: { type: 'radial' } } },
      {
        fill: {
          type: 'gradient',
          shade: { type: 'linear', extra: true },
        },
      },
      {
        fill: {
          type: 'gradient',
          shade: { type: 'linear', angleDegrees: -1 / 60_000 },
        },
      },
      {
        fill: {
          type: 'gradient',
          shade: { type: 'linear', angleDegrees: 360 },
        },
      },
      {
        fill: {
          type: 'gradient',
          shade: { type: 'linear', angleDegrees: 1 / 120_000 },
        },
      },
      {
        fill: {
          type: 'gradient',
          shade: { type: 'linear', scaled: 'true' },
        },
      },
      {
        fill: {
          type: 'gradient',
          shade: { type: 'path', path: 'triangle' },
        },
      },
      {
        fill: {
          type: 'gradient',
          shade: { type: 'path', fillToRectangle: null },
        },
      },
      {
        fill: {
          type: 'gradient',
          shade: { type: 'path', fillToRectangle: { extra: true } },
        },
      },
      {
        fill: {
          type: 'gradient',
          shade: {
            type: 'path',
            fillToRectangle: { leftPercent: minimumPercentage - 0.001 },
          },
        },
      },
      {
        fill: {
          type: 'gradient',
          shade: {
            type: 'path',
            fillToRectangle: { rightPercent: maximumPercentage + 0.001 },
          },
        },
      },
      {
        fill: {
          type: 'gradient',
          shade: {
            type: 'path',
            fillToRectangle: { topPercent: 0.0005 },
          },
        },
      },
      { dash: null },
      { dash: { extra: true } },
      { dash: { preset: 'shortDash' } },
      { join: null },
      { join: { type: 'round', limitPercent: 1 } },
      { join: { type: 'bevel', extra: true } },
      { join: { type: 'miter', limitPercent: -0.001 } },
      {
        join: { type: 'miter', limitPercent: maximumPercentage + 0.001 },
      },
      { join: { type: 'miter', limitPercent: 0.0005 } },
      { join: { type: 'sharp' } },
    ];
    expect(
      invalidEffects.map((effect) =>
        normalizeDocumentEquation(equationWithTextOutlineEffect(effect)),
      ),
    ).toEqual(invalidEffects.map(() => null));

    const wordRun = (properties: string, namespace = WORD_NAMESPACE) =>
      `<m:r xmlns:w="${namespace}" xmlns:w14="${WORD_2010_NAMESPACE}"><w:rPr>${properties}</w:rPr><m:t>x</m:t></m:r>`;
    const fullMarkup =
      '<w14:textOutline w14:w=" +12700 " w14:cap="sq" ' +
      'w14:cmpd="dbl" w14:algn="in"><w14:gradFill><w14:gsLst>' +
      '<w14:gs w14:pos="0"><w14:srgbClr w14:val="a1B2c3">' +
      '<w14:tint w14:val="0"/><w14:alpha w14:val="50000"/>' +
      '<w14:lumMod w14:val="75000"/></w14:srgbClr></w14:gs>' +
      '<w14:gs w14:pos="100000"><w14:schemeClr w14:val="accent3"/></w14:gs>' +
      '</w14:gsLst><w14:lin w14:ang="5400000" w14:scaled="true"/>' +
      '</w14:gradFill><w14:prstDash w14:val="lgDashDotDot"/>' +
      '<w14:miter w14:lim="250000"/></w14:textOutline>';
    for (const namespace of [WORD_NAMESPACE, STRICT_WORD_NAMESPACE]) {
      expect(
        inspectEquationModel(wordRun(fullMarkup, namespace))?.children[0],
      ).toEqual({
        type: 'run',
        text: 'x',
        wordRunProperties: { textOutlineEffect: fullEffect },
      });
    }
    expect(
      inspectEquationModel(wordRun(WORD_2010_TEXT_OUTLINE))?.children[0],
    ).toEqual({
      type: 'run',
      text: 'x',
      wordRunProperties: { textOutlineEffect: {} },
    });
    expect(
      inspectEquationModel(
        wordRun(
          '<w14:textOutline><w14:solidFill/><w14:prstDash/><w14:miter/></w14:textOutline>',
        ),
      )?.children[0],
    ).toEqual({
      type: 'run',
      text: 'x',
      wordRunProperties: {
        textOutlineEffect: {
          fill: { type: 'solid' },
          dash: {},
          join: { type: 'miter' },
        },
      },
    });
    const boundaryStopMarkup = boundaryPositions
      .map(
        (position, index) =>
          `<w14:gs w14:pos="${position * 1_000}"><w14:srgbClr w14:val="${index
            .toString(16)
            .padStart(6, '0')}"/></w14:gs>`,
      )
      .join('');
    const boundaryMarkup =
      '<w14:textOutline w14:w="20116800" w14:cap="rnd" ' +
      'w14:cmpd="tri" w14:algn="in"><w14:gradFill><w14:gsLst>' +
      boundaryStopMarkup +
      '</w14:gsLst><w14:path w14:path="rect"><w14:fillToRect ' +
      'w14:l="-2147483648" w14:t="2147483647" w14:r="0" w14:b="1"/>' +
      '</w14:path></w14:gradFill><w14:prstDash/>' +
      '<w14:miter w14:lim="2147483647"/></w14:textOutline>';
    expect(inspectEquationModel(wordRun(boundaryMarkup))?.children[0]).toEqual({
      type: 'run',
      text: 'x',
      wordRunProperties: { textOutlineEffect: boundaryEffect },
    });

    const attributeEnums = [
      ['cap', 'rnd', 'round'],
      ['cap', 'sq', 'square'],
      ['cap', 'flat', 'flat'],
      ['cmpd', 'sng', 'single'],
      ['cmpd', 'dbl', 'double'],
      ['cmpd', 'thickThin', 'thickThin'],
      ['cmpd', 'thinThick', 'thinThick'],
      ['cmpd', 'tri', 'triple'],
      ['algn', 'ctr', 'center'],
      ['algn', 'in', 'inset'],
    ] as const;
    for (const [attribute, source, value] of attributeEnums) {
      expect(
        inspectEquationModel(
          wordRun(`<w14:textOutline w14:${attribute}="${source}"/>`),
        )?.children[0],
      ).toEqual({
        type: 'run',
        text: 'x',
        wordRunProperties: {
          textOutlineEffect: {
            [attribute === 'algn'
              ? 'alignment'
              : attribute === 'cmpd'
                ? 'compound'
                : 'cap']: value,
          },
        },
      });
    }
    const dashEnums = [
      ['solid', 'solid'],
      ['dot', 'dot'],
      ['sysDot', 'systemDot'],
      ['dash', 'dash'],
      ['sysDash', 'systemDash'],
      ['lgDash', 'longDash'],
      ['dashDot', 'dashDot'],
      ['sysDashDot', 'systemDashDot'],
      ['lgDashDot', 'longDashDot'],
      ['lgDashDotDot', 'longDashDotDot'],
      ['sysDashDotDot', 'systemDashDotDot'],
    ] as const;
    for (const [source, preset] of dashEnums) {
      expect(
        inspectEquationModel(
          wordRun(
            `<w14:textOutline><w14:prstDash w14:val="${source}"/></w14:textOutline>`,
          ),
        )?.children[0],
      ).toEqual({
        type: 'run',
        text: 'x',
        wordRunProperties: { textOutlineEffect: { dash: { preset } } },
      });
    }
    for (const [source, path] of [
      ['shape', 'shape'],
      ['circle', 'circle'],
      ['rect', 'rectangle'],
    ] as const) {
      expect(
        inspectEquationModel(
          wordRun(
            `<w14:textOutline><w14:gradFill><w14:path w14:path="${source}"/></w14:gradFill></w14:textOutline>`,
          ),
        )?.children[0],
      ).toEqual({
        type: 'run',
        text: 'x',
        wordRunProperties: {
          textOutlineEffect: {
            fill: { type: 'gradient', shade: { type: 'path', path } },
          },
        },
      });
    }
    for (const [source, scaled] of [
      ['true', true],
      ['1', true],
      ['false', false],
      ['0', false],
    ] as const) {
      expect(
        inspectEquationModel(
          wordRun(
            `<w14:textOutline><w14:gradFill><w14:lin w14:scaled="${source}"/></w14:gradFill></w14:textOutline>`,
          ),
        )?.children[0],
      ).toEqual({
        type: 'run',
        text: 'x',
        wordRunProperties: {
          textOutlineEffect: {
            fill: { type: 'gradient', shade: { type: 'linear', scaled } },
          },
        },
      });
    }
    expect(
      inspectEquationModel(
        wordRun(
          '<w:shadow w:val="0"/><w:specVanish w:val="0"/>' +
            '<w14:glow><w14:srgbClr w14:val="112233"/></w14:glow>' +
            '<w14:shadow><w14:srgbClr w14:val="445566"/></w14:shadow>' +
            '<w14:reflection w14:stA="25000"/>' +
            '<w14:textOutline><w14:noFill/></w14:textOutline>',
        ),
      )?.children[0],
    ).toEqual({
      type: 'run',
      text: 'x',
      wordRunProperties: {
        shadow: false,
        paragraphMarkAlwaysHidden: false,
        glow: { color: { type: 'rgb', value: '#112233' } },
        shadowEffect: { color: { type: 'rgb', value: '#445566' } },
        reflectionEffect: { startOpacityPercent: 25 },
        textOutlineEffect: { fill: { type: 'none' } },
      },
    });

    const gradientStop = (
      position = '0',
      color = '<w14:srgbClr w14:val="A1B2C3"/>',
    ) => `<w14:gs w14:pos="${position}">${color}</w14:gs>`;
    const twoStops = `${gradientStop()}${gradientStop('100000')}`;
    const invalidMarkup = [
      wordRun('<w14:textOutline/><w14:textOutline/>'),
      wordRun('<w14:textOutline/><w14:reflection/>'),
      wordRun(
        '<w14:textOutline/><w14:shadow><w14:srgbClr w14:val="000000"/></w14:shadow>',
      ),
      wordRun('<w14:textOutline/><w:specVanish/>'),
      wordRun('<w:textOutline/>'),
      wordRun(`<v:textOutline xmlns:v="${VENDOR_NAMESPACE}"/>`),
      wordRun(`<w14:textOutline xmlns:w14="${VENDOR_NAMESPACE}" w14:w="1"/>`),
      wordRun('<w14:textOutline w="1"/>'),
      wordRun('<w14:textOutline w14:w="-1"/>'),
      wordRun('<w14:textOutline w14:w="1.5"/>'),
      wordRun('<w14:textOutline w14:w="20116801"/>'),
      wordRun('<w14:textOutline w14:cap="butt"/>'),
      wordRun('<w14:textOutline w14:cap="RND"/>'),
      wordRun('<w14:textOutline w14:cmpd="quad"/>'),
      wordRun('<w14:textOutline w14:algn="out"/>'),
      wordRun('<w14:textOutline w14:extra="semantic"/>'),
      wordRun(
        `<w14:textOutline xmlns:r="${RELATIONSHIP_NAMESPACE}" r:id="rIdUnsafe"/>`,
      ),
      wordRun('<w14:textOutline>meaningful</w14:textOutline>'),
      wordRun(
        '<w14:textOutline><w14:noFill/><w14:solidFill/></w14:textOutline>',
      ),
      wordRun(
        '<w14:textOutline><w14:prstDash/><w14:noFill/></w14:textOutline>',
      ),
      wordRun('<w14:textOutline><w14:round/><w14:prstDash/></w14:textOutline>'),
      wordRun(
        '<w14:textOutline><w14:prstDash/><w14:prstDash/></w14:textOutline>',
      ),
      wordRun('<w14:textOutline><w14:round/><w14:bevel/></w14:textOutline>'),
      wordRun('<w14:textOutline><w14:noFill w14:extra="1"/></w14:textOutline>'),
      wordRun(
        '<w14:textOutline><w14:noFill>text</w14:noFill></w14:textOutline>',
      ),
      wordRun(
        '<w14:textOutline><w14:noFill><w14:alpha w14:val="1"/></w14:noFill></w14:textOutline>',
      ),
      wordRun(
        '<w14:textOutline><w14:solidFill w14:extra="1"/></w14:textOutline>',
      ),
      wordRun(
        '<w14:textOutline><w14:solidFill><w14:srgbClr w14:val="A1B2C3"/><w14:schemeClr w14:val="accent1"/></w14:solidFill></w14:textOutline>',
      ),
      wordRun(
        '<w14:textOutline><w14:solidFill><w14:alpha w14:val="50000"/></w14:solidFill></w14:textOutline>',
      ),
      wordRun(
        '<w14:textOutline><w14:gradFill w14:extra="1"/></w14:textOutline>',
      ),
      wordRun(
        `<w14:textOutline><w14:gradFill><w14:lin/><w14:gsLst>${twoStops}</w14:gsLst></w14:gradFill></w14:textOutline>`,
      ),
      wordRun(
        '<w14:textOutline><w14:gradFill><w14:lin/><w14:path/></w14:gradFill></w14:textOutline>',
      ),
      wordRun(
        `<w14:textOutline><w14:gradFill><w14:gsLst w14:extra="1">${twoStops}</w14:gsLst></w14:gradFill></w14:textOutline>`,
      ),
      wordRun(
        `<w14:textOutline><w14:gradFill><w14:gsLst>${gradientStop()}</w14:gsLst></w14:gradFill></w14:textOutline>`,
      ),
      wordRun(
        `<w14:textOutline><w14:gradFill><w14:gsLst>${Array.from({ length: 11 }, (_, index) => gradientStop(String(index * 10_000))).join('')}</w14:gsLst></w14:gradFill></w14:textOutline>`,
      ),
      wordRun(
        `<w14:textOutline><w14:gradFill><w14:gsLst><w14:gs><w14:srgbClr w14:val="A1B2C3"/></w14:gs>${gradientStop('100000')}</w14:gsLst></w14:gradFill></w14:textOutline>`,
      ),
      wordRun(
        `<w14:textOutline><w14:gradFill><w14:gsLst>${gradientStop('-1')}${gradientStop('100000')}</w14:gsLst></w14:gradFill></w14:textOutline>`,
      ),
      wordRun(
        `<w14:textOutline><w14:gradFill><w14:gsLst>${gradientStop('100001')}${gradientStop('0')}</w14:gsLst></w14:gradFill></w14:textOutline>`,
      ),
      wordRun(
        `<w14:textOutline><w14:gradFill><w14:gsLst>${gradientStop('1.5')}${gradientStop('100000')}</w14:gsLst></w14:gradFill></w14:textOutline>`,
      ),
      wordRun(
        `<w14:textOutline><w14:gradFill><w14:gsLst><w14:gs w14:pos="0"/>${gradientStop('100000')}</w14:gsLst></w14:gradFill></w14:textOutline>`,
      ),
      wordRun(
        `<w14:textOutline><w14:gradFill><w14:gsLst><w14:gs w14:pos="0"><w14:srgbClr w14:val="A1B2C3"/><w14:schemeClr w14:val="accent1"/></w14:gs>${gradientStop('100000')}</w14:gsLst></w14:gradFill></w14:textOutline>`,
      ),
      wordRun(
        '<w14:textOutline><w14:gradFill><w14:lin w14:extra="1"/></w14:gradFill></w14:textOutline>',
      ),
      wordRun(
        '<w14:textOutline><w14:gradFill><w14:lin w14:ang="-1"/></w14:gradFill></w14:textOutline>',
      ),
      wordRun(
        '<w14:textOutline><w14:gradFill><w14:lin w14:ang="21600000"/></w14:gradFill></w14:textOutline>',
      ),
      wordRun(
        '<w14:textOutline><w14:gradFill><w14:lin w14:ang="1.5"/></w14:gradFill></w14:textOutline>',
      ),
      wordRun(
        '<w14:textOutline><w14:gradFill><w14:lin w14:scaled="on"/></w14:gradFill></w14:textOutline>',
      ),
      wordRun(
        '<w14:textOutline><w14:gradFill><w14:lin><w14:fillToRect/></w14:lin></w14:gradFill></w14:textOutline>',
      ),
      wordRun(
        '<w14:textOutline><w14:gradFill><w14:path w14:path="triangle"/></w14:gradFill></w14:textOutline>',
      ),
      wordRun(
        '<w14:textOutline><w14:gradFill><w14:path><w14:fillToRect/><w14:fillToRect/></w14:path></w14:gradFill></w14:textOutline>',
      ),
      wordRun(
        '<w14:textOutline><w14:gradFill><w14:path><w14:fillToRect l="1"/></w14:path></w14:gradFill></w14:textOutline>',
      ),
      wordRun(
        '<w14:textOutline><w14:gradFill><w14:path><w14:fillToRect w14:l="-2147483649"/></w14:path></w14:gradFill></w14:textOutline>',
      ),
      wordRun(
        '<w14:textOutline><w14:gradFill><w14:path><w14:fillToRect w14:r="2147483648"/></w14:path></w14:gradFill></w14:textOutline>',
      ),
      wordRun(
        '<w14:textOutline><w14:gradFill><w14:path><w14:fillToRect w14:t="1.5"/></w14:path></w14:gradFill></w14:textOutline>',
      ),
      wordRun(
        '<w14:textOutline><w14:gradFill><w14:path><w14:fillToRect><w14:noFill/></w14:fillToRect></w14:path></w14:gradFill></w14:textOutline>',
      ),
      wordRun(
        '<w14:textOutline><w14:prstDash w14:val="shortDash"/></w14:textOutline>',
      ),
      wordRun('<w14:textOutline><w14:prstDash val="solid"/></w14:textOutline>'),
      wordRun(
        '<w14:textOutline><w14:prstDash w14:extra="1"/></w14:textOutline>',
      ),
      wordRun(
        '<w14:textOutline><w14:prstDash><w14:noFill/></w14:prstDash></w14:textOutline>',
      ),
      wordRun('<w14:textOutline><w14:round w14:extra="1"/></w14:textOutline>'),
      wordRun(
        '<w14:textOutline><w14:bevel><w14:noFill/></w14:bevel></w14:textOutline>',
      ),
      wordRun('<w14:textOutline><w14:miter w14:lim="-1"/></w14:textOutline>'),
      wordRun(
        '<w14:textOutline><w14:miter w14:lim="2147483648"/></w14:textOutline>',
      ),
      wordRun('<w14:textOutline><w14:miter w14:lim="1.5"/></w14:textOutline>'),
      wordRun(
        '<w14:textOutline><w14:miter><w14:noFill/></w14:miter></w14:textOutline>',
      ),
      wordRun(
        '<w14:textOutline><v:noFill xmlns:v="urn:a3s:test"/></w14:textOutline>',
      ),
      wordRun(WORD_2010_SCENE_3D),
    ];
    expect(invalidMarkup.map(inspectEquationBody)).toEqual(
      invalidMarkup.map(() => 'unsupported'),
    );

    const document = new DOMParser().parseFromString('', 'text/html');
    const preview = createDocumentEquationElement(document, equation);
    for (const text of [
      'full-text-outline-effect',
      'explicit-default-text-outline-effect',
      'empty-text-outline-effect',
      'empty-components-text-outline-effect',
      'ordered-text-outline-effect',
      'boundary-text-outline-effect',
      'operator-text-outline-effect',
    ]) {
      expect(preview.textContent).toContain(text);
    }
    expect(preview.querySelector('math')?.outerHTML).not.toMatch(
      /text-stroke|paint-order|stroke/iu,
    );
    const sanitized = new DOMParser().parseFromString(
      sanitizeDocumentPageChromeHtml(preview.outerHTML),
      'text/html',
    );
    expect(
      documentEquationFromElement(
        sanitized.body.querySelector<HTMLElement>(
          '[data-document-equation]',
        ) as HTMLElement,
      ),
    ).toEqual(equation);

    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    artifact.content.html = `<p>${preview.outerHTML}</p>`;
    const first = await createArtifactBlob(artifact);
    await expectNativeWordRunTextOutlineEffects(first);
    const imported = await importOfficeFile(
      new File([first], 'word-run-text-outline-effect.docx', {
        type: first.type,
      }),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const importedDocument = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    expect(
      documentEquationFromElement(
        importedDocument.body.querySelector<HTMLElement>(
          '[data-document-equation]',
        ) as HTMLElement,
      ),
    ).toEqual(equation);
    expect(imported.compatibility.issues).not.toContainEqual(
      expect.objectContaining({ code: 'docx.equations.unsupported' }),
    );
    await expectNativeWordRunTextOutlineEffects(
      await createArtifactBlob(imported),
    );
  });

  test('preserves bounded Office 2010 text fill effects and projects only exact solid colors', async () => {
    const transforms = [
      { type: 'tint', value: 0 },
      { type: 'alpha', value: 50_000 },
      { type: 'luminanceModulation', value: 75_000 },
    ];
    const fullEffect = {
      fill: {
        type: 'gradient',
        stops: [
          {
            positionPercent: 0,
            color: { type: 'rgb', value: '#a1b2c3', transforms },
          },
          {
            positionPercent: 100,
            color: { type: 'scheme', value: 'accent3' },
          },
        ],
        shade: { type: 'linear', angleDegrees: 90, scaled: true },
      },
    };
    const boundaryPositions = [0, 0.001, 10, 20, 30, 40, 50, 60, 99.999, 100];
    const boundaryStops = boundaryPositions.map((positionPercent, index) => ({
      positionPercent,
      color: {
        type: 'rgb',
        value: `#${index.toString(16).padStart(6, '0')}`,
      },
    }));
    const minimumPercentage = -2_147_483_648 / 1_000;
    const maximumPercentage = 2_147_483_647 / 1_000;
    const boundaryEffect = {
      fill: {
        type: 'gradient',
        stops: boundaryStops,
        shade: {
          type: 'path',
          path: 'rectangle',
          fillToRectangle: {
            leftPercent: minimumPercentage,
            topPercent: maximumPercentage,
            rightPercent: 0,
            bottomPercent: 0.001,
          },
        },
      },
    };
    const equation = {
      version: 1,
      display: 'inline',
      children: [
        {
          type: 'run',
          text: 'full-text-fill-effect',
          wordRunProperties: { textFillEffect: fullEffect },
        },
        {
          type: 'run',
          text: 'direct-solid-text-fill-effect',
          wordRunProperties: {
            color: { value: '#abcdef' },
            textFillEffect: {
              fill: {
                type: 'solid',
                color: { type: 'rgb', value: '#13579b' },
              },
            },
          },
        },
        {
          type: 'run',
          text: 'empty-text-fill-effect',
          wordRunProperties: { textFillEffect: {} },
        },
        {
          type: 'run',
          text: 'empty-solid-text-fill-effect',
          wordRunProperties: {
            textFillEffect: { fill: { type: 'solid' } },
          },
        },
        {
          type: 'run',
          text: 'empty-gradient-text-fill-effect',
          wordRunProperties: {
            textFillEffect: { fill: { type: 'gradient' } },
          },
        },
        {
          type: 'run',
          text: 'no-text-fill-effect',
          wordRunProperties: {
            color: { value: '#2468ac' },
            textFillEffect: { fill: { type: 'none' } },
          },
        },
        {
          type: 'run',
          text: 'ordered-text-fill-effect',
          wordRunProperties: {
            paragraphMarkAlwaysHidden: false,
            reflectionEffect: {},
            textOutlineEffect: {},
            textFillEffect: { fill: { type: 'none' } },
          },
        },
        {
          type: 'run',
          text: 'boundary-text-fill-effect',
          wordRunProperties: { textFillEffect: boundaryEffect },
        },
        {
          type: 'nary',
          operator: '\u2211',
          limitLocation: 'underOver',
          controlProperties: {
            textFillEffect: {
              fill: {
                type: 'solid',
                color: { type: 'scheme', value: 'hyperlink' },
              },
            },
          },
          children: [{ type: 'run', text: 'operator-text-fill-effect' }],
        },
      ],
    } as unknown as WorkDocumentEquation;
    expect(normalizeDocumentEquation(equation)).toEqual(equation);

    const equationWithTextFillEffect = (textFillEffect: unknown) =>
      ({
        version: 1,
        display: 'inline',
        children: [
          { type: 'run', text: 'x', wordRunProperties: { textFillEffect } },
        ],
      }) as unknown as WorkDocumentEquation;
    expect(
      normalizeDocumentEquation({
        version: 1,
        display: 'inline',
        children: [
          {
            type: 'run',
            text: 'x',
            wordRunProperties: { textFillEffect: undefined },
          },
        ],
      } as unknown as WorkDocumentEquation),
    ).toEqual(simpleEquation('x'));
    expect(
      normalizeDocumentEquation(equationWithTextFillEffect(boundaryEffect))
        ?.children[0],
    ).toEqual({
      type: 'run',
      text: 'x',
      wordRunProperties: { textFillEffect: boundaryEffect },
    });
    const validStop = {
      positionPercent: 50,
      color: { type: 'rgb', value: '#abcdef' },
    };
    const invalidEffects = [
      null,
      false,
      [],
      { extra: true },
      { fill: null },
      { fill: { type: 'none', color: validStop.color } },
      { fill: { type: 'solid', extra: true } },
      { fill: { type: 'solid', color: null } },
      { fill: { type: 'gradient', extra: true } },
      { fill: { type: 'gradient', stops: [] } },
      { fill: { type: 'gradient', stops: [validStop] } },
      {
        fill: {
          type: 'gradient',
          stops: Array.from({ length: 11 }, () => validStop),
        },
      },
      {
        fill: {
          type: 'gradient',
          stops: [validStop, { ...validStop, positionPercent: -0.001 }],
        },
      },
      {
        fill: {
          type: 'gradient',
          stops: [validStop, { ...validStop, positionPercent: 100.001 }],
        },
      },
      {
        fill: {
          type: 'gradient',
          stops: [validStop, { ...validStop, positionPercent: 0.0005 }],
        },
      },
      { fill: { type: 'gradient', shade: null } },
      {
        fill: {
          type: 'gradient',
          shade: { type: 'linear', angleDegrees: 360 },
        },
      },
      {
        fill: {
          type: 'gradient',
          shade: { type: 'linear', scaled: 'true' },
        },
      },
      {
        fill: {
          type: 'gradient',
          shade: { type: 'path', path: 'triangle' },
        },
      },
      {
        fill: {
          type: 'gradient',
          shade: { type: 'path', fillToRectangle: { extra: true } },
        },
      },
      {
        fill: {
          type: 'gradient',
          shade: {
            type: 'path',
            fillToRectangle: { leftPercent: minimumPercentage - 0.001 },
          },
        },
      },
      {
        fill: {
          type: 'gradient',
          shade: {
            type: 'path',
            fillToRectangle: { rightPercent: maximumPercentage + 0.001 },
          },
        },
      },
    ];
    expect(
      invalidEffects.map((effect) =>
        normalizeDocumentEquation(equationWithTextFillEffect(effect)),
      ),
    ).toEqual(invalidEffects.map(() => null));

    const wordRun = (properties: string, namespace = WORD_NAMESPACE) =>
      `<m:r xmlns:w="${namespace}" xmlns:w14="${WORD_2010_NAMESPACE}"><w:rPr>${properties}</w:rPr><m:t>x</m:t></m:r>`;
    const fullMarkup =
      '<w14:textFill><w14:gradFill><w14:gsLst>' +
      '<w14:gs w14:pos="0"><w14:srgbClr w14:val="a1B2c3">' +
      '<w14:tint w14:val="0"/><w14:alpha w14:val="50000"/>' +
      '<w14:lumMod w14:val="75000"/></w14:srgbClr></w14:gs>' +
      '<w14:gs w14:pos="100000"><w14:schemeClr w14:val="accent3"/></w14:gs>' +
      '</w14:gsLst><w14:lin w14:ang="5400000" w14:scaled="true"/>' +
      '</w14:gradFill></w14:textFill>';
    for (const namespace of [WORD_NAMESPACE, STRICT_WORD_NAMESPACE]) {
      expect(
        inspectEquationModel(wordRun(fullMarkup, namespace))?.children[0],
      ).toEqual({
        type: 'run',
        text: 'x',
        wordRunProperties: { textFillEffect: fullEffect },
      });
    }
    expect(
      inspectEquationModel(wordRun(WORD_2010_TEXT_FILL))?.children[0],
    ).toEqual({
      type: 'run',
      text: 'x',
      wordRunProperties: { textFillEffect: {} },
    });
    expect(
      inspectEquationModel(
        wordRun('<w14:textFill><w14:solidFill/></w14:textFill>'),
      )?.children[0],
    ).toEqual({
      type: 'run',
      text: 'x',
      wordRunProperties: {
        textFillEffect: { fill: { type: 'solid' } },
      },
    });
    expect(
      inspectEquationModel(
        wordRun('<w14:textOutline/><w14:textFill><w14:noFill/></w14:textFill>'),
      )?.children[0],
    ).toEqual({
      type: 'run',
      text: 'x',
      wordRunProperties: {
        textOutlineEffect: {},
        textFillEffect: { fill: { type: 'none' } },
      },
    });
    const boundaryStopMarkup = boundaryPositions
      .map(
        (position, index) =>
          `<w14:gs w14:pos="${position * 1_000}"><w14:srgbClr w14:val="${index
            .toString(16)
            .padStart(6, '0')}"/></w14:gs>`,
      )
      .join('');
    const boundaryMarkup =
      '<w14:textFill><w14:gradFill><w14:gsLst>' +
      boundaryStopMarkup +
      '</w14:gsLst><w14:path w14:path="rect"><w14:fillToRect ' +
      'w14:l="-2147483648" w14:t="2147483647" w14:r="0" w14:b="1"/>' +
      '</w14:path></w14:gradFill></w14:textFill>';
    expect(inspectEquationModel(wordRun(boundaryMarkup))?.children[0]).toEqual({
      type: 'run',
      text: 'x',
      wordRunProperties: { textFillEffect: boundaryEffect },
    });

    const gradientStop = (
      position = '0',
      color = '<w14:srgbClr w14:val="A1B2C3"/>',
    ) => `<w14:gs w14:pos="${position}">${color}</w14:gs>`;
    const twoStops = `${gradientStop()}${gradientStop('100000')}`;
    const invalidMarkup = [
      wordRun('<w14:textFill/><w14:textFill/>'),
      wordRun('<w14:textFill/><w14:textOutline/>'),
      wordRun('<w14:textFill/><w14:reflection/>'),
      wordRun('<w14:textFill/><w:specVanish/>'),
      wordRun('<w:textFill/>'),
      wordRun(`<v:textFill xmlns:v="${VENDOR_NAMESPACE}"/>`),
      wordRun(`<w14:textFill xmlns:w14="${VENDOR_NAMESPACE}"/>`),
      wordRun('<w14:textFill w14:extra="semantic"/>'),
      wordRun('<w14:textFill extra="semantic"/>'),
      wordRun(
        `<w14:textFill xmlns:r="${RELATIONSHIP_NAMESPACE}" r:id="rIdUnsafe"/>`,
      ),
      wordRun('<w14:textFill>meaningful</w14:textFill>'),
      wordRun('<w14:textFill><w14:noFill/><w14:solidFill/></w14:textFill>'),
      wordRun('<w14:textFill><w14:prstDash/></w14:textFill>'),
      wordRun('<w14:textFill><w14:noFill w14:extra="1"/></w14:textFill>'),
      wordRun('<w14:textFill><w14:noFill>text</w14:noFill></w14:textFill>'),
      wordRun(
        '<w14:textFill><w14:noFill><w14:alpha w14:val="1"/></w14:noFill></w14:textFill>',
      ),
      wordRun(
        '<w14:textFill><w14:solidFill><w14:srgbClr w14:val="A1B2C3"/><w14:schemeClr w14:val="accent1"/></w14:solidFill></w14:textFill>',
      ),
      wordRun(
        '<w14:textFill><w14:solidFill><w14:alpha w14:val="50000"/></w14:solidFill></w14:textFill>',
      ),
      wordRun(
        `<w14:textFill><w14:gradFill><w14:lin/><w14:gsLst>${twoStops}</w14:gsLst></w14:gradFill></w14:textFill>`,
      ),
      wordRun(
        '<w14:textFill><w14:gradFill><w14:lin/><w14:path/></w14:gradFill></w14:textFill>',
      ),
      wordRun(
        `<w14:textFill><w14:gradFill><w14:gsLst>${gradientStop()}</w14:gsLst></w14:gradFill></w14:textFill>`,
      ),
      wordRun(
        `<w14:textFill><w14:gradFill><w14:gsLst>${Array.from({ length: 11 }, (_, index) => gradientStop(String(index * 10_000))).join('')}</w14:gsLst></w14:gradFill></w14:textFill>`,
      ),
      wordRun(
        `<w14:textFill><w14:gradFill><w14:gsLst>${gradientStop('-1')}${gradientStop('100000')}</w14:gsLst></w14:gradFill></w14:textFill>`,
      ),
      wordRun(
        `<w14:textFill><w14:gradFill><w14:gsLst>${gradientStop('100001')}${gradientStop('0')}</w14:gsLst></w14:gradFill></w14:textFill>`,
      ),
      wordRun(
        '<w14:textFill><w14:gradFill><w14:lin w14:ang="21600000"/></w14:gradFill></w14:textFill>',
      ),
      wordRun(
        '<w14:textFill><w14:gradFill><w14:lin w14:scaled="on"/></w14:gradFill></w14:textFill>',
      ),
      wordRun(
        '<w14:textFill><w14:gradFill><w14:path w14:path="triangle"/></w14:gradFill></w14:textFill>',
      ),
      wordRun(
        '<w14:textFill><w14:gradFill><w14:path><w14:fillToRect w14:l="-2147483649"/></w14:path></w14:gradFill></w14:textFill>',
      ),
      wordRun(
        '<w14:textFill><w14:gradFill><w14:path><w14:fillToRect w14:r="2147483648"/></w14:path></w14:gradFill></w14:textFill>',
      ),
      wordRun(
        '<w14:textFill><v:noFill xmlns:v="urn:a3s:test"/></w14:textFill>',
      ),
      wordRun(WORD_2010_SCENE_3D),
    ];
    expect(invalidMarkup.map(inspectEquationBody)).toEqual(
      invalidMarkup.map(() => 'unsupported'),
    );

    const document = new DOMParser().parseFromString('', 'text/html');
    const preview = createDocumentEquationElement(document, equation);
    const previewRun = (text: string): Element | undefined =>
      Array.from(preview.querySelectorAll('mtext, mo')).find(
        (candidate) => candidate.textContent === text,
      );
    expect(
      previewRun('direct-solid-text-fill-effect')?.getAttribute('mathcolor'),
    ).toBe('#13579b');
    expect(
      previewRun('empty-text-fill-effect')?.getAttribute('mathcolor'),
    ).toBe('#000000');
    expect(
      previewRun('empty-solid-text-fill-effect')?.getAttribute('mathcolor'),
    ).toBe('#000000');
    expect(
      previewRun('empty-gradient-text-fill-effect')?.getAttribute('mathcolor'),
    ).toBe('#000000');
    expect(previewRun('no-text-fill-effect')?.getAttribute('mathcolor')).toBe(
      '#2468ac',
    );
    expect(
      previewRun('full-text-fill-effect')?.getAttribute('mathcolor'),
    ).toBeNull();
    expect(
      previewRun('boundary-text-fill-effect')?.getAttribute('mathcolor'),
    ).toBeNull();
    expect(preview.outerHTML).not.toMatch(
      /background-clip|linear-gradient|text-fill-color|color:\s*transparent/iu,
    );
    const sanitized = new DOMParser().parseFromString(
      sanitizeDocumentPageChromeHtml(preview.outerHTML),
      'text/html',
    );
    expect(
      documentEquationFromElement(
        sanitized.body.querySelector<HTMLElement>(
          '[data-document-equation]',
        ) as HTMLElement,
      ),
    ).toEqual(equation);

    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    artifact.content.html = `<p>${preview.outerHTML}</p>`;
    const first = await createArtifactBlob(artifact);
    await expectNativeWordRunTextFillEffects(first);
    const imported = await importOfficeFile(
      new File([first], 'word-run-text-fill-effect.docx', {
        type: first.type,
      }),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const importedDocument = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    expect(
      documentEquationFromElement(
        importedDocument.body.querySelector<HTMLElement>(
          '[data-document-equation]',
        ) as HTMLElement,
      ),
    ).toEqual(equation);
    expect(imported.compatibility.issues).not.toContainEqual(
      expect.objectContaining({ code: 'docx.equations.unsupported' }),
    );
    await expectNativeWordRunTextFillEffects(
      await createArtifactBlob(imported),
    );
  });

  test('preserves bounded Word control properties across OMML object containers', async () => {
    const equation = controlPropertiesEquation();
    expect(normalizeDocumentEquation(equation)).toEqual(equation);
    expect(
      normalizeDocumentEquation({
        version: 1,
        display: 'inline',
        children: [
          {
            type: 'fraction',
            controlProperties: {},
            fractionType: 'bar',
            numerator: [],
            denominator: [],
          },
        ],
      }),
    ).toEqual({
      version: 1,
      display: 'inline',
      children: [
        {
          type: 'fraction',
          fractionType: 'bar',
          numerator: [],
          denominator: [],
        },
      ],
    });
    expect(
      normalizeDocumentEquation({
        version: 1,
        display: 'inline',
        children: [
          {
            type: 'fraction',
            controlProperties: { extra: true },
            fractionType: 'bar',
            numerator: [],
            denominator: [],
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
            controlProperties: richWordRunProperties(),
          },
        ],
      }),
    ).toBeNull();

    const run = '<m:r><m:t>x</m:t></m:r>';
    for (const controlProperties of [
      '<m:ctrlPr/>',
      `<m:ctrlPr><w:rPr xmlns:w="${WORD_NAMESPACE}"/></m:ctrlPr>`,
    ]) {
      expect(
        inspectEquationModel(
          `<m:f><m:fPr>${controlProperties}</m:fPr><m:num>${run}</m:num><m:den>${run}</m:den></m:f>`,
        )?.children[0],
      ).toEqual({
        type: 'fraction',
        fractionType: 'bar',
        numerator: [{ type: 'run', text: 'x' }],
        denominator: [{ type: 'run', text: 'x' }],
      });
    }

    const invalidControlProperties = (children: string, attributes = '') =>
      `<m:acc><m:accPr><m:chr m:val="&#x303;"/><m:ctrlPr xmlns:w="${WORD_NAMESPACE}" xmlns:r="${RELATIONSHIP_NAMESPACE}" xmlns:v="${VENDOR_NAMESPACE}" ${attributes}>${children}</m:ctrlPr></m:accPr><m:e>${run}</m:e></m:acc>`;
    const unsupported = [
      invalidControlProperties('<w:rPr/>', 'm:extra="semantic"'),
      invalidControlProperties('meaningful<w:rPr/>'),
      invalidControlProperties('<w:rPr/><w:rPr/>'),
      invalidControlProperties('<w:ins><w:rPr/></w:ins>'),
      invalidControlProperties('<w:del><w:rPr/></w:del>'),
      invalidControlProperties('<w:moveFrom><w:rPr/></w:moveFrom>'),
      invalidControlProperties('<w:moveTo><w:rPr/></w:moveTo>'),
      invalidControlProperties('<v:rPr/>'),
      invalidControlProperties('<m:rPr/>'),
      invalidControlProperties('<w:rPr r:id="rIdUnsafe"/>'),
      invalidControlProperties(`<w:rPr>${WORD_2010_SCENE_3D}</w:rPr>`),
      invalidControlProperties(
        '<w:rPr><w:b/><w:rFonts w:ascii="Cambria Math"/></w:rPr>',
      ),
      invalidControlProperties('<w:b/>'),
      `<m:acc><m:accPr><m:ctrlPr/><m:chr m:val="&#x303;"/></m:accPr><m:e>${run}</m:e></m:acc>`,
    ];
    expect(unsupported.map(inspectEquationBody)).toEqual(
      unsupported.map(() => 'unsupported'),
    );

    const document = new DOMParser().parseFromString('', 'text/html');
    const preview = createDocumentEquationElement(document, equation);
    const naryOperator = Array.from(preview.querySelectorAll('mo')).find(
      (operator) => operator.textContent === '\u2211',
    );
    expect(naryOperator?.getAttribute('mathcolor')).toBe('#1a2b3c');
    expect(naryOperator?.getAttribute('mathsize')).toBe('12.5pt');
    expect(naryOperator?.getAttribute('dir')).toBe('ltr');
    expect(naryOperator?.getAttribute('lang')).toBe('en-US');
    expect(naryOperator?.getAttribute('style')).toContain(
      'font-family:"Cambria Math"',
    );
    expect(preview.querySelectorAll('mtext[mathcolor]')).toHaveLength(0);
    expect(
      Array.from(preview.querySelectorAll('mo[fence="true"]')).every(
        (operator) => operator.getAttribute('mathcolor') === '#1a2b3c',
      ),
    ).toBe(true);

    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    artifact.content.html = `<p>${equationHtml(equation)}</p>`;
    const first = await createArtifactBlob(artifact);
    await expectNativeControlProperties(first);

    const firstArchive = await JSZip.loadAsync(await first.arrayBuffer());
    const firstDocument = await xmlEntry(firstArchive, 'word/document.xml');
    const nativeEquation = descendants(firstDocument, 'oMath').find(
      (candidate) => candidate.namespaceURI === MATH_NAMESPACE,
    );
    expect(nativeEquation).toBeDefined();
    const standalone = nativeEquation?.cloneNode(true) as Element;
    standalone.setAttributeNS(XMLNS_NAMESPACE, 'xmlns', MATH_NAMESPACE);
    standalone.setAttributeNS(XMLNS_NAMESPACE, 'xmlns:m', MATH_NAMESPACE);
    standalone.setAttributeNS(XMLNS_NAMESPACE, 'xmlns:w', WORD_NAMESPACE);
    standalone.setAttributeNS(
      XMLNS_NAMESPACE,
      'xmlns:w14',
      WORD_2010_NAMESPACE,
    );
    const strictSource = new XMLSerializer()
      .serializeToString(standalone)
      .replaceAll(MATH_NAMESPACE, STRICT_MATH_NAMESPACE)
      .replaceAll(WORD_NAMESPACE, STRICT_WORD_NAMESPACE);
    const strictInspection = inspectEquationRoot(strictSource);
    expect(strictInspection).toMatchObject({
      status: 'supported',
      equation,
    });

    const imported = await importOfficeFile(
      new File([first], 'control-properties.docx', { type: first.type }),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const importedDocument = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    const importedEquation = importedDocument.body.querySelector<HTMLElement>(
      '[data-document-equation]',
    );
    expect(
      documentEquationFromElement(importedEquation as HTMLElement),
    ).toEqual(equation);
    expect(imported.compatibility.issues).not.toContainEqual(
      expect.objectContaining({ code: 'docx.equations.unsupported' }),
    );
    await expectNativeControlProperties(await createArtifactBlob(imported));
  });

  test('preserves bounded Word control properties on every OMML argument slot', async () => {
    const equation = argumentControlPropertiesEquation();
    expect(normalizeDocumentEquation(equation)).toEqual(equation);

    expect(
      normalizeDocumentEquation({
        version: 1,
        display: 'inline',
        children: [
          {
            type: 'box',
            operatorEmulator: false,
            noBreak: false,
            differential: false,
            alignment: false,
            children: [],
            childrenProperties: { controlProperties: {} },
          },
        ],
      }),
    ).toEqual({
      version: 1,
      display: 'inline',
      children: [
        {
          type: 'box',
          operatorEmulator: false,
          noBreak: false,
          differential: false,
          alignment: false,
          children: [],
        },
      ],
    });

    const emptyDynamicPropertyModels = [
      [
        {
          type: 'matrix',
          baseAlignment: 'center',
          placeholdersHidden: false,
          columnAlignments: ['center'],
          rows: [[[]]],
          cellProperties: [[null]],
        },
        'cellProperties',
      ],
      [
        {
          type: 'equationArray',
          baseAlignment: 'center',
          maximumDistribution: false,
          objectDistribution: false,
          rowSpacingRule: 'single',
          rowSpacing: 0,
          rows: [[]],
          rowProperties: [null],
        },
        'rowProperties',
      ],
      [
        {
          type: 'delimiter',
          opening: '(',
          closing: ')',
          separator: '|',
          arguments: [[]],
          argumentProperties: [null],
        },
        'argumentProperties',
      ],
    ] as const;
    for (const [expression, propertyName] of emptyDynamicPropertyModels) {
      const normalized = normalizeDocumentEquation({
        version: 1,
        display: 'inline',
        children: [expression],
      } as WorkDocumentEquation);
      expect(normalized).not.toBeNull();
      expect(normalized?.children[0]).not.toHaveProperty(propertyName);
    }

    const invalidArgumentModels = [
      {
        type: 'box',
        operatorEmulator: false,
        noBreak: false,
        differential: false,
        alignment: false,
        children: [],
        childrenProperties: { extra: true },
      },
      {
        type: 'matrix',
        baseAlignment: 'center',
        placeholdersHidden: false,
        columnAlignments: ['center', 'center'],
        rows: [[[], []]],
        cellProperties: [[{ controlProperties: { bold: true } }]],
      },
      {
        type: 'equationArray',
        baseAlignment: 'center',
        maximumDistribution: false,
        objectDistribution: false,
        rowSpacingRule: 'single',
        rowSpacing: 0,
        rows: [[], []],
        rowProperties: [{ controlProperties: { bold: true } }],
      },
      {
        type: 'delimiter',
        opening: '(',
        closing: ')',
        separator: '|',
        arguments: [[], []],
        argumentProperties: [{ controlProperties: { bold: true } }],
      },
    ];
    expect(
      invalidArgumentModels.map((expression) =>
        normalizeDocumentEquation({
          version: 1,
          display: 'inline',
          children: [expression],
        } as WorkDocumentEquation),
      ),
    ).toEqual(invalidArgumentModels.map(() => null));

    const run = '<m:r><m:t>x</m:t></m:r>';
    const argument = (content: string) =>
      `<m:box><m:e xmlns:w="${WORD_NAMESPACE}" xmlns:r="${RELATIONSHIP_NAMESPACE}" xmlns:v="${VENDOR_NAMESPACE}">${content}</m:e></m:box>`;
    const richControlProperties =
      '<m:ctrlPr><w:rPr><w:b/><w:color w:val="1A2B3C"/><w:sz w:val="25"/></w:rPr></m:ctrlPr>';
    expect(
      inspectEquationModel(
        argument(
          `<m:argPr><m:argSz m:val="0"/></m:argPr>${run}${richControlProperties}`,
        ),
      )?.children[0],
    ).toEqual({
      type: 'box',
      operatorEmulator: false,
      noBreak: false,
      differential: false,
      alignment: false,
      children: [{ type: 'run', text: 'x' }],
      childrenProperties: {
        controlProperties: {
          bold: true,
          color: { value: '#1a2b3c' },
          fontSize: 12.5,
        },
      },
    });
    for (const controlProperties of [
      '<m:ctrlPr/>',
      '<m:ctrlPr><w:rPr/></m:ctrlPr>',
    ]) {
      expect(
        inspectEquationModel(argument(`${run}${controlProperties}`)),
      ).toEqual({
        version: 1,
        display: 'inline',
        children: [
          {
            type: 'box',
            operatorEmulator: false,
            noBreak: false,
            differential: false,
            alignment: false,
            children: [{ type: 'run', text: 'x' }],
          },
        ],
      });
    }

    const unsupported = [
      argument(`${richControlProperties}${run}`),
      argument(`${run}${richControlProperties}<m:ctrlPr/>`),
      argument(`${run}<m:ctrlPr m:extra="semantic"/>`),
      argument(`${run}<m:ctrlPr>meaningful</m:ctrlPr>`),
      argument(`${run}<m:ctrlPr><w:rPr/><w:rPr/></m:ctrlPr>`),
      argument(`${run}<m:ctrlPr><w:ins><w:rPr/></w:ins></m:ctrlPr>`),
      argument(`${run}<m:ctrlPr><w:del><w:rPr/></w:del></m:ctrlPr>`),
      argument(`${run}<m:ctrlPr><w:moveFrom><w:rPr/></w:moveFrom></m:ctrlPr>`),
      argument(`${run}<m:ctrlPr><w:moveTo><w:rPr/></w:moveTo></m:ctrlPr>`),
      argument(`${run}<m:ctrlPr><v:rPr/></m:ctrlPr>`),
      argument(`${run}<m:ctrlPr><m:rPr/></m:ctrlPr>`),
      argument(`${run}<m:ctrlPr r:id="rIdUnsafe"/>`),
      argument(`${run}<m:ctrlPr><w:rPr r:id="rIdUnsafe"/></m:ctrlPr>`),
      argument(
        `${run}<m:ctrlPr><w:rPr>${WORD_2010_SCENE_3D}</w:rPr></m:ctrlPr>`,
      ),
      argument(
        `${run}<m:ctrlPr><w:rPr><w:b/><w:rFonts w:ascii="Cambria Math"/></w:rPr></m:ctrlPr>`,
      ),
      argument(`${run}<m:ctrlPr><w:rPr><w:b/><w:b/></w:rPr></m:ctrlPr>`),
      argument(`${run}<m:ctrlPr><w:b/></m:ctrlPr>`),
    ];
    expect(unsupported.map(inspectEquationBody)).toEqual(
      unsupported.map(() => 'unsupported'),
    );

    const strictSource = `<m:oMath xmlns:m="${STRICT_MATH_NAMESPACE}" xmlns:w="${STRICT_WORD_NAMESPACE}"><m:box><m:e><m:r><m:t>x</m:t></m:r><m:ctrlPr><w:rPr><w:b/></w:rPr></m:ctrlPr></m:e></m:box></m:oMath>`;
    expect(inspectEquationRoot(strictSource)).toMatchObject({
      status: 'supported',
      equation: {
        children: [
          {
            childrenProperties: { controlProperties: { bold: true } },
          },
        ],
      },
    });

    const document = new DOMParser().parseFromString('', 'text/html');
    const preview = createDocumentEquationElement(document, equation);
    expect(preview.querySelectorAll('[mathcolor], [mathsize]')).toHaveLength(0);

    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    artifact.content.html = `<p>${equationHtml(equation)}</p>`;
    const first = await createArtifactBlob(artifact);
    await expectNativeArgumentControlProperties(first);

    const firstArchive = await JSZip.loadAsync(await first.arrayBuffer());
    const firstDocument = await xmlEntry(firstArchive, 'word/document.xml');
    const nativeEquation = descendants(firstDocument, 'oMath').find(
      (candidate) => candidate.namespaceURI === MATH_NAMESPACE,
    );
    expect(nativeEquation).toBeDefined();
    const standalone = nativeEquation?.cloneNode(true) as Element;
    standalone.setAttributeNS(XMLNS_NAMESPACE, 'xmlns', MATH_NAMESPACE);
    standalone.setAttributeNS(XMLNS_NAMESPACE, 'xmlns:m', MATH_NAMESPACE);
    standalone.setAttributeNS(XMLNS_NAMESPACE, 'xmlns:w', WORD_NAMESPACE);
    standalone.setAttributeNS(
      XMLNS_NAMESPACE,
      'xmlns:w14',
      WORD_2010_NAMESPACE,
    );
    const strictRoundTripSource = new XMLSerializer()
      .serializeToString(standalone)
      .replaceAll(MATH_NAMESPACE, STRICT_MATH_NAMESPACE)
      .replaceAll(WORD_NAMESPACE, STRICT_WORD_NAMESPACE);
    expect(inspectEquationRoot(strictRoundTripSource)).toMatchObject({
      status: 'supported',
      equation,
    });

    const imported = await importOfficeFile(
      new File([first], 'argument-control-properties.docx', {
        type: first.type,
      }),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const importedDocument = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    const importedEquation = importedDocument.body.querySelector<HTMLElement>(
      '[data-document-equation]',
    );
    expect(
      documentEquationFromElement(importedEquation as HTMLElement),
    ).toEqual(equation);
    expect(imported.compatibility.issues).not.toContainEqual(
      expect.objectContaining({ code: 'docx.equations.unsupported' }),
    );
    await expectNativeArgumentControlProperties(
      await createArtifactBlob(imported),
    );
  });

  test('preserves bounded Word math control revisions across object and argument slots', async () => {
    const equation = controlRevisionEquation();
    expect(normalizeDocumentEquation(equation)).toEqual(equation);

    expect(
      normalizeDocumentEquation({
        version: 1,
        display: 'inline',
        children: [
          {
            type: 'fraction',
            controlRevision: {
              kind: 'insertion',
              id: 7,
              author: '  Alice & Bob  ',
              date: '  2026-01-02T03:04:05+08:00  ',
              dateUtc: '  2026-01-01T19:04:05Z  ',
            },
            fractionType: 'bar',
            numerator: [],
            denominator: [],
          },
        ],
      } as unknown as WorkDocumentEquation),
    ).toEqual({
      version: 1,
      display: 'inline',
      children: [
        {
          type: 'fraction',
          controlRevision: {
            kind: 'insertion',
            id: 7,
            author: 'Alice & Bob',
            date: '2026-01-02T03:04:05+08:00',
            dateUtc: '2026-01-01T19:04:05Z',
          },
          fractionType: 'bar',
          numerator: [],
          denominator: [],
        },
      ],
    });

    const validRevision = {
      kind: 'insertion',
      id: 1,
      author: 'Reviewer',
    };
    const invalidRevisions = [
      {},
      { ...validRevision, extra: true },
      { ...validRevision, kind: 'replacement' },
      { ...validRevision, id: -1 },
      { ...validRevision, id: 2_147_483_648 },
      { ...validRevision, id: 1.5 },
      { ...validRevision, id: '1' },
      { ...validRevision, author: '' },
      { ...validRevision, author: 'x'.repeat(256) },
      { ...validRevision, author: 'unsafe\u0000author' },
      { ...validRevision, date: '' },
      { ...validRevision, date: 'not-a-date' },
      { ...validRevision, date: '2026-02-30T03:04:05Z' },
      { ...validRevision, dateUtc: '' },
      { ...validRevision, dateUtc: 'not-a-date' },
      { ...validRevision, dateUtc: '2026-01-02T03:04:05+08:00' },
      { ...validRevision, child: null },
      { ...validRevision, child: validRevision },
      {
        kind: 'deletion',
        id: 2,
        author: 'Reviewer',
        child: { ...validRevision, kind: 'deletion' },
      },
      {
        kind: 'moveFrom',
        id: 2,
        author: 'Reviewer',
        child: { ...validRevision, kind: 'moveTo' },
      },
      {
        kind: 'moveTo',
        id: 2,
        author: 'Reviewer',
        child: { ...validRevision, kind: 'moveFrom' },
      },
    ];
    for (const controlRevision of invalidRevisions) {
      expect(
        normalizeDocumentEquation({
          version: 1,
          display: 'inline',
          children: [
            {
              type: 'fraction',
              controlRevision,
              fractionType: 'bar',
              numerator: [],
              denominator: [],
            },
          ],
        } as unknown as WorkDocumentEquation),
      ).toBeNull();
    }
    expect(
      normalizeDocumentEquation({
        version: 1,
        display: 'inline',
        children: [
          {
            type: 'run',
            text: 'x',
            controlRevision: validRevision,
          },
        ],
      } as unknown as WorkDocumentEquation),
    ).toBeNull();

    const run = '<m:r><m:t>x</m:t></m:r>';
    const revisedObject = (content: string) =>
      `<m:f><m:fPr><m:ctrlPr xmlns:w="${WORD_NAMESPACE}" xmlns:w16du="${WORD_DATE_UTC_NAMESPACE}" xmlns:r="${RELATIONSHIP_NAMESPACE}" xmlns:v="${VENDOR_NAMESPACE}">${content}</m:ctrlPr></m:fPr><m:num>${run}</m:num><m:den>${run}</m:den></m:f>`;
    expect(
      inspectEquationModel(
        revisedObject(
          '<w:ins w:id="+0007" w:author=" Alice &amp; Bob " w:date="2026-01-02T03:04:05+08:00" w16du:dateUtc="2026-01-01T19:04:05Z"><w:del w:id="2147483647" w:author="Final reviewer"><w:rPr><w:b/></w:rPr></w:del></w:ins>',
        ),
      )?.children[0],
    ).toEqual({
      type: 'fraction',
      controlRevision: {
        kind: 'insertion',
        id: 7,
        author: 'Alice & Bob',
        date: '2026-01-02T03:04:05+08:00',
        dateUtc: '2026-01-01T19:04:05Z',
        child: {
          kind: 'deletion',
          id: 2_147_483_647,
          author: 'Final reviewer',
        },
      },
      controlProperties: { bold: true },
      fractionType: 'bar',
      numerator: [{ type: 'run', text: 'x' }],
      denominator: [{ type: 'run', text: 'x' }],
    });

    const malformed = [
      '<w:ins w:id="1"><w:rPr/></w:ins>',
      '<w:ins w:author="Reviewer"><w:rPr/></w:ins>',
      '<w:ins id="1" w:author="Reviewer"><w:rPr/></w:ins>',
      '<w:ins w:id="1" author="Reviewer"><w:rPr/></w:ins>',
      '<w:ins r:id="1" w:author="Reviewer"><w:rPr/></w:ins>',
      '<w:ins w:id="1" w:author="Reviewer" w:extra="semantic"><w:rPr/></w:ins>',
      '<w:ins w:id="-1" w:author="Reviewer"><w:rPr/></w:ins>',
      '<w:ins w:id="2147483648" w:author="Reviewer"><w:rPr/></w:ins>',
      '<w:ins w:id="1.5" w:author="Reviewer"><w:rPr/></w:ins>',
      '<w:ins w:id="1" w:author=""><w:rPr/></w:ins>',
      `<w:ins w:id="1" w:author="${'x'.repeat(256)}"><w:rPr/></w:ins>`,
      '<w:ins w:id="1" w:author="Reviewer" w:date="not-a-date"><w:rPr/></w:ins>',
      '<w:ins w:id="1" w:author="Reviewer" v:dateUtc="2026-01-02T03:04:05Z"><w:rPr/></w:ins>',
      '<w:ins w:id="1" w:author="Reviewer" dateUtc="2026-01-02T03:04:05Z"><w:rPr/></w:ins>',
      '<w:ins w:id="1" w:author="Reviewer" w:dateUtc="2026-01-02T03:04:05Z"><w:rPr/></w:ins>',
      '<w:ins w:id="1" w:author="Reviewer" w16du:dateUtc="2026-01-02T03:04:05+08:00"><w:rPr/></w:ins>',
      '<w:ins w:id="1" w:author="Reviewer" w16du:extra="semantic"><w:rPr/></w:ins>',
      '<w:ins w:id="1" w:author="Reviewer">meaningful<w:rPr/></w:ins>',
      '<w:ins w:id="1" w:author="Reviewer"><w:rPr/><w:del w:id="2" w:author="Reviewer"/></w:ins>',
      '<w:ins w:id="1" w:author="Reviewer"><w:ins w:id="2" w:author="Reviewer"/></w:ins>',
      '<w:del w:id="1" w:author="Reviewer"><w:ins w:id="2" w:author="Reviewer"/></w:del>',
      '<w:moveFrom w:id="1" w:author="Reviewer"><w:moveTo w:id="2" w:author="Reviewer"/></w:moveFrom>',
      '<w:moveTo w:id="1" w:author="Reviewer"><w:moveFrom w:id="2" w:author="Reviewer"/></w:moveTo>',
      '<v:ins v:id="1" v:author="Reviewer"><w:rPr/></v:ins>',
    ];
    expect(malformed.map(revisedObject).map(inspectEquationBody)).toEqual(
      malformed.map(() => 'unsupported'),
    );

    const strictSource = `<m:oMath xmlns:m="${STRICT_MATH_NAMESPACE}" xmlns:w="${STRICT_WORD_NAMESPACE}" xmlns:w16du="${WORD_DATE_UTC_NAMESPACE}"><m:box><m:boxPr><m:ctrlPr><w:moveTo w:id="9" w:author="Strict reviewer" w16du:dateUtc="2026-03-04T05:06:07Z"><w:del w:id="10" w:author="Strict reviewer"><w:rPr><w:i/></w:rPr></w:del></w:moveTo></m:ctrlPr></m:boxPr><m:e>${run}</m:e></m:box></m:oMath>`;
    expect(inspectEquationRoot(strictSource)).toMatchObject({
      status: 'supported',
      equation: {
        children: [
          {
            controlRevision: {
              kind: 'moveTo',
              id: 9,
              author: 'Strict reviewer',
              dateUtc: '2026-03-04T05:06:07Z',
              child: {
                kind: 'deletion',
                id: 10,
                author: 'Strict reviewer',
              },
            },
            controlProperties: { italic: true },
          },
        ],
      },
    });

    const document = new DOMParser().parseFromString('', 'text/html');
    const preview = createDocumentEquationElement(document, equation);
    const slash = Array.from(preview.querySelectorAll('mo')).find(
      (operator) => operator.textContent === '/',
    );
    expect(slash?.getAttribute('mathcolor')).toBe('#1a2b3c');
    expect(preview.querySelectorAll('ins, del, movefrom, moveto')).toHaveLength(
      0,
    );
    expect(preview.querySelectorAll('[data-document-change]')).toHaveLength(0);

    const everyObjectRevision = everyObjectControlRevisionEquation();
    expect(normalizeDocumentEquation(everyObjectRevision)).toEqual(
      everyObjectRevision,
    );
    const everyObjectArtifact = createArtifact('blank-document');
    if (everyObjectArtifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    everyObjectArtifact.content.html = `<p>${equationHtml(everyObjectRevision)}</p>`;
    await expectNativeControlRevisionOnEveryObject(
      await createArtifactBlob(everyObjectArtifact),
    );

    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    artifact.content.html = `<p>${equationHtml(equation)}</p>`;
    const first = await createArtifactBlob(artifact);
    await expectNativeControlRevisions(first);

    const firstArchive = await JSZip.loadAsync(await first.arrayBuffer());
    const firstDocument = await xmlEntry(firstArchive, 'word/document.xml');
    const nativeEquation = descendants(firstDocument, 'oMath').find(
      (candidate) => candidate.namespaceURI === MATH_NAMESPACE,
    );
    expect(nativeEquation).toBeDefined();
    const standalone = nativeEquation?.cloneNode(true) as Element;
    standalone.setAttributeNS(XMLNS_NAMESPACE, 'xmlns', MATH_NAMESPACE);
    standalone.setAttributeNS(XMLNS_NAMESPACE, 'xmlns:m', MATH_NAMESPACE);
    standalone.setAttributeNS(XMLNS_NAMESPACE, 'xmlns:w', WORD_NAMESPACE);
    standalone.setAttributeNS(
      XMLNS_NAMESPACE,
      'xmlns:w14',
      WORD_2010_NAMESPACE,
    );
    standalone.setAttributeNS(
      XMLNS_NAMESPACE,
      'xmlns:w16du',
      WORD_DATE_UTC_NAMESPACE,
    );
    const strictRoundTripSource = new XMLSerializer()
      .serializeToString(standalone)
      .replaceAll(MATH_NAMESPACE, STRICT_MATH_NAMESPACE)
      .replaceAll(WORD_NAMESPACE, STRICT_WORD_NAMESPACE);
    expect(inspectEquationRoot(strictRoundTripSource)).toMatchObject({
      status: 'supported',
      equation,
    });

    const imported = await importOfficeFile(
      new File([first], 'control-revisions.docx', { type: first.type }),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const importedDocument = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    const importedEquation = importedDocument.body.querySelector<HTMLElement>(
      '[data-document-equation]',
    );
    expect(
      documentEquationFromElement(importedEquation as HTMLElement),
    ).toEqual(equation);
    expect(imported.compatibility.issues).not.toContainEqual(
      expect.objectContaining({ code: 'docx.equations.unsupported' }),
    );
    await expectNativeControlRevisions(await createArtifactBlob(imported));
  });

  test('preserves bounded relative argument sizes and projects only Word-effective slots', async () => {
    const equation = argumentSizesEquation();
    expect(normalizeDocumentEquation(equation)).toEqual(equation);

    expect(
      normalizeDocumentEquation({
        version: 1,
        display: 'inline',
        children: [
          {
            type: 'box',
            operatorEmulator: false,
            noBreak: false,
            differential: false,
            alignment: false,
            children: [],
            childrenProperties: {
              size: 0,
              controlProperties: { bold: true },
            },
          },
        ],
      } as unknown as WorkDocumentEquation),
    ).toEqual({
      version: 1,
      display: 'inline',
      children: [
        {
          type: 'box',
          operatorEmulator: false,
          noBreak: false,
          differential: false,
          alignment: false,
          children: [],
          childrenProperties: { controlProperties: { bold: true } },
        },
      ],
    });
    for (const size of [-3, 3, -1.5, Number.NaN, '1', null]) {
      expect(
        normalizeDocumentEquation({
          version: 1,
          display: 'inline',
          children: [
            {
              type: 'box',
              operatorEmulator: false,
              noBreak: false,
              differential: false,
              alignment: false,
              children: [],
              childrenProperties: { size },
            },
          ],
        } as unknown as WorkDocumentEquation),
      ).toBeNull();
    }

    const run = '<m:r><m:t>x</m:t></m:r>';
    const sizedBox = (value: string) =>
      `<m:box><m:e><m:argPr><m:argSz m:val="${value}"/></m:argPr>${run}</m:e></m:box>`;
    expect(
      ['-2', '-1', '1', '+2'].map(sizedBox).map(inspectEquationBody),
    ).toEqual(['supported', 'supported', 'supported', 'supported']);
    expect(
      ['-2', '-1', '1', '+2'].map(sizedBox).map((source) => {
        const expression = inspectEquationModel(source)?.children[0];
        return expression?.type === 'box'
          ? expression.childrenProperties?.size
          : undefined;
      }),
    ).toEqual([-2, -1, 1, 2]);

    const malformed = [
      sizedBox('-3'),
      sizedBox('3'),
      sizedBox('1.0'),
      sizedBox('maybe'),
      sizedBox('9007199254740992'),
      `<m:box><m:e><m:argPr><m:argSz/><m:argSz/></m:argPr>${run}</m:e></m:box>`,
      `<m:box><m:e><m:argPr m:extra="semantic"><m:argSz m:val="1"/></m:argPr>${run}</m:e></m:box>`,
      `<m:box><m:e><m:argPr>meaningful<m:argSz m:val="1"/></m:argPr>${run}</m:e></m:box>`,
      `<m:box><m:e><m:argPr><m:argSz m:val="1" m:extra="semantic"/></m:argPr>${run}</m:e></m:box>`,
      `<m:box xmlns:r="${RELATIONSHIP_NAMESPACE}"><m:e><m:argPr><m:argSz m:val="1" r:id="rIdUnsafe"/></m:argPr>${run}</m:e></m:box>`,
      `<m:box><m:e><m:argPr><m:argSz val="1"/></m:argPr>${run}</m:e></m:box>`,
      `<m:box><m:e><m:argPr><v:argSz xmlns:v="${VENDOR_NAMESPACE}" m:val="1"/></m:argPr>${run}</m:e></m:box>`,
      `<m:box><m:e><m:argPr><m:argSz m:val="1"><m:r/></m:argSz></m:argPr>${run}</m:e></m:box>`,
      `<m:box><m:e>${run}<m:argPr><m:argSz m:val="1"/></m:argPr></m:e></m:box>`,
    ];
    expect(malformed.map(inspectEquationBody)).toEqual(
      malformed.map(() => 'unsupported'),
    );

    const strictSource = `<m:oMath xmlns:m="${STRICT_MATH_NAMESPACE}"><m:box><m:e><m:argPr><m:argSz m:val="-2"/></m:argPr>${run}</m:e></m:box></m:oMath>`;
    expect(inspectEquationRoot(strictSource)).toMatchObject({
      status: 'supported',
      equation: {
        children: [{ childrenProperties: { size: -2 } }],
      },
    });

    const document = new DOMParser().parseFromString('', 'text/html');
    const preview = createDocumentEquationElement(document, equation);
    expect(
      Array.from(preview.querySelectorAll('mstyle[scriptlevel]')).map(
        (style) => [style.textContent, style.getAttribute('scriptlevel')],
      ),
    ).toEqual([
      ['box-sized', '+2'],
      ['group-sized', '+1'],
      ['lower-limit-sized', '-1'],
      ['upper-limit-sized', '-2'],
      ['nary-sub-sized', '+2'],
      ['nary-sup-sized', '+1'],
      ['radical-degree-sized', '-1'],
      ['pre-sub-sized', '-2'],
      ['pre-sup-sized', '+2'],
      ['subscript-sized', '+1'],
      ['sub-super-sub-sized', '-1'],
      ['sub-super-sup-sized', '-2'],
      ['superscript-sized', '+2'],
    ]);
    for (const text of [
      'fraction-noop-sized',
      'function-noop-sized',
      'matrix-noop-sized',
      'array-noop-sized',
      'delimiter-noop-sized',
    ]) {
      const candidate = Array.from(preview.querySelectorAll('mtext')).find(
        (element) => element.textContent === text,
      );
      expect(candidate, text).toBeDefined();
      expect(candidate?.closest('mstyle'), text).toBeNull();
    }
    const sanitized = new DOMParser().parseFromString(
      sanitizeDocumentPageChromeHtml(equationHtml(equation)),
      'text/html',
    );
    expect(sanitized.body.querySelectorAll('mstyle[scriptlevel]')).toHaveLength(
      13,
    );

    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    artifact.content.html = `<p>${equationHtml(equation)}</p>`;
    const first = await createArtifactBlob(artifact);
    await expectNativeArgumentSizes(first);

    const firstArchive = await JSZip.loadAsync(await first.arrayBuffer());
    const firstDocument = await xmlEntry(firstArchive, 'word/document.xml');
    const nativeEquation = descendants(firstDocument, 'oMath').find(
      (candidate) => candidate.namespaceURI === MATH_NAMESPACE,
    );
    expect(nativeEquation).toBeDefined();
    const standalone = nativeEquation?.cloneNode(true) as Element;
    standalone.setAttributeNS(XMLNS_NAMESPACE, 'xmlns', MATH_NAMESPACE);
    standalone.setAttributeNS(XMLNS_NAMESPACE, 'xmlns:m', MATH_NAMESPACE);
    standalone.setAttributeNS(XMLNS_NAMESPACE, 'xmlns:w', WORD_NAMESPACE);
    standalone.setAttributeNS(
      XMLNS_NAMESPACE,
      'xmlns:w14',
      WORD_2010_NAMESPACE,
    );
    const strictRoundTripSource = new XMLSerializer()
      .serializeToString(standalone)
      .replaceAll(MATH_NAMESPACE, STRICT_MATH_NAMESPACE)
      .replaceAll(WORD_NAMESPACE, STRICT_WORD_NAMESPACE);
    expect(inspectEquationRoot(strictRoundTripSource)).toMatchObject({
      status: 'supported',
      equation,
    });

    const imported = await importOfficeFile(
      new File([first], 'argument-sizes.docx', { type: first.type }),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const importedDocument = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    const importedEquation = importedDocument.body.querySelector<HTMLElement>(
      '[data-document-equation]',
    );
    expect(
      documentEquationFromElement(importedEquation as HTMLElement),
    ).toEqual(equation);
    expect(imported.compatibility.issues).not.toContainEqual(
      expect.objectContaining({ code: 'docx.equations.unsupported' }),
    );
    await expectNativeArgumentSizes(await createArtifactBlob(imported));
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
    ];
    expect(unsupported.map(inspectEquationBody)).toEqual(
      unsupported.map(() => 'unsupported'),
    );
  });

  test('preserves empty function slots and strictly validates function arguments', () => {
    const run = '<m:r><m:t>x</m:t></m:r>';
    const supported = [
      `<m:func><m:fName>${run}</m:fName><m:e>${run}</m:e></m:func>`,
      '<m:func><m:funcPr/><m:fName/><m:e/></m:func>',
      `<m:func><m:funcPr><m:ctrlPr/></m:funcPr><m:fName><m:argPr/>${run}<m:ctrlPr/></m:fName><m:e><m:argPr/>${run}<m:ctrlPr/></m:e></m:func>`,
      '<m:func><m:fName><m:argPr/><m:ctrlPr/></m:fName><m:e><m:argPr/><m:ctrlPr/></m:e></m:func>',
    ];
    expect(supported.map(inspectEquationBody)).toEqual(
      supported.map(() => 'supported'),
    );
    expect(
      supported.map((source) => {
        const expression = inspectEquationModel(source)?.children[0];
        return expression?.type === 'function'
          ? [expression.name, expression.children]
          : null;
      }),
    ).toEqual([
      [[{ type: 'run', text: 'x' }], [{ type: 'run', text: 'x' }]],
      [[], []],
      [[{ type: 'run', text: 'x' }], [{ type: 'run', text: 'x' }]],
      [[], []],
    ]);
    expect(
      normalizeDocumentEquation({
        version: 1,
        display: 'inline',
        children: [{ type: 'function', name: [], children: [] }],
      }),
    ).toEqual({
      version: 1,
      display: 'inline',
      children: [{ type: 'function', name: [], children: [] }],
    });
    expect(
      inspectEquationRoot(
        `<m:oMath xmlns:m="${STRICT_MATH_NAMESPACE}"><m:func><m:fName/><m:e/></m:func></m:oMath>`,
      ).status,
    ).toBe('supported');

    const unsupported = [
      `<m:func><m:e>${run}</m:e></m:func>`,
      `<m:func><m:fName>${run}</m:fName></m:func>`,
      `<m:func><m:e>${run}</m:e><m:fName>${run}</m:fName></m:func>`,
      `<m:func><m:fName>${run}</m:fName><m:funcPr/><m:e>${run}</m:e></m:func>`,
      `<m:func><m:funcPr/><m:funcPr/><m:fName>${run}</m:fName><m:e>${run}</m:e></m:func>`,
      `<m:func><m:fName>${run}</m:fName><m:fName>${run}</m:fName><m:e>${run}</m:e></m:func>`,
      `<m:func><m:fName>${run}</m:fName><m:e>${run}</m:e><m:e>${run}</m:e></m:func>`,
      `<m:func><m:funcPr><m:ctrlPr/><m:ctrlPr/></m:funcPr><m:fName>${run}</m:fName><m:e>${run}</m:e></m:func>`,
      `<m:func><m:funcPr><m:ctrlPr xmlns:r="${RELATIONSHIP_NAMESPACE}" r:id="rIdUnsafe"/></m:funcPr><m:fName>${run}</m:fName><m:e>${run}</m:e></m:func>`,
      `<m:func><m:funcPr><m:unknown/></m:funcPr><m:fName>${run}</m:fName><m:e>${run}</m:e></m:func>`,
      `<m:func><m:fName><m:ctrlPr/>${run}</m:fName><m:e>${run}</m:e></m:func>`,
      `<m:func><m:fName>${run}<m:argPr/></m:fName><m:e>${run}</m:e></m:func>`,
      `<m:func><m:fName><m:argPr/><m:argPr/>${run}</m:fName><m:e>${run}</m:e></m:func>`,
      `<m:func><m:fName>${run}<m:ctrlPr/><m:r><m:t>y</m:t></m:r></m:fName><m:e>${run}</m:e></m:func>`,
      `<m:func><m:fName>${run}<m:ctrlPr/><m:ctrlPr/></m:fName><m:e>${run}</m:e></m:func>`,
      `<m:func><m:fName><m:argPr><m:argSz m:val="3"/></m:argPr>${run}</m:fName><m:e>${run}</m:e></m:func>`,
      `<m:func><m:fName><v:argPr xmlns:v="${VENDOR_NAMESPACE}"/>${run}</m:fName><m:e>${run}</m:e></m:func>`,
      `<m:func m:extra="semantic"><m:fName>${run}</m:fName><m:e>${run}</m:e></m:func>`,
      `<m:func>meaningful<m:fName>${run}</m:fName><m:e>${run}</m:e></m:func>`,
      `<m:func><v:fName xmlns:v="${VENDOR_NAMESPACE}">${run}</v:fName><m:e>${run}</m:e></m:func>`,
    ];
    expect(unsupported.map(inspectEquationBody)).toEqual(
      unsupported.map(() => 'unsupported'),
    );
  });

  test('preserves empty math arguments and canonicalizes default argument size', () => {
    const emptyArguments = [
      '<m:acc><m:e/></m:acc>',
      '<m:bar><m:e/></m:bar>',
      '<m:borderBox><m:e/></m:borderBox>',
      '<m:box><m:e/></m:box>',
      '<m:f><m:num/><m:den/></m:f>',
      '<m:sSup><m:e/><m:sup/></m:sSup>',
      '<m:sSub><m:e/><m:sub/></m:sSub>',
      '<m:sSubSup><m:e/><m:sub/><m:sup/></m:sSubSup>',
      '<m:sPre><m:sub/><m:sup/><m:e/></m:sPre>',
      '<m:limLow><m:e/><m:lim/></m:limLow>',
      '<m:limUpp><m:e/><m:lim/></m:limUpp>',
      '<m:rad><m:deg/><m:e/></m:rad>',
      '<m:func><m:fName/><m:e/></m:func>',
      '<m:nary><m:naryPr><m:subHide/><m:supHide/></m:naryPr><m:sub/><m:sup/><m:e/></m:nary>',
      '<m:groupChr><m:e/></m:groupChr>',
      '<m:phant><m:e/></m:phant>',
    ];
    expect(emptyArguments.map(inspectEquationBody)).toEqual(
      emptyArguments.map(() => 'supported'),
    );
    expect(
      emptyArguments.map(
        (source) => inspectEquationModel(source)?.children[0]?.type,
      ),
    ).toEqual([
      'accent',
      'bar',
      'borderBox',
      'box',
      'fraction',
      'superscript',
      'subscript',
      'subSuperScript',
      'preSubSuperScript',
      'lowerLimit',
      'upperLimit',
      'radical',
      'function',
      'nary',
      'groupCharacter',
      'phantom',
    ]);

    const defaultArgumentSizes = [
      '<m:box><m:e><m:argPr/></m:e></m:box>',
      '<m:box><m:e><m:argPr><m:argSz/></m:argPr></m:e></m:box>',
      '<m:box><m:e><m:argPr><m:argSz m:val="0"/></m:argPr></m:e></m:box>',
      '<m:box><m:e><m:argPr><m:argSz m:val="+0"/></m:argPr><m:ctrlPr/></m:e></m:box>',
      '<m:box><m:e><m:argPr><m:argSz m:val="-0"/></m:argPr></m:e></m:box>',
    ];
    expect(defaultArgumentSizes.map(inspectEquationBody)).toEqual(
      defaultArgumentSizes.map(() => 'supported'),
    );
    expect(
      defaultArgumentSizes.map(
        (source) => inspectEquationModel(source)?.children[0],
      ),
    ).toEqual(
      defaultArgumentSizes.map(() => ({
        type: 'box',
        operatorEmulator: false,
        noBreak: false,
        differential: false,
        alignment: false,
        children: [],
      })),
    );

    const emptyFraction: WorkDocumentEquation = {
      version: 1,
      display: 'inline',
      children: [
        {
          type: 'fraction',
          fractionType: 'bar',
          numerator: [],
          denominator: [],
        },
      ],
    };
    expect(normalizeDocumentEquation(emptyFraction)).toEqual(emptyFraction);

    const unsupported = ['-3', '3', '1.0', 'maybe'].map(
      (value) =>
        `<m:box><m:e><m:argPr><m:argSz m:val="${value}"/></m:argPr></m:e></m:box>`,
    );
    unsupported.push(
      '<m:box><m:e><m:argPr><m:argSz/><m:argSz/></m:argPr></m:e></m:box>',
      '<m:box><m:e><m:argPr><m:unknown/></m:argPr></m:e></m:box>',
      '<m:box><m:e><m:argPr m:extra="semantic"/></m:e></m:box>',
      '<m:box><m:e><m:argPr>meaningful</m:argPr></m:e></m:box>',
      `<m:box><m:e><m:argPr><m:argSz xmlns:r="${RELATIONSHIP_NAMESPACE}" m:val="0" r:id="rIdUnsafe"/></m:argPr></m:e></m:box>`,
      `<m:box><m:e><m:argPr><v:argSz xmlns:v="${VENDOR_NAMESPACE}" m:val="0"/></m:argPr></m:e></m:box>`,
    );
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
      `<m:f><m:fPr>meaningful<m:type/></m:fPr><m:num>${run}</m:num><m:den>${run}</m:den></m:f>`,
      `<m:f><m:num>meaningful</m:num><m:den>${run}</m:den></m:f>`,
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
      `<m:rad><m:radPr>meaningful<m:degHide/></m:radPr><m:e>${run}</m:e></m:rad>`,
      `<m:rad><m:radPr><m:degHide/></m:radPr><m:deg>${run}</m:deg><m:e>${run}</m:e></m:rad>`,
      `<m:rad><m:deg>meaningful</m:deg><m:e>${run}</m:e></m:rad>`,
      `<m:rad><m:deg>${run}</m:deg></m:rad>`,
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
      `<m:nary><m:naryPr><m:grow/></m:naryPr><m:sub>${run}</m:sub><m:sup>${run}</m:sup><m:e>${run}</m:e></m:nary>`,
      `<m:nary><m:naryPr><m:grow m:val="true"/></m:naryPr><m:sub>${run}</m:sub><m:sup>${run}</m:sup><m:e>${run}</m:e></m:nary>`,
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
              expression.grow ?? false,
              expression.subScript?.[0]?.type ?? null,
              expression.superScript?.[0]?.type ?? null,
            ]
          : null;
      }),
    ).toEqual([
      ['\u222b', 'subSup', false, 'run', 'run'],
      ['\u222b', 'subSup', false, 'run', 'run'],
      ['\u2211', 'underOver', false, 'run', 'run'],
      ['\u2211', 'underOver', false, 'run', 'run'],
      ['\u2211', 'subSup', false, null, null],
      ['\u220f', 'underOver', false, null, 'run'],
      ['\u222b', 'underOver', false, 'run', 'run'],
      ['\u222b', 'subSup', true, 'run', 'run'],
      ['\u222b', 'subSup', true, 'run', 'run'],
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
      `<m:nary><m:naryPr><m:chr/></m:naryPr><m:sub>${run}</m:sub><m:sup>${run}</m:sup><m:e>${run}</m:e></m:nary>`,
      `<m:nary><m:naryPr><m:chr m:val="+"/></m:naryPr><m:sub>${run}</m:sub><m:sup>${run}</m:sup><m:e>${run}</m:e></m:nary>`,
      `<m:nary><m:naryPr><m:limLoc/><m:chr m:val="&#x2211;"/></m:naryPr><m:sub>${run}</m:sub><m:sup>${run}</m:sup><m:e>${run}</m:e></m:nary>`,
      `<m:nary><m:naryPr><m:supHide/><m:subHide/></m:naryPr><m:sub/><m:sup/><m:e>${run}</m:e></m:nary>`,
      `<m:nary><m:naryPr><m:ctrlPr/><m:subHide/></m:naryPr><m:sub/><m:sup>${run}</m:sup><m:e>${run}</m:e></m:nary>`,
      `<m:nary><m:naryPr><m:chr m:val="&#x2211;"/><m:chr m:val="&#x220F;"/></m:naryPr><m:sub>${run}</m:sub><m:sup>${run}</m:sup><m:e>${run}</m:e></m:nary>`,
      `<m:nary><m:naryPr><m:limLoc m:val="beside"/></m:naryPr><m:sub>${run}</m:sub><m:sup>${run}</m:sup><m:e>${run}</m:e></m:nary>`,
      `<m:nary><m:naryPr><m:grow m:val="maybe"/></m:naryPr><m:sub>${run}</m:sub><m:sup>${run}</m:sup><m:e>${run}</m:e></m:nary>`,
      `<m:nary><m:naryPr><m:grow/><m:grow/></m:naryPr><m:sub>${run}</m:sub><m:sup>${run}</m:sup><m:e>${run}</m:e></m:nary>`,
      `<m:nary><m:naryPr><m:grow/><m:limLoc/></m:naryPr><m:sub>${run}</m:sub><m:sup>${run}</m:sup><m:e>${run}</m:e></m:nary>`,
      `<m:nary><m:naryPr><m:grow val="1"/></m:naryPr><m:sub>${run}</m:sub><m:sup>${run}</m:sup><m:e>${run}</m:e></m:nary>`,
      `<m:nary><m:naryPr><m:grow m:val="1" m:extra="semantic"/></m:naryPr><m:sub>${run}</m:sub><m:sup>${run}</m:sup><m:e>${run}</m:e></m:nary>`,
      `<m:nary xmlns:r="${RELATIONSHIP_NAMESPACE}"><m:naryPr><m:grow r:id="rIdUnsafe"/></m:naryPr><m:sub>${run}</m:sub><m:sup>${run}</m:sup><m:e>${run}</m:e></m:nary>`,
      `<m:nary><m:naryPr><v:grow xmlns:v="${VENDOR_NAMESPACE}" v:val="1"/></m:naryPr><m:sub>${run}</m:sub><m:sup>${run}</m:sup><m:e>${run}</m:e></m:nary>`,
      `<m:nary><m:naryPr><m:subHide/></m:naryPr><m:sub>${run}</m:sub><m:sup>${run}</m:sup><m:e>${run}</m:e></m:nary>`,
      `<m:nary><m:naryPr><m:supHide/></m:naryPr><m:sub>${run}</m:sub><m:sup>${run}</m:sup><m:e>${run}</m:e></m:nary>`,
      `<m:nary><m:naryPr><m:subHide m:val="maybe"/></m:naryPr><m:sub/><m:sup>${run}</m:sup><m:e>${run}</m:e></m:nary>`,
      `<m:nary><m:naryPr><m:chr m:val="&#x2211;" m:extra="semantic"/></m:naryPr><m:sub>${run}</m:sub><m:sup>${run}</m:sup><m:e>${run}</m:e></m:nary>`,
      `<m:nary xmlns:r="${RELATIONSHIP_NAMESPACE}"><m:naryPr><m:chr m:val="&#x2211;" r:id="rIdUnsafe"/></m:naryPr><m:sub>${run}</m:sub><m:sup>${run}</m:sup><m:e>${run}</m:e></m:nary>`,
      `<m:nary><v:naryPr xmlns:v="${VENDOR_NAMESPACE}"><m:chr m:val="&#x2211;"/></v:naryPr><m:sub>${run}</m:sub><m:sup>${run}</m:sup><m:e>${run}</m:e></m:nary>`,
      `<m:nary><m:naryPr><v:chr xmlns:v="${VENDOR_NAMESPACE}" v:val="&#x2211;"/></m:naryPr><m:sub>${run}</m:sub><m:sup>${run}</m:sup><m:e>${run}</m:e></m:nary>`,
      `<m:nary><m:naryPr>meaningful<m:chr m:val="&#x2211;"/></m:naryPr><m:sub>${run}</m:sub><m:sup>${run}</m:sup><m:e>${run}</m:e></m:nary>`,
      `<m:nary m:extra="semantic"><m:sub>${run}</m:sub><m:sup>${run}</m:sup><m:e>${run}</m:e></m:nary>`,
    ];
    expect(unsupported.map(inspectEquationBody)).toEqual(
      unsupported.map(() => 'unsupported'),
    );
  });

  test('normalizes delimiter defaults and strictly validates delimiter structure', () => {
    const run = '<m:r><m:t>x</m:t></m:r>';
    const supported = [
      `<m:d><m:e>${run}</m:e></m:d>`,
      `<m:d><m:dPr/><m:e>${run}</m:e></m:d>`,
      `<m:d><m:dPr><m:begChr/><m:sepChr/><m:endChr/><m:grow/><m:shp/><m:ctrlPr/></m:dPr><m:e>${run}</m:e></m:d>`,
      `<m:d><m:dPr><m:begChr m:val="["/><m:sepChr m:val=";"/><m:endChr m:val="]"/><m:grow m:val="true"/><m:shp m:val="centered"/><m:ctrlPr/></m:dPr><m:e/><m:e>${run}</m:e><m:e/></m:d>`,
      `<m:d><m:dPr><m:begChr m:val=""/><m:sepChr m:val=""/><m:endChr m:val=""/></m:dPr><m:e>${run}</m:e></m:d>`,
      `<m:d><m:dPr><m:grow m:val="0"/></m:dPr><m:e>${run}</m:e></m:d>`,
      `<m:d><m:dPr><m:shp m:val="match"/></m:dPr><m:e>${run}</m:e></m:d>`,
    ];
    expect(supported.map(inspectEquationBody)).toEqual(
      supported.map(() => 'supported'),
    );
    expect(
      supported.map((source) => {
        const expression = inspectEquationModel(source)?.children[0];
        return expression?.type === 'delimiter'
          ? [
              expression.opening,
              expression.separator,
              expression.closing,
              expression.grow ?? true,
              expression.shape ?? 'centered',
              expression.arguments.map((argument) =>
                argument
                  .map((child) => (child.type === 'run' ? child.text : '?'))
                  .join(''),
              ),
            ]
          : null;
      }),
    ).toEqual([
      ['(', '\u2502', ')', true, 'centered', ['x']],
      ['(', '\u2502', ')', true, 'centered', ['x']],
      ['', '', '', true, 'centered', ['x']],
      ['[', ';', ']', true, 'centered', ['', 'x', '']],
      ['', '', '', true, 'centered', ['x']],
      ['(', '\u2502', ')', false, 'centered', ['x']],
      ['(', '\u2502', ')', true, 'match', ['x']],
    ]);

    const unsupported = [
      `<m:d><m:e>${run}</m:e><m:dPr/></m:d>`,
      `<m:d><m:dPr/><m:dPr/><m:e>${run}</m:e></m:d>`,
      '<m:d><m:dPr/></m:d>',
      `<m:d><m:dPr><m:endChr/><m:sepChr/></m:dPr><m:e>${run}</m:e></m:d>`,
      `<m:d><m:dPr><m:begChr/><m:begChr/></m:dPr><m:e>${run}</m:e></m:d>`,
      `<m:d><m:dPr><m:sepChr/><m:sepChr/></m:dPr><m:e>${run}</m:e></m:d>`,
      `<m:d><m:dPr><m:grow/><m:grow/></m:dPr><m:e>${run}</m:e></m:d>`,
      `<m:d><m:dPr><m:shp/><m:shp/></m:dPr><m:e>${run}</m:e></m:d>`,
      `<m:d><m:dPr><m:ctrlPr/><m:ctrlPr/></m:dPr><m:e>${run}</m:e></m:d>`,
      `<m:d><m:dPr><m:begChr m:val="xy"/></m:dPr><m:e>${run}</m:e></m:d>`,
      `<m:d><m:dPr><m:endChr m:val="&#x7f;"/></m:dPr><m:e>${run}</m:e></m:d>`,
      `<m:d><m:dPr><m:grow m:val="maybe"/></m:dPr><m:e>${run}</m:e></m:d>`,
      `<m:d><m:dPr><m:shp m:val="round"/></m:dPr><m:e>${run}</m:e></m:d>`,
      `<m:d><m:dPr><m:shp/><m:grow/></m:dPr><m:e>${run}</m:e></m:d>`,
      `<m:d><m:dPr><m:grow val="0"/></m:dPr><m:e>${run}</m:e></m:d>`,
      `<m:d><m:dPr><m:shp m:val="match" m:extra="semantic"/></m:dPr><m:e>${run}</m:e></m:d>`,
      `<m:d xmlns:r="${RELATIONSHIP_NAMESPACE}"><m:dPr><m:shp r:id="rIdUnsafe"/></m:dPr><m:e>${run}</m:e></m:d>`,
      `<m:d><m:dPr><v:shp xmlns:v="${VENDOR_NAMESPACE}" v:val="match"/></m:dPr><m:e>${run}</m:e></m:d>`,
      `<m:d><m:dPr><m:begChr m:val="[" m:extra="semantic"/></m:dPr><m:e>${run}</m:e></m:d>`,
      `<m:d xmlns:r="${RELATIONSHIP_NAMESPACE}"><m:dPr><m:sepChr r:id="rIdUnsafe"/></m:dPr><m:e>${run}</m:e></m:d>`,
      `<m:d><v:dPr xmlns:v="${VENDOR_NAMESPACE}"/><m:e>${run}</m:e></m:d>`,
      `<m:d><m:dPr><v:begChr xmlns:v="${VENDOR_NAMESPACE}"/></m:dPr><m:e>${run}</m:e></m:d>`,
      `<m:d><m:dPr>meaningful<m:begChr/></m:dPr><m:e>${run}</m:e></m:d>`,
      `<m:d m:extra="semantic"><m:e>${run}</m:e></m:d>`,
      '<m:d><m:e>meaningful</m:e></m:d>',
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

function richWordRunProperties() {
  return {
    fonts: {
      ascii: 'Cambria Math',
      highAnsi: 'Cambria Math',
      eastAsia: '等线',
      complexScript: 'Arial',
      asciiTheme: 'majorAscii' as const,
      highAnsiTheme: 'majorHAnsi' as const,
      eastAsiaTheme: 'minorEastAsia' as const,
      complexScriptTheme: 'majorBidi' as const,
      hint: 'eastAsia' as const,
    },
    bold: true,
    boldComplexScript: false,
    italic: false,
    italicComplexScript: true,
    allCaps: false,
    smallCaps: true,
    strike: false,
    doubleStrike: true,
    outline: true,
    shadow: true,
    emboss: false,
    imprint: false,
    noProof: true,
    snapToGrid: false,
    hidden: false,
    webHidden: true,
    color: {
      value: '#1a2b3c',
      theme: 'accent2' as const,
      tint: '80',
      shade: '40',
    },
    characterSpacingTwips: 20,
    characterScalePercent: 90,
    kerningThresholdHalfPoints: 22,
    positionHalfPoints: 2,
    fontSize: 12.5,
    fontSizeComplexScript: 14,
    underline: {
      style: 'wavyDouble' as const,
      color: {
        value: '#abcdef',
        theme: 'accent3' as const,
        tint: '20',
      },
    },
    textEffect: 'shimmer' as const,
    border: {
      style: 'double' as const,
      color: {
        value: '#112233',
        theme: 'accent1' as const,
        tint: '66',
        shade: '33',
      },
      sizeEighthPoints: 24,
      spacingPoints: 2,
      shadow: true,
      frame: false,
    },
    fitText: { widthTwips: 720, id: 50 },
    verticalAlignment: 'baseline' as const,
    rightToLeft: false,
    complexScript: false,
    emphasisMark: 'circle' as const,
    languages: { latin: 'en-US', eastAsia: 'zh-CN', bidi: 'ar-SA' },
    eastAsianLayout: {
      id: 9,
      combine: true,
      combineBrackets: 'curly' as const,
      vertical: false,
      verticalCompress: true,
    },
    paragraphMarkAlwaysHidden: false,
    glow: {
      radiusEmus: 63_500,
      color: {
        type: 'scheme' as const,
        value: 'accent4' as const,
        transforms: [{ type: 'alpha' as const, value: 50_000 }],
      },
    },
    shadowEffect: {
      blurRadiusEmus: 25_400,
      distanceEmus: 12_700,
      directionDegrees: 90,
      horizontalScalePercent: 90,
      verticalScalePercent: -100,
      horizontalSkewDegrees: -1,
      verticalSkewDegrees: 2,
      alignment: 'topRight' as const,
      color: {
        type: 'scheme' as const,
        value: 'accent5' as const,
        transforms: [{ type: 'luminanceModulation' as const, value: 75_000 }],
      },
    },
    reflectionEffect: {
      blurRadiusEmus: 12_700,
      startOpacityPercent: 60,
      startPositionPercent: 0,
      endOpacityPercent: 10,
      endPositionPercent: 100,
      distanceEmus: 6_350,
      directionDegrees: 90,
      fadeDirectionDegrees: 270,
      horizontalScalePercent: 100,
      verticalScalePercent: -50,
      horizontalSkewDegrees: -1,
      verticalSkewDegrees: 2,
      alignment: 'bottom' as const,
    },
    textOutlineEffect: {
      widthEmus: 12_700,
      cap: 'square' as const,
      compound: 'double' as const,
      alignment: 'inset' as const,
      fill: {
        type: 'solid' as const,
        color: {
          type: 'scheme' as const,
          value: 'accent6' as const,
          transforms: [{ type: 'alpha' as const, value: 80_000 }],
        },
      },
      dash: { preset: 'dashDot' as const },
      join: { type: 'miter' as const, limitPercent: 400 },
    },
    textFillEffect: {
      fill: {
        type: 'solid' as const,
        color: {
          type: 'rgb' as const,
          value: '#2468ac',
          transforms: [{ type: 'shade' as const, value: 25_000 }],
        },
      },
    },
  };
}

function controlPropertiesEquation(): WorkDocumentEquation {
  const run = (text: string) => ({ type: 'run' as const, text });
  const controlProperties = richWordRunProperties();
  return {
    version: 1,
    display: 'inline',
    children: [
      {
        type: 'fraction',
        controlProperties,
        fractionType: 'bar',
        numerator: [run('fraction-numerator')],
        denominator: [run('fraction-denominator')],
      },
      {
        type: 'superscript',
        controlProperties,
        base: [run('superscript-base')],
        superScript: [run('superscript-value')],
      },
      {
        type: 'subscript',
        controlProperties,
        base: [run('subscript-base')],
        subScript: [run('subscript-value')],
      },
      {
        type: 'subSuperScript',
        controlProperties,
        alignScripts: true,
        base: [run('right-script-base')],
        subScript: [run('right-subscript')],
        superScript: [run('right-superscript')],
      },
      {
        type: 'preSubSuperScript',
        controlProperties,
        base: [run('left-script-base')],
        subScript: [run('left-subscript')],
        superScript: [run('left-superscript')],
      },
      {
        type: 'lowerLimit',
        controlProperties,
        base: [run('lower-limit-base')],
        limit: [run('lower-limit')],
      },
      {
        type: 'upperLimit',
        controlProperties,
        base: [run('upper-limit-base')],
        limit: [run('upper-limit')],
      },
      {
        type: 'radical',
        controlProperties,
        children: [run('radical')],
      },
      {
        type: 'function',
        controlProperties,
        name: [run('function')],
        children: [run('function-argument')],
      },
      {
        type: 'nary',
        controlProperties,
        operator: '\u2211',
        limitLocation: 'underOver',
        children: [run('nary-body')],
      },
      {
        type: 'accent',
        controlProperties,
        character: '\u0303',
        children: [run('accent-body')],
      },
      {
        type: 'bar',
        controlProperties,
        position: 'top',
        children: [run('bar-body')],
      },
      {
        type: 'groupCharacter',
        controlProperties,
        character: '\u23de',
        position: 'top',
        verticalJustification: 'bottom',
        children: [run('group-character-body')],
      },
      {
        type: 'phantom',
        controlProperties,
        show: false,
        zeroWidth: true,
        zeroAscent: false,
        zeroDescent: true,
        transparent: true,
        children: [run('phantom-body')],
      },
      {
        type: 'borderBox',
        controlProperties,
        hideTop: false,
        hideBottom: true,
        hideLeft: false,
        hideRight: true,
        strikeHorizontal: true,
        strikeVertical: false,
        strikeBottomLeftToTopRight: true,
        strikeTopLeftToBottomRight: false,
        children: [run('border-box-body')],
      },
      {
        type: 'box',
        controlProperties,
        operatorEmulator: true,
        noBreak: true,
        differential: true,
        alignment: true,
        children: [run('box-body')],
      },
      {
        type: 'matrix',
        controlProperties,
        baseAlignment: 'center',
        placeholdersHidden: false,
        columnAlignments: ['center'],
        rows: [[[run('matrix-cell')]]],
      },
      {
        type: 'equationArray',
        controlProperties,
        baseAlignment: 'center',
        maximumDistribution: false,
        objectDistribution: false,
        rowSpacingRule: 'single',
        rowSpacing: 0,
        rows: [[run('equation-array-row')]],
      },
      {
        type: 'delimiter',
        controlProperties,
        opening: '[',
        closing: ']',
        separator: ';',
        arguments: [[run('delimiter-left')], [run('delimiter-right')]],
      },
    ],
  };
}

function argumentControlPropertiesEquation(): WorkDocumentEquation {
  const run = (text: string) => ({ type: 'run' as const, text });
  const argumentProperties = {
    controlProperties: richWordRunProperties(),
  };
  return {
    version: 1,
    display: 'inline',
    children: [
      {
        type: 'fraction',
        fractionType: 'bar',
        numerator: [run('fraction-numerator')],
        numeratorProperties: argumentProperties,
        denominator: [run('fraction-denominator')],
        denominatorProperties: argumentProperties,
      },
      {
        type: 'superscript',
        base: [run('superscript-base')],
        baseProperties: argumentProperties,
        superScript: [run('superscript-value')],
        superScriptProperties: argumentProperties,
      },
      {
        type: 'subscript',
        base: [run('subscript-base')],
        baseProperties: argumentProperties,
        subScript: [run('subscript-value')],
        subScriptProperties: argumentProperties,
      },
      {
        type: 'subSuperScript',
        base: [run('right-script-base')],
        baseProperties: argumentProperties,
        subScript: [run('right-subscript')],
        subScriptProperties: argumentProperties,
        superScript: [run('right-superscript')],
        superScriptProperties: argumentProperties,
      },
      {
        type: 'preSubSuperScript',
        base: [run('left-script-base')],
        baseProperties: argumentProperties,
        subScript: [run('left-subscript')],
        subScriptProperties: argumentProperties,
        superScript: [run('left-superscript')],
        superScriptProperties: argumentProperties,
      },
      {
        type: 'lowerLimit',
        base: [run('lower-limit-base')],
        baseProperties: argumentProperties,
        limit: [run('lower-limit')],
        limitProperties: argumentProperties,
      },
      {
        type: 'upperLimit',
        base: [run('upper-limit-base')],
        baseProperties: argumentProperties,
        limit: [run('upper-limit')],
        limitProperties: argumentProperties,
      },
      {
        type: 'radical',
        children: [run('radical')],
        childrenProperties: argumentProperties,
        degreeProperties: argumentProperties,
      },
      {
        type: 'function',
        name: [],
        nameProperties: argumentProperties,
        children: [run('function-argument')],
        childrenProperties: argumentProperties,
      },
      {
        type: 'nary',
        operator: '\u2211',
        limitLocation: 'underOver',
        children: [run('nary-body')],
        childrenProperties: argumentProperties,
        subScriptProperties: argumentProperties,
        superScriptProperties: argumentProperties,
      },
      {
        type: 'accent',
        character: '\u0303',
        children: [run('accent-body')],
        childrenProperties: argumentProperties,
      },
      {
        type: 'bar',
        position: 'top',
        children: [run('bar-body')],
        childrenProperties: argumentProperties,
      },
      {
        type: 'groupCharacter',
        character: '\u23de',
        position: 'top',
        verticalJustification: 'bottom',
        children: [run('group-character-body')],
        childrenProperties: argumentProperties,
      },
      {
        type: 'phantom',
        show: false,
        zeroWidth: true,
        zeroAscent: false,
        zeroDescent: true,
        transparent: true,
        children: [run('phantom-body')],
        childrenProperties: argumentProperties,
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
        children: [run('border-box-body')],
        childrenProperties: argumentProperties,
      },
      {
        type: 'box',
        operatorEmulator: true,
        noBreak: true,
        differential: true,
        alignment: true,
        children: [run('box-body')],
        childrenProperties: argumentProperties,
      },
      {
        type: 'matrix',
        baseAlignment: 'center',
        placeholdersHidden: false,
        columnAlignments: ['left', 'right'],
        rows: [
          [[run('matrix-00')], [run('matrix-01')]],
          [[run('matrix-10')], [run('matrix-11')]],
        ],
        cellProperties: [
          [argumentProperties, null],
          [null, argumentProperties],
        ],
      },
      {
        type: 'equationArray',
        baseAlignment: 'center',
        maximumDistribution: false,
        objectDistribution: false,
        rowSpacingRule: 'single',
        rowSpacing: 0,
        rows: [[run('equation-array-0')], [run('equation-array-1')]],
        rowProperties: [null, argumentProperties],
      },
      {
        type: 'delimiter',
        opening: '[',
        closing: ']',
        separator: ';',
        arguments: [[run('delimiter-left')], [run('delimiter-right')]],
        argumentProperties: [null, argumentProperties],
      },
    ],
  };
}

function controlRevisionEquation(): WorkDocumentEquation {
  const run = (text: string) => ({ type: 'run' as const, text });
  return {
    version: 1,
    display: 'inline',
    children: [
      {
        type: 'fraction',
        controlRevision: {
          kind: 'insertion',
          id: 0,
          author: 'Alice & Bob',
          date: '2026-01-02T03:04:05Z',
          dateUtc: '2026-01-02T03:04:05.123Z',
          child: {
            kind: 'deletion',
            id: 2_147_483_647,
            author: 'Final reviewer',
          },
        },
        controlProperties: {
          bold: true,
          color: { value: '#1a2b3c' },
        },
        fractionType: 'linear',
        numerator: [run('revision-numerator')],
        denominator: [run('revision-denominator')],
      },
      {
        type: 'box',
        controlRevision: {
          kind: 'deletion',
          id: 1,
          author: 'Deletion reviewer',
        },
        operatorEmulator: false,
        noBreak: false,
        differential: false,
        alignment: false,
        children: [run('revision-argument')],
        childrenProperties: {
          controlRevision: {
            kind: 'moveFrom',
            id: 2,
            author: 'Move source reviewer',
            date: '2026-02-03T04:05:06.789+08:00',
            dateUtc: '2026-02-02T20:05:06.789Z',
            child: {
              kind: 'insertion',
              id: 3,
              author: 'Nested insertion reviewer',
              child: {
                kind: 'deletion',
                id: 4,
                author: 'Nested deletion reviewer',
              },
            },
          },
          controlProperties: { italic: true, fontSize: 12.5 },
        },
      },
      {
        type: 'delimiter',
        controlRevision: {
          kind: 'moveTo',
          id: 5,
          author: 'Move destination reviewer',
          child: {
            kind: 'deletion',
            id: 6,
            author: 'Destination deletion reviewer',
          },
        },
        controlProperties: {
          underline: { style: 'single' },
        },
        opening: '[',
        closing: ']',
        separator: ';',
        arguments: [[run('revision-delimiter')]],
      },
    ],
  } as unknown as WorkDocumentEquation;
}

function everyObjectControlRevisionEquation(): WorkDocumentEquation {
  const source = controlPropertiesEquation();
  return {
    ...source,
    children: source.children.map((expression, index) => {
      if (expression.type === 'run') {
        throw new Error('Expected an object equation expression.');
      }
      const { controlProperties: _controlProperties, ...rest } = expression;
      return {
        ...rest,
        controlRevision: {
          kind: 'deletion',
          id: index,
          author: `Object reviewer ${index}`,
        },
      } as typeof expression;
    }),
  };
}

function argumentSizesEquation(): WorkDocumentEquation {
  const run = (text: string) => ({ type: 'run' as const, text });
  const sized = (size: -2 | -1 | 1 | 2) => ({ size });
  return {
    version: 1,
    display: 'inline',
    children: [
      {
        type: 'box',
        operatorEmulator: false,
        noBreak: false,
        differential: false,
        alignment: false,
        children: [run('box-sized')],
        childrenProperties: sized(-2),
      },
      {
        type: 'groupCharacter',
        character: '\u23de',
        position: 'top',
        verticalJustification: 'bottom',
        children: [run('group-sized')],
        childrenProperties: sized(-1),
      },
      {
        type: 'lowerLimit',
        base: [run('lower-base')],
        limit: [run('lower-limit-sized')],
        limitProperties: sized(1),
      },
      {
        type: 'upperLimit',
        base: [run('upper-base')],
        limit: [run('upper-limit-sized')],
        limitProperties: sized(2),
      },
      {
        type: 'nary',
        operator: '\u2211',
        limitLocation: 'underOver',
        children: [run('nary-body')],
        subScript: [run('nary-sub-sized')],
        subScriptProperties: sized(-2),
        superScript: [run('nary-sup-sized')],
        superScriptProperties: sized(-1),
      },
      {
        type: 'radical',
        children: [run('radical-body')],
        degree: [run('radical-degree-sized')],
        degreeProperties: sized(1),
      },
      {
        type: 'preSubSuperScript',
        base: [run('pre-base')],
        subScript: [run('pre-sub-sized')],
        subScriptProperties: sized(2),
        superScript: [run('pre-sup-sized')],
        superScriptProperties: sized(-2),
      },
      {
        type: 'subscript',
        base: [run('subscript-base')],
        subScript: [run('subscript-sized')],
        subScriptProperties: sized(-1),
      },
      {
        type: 'subSuperScript',
        base: [run('sub-super-base')],
        subScript: [run('sub-super-sub-sized')],
        subScriptProperties: sized(1),
        superScript: [run('sub-super-sup-sized')],
        superScriptProperties: sized(2),
      },
      {
        type: 'superscript',
        base: [run('superscript-base')],
        superScript: [run('superscript-sized')],
        superScriptProperties: sized(-2),
      },
      {
        type: 'fraction',
        fractionType: 'bar',
        numerator: [run('fraction-noop-sized')],
        numeratorProperties: {
          size: 1,
          controlProperties: { bold: true },
        },
        denominator: [run('fraction-denominator')],
      },
      {
        type: 'function',
        name: [run('function-noop-sized')],
        nameProperties: sized(-1),
        children: [run('function-body')],
      },
      {
        type: 'matrix',
        baseAlignment: 'center',
        placeholdersHidden: false,
        columnAlignments: ['center'],
        rows: [[[run('matrix-noop-sized')]]],
        cellProperties: [[sized(2)]],
      },
      {
        type: 'equationArray',
        baseAlignment: 'center',
        maximumDistribution: false,
        objectDistribution: false,
        rowSpacingRule: 'single',
        rowSpacing: 0,
        rows: [[run('array-noop-sized')]],
        rowProperties: [sized(1)],
      },
      {
        type: 'delimiter',
        opening: '[',
        closing: ']',
        separator: ';',
        arguments: [[run('delimiter-noop-sized')]],
        argumentProperties: [sized(-2)],
      },
    ],
  };
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
        wordRunProperties: richWordRunProperties(),
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
      {
        type: 'fraction',
        fractionType: 'bar',
        numerator: [],
        denominator: [],
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
      { type: 'function', name: [], children: [] },
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
        type: 'delimiter',
        opening: '',
        closing: '',
        separator: '',
        arguments: [[], [run('z')], []],
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
        spacing: {
          rowSpacingRule: 'exact',
          rowSpacing: 12,
          columnGapRule: 'multiple',
          columnGap: 3,
          minimumColumnWidthTwips: 120,
        },
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

async function expectNativeWordRunGeometry(blob: Blob): Promise<void> {
  const archive = await JSZip.loadAsync(await blob.arrayBuffer());
  const document = await xmlEntry(archive, 'word/document.xml');
  const mathRuns = descendants(document, 'r').filter(
    (run) => run.namespaceURI === MATH_NAMESPACE,
  );
  const propertiesFor = (text: string) => {
    const run = mathRuns.find((candidate) => candidate.textContent === text);
    expect(run, text).toBeDefined();
    const properties = directChildren(run as Element, 'rPr').find(
      (candidate) => candidate.namespaceURI === WORD_NAMESPACE,
    );
    expect(properties, text).toBeDefined();
    return directChildren(properties as Element);
  };
  for (const entry of [
    {
      text: 'expanded-raised',
      values: ['200', '75', '24', '6', '28'],
    },
    {
      text: 'explicit-resets',
      values: ['0', '100', '0', '0', '24'],
    },
  ]) {
    const properties = propertiesFor(entry.text);
    expect(
      properties.map((property) => property.localName),
      entry.text,
    ).toEqual(['spacing', 'w', 'kern', 'position', 'sz']);
    expect(
      properties.map((property) => wordAttributes(property).val),
      entry.text,
    ).toEqual(entry.values);
  }
  const belowThreshold = propertiesFor('kerning-below-threshold');
  expect(belowThreshold.map((property) => property.localName)).toEqual([
    'kern',
    'sz',
  ]);
  expect(
    belowThreshold.map((property) => wordAttributes(property).val),
  ).toEqual(['25', '24']);
  const inheritedSize = propertiesFor('kerning-inherited-size');
  expect(inheritedSize.map((property) => property.localName)).toEqual(['kern']);
  expect(wordAttributes(inheritedSize[0]).val).toBe('24');

  const nary = descendants(document, 'nary').find((candidate) =>
    candidate.textContent?.includes('operator-geometry'),
  );
  expect(nary).toBeDefined();
  const naryProperties = directChildren(nary as Element, 'naryPr')[0];
  const controlProperties = directChildren(naryProperties, 'ctrlPr')[0];
  const wordProperties = directChildren(controlProperties, 'rPr').find(
    (candidate) => candidate.namespaceURI === WORD_NAMESPACE,
  );
  expect(wordProperties).toBeDefined();
  const properties = directChildren(wordProperties as Element);
  expect(properties.map((property) => property.localName)).toEqual([
    'spacing',
    'w',
    'kern',
    'position',
    'sz',
  ]);
  expect(properties.map((property) => wordAttributes(property).val)).toEqual([
    '-40',
    '125',
    '30',
    '-4',
    '30',
  ]);
}

async function expectNativeWordRunEffects(blob: Blob): Promise<void> {
  const archive = await JSZip.loadAsync(await blob.arrayBuffer());
  const document = await xmlEntry(archive, 'word/document.xml');
  const mathRuns = descendants(document, 'r').filter(
    (run) => run.namespaceURI === MATH_NAMESPACE,
  );
  const propertiesFor = (text: string) => {
    const run = mathRuns.find((candidate) => candidate.textContent === text);
    expect(run, text).toBeDefined();
    const properties = directChildren(run as Element, 'rPr').find(
      (candidate) => candidate.namespaceURI === WORD_NAMESPACE,
    );
    expect(properties, text).toBeDefined();
    return directChildren(properties as Element);
  };
  for (const entry of [
    {
      text: 'all-caps-outline',
      names: ['caps', 'outline', 'shadow', 'vanish', 'webHidden'],
      values: ['1', '1', '1', '0', '0'],
    },
    {
      text: 'small-caps-hidden',
      names: [
        'caps',
        'smallCaps',
        'outline',
        'shadow',
        'emboss',
        'imprint',
        'vanish',
        'webHidden',
      ],
      values: ['0', '1', '0', '0', '1', '0', '1', '1'],
    },
    {
      text: 'explicit-effect-resets',
      names: [
        'caps',
        'smallCaps',
        'outline',
        'shadow',
        'emboss',
        'imprint',
        'vanish',
        'webHidden',
      ],
      values: ['0', '0', '0', '0', '0', '0', '0', '0'],
    },
  ]) {
    const properties = propertiesFor(entry.text);
    expect(
      properties.map((property) => property.localName),
      entry.text,
    ).toEqual(entry.names);
    expect(
      properties.map((property) => wordAttributes(property).val),
      entry.text,
    ).toEqual(entry.values);
  }

  const nary = descendants(document, 'nary').find((candidate) =>
    candidate.textContent?.includes('operator-effects'),
  );
  expect(nary).toBeDefined();
  const naryProperties = directChildren(nary as Element, 'naryPr')[0];
  const controlProperties = directChildren(naryProperties, 'ctrlPr')[0];
  const wordProperties = directChildren(controlProperties, 'rPr').find(
    (candidate) => candidate.namespaceURI === WORD_NAMESPACE,
  );
  expect(wordProperties).toBeDefined();
  const properties = directChildren(wordProperties as Element);
  expect(properties.map((property) => property.localName)).toEqual([
    'smallCaps',
    'shadow',
  ]);
  expect(properties.map((property) => wordAttributes(property).val)).toEqual([
    '1',
    '1',
  ]);
}

async function expectNativeWordRunAdornments(blob: Blob): Promise<void> {
  const archive = await JSZip.loadAsync(await blob.arrayBuffer());
  const document = await xmlEntry(archive, 'word/document.xml');
  const mathRuns = descendants(document, 'r').filter(
    (run) => run.namespaceURI === MATH_NAMESPACE,
  );
  const propertiesFor = (text: string) => {
    const run = mathRuns.find((candidate) => candidate.textContent === text);
    expect(run, text).toBeDefined();
    const properties = directChildren(run as Element, 'rPr').find(
      (candidate) => candidate.namespaceURI === WORD_NAMESPACE,
    );
    expect(properties, text).toBeDefined();
    return directChildren(properties as Element);
  };
  for (const entry of [
    {
      text: 'animated-bordered',
      effect: 'shimmer',
      border: {
        val: 'double',
        color: '1A2B3C',
        themeColor: 'accent1',
        themeTint: '66',
        themeShade: '33',
        sz: '24',
        space: '2',
        shadow: '1',
        frame: '0',
      },
    },
    {
      text: 'explicit-border-removal',
      effect: 'none',
      border: { val: 'none', shadow: '0', frame: '0' },
    },
    {
      text: 'native-wave-border',
      effect: 'antsBlack',
      border: {
        val: 'wave',
        themeColor: 'accent2',
        sz: '16',
        space: '3',
      },
    },
    {
      text: 'theme-only-border',
      effect: 'lights',
      border: {
        val: 'single',
        themeColor: 'accent4',
        sz: '8',
      },
    },
  ]) {
    const properties = propertiesFor(entry.text);
    expect(
      properties.map((property) => property.localName),
      entry.text,
    ).toEqual(['effect', 'bdr']);
    expect(wordAttributes(properties[0])).toEqual({ val: entry.effect });
    expect(wordAttributes(properties[1])).toEqual(entry.border);
  }

  const nary = descendants(document, 'nary').find((candidate) =>
    candidate.textContent?.includes('operator-adornments'),
  );
  expect(nary).toBeDefined();
  const naryProperties = directChildren(nary as Element, 'naryPr')[0];
  const controlProperties = directChildren(naryProperties, 'ctrlPr')[0];
  const wordProperties = directChildren(controlProperties, 'rPr').find(
    (candidate) => candidate.namespaceURI === WORD_NAMESPACE,
  );
  expect(wordProperties).toBeDefined();
  const properties = directChildren(wordProperties as Element);
  expect(properties.map((property) => property.localName)).toEqual([
    'effect',
    'bdr',
  ]);
  expect(wordAttributes(properties[0])).toEqual({ val: 'sparkle' });
  expect(wordAttributes(properties[1])).toEqual({
    val: 'dotted',
    color: 'auto',
    sz: '8',
    space: '0',
  });
}

async function expectNativeWordRunBackgrounds(blob: Blob): Promise<void> {
  const archive = await JSZip.loadAsync(await blob.arrayBuffer());
  const document = await xmlEntry(archive, 'word/document.xml');
  const mathRuns = descendants(document, 'r').filter(
    (run) => run.namespaceURI === MATH_NAMESPACE,
  );
  const propertiesFor = (text: string) => {
    const run = mathRuns.find((candidate) => candidate.textContent === text);
    expect(run, text).toBeDefined();
    const properties = directChildren(run as Element, 'rPr').find(
      (candidate) => candidate.namespaceURI === WORD_NAMESPACE,
    );
    expect(properties, text).toBeDefined();
    return directChildren(properties as Element);
  };
  const expected = [
    {
      text: 'highlighted',
      names: ['highlight', 'shd'],
      attributes: [{ val: 'yellow' }, { val: 'clear', fill: '112233' }],
    },
    {
      text: 'clear-shading',
      names: ['shd'],
      attributes: [
        {
          val: 'clear',
          fill: '112233',
          themeFill: 'accent4',
          themeFillTint: '80',
        },
      ],
    },
    {
      text: 'solid-shading',
      names: ['shd'],
      attributes: [
        {
          val: 'solid',
          color: 'ABCDEF',
          themeColor: 'text2',
          themeShade: '40',
          fill: '445566',
        },
      ],
    },
    {
      text: 'patterned-shading',
      names: ['shd'],
      attributes: [
        {
          val: 'pct20',
          color: 'FF0000',
          themeFill: 'accent3',
          themeFillTint: '20',
        },
      ],
    },
    {
      text: 'theme-only-shading',
      names: ['shd'],
      attributes: [
        {
          val: 'clear',
          themeFill: 'accent2',
          themeFillTint: '99',
        },
      ],
    },
    {
      text: 'unhighlighted',
      names: ['highlight', 'shd'],
      attributes: [{ val: 'none' }, { val: 'clear', fill: '112233' }],
    },
    {
      text: 'nil-shading',
      names: ['shd'],
      attributes: [{ val: 'nil', fill: '112233' }],
    },
  ];
  for (const entry of expected) {
    const properties = propertiesFor(entry.text);
    expect(
      properties.map((property) => property.localName),
      entry.text,
    ).toEqual(entry.names);
    expect(properties.map(wordAttributes), entry.text).toEqual(
      entry.attributes,
    );
  }

  const nary = descendants(document, 'nary').find((candidate) =>
    candidate.textContent?.includes('operator-background'),
  );
  expect(nary).toBeDefined();
  const naryProperties = directChildren(nary as Element, 'naryPr')[0];
  const controlProperties = directChildren(naryProperties, 'ctrlPr')[0];
  const wordProperties = directChildren(controlProperties, 'rPr').find(
    (candidate) => candidate.namespaceURI === WORD_NAMESPACE,
  );
  expect(wordProperties).toBeDefined();
  const highlight = directChildren(wordProperties as Element, 'highlight')[0];
  expect(wordAttributes(highlight)).toEqual({ val: 'darkCyan' });
}

async function expectNativeWordRunFitText(blob: Blob): Promise<void> {
  const archive = await JSZip.loadAsync(await blob.arrayBuffer());
  const document = await xmlEntry(archive, 'word/document.xml');
  const mathRuns = descendants(document, 'r').filter(
    (run) => run.namespaceURI === MATH_NAMESPACE,
  );
  const fitTextFor = (text: string) => {
    const run = mathRuns.find((candidate) => candidate.textContent === text);
    expect(run, text).toBeDefined();
    const properties = directChildren(run as Element, 'rPr').find(
      (candidate) => candidate.namespaceURI === WORD_NAMESPACE,
    );
    expect(properties, text).toBeDefined();
    const children = directChildren(properties as Element);
    expect(
      children.map((child) => child.localName),
      text,
    ).toEqual(['fitText']);
    return children[0] as Element;
  };
  for (const entry of [
    { text: 'fit-width', attributes: { val: '720', id: '50' } },
    { text: 'fit-zero', attributes: { val: '0', id: '0' } },
    { text: 'fit-without-id', attributes: { val: '31680' } },
  ]) {
    const fitText = fitTextFor(entry.text);
    expect(wordAttributes(fitText), entry.text).toEqual(entry.attributes);
    expect(
      Array.from(fitText.attributes).map((attribute) => attribute.localName),
      entry.text,
    ).toEqual(Object.keys(entry.attributes).map((name) => `w:${name}`));
  }

  const nary = descendants(document, 'nary').find((candidate) =>
    candidate.textContent?.includes('operator-fit-width'),
  );
  expect(nary).toBeDefined();
  const naryProperties = directChildren(nary as Element, 'naryPr')[0];
  const controlProperties = directChildren(naryProperties, 'ctrlPr')[0];
  const wordProperties = directChildren(controlProperties, 'rPr').find(
    (candidate) => candidate.namespaceURI === WORD_NAMESPACE,
  );
  expect(wordProperties).toBeDefined();
  const fitText = directChildren(wordProperties as Element, 'fitText')[0];
  expect(wordAttributes(fitText)).toEqual({ val: '1440', id: '-7' });
}

async function expectNativeWordRunVerticalAlignments(
  blob: Blob,
): Promise<void> {
  const archive = await JSZip.loadAsync(await blob.arrayBuffer());
  const document = await xmlEntry(archive, 'word/document.xml');
  const mathRuns = descendants(document, 'r').filter(
    (run) => run.namespaceURI === MATH_NAMESPACE,
  );
  for (const entry of [
    {
      text: 'baseline-aligned',
      names: ['vertAlign'],
      attributes: [{ val: 'baseline' }],
    },
    {
      text: 'superscript-aligned',
      names: ['sz', 'vertAlign'],
      attributes: [{ val: '24' }, { val: 'superscript' }],
    },
    {
      text: 'subscript-aligned',
      names: ['vertAlign'],
      attributes: [{ val: 'subscript' }],
    },
    {
      text: 'positioned-superscript',
      names: ['position', 'vertAlign'],
      attributes: [{ val: '4' }, { val: 'superscript' }],
    },
  ]) {
    const run = mathRuns.find(
      (candidate) => candidate.textContent === entry.text,
    );
    expect(run, entry.text).toBeDefined();
    const properties = directChildren(run as Element, 'rPr').find(
      (candidate) => candidate.namespaceURI === WORD_NAMESPACE,
    );
    expect(properties, entry.text).toBeDefined();
    const children = directChildren(properties as Element);
    expect(
      children.map((child) => child.localName),
      entry.text,
    ).toEqual(entry.names);
    expect(children.map(wordAttributes), entry.text).toEqual(entry.attributes);
    expect(
      children.every(
        (child) =>
          child.attributes.length === 1 &&
          child.attributes[0]?.name === 'w:val',
      ),
      entry.text,
    ).toBe(true);
  }

  const nary = descendants(document, 'nary').find((candidate) =>
    candidate.textContent?.includes('operator-vertical-alignment'),
  );
  expect(nary).toBeDefined();
  const naryProperties = directChildren(nary as Element, 'naryPr')[0];
  const controlProperties = directChildren(naryProperties, 'ctrlPr')[0];
  const wordProperties = directChildren(controlProperties, 'rPr').find(
    (candidate) => candidate.namespaceURI === WORD_NAMESPACE,
  );
  expect(wordProperties).toBeDefined();
  const verticalAlignment = directChildren(
    wordProperties as Element,
    'vertAlign',
  )[0];
  expect(wordAttributes(verticalAlignment)).toEqual({ val: 'subscript' });
}

async function expectNativeWordRunEmphasisMarks(blob: Blob): Promise<void> {
  const archive = await JSZip.loadAsync(await blob.arrayBuffer());
  const document = await xmlEntry(archive, 'word/document.xml');
  const mathRuns = descendants(document, 'r').filter(
    (run) => run.namespaceURI === MATH_NAMESPACE,
  );
  for (const entry of [
    {
      text: 'no-emphasis',
      names: ['em'],
      attributes: [{ val: 'none' }],
    },
    {
      text: 'dot-emphasis',
      names: ['em'],
      attributes: [{ val: 'dot' }],
    },
    {
      text: 'comma-emphasis',
      names: ['em'],
      attributes: [{ val: 'comma' }],
    },
    {
      text: 'circle-emphasis',
      names: ['em'],
      attributes: [{ val: 'circle' }],
    },
    {
      text: 'under-dot-emphasis',
      names: ['em'],
      attributes: [{ val: 'underDot' }],
    },
    {
      text: 'ordered-emphasis',
      names: ['vertAlign', 'rtl', 'cs', 'em', 'lang'],
      attributes: [
        { val: 'baseline' },
        { val: '0' },
        { val: '1' },
        { val: 'dot' },
        { val: 'en-US' },
      ],
    },
  ]) {
    const run = mathRuns.find(
      (candidate) => candidate.textContent === entry.text,
    );
    expect(run, entry.text).toBeDefined();
    const properties = directChildren(run as Element, 'rPr').find(
      (candidate) => candidate.namespaceURI === WORD_NAMESPACE,
    );
    expect(properties, entry.text).toBeDefined();
    const children = directChildren(properties as Element);
    expect(
      children.map((child) => child.localName),
      entry.text,
    ).toEqual(entry.names);
    expect(children.map(wordAttributes), entry.text).toEqual(entry.attributes);
    expect(
      children.every(
        (child) =>
          child.attributes.length === 1 &&
          child.attributes[0]?.name === 'w:val',
      ),
      entry.text,
    ).toBe(true);
  }

  const nary = descendants(document, 'nary').find((candidate) =>
    candidate.textContent?.includes('operator-emphasis-mark'),
  );
  expect(nary).toBeDefined();
  const naryProperties = directChildren(nary as Element, 'naryPr')[0];
  const controlProperties = directChildren(naryProperties, 'ctrlPr')[0];
  const wordProperties = directChildren(controlProperties, 'rPr').find(
    (candidate) => candidate.namespaceURI === WORD_NAMESPACE,
  );
  expect(wordProperties).toBeDefined();
  const emphasis = directChildren(wordProperties as Element, 'em')[0];
  expect(wordAttributes(emphasis)).toEqual({ val: 'underDot' });
}

async function expectNativeWordRunEastAsianLayouts(blob: Blob): Promise<void> {
  const archive = await JSZip.loadAsync(await blob.arrayBuffer());
  const document = await xmlEntry(archive, 'word/document.xml');
  const mathRuns = descendants(document, 'r').filter(
    (run) => run.namespaceURI === MATH_NAMESPACE,
  );
  for (const entry of [
    {
      text: 'layout-id-zero',
      names: ['eastAsianLayout'],
      attributes: [{ id: '0' }],
      qualifiedAttributes: ['w:id'],
    },
    {
      text: 'two-lines-round',
      names: ['eastAsianLayout'],
      attributes: [{ combine: '1', combineBrackets: 'round' }],
      qualifiedAttributes: ['w:combine', 'w:combineBrackets'],
    },
    {
      text: 'explicit-combine-reset',
      names: ['eastAsianLayout'],
      attributes: [{ combine: '0', combineBrackets: 'none' }],
      qualifiedAttributes: ['w:combine', 'w:combineBrackets'],
    },
    {
      text: 'horizontal-in-vertical',
      names: ['eastAsianLayout'],
      attributes: [{ vert: '1' }],
      qualifiedAttributes: ['w:vert'],
    },
    {
      text: 'explicit-vertical-reset',
      names: ['eastAsianLayout'],
      attributes: [{ vert: '0', vertCompress: '1' }],
      qualifiedAttributes: ['w:vert', 'w:vertCompress'],
    },
    {
      text: 'full-east-asian-layout',
      names: ['lang', 'eastAsianLayout'],
      attributes: [
        { eastAsia: 'ja-JP' },
        {
          id: '-2147483648',
          combine: '1',
          combineBrackets: 'curly',
          vert: '1',
          vertCompress: '0',
        },
      ],
      qualifiedAttributes: [
        'w:id',
        'w:combine',
        'w:combineBrackets',
        'w:vert',
        'w:vertCompress',
      ],
    },
    {
      text: 'layout-id-maximum',
      names: ['eastAsianLayout'],
      attributes: [{ id: '2147483647' }],
      qualifiedAttributes: ['w:id'],
    },
  ]) {
    const run = mathRuns.find(
      (candidate) => candidate.textContent === entry.text,
    );
    expect(run, entry.text).toBeDefined();
    const properties = directChildren(run as Element, 'rPr').find(
      (candidate) => candidate.namespaceURI === WORD_NAMESPACE,
    );
    expect(properties, entry.text).toBeDefined();
    const children = directChildren(properties as Element);
    expect(
      children.map((child) => child.localName),
      entry.text,
    ).toEqual(entry.names);
    expect(children.map(wordAttributes), entry.text).toEqual(entry.attributes);
    const layout = children.find(
      (child) => child.localName === 'eastAsianLayout',
    );
    expect(
      Array.from((layout as Element).attributes).map(
        (attribute) => attribute.name,
      ),
      entry.text,
    ).toEqual(entry.qualifiedAttributes);
  }

  const nary = descendants(document, 'nary').find((candidate) =>
    candidate.textContent?.includes('operator-east-asian-layout'),
  );
  expect(nary).toBeDefined();
  const naryProperties = directChildren(nary as Element, 'naryPr')[0];
  const controlProperties = directChildren(naryProperties, 'ctrlPr')[0];
  const wordProperties = directChildren(controlProperties, 'rPr').find(
    (candidate) => candidate.namespaceURI === WORD_NAMESPACE,
  );
  expect(wordProperties).toBeDefined();
  const layout = directChildren(
    wordProperties as Element,
    'eastAsianLayout',
  )[0];
  expect(wordAttributes(layout)).toEqual({
    id: '-7',
    combine: '0',
    combineBrackets: 'angle',
    vert: '1',
    vertCompress: '0',
  });
  expect(
    Array.from(layout.attributes).map((attribute) => attribute.name),
  ).toEqual([
    'w:id',
    'w:combine',
    'w:combineBrackets',
    'w:vert',
    'w:vertCompress',
  ]);
}

async function expectNativeWordRunSpecVanish(blob: Blob): Promise<void> {
  const archive = await JSZip.loadAsync(await blob.arrayBuffer());
  const document = await xmlEntry(archive, 'word/document.xml');
  const mathRuns = descendants(document, 'r').filter(
    (run) => run.namespaceURI === MATH_NAMESPACE,
  );
  for (const entry of [
    {
      text: 'always-hidden-paragraph-mark',
      names: ['specVanish'],
      attributes: [{ val: '1' }],
    },
    {
      text: 'explicit-paragraph-mark-reset',
      names: ['specVanish'],
      attributes: [{ val: '0' }],
    },
    {
      text: 'spec-vanish-with-hidden-run',
      names: ['vanish', 'specVanish'],
      attributes: [{ val: '1' }, { val: '1' }],
    },
    {
      text: 'ordered-spec-vanish',
      names: ['lang', 'eastAsianLayout', 'specVanish'],
      attributes: [
        { eastAsia: 'zh-CN' },
        { id: '5', combine: '0', vert: '1' },
        { val: '0' },
      ],
    },
  ]) {
    const run = mathRuns.find(
      (candidate) => candidate.textContent === entry.text,
    );
    expect(run, entry.text).toBeDefined();
    const properties = directChildren(run as Element, 'rPr').find(
      (candidate) => candidate.namespaceURI === WORD_NAMESPACE,
    );
    expect(properties, entry.text).toBeDefined();
    const children = directChildren(properties as Element);
    expect(
      children.map((child) => child.localName),
      entry.text,
    ).toEqual(entry.names);
    expect(children.map(wordAttributes), entry.text).toEqual(entry.attributes);
    const specVanish = children.find(
      (child) => child.localName === 'specVanish',
    );
    expect(
      Array.from((specVanish as Element).attributes).map(
        (attribute) => attribute.name,
      ),
      entry.text,
    ).toEqual(['w:val']);
  }

  const nary = descendants(document, 'nary').find((candidate) =>
    candidate.textContent?.includes('operator-spec-vanish'),
  );
  expect(nary).toBeDefined();
  const naryProperties = directChildren(nary as Element, 'naryPr')[0];
  const controlProperties = directChildren(naryProperties, 'ctrlPr')[0];
  const wordProperties = directChildren(controlProperties, 'rPr').find(
    (candidate) => candidate.namespaceURI === WORD_NAMESPACE,
  );
  expect(wordProperties).toBeDefined();
  const specVanish = directChildren(wordProperties as Element, 'specVanish')[0];
  expect(wordAttributes(specVanish)).toEqual({ val: '1' });
  expect(
    Array.from(specVanish.attributes).map((attribute) => attribute.name),
  ).toEqual(['w:val']);
}

async function expectNativeWordRunGlows(blob: Blob): Promise<void> {
  const archive = await JSZip.loadAsync(await blob.arrayBuffer());
  const document = await xmlEntry(archive, 'word/document.xml');
  const root = document.documentElement;
  const ignorable = Array.from(root.attributes).find(
    (attribute) =>
      xmlAttributeNamespace(root, attribute) ===
        MARKUP_COMPATIBILITY_NAMESPACE &&
      xmlAttributeLocalName(attribute) === 'Ignorable',
  );
  expect(ignorable).toBeDefined();
  expect(
    (ignorable?.value ?? '')
      .trim()
      .split(/\s+/u)
      .filter(Boolean)
      .some((prefix) => xmlNamespaceUri(root, prefix) === WORD_2010_NAMESPACE),
  ).toBe(true);

  const mathRuns = descendants(document, 'r').filter(
    (run) => run.namespaceURI === MATH_NAMESPACE,
  );
  const wordPropertiesFor = (text: string): Element => {
    const run = mathRuns.find((candidate) => candidate.textContent === text);
    expect(run, text).toBeDefined();
    const properties = directChildren(run as Element, 'rPr').find(
      (candidate) => candidate.namespaceURI === WORD_NAMESPACE,
    );
    expect(properties, text).toBeDefined();
    return properties as Element;
  };

  const rgbGlow = directChildren(wordPropertiesFor('rgb-glow'), 'glow')[0];
  expect(rgbGlow.namespaceURI).toBe(WORD_2010_NAMESPACE);
  expect(word2010Attributes(rgbGlow)).toEqual({ rad: '63500' });
  const rgbColor = directChildren(rgbGlow)[0];
  expect(rgbColor.localName).toBe('srgbClr');
  expect(rgbColor.namespaceURI).toBe(WORD_2010_NAMESPACE);
  expect(word2010Attributes(rgbColor)).toEqual({ val: 'A1B2C3' });
  const rgbTransforms = directChildren(rgbColor);
  expect(rgbTransforms.map((child) => child.localName)).toEqual([
    'tint',
    'shade',
    'alpha',
    'hueMod',
    'sat',
    'satOff',
    'satMod',
    'lum',
    'lumOff',
    'lumMod',
    'alpha',
  ]);
  expect(rgbTransforms.map(word2010Attributes)).toEqual(
    [
      0, 100_000, 50_000, 200_000, -25_000, 25_000, 50_000, 40_000, -10_000,
      75_000, 100_000,
    ].map((value) => ({ val: String(value) })),
  );

  const schemeGlow = directChildren(
    wordPropertiesFor('scheme-glow'),
    'glow',
  )[0];
  expect(schemeGlow.namespaceURI).toBe(WORD_2010_NAMESPACE);
  expect(word2010Attributes(schemeGlow)).toEqual({ rad: '0' });
  const schemeColor = directChildren(schemeGlow)[0];
  expect(schemeColor.localName).toBe('schemeClr');
  expect(schemeColor.namespaceURI).toBe(WORD_2010_NAMESPACE);
  expect(word2010Attributes(schemeColor)).toEqual({ val: 'phClr' });

  const defaultGlow = directChildren(
    wordPropertiesFor('default-radius-glow'),
    'glow',
  )[0];
  expect(word2010Attributes(defaultGlow)).toEqual({});
  expect(word2010Attributes(directChildren(defaultGlow)[0])).toEqual({
    val: 'accent2',
  });

  const orderedProperties = directChildren(wordPropertiesFor('ordered-glow'));
  expect(orderedProperties.map((child) => child.localName)).toEqual([
    'shadow',
    'specVanish',
    'glow',
  ]);
  expect(orderedProperties.map((child) => child.namespaceURI)).toEqual([
    WORD_NAMESPACE,
    WORD_NAMESPACE,
    WORD_2010_NAMESPACE,
  ]);

  const nary = descendants(document, 'nary').find((candidate) =>
    candidate.textContent?.includes('operator-glow'),
  );
  expect(nary).toBeDefined();
  const naryProperties = directChildren(nary as Element, 'naryPr')[0];
  const controlProperties = directChildren(naryProperties, 'ctrlPr')[0];
  const wordProperties = directChildren(controlProperties, 'rPr').find(
    (candidate) => candidate.namespaceURI === WORD_NAMESPACE,
  );
  expect(wordProperties).toBeDefined();
  const glow = directChildren(wordProperties as Element, 'glow')[0];
  expect(glow.namespaceURI).toBe(WORD_2010_NAMESPACE);
  expect(word2010Attributes(glow)).toEqual({ rad: '12700' });
  expect(word2010Attributes(directChildren(glow)[0])).toEqual({
    val: 'hlink',
  });
}

async function expectNativeWordRunShadowEffects(blob: Blob): Promise<void> {
  const archive = await JSZip.loadAsync(await blob.arrayBuffer());
  const document = await xmlEntry(archive, 'word/document.xml');
  const root = document.documentElement;
  const ignorable = Array.from(root.attributes).find(
    (attribute) =>
      xmlAttributeNamespace(root, attribute) ===
        MARKUP_COMPATIBILITY_NAMESPACE &&
      xmlAttributeLocalName(attribute) === 'Ignorable',
  );
  expect(ignorable).toBeDefined();
  expect(
    (ignorable?.value ?? '')
      .trim()
      .split(/\s+/u)
      .filter(Boolean)
      .some((prefix) => xmlNamespaceUri(root, prefix) === WORD_2010_NAMESPACE),
  ).toBe(true);

  const mathRuns = descendants(document, 'r').filter(
    (run) => run.namespaceURI === MATH_NAMESPACE,
  );
  const wordPropertiesFor = (text: string): Element => {
    const run = mathRuns.find((candidate) => candidate.textContent === text);
    expect(run, text).toBeDefined();
    const properties = directChildren(run as Element, 'rPr').find(
      (candidate) => candidate.namespaceURI === WORD_NAMESPACE,
    );
    expect(properties, text).toBeDefined();
    return properties as Element;
  };

  const rgbShadow = directChildren(
    wordPropertiesFor('rgb-shadow-effect'),
    'shadow',
  )[0];
  expect(rgbShadow.namespaceURI).toBe(WORD_2010_NAMESPACE);
  expect(word2010Attributes(rgbShadow)).toEqual({
    blurRad: '63500',
    dist: '12700',
    dir: '2700000',
    sx: '100000',
    sy: '-50000',
    kx: '-60000',
    ky: '120000',
    algn: 'br',
  });
  const rgbColor = directChildren(rgbShadow)[0];
  expect(rgbColor.localName).toBe('srgbClr');
  expect(rgbColor.namespaceURI).toBe(WORD_2010_NAMESPACE);
  expect(word2010Attributes(rgbColor)).toEqual({ val: 'A1B2C3' });
  const transforms = directChildren(rgbColor);
  expect(transforms.map((child) => child.localName)).toEqual([
    'alpha',
    'satOff',
    'alpha',
  ]);
  expect(transforms.map(word2010Attributes)).toEqual([
    { val: '50000' },
    { val: '-25000' },
    { val: '100000' },
  ]);

  const schemeShadow = directChildren(
    wordPropertiesFor('scheme-shadow-effect'),
    'shadow',
  )[0];
  expect(schemeShadow.namespaceURI).toBe(WORD_2010_NAMESPACE);
  expect(word2010Attributes(schemeShadow)).toEqual({
    blurRad: '0',
    dist: '0',
    dir: '0',
    sx: '0',
    sy: '0',
    kx: '0',
    ky: '0',
    algn: 'none',
  });
  const schemeColor = directChildren(schemeShadow)[0];
  expect(schemeColor.localName).toBe('schemeClr');
  expect(schemeColor.namespaceURI).toBe(WORD_2010_NAMESPACE);
  expect(word2010Attributes(schemeColor)).toEqual({ val: 'phClr' });

  const defaultShadow = directChildren(
    wordPropertiesFor('default-shadow-effect'),
    'shadow',
  )[0];
  expect(word2010Attributes(defaultShadow)).toEqual({});
  expect(word2010Attributes(directChildren(defaultShadow)[0])).toEqual({
    val: 'accent3',
  });

  const boundaryShadow = directChildren(
    wordPropertiesFor('boundary-shadow-effect'),
    'shadow',
  )[0];
  expect(boundaryShadow.namespaceURI).toBe(WORD_2010_NAMESPACE);
  expect(word2010Attributes(boundaryShadow)).toEqual({
    blurRad: '2147483647',
    dist: '2147483647',
    dir: '21599999',
    sx: '-2147483648',
    sy: '2147483647',
    kx: '-5399999',
    ky: '5399999',
    algn: 'tl',
  });
  expect(word2010Attributes(directChildren(boundaryShadow)[0])).toEqual({
    val: 'ABCDEF',
  });

  const orderedProperties = directChildren(
    wordPropertiesFor('ordered-shadow-effect'),
  );
  expect(orderedProperties.map((child) => child.localName)).toEqual([
    'shadow',
    'specVanish',
    'glow',
    'shadow',
  ]);
  expect(orderedProperties.map((child) => child.namespaceURI)).toEqual([
    WORD_NAMESPACE,
    WORD_NAMESPACE,
    WORD_2010_NAMESPACE,
    WORD_2010_NAMESPACE,
  ]);

  const nary = descendants(document, 'nary').find((candidate) =>
    candidate.textContent?.includes('operator-shadow-effect'),
  );
  expect(nary).toBeDefined();
  const naryProperties = directChildren(nary as Element, 'naryPr')[0];
  const controlProperties = directChildren(naryProperties, 'ctrlPr')[0];
  const wordProperties = directChildren(controlProperties, 'rPr').find(
    (candidate) => candidate.namespaceURI === WORD_NAMESPACE,
  );
  expect(wordProperties).toBeDefined();
  const shadow = directChildren(wordProperties as Element, 'shadow').find(
    (candidate) => candidate.namespaceURI === WORD_2010_NAMESPACE,
  );
  expect(shadow).toBeDefined();
  expect(word2010Attributes(shadow as Element)).toEqual({
    blurRad: '25400',
    algn: 'ctr',
  });
  expect(word2010Attributes(directChildren(shadow as Element)[0])).toEqual({
    val: 'hlink',
  });
}

async function expectNativeWordRunReflectionEffects(blob: Blob): Promise<void> {
  const archive = await JSZip.loadAsync(await blob.arrayBuffer());
  const document = await xmlEntry(archive, 'word/document.xml');
  const root = document.documentElement;
  const ignorable = Array.from(root.attributes).find(
    (attribute) =>
      xmlAttributeNamespace(root, attribute) ===
        MARKUP_COMPATIBILITY_NAMESPACE &&
      xmlAttributeLocalName(attribute) === 'Ignorable',
  );
  expect(ignorable).toBeDefined();
  expect(
    (ignorable?.value ?? '')
      .trim()
      .split(/\s+/u)
      .filter(Boolean)
      .some((prefix) => xmlNamespaceUri(root, prefix) === WORD_2010_NAMESPACE),
  ).toBe(true);

  const mathRuns = descendants(document, 'r').filter(
    (run) => run.namespaceURI === MATH_NAMESPACE,
  );
  const wordPropertiesFor = (text: string): Element => {
    const run = mathRuns.find((candidate) => candidate.textContent === text);
    expect(run, text).toBeDefined();
    const properties = directChildren(run as Element, 'rPr').find(
      (candidate) => candidate.namespaceURI === WORD_NAMESPACE,
    );
    expect(properties, text).toBeDefined();
    return properties as Element;
  };
  const reflectionFor = (text: string): Element => {
    const reflection = directChildren(wordPropertiesFor(text), 'reflection')[0];
    expect(reflection, text).toBeDefined();
    expect(reflection.namespaceURI, text).toBe(WORD_2010_NAMESPACE);
    expect(directChildren(reflection), text).toHaveLength(0);
    return reflection;
  };

  expect(word2010Attributes(reflectionFor('full-reflection-effect'))).toEqual({
    blurRad: '63500',
    stA: '50000',
    stPos: '0',
    endA: '10000',
    endPos: '100000',
    dist: '12700',
    dir: '5400000',
    fadeDir: '16200000',
    sx: '100000',
    sy: '-50000',
    kx: '-60000',
    ky: '120000',
    algn: 'b',
  });
  expect(word2010Attributes(reflectionFor('zero-reflection-effect'))).toEqual({
    blurRad: '0',
    stA: '0',
    stPos: '0',
    endA: '0',
    endPos: '0',
    dist: '0',
    dir: '0',
    fadeDir: '0',
    sx: '0',
    sy: '0',
    kx: '0',
    ky: '0',
    algn: 'none',
  });
  expect(
    word2010Attributes(reflectionFor('default-reflection-effect')),
  ).toEqual({});
  expect(
    word2010Attributes(reflectionFor('boundary-reflection-effect')),
  ).toEqual({
    blurRad: '2147483647',
    stA: '0',
    stPos: '100000',
    endA: '100000',
    endPos: '0',
    dist: '2147483647',
    dir: '21599999',
    fadeDir: '21599999',
    sx: '-2147483648',
    sy: '2147483647',
    kx: '-5399999',
    ky: '5399999',
    algn: 'tl',
  });

  const orderedProperties = directChildren(
    wordPropertiesFor('ordered-reflection-effect'),
  );
  expect(orderedProperties.map((child) => child.localName)).toEqual([
    'shadow',
    'specVanish',
    'glow',
    'shadow',
    'reflection',
  ]);
  expect(orderedProperties.map((child) => child.namespaceURI)).toEqual([
    WORD_NAMESPACE,
    WORD_NAMESPACE,
    WORD_2010_NAMESPACE,
    WORD_2010_NAMESPACE,
    WORD_2010_NAMESPACE,
  ]);
  expect(word2010Attributes(orderedProperties.at(-1) as Element)).toEqual({
    stA: '25000',
  });

  const nary = descendants(document, 'nary').find((candidate) =>
    candidate.textContent?.includes('operator-reflection-effect'),
  );
  expect(nary).toBeDefined();
  const naryProperties = directChildren(nary as Element, 'naryPr')[0];
  const controlProperties = directChildren(naryProperties, 'ctrlPr')[0];
  const wordProperties = directChildren(controlProperties, 'rPr').find(
    (candidate) => candidate.namespaceURI === WORD_NAMESPACE,
  );
  expect(wordProperties).toBeDefined();
  const reflection = directChildren(wordProperties as Element, 'reflection')[0];
  expect(reflection.namespaceURI).toBe(WORD_2010_NAMESPACE);
  expect(directChildren(reflection)).toHaveLength(0);
  expect(word2010Attributes(reflection)).toEqual({
    blurRad: '25400',
    endA: '75000',
    algn: 'ctr',
  });
}

async function expectNativeWordRunTextOutlineEffects(
  blob: Blob,
): Promise<void> {
  const archive = await JSZip.loadAsync(await blob.arrayBuffer());
  const document = await xmlEntry(archive, 'word/document.xml');
  const root = document.documentElement;
  const ignorable = Array.from(root.attributes).find(
    (attribute) =>
      xmlAttributeNamespace(root, attribute) ===
        MARKUP_COMPATIBILITY_NAMESPACE &&
      xmlAttributeLocalName(attribute) === 'Ignorable',
  );
  expect(ignorable).toBeDefined();
  expect(
    (ignorable?.value ?? '')
      .trim()
      .split(/\s+/u)
      .filter(Boolean)
      .some((prefix) => xmlNamespaceUri(root, prefix) === WORD_2010_NAMESPACE),
  ).toBe(true);

  const mathRuns = descendants(document, 'r').filter(
    (run) => run.namespaceURI === MATH_NAMESPACE,
  );
  const wordPropertiesFor = (text: string): Element => {
    const run = mathRuns.find((candidate) => candidate.textContent === text);
    expect(run, text).toBeDefined();
    const properties = directChildren(run as Element, 'rPr').find(
      (candidate) => candidate.namespaceURI === WORD_NAMESPACE,
    );
    expect(properties, text).toBeDefined();
    return properties as Element;
  };
  const outlineFor = (text: string): Element => {
    const outline = directChildren(wordPropertiesFor(text), 'textOutline')[0];
    expect(outline, text).toBeDefined();
    expect(outline.namespaceURI, text).toBe(WORD_2010_NAMESPACE);
    return outline;
  };
  const expectWord2010Children = (
    element: Element,
    names: readonly string[],
  ): Element[] => {
    const children = directChildren(element);
    expect(children.map((child) => child.localName)).toEqual(names);
    expect(
      children.every((child) => child.namespaceURI === WORD_2010_NAMESPACE),
    ).toBe(true);
    return children;
  };

  const full = outlineFor('full-text-outline-effect');
  expect(word2010Attributes(full)).toEqual({
    w: '12700',
    cap: 'sq',
    cmpd: 'dbl',
    algn: 'in',
  });
  const [gradient, dash, miter] = expectWord2010Children(full, [
    'gradFill',
    'prstDash',
    'miter',
  ]);
  expect(word2010Attributes(gradient)).toEqual({});
  const [stopList, linear] = expectWord2010Children(gradient, ['gsLst', 'lin']);
  expect(word2010Attributes(stopList)).toEqual({});
  const stops = expectWord2010Children(stopList, ['gs', 'gs']);
  expect(word2010Attributes(stops[0])).toEqual({ pos: '0' });
  expect(word2010Attributes(stops[1])).toEqual({ pos: '100000' });
  const firstColor = expectWord2010Children(stops[0], ['srgbClr'])[0];
  expect(word2010Attributes(firstColor)).toEqual({ val: 'A1B2C3' });
  const transforms = expectWord2010Children(firstColor, [
    'tint',
    'alpha',
    'lumMod',
  ]);
  expect(transforms.map(word2010Attributes)).toEqual([
    { val: '0' },
    { val: '50000' },
    { val: '75000' },
  ]);
  const secondColor = expectWord2010Children(stops[1], ['schemeClr'])[0];
  expect(word2010Attributes(secondColor)).toEqual({ val: 'accent3' });
  expect(directChildren(secondColor)).toHaveLength(0);
  expect(word2010Attributes(linear)).toEqual({
    ang: '5400000',
    scaled: '1',
  });
  expect(directChildren(linear)).toHaveLength(0);
  expect(word2010Attributes(dash)).toEqual({ val: 'lgDashDotDot' });
  expect(directChildren(dash)).toHaveLength(0);
  expect(word2010Attributes(miter)).toEqual({ lim: '250000' });
  expect(directChildren(miter)).toHaveLength(0);

  const explicitDefaults = outlineFor('explicit-default-text-outline-effect');
  expect(word2010Attributes(explicitDefaults)).toEqual({
    w: '0',
    cap: 'flat',
    cmpd: 'sng',
    algn: 'ctr',
  });
  const [noFill, solidDash, bevel] = expectWord2010Children(explicitDefaults, [
    'noFill',
    'prstDash',
    'bevel',
  ]);
  expect(word2010Attributes(noFill)).toEqual({});
  expect(directChildren(noFill)).toHaveLength(0);
  expect(word2010Attributes(solidDash)).toEqual({ val: 'solid' });
  expect(directChildren(solidDash)).toHaveLength(0);
  expect(word2010Attributes(bevel)).toEqual({});
  expect(directChildren(bevel)).toHaveLength(0);

  const empty = outlineFor('empty-text-outline-effect');
  expect(word2010Attributes(empty)).toEqual({});
  expect(directChildren(empty)).toHaveLength(0);
  const emptyComponents = outlineFor('empty-components-text-outline-effect');
  expect(word2010Attributes(emptyComponents)).toEqual({});
  for (const child of expectWord2010Children(emptyComponents, [
    'solidFill',
    'prstDash',
    'miter',
  ])) {
    expect(word2010Attributes(child)).toEqual({});
    expect(directChildren(child)).toHaveLength(0);
  }

  const boundary = outlineFor('boundary-text-outline-effect');
  expect(word2010Attributes(boundary)).toEqual({
    w: '20116800',
    cap: 'rnd',
    cmpd: 'tri',
    algn: 'in',
  });
  const [boundaryGradient, emptyDash, maximumMiter] = expectWord2010Children(
    boundary,
    ['gradFill', 'prstDash', 'miter'],
  );
  const [boundaryStopList, path] = expectWord2010Children(boundaryGradient, [
    'gsLst',
    'path',
  ]);
  const boundaryStops = directChildren(boundaryStopList);
  expect(boundaryStops).toHaveLength(10);
  expect(boundaryStops.map((stop) => word2010Attributes(stop).pos)).toEqual([
    '0',
    '1',
    '10000',
    '20000',
    '30000',
    '40000',
    '50000',
    '60000',
    '99999',
    '100000',
  ]);
  expect(
    boundaryStops.map((stop) =>
      word2010Attributes(expectWord2010Children(stop, ['srgbClr'])[0]),
    ),
  ).toEqual(
    Array.from({ length: 10 }, (_, index) => ({
      val: index.toString(16).padStart(6, '0').toUpperCase(),
    })),
  );
  expect(word2010Attributes(path)).toEqual({ path: 'rect' });
  const fillToRectangle = expectWord2010Children(path, ['fillToRect'])[0];
  expect(word2010Attributes(fillToRectangle)).toEqual({
    l: '-2147483648',
    t: '2147483647',
    r: '0',
    b: '1',
  });
  expect(directChildren(fillToRectangle)).toHaveLength(0);
  expect(word2010Attributes(emptyDash)).toEqual({});
  expect(directChildren(emptyDash)).toHaveLength(0);
  expect(word2010Attributes(maximumMiter)).toEqual({ lim: '2147483647' });
  expect(directChildren(maximumMiter)).toHaveLength(0);

  const orderedProperties = directChildren(
    wordPropertiesFor('ordered-text-outline-effect'),
  );
  expect(orderedProperties.map((child) => child.localName)).toEqual([
    'shadow',
    'specVanish',
    'glow',
    'shadow',
    'reflection',
    'textOutline',
  ]);
  expect(orderedProperties.map((child) => child.namespaceURI)).toEqual([
    WORD_NAMESPACE,
    WORD_NAMESPACE,
    WORD_2010_NAMESPACE,
    WORD_2010_NAMESPACE,
    WORD_2010_NAMESPACE,
    WORD_2010_NAMESPACE,
  ]);
  const orderedOutline = orderedProperties.at(-1) as Element;
  expect(word2010Attributes(orderedOutline)).toEqual({});
  expectWord2010Children(orderedOutline, ['noFill']);

  const nary = descendants(document, 'nary').find((candidate) =>
    candidate.textContent?.includes('operator-text-outline-effect'),
  );
  expect(nary).toBeDefined();
  const naryProperties = directChildren(nary as Element, 'naryPr')[0];
  const controlProperties = directChildren(naryProperties, 'ctrlPr')[0];
  const wordProperties = directChildren(controlProperties, 'rPr').find(
    (candidate) => candidate.namespaceURI === WORD_NAMESPACE,
  );
  expect(wordProperties).toBeDefined();
  const controlOutline = directChildren(
    wordProperties as Element,
    'textOutline',
  )[0];
  expect(controlOutline.namespaceURI).toBe(WORD_2010_NAMESPACE);
  expect(word2010Attributes(controlOutline)).toEqual({ w: '25400' });
  const [controlFill, controlJoin] = expectWord2010Children(controlOutline, [
    'solidFill',
    'round',
  ]);
  const controlColor = expectWord2010Children(controlFill, ['schemeClr'])[0];
  expect(word2010Attributes(controlColor)).toEqual({ val: 'hlink' });
  expect(directChildren(controlColor)).toHaveLength(0);
  expect(word2010Attributes(controlJoin)).toEqual({});
  expect(directChildren(controlJoin)).toHaveLength(0);
}

async function expectNativeWordRunTextFillEffects(blob: Blob): Promise<void> {
  const archive = await JSZip.loadAsync(await blob.arrayBuffer());
  const document = await xmlEntry(archive, 'word/document.xml');
  const root = document.documentElement;
  const ignorable = Array.from(root.attributes).find(
    (attribute) =>
      xmlAttributeNamespace(root, attribute) ===
        MARKUP_COMPATIBILITY_NAMESPACE &&
      xmlAttributeLocalName(attribute) === 'Ignorable',
  );
  expect(ignorable).toBeDefined();
  expect(
    (ignorable?.value ?? '')
      .trim()
      .split(/\s+/u)
      .filter(Boolean)
      .some((prefix) => xmlNamespaceUri(root, prefix) === WORD_2010_NAMESPACE),
  ).toBe(true);

  const mathRuns = descendants(document, 'r').filter(
    (run) => run.namespaceURI === MATH_NAMESPACE,
  );
  const wordPropertiesFor = (text: string): Element => {
    const run = mathRuns.find((candidate) => candidate.textContent === text);
    expect(run, text).toBeDefined();
    const properties = directChildren(run as Element, 'rPr').find(
      (candidate) => candidate.namespaceURI === WORD_NAMESPACE,
    );
    expect(properties, text).toBeDefined();
    return properties as Element;
  };
  const textFillFor = (text: string): Element => {
    const fill = directChildren(wordPropertiesFor(text), 'textFill')[0];
    expect(fill, text).toBeDefined();
    expect(fill.namespaceURI, text).toBe(WORD_2010_NAMESPACE);
    expect(word2010Attributes(fill), text).toEqual({});
    return fill;
  };
  const expectWord2010Children = (
    element: Element,
    names: readonly string[],
  ): Element[] => {
    const children = directChildren(element);
    expect(children.map((child) => child.localName)).toEqual(names);
    expect(
      children.every((child) => child.namespaceURI === WORD_2010_NAMESPACE),
    ).toBe(true);
    return children;
  };

  const full = textFillFor('full-text-fill-effect');
  const gradient = expectWord2010Children(full, ['gradFill'])[0];
  expect(word2010Attributes(gradient)).toEqual({});
  const [stopList, linear] = expectWord2010Children(gradient, ['gsLst', 'lin']);
  expect(word2010Attributes(stopList)).toEqual({});
  const stops = expectWord2010Children(stopList, ['gs', 'gs']);
  expect(word2010Attributes(stops[0])).toEqual({ pos: '0' });
  expect(word2010Attributes(stops[1])).toEqual({ pos: '100000' });
  const firstColor = expectWord2010Children(stops[0], ['srgbClr'])[0];
  expect(word2010Attributes(firstColor)).toEqual({ val: 'A1B2C3' });
  const transforms = expectWord2010Children(firstColor, [
    'tint',
    'alpha',
    'lumMod',
  ]);
  expect(transforms.map(word2010Attributes)).toEqual([
    { val: '0' },
    { val: '50000' },
    { val: '75000' },
  ]);
  const secondColor = expectWord2010Children(stops[1], ['schemeClr'])[0];
  expect(word2010Attributes(secondColor)).toEqual({ val: 'accent3' });
  expect(directChildren(secondColor)).toHaveLength(0);
  expect(word2010Attributes(linear)).toEqual({
    ang: '5400000',
    scaled: '1',
  });
  expect(directChildren(linear)).toHaveLength(0);

  const directProperties = wordPropertiesFor('direct-solid-text-fill-effect');
  expect(
    directChildren(directProperties).map((child) => child.localName),
  ).toEqual(['color', 'textFill']);
  const direct = textFillFor('direct-solid-text-fill-effect');
  const solid = expectWord2010Children(direct, ['solidFill'])[0];
  const solidColor = expectWord2010Children(solid, ['srgbClr'])[0];
  expect(word2010Attributes(solidColor)).toEqual({ val: '13579B' });
  expect(directChildren(solidColor)).toHaveLength(0);

  expect(directChildren(textFillFor('empty-text-fill-effect'))).toHaveLength(0);
  const emptySolid = expectWord2010Children(
    textFillFor('empty-solid-text-fill-effect'),
    ['solidFill'],
  )[0];
  expect(word2010Attributes(emptySolid)).toEqual({});
  expect(directChildren(emptySolid)).toHaveLength(0);
  const emptyGradient = expectWord2010Children(
    textFillFor('empty-gradient-text-fill-effect'),
    ['gradFill'],
  )[0];
  expect(word2010Attributes(emptyGradient)).toEqual({});
  expect(directChildren(emptyGradient)).toHaveLength(0);
  const noFill = expectWord2010Children(textFillFor('no-text-fill-effect'), [
    'noFill',
  ])[0];
  expect(word2010Attributes(noFill)).toEqual({});
  expect(directChildren(noFill)).toHaveLength(0);

  const boundary = textFillFor('boundary-text-fill-effect');
  const boundaryGradient = expectWord2010Children(boundary, ['gradFill'])[0];
  const [boundaryStopList, path] = expectWord2010Children(boundaryGradient, [
    'gsLst',
    'path',
  ]);
  const boundaryStops = directChildren(boundaryStopList);
  expect(boundaryStops).toHaveLength(10);
  expect(boundaryStops.map((stop) => word2010Attributes(stop).pos)).toEqual([
    '0',
    '1',
    '10000',
    '20000',
    '30000',
    '40000',
    '50000',
    '60000',
    '99999',
    '100000',
  ]);
  expect(
    boundaryStops.map((stop) =>
      word2010Attributes(expectWord2010Children(stop, ['srgbClr'])[0]),
    ),
  ).toEqual(
    Array.from({ length: 10 }, (_, index) => ({
      val: index.toString(16).padStart(6, '0').toUpperCase(),
    })),
  );
  expect(word2010Attributes(path)).toEqual({ path: 'rect' });
  const rectangle = expectWord2010Children(path, ['fillToRect'])[0];
  expect(word2010Attributes(rectangle)).toEqual({
    l: '-2147483648',
    t: '2147483647',
    r: '0',
    b: '1',
  });
  expect(directChildren(rectangle)).toHaveLength(0);

  const orderedProperties = directChildren(
    wordPropertiesFor('ordered-text-fill-effect'),
  );
  expect(orderedProperties.map((child) => child.localName)).toEqual([
    'specVanish',
    'reflection',
    'textOutline',
    'textFill',
  ]);
  expect(orderedProperties.map((child) => child.namespaceURI)).toEqual([
    WORD_NAMESPACE,
    WORD_2010_NAMESPACE,
    WORD_2010_NAMESPACE,
    WORD_2010_NAMESPACE,
  ]);

  const nary = descendants(document, 'nary').find((candidate) =>
    candidate.textContent?.includes('operator-text-fill-effect'),
  );
  expect(nary).toBeDefined();
  const naryProperties = directChildren(nary as Element, 'naryPr')[0];
  const controlProperties = directChildren(naryProperties, 'ctrlPr')[0];
  const wordProperties = directChildren(controlProperties, 'rPr').find(
    (candidate) => candidate.namespaceURI === WORD_NAMESPACE,
  );
  expect(wordProperties).toBeDefined();
  const controlTextFill = directChildren(
    wordProperties as Element,
    'textFill',
  )[0];
  expect(controlTextFill.namespaceURI).toBe(WORD_2010_NAMESPACE);
  expect(word2010Attributes(controlTextFill)).toEqual({});
  const controlSolid = expectWord2010Children(controlTextFill, [
    'solidFill',
  ])[0];
  const controlColor = expectWord2010Children(controlSolid, ['schemeClr'])[0];
  expect(word2010Attributes(controlColor)).toEqual({ val: 'hlink' });
  expect(directChildren(controlColor)).toHaveLength(0);
}

async function expectNativeControlProperties(blob: Blob): Promise<void> {
  const archive = await JSZip.loadAsync(await blob.arrayBuffer());
  const document = await xmlEntry(archive, 'word/document.xml');
  const propertyContainerNames = [
    'fPr',
    'sSupPr',
    'sSubPr',
    'sSubSupPr',
    'sPrePr',
    'limLowPr',
    'limUppPr',
    'radPr',
    'funcPr',
    'naryPr',
    'accPr',
    'barPr',
    'groupChrPr',
    'phantPr',
    'borderBoxPr',
    'boxPr',
    'mPr',
    'eqArrPr',
    'dPr',
  ];
  const expectedWordPropertyNames = [
    'rFonts',
    'b',
    'bCs',
    'i',
    'iCs',
    'caps',
    'smallCaps',
    'strike',
    'dstrike',
    'outline',
    'shadow',
    'emboss',
    'imprint',
    'noProof',
    'snapToGrid',
    'vanish',
    'webHidden',
    'color',
    'spacing',
    'w',
    'kern',
    'position',
    'sz',
    'szCs',
    'u',
    'effect',
    'bdr',
    'fitText',
    'vertAlign',
    'rtl',
    'cs',
    'em',
    'lang',
    'eastAsianLayout',
    'specVanish',
    'glow',
    'shadow',
    'reflection',
    'textOutline',
    'textFill',
  ];
  for (const name of propertyContainerNames) {
    const containers = descendants(document, name).filter(
      (candidate) => candidate.namespaceURI === MATH_NAMESPACE,
    );
    expect(containers, name).toHaveLength(1);
    const children = directChildren(containers[0]);
    const controlProperties = children.at(-1);
    expect(controlProperties?.localName, name).toBe('ctrlPr');
    expect(controlProperties?.namespaceURI, name).toBe(MATH_NAMESPACE);
    const wordProperties = directChildren(controlProperties as Element);
    expect(
      wordProperties.map(
        (child) =>
          `${child.namespaceURI === WORD_NAMESPACE ? 'w' : '?'}:${child.localName}`,
      ),
      name,
    ).toEqual(['w:rPr']);
    expect(
      directChildren(wordProperties[0]).map((child) => child.localName),
      name,
    ).toEqual(expectedWordPropertyNames);
  }
}

async function expectNativeArgumentControlProperties(
  blob: Blob,
): Promise<void> {
  const archive = await JSZip.loadAsync(await blob.arrayBuffer());
  const document = await xmlEntry(archive, 'word/document.xml');
  const argumentNames = new Set([
    'deg',
    'den',
    'e',
    'fName',
    'lim',
    'num',
    'sub',
    'sup',
  ]);
  const controlProperties = descendants(document, 'ctrlPr').filter(
    (candidate) => candidate.namespaceURI === MATH_NAMESPACE,
  );
  expect(controlProperties).toHaveLength(33);
  const expectedWordPropertyNames = [
    'rFonts',
    'b',
    'bCs',
    'i',
    'iCs',
    'caps',
    'smallCaps',
    'strike',
    'dstrike',
    'outline',
    'shadow',
    'emboss',
    'imprint',
    'noProof',
    'snapToGrid',
    'vanish',
    'webHidden',
    'color',
    'spacing',
    'w',
    'kern',
    'position',
    'sz',
    'szCs',
    'u',
    'effect',
    'bdr',
    'fitText',
    'vertAlign',
    'rtl',
    'cs',
    'em',
    'lang',
    'eastAsianLayout',
    'specVanish',
    'glow',
    'shadow',
    'reflection',
    'textOutline',
    'textFill',
  ];
  for (const controlProperty of controlProperties) {
    const argument = controlProperty.parentElement;
    expect(argumentNames.has(argument?.localName ?? '')).toBe(true);
    expect(argument?.namespaceURI).toBe(MATH_NAMESPACE);
    expect(directChildren(argument as Element).at(-1)).toBe(controlProperty);
    const wordProperties = directChildren(controlProperty);
    expect(
      wordProperties.map(
        (child) =>
          `${child.namespaceURI === WORD_NAMESPACE ? 'w' : '?'}:${child.localName}`,
      ),
    ).toEqual(['w:rPr']);
    expect(
      directChildren(wordProperties[0]).map((child) => child.localName),
    ).toEqual(expectedWordPropertyNames);
  }
}

async function expectNativeControlRevisions(blob: Blob): Promise<void> {
  const archive = await JSZip.loadAsync(await blob.arrayBuffer());
  const document = await xmlEntry(archive, 'word/document.xml');
  const revisionNames = new Set(['ins', 'del', 'moveFrom', 'moveTo']);
  const controlProperties = descendants(document, 'ctrlPr').filter(
    (candidate) =>
      candidate.namespaceURI === MATH_NAMESPACE &&
      revisionNames.has(directChildren(candidate)[0]?.localName ?? ''),
  );
  expect(controlProperties).toHaveLength(4);
  expect(
    controlProperties.map((controlProperty) => {
      const revisions: Array<{
        kind: string;
        id: string | undefined;
        author: string | undefined;
        date: string | undefined;
        dateUtc: string | undefined;
      }> = [];
      let current = directChildren(controlProperty)[0];
      while (current && revisionNames.has(current.localName)) {
        expect(current.namespaceURI).toBe(WORD_NAMESPACE);
        const attributes = wordAttributes(current);
        revisions.push({
          kind: current.localName,
          id: attributes.id,
          author: attributes.author,
          date: attributes.date,
          dateUtc: wordDateUtcAttribute(current),
        });
        const children = directChildren(current);
        expect(children.length).toBeLessThanOrEqual(1);
        current = children[0];
      }
      if (current) expect(current.namespaceURI).toBe(WORD_NAMESPACE);
      return {
        revisions,
        leaf:
          current?.localName === 'rPr'
            ? directChildren(current).map((child) => child.localName)
            : null,
      };
    }),
  ).toEqual([
    {
      revisions: [
        {
          kind: 'ins',
          id: '0',
          author: 'Alice & Bob',
          date: '2026-01-02T03:04:05Z',
          dateUtc: '2026-01-02T03:04:05.123Z',
        },
        {
          kind: 'del',
          id: '2147483647',
          author: 'Final reviewer',
          date: undefined,
          dateUtc: undefined,
        },
      ],
      leaf: ['b', 'color'],
    },
    {
      revisions: [
        {
          kind: 'del',
          id: '1',
          author: 'Deletion reviewer',
          date: undefined,
          dateUtc: undefined,
        },
      ],
      leaf: null,
    },
    {
      revisions: [
        {
          kind: 'moveFrom',
          id: '2',
          author: 'Move source reviewer',
          date: '2026-02-03T04:05:06.789+08:00',
          dateUtc: '2026-02-02T20:05:06.789Z',
        },
        {
          kind: 'ins',
          id: '3',
          author: 'Nested insertion reviewer',
          date: undefined,
          dateUtc: undefined,
        },
        {
          kind: 'del',
          id: '4',
          author: 'Nested deletion reviewer',
          date: undefined,
          dateUtc: undefined,
        },
      ],
      leaf: ['i', 'sz'],
    },
    {
      revisions: [
        {
          kind: 'moveTo',
          id: '5',
          author: 'Move destination reviewer',
          date: undefined,
          dateUtc: undefined,
        },
        {
          kind: 'del',
          id: '6',
          author: 'Destination deletion reviewer',
          date: undefined,
          dateUtc: undefined,
        },
      ],
      leaf: ['u'],
    },
  ]);
}

async function expectNativeControlRevisionOnEveryObject(
  blob: Blob,
): Promise<void> {
  const archive = await JSZip.loadAsync(await blob.arrayBuffer());
  const document = await xmlEntry(archive, 'word/document.xml');
  const propertyContainerNames = [
    'fPr',
    'sSupPr',
    'sSubPr',
    'sSubSupPr',
    'sPrePr',
    'limLowPr',
    'limUppPr',
    'radPr',
    'funcPr',
    'naryPr',
    'accPr',
    'barPr',
    'groupChrPr',
    'phantPr',
    'borderBoxPr',
    'boxPr',
    'mPr',
    'eqArrPr',
    'dPr',
  ];
  for (const [index, name] of propertyContainerNames.entries()) {
    const containers = descendants(document, name).filter(
      (candidate) => candidate.namespaceURI === MATH_NAMESPACE,
    );
    expect(containers, name).toHaveLength(1);
    const controlProperties = directChildren(containers[0]).at(-1);
    expect(controlProperties?.localName, name).toBe('ctrlPr');
    expect(controlProperties?.namespaceURI, name).toBe(MATH_NAMESPACE);
    const revision = directChildren(controlProperties as Element);
    expect(
      revision.map((child) => `${child.namespaceURI}:${child.localName}`),
      name,
    ).toEqual([`${WORD_NAMESPACE}:del`]);
    expect(wordAttributes(revision[0]), name).toEqual({
      id: String(index),
      author: `Object reviewer ${index}`,
    });
    expect(directChildren(revision[0]), name).toHaveLength(0);
  }
}

async function expectNativeArgumentSizes(blob: Blob): Promise<void> {
  const archive = await JSZip.loadAsync(await blob.arrayBuffer());
  const document = await xmlEntry(archive, 'word/document.xml');
  const sizes = descendants(document, 'argSz').filter(
    (candidate) => candidate.namespaceURI === MATH_NAMESPACE,
  );
  expect(
    sizes.map((size) => [
      size.parentElement?.parentElement?.textContent,
      mathValueAttribute(size),
    ]),
  ).toEqual([
    ['box-sized', '-2'],
    ['group-sized', '-1'],
    ['lower-limit-sized', '1'],
    ['upper-limit-sized', '2'],
    ['nary-sub-sized', '-2'],
    ['nary-sup-sized', '-1'],
    ['radical-degree-sized', '1'],
    ['pre-sub-sized', '2'],
    ['pre-sup-sized', '-2'],
    ['subscript-sized', '-1'],
    ['sub-super-sub-sized', '1'],
    ['sub-super-sup-sized', '2'],
    ['superscript-sized', '-2'],
    ['fraction-noop-sized', '1'],
    ['function-noop-sized', '-1'],
    ['matrix-noop-sized', '2'],
    ['array-noop-sized', '1'],
    ['delimiter-noop-sized', '-2'],
  ]);
  for (const size of sizes) {
    const argumentProperties = size.parentElement;
    const argument = argumentProperties?.parentElement;
    expect(argumentProperties?.localName).toBe('argPr');
    expect(argumentProperties?.namespaceURI).toBe(MATH_NAMESPACE);
    expect(directChildren(argumentProperties as Element)).toEqual([size]);
    expect(directChildren(argument as Element)[0]).toBe(argumentProperties);
  }
  const combinedArgument = sizes.find(
    (size) =>
      size.parentElement?.parentElement?.textContent === 'fraction-noop-sized',
  )?.parentElement?.parentElement;
  expect(
    directChildren(combinedArgument as Element).map((child) => child.localName),
  ).toEqual(['argPr', 'r', 'ctrlPr']);
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
  expect(fractions).toHaveLength(10);
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
    ['num', 'den'],
    ['fPr', 'num', 'den'],
    ['fPr', 'num', 'den'],
    ['fPr', 'num', 'den'],
    ['num', 'den'],
  ]);
  expect(
    fractions.map((fraction) => {
      const properties = directChildren(fraction, 'fPr')[0];
      return properties
        ? mathValueAttribute(directChildren(properties, 'type')[0])
        : null;
    }),
  ).toEqual([
    null,
    'noBar',
    'skw',
    'lin',
    null,
    null,
    'noBar',
    'skw',
    'lin',
    null,
  ]);
  expect(
    fractions.map((fraction) => [
      directChildren(fraction, 'num')[0]?.textContent ?? '',
      directChildren(fraction, 'den')[0]?.textContent ?? '',
    ]),
  ).toEqual([
    ['a+b', 'c'],
    ['n', 'k'],
    ['p', 'q'],
    ['u', 'v'],
    ['', ''],
    ['a+b', 'c'],
    ['n', 'k'],
    ['p', 'q'],
    ['u', 'v'],
    ['', ''],
  ]);
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
  const functions = descendants(document, 'func');
  expect(functions).toHaveLength(4);
  expect(
    functions.map((function_) =>
      directChildren(function_).map((child) => child.localName),
    ),
  ).toEqual(Array.from({ length: 4 }, () => ['funcPr', 'fName', 'e']));
  expect(
    functions.map((function_) => [
      directChildren(function_, 'fName')[0]?.textContent ?? '',
      directChildren(function_, 'e')[0]?.textContent ?? '',
    ]),
  ).toEqual([
    ['sin', 'θ'],
    ['', ''],
    ['sin', 'θ'],
    ['', ''],
  ]);
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
  const delimiters = descendants(document, 'd');
  expect(delimiters).toHaveLength(4);
  expect(
    delimiters.map((delimiter) =>
      directChildren(delimiter).map((child) => child.localName),
    ),
  ).toEqual([
    ['dPr', 'e', 'e'],
    ['dPr', 'e', 'e', 'e'],
    ['dPr', 'e', 'e'],
    ['dPr', 'e', 'e', 'e'],
  ]);
  expect(
    delimiters.map((delimiter) => {
      const properties = directChildren(delimiter, 'dPr')[0];
      return directChildren(properties).map((child) => child.localName);
    }),
  ).toEqual(Array.from({ length: 4 }, () => ['begChr', 'sepChr', 'endChr']));
  expect(
    delimiters.map((delimiter) => {
      const properties = directChildren(delimiter, 'dPr')[0];
      return directChildren(properties).map(mathValueAttribute);
    }),
  ).toEqual([
    ['[', ';', ']'],
    ['', '', ''],
    ['[', ';', ']'],
    ['', '', ''],
  ]);
  expect(
    delimiters.map((delimiter) =>
      directChildren(delimiter, 'e').map((argument) => argument.textContent),
    ),
  ).toEqual([
    ['a', 'b'],
    ['', 'z', ''],
    ['a', 'b'],
    ['', 'z', ''],
  ]);
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
  const styledWordRuns = styledRuns.filter(
    (run) => run.textContent === 'styledF',
  );
  expect(styledWordRuns).toHaveLength(2);
  expect(
    styledWordRuns.map((run) =>
      directChildren(run).map(
        (child) =>
          `${child.namespaceURI === MATH_NAMESPACE ? 'm' : child.namespaceURI === WORD_NAMESPACE ? 'w' : '?'}:${child.localName}`,
      ),
    ),
  ).toEqual([
    ['m:rPr', 'w:rPr', 'm:t'],
    ['m:rPr', 'w:rPr', 'm:t'],
  ]);
  const expectedWordPropertyNames = [
    'rFonts',
    'b',
    'bCs',
    'i',
    'iCs',
    'caps',
    'smallCaps',
    'strike',
    'dstrike',
    'outline',
    'shadow',
    'emboss',
    'imprint',
    'noProof',
    'snapToGrid',
    'vanish',
    'webHidden',
    'color',
    'spacing',
    'w',
    'kern',
    'position',
    'sz',
    'szCs',
    'u',
    'effect',
    'bdr',
    'fitText',
    'vertAlign',
    'rtl',
    'cs',
    'em',
    'lang',
    'eastAsianLayout',
    'specVanish',
    'glow',
    'shadow',
    'reflection',
    'textOutline',
    'textFill',
  ];
  const expectedWordPropertyAttributes = [
    {
      ascii: 'Cambria Math',
      hAnsi: 'Cambria Math',
      eastAsia: '等线',
      cs: 'Arial',
      asciiTheme: 'majorAscii',
      hAnsiTheme: 'majorHAnsi',
      eastAsiaTheme: 'minorEastAsia',
      cstheme: 'majorBidi',
      hint: 'eastAsia',
    },
    { val: '1' },
    { val: '0' },
    { val: '0' },
    { val: '1' },
    { val: '0' },
    { val: '1' },
    { val: '0' },
    { val: '1' },
    { val: '1' },
    { val: '1' },
    { val: '0' },
    { val: '0' },
    { val: '1' },
    { val: '0' },
    { val: '0' },
    { val: '1' },
    {
      val: '1A2B3C',
      themeColor: 'accent2',
      themeTint: '80',
      themeShade: '40',
    },
    { val: '20' },
    { val: '90' },
    { val: '22' },
    { val: '2' },
    { val: '25' },
    { val: '28' },
    {
      val: 'wavyDouble',
      color: 'ABCDEF',
      themeColor: 'accent3',
      themeTint: '20',
    },
    { val: 'shimmer' },
    {
      val: 'double',
      color: '112233',
      themeColor: 'accent1',
      themeTint: '66',
      themeShade: '33',
      sz: '24',
      space: '2',
      shadow: '1',
      frame: '0',
    },
    { val: '720', id: '50' },
    { val: 'baseline' },
    { val: '0' },
    { val: '0' },
    { val: 'circle' },
    { val: 'en-US', eastAsia: 'zh-CN', bidi: 'ar-SA' },
    {
      id: '9',
      combine: '1',
      combineBrackets: 'curly',
      vert: '0',
      vertCompress: '1',
    },
    { val: '0' },
    { rad: '63500' },
    {
      blurRad: '25400',
      dist: '12700',
      dir: '5400000',
      sx: '90000',
      sy: '-100000',
      kx: '-60000',
      ky: '120000',
      algn: 'tr',
    },
    {
      blurRad: '12700',
      stA: '60000',
      stPos: '0',
      endA: '10000',
      endPos: '100000',
      dist: '6350',
      dir: '5400000',
      fadeDir: '16200000',
      sx: '100000',
      sy: '-50000',
      kx: '-60000',
      ky: '120000',
      algn: 'b',
    },
    { w: '12700', cap: 'sq', cmpd: 'dbl', algn: 'in' },
    {},
  ];
  for (const run of styledWordRuns) {
    const properties = directChildren(run, 'rPr').find(
      (candidate) => candidate.namespaceURI === WORD_NAMESPACE,
    );
    expect(properties).toBeDefined();
    expect(directChildren(properties).map((child) => child.localName)).toEqual(
      expectedWordPropertyNames,
    );
    expect(
      directChildren(properties).map((child) =>
        child.namespaceURI === WORD_2010_NAMESPACE
          ? word2010Attributes(child)
          : wordAttributes(child),
      ),
    ).toEqual(expectedWordPropertyAttributes);
    const glow = directChildren(properties, 'glow')[0];
    expect(glow.namespaceURI).toBe(WORD_2010_NAMESPACE);
    const color = directChildren(glow)[0];
    expect(color.namespaceURI).toBe(WORD_2010_NAMESPACE);
    expect(color.localName).toBe('schemeClr');
    expect(word2010Attributes(color)).toEqual({ val: 'accent4' });
    const alpha = directChildren(color)[0];
    expect(alpha.namespaceURI).toBe(WORD_2010_NAMESPACE);
    expect(alpha.localName).toBe('alpha');
    expect(word2010Attributes(alpha)).toEqual({ val: '50000' });
    const shadow = directChildren(properties, 'shadow').find(
      (candidate) => candidate.namespaceURI === WORD_2010_NAMESPACE,
    );
    expect(shadow).toBeDefined();
    const shadowColor = directChildren(shadow as Element)[0];
    expect(shadowColor.namespaceURI).toBe(WORD_2010_NAMESPACE);
    expect(shadowColor.localName).toBe('schemeClr');
    expect(word2010Attributes(shadowColor)).toEqual({ val: 'accent5' });
    const luminanceModulation = directChildren(shadowColor)[0];
    expect(luminanceModulation.namespaceURI).toBe(WORD_2010_NAMESPACE);
    expect(luminanceModulation.localName).toBe('lumMod');
    expect(word2010Attributes(luminanceModulation)).toEqual({ val: '75000' });
    const reflection = directChildren(properties, 'reflection')[0];
    expect(reflection.namespaceURI).toBe(WORD_2010_NAMESPACE);
    expect(directChildren(reflection)).toHaveLength(0);
    const textOutline = directChildren(properties, 'textOutline')[0];
    expect(textOutline.namespaceURI).toBe(WORD_2010_NAMESPACE);
    const [solidFill, dash, miter] = directChildren(textOutline);
    expect(
      [solidFill, dash, miter].map((child) => [
        child.namespaceURI,
        child.localName,
      ]),
    ).toEqual([
      [WORD_2010_NAMESPACE, 'solidFill'],
      [WORD_2010_NAMESPACE, 'prstDash'],
      [WORD_2010_NAMESPACE, 'miter'],
    ]);
    expect(word2010Attributes(dash)).toEqual({ val: 'dashDot' });
    expect(word2010Attributes(miter)).toEqual({ lim: '400000' });
    const outlineColor = directChildren(solidFill)[0];
    expect(outlineColor.namespaceURI).toBe(WORD_2010_NAMESPACE);
    expect(outlineColor.localName).toBe('schemeClr');
    expect(word2010Attributes(outlineColor)).toEqual({ val: 'accent6' });
    const outlineAlpha = directChildren(outlineColor)[0];
    expect(outlineAlpha.namespaceURI).toBe(WORD_2010_NAMESPACE);
    expect(outlineAlpha.localName).toBe('alpha');
    expect(word2010Attributes(outlineAlpha)).toEqual({ val: '80000' });
    const textFill = directChildren(properties, 'textFill')[0];
    expect(textFill.namespaceURI).toBe(WORD_2010_NAMESPACE);
    expect(word2010Attributes(textFill)).toEqual({});
    const fillSolid = directChildren(textFill)[0];
    expect(fillSolid.namespaceURI).toBe(WORD_2010_NAMESPACE);
    expect(fillSolid.localName).toBe('solidFill');
    const fillColor = directChildren(fillSolid)[0];
    expect(fillColor.namespaceURI).toBe(WORD_2010_NAMESPACE);
    expect(fillColor.localName).toBe('srgbClr');
    expect(word2010Attributes(fillColor)).toEqual({ val: '2468AC' });
    const fillShade = directChildren(fillColor)[0];
    expect(fillShade.namespaceURI).toBe(WORD_2010_NAMESPACE);
    expect(fillShade.localName).toBe('shade');
    expect(word2010Attributes(fillShade)).toEqual({ val: '25000' });
  }
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
    expect(directChildren(properties).map((child) => child.localName)).toEqual([
      'baseJc',
      'plcHide',
      'rSpRule',
      'cGpRule',
      'rSp',
      'cSp',
      'cGp',
      'mcs',
    ]);
    expect(
      ['rSpRule', 'cGpRule', 'rSp', 'cSp', 'cGp'].map((name) =>
        mathValueAttribute(directChildren(properties, name)[0]),
      ),
    ).toEqual(['3', '4', '12', '120', '3']);
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
  const unsupportedComponents = directChildren(inline).filter((component) => {
    const equation = document.createElementNS(MATH_NAMESPACE, 'm:oMath');
    equation.setAttributeNS(XMLNS_NAMESPACE, 'xmlns:m', MATH_NAMESPACE);
    equation.setAttributeNS(XMLNS_NAMESPACE, 'xmlns:w', WORD_NAMESPACE);
    equation.setAttributeNS(XMLNS_NAMESPACE, 'xmlns:w14', WORD_2010_NAMESPACE);
    equation.append(component.cloneNode(true));
    return inspectDocxEquation(equation).status !== 'supported';
  });
  expect(unsupportedComponents.map((component) => component.outerHTML)).toEqual(
    [],
  );

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

function wordAttributes(element: Element): Record<string, string> {
  return Object.fromEntries(
    Array.from(element.attributes)
      .filter(
        (attribute) =>
          xmlAttributeNamespace(element, attribute) === WORD_NAMESPACE,
      )
      .map((attribute) => [xmlAttributeLocalName(attribute), attribute.value]),
  );
}

function word2010Attributes(element: Element): Record<string, string> {
  return Object.fromEntries(
    Array.from(element.attributes)
      .filter(
        (attribute) =>
          xmlAttributeNamespace(element, attribute) === WORD_2010_NAMESPACE,
      )
      .map((attribute) => [xmlAttributeLocalName(attribute), attribute.value]),
  );
}

function wordDateUtcAttribute(element: Element): string | undefined {
  return Array.from(element.attributes).find(
    (attribute) =>
      xmlAttributeNamespace(element, attribute) === WORD_DATE_UTC_NAMESPACE &&
      xmlAttributeLocalName(attribute) === 'dateUtc',
  )?.value;
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
