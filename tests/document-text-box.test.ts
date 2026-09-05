import { afterEach, describe, expect, test } from '@rstest/core';
import { Editor } from '@tiptap/core';
import JSZip from 'jszip';
import { createArtifact, importOfficeFile } from '../src/core';
import { createWorkDocumentExtensions } from '../src/internal/features/work/work-document-extensions';
import { createWorkDocumentModelFromContent } from '../src/internal/features/work/work-document-model-codec';
import {
  DOCUMENT_TEXT_BOX_LIMITS,
  normalizeDocumentTextBoxProperties,
  textBoxCss,
  textBoxDomAttributes,
} from '../src/internal/features/work/work-document-text-box';
import { createDocxBlob } from '../src/internal/features/work/work-docx-export';
import {
  applyImportedDocxTextBoxMarkers,
  inspectDocxTextBoxes,
  markDocxTextBoxes,
} from '../src/internal/features/work/work-docx-text-box-import';
import { analyzeDocxCompatibility } from '../src/internal/features/work/work-office-diagnostics';
import {
  attribute,
  descendants,
  directChild,
  parseXml,
} from '../src/internal/features/work/work-ooxml-package';

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const WORDPROCESSING_DRAWING_NAMESPACE =
  'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const DRAWING_NAMESPACE =
  'http://schemas.openxmlformats.org/drawingml/2006/main';
const WORDPROCESSING_SHAPE_NAMESPACE =
  'http://schemas.microsoft.com/office/word/2010/wordprocessingShape';

let editor: Editor | null = null;

afterEach(() => {
  editor?.destroy();
  editor = null;
});

describe('document text boxes', () => {
  test('normalizes bounded geometry, layout, colors, and positioning', () => {
    expect(
      normalizeDocumentTextBoxProperties({
        id: '  box-1  ',
        width: 5,
        height: 900,
        layout: 'floating',
        horizontalOffset: -900,
        verticalOffset: 900,
        horizontalReference: 'page',
        verticalReference: 'margin',
        fill: '#ABCDEF',
        borderColor: 'none',
        borderWidth: -2,
        padding: 50,
        verticalAlign: 'center',
        docPropertiesId: 7,
      }),
    ).toEqual({
      id: 'box-1',
      shapeType: 'rectangle',
      width: DOCUMENT_TEXT_BOX_LIMITS.width.min,
      height: DOCUMENT_TEXT_BOX_LIMITS.height.max,
      layout: 'floating',
      horizontalOffset: DOCUMENT_TEXT_BOX_LIMITS.offset.min,
      verticalOffset: DOCUMENT_TEXT_BOX_LIMITS.offset.max,
      horizontalReference: 'page',
      verticalReference: 'margin',
      fill: '#abcdef',
      borderColor: 'none',
      borderWidth: 0,
      padding: DOCUMENT_TEXT_BOX_LIMITS.padding.max,
      verticalAlign: 'center',
      docPropertiesId: 7,
    });
  });

  test('inserts, edits, serializes, and deletes one undoable text box', () => {
    editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: '<section data-document-section="true"><p>Before</p></section>',
    });
    editor.commands.setTextSelection(2);

    expect(editor.commands.insertDocumentTextBox('Editable text')).toBe(true);
    expect(editor.isActive('documentTextBox')).toBe(true);
    expect(editor.getHTML()).toContain('data-document-text-box="true"');
    expect(editor.getText()).toContain('Editable text');

    expect(
      editor.commands.setDocumentTextBoxProperties({
        width: 76.2,
        height: 25.4,
        layout: 'floating',
        horizontalOffset: 12.7,
        verticalOffset: -6.35,
        fill: '#ddebf7',
        borderColor: '#2f5597',
        borderWidth: 0.7,
        padding: 2.5,
        verticalAlign: 'bottom',
      }),
    ).toBe(true);
    const html = editor.getHTML();
    expect(html).toContain('data-text-box-width="76.2"');
    expect(html).toContain('data-text-box-height="25.4"');
    expect(html).toContain('data-text-box-layout="floating"');
    expect(html).toContain('data-text-box-horizontal-offset="12.7"');
    expect(html).toContain('data-text-box-vertical-offset="-6.35"');
    expect(html).toContain('data-text-box-vertical-align="bottom"');

    expect(editor.commands.undo()).toBe(true);
    expect(editor.getHTML()).not.toContain('data-text-box-layout="floating"');
    expect(editor.commands.redo()).toBe(true);
    expect(editor.getHTML()).toContain('data-text-box-layout="floating"');

    expect(editor.commands.deleteDocumentTextBox()).toBe(true);
    expect(editor.getHTML()).not.toContain('data-document-text-box');
    expect(editor.getText()).toContain('Before');
  });

  test('imports only isolated DrawingML text boxes and keeps their editable text', () => {
    const document = wordXml(`
      <w:p>
        <w:r><w:drawing>${textBoxDrawing('Imported text')}</w:drawing></w:r>
      </w:p>
      <w:p>
        <w:r><w:t>Prefix</w:t></w:r>
        <w:r><w:drawing>${textBoxDrawing('Unsupported mixed text')}</w:drawing></w:r>
      </w:p>
    `);

    const markers = markDocxTextBoxes(document);
    expect(markers.textBoxes).toHaveLength(1);
    expect(markers.textBoxes[0]?.properties).toMatchObject({
      id: 'imported-box',
      width: 76.2,
      height: 25.4,
      layout: 'floating',
      horizontalOffset: 12.7,
      verticalOffset: -6.35,
      horizontalReference: 'page',
      verticalReference: 'margin',
      fill: '#ddebf7',
      borderColor: '#2f5597',
      borderWidth: 0.7,
      padding: 2.5,
      verticalAlign: 'bottom',
      docPropertiesId: 12,
    });
    expect(document.documentElement.textContent).toContain('Imported text');
    expect(document.documentElement.textContent).toContain(
      'Unsupported mixed text',
    );
    expect(document.documentElement.textContent).toContain(
      markers.textBoxes[0]?.marker,
    );

    const html = new DOMParser().parseFromString(
      `<p>${markers.textBoxes[0]?.marker}<strong>Imported text</strong></p>`,
      'text/html',
    );
    applyImportedDocxTextBoxMarkers(html, markers);
    const box = html.body.querySelector<HTMLElement>(
      '[data-document-text-box]',
    );
    expect(box?.tagName).toBe('DIV');
    expect(box?.textContent).toBe('Imported text');
    expect(box?.dataset.textBoxWidth).toBe('76.2');
    expect(box?.dataset.textBoxVerticalOffset).toBe('-6.35');
    expect(box?.querySelector('strong')?.textContent).toBe('Imported text');
    expect(html.body.textContent).not.toContain('__A3S_WORK_TEXT_BOX_');
  });

  test('reports explicitly declared text boxes with malformed bodies', () => {
    const document = wordXml(`
      <w:p>
        <w:r><w:drawing>${malformedTextBoxDrawing()}</w:drawing></w:r>
      </w:p>
    `);

    expect(inspectDocxTextBoxes(document)).toEqual({
      supported: 0,
      unsupported: 1,
    });
    expect(markDocxTextBoxes(document).textBoxes).toHaveLength(0);
  });

  test('imports a WPS ellipse shape with text as the bounded shape subset', () => {
    const document = wordXml(`
      <w:p>
        <w:bookmarkStart w:id="0" w:name="_GoBack"/><w:bookmarkEnd w:id="0"/>
        <w:r>${wpsAlternateShapeDrawing('Ellipse text', 'ellipse')}</w:r>
      </w:p>
    `);

    const markers = markDocxTextBoxes(document);
    expect(markers.textBoxes).toHaveLength(1);
    expect(markers.textBoxes[0]?.properties).toMatchObject({
      shapeType: 'ellipse',
      width: 76.2,
      height: 25.4,
    });

    const html = new DOMParser().parseFromString(
      `<p>${markers.textBoxes[0]?.marker}Ellipse text</p>`,
      'text/html',
    );
    applyImportedDocxTextBoxMarkers(html, markers);
    expect(
      html.body.querySelector('[data-document-text-box]'),
    ).toHaveAttribute('data-text-box-shape', 'ellipse');
  });

  test.each([
    ['rect', 'rectangle'],
    ['roundRect', 'roundedRectangle'],
    ['ellipse', 'ellipse'],
    ['diamond', 'diamond'],
    ['triangle', 'triangle'],
  ] as const)('maps the supported WPS %s preset to %s', (preset, shapeType) => {
    const document = wordXml(`
      <w:p>
        <w:r><w:drawing>${shapeDrawing('Preset text', preset)}</w:drawing></w:r>
      </w:p>
    `);

    const markers = markDocxTextBoxes(document);
    expect(markers.textBoxes).toHaveLength(1);
    expect(markers.textBoxes[0]?.properties.shapeType).toBe(shapeType);
  });

  test('preserves imported text-box attributes in the structured document model', () => {
    const content = createWorkDocumentModelFromContent({
      type: 'document',
      html: `<section data-document-section="true"><div data-document-text-box="true" data-text-box-id="wps-box" data-text-box-shape="roundedRectangle" data-text-box-width="50.8" data-text-box-height="25.4" data-text-box-layout="floating" data-text-box-horizontal-offset="-6.35" data-text-box-vertical-offset="0" data-text-box-horizontal-reference="column" data-text-box-vertical-reference="paragraph" data-text-box-fill="#d9ead3" data-text-box-border-color="#4472c4" data-text-box-border-width="0.7" data-text-box-padding="1.27" data-text-box-vertical-align="center" data-text-box-doc-properties-id="1">WPS native shape</div></section>`,
      pageSize: 'a4',
    });
    const textBox = content.model?.root.content?.[0]?.content?.[0];
    expect(textBox).toMatchObject({
      type: 'documentTextBox',
      attrs: expect.objectContaining({
        id: 'wps-box',
        shapeType: 'roundedRectangle',
        width: '50.8',
        height: '25.4',
        layout: 'floating',
        fill: '#d9ead3',
        padding: '1.27',
        verticalAlign: 'center',
        docPropertiesId: '1',
      }),
    });
  });

  test('surfaces malformed text boxes in compatibility diagnostics', async () => {
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    const blob = await createDocxBlob(artifact.content);
    const archive = await JSZip.loadAsync(await blob.arrayBuffer());
    const source =
      (await archive.file('word/document.xml')?.async('text')) ?? '';
    archive.file(
      'word/document.xml',
      source.replace(
        '</w:body>',
        `<w:p><w:r><w:drawing>${malformedTextBoxDrawing()}</w:drawing></w:r></w:p></w:body>`,
      ),
    );
    const malformed = await archive.generateAsync({
      type: 'blob',
      mimeType: blob.type,
    });
    const compatibility = await analyzeDocxCompatibility(
      new File([malformed], 'malformed-text-box.docx', { type: blob.type }),
      [],
    );
    expect(compatibility.issues).toContainEqual(
      expect.objectContaining({
        code: 'docx.text-boxes.unsupported',
        severity: 'warning',
      }),
    );
  });

  test('exports native WPS shape geometry and reopens the editable box', async () => {
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    const properties = normalizeDocumentTextBoxProperties({
      id: 'round-trip-box',
      width: 76.2,
      height: 25.4,
      layout: 'floating',
      horizontalOffset: 12.7,
      verticalOffset: -6.35,
      horizontalReference: 'page',
      verticalReference: 'margin',
      fill: '#ddebf7',
      borderColor: '#2f5597',
      borderWidth: 0.7,
      padding: 2.5,
      verticalAlign: 'bottom',
      docPropertiesId: 12,
    });
    const attributes = Object.entries(textBoxDomAttributes(properties))
      .filter((entry): entry is [string, string] => entry[1] !== undefined)
      .map(([name, value]) => `${name}="${value}"`)
      .join(' ');
    artifact.content.html = `<p>Before</p><div ${attributes} style="${textBoxCss(properties)}"><strong>Round-trip text</strong></div><p>After</p>`;

    const blob = await createDocxBlob(artifact.content);
    const archive = await JSZip.loadAsync(await blob.arrayBuffer());
    const source =
      (await archive.file('word/document.xml')?.async('text')) ?? '';
    const document = parseXml(source);
    const shape = descendants(document, 'wsp')[0];
    if (!shape) throw new Error('Expected a native WPS shape.');
    expect(attribute(directChild(shape, 'cNvSpPr') ?? shape, 'txBox')).toBe(
      '1',
    );
    expect(source).toContain('<wp:extent cx="2743200" cy="914400"');
    expect(source).toContain('<wps:txbx>');
    expect(source).toContain('Round-trip text');
    expect(source).toContain('val="ddebf7"');
    expect(source).toContain('val="2f5597"');
    expect(source).toContain('anchor="b"');

    expect(
      descendants(document, 'prstGeom').map((geometry) =>
        attribute(geometry, 'prst'),
      ),
    ).toContain('rect');

    const reopened = await importOfficeFile(
      new File([blob], 'text-box.docx', { type: blob.type }),
    );
    if (reopened.content.type !== 'document') {
      throw new Error('Expected a reopened document artifact.');
    }
    expect(reopened.content.html).toContain('data-document-text-box="true"');
    expect(reopened.content.html).toContain(
      'data-text-box-id="round-trip-box"',
    );
    expect(reopened.content.html).toContain('data-text-box-width="76.2"');
    expect(reopened.content.html).toContain(
      'data-text-box-vertical-offset="-6.35"',
    );
    expect(reopened.content.html).toContain('Round-trip text');

    const reexported = await createDocxBlob(reopened.content);
    const reexportedArchive = await JSZip.loadAsync(
      await reexported.arrayBuffer(),
    );
    const reexportedSource =
      (await reexportedArchive.file('word/document.xml')?.async('text')) ?? '';
    expect(reexportedSource).not.toContain('__A3S_TEXT_BOX_ID_');
    expect(reexportedSource).toContain('name="A3S Text Box round-trip-box"');
    expect(reexportedSource).toMatch(
      /<wp:docPr id="12" name="A3S Text Box round-trip-box"/,
    );

    const compatibility = await analyzeDocxCompatibility(
      new File([blob], 'text-box.docx', { type: blob.type }),
      [],
    );
    expect(compatibility.issues).toContainEqual(
      expect.objectContaining({ code: 'docx.text-boxes', severity: 'info' }),
    );
    expect(compatibility.issues).not.toContainEqual(
      expect.objectContaining({ code: 'docx.images' }),
    );
  });

  test('exports the selected WPS preset geometry and reopens it', async () => {
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    const properties = normalizeDocumentTextBoxProperties({
      id: 'ellipse-box',
      shapeType: 'ellipse',
      width: 80,
      height: 40,
    });
    const attributes = Object.entries(textBoxDomAttributes(properties))
      .filter((entry): entry is [string, string] => entry[1] !== undefined)
      .map(([name, value]) => `${name}="${value}"`)
      .join(' ');
    artifact.content.html = `<div ${attributes} style="${textBoxCss(properties)}">Ellipse</div>`;
    const blob = await createDocxBlob(artifact.content);
    const archive = await JSZip.loadAsync(await blob.arrayBuffer());
    const source =
      (await archive.file('word/document.xml')?.async('text')) ?? '';
    expect(source).toContain('<a:prstGeom prst="ellipse"');
    const reopened = await importOfficeFile(
      new File([blob], 'ellipse-shape.docx', { type: blob.type }),
    );
    if (reopened.content.type !== 'document') {
      throw new Error('Expected a reopened document artifact.');
    }
    expect(reopened.content.html).toContain(
      'data-text-box-shape="ellipse"',
    );
  });

  test('assigns unique drawing-property IDs to repeated text boxes', async () => {
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    const sharedAttributes =
      'data-document-text-box="true" data-text-box-width="50" data-text-box-height="20" data-text-box-doc-properties-id="12"';
    artifact.content.html = `<div ${sharedAttributes}>First</div><div ${sharedAttributes}>Second</div>`;

    const blob = await createDocxBlob(artifact.content);
    const archive = await JSZip.loadAsync(await blob.arrayBuffer());
    const source =
      (await archive.file('word/document.xml')?.async('text')) ?? '';
    const document = parseXml(source);
    const docProperties = descendants(document, 'docPr');
    const ids = descendants(document, 'wsp')
      .map((shape) =>
        attribute(directChild(shape, 'cNvSpPr') ?? shape, 'txBox'),
      )
      .filter((value) => value === '1')
      .flatMap((_, index) => {
        const docProperty = docProperties[index];
        return docProperty ? [attribute(docProperty, 'id')] : [];
      });
    expect(ids).toHaveLength(2);
    expect(new Set(ids).size).toBe(2);
    expect(ids).toContain('12');
  });
});

function wordXml(body: string): Document {
  return parseXml(
    `<w:document xmlns:w="${WORD_NAMESPACE}" xmlns:wp="${WORDPROCESSING_DRAWING_NAMESPACE}" xmlns:a="${DRAWING_NAMESPACE}" xmlns:wps="${WORDPROCESSING_SHAPE_NAMESPACE}"><w:body>${body}</w:body></w:document>`,
  );
}

function textBoxDrawing(text: string): string {
  return `
    <wp:anchor>
      <wp:positionH relativeFrom="page"><wp:posOffset>457200</wp:posOffset></wp:positionH>
      <wp:positionV relativeFrom="margin"><wp:posOffset>-228600</wp:posOffset></wp:positionV>
      <wp:extent cx="2743200" cy="914400"/>
      <wp:docPr id="12" name="A3S Text Box imported-box"/>
      <a:graphic>
        <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
          <wps:wsp>
            <wps:cNvSpPr txBox="1"/>
            <wps:spPr>
              <a:xfrm><a:ext cx="2743200" cy="914400"/></a:xfrm>
              <a:solidFill><a:srgbClr val="DDEBF7"/></a:solidFill>
              <a:ln w="25200"><a:solidFill><a:srgbClr val="2F5597"/></a:solidFill></a:ln>
            </wps:spPr>
            <wps:txbx><w:txbxContent><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:txbxContent></wps:txbx>
            <wps:bodyPr lIns="90000" rIns="90000" tIns="90000" bIns="90000" anchor="b"/>
          </wps:wsp>
        </a:graphicData>
      </a:graphic>
    </wp:anchor>
  `;
}

function shapeDrawing(text: string, preset: string): string {
  return textBoxDrawing(text).replace(
    /<wps:cNvSpPr txBox="1"\s*\/>\s*<wps:spPr>/,
    `<wps:cNvSpPr/><wps:spPr><a:prstGeom prst="${preset}"><a:avLst/></a:prstGeom>`,
  );
}

function wpsAlternateShapeDrawing(text: string, preset: string): string {
  return `
    <mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
      <mc:Choice Requires="wps">
        <w:drawing><wp:anchor>
          <wp:extent cx="2743200" cy="914400"/>
          <wp:docPr id="22" name="WPS ${preset}"/>
          <a:graphic>
            <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
              <wps:wsp>
                <wps:cNvSpPr/>
                <wps:spPr>
                  <a:xfrm><a:ext cx="2743200" cy="914400"/></a:xfrm>
                  <a:prstGeom prst="${preset}"><a:avLst/></a:prstGeom>
                  <a:solidFill><a:srgbClr val="D9EAD3"/></a:solidFill>
                </wps:spPr>
                <wps:txbx><w:txbxContent><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:txbxContent></wps:txbx>
                <wps:bodyPr anchor="ctr"/>
              </wps:wsp>
            </a:graphicData>
          </a:graphic>
        </wp:anchor></w:drawing>
      </mc:Choice>
      <mc:Fallback><w:pict/></mc:Fallback>
    </mc:AlternateContent>
  `;
}

function malformedTextBoxDrawing(): string {
  return `
    <wp:inline xmlns:wp="${WORDPROCESSING_DRAWING_NAMESPACE}">
      <wp:extent cx="2743200" cy="914400"/>
      <wp:docPr id="13" name="Malformed text box"/>
      <a:graphic xmlns:a="${DRAWING_NAMESPACE}">
        <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
          <wps:wsp xmlns:wps="${WORDPROCESSING_SHAPE_NAMESPACE}"><wps:cNvSpPr txBox="1"/></wps:wsp>
        </a:graphicData>
      </a:graphic>
    </wp:inline>
  `;
}
