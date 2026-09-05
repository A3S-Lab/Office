import { afterEach, describe, expect, test } from '@rstest/core';
import { Editor } from '@tiptap/core';
import JSZip from 'jszip';
import { createArtifact, importOfficeFile } from '../src/core';
import {
  connectorCss,
  connectorDomAttributes,
  DOCUMENT_CONNECTOR_LIMITS,
  normalizeDocumentConnectorProperties,
} from '../src/internal/features/work/work-document-connector';
import { createWorkDocumentExtensions } from '../src/internal/features/work/work-document-extensions';
import {
  applyImportedDocxConnectorMarkers,
  inspectDocxConnectorShapes,
  markDocxConnectors,
} from '../src/internal/features/work/work-docx-connector-import';
import { createDocxBlob } from '../src/internal/features/work/work-docx-export';
import { parseXml } from '../src/internal/features/work/work-ooxml-package';

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const VML_NAMESPACE = 'urn:schemas-microsoft-com:vml';
const OFFICE_NAMESPACE = 'urn:schemas-microsoft-com:office:office';
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

describe('document connectors', () => {
  test('normalizes endpoints, geometry, color, and arrow state', () => {
    expect(
      normalizeDocumentConnectorProperties({
        id: ' connector-1 ',
        width: 1,
        height: 900,
        startX: -5,
        startY: 120,
        endX: 35,
        endY: 65,
        lineColor: '#ABCDEF',
        lineWidth: 50,
        lineStyle: 'dashDot',
        startArrow: 'open',
        endArrow: 'diamond',
      }),
    ).toEqual({
      id: 'connector-1',
      connectorKind: 'straight',
      width: DOCUMENT_CONNECTOR_LIMITS.width.min,
      height: DOCUMENT_CONNECTOR_LIMITS.height.max,
      layout: 'inline',
      horizontalOffset: null,
      verticalOffset: null,
      horizontalReference: 'column',
      verticalReference: 'paragraph',
      startX: 0,
      startY: 100,
      endX: 35,
      endY: 65,
      lineColor: '#abcdef',
      lineWidth: DOCUMENT_CONNECTOR_LIMITS.lineWidth.max,
      lineStyle: 'dashDot',
      startArrow: 'open',
      endArrow: 'diamond',
      docPropertiesId: null,
    });
  });

  test('inserts, edits, serializes, and deletes one connector', () => {
    editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: '<section data-document-section="true"><p>Before</p></section>',
    });
    editor.commands.setTextSelection(2);

    expect(
      editor.commands.insertDocumentConnector({
        endY: 0,
        endArrow: 'triangle',
      }),
    ).toBe(true);
    expect(editor.isActive('documentConnector')).toBe(true);
    expect(editor.getHTML()).toContain('data-document-connector="true"');
    expect(editor.getHTML()).toContain('data-connector-end-arrow="triangle"');

    expect(
      editor.commands.setDocumentConnectorProperties({
        width: 76.2,
        height: 25.4,
        layout: 'floating',
        horizontalOffset: 12.7,
        verticalOffset: -6.35,
        lineColor: '#2f5597',
        lineWidth: 0.7,
        startX: 15,
        startY: 20,
        endX: 90,
        endY: 80,
      }),
    ).toBe(true);
    expect(editor.getHTML()).toContain('data-connector-layout="floating"');
    expect(editor.getHTML()).toContain('data-connector-start-x="15"');
    expect(editor.getHTML()).toContain('data-connector-line-color="#2f5597"');

    expect(editor.commands.undo()).toBe(true);
    expect(editor.getHTML()).not.toContain('data-connector-layout="floating"');
    expect(editor.commands.redo()).toBe(true);
    expect(editor.getHTML()).toContain('data-connector-layout="floating"');

    expect(editor.commands.deleteDocumentConnector()).toBe(true);
    expect(editor.getHTML()).not.toContain('data-document-connector');
    expect(editor.getText()).toContain('Before');
  });

  test('imports an isolated WPS VML connector without fabricating a text box', () => {
    const document = wordXml(`
      <w:p>
        <w:bookmarkStart w:id="0" w:name="_GoBack"/><w:bookmarkEnd w:id="0"/>
        <w:r><w:pict>${vmlConnector()}</w:pict></w:r>
      </w:p>
    `);
    expect(inspectDocxConnectorShapes(document)).toEqual({
      supported: 1,
      unsupported: 0,
    });
    const markers = markDocxConnectors(document);
    expect(markers.connectors).toHaveLength(1);
    expect(markers.connectors[0]?.properties).toMatchObject({
      connectorKind: 'straight',
      width: 50.8,
      layout: 'floating',
      horizontalOffset: -6.35,
      lineColor: '#c00000',
      lineStyle: 'dash',
      startArrow: 'open',
      endArrow: 'stealth',
      endY: 50,
    });
    expect(markers.connectors[0]?.properties.verticalOffset).toBeCloseTo(
      45.1556,
      3,
    );
    const html = new DOMParser().parseFromString(
      `<p>${markers.connectors[0]?.marker}</p>`,
      'text/html',
    );
    applyImportedDocxConnectorMarkers(html, markers);
    const connector = html.body.querySelector('[data-document-connector]');
    expect(connector).toHaveAttribute('data-connector-width', '50.8');
    expect(connector).toHaveAttribute('data-connector-kind', 'straight');
    expect(connector).toHaveAttribute('data-connector-layout', 'floating');
    expect(connector).toHaveAttribute('data-connector-start-arrow', 'open');
    expect(connector).toHaveAttribute('data-connector-end-arrow', 'stealth');
    expect(html.body.querySelector('[data-document-text-box]')).toBeNull();
  });

  test('maps WPS VML shape types to typed connector kinds', () => {
    const document = wordXml(
      [32, 33, 37]
        .map(
          (shapeType) =>
            `<w:p><w:r><w:pict>${vmlConnector(shapeType)}</w:pict></w:r></w:p>`,
        )
        .join(''),
    );
    const markers = markDocxConnectors(document);
    expect(
      markers.connectors.map(({ properties }) => properties.connectorKind),
    ).toEqual(['straight', 'elbow', 'curved']);
  });

  test('exports a DrawingML connector and reopens its native properties', async () => {
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    const properties = normalizeDocumentConnectorProperties({
      id: 'round-trip-connector',
      width: 76.2,
      height: 25.4,
      layout: 'floating',
      horizontalOffset: 12.7,
      verticalOffset: -6.35,
      horizontalReference: 'page',
      verticalReference: 'margin',
      startX: 15,
      startY: 20,
      endX: 90,
      endY: 80,
      connectorKind: 'elbow',
      lineColor: '#2f5597',
      lineWidth: 0.7,
      lineStyle: 'dash',
      startArrow: 'open',
      endArrow: 'diamond',
      docPropertiesId: 12,
    });
    const attributes = Object.entries(connectorDomAttributes(properties))
      .filter((entry): entry is [string, string] => entry[1] !== undefined)
      .map(([name, value]) => `${name}="${value}"`)
      .join(' ');
    artifact.content.html = `<p>Before</p><div ${attributes} style="${connectorCss(properties)}"></div><p>After</p>`;

    const blob = await createDocxBlob(artifact.content);
    const archive = await JSZip.loadAsync(await blob.arrayBuffer());
    const source =
      (await archive.file('word/document.xml')?.async('text')) ?? '';
    expect(source).toContain('custGeom');
    expect(source).toContain('<a:lnTo>');
    expect(source).toContain('headEnd');
    expect(source).toContain('tailEnd');
    expect(source).toContain('<a:headEnd type="open"');
    expect(source).toContain('<a:tailEnd type="diamond"');
    expect(source).toContain('<a:prstDash val="dash"');
    expect(source).toContain('val="2F5597"');
    expect(source).not.toContain('__A3S_CONNECTOR_ID_');

    const reopened = await importOfficeFile(
      new File([blob], 'connector.docx', { type: blob.type }),
    );
    if (reopened.content.type !== 'document') {
      throw new Error('Expected a reopened document artifact.');
    }
    expect(reopened.content.html).toContain('data-document-connector="true"');
    expect(reopened.content.html).toContain('data-connector-kind="elbow"');
    expect(reopened.content.html).toContain(
      'data-connector-line-color="#2f5597"',
    );
    expect(reopened.content.html).toContain(
      'data-connector-start-arrow="open"',
    );
    expect(reopened.content.html).toContain(
      'data-connector-end-arrow="diamond"',
    );
    expect(reopened.content.html).toContain('data-connector-line-style="dash"');
    expect(reopened.content.html).not.toContain('data-document-text-box');
  });

  test('exports curved connector geometry as a native quadratic path', async () => {
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    const properties = normalizeDocumentConnectorProperties({
      id: 'curved-connector',
      connectorKind: 'curved',
      startX: 10,
      startY: 20,
      endX: 90,
      endY: 80,
    });
    const attributes = Object.entries(connectorDomAttributes(properties))
      .filter((entry): entry is [string, string] => entry[1] !== undefined)
      .map(([name, value]) => `${name}="${value}"`)
      .join(' ');
    artifact.content.html = `<p>Before</p><div ${attributes} style="${connectorCss(properties)}"></div><p>After</p>`;
    const blob = await createDocxBlob(artifact.content);
    const archive = await JSZip.loadAsync(await blob.arrayBuffer());
    const source =
      (await archive.file('word/document.xml')?.async('text')) ?? '';
    expect(source).toContain('<a:quadBezTo>');
    const reopened = await importOfficeFile(
      new File([blob], 'curved-connector.docx', { type: blob.type }),
    );
    if (reopened.content.type !== 'document') {
      throw new Error('Expected a reopened document artifact.');
    }
    expect(reopened.content.html).toContain('data-connector-kind="curved"');
  });
});

function wordXml(body: string): Document {
  return parseXml(
    `<w:document xmlns:w="${WORD_NAMESPACE}" xmlns:v="${VML_NAMESPACE}" xmlns:o="${OFFICE_NAMESPACE}" xmlns:wp="${WORDPROCESSING_DRAWING_NAMESPACE}" xmlns:a="${DRAWING_NAMESPACE}" xmlns:wps="${WORDPROCESSING_SHAPE_NAMESPACE}"><w:body>${body}</w:body></w:document>`,
  );
}

function vmlConnector(shapeType = 32): string {
  return `
    <v:shape id="A3S Connector" o:spid="_x0000_s1026" o:spt="${shapeType}" type="#_x0000_t${shapeType}" style="position:absolute;left:0pt;margin-left:-18pt;margin-top:128pt;height:1pt;width:144pt" filled="f" stroked="f" coordsize="21600,21600">
      <v:path arrowok="t"/>
      <v:stroke on="f" color="#C00000" weight="1pt" dashstyle="dash" startarrow="open" endarrow="classic"/>
    </v:shape>
  `;
}
