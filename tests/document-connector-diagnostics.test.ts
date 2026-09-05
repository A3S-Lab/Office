import { describe, expect, test } from '@rstest/core';
import JSZip from 'jszip';
import { createArtifact } from '../src/core';
import { inspectDocxConnectors } from '../src/internal/features/work/work-docx-connector-diagnostics';
import { createDocxBlob } from '../src/internal/features/work/work-docx-export';
import { analyzeDocxCompatibility } from '../src/internal/features/work/work-office-diagnostics';
import { parseXml } from '../src/internal/features/work/work-ooxml-package';

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const VML_NAMESPACE = 'urn:schemas-microsoft-com:vml';
const OFFICE_NAMESPACE = 'urn:schemas-microsoft-com:office:office';

describe('WPS connector compatibility boundary', () => {
  test('counts VML connectors without treating ordinary VML shapes as connectors', () => {
    const document = parseXml(`
      <w:document xmlns:w="${WORD_NAMESPACE}"
        xmlns:v="${VML_NAMESPACE}"
        xmlns:o="${OFFICE_NAMESPACE}">
        <w:body>
          <w:p><w:r><w:pict>
            <v:shape o:spt="32" type="#_x0000_t32"/>
            <v:shape o:spt="1" type="#_x0000_t1"/>
          </w:pict></w:r></w:p>
        </w:body>
      </w:document>
    `);

    expect(inspectDocxConnectors(document)).toEqual({
      detected: 1,
      connectorPictContainers: 0,
    });
  });

  test('emits a dedicated connector diagnostic and avoids the generic image warning', async () => {
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
        `<w:p><w:r><w:pict xmlns:v="${VML_NAMESPACE}" xmlns:o="${OFFICE_NAMESPACE}"><v:shape o:spt="32" type="#_x0000_t32" style="position:absolute;left:0pt;top:0pt;width:144pt;height:1pt"/></w:pict></w:r></w:p></w:body>`,
      ),
    );
    const connectorDocx = await archive.generateAsync({
      type: 'blob',
      mimeType: blob.type,
    });

    const compatibility = await analyzeDocxCompatibility(
      new File([connectorDocx], 'wps-connector.docx', { type: blob.type }),
      [],
    );
    expect(compatibility.issues).toContainEqual(
      expect.objectContaining({
        code: 'docx.connectors',
        severity: 'warning',
      }),
    );
    expect(compatibility.issues).not.toContainEqual(
      expect.objectContaining({ code: 'docx.images' }),
    );
  });
});
