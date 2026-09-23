import { afterEach, describe, expect, test } from '@rstest/core';
import { Editor } from '@tiptap/core';
import JSZip from 'jszip';
import { createArtifact, importOfficeFile } from '../src/core';
import { createWorkDocumentExtensions } from '../src/internal/features/work/work-document-extensions';
import { createDocxBlob } from '../src/internal/features/work/work-docx-export';
import {
  applyImportedDocxRepeatingSectionMarkers,
  inspectDocxRepeatingSections,
  markDocxRepeatingSections,
} from '../src/internal/features/work/work-docx-repeating-section-import';
import { parseXml } from '../src/internal/features/work/work-ooxml-package';

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const WORD_2012_NAMESPACE =
  'http://schemas.microsoft.com/office/word/2012/wordml';

let editor: Editor | null = null;

afterEach(() => {
  editor?.destroy();
  editor = null;
});

describe('document repeating sections', () => {
  test('inserts a repeating section with one item', () => {
    editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: '<section data-document-section="true"><p></p></section>',
    });
    expect(
      editor.commands.insertDocumentRepeatingSection('第一项', {
        alias: '行项目',
        tag: 'lineItems',
      }),
    ).toBe(true);
    expect(editor.getHTML()).toContain(
      'data-document-repeating-section="true"',
    );
    expect(editor.getHTML()).toContain(
      'data-document-repeating-section-item="true"',
    );
    expect(editor.getHTML()).toContain('第一项');
  });

  test('marks and inspects a bounded body-level repeating section', () => {
    const document = parseXml(
      `<w:document xmlns:w="${WORD_NAMESPACE}" xmlns:w15="${WORD_2012_NAMESPACE}"><w:body><w:sdt><w:sdtPr><w:id w:val="11"/><w:alias w:val="Line items"/><w:tag w:val="lines"/><w15:repeatingSection/></w:sdtPr><w:sdtContent><w:sdt><w:sdtPr><w:id w:val="12"/><w15:repeatingSectionItem/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>Item A</w:t></w:r></w:p></w:sdtContent></w:sdt></w:sdtContent></w:sdt></w:body></w:document>`,
    );
    expect(inspectDocxRepeatingSections(document)).toEqual({
      supported: 1,
      unsupported: 0,
    });
    const markers = markDocxRepeatingSections(document);
    expect(markers.sections).toHaveLength(1);
    expect(markers.unsupported).toBe(0);
    expect(markers.sections[0]?.properties.alias).toBe('Line items');
    expect(markers.sections[0]?.items).toHaveLength(1);
    expect(document.documentElement.textContent).toContain('Item A');
    expect(document.documentElement.textContent).toContain(
      markers.sections[0]?.start,
    );

    const html = new DOMParser().parseFromString(
      `<p>${markers.sections[0]?.start}</p><p>${markers.sections[0]?.items[0]?.start}</p><p>Item A</p><p>${markers.sections[0]?.items[0]?.end}</p><p>${markers.sections[0]?.end}</p>`,
      'text/html',
    );
    applyImportedDocxRepeatingSectionMarkers(html, markers);
    const section = html.body.querySelector<HTMLElement>(
      '[data-document-repeating-section]',
    );
    const item = section?.querySelector<HTMLElement>(
      '[data-document-repeating-section-item]',
    );
    expect(section).toBeTruthy();
    expect(section?.getAttribute('data-repeating-section-alias')).toBe(
      'Line items',
    );
    expect(item).toBeTruthy();
    expect(item?.textContent ?? '').toContain('Item A');
    expect(html.body.textContent ?? '').not.toContain(
      '__A3S_WORK_REPEATING_SECTION_',
    );
  });

  test('exports a native w15:repeatingSection package', async () => {
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    artifact.content.html =
      '<section data-document-section="true"><div data-document-repeating-section="true" data-repeating-section-id="lines" data-repeating-section-native-id="11" data-repeating-section-alias="Line items" data-repeating-section-tag="lines" data-repeating-section-lock="unlocked" data-repeating-section-appearance="boundingBox"><div data-document-repeating-section-item="true" data-repeating-section-item-id="item-1" data-repeating-section-item-native-id="12"><p>Item A</p></div></div></section>';
    const blob = await createDocxBlob(artifact.content);
    const archive = await JSZip.loadAsync(await blob.arrayBuffer());
    const xml = (await archive.file('word/document.xml')?.async('text')) ?? '';
    expect(xml).toContain('repeatingSection');
    expect(xml).toContain('repeatingSectionItem');
    expect(xml).toContain('Item A');
    expect(xml).toContain('Line items');
    expect(xml).not.toContain('__A3S_WORK_REPEATING_SECTION_');

    const exported = parseXml(xml, 'exported repeating section');
    expect(inspectDocxRepeatingSections(exported)).toEqual({
      supported: 1,
      unsupported: 0,
    });

    const reopened = await importOfficeFile(
      new File([blob], 'repeating-section.docx', { type: blob.type }),
    );
    if (reopened.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    expect(reopened.content.html).toContain(
      'data-document-repeating-section="true"',
    );
    expect(reopened.content.html).toContain('Item A');
  });
});
