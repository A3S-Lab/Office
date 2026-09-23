import { afterEach, describe, expect, test } from '@rstest/core';
import { Editor } from '@tiptap/core';
import JSZip from 'jszip';
import { createArtifact, importOfficeFile } from '../src/core';
import { createWorkDocumentExtensions } from '../src/internal/features/work/work-document-extensions';
import {
  applyImportedDocxBlockContentControlMarkers,
  inspectDocxBlockContentControls,
  markDocxBlockContentControls,
} from '../src/internal/features/work/work-docx-block-content-control-import';
import { inspectDocxContentControls } from '../src/internal/features/work/work-docx-content-control-import';
import { createDocxBlob } from '../src/internal/features/work/work-docx-export';
import { parseXml } from '../src/internal/features/work/work-ooxml-package';

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

let editor: Editor | null = null;

afterEach(() => {
  editor?.destroy();
  editor = null;
});

describe('document block content controls', () => {
  test('inserts a body-level block content control', () => {
    editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: '<section data-document-section="true"><p></p></section>',
    });
    expect(
      editor.commands.insertDocumentBlockContentControl('条款正文', {
        alias: 'Clause',
        tag: 'clause',
        type: 'richText',
      }),
    ).toBe(true);
    expect(editor.getHTML()).toContain(
      'data-document-block-content-control="true"',
    );
    expect(editor.getHTML()).toContain('条款正文');
  });

  test('marks and inspects a bounded body-level rich-text control', () => {
    const document = parseXml(
      `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:sdt><w:sdtPr><w:id w:val="21"/><w:alias w:val="Clause"/><w:tag w:val="clause"/><w:richText/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>Block body</w:t></w:r></w:p></w:sdtContent></w:sdt></w:body></w:document>`,
    );
    expect(inspectDocxBlockContentControls(document)).toEqual({
      supported: 1,
      unsupported: 0,
    });
    expect(inspectDocxContentControls(document)).toEqual({
      supported: 0,
      unsupported: 0,
    });
    const markers = markDocxBlockContentControls(document);
    expect(markers.controls).toHaveLength(1);
    expect(markers.controls[0]?.properties).toMatchObject({
      type: 'richText',
      alias: 'Clause',
      tag: 'clause',
    });
    expect(markers.controls[0]?.text).toBe('Block body');

    const html = new DOMParser().parseFromString(
      `<p>${markers.controls[0]?.start}</p><p>Block body</p><p>${markers.controls[0]?.end}</p>`,
      'text/html',
    );
    applyImportedDocxBlockContentControlMarkers(html, markers);
    const control = html.body.querySelector<HTMLElement>(
      '[data-document-block-content-control]',
    );
    expect(control).toBeTruthy();
    expect(control?.getAttribute('data-block-content-control-alias')).toBe(
      'Clause',
    );
    expect(control?.textContent ?? '').toContain('Block body');
    expect(html.body.textContent ?? '').not.toContain(
      '__A3S_WORK_BLOCK_CONTENT_CONTROL_',
    );
  });

  test('exports a native body-level w:sdt package', async () => {
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    artifact.content.html =
      '<section data-document-section="true"><div data-document-block-content-control="true" data-block-content-control-id="clause" data-block-content-control-native-id="21" data-block-content-control-type="richText" data-block-content-control-alias="Clause" data-block-content-control-tag="clause" data-block-content-control-lock="unlocked" data-block-content-control-appearance="boundingBox"><p>Block body</p></div></section>';
    const blob = await createDocxBlob(artifact.content);
    const archive = await JSZip.loadAsync(await blob.arrayBuffer());
    const xml = (await archive.file('word/document.xml')?.async('text')) ?? '';
    expect(xml).toContain('<w:sdt>');
    expect(xml).toContain('<w:richText');
    expect(xml).toContain('Clause');
    expect(xml).toContain('Block body');
    expect(xml).not.toContain('__A3S_WORK_BLOCK_CONTENT_CONTROL_');

    const exported = parseXml(xml, 'exported block content control');
    expect(inspectDocxBlockContentControls(exported)).toEqual({
      supported: 1,
      unsupported: 0,
    });

    const reopened = await importOfficeFile(
      new File([blob], 'block-content-control.docx', { type: blob.type }),
    );
    if (reopened.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    expect(reopened.content.html).toContain(
      'data-document-block-content-control="true"',
    );
    expect(reopened.content.html).toContain('Block body');
  });

  test('round-trips a multi-paragraph body-level rich-text control', async () => {
    const document = parseXml(
      `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:sdt><w:sdtPr><w:id w:val="22"/><w:alias w:val="Clause"/><w:tag w:val="clause"/><w:richText/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>First paragraph</w:t></w:r></w:p><w:p><w:r><w:t>Second paragraph</w:t></w:r></w:p></w:sdtContent></w:sdt></w:body></w:document>`,
    );
    expect(inspectDocxBlockContentControls(document)).toEqual({
      supported: 1,
      unsupported: 0,
    });
    const markers = markDocxBlockContentControls(document);
    expect(markers.controls).toHaveLength(1);
    expect(markers.controls[0]?.text).toBe(
      'First paragraph\nSecond paragraph',
    );

    const html = new DOMParser().parseFromString(
      `<p>${markers.controls[0]?.start}</p><p>First paragraph</p><p>Second paragraph</p><p>${markers.controls[0]?.end}</p>`,
      'text/html',
    );
    applyImportedDocxBlockContentControlMarkers(html, markers);
    const control = html.body.querySelector<HTMLElement>(
      '[data-document-block-content-control]',
    );
    expect(control?.querySelectorAll('p')).toHaveLength(2);
    expect(control?.textContent ?? '').toContain('First paragraph');
    expect(control?.textContent ?? '').toContain('Second paragraph');

    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    artifact.content.html =
      '<section data-document-section="true"><div data-document-block-content-control="true" data-block-content-control-id="clause" data-block-content-control-native-id="22" data-block-content-control-type="richText" data-block-content-control-alias="Clause" data-block-content-control-tag="clause" data-block-content-control-lock="unlocked" data-block-content-control-appearance="boundingBox"><p>First paragraph</p><p>Second paragraph</p></div></section>';
    const blob = await createDocxBlob(artifact.content);
    const archive = await JSZip.loadAsync(await blob.arrayBuffer());
    const xml = (await archive.file('word/document.xml')?.async('text')) ?? '';
    expect(xml).toContain('First paragraph');
    expect(xml).toContain('Second paragraph');
    expect(xml.match(/<w:p>/g)?.length ?? 0).toBeGreaterThanOrEqual(2);

    const reopened = await importOfficeFile(
      new File([blob], 'multi-paragraph-block-content-control.docx', {
        type: blob.type,
      }),
    );
    if (reopened.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    expect(reopened.content.html).toContain('First paragraph');
    expect(reopened.content.html).toContain('Second paragraph');
  });

  test('keeps depth-3 nested block content controls fail-closed', () => {
    const document = parseXml(
      `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:sdt><w:sdtPr><w:id w:val="30"/><w:richText/></w:sdtPr><w:sdtContent><w:sdt><w:sdtPr><w:id w:val="31"/><w:richText/></w:sdtPr><w:sdtContent><w:sdt><w:sdtPr><w:id w:val="32"/><w:richText/></w:sdtPr><w:sdtContent><w:sdt><w:sdtPr><w:id w:val="33"/><w:richText/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>Too deep</w:t></w:r></w:p></w:sdtContent></w:sdt></w:sdtContent></w:sdt></w:sdtContent></w:sdt></w:sdtContent></w:sdt></w:body></w:document>`,
    );
    expect(inspectDocxBlockContentControls(document)).toEqual({
      supported: 0,
      unsupported: 1,
    });
  });

  test('round-trips two-level nested body-level rich-text controls', async () => {
    const document = parseXml(
      `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:sdt><w:sdtPr><w:id w:val="50"/><w:alias w:val="Chapter"/><w:tag w:val="chapter"/><w:richText/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>Chapter intro</w:t></w:r></w:p><w:sdt><w:sdtPr><w:id w:val="51"/><w:alias w:val="Section"/><w:tag w:val="section"/><w:richText/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>Section intro</w:t></w:r></w:p><w:sdt><w:sdtPr><w:id w:val="52"/><w:alias w:val="Clause"/><w:tag w:val="clause"/><w:richText/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>Leaf body</w:t></w:r></w:p></w:sdtContent></w:sdt><w:p><w:r><w:t>Section outro</w:t></w:r></w:p></w:sdtContent></w:sdt><w:p><w:r><w:t>Chapter outro</w:t></w:r></w:p></w:sdtContent></w:sdt></w:body></w:document>`,
    );
    expect(inspectDocxBlockContentControls(document)).toEqual({
      supported: 1,
      unsupported: 0,
    });
    const markers = markDocxBlockContentControls(document);
    expect(markers.controls).toHaveLength(3);
    expect(markers.controls.map((entry) => entry.properties.alias)).toEqual([
      'Clause',
      'Section',
      'Chapter',
    ]);

    const html = new DOMParser().parseFromString(
      [
        `<p>${markers.controls[2]?.start}</p>`,
        '<p>Chapter intro</p>',
        `<p>${markers.controls[1]?.start}</p>`,
        '<p>Section intro</p>',
        `<p>${markers.controls[0]?.start}</p>`,
        '<p>Leaf body</p>',
        `<p>${markers.controls[0]?.end}</p>`,
        '<p>Section outro</p>',
        `<p>${markers.controls[1]?.end}</p>`,
        '<p>Chapter outro</p>',
        `<p>${markers.controls[2]?.end}</p>`,
      ].join(''),
      'text/html',
    );
    applyImportedDocxBlockContentControlMarkers(html, markers);
    const chapter = html.body.querySelector<HTMLElement>(
      '[data-document-block-content-control][data-block-content-control-alias="Chapter"]',
    );
    const section = chapter?.querySelector<HTMLElement>(
      '[data-document-block-content-control][data-block-content-control-alias="Section"]',
    );
    const clause = section?.querySelector<HTMLElement>(
      '[data-document-block-content-control][data-block-content-control-alias="Clause"]',
    );
    expect(chapter).toBeTruthy();
    expect(section).toBeTruthy();
    expect(clause).toBeTruthy();
    expect(clause?.textContent ?? '').toContain('Leaf body');

    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    artifact.content.html =
      '<section data-document-section="true"><div data-document-block-content-control="true" data-block-content-control-id="chapter" data-block-content-control-native-id="50" data-block-content-control-type="richText" data-block-content-control-alias="Chapter" data-block-content-control-tag="chapter" data-block-content-control-lock="unlocked" data-block-content-control-appearance="boundingBox"><p>Chapter intro</p><div data-document-block-content-control="true" data-block-content-control-id="section" data-block-content-control-native-id="51" data-block-content-control-type="richText" data-block-content-control-alias="Section" data-block-content-control-tag="section" data-block-content-control-lock="unlocked" data-block-content-control-appearance="boundingBox"><p>Section intro</p><div data-document-block-content-control="true" data-block-content-control-id="clause" data-block-content-control-native-id="52" data-block-content-control-type="richText" data-block-content-control-alias="Clause" data-block-content-control-tag="clause" data-block-content-control-lock="unlocked" data-block-content-control-appearance="boundingBox"><p>Leaf body</p></div><p>Section outro</p></div><p>Chapter outro</p></div></section>';
    const blob = await createDocxBlob(artifact.content);
    const archive = await JSZip.loadAsync(await blob.arrayBuffer());
    const xml = (await archive.file('word/document.xml')?.async('text')) ?? '';
    expect(xml).toContain('Leaf body');
    expect((xml.match(/<w:sdt>/g) ?? []).length).toBeGreaterThanOrEqual(3);

    const reopened = await importOfficeFile(
      new File([blob], 'two-level-nested-block-content-control.docx', {
        type: blob.type,
      }),
    );
    if (reopened.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    expect(reopened.content.html).toContain('Leaf body');
    expect(reopened.content.html).toContain(
      'data-block-content-control-alias="Chapter"',
    );
    expect(reopened.content.html).toContain(
      'data-block-content-control-alias="Section"',
    );
    expect(reopened.content.html).toContain(
      'data-block-content-control-alias="Clause"',
    );
  });

  test('inserts a second nesting level inside a nested block control', () => {
    editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: '<section data-document-section="true"><p></p></section>',
    });
    expect(
      editor.commands.insertDocumentBlockContentControl('Outer', {
        alias: 'Outer',
        type: 'richText',
      }),
    ).toBe(true);
    const outer = editor.view.dom.querySelector(
      '[data-document-block-content-control]',
    );
    expect(outer).toBeTruthy();
    const paragraph = outer?.querySelector('p');
    if (!paragraph) throw new Error('Expected outer paragraph.');
    const pos = editor.view.posAtDOM(paragraph, 0);
    editor.commands.setTextSelection(pos + 1);
    expect(
      editor.commands.insertDocumentBlockContentControl('Mid', {
        alias: 'Mid',
        type: 'richText',
      }),
    ).toBe(true);
    const mid = editor.view.dom.querySelector(
      '[data-block-content-control-alias="Mid"]',
    );
    expect(mid).toBeTruthy();
    const midParagraph = mid?.querySelector('p');
    if (!midParagraph) throw new Error('Expected mid paragraph.');
    const midPos = editor.view.posAtDOM(midParagraph, 0);
    editor.commands.setTextSelection(midPos + 1);
    expect(
      editor.commands.insertDocumentBlockContentControl('Leaf', {
        alias: 'Leaf',
        type: 'richText',
      }),
    ).toBe(true);
    expect(
      editor.view.dom.querySelector(
        '[data-block-content-control-alias="Outer"] [data-block-content-control-alias="Mid"] [data-block-content-control-alias="Leaf"]',
      ),
    ).toBeTruthy();
    const leafParagraph = editor.view.dom.querySelector(
      '[data-block-content-control-alias="Leaf"] p',
    );
    if (!leafParagraph) throw new Error('Expected leaf paragraph.');
    const leafPos = editor.view.posAtDOM(leafParagraph, 0);
    editor.commands.setTextSelection(leafPos + 1);
    expect(
      editor.commands.insertDocumentBlockContentControl('Too deep', {
        alias: 'TooDeep',
        type: 'richText',
      }),
    ).toBe(false);
  });

  test('round-trips one-level nested body-level rich-text controls', async () => {
    const document = parseXml(
      `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:sdt><w:sdtPr><w:id w:val="40"/><w:alias w:val="Section"/><w:tag w:val="section"/><w:richText/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>Intro</w:t></w:r></w:p><w:sdt><w:sdtPr><w:id w:val="41"/><w:alias w:val="Clause"/><w:tag w:val="clause"/><w:richText/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>Nested body</w:t></w:r></w:p></w:sdtContent></w:sdt><w:p><w:r><w:t>Outro</w:t></w:r></w:p></w:sdtContent></w:sdt></w:body></w:document>`,
    );
    expect(inspectDocxBlockContentControls(document)).toEqual({
      supported: 1,
      unsupported: 0,
    });
    const markers = markDocxBlockContentControls(document);
    expect(markers.controls).toHaveLength(2);
    expect(markers.controls.map((entry) => entry.properties.alias)).toEqual([
      'Clause',
      'Section',
    ]);

    const html = new DOMParser().parseFromString(
      `<p>${markers.controls[1]?.start}</p><p>Intro</p><p>${markers.controls[0]?.start}</p><p>Nested body</p><p>${markers.controls[0]?.end}</p><p>Outro</p><p>${markers.controls[1]?.end}</p>`,
      'text/html',
    );
    applyImportedDocxBlockContentControlMarkers(html, markers);
    const outer = html.body.querySelector<HTMLElement>(
      '[data-document-block-content-control][data-block-content-control-alias="Section"]',
    );
    const inner = outer?.querySelector<HTMLElement>(
      '[data-document-block-content-control][data-block-content-control-alias="Clause"]',
    );
    expect(outer).toBeTruthy();
    expect(inner).toBeTruthy();
    expect(inner?.textContent ?? '').toContain('Nested body');
    expect(outer?.textContent ?? '').toContain('Intro');
    expect(outer?.textContent ?? '').toContain('Outro');

    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    artifact.content.html =
      '<section data-document-section="true"><div data-document-block-content-control="true" data-block-content-control-id="section" data-block-content-control-native-id="40" data-block-content-control-type="richText" data-block-content-control-alias="Section" data-block-content-control-tag="section" data-block-content-control-lock="unlocked" data-block-content-control-appearance="boundingBox"><p>Intro</p><div data-document-block-content-control="true" data-block-content-control-id="clause" data-block-content-control-native-id="41" data-block-content-control-type="richText" data-block-content-control-alias="Clause" data-block-content-control-tag="clause" data-block-content-control-lock="unlocked" data-block-content-control-appearance="boundingBox"><p>Nested body</p></div><p>Outro</p></div></section>';
    const blob = await createDocxBlob(artifact.content);
    const archive = await JSZip.loadAsync(await blob.arrayBuffer());
    const xml = (await archive.file('word/document.xml')?.async('text')) ?? '';
    expect(xml).toContain('Intro');
    expect(xml).toContain('Nested body');
    expect(xml).toContain('Outro');
    expect((xml.match(/<w:sdt>/g) ?? []).length).toBeGreaterThanOrEqual(2);

    const reopened = await importOfficeFile(
      new File([blob], 'nested-block-content-control.docx', {
        type: blob.type,
      }),
    );
    if (reopened.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    expect(reopened.content.html).toContain('Nested body');
    expect(reopened.content.html).toContain(
      'data-block-content-control-alias="Section"',
    );
    expect(reopened.content.html).toContain(
      'data-block-content-control-alias="Clause"',
    );
  });

  test('inserts multi-line block content as multiple paragraphs', () => {
    editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: '<section data-document-section="true"><p></p></section>',
    });
    expect(
      editor.commands.insertDocumentBlockContentControl(
        'First line\nSecond line',
        { alias: 'Clause', tag: 'clause', type: 'richText' },
      ),
    ).toBe(true);
    const control = editor.view.dom.querySelector(
      '[data-document-block-content-control]',
    );
    expect(control?.querySelectorAll('p')).toHaveLength(2);
    expect(control?.textContent ?? '').toContain('First line');
    expect(control?.textContent ?? '').toContain('Second line');
  });
});
