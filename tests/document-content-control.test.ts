import { afterEach, describe, expect, test } from '@rstest/core';
import { Editor } from '@tiptap/core';
import JSZip from 'jszip';
import { createArtifact, importOfficeFile } from '../src/core';
import {
  contentControlDomAttributes,
  DOCUMENT_CONTENT_CONTROL_DEFAULTS,
  normalizeDocumentContentControlProperties,
} from '../src/internal/features/work/work-document-content-control';
import {
  applyImportedDocxContentControlMarkers,
  inspectDocxContentControls,
  markDocxContentControls,
} from '../src/internal/features/work/work-docx-content-control-import';
import { createDocxBlob } from '../src/internal/features/work/work-docx-export';
import { analyzeDocxCompatibility } from '../src/internal/features/work/work-office-diagnostics';
import { parseXml } from '../src/internal/features/work/work-ooxml-package';

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const WORD_2012_NAMESPACE =
  'http://schemas.microsoft.com/office/word/2012/wordml';

import { createWorkDocumentExtensions } from '../src/internal/features/work/work-document-extensions';

let editor: Editor | null = null;

afterEach(() => {
  editor?.destroy();
  editor = null;
});

describe('document content controls', () => {
  test('normalizes bounded metadata and projects accessible HTML attributes', () => {
    expect(
      normalizeDocumentContentControlProperties({
        id: '  control-1\u0000 ',
        nativeId: '12',
        type: 'richText',
        alias: '  Customer name  ',
        tag: 'customer-name',
        lock: 'sdtContentLocked',
        multiLine: 'true',
        appearance: 'tags',
        color: 'DDEBF7',
      }),
    ).toEqual({
      id: 'control-1',
      nativeId: 12,
      type: 'richText',
      alias: 'Customer name',
      tag: 'customer-name',
      lock: 'sdtContentLocked',
      multiLine: true,
      appearance: 'tags',
      color: '#ddebf7',
    });
    expect(
      contentControlDomAttributes(DOCUMENT_CONTENT_CONTROL_DEFAULTS),
    ).toMatchObject({
      'data-document-content-control': 'true',
      'data-content-control-type': 'text',
      'data-content-control-lock': 'unlocked',
      'data-content-control-multiline': 'false',
      'data-content-control-appearance': 'boundingBox',
    });
    expect(
      contentControlDomAttributes({
        tag: 'customer-name',
        appearance: 'tags',
      })['data-content-control-label'],
    ).toBe('customer-name');
  });

  test('wraps a selection, edits its text, and keeps the operation undoable', () => {
    editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content:
        '<section data-document-section="true"><p>Before value after</p></section>',
    });
    editor.commands.setTextSelection({ from: 9, to: 14 });
    expect(
      editor.commands.insertDocumentContentControl({
        id: 'customer-name',
        alias: 'Customer name',
        tag: 'customer',
      }),
    ).toBe(true);
    expect(editor.getHTML()).toContain('data-document-content-control="true"');
    expect(editor.getHTML()).toContain(
      'data-content-control-alias="Customer name"',
    );
    expect(editor.getText()).toContain('value');
    editor.commands.insertContent('updated');
    expect(editor.getText()).toContain('updated');
    expect(editor.commands.undo()).toBe(true);
    expect(editor.getText()).toContain('value');
  });

  test('rejects edits to content-locked controls and requires explicit unlock to delete', () => {
    editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content:
        '<section data-document-section="true"><p><span data-document-content-control="true" data-content-control-id="locked" data-content-control-lock="contentLocked">Fixed</span></p></section>',
    });
    expect(editor.getHTML()).toContain('data-content-control-id="locked"');
    editor.commands.setTextSelection({ from: 3, to: 8 });
    editor.commands.insertContent('Changed');
    expect(editor.getText()).toContain('Fixed');
    expect(editor.getText()).not.toContain('Changed');
  });

  test('keeps a shell-locked control intact until deletion is explicitly allowed', () => {
    editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content:
        '<section data-document-section="true"><p><span data-document-content-control="true" data-content-control-id="shell-locked" data-content-control-lock="sdtLocked">Fixed</span></p></section>',
    });
    const position = contentControlPosition(editor, 'shell-locked');
    editor.commands.setNodeSelection(position);
    expect(editor.commands.deleteDocumentContentControl()).toBe(false);
    expect(editor.getText()).toContain('Fixed');
    expect(
      editor.commands.deleteDocumentContentControl({ allowLocked: true }),
    ).toBe(true);
    expect(editor.getText()).not.toContain('Fixed');
  });

  test('rejects generic metadata mutations for locked control shells', () => {
    editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content:
        '<section data-document-section="true"><p><span data-document-content-control="true" data-content-control-id="shell-locked" data-content-control-lock="sdtLocked">Editable value</span></p></section>',
    });
    const position = contentControlPosition(editor, 'shell-locked');
    const node = editor.state.doc.nodeAt(position);
    if (!node) throw new Error('Content control was not found.');
    const transaction = editor.state.tr.setNodeMarkup(position, undefined, {
      ...node.attrs,
      alias: 'Should not apply',
    });
    editor.view.dispatch(transaction);
    expect(editor.getHTML()).not.toContain('Should not apply');
    expect(editor.getText()).toContain('Editable value');
  });

  test('allows shell-locked content edits but blocks content-locked metadata and text', () => {
    editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content:
        '<section data-document-section="true"><p><span data-document-content-control="true" data-content-control-id="shell" data-content-control-lock="sdtLocked">Shell</span> <span data-document-content-control="true" data-content-control-id="content" data-content-control-lock="contentLocked">Content</span></p></section>',
    });
    const shellPosition = contentControlPosition(editor, 'shell');
    const shellNode = editor.state.doc.nodeAt(shellPosition);
    if (!shellNode) throw new Error('Shell-locked control was not found.');
    editor.commands.setTextSelection({
      from: shellPosition + 1,
      to: shellPosition + 1 + shellNode.content.size,
    });
    expect(editor.commands.insertContent('Edited shell')).toBe(true);
    expect(editor.getText()).toContain('Edited shell');

    const contentPosition = contentControlPosition(editor, 'content');
    const contentNode = editor.state.doc.nodeAt(contentPosition);
    if (!contentNode) throw new Error('Content-locked control was not found.');
    const metadataTransaction = editor.state.tr.setNodeMarkup(
      contentPosition,
      undefined,
      { ...contentNode.attrs, alias: 'Should not apply' },
    );
    editor.view.dispatch(metadataTransaction);
    expect(editor.getHTML()).not.toContain('Should not apply');
    editor.commands.setTextSelection({
      from: contentPosition + 1,
      to: contentPosition + 1 + contentNode.content.size,
    });
    editor.commands.insertContent('Changed');
    expect(editor.getText()).toContain('Content');
    expect(editor.getText()).not.toContain('Changed');
  });

  test('marks and restores a safe inline DOCX control while rejecting active forms', () => {
    const document = parseXml(
      `<w:document xmlns:w="${WORD_NAMESPACE}" xmlns:w15="${WORD_2012_NAMESPACE}"><w:body><w:p><w:r><w:t>Before </w:t></w:r><w:sdt><w:sdtPr><w:id w:val="41"/><w:alias w:val="Customer name"/><w:tag w:val="customer"/><w:text w:multiLine="0"/><w:lock w:val="contentLocked"/><w15:appearance w15:val="tags"/><w15:color w15:val="DDEBF7"/></w:sdtPr><w:sdtContent><w:r><w:rPr><w:b/></w:rPr><w:t>Lin</w:t></w:r><w:r><w:t>Da</w:t></w:r></w:sdtContent></w:sdt><w:r><w:t> after</w:t></w:r></w:p><w:p><w:sdt><w:sdtPr><w:dataBinding w:storeItemID="{unsafe}"/></w:sdtPr><w:sdtContent><w:r><w:t>Do not bind</w:t></w:r></w:sdtContent></w:sdt></w:p></w:body></w:document>`,
    );
    expect(inspectDocxContentControls(document)).toEqual({
      supported: 1,
      unsupported: 1,
    });
    const markers = markDocxContentControls(document);
    expect(markers.controls).toHaveLength(1);
    expect(markers.unsupported).toBe(1);
    expect(document.documentElement.textContent).toContain(
      markers.controls[0]?.start,
    );
    const html = new DOMParser().parseFromString(
      `<p>Before ${markers.controls[0]?.start}<strong>Lin</strong>Da${markers.controls[0]?.end} after</p>`,
      'text/html',
    );
    applyImportedDocxContentControlMarkers(html, markers);
    const control = html.body.querySelector<HTMLElement>(
      '[data-document-content-control]',
    );
    expect(control?.textContent).toBe('LinDa');
    expect(control?.dataset.contentControlNativeId).toBe('41');
    expect(control?.dataset.contentControlAlias).toBe('Customer name');
    expect(control?.dataset.contentControlLock).toBe('contentLocked');
    expect(control?.querySelector('strong')?.textContent).toBe('Lin');
    expect(html.body.textContent).not.toContain('__A3S_WORK_CONTENT_CONTROL_');
  });

  test('moves content-control markers away from source text collisions', () => {
    const document = parseXml(
      `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:p><w:r><w:t>__A3S_WORK_CONTENT_CONTROL_START_1__</w:t></w:r><w:sdt><w:sdtPr><w:id w:val="9"/><w:text/></w:sdtPr><w:sdtContent><w:r><w:t>Value</w:t></w:r></w:sdtContent></w:sdt></w:p></w:body></w:document>`,
    );
    const markers = markDocxContentControls(document);
    expect(markers.controls).toHaveLength(1);
    expect(markers.controls[0]?.start).not.toBe(
      '__A3S_WORK_CONTENT_CONTROL_START_1__',
    );
    expect(document.documentElement.textContent).toContain(
      markers.controls[0]?.start,
    );
  });

  test('diagnoses block, nested, and namespace-spoofed controls without reviving them', () => {
    const document = parseXml(
      `<w:document xmlns:w="${WORD_NAMESPACE}" xmlns:s="${'http://purl.oclc.org/ooxml/wordprocessingml/main'}"><w:body><w:sdt><w:sdtPr><w:text/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>Block control</w:t></w:r></w:p></w:sdtContent></w:sdt><w:p><w:sdt><w:sdtPr><w:text/></w:sdtPr><w:sdtContent><w:p><w:sdt><w:sdtPr><w:text/></w:sdtPr><w:sdtContent><w:r><w:t>Nested control</w:t></w:r></w:sdtContent></w:sdt></w:p></w:sdtContent></w:sdt></w:p><w:p><w:sdt><w:sdtPr><w:id s:val="7"/><w:text/></w:sdtPr><w:sdtContent><w:r><w:t>Spoofed id</w:t></w:r></w:sdtContent></w:sdt></w:p></w:body></w:document>`,
    );
    expect(inspectDocxContentControls(document)).toEqual({
      supported: 0,
      unsupported: 3,
    });
    const markers = markDocxContentControls(document);
    expect(markers.controls).toHaveLength(0);
    expect(markers.unsupported).toBe(3);
    expect(document.documentElement.textContent).toContain('Nested control');
    expect(document.documentElement.textContent).not.toContain(
      '__A3S_WORK_CONTENT_CONTROL_START_',
    );
  });

  test('exports a native w:sdt and reopens it as an editable control', async () => {
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    artifact.content.html =
      '<section data-document-section="true"><p>Before <span data-document-content-control="true" data-content-control-id="customer" data-content-control-native-id="41" data-content-control-type="text" data-content-control-alias="Customer name" data-content-control-tag="customer" data-content-control-lock="unlocked" data-content-control-multiline="false" data-content-control-appearance="tags" data-content-control-color="#ddebf7">Lin<strong>Da</strong></span> after</p></section>';
    const blob = await createDocxBlob(artifact.content);
    const archive = await JSZip.loadAsync(await blob.arrayBuffer());
    const xml = (await archive.file('word/document.xml')?.async('text')) ?? '';
    expect(xml).toContain('<w:sdt>');
    expect(xml).toContain('<w:sdtPr>');
    expect(xml).toContain('w:val="41"');
    expect(xml).toContain('w:val="Customer name"');
    expect(xml).toContain('w:val="customer"');
    expect(xml).toContain('w15:appearance');
    expect(xml).toContain('w15:color');
    expect(xml).not.toContain('__A3S_WORK_CONTENT_CONTROL_');
    const reopened = await importOfficeFile(
      new File([blob], 'content-control.docx', { type: blob.type }),
    );
    if (reopened.content.type !== 'document') {
      throw new Error('Expected a reopened document artifact.');
    }
    expect(reopened.content.html).toContain(
      'data-document-content-control="true"',
    );
    expect(reopened.content.html).toContain(
      'data-content-control-native-id="41"',
    );
    expect(reopened.content.html).toContain('Customer name');
    expect(reopened.content.html).toContain('Lin');
    expect(reopened.content.html).toContain('Da');
    const compatibility = await analyzeDocxCompatibility(
      new File([blob], 'content-control.docx', { type: blob.type }),
      [],
    );
    expect(compatibility.issues).toContainEqual(
      expect.objectContaining({
        code: 'docx.content-controls',
        severity: 'info',
      }),
    );
  });

  test('allocates collision-free native IDs and round-trips empty controls', async () => {
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    artifact.content.html =
      '<section data-document-section="true"><p><span data-document-content-control="true" data-content-control-native-id="7" data-content-control-id="first">A</span><span data-document-content-control="true" data-content-control-native-id="7" data-content-control-id="second"></span></p></section>';
    const blob = await createDocxBlob(artifact.content);
    const archive = await JSZip.loadAsync(await blob.arrayBuffer());
    const xml = (await archive.file('word/document.xml')?.async('text')) ?? '';
    const ids = Array.from(
      xml.matchAll(/<w:id w:val="(-?\d+)"\s*\/>/g),
      (match) => Number(match[1]),
    );
    expect(ids).toContain(7);
    expect(new Set(ids).size).toBe(ids.length);
    expect(xml.match(/<w:sdt>/g)?.length).toBe(2);
    expect(xml).not.toContain('__A3S_WORK_CONTENT_CONTROL_');

    const reopened = await importOfficeFile(
      new File([blob], 'empty-content-control.docx', { type: blob.type }),
    );
    if (reopened.content.type !== 'document') {
      throw new Error('Expected a reopened document artifact.');
    }
    const reopenedControls = new DOMParser()
      .parseFromString(reopened.content.html, 'text/html')
      .body.querySelectorAll('[data-document-content-control]');
    expect(reopenedControls).toHaveLength(2);
    expect(reopenedControls[0]?.textContent).toBe('A');
    expect(reopenedControls[1]?.textContent).toBe('');
    expect(reopenedControls[0]?.dataset.contentControlId).not.toBe(
      reopenedControls[1]?.dataset.contentControlId,
    );
  });
});

function contentControlPosition(currentEditor: Editor, id: string): number {
  let found: number | null = null;
  currentEditor.state.doc.descendants((node, position) => {
    if (found !== null || node.type.name !== 'documentContentControl') return;
    if (node.attrs.id === id) found = position;
  });
  if (found === null) throw new Error(`Content control "${id}" was not found.`);
  return found;
}
