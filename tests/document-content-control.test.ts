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
      checked: false,
      options: [],
      selectedValue: '',
      fullDate: '',
      dateFormat: 'yyyy年M月d日',
      dateLanguage: 'zh-CN',
      dateMapping: 'dateTime',
      bindingStoreItemId: '',
      bindingXPath: '',
      bindingPrefixMappings: '',
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
      normalizeDocumentContentControlProperties({
        type: 'checkbox',
        checked: 'true',
        multiLine: true,
      }),
    ).toEqual(
      expect.objectContaining({
        type: 'checkbox',
        checked: true,
        multiLine: false,
        options: [],
        selectedValue: '',
      }),
    );
    expect(
      normalizeDocumentContentControlProperties({
        type: 'dropDownList',
        options: [
          { displayText: 'Yes', value: 'yes' },
          { displayText: 'No', value: 'no' },
          { displayText: 'Yes', value: 'yes' },
        ],
        selectedValue: 'no',
        multiLine: true,
      }),
    ).toEqual(
      expect.objectContaining({
        type: 'dropDownList',
        multiLine: false,
        options: [
          { displayText: 'Yes', value: 'yes' },
          { displayText: 'No', value: 'no' },
        ],
        selectedValue: 'no',
      }),
    );
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

  test('inserts a checkbox glyph, toggles checked state, and round-trips w14:checkbox', async () => {
    editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: '<section data-document-section="true"><p>Ready</p></section>',
    });
    editor.commands.setTextSelection({ from: 2, to: 2 });
    expect(
      editor.commands.insertDocumentContentControl({
        id: 'accept',
        type: 'checkbox',
        alias: 'Accept terms',
        tag: 'accept',
        checked: false,
      }),
    ).toBe(true);
    expect(editor.getHTML()).toContain('data-content-control-type="checkbox"');
    expect(editor.getHTML()).toContain('data-content-control-checked="false"');
    expect(editor.getHTML()).toContain('role="checkbox"');
    expect(editor.getText()).toContain('\u2610');
    const position = contentControlPosition(editor, 'accept');
    editor.commands.setNodeSelection(position);
    expect(
      editor.commands.setDocumentContentControlProperties({ checked: true }),
    ).toBe(true);
    expect(editor.getHTML()).toContain('data-content-control-checked="true"');
    expect(editor.getText()).toContain('\u2611');

    const document = parseXml(
      `<w:document xmlns:w="${WORD_NAMESPACE}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body><w:p><w:sdt><w:sdtPr><w:id w:val="88"/><w:alias w:val="Accept terms"/><w:tag w:val="accept"/><w14:checkbox><w14:checked w14:val="1"/><w14:checkedState w14:val="2611" w14:font="Segoe UI Symbol"/><w14:uncheckedState w14:val="2610" w14:font="Segoe UI Symbol"/></w14:checkbox></w:sdtPr><w:sdtContent><w:r><w:t>\u2611</w:t></w:r></w:sdtContent></w:sdt></w:p></w:body></w:document>`,
    );
    expect(inspectDocxContentControls(document)).toEqual({
      supported: 1,
      unsupported: 0,
    });
    const markers = markDocxContentControls(document);
    expect(markers.controls).toHaveLength(1);
    expect(markers.controls[0]?.properties).toMatchObject({
      type: 'checkbox',
      checked: true,
      alias: 'Accept terms',
      tag: 'accept',
      nativeId: 88,
    });

    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    artifact.content.html =
      '<section data-document-section="true"><p><span data-document-content-control="true" data-content-control-id="accept" data-content-control-native-id="88" data-content-control-type="checkbox" data-content-control-alias="Accept terms" data-content-control-tag="accept" data-content-control-lock="unlocked" data-content-control-multiline="false" data-content-control-appearance="boundingBox" data-content-control-checked="true">\u2611</span></p></section>';
    const blob = await createDocxBlob(artifact.content);
    const archive = await JSZip.loadAsync(await blob.arrayBuffer());
    const xml = (await archive.file('word/document.xml')?.async('text')) ?? '';
    expect(xml).toContain('w14:checkbox');
    expect(xml).toContain('w14:checked');
    expect(xml).toContain('w14:val="1"');
    expect(xml).toContain('xmlns:w14=');
    const reopened = await importOfficeFile(
      new File([blob], 'checkbox-content-control.docx', { type: blob.type }),
    );
    if (reopened.content.type !== 'document') {
      throw new Error('Expected a reopened document artifact.');
    }
    expect(reopened.content.html).toContain(
      'data-content-control-type="checkbox"',
    );
    expect(reopened.content.html).toContain(
      'data-content-control-checked="true"',
    );
    expect(reopened.content.html).toContain('\u2611');
  });

  test('inserts a drop-down list, changes selection, and round-trips w:dropDownList', async () => {
    editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: '<section data-document-section="true"><p>Ready</p></section>',
    });
    editor.commands.setTextSelection({ from: 2, to: 2 });
    expect(
      editor.commands.insertDocumentContentControl({
        id: 'status',
        type: 'dropDownList',
        alias: 'Status',
        tag: 'status',
        options: [
          { displayText: 'Open', value: 'open' },
          { displayText: 'Closed', value: 'closed' },
        ],
        selectedValue: 'open',
      }),
    ).toBe(true);
    expect(editor.getHTML()).toContain(
      'data-content-control-type="dropDownList"',
    );
    expect(editor.getHTML()).toContain('role="listbox"');
    expect(editor.getText()).toContain('Open');
    const position = contentControlPosition(editor, 'status');
    editor.commands.setNodeSelection(position);
    expect(
      editor.commands.setDocumentContentControlProperties({
        selectedValue: 'closed',
      }),
    ).toBe(true);
    expect(editor.getText()).toContain('Closed');

    const document = parseXml(
      `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:p><w:sdt><w:sdtPr><w:id w:val="77"/><w:alias w:val="Status"/><w:tag w:val="status"/><w:dropDownList><w:listItem w:displayText="Open" w:value="open"/><w:listItem w:displayText="Closed" w:value="closed"/></w:dropDownList></w:sdtPr><w:sdtContent><w:r><w:t>Closed</w:t></w:r></w:sdtContent></w:sdt></w:p></w:body></w:document>`,
    );
    expect(inspectDocxContentControls(document)).toEqual({
      supported: 1,
      unsupported: 0,
    });
    const markers = markDocxContentControls(document);
    expect(markers.controls).toHaveLength(1);
    expect(markers.controls[0]?.properties).toMatchObject({
      type: 'dropDownList',
      selectedValue: 'closed',
      alias: 'Status',
      options: [
        { displayText: 'Open', value: 'open' },
        { displayText: 'Closed', value: 'closed' },
      ],
    });

    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    artifact.content.html =
      '<section data-document-section="true"><p><span data-document-content-control="true" data-content-control-id="status" data-content-control-native-id="77" data-content-control-type="dropDownList" data-content-control-alias="Status" data-content-control-tag="status" data-content-control-lock="unlocked" data-content-control-multiline="false" data-content-control-appearance="boundingBox" data-content-control-options=\'[{"displayText":"Open","value":"open"},{"displayText":"Closed","value":"closed"}]\' data-content-control-selected-value="closed">Closed</span></p></section>';
    const blob = await createDocxBlob(artifact.content);
    const archive = await JSZip.loadAsync(await blob.arrayBuffer());
    const xml = (await archive.file('word/document.xml')?.async('text')) ?? '';
    expect(xml).toContain('w:dropDownList');
    expect(xml).toContain('w:listItem');
    expect(xml).toContain('w:value="open"');
    expect(xml).toContain('w:displayText="Closed"');
    const reopened = await importOfficeFile(
      new File([blob], 'dropdown-content-control.docx', { type: blob.type }),
    );
    if (reopened.content.type !== 'document') {
      throw new Error('Expected a reopened document artifact.');
    }
    expect(reopened.content.html).toContain(
      'data-content-control-type="dropDownList"',
    );
    expect(reopened.content.html).toContain(
      'data-content-control-selected-value="closed"',
    );
    expect(reopened.content.html).toContain('Closed');
  });

  test('inserts a combo box, keeps free text, and round-trips w:comboBox', async () => {
    editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: '<section data-document-section="true"><p>Ready</p></section>',
    });
    editor.commands.setTextSelection({ from: 2, to: 2 });
    expect(
      editor.commands.insertDocumentContentControl({
        id: 'dept',
        type: 'comboBox',
        alias: 'Department',
        tag: 'dept',
        options: [
          { displayText: 'Sales', value: 'sales' },
          { displayText: 'Engineering', value: 'eng' },
        ],
        selectedValue: 'sales',
      }),
    ).toBe(true);
    expect(editor.getHTML()).toContain('data-content-control-type="comboBox"');
    expect(editor.getHTML()).toContain('role="combobox"');
    expect(editor.getText()).toContain('Sales');
    const position = contentControlPosition(editor, 'dept');
    editor.commands.setNodeSelection(position);
    expect(
      editor.commands.setDocumentContentControlProperties({
        selectedValue: 'eng',
      }),
    ).toBe(true);
    expect(editor.getText()).toContain('Engineering');

    editor.destroy();
    editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: '<section data-document-section="true"><p>Ready</p></section>',
    });
    editor.commands.setTextSelection({ from: 2, to: 2 });
    expect(
      editor.commands.insertDocumentContentControl({
        id: 'custom',
        type: 'comboBox',
        alias: 'Custom',
        tag: 'custom',
        options: [
          { displayText: 'Sales', value: 'sales' },
          { displayText: 'Engineering', value: 'eng' },
        ],
        selectedValue: '',
        text: 'Custom desk',
      }),
    ).toBe(true);
    expect(editor.getText()).toContain('Custom desk');
    expect(editor.getHTML()).toContain('data-content-control-type="comboBox"');

    const document = parseXml(
      `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:p><w:sdt><w:sdtPr><w:id w:val="66"/><w:alias w:val="Department"/><w:tag w:val="dept"/><w:comboBox><w:listItem w:displayText="Sales" w:value="sales"/><w:listItem w:displayText="Engineering" w:value="eng"/></w:comboBox></w:sdtPr><w:sdtContent><w:r><w:t>Custom desk</w:t></w:r></w:sdtContent></w:sdt></w:p></w:body></w:document>`,
    );
    expect(inspectDocxContentControls(document)).toEqual({
      supported: 1,
      unsupported: 0,
    });
    const markers = markDocxContentControls(document);
    expect(markers.controls).toHaveLength(1);
    expect(markers.controls[0]?.properties).toMatchObject({
      type: 'comboBox',
      selectedValue: '',
      alias: 'Department',
      options: [
        { displayText: 'Sales', value: 'sales' },
        { displayText: 'Engineering', value: 'eng' },
      ],
    });

    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    artifact.content.html =
      '<section data-document-section="true"><p><span data-document-content-control="true" data-content-control-id="dept" data-content-control-native-id="66" data-content-control-type="comboBox" data-content-control-alias="Department" data-content-control-tag="dept" data-content-control-lock="unlocked" data-content-control-multiline="false" data-content-control-appearance="boundingBox" data-content-control-options=\'[{"displayText":"Sales","value":"sales"},{"displayText":"Engineering","value":"eng"}]\'>Custom desk</span></p></section>';
    const blob = await createDocxBlob(artifact.content);
    const archive = await JSZip.loadAsync(await blob.arrayBuffer());
    const xml = (await archive.file('word/document.xml')?.async('text')) ?? '';
    expect(xml).toContain('w:comboBox');
    expect(xml).toContain('w:listItem');
    expect(xml).toContain('w:value="sales"');
    expect(xml).toContain('Custom desk');
    const reopened = await importOfficeFile(
      new File([blob], 'combo-content-control.docx', { type: blob.type }),
    );
    if (reopened.content.type !== 'document') {
      throw new Error('Expected a reopened document artifact.');
    }
    expect(reopened.content.html).toContain(
      'data-content-control-type="comboBox"',
    );
    expect(reopened.content.html).toContain('Custom desk');
  });

  test('inserts a date control, changes value, and round-trips w:date', async () => {
    editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: '<section data-document-section="true"><p>Ready</p></section>',
    });
    editor.commands.setTextSelection({ from: 2, to: 2 });
    expect(
      editor.commands.insertDocumentContentControl({
        id: 'due',
        type: 'date',
        alias: 'Due date',
        tag: 'due',
        fullDate: '2026-09-22T00:00:00Z',
        dateFormat: 'yyyy-MM-dd',
        dateLanguage: 'en-US',
        dateMapping: 'dateTime',
      }),
    ).toBe(true);
    expect(editor.getHTML()).toContain('data-content-control-type="date"');
    expect(editor.getHTML()).toContain(
      'data-content-control-full-date="2026-09-22T00:00:00Z"',
    );
    expect(editor.getText()).toContain('2026-09-22');
    const position = contentControlPosition(editor, 'due');
    editor.commands.setNodeSelection(position);
    expect(
      editor.commands.setDocumentContentControlProperties({
        fullDate: '2026-10-01T00:00:00Z',
        dateFormat: 'yyyy年M月d日',
      }),
    ).toBe(true);
    expect(editor.getText()).toContain('2026年10月1日');

    const document = parseXml(
      `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:p><w:sdt><w:sdtPr><w:id w:val="55"/><w:alias w:val="Due date"/><w:tag w:val="due"/><w:date w:fullDate="2026-10-01T00:00:00Z"><w:dateFormat w:val="yyyy年M月d日"/><w:lid w:val="zh-CN"/><w:storeMappedDataAs w:val="dateTime"/><w:calendar w:val="gregorian"/></w:date></w:sdtPr><w:sdtContent><w:r><w:t>2026年10月1日</w:t></w:r></w:sdtContent></w:sdt></w:p></w:body></w:document>`,
    );
    expect(inspectDocxContentControls(document)).toEqual({
      supported: 1,
      unsupported: 0,
    });
    const markers = markDocxContentControls(document);
    expect(markers.controls).toHaveLength(1);
    expect(markers.controls[0]?.properties).toMatchObject({
      type: 'date',
      fullDate: '2026-10-01T00:00:00Z',
      dateFormat: 'yyyy年M月d日',
      dateLanguage: 'zh-CN',
      dateMapping: 'dateTime',
      alias: 'Due date',
    });

    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    artifact.content.html =
      '<section data-document-section="true"><p><span data-document-content-control="true" data-content-control-id="due" data-content-control-native-id="55" data-content-control-type="date" data-content-control-alias="Due date" data-content-control-tag="due" data-content-control-lock="unlocked" data-content-control-multiline="false" data-content-control-appearance="boundingBox" data-content-control-full-date="2026-10-01T00:00:00Z" data-content-control-date-format="yyyy年M月d日" data-content-control-date-language="zh-CN" data-content-control-date-mapping="dateTime">2026年10月1日</span></p></section>';
    const blob = await createDocxBlob(artifact.content);
    const archive = await JSZip.loadAsync(await blob.arrayBuffer());
    const xml = (await archive.file('word/document.xml')?.async('text')) ?? '';
    expect(xml).toContain('w:date');
    expect(xml).toContain('w:fullDate="2026-10-01T00:00:00Z"');
    expect(xml).toContain('w:dateFormat');
    expect(xml).toContain('yyyy年M月d日');
    const reopened = await importOfficeFile(
      new File([blob], 'date-content-control.docx', { type: blob.type }),
    );
    if (reopened.content.type !== 'document') {
      throw new Error('Expected a reopened document artifact.');
    }
    expect(reopened.content.html).toContain('data-content-control-type="date"');
    expect(reopened.content.html).toContain(
      'data-content-control-full-date="2026-10-01T00:00:00Z"',
    );
    expect(reopened.content.html).toContain('2026年10月1日');
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

  test('admits bounded dataBinding metadata and round-trips native w:dataBinding', async () => {
    expect(
      normalizeDocumentContentControlProperties({
        type: 'text',
        bindingStoreItemId: '{55CF591A-3D7D-47A0-974C-028795EBCBC9}',
        bindingXPath: '/ns0:root[1]/ns0:name[1]',
        bindingPrefixMappings:
          "xmlns:ns0='http://schemas.example.com/customer'",
      }),
    ).toMatchObject({
      bindingStoreItemId: '{55CF591A-3D7D-47A0-974C-028795EBCBC9}',
      bindingXPath: '/ns0:root[1]/ns0:name[1]',
      bindingPrefixMappings: "xmlns:ns0='http://schemas.example.com/customer'",
    });
    expect(
      normalizeDocumentContentControlProperties({
        type: 'text',
        bindingStoreItemId: '{unsafe}',
        bindingXPath: '/ns0:root[1]',
      }),
    ).toMatchObject({
      bindingStoreItemId: '',
      bindingXPath: '',
      bindingPrefixMappings: '',
    });

    const document = parseXml(
      `<w:document xmlns:w="${WORD_NAMESPACE}"><w:body><w:p><w:sdt><w:sdtPr><w:id w:val="77"/><w:alias w:val="Customer"/><w:tag w:val="cust"/><w:dataBinding w:prefixMappings="xmlns:ns0='http://schemas.example.com/customer'" w:xpath="/ns0:root[1]/ns0:name[1]" w:storeItemID="{55CF591A-3D7D-47A0-974C-028795EBCBC9}"/><w:text w:multiLine="0"/></w:sdtPr><w:sdtContent><w:r><w:t>LinDa</w:t></w:r></w:sdtContent></w:sdt></w:p></w:body></w:document>`,
    );
    expect(inspectDocxContentControls(document)).toEqual({
      supported: 1,
      unsupported: 0,
    });
    const markers = markDocxContentControls(document);
    expect(markers.controls).toHaveLength(1);
    expect(markers.controls[0]?.properties).toMatchObject({
      type: 'text',
      alias: 'Customer',
      bindingStoreItemId: '{55CF591A-3D7D-47A0-974C-028795EBCBC9}',
      bindingXPath: '/ns0:root[1]/ns0:name[1]',
      bindingPrefixMappings: "xmlns:ns0='http://schemas.example.com/customer'",
    });

    editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: '<section data-document-section="true"><p>Ready</p></section>',
    });
    editor.commands.setTextSelection({ from: 2, to: 2 });
    expect(
      editor.commands.insertDocumentContentControl({
        id: 'bound-name',
        type: 'text',
        alias: 'Customer',
        tag: 'cust',
        text: 'LinDa',
        bindingStoreItemId: '{55CF591A-3D7D-47A0-974C-028795EBCBC9}',
        bindingXPath: '/ns0:root[1]/ns0:name[1]',
        bindingPrefixMappings:
          "xmlns:ns0='http://schemas.example.com/customer'",
      }),
    ).toBe(true);
    expect(editor.getHTML()).toContain(
      'data-content-control-binding-store-item-id="{55CF591A-3D7D-47A0-974C-028795EBCBC9}"',
    );
    expect(editor.getHTML()).toContain(
      'data-content-control-binding-xpath="/ns0:root[1]/ns0:name[1]"',
    );

    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    artifact.content.html =
      '<section data-document-section="true"><p><span data-document-content-control="true" data-content-control-id="bound-name" data-content-control-native-id="77" data-content-control-type="text" data-content-control-alias="Customer" data-content-control-tag="cust" data-content-control-lock="unlocked" data-content-control-multiline="false" data-content-control-appearance="boundingBox" data-content-control-binding-store-item-id="{55CF591A-3D7D-47A0-974C-028795EBCBC9}" data-content-control-binding-xpath="/ns0:root[1]/ns0:name[1]" data-content-control-binding-prefix-mappings="xmlns:ns0=\'http://schemas.example.com/customer\'">LinDa</span></p></section>';
    const blob = await createDocxBlob(artifact.content);
    const archive = await JSZip.loadAsync(await blob.arrayBuffer());
    const xml = (await archive.file('word/document.xml')?.async('text')) ?? '';
    expect(xml).toContain('w:dataBinding');
    expect(xml).toContain(
      'w:storeItemID="{55CF591A-3D7D-47A0-974C-028795EBCBC9}"',
    );
    expect(xml).toContain('w:xpath="/ns0:root[1]/ns0:name[1]"');
    expect(xml).toContain(
      'w:prefixMappings="xmlns:ns0=\'http://schemas.example.com/customer\'"',
    );
    const reopened = await importOfficeFile(
      new File([blob], 'bound-content-control.docx', { type: blob.type }),
    );
    if (reopened.content.type !== 'document') {
      throw new Error('Expected a reopened document artifact.');
    }
    expect(reopened.content.html).toContain(
      'data-content-control-binding-store-item-id="{55CF591A-3D7D-47A0-974C-028795EBCBC9}"',
    );
    expect(reopened.content.html).toContain(
      'data-content-control-binding-xpath="/ns0:root[1]/ns0:name[1]"',
    );
    expect(reopened.content.html).toContain('LinDa');
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
      unsupported: 2,
    });
    const markers = markDocxContentControls(document);
    expect(markers.controls).toHaveLength(0);
    expect(markers.unsupported).toBe(2);
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
