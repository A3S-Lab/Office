import { Editor } from '@tiptap/core';
import { describe, expect, test } from '@rstest/core';
import JSZip from 'jszip';
import {
  createArtifact,
  createArtifactBlob,
  importOfficeFile,
} from '../src/core';
import { collectDocumentChanges } from '../src/internal/features/work/work-document-changes';
import { createWorkDocumentExtensions } from '../src/internal/features/work/work-document-extensions';
import { parseDocumentRowFormatting } from '../src/internal/features/work/work-document-row-format-changes';
import { analyzeDocxCompatibility } from '../src/internal/features/work/work-office-diagnostics';
import {
  descendants,
  directChild,
  parseXml,
} from '../src/internal/features/work/work-ooxml-package';
import {
  xmlAttributeLocalName,
  xmlAttributeNamespace,
} from '../src/internal/features/work/work-docx-settings-xml';

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

describe('DOCX row-formatting revisions', () => {
  test('imports cantSplit-only w:trPrChange as a reviewable row-formatting change', async () => {
    const source = await rowDocxWithCantSplitChange({ prior: true });
    const imported = await importOfficeFile(
      new File([source], 'row-formatting.docx'),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const html = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    const row = html.body.querySelector('tr');
    expect(row?.dataset.changeKind).toBe('row-formatting');
    expect(row?.dataset.officeRowPropertyRevisionOmml).toBeUndefined();
    expect(parseDocumentRowFormatting(row?.dataset.changeBefore)).toEqual({
      cantSplit: true,
    });

    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: imported.content.html,
    });
    try {
      const changes = collectDocumentChanges(editor.state.doc);
      expect(changes).toHaveLength(1);
      expect(changes[0]?.kind).toBe('row-formatting');
      expect(editor.commands.acceptDocumentChange(changes[0]?.id ?? '')).toBe(
        true,
      );
      expect(collectDocumentChanges(editor.state.doc)).toHaveLength(0);
    } finally {
      editor.destroy();
    }
  });

  test('reject restores prior cantSplit and drops the pending change', async () => {
    const source = await rowDocxWithCantSplitChange({ prior: false });
    const imported = await importOfficeFile(
      new File([source], 'row-formatting-reject.docx'),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: imported.content.html,
    });
    try {
      const change = collectDocumentChanges(editor.state.doc)[0];
      expect(editor.commands.rejectDocumentChange(change?.id ?? '')).toBe(true);
      const html = new DOMParser().parseFromString(
        editor.getHTML(),
        'text/html',
      );
      const row = html.body.querySelector('tr');
      expect(row?.dataset.changeKind).toBeUndefined();
      expect(row?.dataset.officeCantSplit).toBe('false');
    } finally {
      editor.destroy();
    }
  });

  test('pending row-formatting change round-trips as native w:trPrChange', async () => {
    const source = await rowDocxWithCantSplitChange({ prior: true });
    const imported = await importOfficeFile(
      new File([source], 'row-formatting-roundtrip.docx'),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: imported.content.html,
    });
    try {
      imported.content.html = editor.getHTML();
    } finally {
      editor.destroy();
    }
    const exported = await createArtifactBlob(imported);
    const archive = await JSZip.loadAsync(await exported.arrayBuffer());
    const document = await xmlEntry(archive, 'word/document.xml');
    const change = directChild(
      directChild(descendants(document, 'tr')[0], 'trPr'),
      'trPrChange',
    );
    expect(change).toBeTruthy();
    expect(wordAttribute(change!, 'author')).toBe('Reviewer');
    expect(directChild(directChild(change, 'trPr'), 'cantSplit')).toBeTruthy();
  });

  test('imports tblHeader-only w:trPrChange as a reviewable row-formatting change', async () => {
    const source = await rowDocxWithHeaderChange({ prior: true });
    const imported = await importOfficeFile(
      new File([source], 'row-formatting-header.docx'),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const html = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    const row = html.body.querySelector('tr');
    expect(row?.dataset.changeKind).toBe('row-formatting');
    expect(parseDocumentRowFormatting(row?.dataset.changeBefore)).toEqual({
      repeatHeader: true,
    });
  });

  test('reports cantSplit-only row-property revisions as reviewable diagnostics', async () => {
    const source = await rowDocxWithCantSplitChange({ prior: true });
    const summary = await analyzeDocxCompatibility(
      new File([source], 'row-formatting-diag.docx'),
      [],
    );
    expect(
      summary.issues.some(
        (issue) => issue.code === 'docx.revisions.row-formatting',
      ),
    ).toBe(true);
  });

  test('live row cantSplit edits become reviewable when track changes is on', () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions({
        isTracking: () => true,
      }),
      content:
        '<table><tbody><tr data-office-cant-split="false"><td><p>Cell</p></td></tr></tbody></table>',
    });
    try {
      expect(
        editor.commands.setDocumentTableRowOptions({
          cantSplit: true,
          repeatHeader: false,
        }),
      ).toBe(true);
      const changes = collectDocumentChanges(editor.state.doc);
      expect(changes).toHaveLength(1);
      expect(changes[0]?.kind).toBe('row-formatting');
      const html = new DOMParser().parseFromString(
        editor.getHTML(),
        'text/html',
      );
      const row = html.body.querySelector('tr');
      expect(row?.dataset.changeKind).toBe('row-formatting');
      expect(parseDocumentRowFormatting(row?.dataset.changeBefore)).toEqual({
        cantSplit: false,
        repeatHeader: false,
        hidden: false,
        alignment: 'left',
        gridBefore: 0,
        gridAfter: 0,
        widthBefore: { type: 'auto', value: null },
        widthAfter: { type: 'auto', value: null },
      });
      expect(row?.dataset.officeCantSplit).toBe('true');
    } finally {
      editor.destroy();
    }
  });

  test('imports hidden-only w:trPrChange as a reviewable row-formatting change', async () => {
    const source = await rowDocxWithHiddenChange({ prior: true });
    const imported = await importOfficeFile(
      new File([source], 'row-formatting-hidden.docx'),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const html = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    const row = html.body.querySelector('tr');
    expect(row?.dataset.changeKind).toBe('row-formatting');
    expect(row?.dataset.officeRowPropertyRevisionOmml).toBeUndefined();
    expect(parseDocumentRowFormatting(row?.dataset.changeBefore)).toEqual({
      hidden: true,
    });

    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: imported.content.html,
    });
    try {
      const change = collectDocumentChanges(editor.state.doc)[0];
      expect(change?.kind).toBe('row-formatting');
      expect(editor.commands.rejectDocumentChange(change?.id ?? '')).toBe(true);
      const rejected = new DOMParser().parseFromString(
        editor.getHTML(),
        'text/html',
      );
      const rejectedRow = rejected.body.querySelector('tr');
      expect(rejectedRow?.dataset.changeKind).toBeUndefined();
      expect(rejectedRow?.dataset.officeRowHidden).toBe('true');
    } finally {
      editor.destroy();
    }
  });

  test('pending hidden row-formatting change round-trips as native w:trPrChange', async () => {
    const source = await rowDocxWithHiddenChange({ prior: true });
    const imported = await importOfficeFile(
      new File([source], 'row-formatting-hidden-roundtrip.docx'),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const exported = await createArtifactBlob(imported);
    const archive = await JSZip.loadAsync(await exported.arrayBuffer());
    const document = await xmlEntry(archive, 'word/document.xml');
    const change = directChild(
      directChild(descendants(document, 'tr')[0], 'trPr'),
      'trPrChange',
    );
    expect(change).toBeTruthy();
    expect(directChild(directChild(change, 'trPr'), 'hidden')).toBeTruthy();
  });

  test('live row hidden edits become reviewable when track changes is on', () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions({
        isTracking: () => true,
      }),
      content:
        '<table><tbody><tr data-office-row-hidden="false"><td><p>Cell</p></td></tr></tbody></table>',
    });
    try {
      expect(
        editor.commands.setDocumentTableRowOptions({
          cantSplit: false,
          repeatHeader: false,
          hidden: true,
        }),
      ).toBe(true);
      const changes = collectDocumentChanges(editor.state.doc).filter(
        (change) => change.kind === 'row-formatting',
      );
      expect(changes).toHaveLength(1);
      const html = new DOMParser().parseFromString(
        editor.getHTML(),
        'text/html',
      );
      const row = html.body.querySelector('tr');
      expect(row?.dataset.changeKind).toBe('row-formatting');
      expect(parseDocumentRowFormatting(row?.dataset.changeBefore)).toEqual({
        cantSplit: false,
        repeatHeader: false,
        hidden: false,
        alignment: 'left',
        gridBefore: 0,
        gridAfter: 0,
        widthBefore: { type: 'auto', value: null },
        widthAfter: { type: 'auto', value: null },
      });
      expect(row?.dataset.officeRowHidden).toBe('true');
    } finally {
      editor.destroy();
    }
  });

  test('imports jc-only w:trPrChange as a reviewable row-formatting change', async () => {
    const source = await rowDocxWithAlignmentChange({ prior: 'center' });
    const imported = await importOfficeFile(
      new File([source], 'row-formatting-alignment.docx'),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const html = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    const row = html.body.querySelector('tr');
    expect(row?.dataset.changeKind).toBe('row-formatting');
    expect(row?.dataset.officeRowPropertyRevisionOmml).toBeUndefined();
    expect(parseDocumentRowFormatting(row?.dataset.changeBefore)).toEqual({
      alignment: 'center',
    });

    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: imported.content.html,
    });
    try {
      const change = collectDocumentChanges(editor.state.doc)[0];
      expect(change?.kind).toBe('row-formatting');
      expect(editor.commands.rejectDocumentChange(change?.id ?? '')).toBe(true);
      const rejected = new DOMParser().parseFromString(
        editor.getHTML(),
        'text/html',
      );
      const rejectedRow = rejected.body.querySelector('tr');
      expect(rejectedRow?.dataset.changeKind).toBeUndefined();
      expect(rejectedRow?.dataset.officeRowAlignment).toBe('center');
    } finally {
      editor.destroy();
    }
  });

  test('pending jc row-formatting change round-trips as native w:trPrChange', async () => {
    const source = await rowDocxWithAlignmentChange({ prior: 'right' });
    const imported = await importOfficeFile(
      new File([source], 'row-formatting-alignment-roundtrip.docx'),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const exported = await createArtifactBlob(imported);
    const archive = await JSZip.loadAsync(await exported.arrayBuffer());
    const document = await xmlEntry(archive, 'word/document.xml');
    const change = directChild(
      directChild(descendants(document, 'tr')[0], 'trPr'),
      'trPrChange',
    );
    expect(change).toBeTruthy();
    const priorJc = directChild(directChild(change, 'trPr'), 'jc');
    expect(
      priorJc?.getAttributeNS(WORD_NAMESPACE, 'val') ??
        priorJc?.getAttribute('w:val') ??
        priorJc?.getAttribute('val'),
    ).toBe('right');
  });

  test('live row alignment edits become reviewable when track changes is on', () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions({
        isTracking: () => true,
      }),
      content:
        '<table><tbody><tr data-office-row-alignment="left"><td><p>Cell</p></td></tr></tbody></table>',
    });
    try {
      expect(
        editor.commands.setDocumentTableRowOptions({
          cantSplit: false,
          repeatHeader: false,
          alignment: 'center',
        }),
      ).toBe(true);
      const changes = collectDocumentChanges(editor.state.doc).filter(
        (change) => change.kind === 'row-formatting',
      );
      expect(changes).toHaveLength(1);
      const html = new DOMParser().parseFromString(
        editor.getHTML(),
        'text/html',
      );
      const row = html.body.querySelector('tr');
      expect(row?.dataset.changeKind).toBe('row-formatting');
      expect(parseDocumentRowFormatting(row?.dataset.changeBefore)).toEqual({
        cantSplit: false,
        repeatHeader: false,
        hidden: false,
        alignment: 'left',
        gridBefore: 0,
        gridAfter: 0,
        widthBefore: { type: 'auto', value: null },
        widthAfter: { type: 'auto', value: null },
      });
      expect(row?.dataset.officeRowAlignment).toBe('center');
    } finally {
      editor.destroy();
    }
  });

  test('imports trHeight-only w:trPrChange as a reviewable row-formatting change', async () => {
    const source = await rowDocxWithHeightChange({
      priorValue: 480,
      priorRule: 'exact',
      currentValue: 720,
    });
    const imported = await importOfficeFile(
      new File([source], 'row-formatting-height.docx'),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const html = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    const row = html.body.querySelector('tr');
    expect(row?.dataset.changeKind).toBe('row-formatting');
    expect(parseDocumentRowFormatting(row?.dataset.changeBefore)).toEqual({
      height: { value: 32, rule: 'exact' },
    });

    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: imported.content.html,
    });
    try {
      const change = collectDocumentChanges(editor.state.doc)[0];
      expect(change?.kind).toBe('row-formatting');
      expect(editor.commands.rejectDocumentChange(change?.id ?? '')).toBe(true);
      const rejected = new DOMParser().parseFromString(
        editor.getHTML(),
        'text/html',
      );
      const rejectedRow = rejected.body.querySelector('tr');
      expect(rejectedRow?.dataset.changeKind).toBeUndefined();
      expect(rejectedRow?.dataset.officeRowHeight).toBe('32');
      expect(rejectedRow?.dataset.officeRowHeightRule).toBe('exact');
    } finally {
      editor.destroy();
    }
  });

  test('pending trHeight row-formatting change round-trips as native w:trPrChange', async () => {
    const source = await rowDocxWithHeightChange({
      priorValue: 480,
      priorRule: 'atLeast',
      currentValue: 720,
    });
    const imported = await importOfficeFile(
      new File([source], 'row-formatting-height-roundtrip.docx'),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: imported.content.html,
    });
    try {
      imported.content.html = editor.getHTML();
    } finally {
      editor.destroy();
    }
    const exported = await createArtifactBlob(imported);
    const archive = await JSZip.loadAsync(await exported.arrayBuffer());
    const document = await xmlEntry(archive, 'word/document.xml');
    const change = directChild(
      directChild(descendants(document, 'tr')[0], 'trPr'),
      'trPrChange',
    );
    expect(change).toBeTruthy();
    const priorHeight = directChild(directChild(change, 'trPr'), 'trHeight');
    expect(wordAttribute(priorHeight!, 'val')).toBe('480');
    expect(wordAttribute(priorHeight!, 'hRule')).toBe('atLeast');
  });

  test('live row height edits become reviewable when track changes is on', () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions({
        isTracking: () => true,
      }),
      content:
        '<table><tbody><tr data-office-row-height="24" data-office-row-height-rule="atLeast"><td><p>Cell</p></td></tr></tbody></table>',
    });
    try {
      expect(editor.commands.setDocumentTableRowHeight(40, 'exact')).toBe(true);
      const changes = collectDocumentChanges(editor.state.doc);
      expect(changes).toHaveLength(1);
      expect(changes[0]?.kind).toBe('row-formatting');
      const html = new DOMParser().parseFromString(
        editor.getHTML(),
        'text/html',
      );
      const row = html.body.querySelector('tr');
      expect(row?.dataset.changeKind).toBe('row-formatting');
      expect(parseDocumentRowFormatting(row?.dataset.changeBefore)).toEqual({
        cantSplit: false,
        repeatHeader: false,
        hidden: false,
        alignment: 'left',
        gridBefore: 0,
        gridAfter: 0,
        widthBefore: { type: 'auto', value: null },
        widthAfter: { type: 'auto', value: null },
        height: { value: 24, rule: 'atLeast' },
      });
      expect(row?.dataset.officeRowHeight).toBe('40');
      expect(row?.dataset.officeRowHeightRule).toBe('exact');
    } finally {
      editor.destroy();
    }
  });
  test('imports gridBefore-only w:trPrChange as a reviewable row-formatting change', async () => {
    const source = await rowDocxWithGridBeforeChange({ prior: 2, current: 1 });
    const imported = await importOfficeFile(
      new File([source], 'row-formatting-grid-before.docx'),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const html = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    const row = html.body.querySelector('tr');
    expect(row?.dataset.changeKind).toBe('row-formatting');
    expect(row?.dataset.officeRowPropertyRevisionOmml).toBeUndefined();
    expect(parseDocumentRowFormatting(row?.dataset.changeBefore)).toEqual({
      gridBefore: 2,
    });
    expect(row?.dataset.officeRowGridBefore).toBe('1');

    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: imported.content.html,
    });
    try {
      const change = collectDocumentChanges(editor.state.doc)[0];
      expect(change?.kind).toBe('row-formatting');
      expect(editor.commands.rejectDocumentChange(change?.id ?? '')).toBe(true);
      const rejected = new DOMParser().parseFromString(
        editor.getHTML(),
        'text/html',
      );
      const rejectedRow = rejected.body.querySelector('tr');
      expect(rejectedRow?.dataset.changeKind).toBeUndefined();
      expect(rejectedRow?.dataset.officeRowGridBefore).toBe('2');
    } finally {
      editor.destroy();
    }
  });

  test('pending gridBefore row-formatting change round-trips as native w:trPrChange', async () => {
    const source = await rowDocxWithGridBeforeChange({ prior: 3 });
    const imported = await importOfficeFile(
      new File([source], 'row-formatting-grid-before-roundtrip.docx'),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const exported = await xmlEntry(
      await JSZip.loadAsync(
        await (await createArtifactBlob(imported)).arrayBuffer(),
      ),
      'word/document.xml',
    );
    const change = directChild(
      directChild(descendants(exported, 'tr')[0], 'trPr'),
      'trPrChange',
    );
    expect(change).toBeTruthy();
    const priorGrid = directChild(directChild(change!, 'trPr'), 'gridBefore');
    expect(
      priorGrid?.getAttributeNS(WORD_NAMESPACE, 'val') ??
        priorGrid?.getAttribute('w:val') ??
        priorGrid?.getAttribute('val'),
    ).toBe('3');
  });

  test('imports gridAfter-only w:trPrChange as a reviewable row-formatting change', async () => {
    const source = await rowDocxWithGridAfterChange({ prior: 2, current: 1 });
    const imported = await importOfficeFile(
      new File([source], 'row-formatting-grid-after.docx'),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const html = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    const row = html.body.querySelector('tr');
    expect(row?.dataset.changeKind).toBe('row-formatting');
    expect(row?.dataset.officeRowPropertyRevisionOmml).toBeUndefined();
    expect(parseDocumentRowFormatting(row?.dataset.changeBefore)).toEqual({
      gridAfter: 2,
    });
    expect(row?.dataset.officeRowGridAfter).toBe('1');

    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: imported.content.html,
    });
    try {
      const change = collectDocumentChanges(editor.state.doc)[0];
      expect(change?.kind).toBe('row-formatting');
      expect(editor.commands.rejectDocumentChange(change?.id ?? '')).toBe(true);
      const rejected = new DOMParser().parseFromString(
        editor.getHTML(),
        'text/html',
      );
      const rejectedRow = rejected.body.querySelector('tr');
      expect(rejectedRow?.dataset.changeKind).toBeUndefined();
      expect(rejectedRow?.dataset.officeRowGridAfter).toBe('2');
    } finally {
      editor.destroy();
    }
  });

  test('pending gridAfter row-formatting change round-trips as native w:trPrChange', async () => {
    const source = await rowDocxWithGridAfterChange({ prior: 3 });
    const imported = await importOfficeFile(
      new File([source], 'row-formatting-grid-after-roundtrip.docx'),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const exported = await xmlEntry(
      await JSZip.loadAsync(
        await (await createArtifactBlob(imported)).arrayBuffer(),
      ),
      'word/document.xml',
    );
    const change = directChild(
      directChild(descendants(exported, 'tr')[0], 'trPr'),
      'trPrChange',
    );
    expect(change).toBeTruthy();
    const priorGrid = directChild(directChild(change!, 'trPr'), 'gridAfter');
    expect(
      priorGrid?.getAttributeNS(WORD_NAMESPACE, 'val') ??
        priorGrid?.getAttribute('w:val') ??
        priorGrid?.getAttribute('val'),
    ).toBe('3');
  });

  test('live row gridAfter edits become reviewable when track changes is on', () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions({
        isTracking: () => true,
      }),
      content:
        '<table><tbody><tr data-office-row-grid-after="0"><td><p>Cell</p></td></tr></tbody></table>',
    });
    try {
      let rowPos: number | null = null;
      editor.state.doc.descendants((node, position) => {
        if (node.type.name === 'tableRow' && rowPos === null) {
          rowPos = position;
        }
      });
      expect(rowPos).not.toBeNull();
      editor.view.dispatch(
        editor.state.tr.setNodeMarkup(rowPos!, undefined, {
          ...editor.state.doc.nodeAt(rowPos!)!.attrs,
          gridAfter: 2,
        }),
      );
      const changes = collectDocumentChanges(editor.state.doc).filter(
        (change) => change.kind === 'row-formatting',
      );
      expect(changes).toHaveLength(1);
      const html = new DOMParser().parseFromString(
        editor.getHTML(),
        'text/html',
      );
      const row = html.body.querySelector('tr');
      expect(row?.dataset.changeKind).toBe('row-formatting');
      expect(
        parseDocumentRowFormatting(row?.dataset.changeBefore),
      ).toMatchObject({
        gridAfter: 0,
      });
      expect(row?.dataset.officeRowGridAfter).toBe('2');
    } finally {
      editor.destroy();
    }
  });

  test('imports wBefore-only w:trPrChange as a reviewable row-formatting change', async () => {
    const source = await rowDocxWithWidthBeforeChange({
      prior: { type: 'dxa', w: '288' },
      current: { type: 'dxa', w: '144' },
    });
    const imported = await importOfficeFile(
      new File([source], 'row-formatting-w-before.docx'),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const html = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    const row = html.body.querySelector('tr');
    expect(row?.dataset.changeKind).toBe('row-formatting');
    expect(row?.dataset.officeRowPropertyRevisionOmml).toBeUndefined();
    expect(parseDocumentRowFormatting(row?.dataset.changeBefore)).toEqual({
      widthBefore: { type: 'pixels', value: 19.2 },
    });
    expect(row?.dataset.officeRowWidthBeforeType).toBe('pixels');
    expect(row?.dataset.officeRowWidthBefore).toBe('9.6');

    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: imported.content.html,
    });
    try {
      const change = collectDocumentChanges(editor.state.doc)[0];
      expect(change?.kind).toBe('row-formatting');
      expect(editor.commands.rejectDocumentChange(change?.id ?? '')).toBe(true);
      const rejected = new DOMParser().parseFromString(
        editor.getHTML(),
        'text/html',
      );
      const rejectedRow = rejected.body.querySelector('tr');
      expect(rejectedRow?.dataset.changeKind).toBeUndefined();
      expect(rejectedRow?.dataset.officeRowWidthBeforeType).toBe('pixels');
      expect(rejectedRow?.dataset.officeRowWidthBefore).toBe('19.2');
    } finally {
      editor.destroy();
    }
  });

  test('pending wBefore row-formatting change round-trips as native w:trPrChange', async () => {
    const source = await rowDocxWithWidthBeforeChange({
      prior: { type: 'dxa', w: '432' },
    });
    const imported = await importOfficeFile(
      new File([source], 'row-formatting-w-before-roundtrip.docx'),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const exported = await xmlEntry(
      await JSZip.loadAsync(
        await (await createArtifactBlob(imported)).arrayBuffer(),
      ),
      'word/document.xml',
    );
    const change = directChild(
      directChild(descendants(exported, 'tr')[0], 'trPr'),
      'trPrChange',
    );
    expect(change).toBeTruthy();
    const priorWidth = directChild(directChild(change!, 'trPr'), 'wBefore');
    expect(
      priorWidth?.getAttributeNS(WORD_NAMESPACE, 'type') ??
        priorWidth?.getAttribute('w:type') ??
        priorWidth?.getAttribute('type'),
    ).toBe('dxa');
    expect(
      priorWidth?.getAttributeNS(WORD_NAMESPACE, 'w') ??
        priorWidth?.getAttribute('w:w') ??
        priorWidth?.getAttribute('w'),
    ).toBe('432');
  });

  test('live row widthBefore edits become reviewable when track changes is on', () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions({
        isTracking: () => true,
      }),
      content:
        '<table><tbody><tr data-office-row-width-before-type="auto"><td><p>Cell</p></td></tr></tbody></table>',
    });
    try {
      let rowPos: number | null = null;
      editor.state.doc.descendants((node, position) => {
        if (node.type.name === 'tableRow' && rowPos === null) {
          rowPos = position;
        }
      });
      expect(rowPos).not.toBeNull();
      editor.view.dispatch(
        editor.state.tr.setNodeMarkup(rowPos!, undefined, {
          ...editor.state.doc.nodeAt(rowPos!)!.attrs,
          widthBefore: { type: 'pixels', value: 48 },
        }),
      );
      const changes = collectDocumentChanges(editor.state.doc).filter(
        (change) => change.kind === 'row-formatting',
      );
      expect(changes).toHaveLength(1);
      const html = new DOMParser().parseFromString(
        editor.getHTML(),
        'text/html',
      );
      const row = html.body.querySelector('tr');
      expect(row?.dataset.changeKind).toBe('row-formatting');
      expect(
        parseDocumentRowFormatting(row?.dataset.changeBefore),
      ).toMatchObject({
        widthBefore: { type: 'auto', value: null },
      });
      expect(row?.dataset.officeRowWidthBeforeType).toBe('pixels');
      expect(row?.dataset.officeRowWidthBefore).toBe('48');
    } finally {
      editor.destroy();
    }
  });

  test('imports wAfter-only w:trPrChange as a reviewable row-formatting change', async () => {
    const source = await rowDocxWithWidthAfterChange({
      prior: { type: 'dxa', w: '288' },
      current: { type: 'dxa', w: '144' },
    });
    const imported = await importOfficeFile(
      new File([source], 'row-formatting-w-after.docx'),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const html = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    const row = html.body.querySelector('tr');
    expect(row?.dataset.changeKind).toBe('row-formatting');
    expect(row?.dataset.officeRowPropertyRevisionOmml).toBeUndefined();
    expect(parseDocumentRowFormatting(row?.dataset.changeBefore)).toEqual({
      widthAfter: { type: 'pixels', value: 19.2 },
    });
    expect(row?.dataset.officeRowWidthAfterType).toBe('pixels');
    expect(row?.dataset.officeRowWidthAfter).toBe('9.6');

    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: imported.content.html,
    });
    try {
      const change = collectDocumentChanges(editor.state.doc)[0];
      expect(change?.kind).toBe('row-formatting');
      expect(editor.commands.rejectDocumentChange(change?.id ?? '')).toBe(true);
      const rejected = new DOMParser().parseFromString(
        editor.getHTML(),
        'text/html',
      );
      const rejectedRow = rejected.body.querySelector('tr');
      expect(rejectedRow?.dataset.changeKind).toBeUndefined();
      expect(rejectedRow?.dataset.officeRowWidthAfterType).toBe('pixels');
      expect(rejectedRow?.dataset.officeRowWidthAfter).toBe('19.2');
    } finally {
      editor.destroy();
    }
  });

  test('pending wAfter row-formatting change round-trips as native w:trPrChange', async () => {
    const source = await rowDocxWithWidthAfterChange({
      prior: { type: 'dxa', w: '432' },
    });
    const imported = await importOfficeFile(
      new File([source], 'row-formatting-w-after-roundtrip.docx'),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const exported = await xmlEntry(
      await JSZip.loadAsync(
        await (await createArtifactBlob(imported)).arrayBuffer(),
      ),
      'word/document.xml',
    );
    const change = directChild(
      directChild(descendants(exported, 'tr')[0], 'trPr'),
      'trPrChange',
    );
    expect(change).toBeTruthy();
    const priorWidth = directChild(directChild(change!, 'trPr'), 'wAfter');
    expect(
      priorWidth?.getAttributeNS(WORD_NAMESPACE, 'type') ??
        priorWidth?.getAttribute('w:type') ??
        priorWidth?.getAttribute('type'),
    ).toBe('dxa');
    expect(
      priorWidth?.getAttributeNS(WORD_NAMESPACE, 'w') ??
        priorWidth?.getAttribute('w:w') ??
        priorWidth?.getAttribute('w'),
    ).toBe('432');
  });

  test('live row widthAfter edits become reviewable when track changes is on', () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions({
        isTracking: () => true,
      }),
      content:
        '<table><tbody><tr data-office-row-width-after-type="auto"><td><p>Cell</p></td></tr></tbody></table>',
    });
    try {
      let rowPos: number | null = null;
      editor.state.doc.descendants((node, position) => {
        if (node.type.name === 'tableRow' && rowPos === null) {
          rowPos = position;
        }
      });
      expect(rowPos).not.toBeNull();
      editor.view.dispatch(
        editor.state.tr.setNodeMarkup(rowPos!, undefined, {
          ...editor.state.doc.nodeAt(rowPos!)!.attrs,
          widthAfter: { type: 'pixels', value: 48 },
        }),
      );
      const changes = collectDocumentChanges(editor.state.doc).filter(
        (change) => change.kind === 'row-formatting',
      );
      expect(changes).toHaveLength(1);
      const html = new DOMParser().parseFromString(
        editor.getHTML(),
        'text/html',
      );
      const row = html.body.querySelector('tr');
      expect(row?.dataset.changeKind).toBe('row-formatting');
      expect(
        parseDocumentRowFormatting(row?.dataset.changeBefore),
      ).toMatchObject({
        widthAfter: { type: 'auto', value: null },
      });
      expect(row?.dataset.officeRowWidthAfterType).toBe('pixels');
      expect(row?.dataset.officeRowWidthAfter).toBe('48');
    } finally {
      editor.destroy();
    }
  });

  test('live row gridBefore edits become reviewable when track changes is on', () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions({
        isTracking: () => true,
      }),
      content:
        '<table><tbody><tr data-office-row-grid-before="0"><td><p>Cell</p></td></tr></tbody></table>',
    });
    try {
      let rowPos: number | null = null;
      editor.state.doc.descendants((node, position) => {
        if (node.type.name === 'tableRow' && rowPos === null) {
          rowPos = position;
        }
      });
      expect(rowPos).not.toBeNull();
      editor.view.dispatch(
        editor.state.tr.setNodeMarkup(rowPos!, undefined, {
          ...editor.state.doc.nodeAt(rowPos!)!.attrs,
          gridBefore: 2,
        }),
      );
      const changes = collectDocumentChanges(editor.state.doc).filter(
        (change) => change.kind === 'row-formatting',
      );
      expect(changes).toHaveLength(1);
      const html = new DOMParser().parseFromString(
        editor.getHTML(),
        'text/html',
      );
      const row = html.body.querySelector('tr');
      expect(row?.dataset.changeKind).toBe('row-formatting');
      expect(
        parseDocumentRowFormatting(row?.dataset.changeBefore),
      ).toMatchObject({
        gridBefore: 0,
      });
      expect(row?.dataset.officeRowGridBefore).toBe('2');
    } finally {
      editor.destroy();
    }
  });
});

async function rowDocxWithCantSplitChange(options: {
  prior: boolean;
}): Promise<ArrayBuffer> {
  const artifact = createArtifact('blank-document');
  if (artifact.content.type !== 'document') {
    throw new Error('Expected a document artifact.');
  }
  artifact.content.html =
    '<table><tbody><tr><td><p>Cell</p></td></tr></tbody></table>';
  const seed = await createArtifactBlob(artifact);
  const archive = await JSZip.loadAsync(await seed.arrayBuffer());
  const document = await xmlEntry(archive, 'word/document.xml');
  const row = descendants(document, 'tr')[0];
  const properties =
    directChild(row, 'trPr') ??
    (() => {
      const created = document.createElementNS(WORD_NAMESPACE, 'w:trPr');
      row.insertBefore(created, row.firstChild);
      return created;
    })();
  for (const existing of Array.from(properties.children).filter(
    (child) =>
      child.localName === 'cantSplit' || child.localName === 'trPrChange',
  )) {
    existing.remove();
  }
  const current = document.createElementNS(WORD_NAMESPACE, 'w:cantSplit');
  properties.append(current);
  const change = document.createElementNS(WORD_NAMESPACE, 'w:trPrChange');
  change.setAttributeNS(WORD_NAMESPACE, 'w:id', '21');
  change.setAttributeNS(WORD_NAMESPACE, 'w:author', 'Reviewer');
  change.setAttributeNS(WORD_NAMESPACE, 'w:date', '2026-09-08T00:00:00Z');
  const prior = document.createElementNS(WORD_NAMESPACE, 'w:trPr');
  const priorCantSplit = document.createElementNS(
    WORD_NAMESPACE,
    'w:cantSplit',
  );
  if (!options.prior) {
    priorCantSplit.setAttributeNS(WORD_NAMESPACE, 'w:val', '0');
  }
  prior.append(priorCantSplit);
  change.append(prior);
  properties.append(change);
  archive.file(
    'word/document.xml',
    new XMLSerializer().serializeToString(document),
  );
  return archive.generateAsync({ type: 'arraybuffer' });
}

async function rowDocxWithHeaderChange(options: {
  prior: boolean;
}): Promise<ArrayBuffer> {
  const artifact = createArtifact('blank-document');
  if (artifact.content.type !== 'document') {
    throw new Error('Expected a document artifact.');
  }
  artifact.content.html =
    '<table><tbody><tr><td><p>Cell</p></td></tr></tbody></table>';
  const seed = await createArtifactBlob(artifact);
  const archive = await JSZip.loadAsync(await seed.arrayBuffer());
  const document = await xmlEntry(archive, 'word/document.xml');
  const row = descendants(document, 'tr')[0];
  const properties =
    directChild(row, 'trPr') ??
    (() => {
      const created = document.createElementNS(WORD_NAMESPACE, 'w:trPr');
      row.insertBefore(created, row.firstChild);
      return created;
    })();
  for (const existing of Array.from(properties.children).filter(
    (child) =>
      child.localName === 'tblHeader' || child.localName === 'trPrChange',
  )) {
    existing.remove();
  }
  const current = document.createElementNS(WORD_NAMESPACE, 'w:tblHeader');
  properties.append(current);
  const change = document.createElementNS(WORD_NAMESPACE, 'w:trPrChange');
  change.setAttributeNS(WORD_NAMESPACE, 'w:id', '22');
  change.setAttributeNS(WORD_NAMESPACE, 'w:author', 'Reviewer');
  change.setAttributeNS(WORD_NAMESPACE, 'w:date', '2026-09-08T00:00:00Z');
  const prior = document.createElementNS(WORD_NAMESPACE, 'w:trPr');
  const priorHeader = document.createElementNS(WORD_NAMESPACE, 'w:tblHeader');
  if (!options.prior) {
    priorHeader.setAttributeNS(WORD_NAMESPACE, 'w:val', '0');
  }
  prior.append(priorHeader);
  change.append(prior);
  properties.append(change);
  archive.file(
    'word/document.xml',
    new XMLSerializer().serializeToString(document),
  );
  return archive.generateAsync({ type: 'arraybuffer' });
}

async function rowDocxWithHeightChange(options: {
  priorValue: number;
  priorRule?: 'exact' | 'atLeast';
  currentValue: number;
  currentRule?: 'exact' | 'atLeast';
}): Promise<ArrayBuffer> {
  const artifact = createArtifact('blank-document');
  if (artifact.content.type !== 'document') {
    throw new Error('Expected a document artifact.');
  }
  artifact.content.html =
    '<table><tbody><tr><td><p>Cell</p></td></tr></tbody></table>';
  const seed = await createArtifactBlob(artifact);
  const archive = await JSZip.loadAsync(await seed.arrayBuffer());
  const document = await xmlEntry(archive, 'word/document.xml');
  const row = descendants(document, 'tr')[0];
  const properties =
    directChild(row, 'trPr') ??
    (() => {
      const created = document.createElementNS(WORD_NAMESPACE, 'w:trPr');
      row.insertBefore(created, row.firstChild);
      return created;
    })();
  for (const existing of Array.from(properties.children).filter(
    (child) =>
      child.localName === 'trHeight' || child.localName === 'trPrChange',
  )) {
    existing.remove();
  }
  const current = document.createElementNS(WORD_NAMESPACE, 'w:trHeight');
  current.setAttributeNS(WORD_NAMESPACE, 'w:val', String(options.currentValue));
  current.setAttributeNS(
    WORD_NAMESPACE,
    'w:hRule',
    options.currentRule ?? 'atLeast',
  );
  properties.append(current);
  const change = document.createElementNS(WORD_NAMESPACE, 'w:trPrChange');
  change.setAttributeNS(WORD_NAMESPACE, 'w:id', '23');
  change.setAttributeNS(WORD_NAMESPACE, 'w:author', 'Reviewer');
  change.setAttributeNS(WORD_NAMESPACE, 'w:date', '2026-09-08T00:00:00Z');
  const prior = document.createElementNS(WORD_NAMESPACE, 'w:trPr');
  const priorHeight = document.createElementNS(WORD_NAMESPACE, 'w:trHeight');
  priorHeight.setAttributeNS(
    WORD_NAMESPACE,
    'w:val',
    String(options.priorValue),
  );
  priorHeight.setAttributeNS(
    WORD_NAMESPACE,
    'w:hRule',
    options.priorRule ?? 'atLeast',
  );
  prior.append(priorHeight);
  change.append(prior);
  properties.append(change);
  archive.file(
    'word/document.xml',
    new XMLSerializer().serializeToString(document),
  );
  return archive.generateAsync({ type: 'arraybuffer' });
}

async function rowDocxWithHiddenChange(options: {
  prior: boolean;
}): Promise<ArrayBuffer> {
  const artifact = createArtifact('blank-document');
  if (artifact.content.type !== 'document') {
    throw new Error('Expected a document artifact.');
  }
  artifact.content.html =
    '<table><tbody><tr><td><p>Cell</p></td></tr></tbody></table>';
  const seed = await createArtifactBlob(artifact);
  const archive = await JSZip.loadAsync(await seed.arrayBuffer());
  const document = await xmlEntry(archive, 'word/document.xml');
  const row = descendants(document, 'tr')[0];
  const properties =
    directChild(row, 'trPr') ??
    (() => {
      const created = document.createElementNS(WORD_NAMESPACE, 'w:trPr');
      row.insertBefore(created, row.firstChild);
      return created;
    })();
  for (const existing of Array.from(properties.children).filter(
    (child) => child.localName === 'hidden' || child.localName === 'trPrChange',
  )) {
    existing.remove();
  }
  const change = document.createElementNS(WORD_NAMESPACE, 'w:trPrChange');
  change.setAttributeNS(WORD_NAMESPACE, 'w:id', '27');
  change.setAttributeNS(WORD_NAMESPACE, 'w:author', 'Reviewer');
  change.setAttributeNS(WORD_NAMESPACE, 'w:date', '2026-09-08T00:00:00Z');
  const prior = document.createElementNS(WORD_NAMESPACE, 'w:trPr');
  const priorHidden = document.createElementNS(WORD_NAMESPACE, 'w:hidden');
  if (!options.prior) {
    priorHidden.setAttributeNS(WORD_NAMESPACE, 'w:val', '0');
  }
  prior.append(priorHidden);
  change.append(prior);
  properties.append(change);
  archive.file(
    'word/document.xml',
    new XMLSerializer().serializeToString(document),
  );
  return archive.generateAsync({ type: 'arraybuffer' });
}

async function rowDocxWithAlignmentChange(options: {
  prior: 'left' | 'center' | 'right';
  current?: 'left' | 'center' | 'right';
}): Promise<ArrayBuffer> {
  const artifact = createArtifact('blank-document');
  if (artifact.content.type !== 'document') {
    throw new Error('Expected a document artifact.');
  }
  artifact.content.html =
    '<table><tbody><tr><td><p>Cell</p></td></tr></tbody></table>';
  const seed = await createArtifactBlob(artifact);
  const archive = await JSZip.loadAsync(await seed.arrayBuffer());
  const document = await xmlEntry(archive, 'word/document.xml');
  const row = descendants(document, 'tr')[0];
  const properties =
    directChild(row, 'trPr') ??
    (() => {
      const created = document.createElementNS(WORD_NAMESPACE, 'w:trPr');
      row.insertBefore(created, row.firstChild);
      return created;
    })();
  for (const existing of Array.from(properties.children).filter(
    (child) => child.localName === 'jc' || child.localName === 'trPrChange',
  )) {
    existing.remove();
  }
  const current = document.createElementNS(WORD_NAMESPACE, 'w:jc');
  current.setAttributeNS(WORD_NAMESPACE, 'w:val', options.current ?? 'left');
  properties.append(current);
  const change = document.createElementNS(WORD_NAMESPACE, 'w:trPrChange');
  change.setAttributeNS(WORD_NAMESPACE, 'w:id', '28');
  change.setAttributeNS(WORD_NAMESPACE, 'w:author', 'Reviewer');
  change.setAttributeNS(WORD_NAMESPACE, 'w:date', '2026-09-08T00:00:00Z');
  const prior = document.createElementNS(WORD_NAMESPACE, 'w:trPr');
  const priorJc = document.createElementNS(WORD_NAMESPACE, 'w:jc');
  priorJc.setAttributeNS(WORD_NAMESPACE, 'w:val', options.prior);
  prior.append(priorJc);
  change.append(prior);
  properties.append(change);
  archive.file(
    'word/document.xml',
    new XMLSerializer().serializeToString(document),
  );
  return archive.generateAsync({ type: 'arraybuffer' });
}

async function rowDocxWithGridAfterChange(options: {
  prior: number;
  current?: number;
}): Promise<ArrayBuffer> {
  const artifact = createArtifact('blank-document');
  if (artifact.content.type !== 'document') {
    throw new Error('Expected a document artifact.');
  }
  artifact.content.html =
    '<table><tbody><tr><td><p>Cell</p></td></tr></tbody></table>';
  const seed = await createArtifactBlob(artifact);
  const archive = await JSZip.loadAsync(await seed.arrayBuffer());
  const document = await xmlEntry(archive, 'word/document.xml');
  const row = descendants(document, 'tr')[0];
  const properties =
    directChild(row, 'trPr') ??
    (() => {
      const created = document.createElementNS(WORD_NAMESPACE, 'w:trPr');
      row.insertBefore(created, row.firstChild);
      return created;
    })();
  for (const existing of Array.from(properties.children).filter(
    (child) =>
      child.localName === 'gridAfter' || child.localName === 'trPrChange',
  )) {
    existing.remove();
  }
  const current = document.createElementNS(WORD_NAMESPACE, 'w:gridAfter');
  current.setAttributeNS(
    WORD_NAMESPACE,
    'w:val',
    String(options.current ?? options.prior),
  );
  properties.append(current);
  const change = document.createElementNS(WORD_NAMESPACE, 'w:trPrChange');
  change.setAttributeNS(WORD_NAMESPACE, 'w:id', '30');
  change.setAttributeNS(WORD_NAMESPACE, 'w:author', 'Reviewer');
  change.setAttributeNS(WORD_NAMESPACE, 'w:date', '2026-09-08T00:00:00Z');
  const prior = document.createElementNS(WORD_NAMESPACE, 'w:trPr');
  const priorGrid = document.createElementNS(WORD_NAMESPACE, 'w:gridAfter');
  priorGrid.setAttributeNS(WORD_NAMESPACE, 'w:val', String(options.prior));
  prior.append(priorGrid);
  change.append(prior);
  properties.append(change);
  archive.file(
    'word/document.xml',
    new XMLSerializer().serializeToString(document),
  );
  return archive.generateAsync({ type: 'arraybuffer' });
}

async function rowDocxWithWidthAfterChange(options: {
  prior: { type: string; w: string };
  current?: { type: string; w: string };
}): Promise<ArrayBuffer> {
  const artifact = createArtifact('blank-document');
  if (artifact.content.type !== 'document') {
    throw new Error('Expected a document artifact.');
  }
  artifact.content.html =
    '<table><tbody><tr><td><p>Cell</p></td></tr></tbody></table>';
  const seed = await createArtifactBlob(artifact);
  const archive = await JSZip.loadAsync(await seed.arrayBuffer());
  const document = await xmlEntry(archive, 'word/document.xml');
  const row = descendants(document, 'tr')[0];
  const properties =
    directChild(row, 'trPr') ??
    (() => {
      const created = document.createElementNS(WORD_NAMESPACE, 'w:trPr');
      row.insertBefore(created, row.firstChild);
      return created;
    })();
  for (const existing of Array.from(properties.children).filter(
    (child) => child.localName === 'wAfter' || child.localName === 'trPrChange',
  )) {
    existing.remove();
  }
  const currentSpec = options.current ?? options.prior;
  const current = document.createElementNS(WORD_NAMESPACE, 'w:wAfter');
  current.setAttributeNS(WORD_NAMESPACE, 'w:type', currentSpec.type);
  current.setAttributeNS(WORD_NAMESPACE, 'w:w', currentSpec.w);
  properties.append(current);
  const change = document.createElementNS(WORD_NAMESPACE, 'w:trPrChange');
  change.setAttributeNS(WORD_NAMESPACE, 'w:id', '32');
  change.setAttributeNS(WORD_NAMESPACE, 'w:author', 'Reviewer');
  change.setAttributeNS(WORD_NAMESPACE, 'w:date', '2026-09-08T00:00:00Z');
  const prior = document.createElementNS(WORD_NAMESPACE, 'w:trPr');
  const priorWidth = document.createElementNS(WORD_NAMESPACE, 'w:wAfter');
  priorWidth.setAttributeNS(WORD_NAMESPACE, 'w:type', options.prior.type);
  priorWidth.setAttributeNS(WORD_NAMESPACE, 'w:w', options.prior.w);
  prior.append(priorWidth);
  change.append(prior);
  properties.append(change);
  archive.file(
    'word/document.xml',
    new XMLSerializer().serializeToString(document),
  );
  return archive.generateAsync({ type: 'arraybuffer' });
}

async function rowDocxWithWidthBeforeChange(options: {
  prior: { type: string; w: string };
  current?: { type: string; w: string };
}): Promise<ArrayBuffer> {
  const artifact = createArtifact('blank-document');
  if (artifact.content.type !== 'document') {
    throw new Error('Expected a document artifact.');
  }
  artifact.content.html =
    '<table><tbody><tr><td><p>Cell</p></td></tr></tbody></table>';
  const seed = await createArtifactBlob(artifact);
  const archive = await JSZip.loadAsync(await seed.arrayBuffer());
  const document = await xmlEntry(archive, 'word/document.xml');
  const row = descendants(document, 'tr')[0];
  const properties =
    directChild(row, 'trPr') ??
    (() => {
      const created = document.createElementNS(WORD_NAMESPACE, 'w:trPr');
      row.insertBefore(created, row.firstChild);
      return created;
    })();
  for (const existing of Array.from(properties.children).filter(
    (child) =>
      child.localName === 'wBefore' || child.localName === 'trPrChange',
  )) {
    existing.remove();
  }
  const currentSpec = options.current ?? options.prior;
  const current = document.createElementNS(WORD_NAMESPACE, 'w:wBefore');
  current.setAttributeNS(WORD_NAMESPACE, 'w:type', currentSpec.type);
  current.setAttributeNS(WORD_NAMESPACE, 'w:w', currentSpec.w);
  properties.append(current);
  const change = document.createElementNS(WORD_NAMESPACE, 'w:trPrChange');
  change.setAttributeNS(WORD_NAMESPACE, 'w:id', '31');
  change.setAttributeNS(WORD_NAMESPACE, 'w:author', 'Reviewer');
  change.setAttributeNS(WORD_NAMESPACE, 'w:date', '2026-09-08T00:00:00Z');
  const prior = document.createElementNS(WORD_NAMESPACE, 'w:trPr');
  const priorWidth = document.createElementNS(WORD_NAMESPACE, 'w:wBefore');
  priorWidth.setAttributeNS(WORD_NAMESPACE, 'w:type', options.prior.type);
  priorWidth.setAttributeNS(WORD_NAMESPACE, 'w:w', options.prior.w);
  prior.append(priorWidth);
  change.append(prior);
  properties.append(change);
  archive.file(
    'word/document.xml',
    new XMLSerializer().serializeToString(document),
  );
  return archive.generateAsync({ type: 'arraybuffer' });
}

async function rowDocxWithGridBeforeChange(options: {
  prior: number;
  current?: number;
}): Promise<ArrayBuffer> {
  const artifact = createArtifact('blank-document');
  if (artifact.content.type !== 'document') {
    throw new Error('Expected a document artifact.');
  }
  artifact.content.html =
    '<table><tbody><tr><td><p>Cell</p></td></tr></tbody></table>';
  const seed = await createArtifactBlob(artifact);
  const archive = await JSZip.loadAsync(await seed.arrayBuffer());
  const document = await xmlEntry(archive, 'word/document.xml');
  const row = descendants(document, 'tr')[0];
  const properties =
    directChild(row, 'trPr') ??
    (() => {
      const created = document.createElementNS(WORD_NAMESPACE, 'w:trPr');
      row.insertBefore(created, row.firstChild);
      return created;
    })();
  for (const existing of Array.from(properties.children).filter(
    (child) =>
      child.localName === 'gridBefore' || child.localName === 'trPrChange',
  )) {
    existing.remove();
  }
  const current = document.createElementNS(WORD_NAMESPACE, 'w:gridBefore');
  current.setAttributeNS(
    WORD_NAMESPACE,
    'w:val',
    String(options.current ?? options.prior),
  );
  properties.append(current);
  const change = document.createElementNS(WORD_NAMESPACE, 'w:trPrChange');
  change.setAttributeNS(WORD_NAMESPACE, 'w:id', '29');
  change.setAttributeNS(WORD_NAMESPACE, 'w:author', 'Reviewer');
  change.setAttributeNS(WORD_NAMESPACE, 'w:date', '2026-09-08T00:00:00Z');
  const prior = document.createElementNS(WORD_NAMESPACE, 'w:trPr');
  const priorGrid = document.createElementNS(WORD_NAMESPACE, 'w:gridBefore');
  priorGrid.setAttributeNS(WORD_NAMESPACE, 'w:val', String(options.prior));
  prior.append(priorGrid);
  change.append(prior);
  properties.append(change);
  archive.file(
    'word/document.xml',
    new XMLSerializer().serializeToString(document),
  );
  return archive.generateAsync({ type: 'arraybuffer' });
}

async function xmlEntry(archive: JSZip, path: string): Promise<Document> {
  const entry = archive.file(path);
  if (!entry) throw new Error(`Missing ${path}`);
  return parseXml(await entry.async('string'), path);
}

function wordAttribute(element: Element, localName: string): string | null {
  const matches = Array.from(element.attributes).filter(
    (candidate) =>
      xmlAttributeLocalName(candidate) === localName &&
      xmlAttributeNamespace(element, candidate) === element.namespaceURI,
  );
  return matches.length === 1 ? (matches[0]?.value ?? null) : null;
}
