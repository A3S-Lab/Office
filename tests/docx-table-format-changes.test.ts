import { Editor } from '@tiptap/core';
import { describe, expect, test } from '@rstest/core';
import JSZip from 'jszip';
import {
  createArtifact,
  createArtifactBlob,
  importOfficeFile,
} from '../src/core';
import {
  collectDocumentChanges,
} from '../src/internal/features/work/work-document-changes';
import { createWorkDocumentExtensions } from '../src/internal/features/work/work-document-extensions';
import { parseDocumentTableFormatting } from '../src/internal/features/work/work-document-table-format-changes';
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

describe('DOCX table-formatting revisions', () => {
  test('imports jc-only w:tblPrChange as a reviewable table-formatting change', async () => {
    const source = await tableDocxWithAlignmentChange({
      current: 'right',
      prior: 'center',
    });
    const imported = await importOfficeFile(
      new File([source], 'table-formatting.docx'),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const html = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    const table = html.body.querySelector('table');
    expect(table?.dataset.changeKind).toBe('table-formatting');
    expect(table?.dataset.changeAuthor).toBe('Reviewer');
    expect(table?.dataset.officeTablePropertyRevisionOmml).toBeUndefined();
    expect(parseDocumentTableFormatting(table?.dataset.changeBefore)).toEqual({
      alignment: 'center',
    });

    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: imported.content.html,
    });
    try {
      const changes = collectDocumentChanges(editor.state.doc);
      expect(changes).toHaveLength(1);
      expect(changes[0]).toMatchObject({
        kind: 'table-formatting',
        author: 'Reviewer',
      });
      expect(editor.commands.acceptDocumentChange(changes[0]?.id ?? '')).toBe(
        true,
      );
      expect(collectDocumentChanges(editor.state.doc)).toHaveLength(0);
      const accepted = new DOMParser().parseFromString(
        editor.getHTML(),
        'text/html',
      );
      expect(
        accepted.body.querySelector('table')?.dataset.changeKind,
      ).toBeUndefined();
    } finally {
      editor.destroy();
    }
  });

  test('reject restores prior table alignment and drops the pending change', async () => {
    const source = await tableDocxWithAlignmentChange({
      current: 'right',
      prior: 'left',
    });
    const imported = await importOfficeFile(
      new File([source], 'table-formatting-reject.docx'),
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
      expect(change?.kind).toBe('table-formatting');
      expect(editor.commands.rejectDocumentChange(change?.id ?? '')).toBe(true);
      expect(collectDocumentChanges(editor.state.doc)).toHaveLength(0);
      const html = new DOMParser().parseFromString(
        editor.getHTML(),
        'text/html',
      );
      const table = html.body.querySelector('table');
      expect(table?.dataset.changeKind).toBeUndefined();
      expect(table?.dataset.officeTableAlignment).toBe('left');
    } finally {
      editor.destroy();
    }
  });

  test('pending table-formatting change round-trips as native w:tblPrChange', async () => {
    const source = await tableDocxWithAlignmentChange({
      current: 'center',
      prior: 'left',
    });
    const imported = await importOfficeFile(
      new File([source], 'table-formatting-roundtrip.docx'),
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
    const table = descendants(document, 'tbl')[0];
    const properties = directChild(table, 'tblPr');
    const change = directChild(properties, 'tblPrChange');
    expect(change).toBeTruthy();
    expect(wordAttribute(change!, 'author')).toBe('Reviewer');
    const prior = directChild(change, 'tblPr');
    const jc = directChild(prior, 'jc');
    expect(wordAttribute(jc!, 'val')).toBe('left');
  });

  test('reports alignment-only table-property revisions as reviewable diagnostics', async () => {
    const source = await tableDocxWithAlignmentChange({
      current: 'right',
      prior: 'center',
    });
    const summary = await analyzeDocxCompatibility(
      new File([source], 'table-formatting-diag.docx'),
      [],
    );
    expect(
      summary.issues.some(
        (issue) => issue.code === 'docx.revisions.table-formatting',
      ),
    ).toBe(true);
  });

  test('live table alignment edits become reviewable when track changes is on', () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions({
        isTracking: () => true,
      }),
      content:
        '<table data-office-table-alignment="left"><tbody><tr><td><p>Cell</p></td></tr></tbody></table>',
    });
    try {
      expect(editor.commands.setDocumentTableAlignment('center')).toBe(true);
      const changes = collectDocumentChanges(editor.state.doc);
      expect(changes).toHaveLength(1);
      expect(changes[0]?.kind).toBe('table-formatting');
      const html = new DOMParser().parseFromString(
        editor.getHTML(),
        'text/html',
      );
      const table = html.body.querySelector('table');
      expect(table?.dataset.changeKind).toBe('table-formatting');
      const before = parseDocumentTableFormatting(table?.dataset.changeBefore);
      expect(before?.alignment).toBe('left');
      expect(before?.width).toBeTruthy();
      expect(table?.dataset.officeTableAlignment).toBe('center');
    } finally {
      editor.destroy();
    }
  });

  test('imports preferred-width-only w:tblPrChange as a reviewable table-formatting change', async () => {
    const source = await tableDocxWithWidthChange({
      currentType: 'pct',
      currentW: '5000',
      priorType: 'pct',
      priorW: '2500',
    });
    const imported = await importOfficeFile(
      new File([source], 'table-width-formatting.docx'),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const html = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    const table = html.body.querySelector('table');
    expect(table?.dataset.changeKind).toBe('table-formatting');
    expect(table?.dataset.officeTablePropertyRevisionOmml).toBeUndefined();
    expect(parseDocumentTableFormatting(table?.dataset.changeBefore)).toEqual({
      width: { type: 'percent', value: 50 },
    });

    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: imported.content.html,
    });
    try {
      const change = collectDocumentChanges(editor.state.doc)[0];
      expect(change?.kind).toBe('table-formatting');
      expect(editor.commands.rejectDocumentChange(change?.id ?? '')).toBe(true);
      const rejected = new DOMParser().parseFromString(
        editor.getHTML(),
        'text/html',
      );
      const rejectedTable = rejected.body.querySelector('table');
      expect(rejectedTable?.dataset.changeKind).toBeUndefined();
      expect(rejectedTable?.dataset.officeTableWidthType).toBe('percent');
      expect(rejectedTable?.dataset.officeTableWidth).toBe('50');
    } finally {
      editor.destroy();
    }
  });

  test('imports jc+tblW prior snapshot as one reviewable table-formatting change', async () => {
    const source = await tableDocxWithAlignmentAndWidthChange();
    const imported = await importOfficeFile(
      new File([source], 'table-jc-width-formatting.docx'),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const html = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    const table = html.body.querySelector('table');
    expect(parseDocumentTableFormatting(table?.dataset.changeBefore)).toEqual({
      alignment: 'left',
      width: { type: 'auto', value: null },
    });
  });

  test('imports tblInd-only w:tblPrChange as a reviewable table-formatting change', async () => {
    const source = await tableDocxWithIndentChange({
      currentTwips: '1440',
      priorTwips: '720',
    });
    const imported = await importOfficeFile(
      new File([source], 'table-indent-formatting.docx'),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const html = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    const table = html.body.querySelector('table');
    expect(table?.dataset.changeKind).toBe('table-formatting');
    expect(table?.dataset.officeTablePropertyRevisionOmml).toBeUndefined();
    expect(parseDocumentTableFormatting(table?.dataset.changeBefore)).toEqual({
      indent: 48,
    });

    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: imported.content.html,
    });
    try {
      const change = collectDocumentChanges(editor.state.doc)[0];
      expect(change?.kind).toBe('table-formatting');
      expect(editor.commands.rejectDocumentChange(change?.id ?? '')).toBe(true);
      const rejected = new DOMParser().parseFromString(
        editor.getHTML(),
        'text/html',
      );
      const rejectedTable = rejected.body.querySelector('table');
      expect(rejectedTable?.dataset.changeKind).toBeUndefined();
      expect(rejectedTable?.dataset.officeTableIndent).toBe('48');
    } finally {
      editor.destroy();
    }
  });

  test('live table indent edits become reviewable when track changes is on', () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions({
        isTracking: () => true,
      }),
      content:
        '<table data-office-table-indent="12"><tbody><tr><td><p>Cell</p></td></tr></tbody></table>',
    });
    try {
      expect(
        editor.commands.setDocumentTableProperties({
          width: { type: 'auto', value: null },
          alignment: 'left',
          indent: 48,
        }),
      ).toBe(true);
      const changes = collectDocumentChanges(editor.state.doc);
      expect(changes).toHaveLength(1);
      expect(changes[0]?.kind).toBe('table-formatting');
      const html = new DOMParser().parseFromString(
        editor.getHTML(),
        'text/html',
      );
      const table = html.body.querySelector('table');
      expect(table?.dataset.changeKind).toBe('table-formatting');
      const before = parseDocumentTableFormatting(table?.dataset.changeBefore);
      expect(before?.indent).toBe(12);
      expect(table?.dataset.officeTableIndent).toBe('48');
    } finally {
      editor.destroy();
    }
  });

  test('imports tblCellMar-only w:tblPrChange as a reviewable table-formatting change', async () => {
    const source = await tableDocxWithCellMarginChange({
      priorTwips: 720,
      currentTwips: 1440,
    });
    const imported = await importOfficeFile(
      new File([source], 'table-cell-margin-formatting.docx'),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const html = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    const table = html.body.querySelector('table');
    expect(table?.dataset.changeKind).toBe('table-formatting');
    expect(table?.dataset.officeTablePropertyRevisionOmml).toBeUndefined();
    expect(parseDocumentTableFormatting(table?.dataset.changeBefore)).toEqual({
      cellMargins: { top: 48, right: 48, bottom: 48, left: 48 },
    });

    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: imported.content.html,
    });
    try {
      const change = collectDocumentChanges(editor.state.doc)[0];
      expect(change?.kind).toBe('table-formatting');
      expect(editor.commands.rejectDocumentChange(change?.id ?? '')).toBe(true);
      const rejected = new DOMParser().parseFromString(
        editor.getHTML(),
        'text/html',
      );
      const rejectedTable = rejected.body.querySelector('table');
      expect(rejectedTable?.dataset.changeKind).toBeUndefined();
      expect(rejectedTable?.dataset.officeTableCellMarginTop).toBe('48');
      expect(rejectedTable?.dataset.officeTableCellMarginLeft).toBe('48');
    } finally {
      editor.destroy();
    }
  });

  test('live table cell-margin edits become reviewable when track changes is on', () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions({
        isTracking: () => true,
      }),
      content: [
        '<table data-office-table-cell-margin-top="12"',
        ' data-office-table-cell-margin-right="12"',
        ' data-office-table-cell-margin-bottom="12"',
        ' data-office-table-cell-margin-left="12">',
        '<tbody><tr><td><p>Cell</p></td></tr></tbody></table>',
      ].join(''),
    });
    try {
      expect(
        editor.commands.setDocumentTableCellMargins({
          top: 48,
          right: 48,
          bottom: 48,
          left: 48,
        }),
      ).toBe(true);
      const changes = collectDocumentChanges(editor.state.doc);
      expect(changes).toHaveLength(1);
      expect(changes[0]?.kind).toBe('table-formatting');
      const html = new DOMParser().parseFromString(
        editor.getHTML(),
        'text/html',
      );
      const table = html.body.querySelector('table');
      expect(table?.dataset.changeKind).toBe('table-formatting');
      const before = parseDocumentTableFormatting(table?.dataset.changeBefore);
      expect(before?.cellMargins).toEqual({
        top: 12,
        right: 12,
        bottom: 12,
        left: 12,
      });
      expect(table?.dataset.officeTableCellMarginTop).toBe('48');
    } finally {
      editor.destroy();
    }
  });

  test('imports tblLayout-only w:tblPrChange as a reviewable table-formatting change', async () => {
    const source = await tableDocxWithLayoutChange({
      prior: 'fixed',
      current: 'autofit',
    });
    const imported = await importOfficeFile(
      new File([source], 'table-layout-formatting.docx'),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const html = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    const table = html.body.querySelector('table');
    expect(table?.dataset.changeKind).toBe('table-formatting');
    expect(table?.dataset.officeTablePropertyRevisionOmml).toBeUndefined();
    expect(parseDocumentTableFormatting(table?.dataset.changeBefore)).toEqual({
      layout: 'fixed',
    });

    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: imported.content.html,
    });
    try {
      const change = collectDocumentChanges(editor.state.doc)[0];
      expect(change?.kind).toBe('table-formatting');
      expect(editor.commands.rejectDocumentChange(change?.id ?? '')).toBe(true);
      const rejected = new DOMParser().parseFromString(
        editor.getHTML(),
        'text/html',
      );
      const rejectedTable = rejected.body.querySelector('table');
      expect(rejectedTable?.dataset.changeKind).toBeUndefined();
      expect(rejectedTable?.dataset.officeTableLayout).toBe('fixed');
    } finally {
      editor.destroy();
    }
  });



  test('imports solid-shd-only w:tblPrChange as a reviewable table-formatting change', async () => {
    const source = await tableDocxWithSolidFillChange({
      prior: '#ffcc00',
      current: null,
    });
    const imported = await importOfficeFile(
      new File([source], 'table-fill-formatting.docx'),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const html = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    const table = html.body.querySelector('table');
    expect(table?.dataset.changeKind).toBe('table-formatting');
    expect(table?.dataset.officeTablePropertyRevisionOmml).toBeUndefined();
    expect(parseDocumentTableFormatting(table?.dataset.changeBefore)).toEqual({
      fill: '#ffcc00',
    });
  });

  test('pending solid-shd table-formatting change round-trips as native w:tblPrChange', async () => {
    const source = await tableDocxWithSolidFillChange({
      prior: '#ffcc00',
      current: null,
    });
    const imported = await importOfficeFile(
      new File([source], 'table-fill-roundtrip.docx'),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const output = await JSZip.loadAsync(
      await (await createArtifactBlob(imported)).arrayBuffer(),
    );
    const exported = await xmlEntry(output, 'word/document.xml');
    const properties = directChild(descendants(exported, 'tbl')[0], 'tblPr');
    const change = directChild(properties, 'tblPrChange');
    expect(change).toBeTruthy();
    expect(directChild(directChild(change!, 'tblPr'), 'shd')).toBeTruthy();
  });

  test('live table fill edits become reviewable when track changes is on', () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions({
        isTracking: () => true,
      }),
      content: [
        '<table data-office-table-layout="autofit"',
        ' data-office-table-width-type="percent"',
        ' data-office-table-width="100"',
        ' data-office-table-bidi-visual="false">',
        '<tbody><tr><td><p>Cell</p></td></tr></tbody></table>',
      ].join(''),
    });
    try {
      let tablePos: number | null = null;
      editor.state.doc.descendants((node, position) => {
        if (node.type.name === 'table' && tablePos === null) {
          tablePos = position;
        }
      });
      expect(tablePos).not.toBeNull();
      editor.view.dispatch(
        editor.state.tr.setNodeMarkup(tablePos!, undefined, {
          ...editor.state.doc.nodeAt(tablePos!)!.attrs,
          fill: '#ffcc00',
        }),
      );
      const changes = collectDocumentChanges(editor.state.doc).filter(
        (change) => change.kind === 'table-formatting',
      );
      expect(changes).toHaveLength(1);
      const html = new DOMParser().parseFromString(
        editor.getHTML(),
        'text/html',
      );
      const table = html.body.querySelector('table');
      expect(table?.dataset.changeKind).toBe('table-formatting');
      expect(
        parseDocumentTableFormatting(table?.dataset.changeBefore),
      ).toMatchObject({ bidiVisual: false });
      expect(
        parseDocumentTableFormatting(table?.dataset.changeBefore)?.fill,
      ).toBeUndefined();
      expect(table?.dataset.officeTableFill).toBe('#ffcc00');
    } finally {
      editor.destroy();
    }
  });

  test('imports bidiVisual-only w:tblPrChange as a reviewable table-formatting change', async () => {
    const source = await tableDocxWithBidiVisualChange({
      prior: true,
      current: false,
    });
    const imported = await importOfficeFile(
      new File([source], 'table-bidivisual-formatting.docx'),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const html = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    const table = html.body.querySelector('table');
    expect(table?.dataset.changeKind).toBe('table-formatting');
    expect(table?.dataset.officeTablePropertyRevisionOmml).toBeUndefined();
    expect(parseDocumentTableFormatting(table?.dataset.changeBefore)).toEqual({
      bidiVisual: true,
    });
  });

  test('pending bidiVisual table-formatting change round-trips as native w:tblPrChange', async () => {
    const source = await tableDocxWithBidiVisualChange({
      prior: true,
      current: false,
    });
    const imported = await importOfficeFile(
      new File([source], 'table-bidivisual-roundtrip.docx'),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const output = await JSZip.loadAsync(
      await (await createArtifactBlob(imported)).arrayBuffer(),
    );
    const exported = await xmlEntry(output, 'word/document.xml');
    const properties = directChild(descendants(exported, 'tbl')[0], 'tblPr');
    const change = directChild(properties, 'tblPrChange');
    expect(change).toBeTruthy();
    expect(directChild(directChild(change!, 'tblPr'), 'bidiVisual')).toBeTruthy();
  });

  test('live table bidiVisual edits become reviewable when track changes is on', () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions({
        isTracking: () => true,
      }),
      content: [
        '<table data-office-table-layout="autofit"',
        ' data-office-table-width-type="percent"',
        ' data-office-table-width="100"',
        ' data-office-table-bidi-visual="false">',
        '<tbody><tr><td><p>Cell</p></td></tr></tbody></table>',
      ].join(''),
    });
    try {
      let tablePos: number | null = null;
      editor.state.doc.descendants((node, position) => {
        if (node.type.name === 'table' && tablePos === null) {
          tablePos = position;
        }
      });
      expect(tablePos).not.toBeNull();
      editor.view.dispatch(
        editor.state.tr.setNodeMarkup(tablePos!, undefined, {
          ...editor.state.doc.nodeAt(tablePos!)!.attrs,
          bidiVisual: true,
        }),
      );
      const changes = collectDocumentChanges(editor.state.doc).filter(
        (change) => change.kind === 'table-formatting',
      );
      expect(changes).toHaveLength(1);
      const html = new DOMParser().parseFromString(
        editor.getHTML(),
        'text/html',
      );
      const table = html.body.querySelector('table');
      expect(table?.dataset.changeKind).toBe('table-formatting');
      expect(parseDocumentTableFormatting(table?.dataset.changeBefore)).toMatchObject(
        {
          bidiVisual: false,
        },
      );
      expect(table?.dataset.officeTableBidiVisual).toBe('true');
    } finally {
      editor.destroy();
    }
  });

  test('imports tblLook-only w:tblPrChange as a reviewable table-formatting change', async () => {
    const source = await tableDocxWithLookChange({
      prior: {
        firstRow: true,
        lastRow: false,
        firstColumn: true,
        lastColumn: false,
        noHorizontalBand: false,
        noVerticalBand: true,
      },
      current: {
        firstRow: true,
        lastRow: true,
        firstColumn: true,
        lastColumn: false,
        noHorizontalBand: false,
        noVerticalBand: true,
      },
    });
    const imported = await importOfficeFile(
      new File([source], 'table-look-formatting.docx'),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const html = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    const table = html.body.querySelector('table');
    expect(table?.dataset.changeKind).toBe('table-formatting');
    expect(table?.dataset.officeTablePropertyRevisionOmml).toBeUndefined();
    expect(parseDocumentTableFormatting(table?.dataset.changeBefore)).toEqual({
      look: {
        firstRow: true,
        lastRow: false,
        firstColumn: true,
        lastColumn: false,
        noHorizontalBand: false,
        noVerticalBand: true,
      },
    });
    expect(JSON.parse(table?.dataset.officeTableLook ?? 'null')).toEqual({
      firstRow: true,
      lastRow: true,
      firstColumn: true,
      lastColumn: false,
      noHorizontalBand: false,
      noVerticalBand: true,
    });
  });

  test('pending tblLook table-formatting change round-trips as native w:tblPrChange', async () => {
    const source = await tableDocxWithLookChange({
      prior: {
        firstRow: true,
        lastRow: false,
        firstColumn: true,
        lastColumn: false,
        noHorizontalBand: false,
        noVerticalBand: true,
      },
      current: {
        firstRow: false,
        lastRow: false,
        firstColumn: true,
        lastColumn: true,
        noHorizontalBand: true,
        noVerticalBand: true,
      },
    });
    const imported = await importOfficeFile(
      new File([source], 'table-look-roundtrip.docx'),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const output = await JSZip.loadAsync(
      await (await createArtifactBlob(imported)).arrayBuffer(),
    );
    const exported = await xmlEntry(output, 'word/document.xml');
    const properties = directChild(descendants(exported, 'tbl')[0], 'tblPr');
    const change = directChild(properties, 'tblPrChange');
    expect(change).toBeTruthy();
    expect(directChild(directChild(change!, 'tblPr'), 'tblLook')).toBeTruthy();
    expect(directChild(properties, 'tblLook')).toBeTruthy();
  });

  test('live table look edits become reviewable when track changes is on', () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions({
        isTracking: () => true,
      }),
      content: [
        '<table data-office-table-layout="autofit"',
        ' data-office-table-width-type="percent"',
        ' data-office-table-width="100"',
        ' data-office-table-bidi-visual="false"',
        ' data-office-table-look=\'{"firstRow":true,"lastRow":false,"firstColumn":true,"lastColumn":false,"noHorizontalBand":false,"noVerticalBand":true}\'>',
        '<tbody><tr><td><p>Cell</p></td></tr></tbody></table>',
      ].join(''),
    });
    try {
      let tablePos: number | null = null;
      editor.state.doc.descendants((node, position) => {
        if (node.type.name === 'table' && tablePos === null) {
          tablePos = position;
        }
      });
      expect(tablePos).not.toBeNull();
      editor.view.dispatch(
        editor.state.tr.setNodeMarkup(tablePos!, undefined, {
          ...editor.state.doc.nodeAt(tablePos!)!.attrs,
          look: {
            firstRow: true,
            lastRow: true,
            firstColumn: true,
            lastColumn: false,
            noHorizontalBand: false,
            noVerticalBand: true,
          },
        }),
      );
      const changes = collectDocumentChanges(editor.state.doc).filter(
        (change) => change.kind === 'table-formatting',
      );
      expect(changes).toHaveLength(1);
      const html = new DOMParser().parseFromString(
        editor.getHTML(),
        'text/html',
      );
      const table = html.body.querySelector('table');
      expect(table?.dataset.changeKind).toBe('table-formatting');
      expect(parseDocumentTableFormatting(table?.dataset.changeBefore)?.look).toEqual(
        {
          firstRow: true,
          lastRow: false,
          firstColumn: true,
          lastColumn: false,
          noHorizontalBand: false,
          noVerticalBand: true,
        },
      );
      expect(JSON.parse(table?.dataset.officeTableLook ?? 'null')).toEqual({
        firstRow: true,
        lastRow: true,
        firstColumn: true,
        lastColumn: false,
        noHorizontalBand: false,
        noVerticalBand: true,
      });
    } finally {
      editor.destroy();
    }
  });


  test('imports tblOverlap-only w:tblPrChange as a reviewable table-formatting change', async () => {
    const source = await tableDocxWithOverlapChange({
      prior: 'never',
      current: 'overlap',
    });
    const imported = await importOfficeFile(
      new File([source], 'table-overlap-formatting.docx'),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const html = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    const table = html.body.querySelector('table');
    expect(table?.dataset.changeKind).toBe('table-formatting');
    expect(table?.dataset.officeTablePropertyRevisionOmml).toBeUndefined();
    expect(parseDocumentTableFormatting(table?.dataset.changeBefore)).toEqual({
      overlap: 'never',
    });
    expect(table?.dataset.officeTableOverlap).toBe('overlap');
  });

  test('pending tblOverlap table-formatting change round-trips as native w:tblPrChange', async () => {
    const source = await tableDocxWithOverlapChange({
      prior: 'never',
      current: 'overlap',
    });
    const imported = await importOfficeFile(
      new File([source], 'table-overlap-roundtrip.docx'),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const output = await JSZip.loadAsync(
      await (await createArtifactBlob(imported)).arrayBuffer(),
    );
    const exported = await xmlEntry(output, 'word/document.xml');
    const properties = directChild(descendants(exported, 'tbl')[0], 'tblPr');
    const change = directChild(properties, 'tblPrChange');
    expect(change).toBeTruthy();
    expect(directChild(directChild(change!, 'tblPr'), 'tblOverlap')).toBeTruthy();
    expect(directChild(properties, 'tblOverlap')).toBeTruthy();
  });

  test('live table overlap edits become reviewable when track changes is on', () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions({
        isTracking: () => true,
      }),
      content: [
        '<table data-office-table-layout="autofit"',
        ' data-office-table-width-type="percent"',
        ' data-office-table-width="100"',
        ' data-office-table-bidi-visual="false"',
        ' data-office-table-overlap="never">',
        '<tbody><tr><td><p>Cell</p></td></tr></tbody></table>',
      ].join(''),
    });
    try {
      let tablePos: number | null = null;
      editor.state.doc.descendants((node, position) => {
        if (node.type.name === 'table' && tablePos === null) {
          tablePos = position;
        }
      });
      expect(tablePos).not.toBeNull();
      editor.view.dispatch(
        editor.state.tr.setNodeMarkup(tablePos!, undefined, {
          ...editor.state.doc.nodeAt(tablePos!)!.attrs,
          overlap: 'overlap',
        }),
      );
      const changes = collectDocumentChanges(editor.state.doc).filter(
        (change) => change.kind === 'table-formatting',
      );
      expect(changes).toHaveLength(1);
      const html = new DOMParser().parseFromString(
        editor.getHTML(),
        'text/html',
      );
      const table = html.body.querySelector('table');
      expect(table?.dataset.changeKind).toBe('table-formatting');
      expect(parseDocumentTableFormatting(table?.dataset.changeBefore)?.overlap).toBe(
        'never',
      );
      expect(table?.dataset.officeTableOverlap).toBe('overlap');
    } finally {
      editor.destroy();
    }
  });

  test('imports tblStyle-only w:tblPrChange as a reviewable table-formatting change', async () => {
    const source = await tableDocxWithStyleIdChange({
      prior: 'TableNormal',
      current: 'TableGrid',
    });
    const imported = await importOfficeFile(
      new File([source], 'table-style-formatting.docx'),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const html = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    const table = html.body.querySelector('table');
    expect(table?.dataset.changeKind).toBe('table-formatting');
    expect(table?.dataset.officeTablePropertyRevisionOmml).toBeUndefined();
    expect(parseDocumentTableFormatting(table?.dataset.changeBefore)).toEqual({
      styleId: 'TableNormal',
    });
    expect(table?.dataset.officeTableStyleId).toBe('TableGrid');
  });

  test('pending tblStyle table-formatting change round-trips as native w:tblPrChange', async () => {
    const source = await tableDocxWithStyleIdChange({
      prior: 'TableNormal',
      current: 'TableGrid',
    });
    const imported = await importOfficeFile(
      new File([source], 'table-style-roundtrip.docx'),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const output = await JSZip.loadAsync(
      await (await createArtifactBlob(imported)).arrayBuffer(),
    );
    const exported = await xmlEntry(output, 'word/document.xml');
    const properties = directChild(descendants(exported, 'tbl')[0], 'tblPr');
    const change = directChild(properties, 'tblPrChange');
    expect(change).toBeTruthy();
    expect(directChild(directChild(change!, 'tblPr'), 'tblStyle')).toBeTruthy();
    expect(directChild(properties, 'tblStyle')).toBeTruthy();
  });

  test('live table styleId edits become reviewable when track changes is on', () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions({
        isTracking: () => true,
      }),
      content: [
        '<table data-office-table-layout="autofit"',
        ' data-office-table-width-type="percent"',
        ' data-office-table-width="100"',
        ' data-office-table-bidi-visual="false"',
        ' data-office-table-style-id="TableNormal">',
        '<tbody><tr><td><p>Cell</p></td></tr></tbody></table>',
      ].join(''),
    });
    try {
      let tablePos: number | null = null;
      editor.state.doc.descendants((node, position) => {
        if (node.type.name === 'table' && tablePos === null) {
          tablePos = position;
        }
      });
      expect(tablePos).not.toBeNull();
      editor.view.dispatch(
        editor.state.tr.setNodeMarkup(tablePos!, undefined, {
          ...editor.state.doc.nodeAt(tablePos!)!.attrs,
          styleId: 'TableGrid',
        }),
      );
      const changes = collectDocumentChanges(editor.state.doc).filter(
        (change) => change.kind === 'table-formatting',
      );
      expect(changes).toHaveLength(1);
      const html = new DOMParser().parseFromString(
        editor.getHTML(),
        'text/html',
      );
      const table = html.body.querySelector('table');
      expect(table?.dataset.changeKind).toBe('table-formatting');
      expect(parseDocumentTableFormatting(table?.dataset.changeBefore)?.styleId).toBe(
        'TableNormal',
      );
      expect(table?.dataset.officeTableStyleId).toBe('TableGrid');
    } finally {
      editor.destroy();
    }
  });

  test('imports tblCellSpacing-only w:tblPrChange as a reviewable table-formatting change', async () => {
    const source = await tableDocxWithCellSpacingChange({
      priorTwips: 120,
      currentTwips: 240,
    });
    const imported = await importOfficeFile(
      new File([source], 'table-cell-spacing-formatting.docx'),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const html = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    const table = html.body.querySelector('table');
    expect(table?.dataset.changeKind).toBe('table-formatting');
    expect(table?.dataset.officeTablePropertyRevisionOmml).toBeUndefined();
    expect(parseDocumentTableFormatting(table?.dataset.changeBefore)).toEqual({
      cellSpacing: 8,
    });
    expect(table?.dataset.officeTableCellSpacing).toBe('16');
  });

  test('pending tblCellSpacing table-formatting change round-trips as native w:tblPrChange', async () => {
    const source = await tableDocxWithCellSpacingChange({
      priorTwips: 120,
      currentTwips: 240,
    });
    const imported = await importOfficeFile(
      new File([source], 'table-cell-spacing-roundtrip.docx'),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const output = await JSZip.loadAsync(
      await (await createArtifactBlob(imported)).arrayBuffer(),
    );
    const exported = await xmlEntry(output, 'word/document.xml');
    const properties = directChild(descendants(exported, 'tbl')[0], 'tblPr');
    const change = directChild(properties, 'tblPrChange');
    expect(change).toBeTruthy();
    expect(directChild(directChild(change!, 'tblPr'), 'tblCellSpacing')).toBeTruthy();
    expect(directChild(properties, 'tblCellSpacing')).toBeTruthy();
  });

  test('live table cellSpacing edits become reviewable when track changes is on', () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions({
        isTracking: () => true,
      }),
      content: [
        '<table data-office-table-layout="autofit"',
        ' data-office-table-width-type="percent"',
        ' data-office-table-width="100"',
        ' data-office-table-bidi-visual="false"',
        ' data-office-table-cell-spacing="8">',
        '<tbody><tr><td><p>Cell</p></td></tr></tbody></table>',
      ].join(''),
    });
    try {
      let tablePos: number | null = null;
      editor.state.doc.descendants((node, position) => {
        if (node.type.name === 'table' && tablePos === null) {
          tablePos = position;
        }
      });
      expect(tablePos).not.toBeNull();
      editor.view.dispatch(
        editor.state.tr.setNodeMarkup(tablePos!, undefined, {
          ...editor.state.doc.nodeAt(tablePos!)!.attrs,
          cellSpacing: 16,
        }),
      );
      const changes = collectDocumentChanges(editor.state.doc).filter(
        (change) => change.kind === 'table-formatting',
      );
      expect(changes).toHaveLength(1);
      const html = new DOMParser().parseFromString(
        editor.getHTML(),
        'text/html',
      );
      const table = html.body.querySelector('table');
      expect(table?.dataset.changeKind).toBe('table-formatting');
      expect(
        parseDocumentTableFormatting(table?.dataset.changeBefore)?.cellSpacing,
      ).toBe(8);
      expect(table?.dataset.officeTableCellSpacing).toBe('16');
    } finally {
      editor.destroy();
    }
  });

  test('imports tblBorders-only w:tblPrChange as a reviewable table-formatting change', async () => {
    const source = await tableDocxWithBordersChange({
      prior: {
        top: { val: 'single', sz: '12', color: 'FF0000' },
      },
      current: {
        top: { val: 'double', sz: '24', color: '0000FF' },
      },
    });
    const imported = await importOfficeFile(
      new File([source], 'table-borders-formatting.docx'),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const html = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    const table = html.body.querySelector('table');
    expect(table?.dataset.changeKind).toBe('table-formatting');
    expect(table?.dataset.officeTablePropertyRevisionOmml).toBeUndefined();
    expect(parseDocumentTableFormatting(table?.dataset.changeBefore)).toEqual({
      borders: {
        top: { color: '#ff0000', style: 'solid', width: 2 },
      },
    });
    expect(table?.dataset.officeTableBorders).toContain('#0000ff');
  });

  test('pending tblBorders table-formatting change round-trips as native w:tblPrChange', async () => {
    const source = await tableDocxWithBordersChange({
      prior: {
        top: { val: 'single', sz: '12', color: 'FF0000' },
      },
      current: {
        top: { val: 'double', sz: '24', color: '0000FF' },
      },
    });
    const imported = await importOfficeFile(
      new File([source], 'table-borders-roundtrip.docx'),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const output = await JSZip.loadAsync(
      await (await createArtifactBlob(imported)).arrayBuffer(),
    );
    const exported = await xmlEntry(output, 'word/document.xml');
    const properties = directChild(descendants(exported, 'tbl')[0], 'tblPr');
    const change = directChild(properties, 'tblPrChange');
    expect(change).toBeTruthy();
    expect(
      directChild(directChild(change!, 'tblPr'), 'tblBorders'),
    ).toBeTruthy();
    expect(directChild(properties, 'tblBorders')).toBeTruthy();
  });

  test('live table borders edits become reviewable when track changes is on', () => {
    const priorBorders = JSON.stringify({
      top: { color: '#ff0000', style: 'solid', width: 1 },
    });
    const editor = new Editor({
      extensions: createWorkDocumentExtensions({
        isTracking: () => true,
      }),
      content: [
        '<table data-office-table-layout="autofit"',
        ' data-office-table-width-type="percent"',
        ' data-office-table-width="100"',
        ' data-office-table-bidi-visual="false"',
        ` data-office-table-borders='${priorBorders}'>`,
        '<tbody><tr><td><p>Cell</p></td></tr></tbody></table>',
      ].join(''),
    });
    try {
      let tablePos: number | null = null;
      editor.state.doc.descendants((node, position) => {
        if (node.type.name === 'table' && tablePos === null) {
          tablePos = position;
        }
      });
      expect(tablePos).not.toBeNull();
      editor.view.dispatch(
        editor.state.tr.setNodeMarkup(tablePos!, undefined, {
          ...editor.state.doc.nodeAt(tablePos!)!.attrs,
          borders: {
            top: { color: '#0000ff', style: 'double', width: 2 },
          },
        }),
      );
      const changes = collectDocumentChanges(editor.state.doc).filter(
        (change) => change.kind === 'table-formatting',
      );
      expect(changes).toHaveLength(1);
      const html = new DOMParser().parseFromString(
        editor.getHTML(),
        'text/html',
      );
      const table = html.body.querySelector('table');
      expect(table?.dataset.changeKind).toBe('table-formatting');
      expect(
        parseDocumentTableFormatting(table?.dataset.changeBefore)?.borders,
      ).toEqual({
        top: { color: '#ff0000', style: 'solid', width: 1 },
      });
      expect(table?.dataset.officeTableBorders).toContain('#0000ff');
    } finally {
      editor.destroy();
    }
  });

  test('live table layout-mode edits become reviewable when track changes is on', () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions({
        isTracking: () => true,
      }),
      content: [
        '<table data-office-table-layout="autofit"',
        ' data-office-table-width-type="percent"',
        ' data-office-table-width="100">',
        '<tbody><tr><td><p>Cell</p></td></tr></tbody></table>',
      ].join(''),
    });
    try {
      expect(editor.commands.setDocumentTableLayoutMode('fixed', 400)).toBe(
        true,
      );
      const changes = collectDocumentChanges(editor.state.doc).filter(
        (change) => change.kind === 'table-formatting',
      );
      expect(changes).toHaveLength(1);
      expect(changes[0]?.kind).toBe('table-formatting');
      const html = new DOMParser().parseFromString(
        editor.getHTML(),
        'text/html',
      );
      const table = html.body.querySelector('table');
      expect(table?.dataset.changeKind).toBe('table-formatting');
      const before = parseDocumentTableFormatting(table?.dataset.changeBefore);
      expect(before?.layout).toBe('autofit');
      expect(table?.dataset.officeTableLayout).toBe('fixed');
    } finally {
      editor.destroy();
    }
  });
});


async function tableDocxWithSolidFillChange(options: {
  prior: string | null;
  current: string | null;
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
  const table = descendants(document, 'tbl')[0];
  const properties =
    directChild(table, 'tblPr') ??
    (() => {
      const created = document.createElementNS(WORD_NAMESPACE, 'w:tblPr');
      table.insertBefore(created, table.firstChild);
      return created;
    })();
  for (const existing of Array.from(properties.children).filter(
    (child) => child.localName === 'shd' || child.localName === 'tblPrChange',
  )) {
    existing.remove();
  }
  if (options.current) {
    const shading = document.createElementNS(WORD_NAMESPACE, 'w:shd');
    shading.setAttributeNS(WORD_NAMESPACE, 'w:val', 'clear');
    shading.setAttributeNS(
      WORD_NAMESPACE,
      'w:fill',
      options.current.replace(/^#/, '').toUpperCase(),
    );
    properties.append(shading);
  }
  const change = document.createElementNS(WORD_NAMESPACE, 'w:tblPrChange');
  change.setAttributeNS(WORD_NAMESPACE, 'w:id', '41');
  change.setAttributeNS(WORD_NAMESPACE, 'w:author', 'Reviewer');
  change.setAttributeNS(WORD_NAMESPACE, 'w:date', '2026-09-09T00:00:00Z');
  const prior = document.createElementNS(WORD_NAMESPACE, 'w:tblPr');
  if (options.prior) {
    const shading = document.createElementNS(WORD_NAMESPACE, 'w:shd');
    shading.setAttributeNS(WORD_NAMESPACE, 'w:val', 'clear');
    shading.setAttributeNS(
      WORD_NAMESPACE,
      'w:fill',
      options.prior.replace(/^#/, '').toUpperCase(),
    );
    prior.append(shading);
  }
  change.append(prior);
  properties.append(change);
  archive.file(
    'word/document.xml',
    new XMLSerializer().serializeToString(document),
  );
  return archive.generateAsync({ type: 'arraybuffer' });
}


async function tableDocxWithOverlapChange(options: {
  prior: 'never' | 'overlap';
  current: 'never' | 'overlap' | null;
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
  const table = descendants(document, 'tbl')[0];
  const properties =
    directChild(table, 'tblPr') ??
    (() => {
      const created = document.createElementNS(WORD_NAMESPACE, 'w:tblPr');
      table.insertBefore(created, table.firstChild);
      return created;
    })();
  for (const existing of Array.from(properties.children).filter(
    (child) =>
      child.localName === 'tblOverlap' || child.localName === 'tblPrChange',
  )) {
    existing.remove();
  }
  if (options.current) {
    const current = document.createElementNS(WORD_NAMESPACE, 'w:tblOverlap');
    current.setAttributeNS(WORD_NAMESPACE, 'w:val', options.current);
    properties.append(current);
  }
  const change = document.createElementNS(WORD_NAMESPACE, 'w:tblPrChange');
  change.setAttributeNS(WORD_NAMESPACE, 'w:id', '61');
  change.setAttributeNS(WORD_NAMESPACE, 'w:author', 'Reviewer');
  change.setAttributeNS(WORD_NAMESPACE, 'w:date', '2026-09-09T00:00:00Z');
  const prior = document.createElementNS(WORD_NAMESPACE, 'w:tblPr');
  const priorOverlap = document.createElementNS(WORD_NAMESPACE, 'w:tblOverlap');
  priorOverlap.setAttributeNS(WORD_NAMESPACE, 'w:val', options.prior);
  prior.append(priorOverlap);
  change.append(prior);
  properties.append(change);
  archive.file(
    'word/document.xml',
    new XMLSerializer().serializeToString(document),
  );
  return archive.generateAsync({ type: 'arraybuffer' });
}

async function tableDocxWithStyleIdChange(options: {
  prior: string;
  current: string | null;
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
  const table = descendants(document, 'tbl')[0];
  const properties =
    directChild(table, 'tblPr') ??
    (() => {
      const created = document.createElementNS(WORD_NAMESPACE, 'w:tblPr');
      table.insertBefore(created, table.firstChild);
      return created;
    })();
  for (const existing of Array.from(properties.children).filter(
    (child) =>
      child.localName === 'tblStyle' || child.localName === 'tblPrChange',
  )) {
    existing.remove();
  }
  if (options.current) {
    const current = document.createElementNS(WORD_NAMESPACE, 'w:tblStyle');
    current.setAttributeNS(WORD_NAMESPACE, 'w:val', options.current);
    properties.append(current);
  }
  const change = document.createElementNS(WORD_NAMESPACE, 'w:tblPrChange');
  change.setAttributeNS(WORD_NAMESPACE, 'w:id', '62');
  change.setAttributeNS(WORD_NAMESPACE, 'w:author', 'Reviewer');
  change.setAttributeNS(WORD_NAMESPACE, 'w:date', '2026-09-09T00:00:00Z');
  const prior = document.createElementNS(WORD_NAMESPACE, 'w:tblPr');
  const priorStyle = document.createElementNS(WORD_NAMESPACE, 'w:tblStyle');
  priorStyle.setAttributeNS(WORD_NAMESPACE, 'w:val', options.prior);
  prior.append(priorStyle);
  change.append(prior);
  properties.append(change);
  archive.file(
    'word/document.xml',
    new XMLSerializer().serializeToString(document),
  );
  return archive.generateAsync({ type: 'arraybuffer' });
}

async function tableDocxWithCellSpacingChange(options: {
  priorTwips: number;
  currentTwips: number | null;
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
  const table = descendants(document, 'tbl')[0];
  const properties =
    directChild(table, 'tblPr') ??
    (() => {
      const created = document.createElementNS(WORD_NAMESPACE, 'w:tblPr');
      table.insertBefore(created, table.firstChild);
      return created;
    })();
  for (const existing of Array.from(properties.children).filter(
    (child) =>
      child.localName === 'tblCellSpacing' ||
      child.localName === 'tblPrChange',
  )) {
    existing.remove();
  }
  if (options.currentTwips !== null) {
    const current = document.createElementNS(WORD_NAMESPACE, 'w:tblCellSpacing');
    current.setAttributeNS(WORD_NAMESPACE, 'w:type', 'dxa');
    current.setAttributeNS(WORD_NAMESPACE, 'w:w', String(options.currentTwips));
    properties.append(current);
  }
  const change = document.createElementNS(WORD_NAMESPACE, 'w:tblPrChange');
  change.setAttributeNS(WORD_NAMESPACE, 'w:id', '63');
  change.setAttributeNS(WORD_NAMESPACE, 'w:author', 'Reviewer');
  change.setAttributeNS(WORD_NAMESPACE, 'w:date', '2026-09-09T00:00:00Z');
  const prior = document.createElementNS(WORD_NAMESPACE, 'w:tblPr');
  const priorSpacing = document.createElementNS(WORD_NAMESPACE, 'w:tblCellSpacing');
  priorSpacing.setAttributeNS(WORD_NAMESPACE, 'w:type', 'dxa');
  priorSpacing.setAttributeNS(WORD_NAMESPACE, 'w:w', String(options.priorTwips));
  prior.append(priorSpacing);
  change.append(prior);
  properties.append(change);
  archive.file(
    'word/document.xml',
    new XMLSerializer().serializeToString(document),
  );
  return archive.generateAsync({ type: 'arraybuffer' });
}

async function tableDocxWithBordersChange(options: {
  prior: Record<string, { val: string; sz: string; color: string }>;
  current: Record<string, { val: string; sz: string; color: string }> | null;
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
  const table = descendants(document, 'tbl')[0];
  const properties =
    directChild(table, 'tblPr') ??
    (() => {
      const created = document.createElementNS(WORD_NAMESPACE, 'w:tblPr');
      table.insertBefore(created, table.firstChild);
      return created;
    })();
  for (const existing of Array.from(properties.children).filter(
    (child) =>
      child.localName === 'tblBorders' || child.localName === 'tblPrChange',
  )) {
    existing.remove();
  }
  if (options.current) {
    properties.append(createBordersElement(document, options.current));
  }
  const change = document.createElementNS(WORD_NAMESPACE, 'w:tblPrChange');
  change.setAttributeNS(WORD_NAMESPACE, 'w:id', '64');
  change.setAttributeNS(WORD_NAMESPACE, 'w:author', 'Reviewer');
  change.setAttributeNS(WORD_NAMESPACE, 'w:date', '2026-09-09T00:00:00Z');
  const prior = document.createElementNS(WORD_NAMESPACE, 'w:tblPr');
  prior.append(createBordersElement(document, options.prior));
  change.append(prior);
  properties.append(change);
  archive.file(
    'word/document.xml',
    new XMLSerializer().serializeToString(document),
  );
  return archive.generateAsync({ type: 'arraybuffer' });
}

function createBordersElement(
  document: Document,
  edges: Record<string, { val: string; sz: string; color: string }>,
): Element {
  const borders = document.createElementNS(WORD_NAMESPACE, 'w:tblBorders');
  for (const [name, edge] of Object.entries(edges)) {
    const child = document.createElementNS(WORD_NAMESPACE, `w:${name}`);
    child.setAttributeNS(WORD_NAMESPACE, 'w:val', edge.val);
    child.setAttributeNS(WORD_NAMESPACE, 'w:sz', edge.sz);
    child.setAttributeNS(WORD_NAMESPACE, 'w:space', '0');
    child.setAttributeNS(WORD_NAMESPACE, 'w:color', edge.color);
    borders.append(child);
  }
  return borders;
}

async function tableDocxWithLookChange(options: {
  prior: {
    firstRow: boolean;
    lastRow: boolean;
    firstColumn: boolean;
    lastColumn: boolean;
    noHorizontalBand: boolean;
    noVerticalBand: boolean;
  };
  current: {
    firstRow: boolean;
    lastRow: boolean;
    firstColumn: boolean;
    lastColumn: boolean;
    noHorizontalBand: boolean;
    noVerticalBand: boolean;
  } | null;
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
  const table = descendants(document, 'tbl')[0];
  const properties =
    directChild(table, 'tblPr') ??
    (() => {
      const created = document.createElementNS(WORD_NAMESPACE, 'w:tblPr');
      table.insertBefore(created, table.firstChild);
      return created;
    })();
  for (const existing of Array.from(properties.children).filter(
    (child) =>
      child.localName === 'tblLook' || child.localName === 'tblPrChange',
  )) {
    existing.remove();
  }
  if (options.current) {
    properties.append(createLookElement(document, options.current));
  }
  const change = document.createElementNS(WORD_NAMESPACE, 'w:tblPrChange');
  change.setAttributeNS(WORD_NAMESPACE, 'w:id', '51');
  change.setAttributeNS(WORD_NAMESPACE, 'w:author', 'Reviewer');
  change.setAttributeNS(WORD_NAMESPACE, 'w:date', '2026-09-09T00:00:00Z');
  const prior = document.createElementNS(WORD_NAMESPACE, 'w:tblPr');
  prior.append(createLookElement(document, options.prior));
  change.append(prior);
  properties.append(change);
  archive.file(
    'word/document.xml',
    new XMLSerializer().serializeToString(document),
  );
  return archive.generateAsync({ type: 'arraybuffer' });
}

function createLookElement(
  document: Document,
  look: {
    firstRow: boolean;
    lastRow: boolean;
    firstColumn: boolean;
    lastColumn: boolean;
    noHorizontalBand: boolean;
    noVerticalBand: boolean;
  },
): Element {
  const element = document.createElementNS(WORD_NAMESPACE, 'w:tblLook');
  let value = 0;
  if (look.firstRow) value |= 0x0020;
  if (look.lastRow) value |= 0x0040;
  if (look.firstColumn) value |= 0x0080;
  if (look.lastColumn) value |= 0x0100;
  if (look.noHorizontalBand) value |= 0x0200;
  if (look.noVerticalBand) value |= 0x0400;
  element.setAttributeNS(
    WORD_NAMESPACE,
    'w:val',
    value.toString(16).toUpperCase().padStart(4, '0'),
  );
  element.setAttributeNS(
    WORD_NAMESPACE,
    'w:firstRow',
    look.firstRow ? '1' : '0',
  );
  element.setAttributeNS(WORD_NAMESPACE, 'w:lastRow', look.lastRow ? '1' : '0');
  element.setAttributeNS(
    WORD_NAMESPACE,
    'w:firstColumn',
    look.firstColumn ? '1' : '0',
  );
  element.setAttributeNS(
    WORD_NAMESPACE,
    'w:lastColumn',
    look.lastColumn ? '1' : '0',
  );
  element.setAttributeNS(
    WORD_NAMESPACE,
    'w:noHBand',
    look.noHorizontalBand ? '1' : '0',
  );
  element.setAttributeNS(
    WORD_NAMESPACE,
    'w:noVBand',
    look.noVerticalBand ? '1' : '0',
  );
  return element;
}

async function tableDocxWithBidiVisualChange(options: {
  prior: boolean;
  current: boolean;
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
  const table = descendants(document, 'tbl')[0];
  const properties =
    directChild(table, 'tblPr') ??
    (() => {
      const created = document.createElementNS(WORD_NAMESPACE, 'w:tblPr');
      table.insertBefore(created, table.firstChild);
      return created;
    })();
  for (const existing of Array.from(properties.children).filter(
    (child) =>
      child.localName === 'bidiVisual' || child.localName === 'tblPrChange',
  )) {
    existing.remove();
  }
  if (options.current) {
    properties.append(
      document.createElementNS(WORD_NAMESPACE, 'w:bidiVisual'),
    );
  }
  const change = document.createElementNS(WORD_NAMESPACE, 'w:tblPrChange');
  change.setAttributeNS(WORD_NAMESPACE, 'w:id', '21');
  change.setAttributeNS(WORD_NAMESPACE, 'w:author', 'Reviewer');
  change.setAttributeNS(WORD_NAMESPACE, 'w:date', '2026-09-08T00:00:00Z');
  const prior = document.createElementNS(WORD_NAMESPACE, 'w:tblPr');
  if (options.prior) {
    prior.append(document.createElementNS(WORD_NAMESPACE, 'w:bidiVisual'));
  } else {
    const off = document.createElementNS(WORD_NAMESPACE, 'w:bidiVisual');
    off.setAttributeNS(WORD_NAMESPACE, 'w:val', '0');
    prior.append(off);
  }
  change.append(prior);
  properties.append(change);
  archive.file(
    'word/document.xml',
    new XMLSerializer().serializeToString(document),
  );
  return archive.generateAsync({ type: 'arraybuffer' });
}

async function tableDocxWithAlignmentChange(options: {
  current: 'left' | 'center' | 'right';
  prior: 'left' | 'center' | 'right';
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
  const table = descendants(document, 'tbl')[0];
  const properties =
    directChild(table, 'tblPr') ??
    (() => {
      const created = document.createElementNS(WORD_NAMESPACE, 'w:tblPr');
      table.insertBefore(created, table.firstChild);
      return created;
    })();
  for (const existing of Array.from(properties.children).filter(
    (child) =>
      child.localName === 'jc' ||
      child.localName === 'tblW' ||
      child.localName === 'tblPrChange',
  )) {
    existing.remove();
  }
  const current = document.createElementNS(WORD_NAMESPACE, 'w:jc');
  current.setAttributeNS(WORD_NAMESPACE, 'w:val', options.current);
  properties.append(current);
  const change = document.createElementNS(WORD_NAMESPACE, 'w:tblPrChange');
  change.setAttributeNS(WORD_NAMESPACE, 'w:id', '11');
  change.setAttributeNS(WORD_NAMESPACE, 'w:author', 'Reviewer');
  change.setAttributeNS(WORD_NAMESPACE, 'w:date', '2026-09-08T00:00:00Z');
  const prior = document.createElementNS(WORD_NAMESPACE, 'w:tblPr');
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

async function tableDocxWithWidthChange(options: {
  currentType: string;
  currentW: string;
  priorType: string;
  priorW: string;
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
  const table = descendants(document, 'tbl')[0];
  const properties =
    directChild(table, 'tblPr') ??
    (() => {
      const created = document.createElementNS(WORD_NAMESPACE, 'w:tblPr');
      table.insertBefore(created, table.firstChild);
      return created;
    })();
  for (const existing of Array.from(properties.children).filter(
    (child) =>
      child.localName === 'jc' ||
      child.localName === 'tblW' ||
      child.localName === 'tblPrChange',
  )) {
    existing.remove();
  }
  const current = document.createElementNS(WORD_NAMESPACE, 'w:tblW');
  current.setAttributeNS(WORD_NAMESPACE, 'w:type', options.currentType);
  current.setAttributeNS(WORD_NAMESPACE, 'w:w', options.currentW);
  properties.append(current);
  const change = document.createElementNS(WORD_NAMESPACE, 'w:tblPrChange');
  change.setAttributeNS(WORD_NAMESPACE, 'w:id', '12');
  change.setAttributeNS(WORD_NAMESPACE, 'w:author', 'Reviewer');
  change.setAttributeNS(WORD_NAMESPACE, 'w:date', '2026-09-08T00:00:00Z');
  const prior = document.createElementNS(WORD_NAMESPACE, 'w:tblPr');
  const priorWidth = document.createElementNS(WORD_NAMESPACE, 'w:tblW');
  priorWidth.setAttributeNS(WORD_NAMESPACE, 'w:type', options.priorType);
  priorWidth.setAttributeNS(WORD_NAMESPACE, 'w:w', options.priorW);
  prior.append(priorWidth);
  change.append(prior);
  properties.append(change);
  archive.file(
    'word/document.xml',
    new XMLSerializer().serializeToString(document),
  );
  return archive.generateAsync({ type: 'arraybuffer' });
}

async function tableDocxWithAlignmentAndWidthChange(): Promise<ArrayBuffer> {
  const artifact = createArtifact('blank-document');
  if (artifact.content.type !== 'document') {
    throw new Error('Expected a document artifact.');
  }
  artifact.content.html =
    '<table><tbody><tr><td><p>Cell</p></td></tr></tbody></table>';
  const seed = await createArtifactBlob(artifact);
  const archive = await JSZip.loadAsync(await seed.arrayBuffer());
  const document = await xmlEntry(archive, 'word/document.xml');
  const table = descendants(document, 'tbl')[0];
  const properties =
    directChild(table, 'tblPr') ??
    (() => {
      const created = document.createElementNS(WORD_NAMESPACE, 'w:tblPr');
      table.insertBefore(created, table.firstChild);
      return created;
    })();
  for (const existing of Array.from(properties.children).filter(
    (child) =>
      child.localName === 'jc' ||
      child.localName === 'tblW' ||
      child.localName === 'tblPrChange',
  )) {
    existing.remove();
  }
  const currentJc = document.createElementNS(WORD_NAMESPACE, 'w:jc');
  currentJc.setAttributeNS(WORD_NAMESPACE, 'w:val', 'center');
  const currentWidth = document.createElementNS(WORD_NAMESPACE, 'w:tblW');
  currentWidth.setAttributeNS(WORD_NAMESPACE, 'w:type', 'pct');
  currentWidth.setAttributeNS(WORD_NAMESPACE, 'w:w', '5000');
  properties.append(currentJc, currentWidth);
  const change = document.createElementNS(WORD_NAMESPACE, 'w:tblPrChange');
  change.setAttributeNS(WORD_NAMESPACE, 'w:id', '13');
  change.setAttributeNS(WORD_NAMESPACE, 'w:author', 'Reviewer');
  change.setAttributeNS(WORD_NAMESPACE, 'w:date', '2026-09-08T00:00:00Z');
  const prior = document.createElementNS(WORD_NAMESPACE, 'w:tblPr');
  const priorJc = document.createElementNS(WORD_NAMESPACE, 'w:jc');
  priorJc.setAttributeNS(WORD_NAMESPACE, 'w:val', 'left');
  const priorWidth = document.createElementNS(WORD_NAMESPACE, 'w:tblW');
  priorWidth.setAttributeNS(WORD_NAMESPACE, 'w:type', 'auto');
  priorWidth.setAttributeNS(WORD_NAMESPACE, 'w:w', '0');
  prior.append(priorJc, priorWidth);
  change.append(prior);
  properties.append(change);
  archive.file(
    'word/document.xml',
    new XMLSerializer().serializeToString(document),
  );
  return archive.generateAsync({ type: 'arraybuffer' });
}

async function tableDocxWithIndentChange(options: {
  currentTwips: string;
  priorTwips: string;
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
  const table = descendants(document, 'tbl')[0];
  const properties =
    directChild(table, 'tblPr') ??
    (() => {
      const created = document.createElementNS(WORD_NAMESPACE, 'w:tblPr');
      table.insertBefore(created, table.firstChild);
      return created;
    })();
  for (const existing of Array.from(properties.children).filter(
    (child) =>
      child.localName === 'jc' ||
      child.localName === 'tblW' ||
      child.localName === 'tblInd' ||
      child.localName === 'tblPrChange',
  )) {
    existing.remove();
  }
  const current = document.createElementNS(WORD_NAMESPACE, 'w:tblInd');
  current.setAttributeNS(WORD_NAMESPACE, 'w:type', 'dxa');
  current.setAttributeNS(WORD_NAMESPACE, 'w:w', options.currentTwips);
  properties.append(current);
  const change = document.createElementNS(WORD_NAMESPACE, 'w:tblPrChange');
  change.setAttributeNS(WORD_NAMESPACE, 'w:id', '14');
  change.setAttributeNS(WORD_NAMESPACE, 'w:author', 'Reviewer');
  change.setAttributeNS(WORD_NAMESPACE, 'w:date', '2026-09-08T00:00:00Z');
  const prior = document.createElementNS(WORD_NAMESPACE, 'w:tblPr');
  const priorIndent = document.createElementNS(WORD_NAMESPACE, 'w:tblInd');
  priorIndent.setAttributeNS(WORD_NAMESPACE, 'w:type', 'dxa');
  priorIndent.setAttributeNS(WORD_NAMESPACE, 'w:w', options.priorTwips);
  prior.append(priorIndent);
  change.append(prior);
  properties.append(change);
  archive.file(
    'word/document.xml',
    new XMLSerializer().serializeToString(document),
  );
  return archive.generateAsync({ type: 'arraybuffer' });
}

async function tableDocxWithCellMarginChange(options: {
  priorTwips: number;
  currentTwips: number;
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
  const table = descendants(document, 'tbl')[0];
  const properties =
    directChild(table, 'tblPr') ??
    (() => {
      const created = document.createElementNS(WORD_NAMESPACE, 'w:tblPr');
      table.insertBefore(created, table.firstChild);
      return created;
    })();
  for (const existing of Array.from(properties.children).filter(
    (child) =>
      child.localName === 'tblCellMar' || child.localName === 'tblPrChange',
  )) {
    existing.remove();
  }
  properties.append(createTblCellMar(document, options.currentTwips));
  const change = document.createElementNS(WORD_NAMESPACE, 'w:tblPrChange');
  change.setAttributeNS(WORD_NAMESPACE, 'w:id', '15');
  change.setAttributeNS(WORD_NAMESPACE, 'w:author', 'Reviewer');
  change.setAttributeNS(WORD_NAMESPACE, 'w:date', '2026-09-08T00:00:00Z');
  const prior = document.createElementNS(WORD_NAMESPACE, 'w:tblPr');
  prior.append(createTblCellMar(document, options.priorTwips));
  change.append(prior);
  properties.append(change);
  archive.file(
    'word/document.xml',
    new XMLSerializer().serializeToString(document),
  );
  return archive.generateAsync({ type: 'arraybuffer' });
}

function createTblCellMar(document: Document, twips: number): Element {
  const margins = document.createElementNS(WORD_NAMESPACE, 'w:tblCellMar');
  for (const side of ['top', 'left', 'bottom', 'right'] as const) {
    const edge = document.createElementNS(WORD_NAMESPACE, `w:${side}`);
    edge.setAttributeNS(WORD_NAMESPACE, 'w:type', 'dxa');
    edge.setAttributeNS(WORD_NAMESPACE, 'w:w', String(twips));
    margins.append(edge);
  }
  return margins;
}

async function tableDocxWithLayoutChange(options: {
  prior: 'autofit' | 'fixed';
  current: 'autofit' | 'fixed';
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
  const table = descendants(document, 'tbl')[0];
  const properties =
    directChild(table, 'tblPr') ??
    (() => {
      const created = document.createElementNS(WORD_NAMESPACE, 'w:tblPr');
      table.insertBefore(created, table.firstChild);
      return created;
    })();
  for (const existing of Array.from(properties.children).filter(
    (child) =>
      child.localName === 'tblLayout' || child.localName === 'tblPrChange',
  )) {
    existing.remove();
  }
  const current = document.createElementNS(WORD_NAMESPACE, 'w:tblLayout');
  current.setAttributeNS(WORD_NAMESPACE, 'w:type', options.current);
  properties.append(current);
  const change = document.createElementNS(WORD_NAMESPACE, 'w:tblPrChange');
  change.setAttributeNS(WORD_NAMESPACE, 'w:id', '16');
  change.setAttributeNS(WORD_NAMESPACE, 'w:author', 'Reviewer');
  change.setAttributeNS(WORD_NAMESPACE, 'w:date', '2026-09-08T00:00:00Z');
  const prior = document.createElementNS(WORD_NAMESPACE, 'w:tblPr');
  const priorLayout = document.createElementNS(WORD_NAMESPACE, 'w:tblLayout');
  priorLayout.setAttributeNS(WORD_NAMESPACE, 'w:type', options.prior);
  prior.append(priorLayout);
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
