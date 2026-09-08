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
