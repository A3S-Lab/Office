import { Editor } from '@tiptap/core';
import { describe, expect, test } from '@rstest/core';
import JSZip from 'jszip';
import {
  createArtifact,
  createArtifactBlob,
  importOfficeFile,
} from '../src/core';
import { collectDocumentChanges } from '../src/internal/features/work/work-document-changes';
import { parseDocumentCellFormatting } from '../src/internal/features/work/work-document-cell-format-changes';
import { createWorkDocumentExtensions } from '../src/internal/features/work/work-document-extensions';
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

describe('DOCX cell-formatting revisions', () => {
  test('imports vAlign-only w:tcPrChange as a reviewable cell-formatting change', async () => {
    const source = await cellDocxWithVAlignChange({ prior: 'center' });
    const imported = await importOfficeFile(
      new File([source], 'cell-formatting.docx'),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const html = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    const cell = html.body.querySelector('td');
    expect(cell?.dataset.changeKind).toBe('cell-formatting');
    expect(cell?.dataset.officeCellPropertyRevisionOmml).toBeUndefined();
    expect(parseDocumentCellFormatting(cell?.dataset.changeBefore)).toEqual({
      verticalAlign: 'middle',
    });

    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: imported.content.html,
    });
    try {
      const changes = collectDocumentChanges(editor.state.doc);
      expect(changes).toHaveLength(1);
      expect(changes[0]?.kind).toBe('cell-formatting');
      expect(editor.commands.acceptDocumentChange(changes[0]?.id ?? '')).toBe(
        true,
      );
      expect(collectDocumentChanges(editor.state.doc)).toHaveLength(0);
    } finally {
      editor.destroy();
    }
  });

  test('reject restores prior verticalAlign and drops the pending change', async () => {
    const source = await cellDocxWithVAlignChange({
      prior: 'top',
      current: 'bottom',
    });
    const imported = await importOfficeFile(
      new File([source], 'cell-formatting-reject.docx'),
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
      const cell = html.body.querySelector('td');
      expect(cell?.dataset.changeKind).toBeUndefined();
      expect(cell?.dataset.officeCellVerticalAlign).toBe('top');
    } finally {
      editor.destroy();
    }
  });

  test('pending cell-formatting change round-trips as native w:tcPrChange', async () => {
    const source = await cellDocxWithVAlignChange({ prior: 'center' });
    const imported = await importOfficeFile(
      new File([source], 'cell-formatting-roundtrip.docx'),
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
      directChild(descendants(document, 'tc')[0], 'tcPr'),
      'tcPrChange',
    );
    expect(change).toBeTruthy();
    expect(wordAttribute(change!, 'author')).toBe('Reviewer');
    const priorAlign = directChild(directChild(change, 'tcPr'), 'vAlign');
    expect(
      priorAlign?.getAttributeNS(WORD_NAMESPACE, 'val') ??
        priorAlign?.getAttribute('w:val') ??
        priorAlign?.getAttribute('val'),
    ).toBe('center');
  });

  test('reports vAlign-only cell-property revisions as reviewable diagnostics', async () => {
    const source = await cellDocxWithVAlignChange({ prior: 'center' });
    const summary = await analyzeDocxCompatibility(
      new File([source], 'cell-formatting-diag.docx'),
      [],
    );
    expect(
      summary.issues.some(
        (issue) => issue.code === 'docx.revisions.cell-formatting',
      ),
    ).toBe(true);
  });

  test('imports solid-shd-only w:tcPrChange as a reviewable cell-formatting change', async () => {
    const source = await cellDocxWithFillChange({ prior: 'FFAA00' });
    const imported = await importOfficeFile(
      new File([source], 'cell-formatting-fill.docx'),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const html = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    const cell = html.body.querySelector('td');
    expect(cell?.dataset.changeKind).toBe('cell-formatting');
    expect(parseDocumentCellFormatting(cell?.dataset.changeBefore)).toEqual({
      fill: '#ffaa00',
    });
  });

  test('reject restores prior solid fill and drops the pending change', async () => {
    const source = await cellDocxWithFillChange({
      prior: '112233',
      current: 'AABBCC',
    });
    const imported = await importOfficeFile(
      new File([source], 'cell-formatting-fill-reject.docx'),
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
      const cell = html.body.querySelector('td');
      expect(cell?.dataset.changeKind).toBeUndefined();
      expect(cell?.dataset.officeCellFill).toBe('#112233');
    } finally {
      editor.destroy();
    }
  });

  test('live cell verticalAlign edits become reviewable when track changes is on', () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions({
        isTracking: () => true,
      }),
      content:
        '<table><tbody><tr><td data-office-cell-vertical-align="top"><p>Cell</p></td></tr></tbody></table>',
    });
    try {
      expect(
        editor.commands.setDocumentTableCellFormat({
          verticalAlign: 'bottom',
        }),
      ).toBe(true);
      const changes = collectDocumentChanges(editor.state.doc);
      expect(changes).toHaveLength(1);
      expect(changes[0]?.kind).toBe('cell-formatting');
      const html = new DOMParser().parseFromString(
        editor.getHTML(),
        'text/html',
      );
      const cell = html.body.querySelector('td');
      expect(cell?.dataset.changeKind).toBe('cell-formatting');
      const before = parseDocumentCellFormatting(cell?.dataset.changeBefore);
      expect(before?.verticalAlign).toBe('top');
      expect(cell?.dataset.officeCellVerticalAlign).toBe('bottom');
    } finally {
      editor.destroy();
    }
  });

  test('imports tcMar-only w:tcPrChange as a reviewable cell-formatting change', async () => {
    const source = await cellDocxWithMarginChange({
      priorTwips: 720,
      currentTwips: 1440,
    });
    const imported = await importOfficeFile(
      new File([source], 'cell-formatting-margins.docx'),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const html = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    const cell = html.body.querySelector('td');
    expect(cell?.dataset.changeKind).toBe('cell-formatting');
    expect(cell?.dataset.officeCellPropertyRevisionOmml).toBeUndefined();
    expect(parseDocumentCellFormatting(cell?.dataset.changeBefore)).toEqual({
      margins: { top: 48, right: 48, bottom: 48, left: 48 },
    });

    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: imported.content.html,
    });
    try {
      const change = collectDocumentChanges(editor.state.doc)[0];
      expect(change?.kind).toBe('cell-formatting');
      expect(editor.commands.rejectDocumentChange(change?.id ?? '')).toBe(true);
      const rejected = new DOMParser().parseFromString(
        editor.getHTML(),
        'text/html',
      );
      const rejectedCell = rejected.body.querySelector('td');
      expect(rejectedCell?.dataset.changeKind).toBeUndefined();
      expect(rejectedCell?.dataset.officeCellMarginTop).toBe('48');
      expect(rejectedCell?.dataset.officeCellMarginLeft).toBe('48');
    } finally {
      editor.destroy();
    }
  });

  test('live cell margin edits become reviewable when track changes is on', () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions({
        isTracking: () => true,
      }),
      content: [
        '<table><tbody><tr>',
        '<td data-office-cell-margin-top="12" data-office-cell-margin-right="12"',
        ' data-office-cell-margin-bottom="12" data-office-cell-margin-left="12">',
        '<p>Cell</p></td>',
        '</tr></tbody></table>',
      ].join(''),
    });
    try {
      expect(
        editor.commands.setDocumentTableCellFormat({
          margins: { top: 48, right: 48, bottom: 48, left: 48 },
        }),
      ).toBe(true);
      const changes = collectDocumentChanges(editor.state.doc);
      expect(changes).toHaveLength(1);
      expect(changes[0]?.kind).toBe('cell-formatting');
      const html = new DOMParser().parseFromString(
        editor.getHTML(),
        'text/html',
      );
      const cell = html.body.querySelector('td');
      expect(cell?.dataset.changeKind).toBe('cell-formatting');
      const before = parseDocumentCellFormatting(cell?.dataset.changeBefore);
      expect(before?.margins).toEqual({
        top: 12,
        right: 12,
        bottom: 12,
        left: 12,
      });
      expect(cell?.dataset.officeCellMarginTop).toBe('48');
    } finally {
      editor.destroy();
    }
  });

  test('imports tcW-only w:tcPrChange as a reviewable cell-formatting change', async () => {
    const source = await cellDocxWithWidthChange({
      priorTwips: 1440,
      currentTwips: 2880,
    });
    const imported = await importOfficeFile(
      new File([source], 'cell-formatting-width.docx'),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const html = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    const cell = html.body.querySelector('td');
    expect(cell?.dataset.changeKind).toBe('cell-formatting');
    expect(cell?.dataset.officeCellPropertyRevisionOmml).toBeUndefined();
    expect(parseDocumentCellFormatting(cell?.dataset.changeBefore)).toEqual({
      width: { type: 'pixels', value: 96 },
    });

    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: imported.content.html,
    });
    try {
      const change = collectDocumentChanges(editor.state.doc)[0];
      expect(change?.kind).toBe('cell-formatting');
      expect(editor.commands.rejectDocumentChange(change?.id ?? '')).toBe(true);
      const rejected = new DOMParser().parseFromString(
        editor.getHTML(),
        'text/html',
      );
      const rejectedCell = rejected.body.querySelector('td');
      expect(rejectedCell?.dataset.changeKind).toBeUndefined();
      expect(rejectedCell?.getAttribute('colwidth')).toBe('96');
    } finally {
      editor.destroy();
    }
  });

  test('pending tcW cell-formatting change round-trips as native w:tcPrChange', async () => {
    const source = await cellDocxWithWidthChange({ priorTwips: 1440 });
    const imported = await importOfficeFile(
      new File([source], 'cell-formatting-width-roundtrip.docx'),
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
      directChild(descendants(exported, 'tc')[0], 'tcPr'),
      'tcPrChange',
    );
    expect(change).toBeTruthy();
    const priorWidth = directChild(directChild(change!, 'tcPr'), 'tcW');
    expect(
      priorWidth?.getAttributeNS(WORD_NAMESPACE, 'type') ??
        priorWidth?.getAttribute('w:type') ??
        priorWidth?.getAttribute('type'),
    ).toBe('dxa');
    expect(
      priorWidth?.getAttributeNS(WORD_NAMESPACE, 'w') ??
        priorWidth?.getAttribute('w:w') ??
        priorWidth?.getAttribute('w'),
    ).toBe('1440');
  });

  test('live cell preferred-width edits become reviewable when track changes is on', () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions({
        isTracking: () => true,
      }),
      content:
        '<table><tbody><tr><td colwidth="96"><p>Cell</p></td></tr></tbody></table>',
    });
    try {
      let cellPos: number | null = null;
      editor.state.doc.descendants((node, position) => {
        if (node.type.name === 'tableCell' && cellPos === null) {
          cellPos = position;
        }
      });
      expect(cellPos).not.toBeNull();
      editor.view.dispatch(
        editor.state.tr.setNodeMarkup(cellPos!, undefined, {
          ...editor.state.doc.nodeAt(cellPos!)!.attrs,
          colwidth: [192],
        }),
      );
      const changes = collectDocumentChanges(editor.state.doc).filter(
        (change) => change.kind === 'cell-formatting',
      );
      expect(changes).toHaveLength(1);
      expect(changes[0]?.kind).toBe('cell-formatting');
      const html = new DOMParser().parseFromString(
        editor.getHTML(),
        'text/html',
      );
      const cell = html.body.querySelector('td');
      expect(cell?.dataset.changeKind).toBe('cell-formatting');
      expect(parseDocumentCellFormatting(cell?.dataset.changeBefore)).toEqual({
        verticalAlign: 'top',
        fill: '#ffffff',
        width: { type: 'pixels', value: 96 },
        noWrap: false,
        textDirection: 'lrTb',
        fitText: false,
        hideMark: false,
      });
      expect(cell?.getAttribute('colwidth')).toBe('192');
    } finally {
      editor.destroy();
    }
  });

  test('imports noWrap-only w:tcPrChange as a reviewable cell-formatting change', async () => {
    const source = await cellDocxWithNoWrapChange({
      prior: true,
      current: false,
    });
    const imported = await importOfficeFile(
      new File([source], 'cell-formatting-nowrap.docx'),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const html = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    const cell = html.body.querySelector('td');
    expect(cell?.dataset.changeKind).toBe('cell-formatting');
    expect(cell?.dataset.officeCellPropertyRevisionOmml).toBeUndefined();
    expect(parseDocumentCellFormatting(cell?.dataset.changeBefore)).toEqual({
      noWrap: true,
    });

    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: imported.content.html,
    });
    try {
      const change = collectDocumentChanges(editor.state.doc)[0];
      expect(change?.kind).toBe('cell-formatting');
      expect(editor.commands.rejectDocumentChange(change?.id ?? '')).toBe(true);
      const rejected = new DOMParser().parseFromString(
        editor.getHTML(),
        'text/html',
      );
      const rejectedCell = rejected.body.querySelector('td');
      expect(rejectedCell?.dataset.changeKind).toBeUndefined();
      expect(rejectedCell?.dataset.officeCellNoWrap).toBe('true');
    } finally {
      editor.destroy();
    }
  });

  test('pending noWrap cell-formatting change round-trips as native w:tcPrChange', async () => {
    const source = await cellDocxWithNoWrapChange({ prior: true });
    const imported = await importOfficeFile(
      new File([source], 'cell-formatting-nowrap-roundtrip.docx'),
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
      directChild(descendants(exported, 'tc')[0], 'tcPr'),
      'tcPrChange',
    );
    expect(change).toBeTruthy();
    expect(directChild(directChild(change!, 'tcPr'), 'noWrap')).toBeTruthy();
  });

  test('live cell noWrap edits become reviewable when track changes is on', () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions({
        isTracking: () => true,
      }),
      content:
        '<table><tbody><tr><td data-office-cell-no-wrap="false"><p>Cell</p></td></tr></tbody></table>',
    });
    try {
      let cellPos: number | null = null;
      editor.state.doc.descendants((node, position) => {
        if (node.type.name === 'tableCell' && cellPos === null) {
          cellPos = position;
        }
      });
      expect(cellPos).not.toBeNull();
      editor.view.dispatch(
        editor.state.tr.setNodeMarkup(cellPos!, undefined, {
          ...editor.state.doc.nodeAt(cellPos!)!.attrs,
          noWrap: true,
        }),
      );
      const changes = collectDocumentChanges(editor.state.doc).filter(
        (change) => change.kind === 'cell-formatting',
      );
      expect(changes).toHaveLength(1);
      const html = new DOMParser().parseFromString(
        editor.getHTML(),
        'text/html',
      );
      const cell = html.body.querySelector('td');
      expect(cell?.dataset.changeKind).toBe('cell-formatting');
      expect(
        parseDocumentCellFormatting(cell?.dataset.changeBefore),
      ).toMatchObject({
        noWrap: false,
      });
      expect(cell?.dataset.officeCellNoWrap).toBe('true');
    } finally {
      editor.destroy();
    }
  });

  test('imports textDirection-only w:tcPrChange as a reviewable cell-formatting change', async () => {
    const source = await cellDocxWithTextDirectionChange({
      prior: 'btLr',
      current: 'tbRl',
    });
    const imported = await importOfficeFile(
      new File([source], 'cell-formatting-text-direction.docx'),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const html = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    const cell = html.body.querySelector('td');
    expect(cell?.dataset.changeKind).toBe('cell-formatting');
    expect(cell?.dataset.officeCellPropertyRevisionOmml).toBeUndefined();
    expect(parseDocumentCellFormatting(cell?.dataset.changeBefore)).toEqual({
      textDirection: 'btLr',
    });
    expect(cell?.dataset.officeCellTextDirection).toBe('tbRl');

    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: imported.content.html,
    });
    try {
      const change = collectDocumentChanges(editor.state.doc)[0];
      expect(change?.kind).toBe('cell-formatting');
      expect(editor.commands.rejectDocumentChange(change?.id ?? '')).toBe(true);
      const rejected = new DOMParser().parseFromString(
        editor.getHTML(),
        'text/html',
      );
      const rejectedCell = rejected.body.querySelector('td');
      expect(rejectedCell?.dataset.changeKind).toBeUndefined();
      expect(rejectedCell?.dataset.officeCellTextDirection).toBe('btLr');
    } finally {
      editor.destroy();
    }
  });

  test('pending textDirection cell-formatting change round-trips as native w:tcPrChange', async () => {
    const source = await cellDocxWithTextDirectionChange({ prior: 'tbRl' });
    const imported = await importOfficeFile(
      new File([source], 'cell-formatting-text-direction-roundtrip.docx'),
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
      directChild(descendants(exported, 'tc')[0], 'tcPr'),
      'tcPrChange',
    );
    expect(change).toBeTruthy();
    const priorDirection = directChild(
      directChild(change!, 'tcPr'),
      'textDirection',
    );
    expect(
      priorDirection?.getAttributeNS(WORD_NAMESPACE, 'val') ??
        priorDirection?.getAttribute('w:val') ??
        priorDirection?.getAttribute('val'),
    ).toBe('tbRl');
  });

  test('live cell textDirection edits become reviewable when track changes is on', () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions({
        isTracking: () => true,
      }),
      content:
        '<table><tbody><tr><td data-office-cell-text-direction="lrTb"><p>Cell</p></td></tr></tbody></table>',
    });
    try {
      let cellPos: number | null = null;
      editor.state.doc.descendants((node, position) => {
        if (node.type.name === 'tableCell' && cellPos === null) {
          cellPos = position;
        }
      });
      expect(cellPos).not.toBeNull();
      editor.view.dispatch(
        editor.state.tr.setNodeMarkup(cellPos!, undefined, {
          ...editor.state.doc.nodeAt(cellPos!)!.attrs,
          textDirection: 'tbRl',
        }),
      );
      const changes = collectDocumentChanges(editor.state.doc).filter(
        (change) => change.kind === 'cell-formatting',
      );
      expect(changes).toHaveLength(1);
      const html = new DOMParser().parseFromString(
        editor.getHTML(),
        'text/html',
      );
      const cell = html.body.querySelector('td');
      expect(cell?.dataset.changeKind).toBe('cell-formatting');
      expect(
        parseDocumentCellFormatting(cell?.dataset.changeBefore),
      ).toMatchObject({
        textDirection: 'lrTb',
      });
      expect(cell?.dataset.officeCellTextDirection).toBe('tbRl');
    } finally {
      editor.destroy();
    }
  });

  test('imports tcFitText-only w:tcPrChange as a reviewable cell-formatting change', async () => {
    const source = await cellDocxWithFitTextChange({
      prior: true,
      current: false,
    });
    const imported = await importOfficeFile(
      new File([source], 'cell-formatting-fit-text.docx'),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const html = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    const cell = html.body.querySelector('td');
    expect(cell?.dataset.changeKind).toBe('cell-formatting');
    expect(cell?.dataset.officeCellPropertyRevisionOmml).toBeUndefined();
    expect(parseDocumentCellFormatting(cell?.dataset.changeBefore)).toEqual({
      fitText: true,
    });

    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: imported.content.html,
    });
    try {
      const change = collectDocumentChanges(editor.state.doc)[0];
      expect(change?.kind).toBe('cell-formatting');
      expect(editor.commands.rejectDocumentChange(change?.id ?? '')).toBe(true);
      const rejected = new DOMParser().parseFromString(
        editor.getHTML(),
        'text/html',
      );
      const rejectedCell = rejected.body.querySelector('td');
      expect(rejectedCell?.dataset.changeKind).toBeUndefined();
      expect(rejectedCell?.dataset.officeCellFitText).toBe('true');
    } finally {
      editor.destroy();
    }
  });

  test('pending tcFitText cell-formatting change round-trips as native w:tcPrChange', async () => {
    const source = await cellDocxWithFitTextChange({ prior: true });
    const imported = await importOfficeFile(
      new File([source], 'cell-formatting-fit-text-roundtrip.docx'),
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
      directChild(descendants(exported, 'tc')[0], 'tcPr'),
      'tcPrChange',
    );
    expect(change).toBeTruthy();
    expect(directChild(directChild(change!, 'tcPr'), 'tcFitText')).toBeTruthy();
  });

  test('live cell fitText edits become reviewable when track changes is on', () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions({
        isTracking: () => true,
      }),
      content:
        '<table><tbody><tr><td data-office-cell-fit-text="false"><p>Cell</p></td></tr></tbody></table>',
    });
    try {
      let cellPos: number | null = null;
      editor.state.doc.descendants((node, position) => {
        if (node.type.name === 'tableCell' && cellPos === null) {
          cellPos = position;
        }
      });
      expect(cellPos).not.toBeNull();
      editor.view.dispatch(
        editor.state.tr.setNodeMarkup(cellPos!, undefined, {
          ...editor.state.doc.nodeAt(cellPos!)!.attrs,
          fitText: true,
        }),
      );
      const changes = collectDocumentChanges(editor.state.doc).filter(
        (change) => change.kind === 'cell-formatting',
      );
      expect(changes).toHaveLength(1);
      const html = new DOMParser().parseFromString(
        editor.getHTML(),
        'text/html',
      );
      const cell = html.body.querySelector('td');
      expect(cell?.dataset.changeKind).toBe('cell-formatting');
      expect(
        parseDocumentCellFormatting(cell?.dataset.changeBefore),
      ).toMatchObject({
        fitText: false,
      });
      expect(cell?.dataset.officeCellFitText).toBe('true');
    } finally {
      editor.destroy();
    }
  });

  test('imports hideMark-only w:tcPrChange as a reviewable cell-formatting change', async () => {
    const source = await cellDocxWithHideMarkChange({
      prior: true,
      current: false,
    });
    const imported = await importOfficeFile(
      new File([source], 'cell-formatting-hide-mark.docx'),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const html = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    const cell = html.body.querySelector('td');
    expect(cell?.dataset.changeKind).toBe('cell-formatting');
    expect(cell?.dataset.officeCellPropertyRevisionOmml).toBeUndefined();
    expect(parseDocumentCellFormatting(cell?.dataset.changeBefore)).toEqual({
      hideMark: true,
    });

    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: imported.content.html,
    });
    try {
      const change = collectDocumentChanges(editor.state.doc)[0];
      expect(change?.kind).toBe('cell-formatting');
      expect(editor.commands.rejectDocumentChange(change?.id ?? '')).toBe(true);
      const rejected = new DOMParser().parseFromString(
        editor.getHTML(),
        'text/html',
      );
      const rejectedCell = rejected.body.querySelector('td');
      expect(rejectedCell?.dataset.changeKind).toBeUndefined();
      expect(rejectedCell?.dataset.officeCellHideMark).toBe('true');
    } finally {
      editor.destroy();
    }
  });

  test('pending hideMark cell-formatting change round-trips as native w:tcPrChange', async () => {
    const source = await cellDocxWithHideMarkChange({ prior: true });
    const imported = await importOfficeFile(
      new File([source], 'cell-formatting-hide-mark-roundtrip.docx'),
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
      directChild(descendants(exported, 'tc')[0], 'tcPr'),
      'tcPrChange',
    );
    expect(change).toBeTruthy();
    expect(directChild(directChild(change!, 'tcPr'), 'hideMark')).toBeTruthy();
  });

  test('live cell hideMark edits become reviewable when track changes is on', () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions({
        isTracking: () => true,
      }),
      content:
        '<table><tbody><tr><td data-office-cell-hide-mark="false"><p>Cell</p></td></tr></tbody></table>',
    });
    try {
      let cellPos: number | null = null;
      editor.state.doc.descendants((node, position) => {
        if (node.type.name === 'tableCell' && cellPos === null) {
          cellPos = position;
        }
      });
      expect(cellPos).not.toBeNull();
      editor.view.dispatch(
        editor.state.tr.setNodeMarkup(cellPos!, undefined, {
          ...editor.state.doc.nodeAt(cellPos!)!.attrs,
          hideMark: true,
        }),
      );
      const changes = collectDocumentChanges(editor.state.doc).filter(
        (change) => change.kind === 'cell-formatting',
      );
      expect(changes).toHaveLength(1);
      const html = new DOMParser().parseFromString(
        editor.getHTML(),
        'text/html',
      );
      const cell = html.body.querySelector('td');
      expect(cell?.dataset.changeKind).toBe('cell-formatting');
      expect(
        parseDocumentCellFormatting(cell?.dataset.changeBefore),
      ).toMatchObject({
        hideMark: false,
      });
      expect(cell?.dataset.officeCellHideMark).toBe('true');
    } finally {
      editor.destroy();
    }
  });

  test('imports cnfStyle-only w:tcPrChange as a reviewable cell-formatting change', async () => {
    const source = await cellDocxWithCnfStyleChange({
      prior: '100000000000',
      current: '010000000000',
    });
    const imported = await importOfficeFile(
      new File([source], 'cell-formatting-cnf-style.docx'),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const html = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    const cell = html.body.querySelector('td');
    expect(cell?.dataset.changeKind).toBe('cell-formatting');
    expect(cell?.dataset.officeCellPropertyRevisionOmml).toBeUndefined();
    expect(parseDocumentCellFormatting(cell?.dataset.changeBefore)).toEqual({
      cnfStyle: '100000000000',
    });
    expect(cell?.dataset.officeCellCnfStyle).toBe('010000000000');

    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: imported.content.html,
    });
    try {
      const change = collectDocumentChanges(editor.state.doc)[0];
      expect(change?.kind).toBe('cell-formatting');
      expect(editor.commands.rejectDocumentChange(change?.id ?? '')).toBe(true);
      const rejected = new DOMParser().parseFromString(
        editor.getHTML(),
        'text/html',
      );
      const rejectedCell = rejected.body.querySelector('td');
      expect(rejectedCell?.dataset.changeKind).toBeUndefined();
      expect(rejectedCell?.dataset.officeCellCnfStyle).toBe('100000000000');
    } finally {
      editor.destroy();
    }
  });

  test('pending cnfStyle cell-formatting change round-trips as native w:tcPrChange', async () => {
    const source = await cellDocxWithCnfStyleChange({
      prior: '101000000100',
    });
    const imported = await importOfficeFile(
      new File([source], 'cell-formatting-cnf-style-roundtrip.docx'),
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
      directChild(descendants(exported, 'tc')[0], 'tcPr'),
      'tcPrChange',
    );
    expect(change).toBeTruthy();
    const priorCnfStyle = directChild(directChild(change!, 'tcPr'), 'cnfStyle');
    expect(
      priorCnfStyle?.getAttributeNS(WORD_NAMESPACE, 'val') ??
        priorCnfStyle?.getAttribute('w:val') ??
        priorCnfStyle?.getAttribute('val'),
    ).toBe('101000000100');
  });

  test('malformed cnfStyle w:tcPrChange stays on the opaque cell metadata path', async () => {
    const source = await cellDocxWithCnfStyleChange({
      prior: '100000000000',
      malformedPrior: true,
    });
    const imported = await importOfficeFile(
      new File([source], 'cell-formatting-cnf-style-opaque.docx'),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const html = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    const cell = html.body.querySelector('td');
    expect(cell?.dataset.changeKind).toBeUndefined();
    expect(cell?.dataset.officeCellPropertyRevisionOmml).toBeTruthy();
  });

  test('imports tcBorders-only w:tcPrChange as a reviewable cell-formatting change', async () => {
    const source = await cellDocxWithBordersChange({
      prior: {
        top: { val: 'single', sz: '12', color: 'FF0000' },
      },
      current: {
        top: { val: 'double', sz: '24', color: '0000FF' },
      },
    });
    const imported = await importOfficeFile(
      new File([source], 'cell-borders-formatting.docx'),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const html = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    const cell = html.body.querySelector('td');
    expect(cell?.dataset.changeKind).toBe('cell-formatting');
    expect(cell?.dataset.officeCellPropertyRevisionOmml).toBeUndefined();
    expect(parseDocumentCellFormatting(cell?.dataset.changeBefore)).toEqual({
      borders: {
        top: { color: '#ff0000', style: 'solid', width: 2 },
      },
    });
    expect(cell?.dataset.officeCellBorderTopColor?.toLowerCase()).toBe(
      '#0000ff',
    );
  });

  test('pending tcBorders cell-formatting change round-trips as native w:tcPrChange', async () => {
    const source = await cellDocxWithBordersChange({
      prior: {
        top: { val: 'single', sz: '12', color: 'FF0000' },
      },
      current: {
        top: { val: 'double', sz: '24', color: '0000FF' },
      },
    });
    const imported = await importOfficeFile(
      new File([source], 'cell-borders-roundtrip.docx'),
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
    const properties = directChild(descendants(exported, 'tc')[0], 'tcPr');
    const change = directChild(properties, 'tcPrChange');
    expect(change).toBeTruthy();
    expect(
      directChild(directChild(change!, 'tcPr'), 'tcBorders'),
    ).toBeTruthy();
    expect(directChild(properties, 'tcBorders')).toBeTruthy();
  });

  test('theme-bound tcBorders w:tcPrChange stays on the opaque cell metadata path', async () => {
    const source = await cellDocxWithBordersChange({
      prior: {
        top: { val: 'single', sz: '12', color: 'FF0000', themeColor: 'accent1' },
      },
      current: {
        top: { val: 'double', sz: '24', color: '0000FF' },
      },
    });
    const imported = await importOfficeFile(
      new File([source], 'cell-borders-opaque.docx'),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const html = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    const cell = html.body.querySelector('td');
    expect(cell?.dataset.changeKind).toBeUndefined();
    expect(cell?.dataset.officeCellPropertyRevisionOmml).toBeTruthy();
  });

  test('live cell border edits become reviewable when track changes is on', () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions({
        isTracking: () => true,
      }),
      content: [
        '<table><tbody><tr><td',
        ' data-office-cell-border-top-color="#ff0000"',
        ' data-office-cell-border-top-style="solid"',
        ' data-office-cell-border-top-width="1"',
        ' data-office-cell-border-right-color="#ff0000"',
        ' data-office-cell-border-right-style="solid"',
        ' data-office-cell-border-right-width="1"',
        ' data-office-cell-border-bottom-color="#ff0000"',
        ' data-office-cell-border-bottom-style="solid"',
        ' data-office-cell-border-bottom-width="1"',
        ' data-office-cell-border-left-color="#ff0000"',
        ' data-office-cell-border-left-style="solid"',
        ' data-office-cell-border-left-width="1"',
        ' style="border-top: 1px solid #ff0000; border-right: 1px solid #ff0000; border-bottom: 1px solid #ff0000; border-left: 1px solid #ff0000">',
        '<p>Cell</p></td></tr></tbody></table>',
      ].join(''),
    });
    try {
      let cellPos: number | null = null;
      editor.state.doc.descendants((node, position) => {
        if (node.type.name === 'tableCell' && cellPos === null) {
          cellPos = position;
        }
      });
      expect(cellPos).not.toBeNull();
      const cell = editor.state.doc.nodeAt(cellPos!)!;
      editor.view.dispatch(
        editor.state.tr.setNodeMarkup(cellPos!, undefined, {
          ...cell.attrs,
          borders: {
            top: { color: '#0000ff', style: 'double', width: 2 },
            right: cell.attrs.borders.right,
            bottom: cell.attrs.borders.bottom,
            left: cell.attrs.borders.left,
          },
          borderColor: '#0000ff',
          borderStyle: 'double',
          borderWidth: 2,
        }),
      );
      const changes = collectDocumentChanges(editor.state.doc).filter(
        (change) => change.kind === 'cell-formatting',
      );
      expect(changes).toHaveLength(1);
      const html = new DOMParser().parseFromString(
        editor.getHTML(),
        'text/html',
      );
      const rendered = html.body.querySelector('td');
      expect(rendered?.dataset.changeKind).toBe('cell-formatting');
      expect(
        parseDocumentCellFormatting(rendered?.dataset.changeBefore),
      ).toMatchObject({
        borders: {
          top: { color: '#ff0000', style: 'solid', width: 1 },
          right: { color: '#ff0000', style: 'solid', width: 1 },
          bottom: { color: '#ff0000', style: 'solid', width: 1 },
          left: { color: '#ff0000', style: 'solid', width: 1 },
        },
      });
      expect(rendered?.dataset.officeCellBorderTopColor?.toLowerCase()).toBe(
        '#0000ff',
      );
    } finally {
      editor.destroy();
    }
  });

  test('live cell cnfStyle edits become reviewable when track changes is on', () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions({
        isTracking: () => true,
      }),
      content:
        '<table><tbody><tr><td data-office-cell-cnf-style="000000000000"><p>Cell</p></td></tr></tbody></table>',
    });
    try {
      let cellPos: number | null = null;
      editor.state.doc.descendants((node, position) => {
        if (node.type.name === 'tableCell' && cellPos === null) {
          cellPos = position;
        }
      });
      expect(cellPos).not.toBeNull();
      editor.view.dispatch(
        editor.state.tr.setNodeMarkup(cellPos!, undefined, {
          ...editor.state.doc.nodeAt(cellPos!)!.attrs,
          cnfStyle: '100000000000',
        }),
      );
      const changes = collectDocumentChanges(editor.state.doc).filter(
        (change) => change.kind === 'cell-formatting',
      );
      expect(changes).toHaveLength(1);
      const html = new DOMParser().parseFromString(
        editor.getHTML(),
        'text/html',
      );
      const cell = html.body.querySelector('td');
      expect(cell?.dataset.changeKind).toBe('cell-formatting');
      expect(
        parseDocumentCellFormatting(cell?.dataset.changeBefore),
      ).toMatchObject({
        cnfStyle: '000000000000',
      });
      expect(cell?.dataset.officeCellCnfStyle).toBe('100000000000');
    } finally {
      editor.destroy();
    }
  });

  test('imports hMerge-only w:tcPrChange as a reviewable cell-formatting change', async () => {
    const source = await cellDocxWithHMergeChange({
      prior: 'restart',
      current: 'continue',
    });
    const imported = await importOfficeFile(
      new File([source], 'cell-formatting-h-merge.docx'),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const html = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    const cell = html.body.querySelector('td');
    expect(cell?.dataset.changeKind).toBe('cell-formatting');
    expect(cell?.dataset.officeCellPropertyRevisionOmml).toBeUndefined();
    expect(parseDocumentCellFormatting(cell?.dataset.changeBefore)).toEqual({
      hMerge: 'restart',
    });
    expect(cell?.dataset.officeCellHMerge).toBe('continue');

    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: imported.content.html,
    });
    try {
      const change = collectDocumentChanges(editor.state.doc)[0];
      expect(change?.kind).toBe('cell-formatting');
      expect(editor.commands.rejectDocumentChange(change?.id ?? '')).toBe(true);
      const rejected = new DOMParser().parseFromString(
        editor.getHTML(),
        'text/html',
      );
      const rejectedCell = rejected.body.querySelector('td');
      expect(rejectedCell?.dataset.changeKind).toBeUndefined();
      expect(rejectedCell?.dataset.officeCellHMerge).toBe('restart');
    } finally {
      editor.destroy();
    }
  });

  test('pending hMerge cell-formatting change round-trips as native w:tcPrChange', async () => {
    const source = await cellDocxWithHMergeChange({
      prior: 'continue',
      omitPriorVal: true,
    });
    const imported = await importOfficeFile(
      new File([source], 'cell-formatting-h-merge-roundtrip.docx'),
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
      directChild(descendants(exported, 'tc')[0], 'tcPr'),
      'tcPrChange',
    );
    expect(change).toBeTruthy();
    const priorHMerge = directChild(directChild(change!, 'tcPr'), 'hMerge');
    expect(
      priorHMerge?.getAttributeNS(WORD_NAMESPACE, 'val') ??
        priorHMerge?.getAttribute('w:val') ??
        priorHMerge?.getAttribute('val'),
    ).toBe('continue');
  });

  test('malformed hMerge w:tcPrChange stays on the opaque cell metadata path', async () => {
    const source = await cellDocxWithHMergeChange({
      prior: 'restart',
      malformedPrior: true,
    });
    const imported = await importOfficeFile(
      new File([source], 'cell-formatting-h-merge-opaque.docx'),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const html = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    const cell = html.body.querySelector('td');
    expect(cell?.dataset.changeKind).toBeUndefined();
    expect(cell?.dataset.officeCellPropertyRevisionOmml).toBeTruthy();
  });

  test('live cell hMerge edits become reviewable when track changes is on', () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions({
        isTracking: () => true,
      }),
      content:
        '<table><tbody><tr><td data-office-cell-h-merge="continue"><p>Cell</p></td></tr></tbody></table>',
    });
    try {
      let cellPos: number | null = null;
      editor.state.doc.descendants((node, position) => {
        if (node.type.name === 'tableCell' && cellPos === null) {
          cellPos = position;
        }
      });
      expect(cellPos).not.toBeNull();
      editor.view.dispatch(
        editor.state.tr.setNodeMarkup(cellPos!, undefined, {
          ...editor.state.doc.nodeAt(cellPos!)!.attrs,
          hMerge: 'restart',
        }),
      );
      const changes = collectDocumentChanges(editor.state.doc).filter(
        (change) => change.kind === 'cell-formatting',
      );
      expect(changes).toHaveLength(1);
      const html = new DOMParser().parseFromString(
        editor.getHTML(),
        'text/html',
      );
      const cell = html.body.querySelector('td');
      expect(cell?.dataset.changeKind).toBe('cell-formatting');
      expect(
        parseDocumentCellFormatting(cell?.dataset.changeBefore),
      ).toMatchObject({
        hMerge: 'continue',
      });
      expect(cell?.dataset.officeCellHMerge).toBe('restart');
    } finally {
      editor.destroy();
    }
  });

  test('imports vMerge-only w:tcPrChange as a reviewable cell-formatting change', async () => {
    const source = await cellDocxWithVMergeChange({
      prior: 'restart',
      current: 'continue',
    });
    const imported = await importOfficeFile(
      new File([source], 'cell-formatting-v-merge.docx'),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const html = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    const cell = html.body.querySelector('td');
    expect(cell?.dataset.changeKind).toBe('cell-formatting');
    expect(cell?.dataset.officeCellPropertyRevisionOmml).toBeUndefined();
    expect(parseDocumentCellFormatting(cell?.dataset.changeBefore)).toEqual({
      vMerge: 'restart',
    });
    expect(cell?.dataset.officeCellVMerge).toBe('continue');

    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: imported.content.html,
    });
    try {
      const change = collectDocumentChanges(editor.state.doc)[0];
      expect(change?.kind).toBe('cell-formatting');
      expect(editor.commands.rejectDocumentChange(change?.id ?? '')).toBe(true);
      const rejected = new DOMParser().parseFromString(
        editor.getHTML(),
        'text/html',
      );
      const rejectedCell = rejected.body.querySelector('td');
      expect(rejectedCell?.dataset.changeKind).toBeUndefined();
      expect(rejectedCell?.dataset.officeCellVMerge).toBe('restart');
    } finally {
      editor.destroy();
    }
  });

  test('pending vMerge cell-formatting change round-trips as native w:tcPrChange', async () => {
    const source = await cellDocxWithVMergeChange({
      prior: 'continue',
      omitPriorVal: true,
    });
    const imported = await importOfficeFile(
      new File([source], 'cell-formatting-v-merge-roundtrip.docx'),
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
      directChild(descendants(exported, 'tc')[0], 'tcPr'),
      'tcPrChange',
    );
    expect(change).toBeTruthy();
    const priorVMerge = directChild(directChild(change!, 'tcPr'), 'vMerge');
    expect(
      priorVMerge?.getAttributeNS(WORD_NAMESPACE, 'val') ??
        priorVMerge?.getAttribute('w:val') ??
        priorVMerge?.getAttribute('val'),
    ).toBe('continue');
  });

  test('malformed vMerge w:tcPrChange stays on the opaque cell metadata path', async () => {
    const source = await cellDocxWithVMergeChange({
      prior: 'restart',
      malformedPrior: true,
    });
    const imported = await importOfficeFile(
      new File([source], 'cell-formatting-v-merge-opaque.docx'),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const html = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    const cell = html.body.querySelector('td');
    expect(cell?.dataset.changeKind).toBeUndefined();
    expect(cell?.dataset.officeCellPropertyRevisionOmml).toBeTruthy();
  });

  test('live cell vMerge edits become reviewable when track changes is on', () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions({
        isTracking: () => true,
      }),
      content:
        '<table><tbody><tr><td data-office-cell-v-merge="continue"><p>Cell</p></td></tr></tbody></table>',
    });
    try {
      let cellPos: number | null = null;
      editor.state.doc.descendants((node, position) => {
        if (node.type.name === 'tableCell' && cellPos === null) {
          cellPos = position;
        }
      });
      expect(cellPos).not.toBeNull();
      editor.view.dispatch(
        editor.state.tr.setNodeMarkup(cellPos!, undefined, {
          ...editor.state.doc.nodeAt(cellPos!)!.attrs,
          vMerge: 'restart',
        }),
      );
      const changes = collectDocumentChanges(editor.state.doc).filter(
        (change) => change.kind === 'cell-formatting',
      );
      expect(changes).toHaveLength(1);
      const html = new DOMParser().parseFromString(
        editor.getHTML(),
        'text/html',
      );
      const cell = html.body.querySelector('td');
      expect(cell?.dataset.changeKind).toBe('cell-formatting');
      expect(
        parseDocumentCellFormatting(cell?.dataset.changeBefore),
      ).toMatchObject({
        vMerge: 'continue',
      });
      expect(cell?.dataset.officeCellVMerge).toBe('restart');
    } finally {
      editor.destroy();
    }
  });

  test('imports gridSpan-only w:tcPrChange as a reviewable cell-formatting change', async () => {
    const source = await cellDocxWithGridSpanChange({
      prior: 2,
      current: 3,
    });
    const imported = await importOfficeFile(
      new File([source], 'cell-formatting-grid-span.docx'),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const html = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    const cell = html.body.querySelector('td');
    expect(cell?.dataset.changeKind).toBe('cell-formatting');
    expect(cell?.dataset.officeCellPropertyRevisionOmml).toBeUndefined();
    expect(parseDocumentCellFormatting(cell?.dataset.changeBefore)).toEqual({
      gridSpan: 2,
    });
    expect(cell?.dataset.officeCellGridSpan).toBe('3');

    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: imported.content.html,
    });
    try {
      const change = collectDocumentChanges(editor.state.doc)[0];
      expect(change?.kind).toBe('cell-formatting');
      expect(editor.commands.rejectDocumentChange(change?.id ?? '')).toBe(true);
      const rejected = new DOMParser().parseFromString(
        editor.getHTML(),
        'text/html',
      );
      const rejectedCell = rejected.body.querySelector('td');
      expect(rejectedCell?.dataset.changeKind).toBeUndefined();
      expect(rejectedCell?.dataset.officeCellGridSpan).toBe('2');
    } finally {
      editor.destroy();
    }
  });

  test('pending gridSpan cell-formatting change round-trips as native w:tcPrChange', async () => {
    const source = await cellDocxWithGridSpanChange({
      prior: 4,
      current: 2,
    });
    const imported = await importOfficeFile(
      new File([source], 'cell-formatting-grid-span-roundtrip.docx'),
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
      directChild(descendants(exported, 'tc')[0], 'tcPr'),
      'tcPrChange',
    );
    expect(change).toBeTruthy();
    const priorGridSpan = directChild(directChild(change!, 'tcPr'), 'gridSpan');
    expect(
      priorGridSpan?.getAttributeNS(WORD_NAMESPACE, 'val') ??
        priorGridSpan?.getAttribute('w:val') ??
        priorGridSpan?.getAttribute('val'),
    ).toBe('4');
  });

  test('malformed gridSpan w:tcPrChange stays on the opaque cell metadata path', async () => {
    const source = await cellDocxWithGridSpanChange({
      prior: 2,
      malformedPrior: true,
    });
    const imported = await importOfficeFile(
      new File([source], 'cell-formatting-grid-span-opaque.docx'),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const html = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    const cell = html.body.querySelector('td');
    expect(cell?.dataset.changeKind).toBeUndefined();
    expect(cell?.dataset.officeCellPropertyRevisionOmml).toBeTruthy();
  });

  test('live cell gridSpan edits become reviewable when track changes is on', () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions({
        isTracking: () => true,
      }),
      content:
        '<table><tbody><tr><td data-office-cell-grid-span="2"><p>Cell</p></td></tr></tbody></table>',
    });
    try {
      let cellPos: number | null = null;
      editor.state.doc.descendants((node, position) => {
        if (node.type.name === 'tableCell' && cellPos === null) {
          cellPos = position;
        }
      });
      expect(cellPos).not.toBeNull();
      editor.view.dispatch(
        editor.state.tr.setNodeMarkup(cellPos!, undefined, {
          ...editor.state.doc.nodeAt(cellPos!)!.attrs,
          gridSpan: 3,
        }),
      );
      const changes = collectDocumentChanges(editor.state.doc).filter(
        (change) => change.kind === 'cell-formatting',
      );
      expect(changes).toHaveLength(1);
      const html = new DOMParser().parseFromString(
        editor.getHTML(),
        'text/html',
      );
      const cell = html.body.querySelector('td');
      expect(cell?.dataset.changeKind).toBe('cell-formatting');
      expect(
        parseDocumentCellFormatting(cell?.dataset.changeBefore),
      ).toMatchObject({
        gridSpan: 2,
      });
      expect(cell?.dataset.officeCellGridSpan).toBe('3');
    } finally {
      editor.destroy();
    }
  });

});

async function cellDocxWithVAlignChange(options: {
  prior: 'top' | 'center' | 'bottom';
  current?: 'top' | 'center' | 'bottom';
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
  const cell = descendants(document, 'tc')[0];
  const properties =
    directChild(cell, 'tcPr') ??
    (() => {
      const created = document.createElementNS(WORD_NAMESPACE, 'w:tcPr');
      cell.insertBefore(created, cell.firstChild);
      return created;
    })();
  for (const existing of Array.from(properties.children).filter(
    (child) => child.localName === 'vAlign' || child.localName === 'tcPrChange',
  )) {
    existing.remove();
  }
  const currentAlign = options.current ?? 'bottom';
  const current = document.createElementNS(WORD_NAMESPACE, 'w:vAlign');
  current.setAttributeNS(WORD_NAMESPACE, 'w:val', currentAlign);
  properties.append(current);
  const change = document.createElementNS(WORD_NAMESPACE, 'w:tcPrChange');
  change.setAttributeNS(WORD_NAMESPACE, 'w:id', '31');
  change.setAttributeNS(WORD_NAMESPACE, 'w:author', 'Reviewer');
  change.setAttributeNS(WORD_NAMESPACE, 'w:date', '2026-09-08T00:00:00Z');
  const prior = document.createElementNS(WORD_NAMESPACE, 'w:tcPr');
  const priorAlign = document.createElementNS(WORD_NAMESPACE, 'w:vAlign');
  priorAlign.setAttributeNS(WORD_NAMESPACE, 'w:val', options.prior);
  prior.append(priorAlign);
  change.append(prior);
  properties.append(change);
  archive.file(
    'word/document.xml',
    new XMLSerializer().serializeToString(document),
  );
  return archive.generateAsync({ type: 'arraybuffer' });
}

async function cellDocxWithFillChange(options: {
  prior: string;
  current?: string;
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
  const cell = descendants(document, 'tc')[0];
  const properties =
    directChild(cell, 'tcPr') ??
    (() => {
      const created = document.createElementNS(WORD_NAMESPACE, 'w:tcPr');
      cell.insertBefore(created, cell.firstChild);
      return created;
    })();
  for (const existing of Array.from(properties.children).filter(
    (child) => child.localName === 'shd' || child.localName === 'tcPrChange',
  )) {
    existing.remove();
  }
  const currentFill = options.current ?? 'FFFFFF';
  const current = document.createElementNS(WORD_NAMESPACE, 'w:shd');
  current.setAttributeNS(WORD_NAMESPACE, 'w:val', 'clear');
  current.setAttributeNS(WORD_NAMESPACE, 'w:fill', currentFill);
  properties.append(current);
  const change = document.createElementNS(WORD_NAMESPACE, 'w:tcPrChange');
  change.setAttributeNS(WORD_NAMESPACE, 'w:id', '32');
  change.setAttributeNS(WORD_NAMESPACE, 'w:author', 'Reviewer');
  change.setAttributeNS(WORD_NAMESPACE, 'w:date', '2026-09-08T00:00:00Z');
  const prior = document.createElementNS(WORD_NAMESPACE, 'w:tcPr');
  const priorShading = document.createElementNS(WORD_NAMESPACE, 'w:shd');
  priorShading.setAttributeNS(WORD_NAMESPACE, 'w:val', 'clear');
  priorShading.setAttributeNS(WORD_NAMESPACE, 'w:fill', options.prior);
  prior.append(priorShading);
  change.append(prior);
  properties.append(change);
  archive.file(
    'word/document.xml',
    new XMLSerializer().serializeToString(document),
  );
  return archive.generateAsync({ type: 'arraybuffer' });
}

async function cellDocxWithMarginChange(options: {
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
  const cell = descendants(document, 'tc')[0];
  const properties =
    directChild(cell, 'tcPr') ??
    (() => {
      const created = document.createElementNS(WORD_NAMESPACE, 'w:tcPr');
      cell.insertBefore(created, cell.firstChild);
      return created;
    })();
  for (const existing of Array.from(properties.children).filter(
    (child) => child.localName === 'tcMar' || child.localName === 'tcPrChange',
  )) {
    existing.remove();
  }
  properties.append(createTcMar(document, options.currentTwips));
  const change = document.createElementNS(WORD_NAMESPACE, 'w:tcPrChange');
  change.setAttributeNS(WORD_NAMESPACE, 'w:id', '33');
  change.setAttributeNS(WORD_NAMESPACE, 'w:author', 'Reviewer');
  change.setAttributeNS(WORD_NAMESPACE, 'w:date', '2026-09-08T00:00:00Z');
  const prior = document.createElementNS(WORD_NAMESPACE, 'w:tcPr');
  prior.append(createTcMar(document, options.priorTwips));
  change.append(prior);
  properties.append(change);
  archive.file(
    'word/document.xml',
    new XMLSerializer().serializeToString(document),
  );
  return archive.generateAsync({ type: 'arraybuffer' });
}

function createTcMar(document: Document, twips: number): Element {
  const margins = document.createElementNS(WORD_NAMESPACE, 'w:tcMar');
  for (const side of ['top', 'left', 'bottom', 'right'] as const) {
    const edge = document.createElementNS(WORD_NAMESPACE, `w:${side}`);
    edge.setAttributeNS(WORD_NAMESPACE, 'w:type', 'dxa');
    edge.setAttributeNS(WORD_NAMESPACE, 'w:w', String(twips));
    margins.append(edge);
  }
  return margins;
}

async function cellDocxWithWidthChange(options: {
  priorTwips: number;
  currentTwips?: number;
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
  const cell = descendants(document, 'tc')[0];
  const properties =
    directChild(cell, 'tcPr') ??
    (() => {
      const created = document.createElementNS(WORD_NAMESPACE, 'w:tcPr');
      cell.insertBefore(created, cell.firstChild);
      return created;
    })();
  for (const existing of Array.from(properties.children).filter(
    (child) => child.localName === 'tcW' || child.localName === 'tcPrChange',
  )) {
    existing.remove();
  }
  const current = document.createElementNS(WORD_NAMESPACE, 'w:tcW');
  current.setAttributeNS(WORD_NAMESPACE, 'w:type', 'dxa');
  current.setAttributeNS(
    WORD_NAMESPACE,
    'w:w',
    String(options.currentTwips ?? options.priorTwips * 2),
  );
  properties.append(current);
  const change = document.createElementNS(WORD_NAMESPACE, 'w:tcPrChange');
  change.setAttributeNS(WORD_NAMESPACE, 'w:id', '41');
  change.setAttributeNS(WORD_NAMESPACE, 'w:author', 'Reviewer');
  change.setAttributeNS(WORD_NAMESPACE, 'w:date', '2026-09-08T00:00:00Z');
  const prior = document.createElementNS(WORD_NAMESPACE, 'w:tcPr');
  const priorWidth = document.createElementNS(WORD_NAMESPACE, 'w:tcW');
  priorWidth.setAttributeNS(WORD_NAMESPACE, 'w:type', 'dxa');
  priorWidth.setAttributeNS(WORD_NAMESPACE, 'w:w', String(options.priorTwips));
  prior.append(priorWidth);
  change.append(prior);
  properties.append(change);
  archive.file(
    'word/document.xml',
    new XMLSerializer().serializeToString(document),
  );
  return archive.generateAsync({ type: 'arraybuffer' });
}

async function cellDocxWithNoWrapChange(options: {
  prior: boolean;
  current?: boolean;
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
  const cell = descendants(document, 'tc')[0];
  const properties =
    directChild(cell, 'tcPr') ??
    (() => {
      const created = document.createElementNS(WORD_NAMESPACE, 'w:tcPr');
      cell.insertBefore(created, cell.firstChild);
      return created;
    })();
  for (const existing of Array.from(properties.children).filter(
    (child) => child.localName === 'noWrap' || child.localName === 'tcPrChange',
  )) {
    existing.remove();
  }
  const currentNoWrap = options.current ?? false;
  if (currentNoWrap) {
    properties.append(document.createElementNS(WORD_NAMESPACE, 'w:noWrap'));
  }
  const change = document.createElementNS(WORD_NAMESPACE, 'w:tcPrChange');
  change.setAttributeNS(WORD_NAMESPACE, 'w:id', '42');
  change.setAttributeNS(WORD_NAMESPACE, 'w:author', 'Reviewer');
  change.setAttributeNS(WORD_NAMESPACE, 'w:date', '2026-09-08T00:00:00Z');
  const prior = document.createElementNS(WORD_NAMESPACE, 'w:tcPr');
  if (options.prior) {
    prior.append(document.createElementNS(WORD_NAMESPACE, 'w:noWrap'));
  } else {
    const noWrap = document.createElementNS(WORD_NAMESPACE, 'w:noWrap');
    noWrap.setAttributeNS(WORD_NAMESPACE, 'w:val', '0');
    prior.append(noWrap);
  }
  change.append(prior);
  properties.append(change);
  archive.file(
    'word/document.xml',
    new XMLSerializer().serializeToString(document),
  );
  return archive.generateAsync({ type: 'arraybuffer' });
}

async function cellDocxWithTextDirectionChange(options: {
  prior: 'lrTb' | 'tbRl' | 'btLr' | 'lrTbV' | 'tbRlV' | 'tbLrV';
  current?: 'lrTb' | 'tbRl' | 'btLr' | 'lrTbV' | 'tbRlV' | 'tbLrV';
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
  const cell = descendants(document, 'tc')[0];
  const properties =
    directChild(cell, 'tcPr') ??
    (() => {
      const created = document.createElementNS(WORD_NAMESPACE, 'w:tcPr');
      cell.insertBefore(created, cell.firstChild);
      return created;
    })();
  for (const existing of Array.from(properties.children).filter(
    (child) =>
      child.localName === 'textDirection' || child.localName === 'tcPrChange',
  )) {
    existing.remove();
  }
  const current = document.createElementNS(WORD_NAMESPACE, 'w:textDirection');
  current.setAttributeNS(
    WORD_NAMESPACE,
    'w:val',
    options.current ?? options.prior,
  );
  properties.append(current);
  const change = document.createElementNS(WORD_NAMESPACE, 'w:tcPrChange');
  change.setAttributeNS(WORD_NAMESPACE, 'w:id', '43');
  change.setAttributeNS(WORD_NAMESPACE, 'w:author', 'Reviewer');
  change.setAttributeNS(WORD_NAMESPACE, 'w:date', '2026-09-08T00:00:00Z');
  const prior = document.createElementNS(WORD_NAMESPACE, 'w:tcPr');
  const priorDirection = document.createElementNS(
    WORD_NAMESPACE,
    'w:textDirection',
  );
  priorDirection.setAttributeNS(WORD_NAMESPACE, 'w:val', options.prior);
  prior.append(priorDirection);
  change.append(prior);
  properties.append(change);
  archive.file(
    'word/document.xml',
    new XMLSerializer().serializeToString(document),
  );
  return archive.generateAsync({ type: 'arraybuffer' });
}

async function cellDocxWithFitTextChange(options: {
  prior: boolean;
  current?: boolean;
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
  const cell = descendants(document, 'tc')[0];
  const properties =
    directChild(cell, 'tcPr') ??
    (() => {
      const created = document.createElementNS(WORD_NAMESPACE, 'w:tcPr');
      cell.insertBefore(created, cell.firstChild);
      return created;
    })();
  for (const existing of Array.from(properties.children).filter(
    (child) =>
      child.localName === 'tcFitText' || child.localName === 'tcPrChange',
  )) {
    existing.remove();
  }
  const currentFitText = options.current ?? false;
  if (currentFitText) {
    properties.append(document.createElementNS(WORD_NAMESPACE, 'w:tcFitText'));
  }
  const change = document.createElementNS(WORD_NAMESPACE, 'w:tcPrChange');
  change.setAttributeNS(WORD_NAMESPACE, 'w:id', '44');
  change.setAttributeNS(WORD_NAMESPACE, 'w:author', 'Reviewer');
  change.setAttributeNS(WORD_NAMESPACE, 'w:date', '2026-09-08T00:00:00Z');
  const prior = document.createElementNS(WORD_NAMESPACE, 'w:tcPr');
  if (options.prior) {
    prior.append(document.createElementNS(WORD_NAMESPACE, 'w:tcFitText'));
  } else {
    const fitText = document.createElementNS(WORD_NAMESPACE, 'w:tcFitText');
    fitText.setAttributeNS(WORD_NAMESPACE, 'w:val', '0');
    prior.append(fitText);
  }
  change.append(prior);
  properties.append(change);
  archive.file(
    'word/document.xml',
    new XMLSerializer().serializeToString(document),
  );
  return archive.generateAsync({ type: 'arraybuffer' });
}

async function cellDocxWithBordersChange(options: {
  prior: Record<
    string,
    { val: string; sz: string; color: string; themeColor?: string }
  >;
  current: Record<
    string,
    { val: string; sz: string; color: string; themeColor?: string }
  > | null;
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
  const cell = descendants(document, 'tc')[0];
  const properties =
    directChild(cell, 'tcPr') ??
    (() => {
      const created = document.createElementNS(WORD_NAMESPACE, 'w:tcPr');
      cell.insertBefore(created, cell.firstChild);
      return created;
    })();
  for (const existing of Array.from(properties.children).filter(
    (child) =>
      child.localName === 'tcBorders' || child.localName === 'tcPrChange',
  )) {
    existing.remove();
  }
  if (options.current) {
    properties.append(createCellBordersElement(document, options.current));
  }
  const change = document.createElementNS(WORD_NAMESPACE, 'w:tcPrChange');
  change.setAttributeNS(WORD_NAMESPACE, 'w:id', '47');
  change.setAttributeNS(WORD_NAMESPACE, 'w:author', 'Reviewer');
  change.setAttributeNS(WORD_NAMESPACE, 'w:date', '2026-09-10T00:00:00Z');
  const prior = document.createElementNS(WORD_NAMESPACE, 'w:tcPr');
  prior.append(createCellBordersElement(document, options.prior));
  change.append(prior);
  properties.append(change);
  archive.file(
    'word/document.xml',
    new XMLSerializer().serializeToString(document),
  );
  return archive.generateAsync({ type: 'arraybuffer' });
}

function createCellBordersElement(
  document: Document,
  edges: Record<
    string,
    { val: string; sz: string; color: string; themeColor?: string }
  >,
): Element {
  const borders = document.createElementNS(WORD_NAMESPACE, 'w:tcBorders');
  for (const [name, edge] of Object.entries(edges)) {
    const child = document.createElementNS(WORD_NAMESPACE, `w:${name}`);
    child.setAttributeNS(WORD_NAMESPACE, 'w:val', edge.val);
    child.setAttributeNS(WORD_NAMESPACE, 'w:sz', edge.sz);
    child.setAttributeNS(WORD_NAMESPACE, 'w:space', '0');
    child.setAttributeNS(WORD_NAMESPACE, 'w:color', edge.color);
    if (edge.themeColor) {
      child.setAttributeNS(WORD_NAMESPACE, 'w:themeColor', edge.themeColor);
    }
    borders.append(child);
  }
  return borders;
}

async function cellDocxWithCnfStyleChange(options: {
  prior: string;
  current?: string;
  malformedPrior?: boolean;
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
  const cell = descendants(document, 'tc')[0];
  const properties =
    directChild(cell, 'tcPr') ??
    (() => {
      const created = document.createElementNS(WORD_NAMESPACE, 'w:tcPr');
      cell.insertBefore(created, cell.firstChild);
      return created;
    })();
  for (const existing of Array.from(properties.children).filter(
    (child) =>
      child.localName === 'cnfStyle' || child.localName === 'tcPrChange',
  )) {
    existing.remove();
  }
  const current = document.createElementNS(WORD_NAMESPACE, 'w:cnfStyle');
  current.setAttributeNS(
    WORD_NAMESPACE,
    'w:val',
    options.current ?? options.prior,
  );
  properties.append(current);
  const change = document.createElementNS(WORD_NAMESPACE, 'w:tcPrChange');
  change.setAttributeNS(WORD_NAMESPACE, 'w:id', '46');
  change.setAttributeNS(WORD_NAMESPACE, 'w:author', 'Reviewer');
  change.setAttributeNS(WORD_NAMESPACE, 'w:date', '2026-09-08T00:00:00Z');
  const prior = document.createElementNS(WORD_NAMESPACE, 'w:tcPr');
  const priorCnfStyle = document.createElementNS(WORD_NAMESPACE, 'w:cnfStyle');
  priorCnfStyle.setAttributeNS(
    WORD_NAMESPACE,
    'w:val',
    options.malformedPrior ? '10000000000' : options.prior,
  );
  prior.append(priorCnfStyle);
  change.append(prior);
  properties.append(change);
  archive.file(
    'word/document.xml',
    new XMLSerializer().serializeToString(document),
  );
  return archive.generateAsync({ type: 'arraybuffer' });
}


async function cellDocxWithHMergeChange(options: {
  prior: 'restart' | 'continue';
  current?: 'restart' | 'continue';
  omitPriorVal?: boolean;
  malformedPrior?: boolean;
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
  const cell = descendants(document, 'tc')[0];
  const properties =
    directChild(cell, 'tcPr') ??
    (() => {
      const created = document.createElementNS(WORD_NAMESPACE, 'w:tcPr');
      cell.insertBefore(created, cell.firstChild);
      return created;
    })();
  for (const existing of Array.from(properties.children).filter(
    (child) =>
      child.localName === 'hMerge' || child.localName === 'tcPrChange',
  )) {
    existing.remove();
  }
  const current = document.createElementNS(WORD_NAMESPACE, 'w:hMerge');
  current.setAttributeNS(
    WORD_NAMESPACE,
    'w:val',
    options.current ?? options.prior,
  );
  properties.append(current);
  const change = document.createElementNS(WORD_NAMESPACE, 'w:tcPrChange');
  change.setAttributeNS(WORD_NAMESPACE, 'w:id', '47');
  change.setAttributeNS(WORD_NAMESPACE, 'w:author', 'Reviewer');
  change.setAttributeNS(WORD_NAMESPACE, 'w:date', '2026-09-08T00:00:00Z');
  const prior = document.createElementNS(WORD_NAMESPACE, 'w:tcPr');
  const priorHMerge = document.createElementNS(WORD_NAMESPACE, 'w:hMerge');
  if (options.malformedPrior) {
    priorHMerge.setAttributeNS(WORD_NAMESPACE, 'w:val', 'diagonal');
  } else if (!options.omitPriorVal) {
    priorHMerge.setAttributeNS(WORD_NAMESPACE, 'w:val', options.prior);
  }
  prior.append(priorHMerge);
  change.append(prior);
  properties.append(change);
  archive.file(
    'word/document.xml',
    new XMLSerializer().serializeToString(document),
  );
  return archive.generateAsync({ type: 'arraybuffer' });
}

async function cellDocxWithGridSpanChange(options: {
  prior: number;
  current?: number;
  malformedPrior?: boolean;
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
  const cell = descendants(document, 'tc')[0];
  const properties =
    directChild(cell, 'tcPr') ??
    (() => {
      const created = document.createElementNS(WORD_NAMESPACE, 'w:tcPr');
      cell.insertBefore(created, cell.firstChild);
      return created;
    })();
  for (const existing of Array.from(properties.children).filter(
    (child) =>
      child.localName === 'gridSpan' || child.localName === 'tcPrChange',
  )) {
    existing.remove();
  }
  const current = document.createElementNS(WORD_NAMESPACE, 'w:gridSpan');
  current.setAttributeNS(
    WORD_NAMESPACE,
    'w:val',
    String(options.current ?? options.prior),
  );
  properties.append(current);
  const change = document.createElementNS(WORD_NAMESPACE, 'w:tcPrChange');
  change.setAttributeNS(WORD_NAMESPACE, 'w:id', '49');
  change.setAttributeNS(WORD_NAMESPACE, 'w:author', 'Reviewer');
  change.setAttributeNS(WORD_NAMESPACE, 'w:date', '2026-09-08T00:00:00Z');
  const prior = document.createElementNS(WORD_NAMESPACE, 'w:tcPr');
  const priorGridSpan = document.createElementNS(WORD_NAMESPACE, 'w:gridSpan');
  if (options.malformedPrior) {
    priorGridSpan.setAttributeNS(WORD_NAMESPACE, 'w:val', '0');
  } else {
    priorGridSpan.setAttributeNS(
      WORD_NAMESPACE,
      'w:val',
      String(options.prior),
    );
  }
  prior.append(priorGridSpan);
  change.append(prior);
  properties.append(change);
  archive.file(
    'word/document.xml',
    new XMLSerializer().serializeToString(document),
  );
  return archive.generateAsync({ type: 'arraybuffer' });
}

async function cellDocxWithVMergeChange(options: {
  prior: 'restart' | 'continue';
  current?: 'restart' | 'continue';
  omitPriorVal?: boolean;
  malformedPrior?: boolean;
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
  const cell = descendants(document, 'tc')[0];
  const properties =
    directChild(cell, 'tcPr') ??
    (() => {
      const created = document.createElementNS(WORD_NAMESPACE, 'w:tcPr');
      cell.insertBefore(created, cell.firstChild);
      return created;
    })();
  for (const existing of Array.from(properties.children).filter(
    (child) =>
      child.localName === 'vMerge' || child.localName === 'tcPrChange',
  )) {
    existing.remove();
  }
  const current = document.createElementNS(WORD_NAMESPACE, 'w:vMerge');
  current.setAttributeNS(
    WORD_NAMESPACE,
    'w:val',
    options.current ?? options.prior,
  );
  properties.append(current);
  const change = document.createElementNS(WORD_NAMESPACE, 'w:tcPrChange');
  change.setAttributeNS(WORD_NAMESPACE, 'w:id', '48');
  change.setAttributeNS(WORD_NAMESPACE, 'w:author', 'Reviewer');
  change.setAttributeNS(WORD_NAMESPACE, 'w:date', '2026-09-08T00:00:00Z');
  const prior = document.createElementNS(WORD_NAMESPACE, 'w:tcPr');
  const priorVMerge = document.createElementNS(WORD_NAMESPACE, 'w:vMerge');
  if (options.malformedPrior) {
    priorVMerge.setAttributeNS(WORD_NAMESPACE, 'w:val', 'diagonal');
  } else if (!options.omitPriorVal) {
    priorVMerge.setAttributeNS(WORD_NAMESPACE, 'w:val', options.prior);
  }
  prior.append(priorVMerge);
  change.append(prior);
  properties.append(change);
  archive.file(
    'word/document.xml',
    new XMLSerializer().serializeToString(document),
  );
  return archive.generateAsync({ type: 'arraybuffer' });
}

async function cellDocxWithHideMarkChange(options: {
  prior: boolean;
  current?: boolean;
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
  const cell = descendants(document, 'tc')[0];
  const properties =
    directChild(cell, 'tcPr') ??
    (() => {
      const created = document.createElementNS(WORD_NAMESPACE, 'w:tcPr');
      cell.insertBefore(created, cell.firstChild);
      return created;
    })();
  for (const existing of Array.from(properties.children).filter(
    (child) =>
      child.localName === 'hideMark' || child.localName === 'tcPrChange',
  )) {
    existing.remove();
  }
  const currentHideMark = options.current ?? false;
  if (currentHideMark) {
    properties.append(document.createElementNS(WORD_NAMESPACE, 'w:hideMark'));
  }
  const change = document.createElementNS(WORD_NAMESPACE, 'w:tcPrChange');
  change.setAttributeNS(WORD_NAMESPACE, 'w:id', '45');
  change.setAttributeNS(WORD_NAMESPACE, 'w:author', 'Reviewer');
  change.setAttributeNS(WORD_NAMESPACE, 'w:date', '2026-09-08T00:00:00Z');
  const prior = document.createElementNS(WORD_NAMESPACE, 'w:tcPr');
  if (options.prior) {
    prior.append(document.createElementNS(WORD_NAMESPACE, 'w:hideMark'));
  } else {
    const hideMark = document.createElementNS(WORD_NAMESPACE, 'w:hideMark');
    hideMark.setAttributeNS(WORD_NAMESPACE, 'w:val', '0');
    prior.append(hideMark);
  }
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
