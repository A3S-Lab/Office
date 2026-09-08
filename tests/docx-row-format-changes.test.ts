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
    expect(
      directChild(directChild(change, 'trPr'), 'cantSplit'),
    ).toBeTruthy();
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
        height: { value: 24, rule: 'atLeast' },
      });
      expect(row?.dataset.officeRowHeight).toBe('40');
      expect(row?.dataset.officeRowHeightRule).toBe('exact');
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
    (child) =>
      child.localName === 'hidden' || child.localName === 'trPrChange',
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
  current.setAttributeNS(
    WORD_NAMESPACE,
    'w:val',
    options.current ?? 'left',
  );
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
