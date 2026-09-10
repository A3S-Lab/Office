import { Editor } from '@tiptap/core';
import { describe, expect, test } from '@rstest/core';
import JSZip from 'jszip';
import {
  createArtifact,
  createArtifactBlob,
  importOfficeFile,
} from '../src/core';
import { collectDocumentChanges } from '../src/internal/features/work/work-document-changes';
import { parseDocumentSectionFormatting } from '../src/internal/features/work/work-document-section-format-changes';
import { createWorkDocumentExtensions } from '../src/internal/features/work/work-document-extensions';
import { activeDocumentSection } from '../src/internal/features/work/work-document-section-editor';
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

describe('DOCX section-formatting revisions', () => {
  test('imports orientation-only w:sectPrChange as a reviewable section-formatting change', async () => {
    const source = await sectionDocxWithOrientationChange({
      prior: 'landscape',
    });
    const imported = await importOfficeFile(
      new File([source], 'section-formatting.docx'),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const html = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    const section = html.body.querySelector('section[data-document-section]');
    expect(section?.getAttribute('data-change-kind')).toBe(
      'section-formatting',
    );
    expect(section?.dataset.sectionPropertyRevisionOmml).toBeFalsy();
    expect(
      parseDocumentSectionFormatting(
        section?.getAttribute('data-change-before'),
      ),
    ).toEqual({ orientation: 'landscape' });

    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: imported.content.html,
    });
    try {
      const changes = collectDocumentChanges(editor.state.doc);
      expect(changes).toHaveLength(1);
      expect(changes[0]?.kind).toBe('section-formatting');
      expect(editor.commands.acceptDocumentChange(changes[0]?.id ?? '')).toBe(
        true,
      );
      expect(collectDocumentChanges(editor.state.doc)).toHaveLength(0);
    } finally {
      editor.destroy();
    }
  });

  test('reject restores prior orientation and drops the pending change', async () => {
    const source = await sectionDocxWithOrientationChange({
      prior: 'landscape',
      current: 'portrait',
    });
    const imported = await importOfficeFile(
      new File([source], 'section-formatting-reject.docx'),
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
      const section = html.body.querySelector('section[data-document-section]');
      expect(section?.getAttribute('data-change-kind')).toBeNull();
      expect(section?.dataset.sectionOrientation).toBe('landscape');
    } finally {
      editor.destroy();
    }
  });

  test('pending section-formatting change round-trips as native w:sectPrChange', async () => {
    const source = await sectionDocxWithOrientationChange({
      prior: 'landscape',
    });
    const imported = await importOfficeFile(
      new File([source], 'section-formatting-roundtrip.docx'),
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
    const section = descendants(document, 'sectPr').find(
      (element) => element.parentElement?.localName !== 'sectPrChange',
    );
    const change = directChild(section!, 'sectPrChange');
    expect(change).toBeTruthy();
    expect(wordAttribute(change!, 'author')).toBe('Reviewer');
    const priorSize = directChild(directChild(change, 'sectPr'), 'pgSz');
    expect(
      priorSize?.getAttributeNS(WORD_NAMESPACE, 'orient') ??
        priorSize?.getAttribute('w:orient') ??
        priorSize?.getAttribute('orient'),
    ).toBe('landscape');
  });

  test('reports orientation-only section-property revisions as reviewable diagnostics', async () => {
    const source = await sectionDocxWithOrientationChange({
      prior: 'landscape',
    });
    const summary = await analyzeDocxCompatibility(
      new File([source], 'section-formatting-diag.docx'),
      [],
    );
    expect(
      summary.issues.some(
        (issue) => issue.code === 'docx.revisions.section-formatting',
      ),
    ).toBe(true);
  });

  test('live section orientation edits become reviewable when track changes is on', () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions({
        isTracking: () => true,
      }),
      content: [
        '<section data-document-section="true" data-section-id="section-1"',
        ' data-section-orientation="portrait">',
        '<p>Body</p>',
        '</section>',
      ].join(''),
    });
    try {
      const active = activeDocumentSection(editor);
      expect(active).not.toBeNull();
      if (!active) throw new Error('Expected an active document section.');
      expect(
        editor.commands.updateActiveDocumentSection({
          ...active.layout,
          orientation: 'landscape',
        }),
      ).toBe(true);
      const changes = collectDocumentChanges(editor.state.doc);
      expect(changes).toHaveLength(1);
      expect(changes[0]?.kind).toBe('section-formatting');
      const html = new DOMParser().parseFromString(
        editor.getHTML(),
        'text/html',
      );
      const section = html.body.querySelector('section[data-document-section]');
      expect(section?.getAttribute('data-change-kind')).toBe(
        'section-formatting',
      );
      expect(
        parseDocumentSectionFormatting(
          section?.getAttribute('data-change-before'),
        ),
      ).toMatchObject({ orientation: 'portrait' });
      expect(
        parseDocumentSectionFormatting(
          section?.getAttribute('data-change-before'),
        )?.pageMargins,
      ).toBeTruthy();
      expect(section?.dataset.sectionOrientation).toBe('landscape');
    } finally {
      editor.destroy();
    }
  });

  test('imports pgMar-only w:sectPrChange as a reviewable section-formatting change', async () => {
    const source = await sectionDocxWithMarginChange({
      prior: {
        top: 1440,
        right: 1440,
        bottom: 1440,
        left: 1440,
        header: 708,
        footer: 708,
        gutter: 0,
      },
      current: {
        top: 720,
        right: 720,
        bottom: 720,
        left: 720,
        header: 708,
        footer: 708,
        gutter: 0,
      },
    });
    const imported = await importOfficeFile(
      new File([source], 'section-margin-formatting.docx'),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const html = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    const section = html.body.querySelector('section[data-document-section]');
    expect(section?.getAttribute('data-change-kind')).toBe(
      'section-formatting',
    );
    expect(section?.dataset.sectionPropertyRevisionOmml).toBeFalsy();
    expect(
      parseDocumentSectionFormatting(
        section?.getAttribute('data-change-before'),
      ),
    ).toEqual({
      pageMargins: {
        top: 1440,
        right: 1440,
        bottom: 1440,
        left: 1440,
        header: 708,
        footer: 708,
        gutter: 0,
      },
    });

    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: imported.content.html,
    });
    try {
      const change = collectDocumentChanges(editor.state.doc)[0];
      expect(change?.kind).toBe('section-formatting');
      expect(editor.commands.rejectDocumentChange(change?.id ?? '')).toBe(true);
      const rejected = new DOMParser().parseFromString(
        editor.getHTML(),
        'text/html',
      );
      const rejectedSection = rejected.body.querySelector(
        'section[data-document-section]',
      );
      expect(rejectedSection?.getAttribute('data-change-kind')).toBeNull();
      expect(rejectedSection?.dataset.sectionPageMargins).toContain(
        '"top":1440',
      );
    } finally {
      editor.destroy();
    }
  });

  test('live section page-margin edits become reviewable when track changes is on', () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions({
        isTracking: () => true,
      }),
      content: [
        '<section data-document-section="true" data-section-id="section-1"',
        ' data-section-orientation="portrait"',
        ' data-section-page-margins=\'{"top":1440,"right":1440,"bottom":1440,"left":1440,"header":708,"footer":708,"gutter":0}\'>',
        '<p>Body</p>',
        '</section>',
      ].join(''),
    });
    try {
      const active = activeDocumentSection(editor);
      expect(active).not.toBeNull();
      if (!active) throw new Error('Expected an active document section.');
      expect(
        editor.commands.updateActiveDocumentSection({
          ...active.layout,
          pageMargins: {
            top: 720,
            right: 720,
            bottom: 720,
            left: 720,
            header: 708,
            footer: 708,
            gutter: 0,
          },
          margins: { top: 12.7, right: 12.7, bottom: 12.7, left: 12.7 },
        }),
      ).toBe(true);
      const changes = collectDocumentChanges(editor.state.doc);
      expect(changes).toHaveLength(1);
      expect(changes[0]?.kind).toBe('section-formatting');
      const html = new DOMParser().parseFromString(
        editor.getHTML(),
        'text/html',
      );
      const section = html.body.querySelector('section[data-document-section]');
      expect(section?.getAttribute('data-change-kind')).toBe(
        'section-formatting',
      );
      expect(
        parseDocumentSectionFormatting(
          section?.getAttribute('data-change-before'),
        )?.pageMargins?.top,
      ).toBe(1440);
    } finally {
      editor.destroy();
    }
  });

  test('imports full w:pgSz w:sectPrChange as a reviewable section-formatting change', async () => {
    const source = await sectionDocxWithPageGeometryChange({
      prior: { width: 12240, height: 15840, orientation: 'portrait' },
      current: { width: 16838, height: 11906, orientation: 'landscape' },
    });
    const imported = await importOfficeFile(
      new File([source], 'section-page-geometry-formatting.docx'),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const html = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    const section = html.body.querySelector('section[data-document-section]');
    expect(section?.getAttribute('data-change-kind')).toBe(
      'section-formatting',
    );
    expect(section?.dataset.sectionPropertyRevisionOmml).toBeFalsy();
    expect(
      parseDocumentSectionFormatting(
        section?.getAttribute('data-change-before'),
      ),
    ).toEqual({
      pageGeometry: {
        width: 12240,
        height: 15840,
        orientation: 'portrait',
      },
    });

    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: imported.content.html,
    });
    try {
      const change = collectDocumentChanges(editor.state.doc)[0];
      expect(change?.kind).toBe('section-formatting');
      expect(editor.commands.rejectDocumentChange(change?.id ?? '')).toBe(true);
      const rejected = new DOMParser().parseFromString(
        editor.getHTML(),
        'text/html',
      );
      const rejectedSection = rejected.body.querySelector(
        'section[data-document-section]',
      );
      expect(rejectedSection?.getAttribute('data-change-kind')).toBeNull();
      expect(rejectedSection?.dataset.sectionOrientation).toBe('portrait');
      expect(rejectedSection?.dataset.sectionPageGeometry).toContain(
        '"width":12240',
      );
      expect(rejectedSection?.dataset.sectionPageGeometry).toContain(
        '"height":15840',
      );
    } finally {
      editor.destroy();
    }
  });

  test('live section page-geometry edits become reviewable when track changes is on', () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions({
        isTracking: () => true,
      }),
      content: [
        '<section data-document-section="true" data-section-id="section-1"',
        ' data-section-orientation="portrait"',
        ' data-section-page-geometry=\'{"width":11906,"height":16838,"orientation":"portrait"}\'>',
        '<p>Body</p>',
        '</section>',
      ].join(''),
    });
    try {
      const active = activeDocumentSection(editor);
      expect(active).not.toBeNull();
      if (!active) throw new Error('Expected an active document section.');
      expect(
        editor.commands.updateActiveDocumentSection({
          ...active.layout,
          pageGeometry: {
            width: 12240,
            height: 15840,
            orientation: 'portrait',
          },
          pageSize: 'letter',
        }),
      ).toBe(true);
      const changes = collectDocumentChanges(editor.state.doc);
      expect(changes).toHaveLength(1);
      expect(changes[0]?.kind).toBe('section-formatting');
      const html = new DOMParser().parseFromString(
        editor.getHTML(),
        'text/html',
      );
      const section = html.body.querySelector('section[data-document-section]');
      expect(section?.getAttribute('data-change-kind')).toBe(
        'section-formatting',
      );
      expect(
        parseDocumentSectionFormatting(
          section?.getAttribute('data-change-before'),
        )?.pageGeometry,
      ).toEqual({
        width: 11906,
        height: 16838,
        orientation: 'portrait',
      });
      expect(section?.dataset.sectionPageGeometry).toContain('"width":12240');
    } finally {
      editor.destroy();
    }
  });

  test('imports paperSrc-only w:sectPrChange as a reviewable section-formatting change', async () => {
    const source = await sectionDocxWithPaperSourceChange({
      prior: { first: 1, other: 2 },
      current: { first: 5, other: 6 },
    });
    const imported = await importOfficeFile(
      new File([source], 'section-paper-source-formatting.docx'),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const html = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    const section = html.body.querySelector('section[data-document-section]');
    expect(section?.getAttribute('data-change-kind')).toBe(
      'section-formatting',
    );
    expect(section?.dataset.sectionPropertyRevisionOmml).toBeFalsy();
    expect(
      parseDocumentSectionFormatting(
        section?.getAttribute('data-change-before'),
      ),
    ).toEqual({
      paperSource: { first: 1, other: 2 },
    });

    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: imported.content.html,
    });
    try {
      const change = collectDocumentChanges(editor.state.doc)[0];
      expect(change?.kind).toBe('section-formatting');
      expect(editor.commands.rejectDocumentChange(change?.id ?? '')).toBe(true);
      const rejected = new DOMParser().parseFromString(
        editor.getHTML(),
        'text/html',
      );
      const rejectedSection = rejected.body.querySelector(
        'section[data-document-section]',
      );
      expect(rejectedSection?.getAttribute('data-change-kind')).toBeNull();
      expect(rejectedSection?.dataset.sectionPaperSource).toContain(
        '"first":1',
      );
      expect(rejectedSection?.dataset.sectionPaperSource).toContain(
        '"other":2',
      );
    } finally {
      editor.destroy();
    }
  });

  test('live section paper-source edits become reviewable when track changes is on', () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions({
        isTracking: () => true,
      }),
      content: [
        '<section data-document-section="true" data-section-id="section-1"',
        ' data-section-orientation="portrait"',
        ' data-section-paper-source=\'{"first":1,"other":2}\'>',
        '<p>Body</p>',
        '</section>',
      ].join(''),
    });
    try {
      const active = activeDocumentSection(editor);
      expect(active).not.toBeNull();
      if (!active) throw new Error('Expected an active document section.');
      expect(
        editor.commands.updateActiveDocumentSection({
          ...active.layout,
          paperSource: { first: 5, other: 6 },
        }),
      ).toBe(true);
      const changes = collectDocumentChanges(editor.state.doc);
      expect(changes).toHaveLength(1);
      expect(changes[0]?.kind).toBe('section-formatting');
      const html = new DOMParser().parseFromString(
        editor.getHTML(),
        'text/html',
      );
      const section = html.body.querySelector('section[data-document-section]');
      expect(section?.getAttribute('data-change-kind')).toBe(
        'section-formatting',
      );
      expect(
        parseDocumentSectionFormatting(
          section?.getAttribute('data-change-before'),
        )?.paperSource,
      ).toEqual({ first: 1, other: 2 });
      expect(section?.dataset.sectionPaperSource).toContain('"first":5');
    } finally {
      editor.destroy();
    }
  });

  test('imports docGrid-only w:sectPrChange as a reviewable section-formatting change', async () => {
    const source = await sectionDocxWithDocGridChange({
      prior: { type: 'lines', linePitch: 18 },
      current: { type: 'linesAndChars', linePitch: 24 },
    });
    const imported = await importOfficeFile(
      new File([source], 'section-docgrid-formatting.docx'),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const html = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    const section = html.body.querySelector('section[data-document-section]');
    expect(section?.getAttribute('data-change-kind')).toBe(
      'section-formatting',
    );
    expect(
      section?.getAttribute('data-section-property-revision-omml'),
    ).toBeNull();
    expect(
      parseDocumentSectionFormatting(
        section?.getAttribute('data-change-before'),
      ),
    ).toEqual({ documentGrid: { type: 'lines', linePitch: 18 } });
    expect(section?.dataset.sectionDocumentGridType).toBe('linesAndChars');
    expect(section?.dataset.sectionDocumentGridLinePitch).toBe('24');
  });

  test('pending docGrid section-formatting change round-trips as native w:sectPrChange', async () => {
    const source = await sectionDocxWithDocGridChange({
      prior: { type: 'lines', linePitch: 18 },
      current: { type: 'linesAndChars', linePitch: 24 },
    });
    const imported = await importOfficeFile(
      new File([source], 'section-docgrid-roundtrip.docx'),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const output = await JSZip.loadAsync(
      await (await createArtifactBlob(imported)).arrayBuffer(),
    );
    const exported = await xmlEntry(output, 'word/document.xml');
    const section = descendants(exported, 'sectPr').find(
      (element) => element.parentElement?.localName !== 'sectPrChange',
    );
    const change = directChild(section!, 'sectPrChange');
    expect(change).toBeTruthy();
    const priorDocGrid = directChild(directChild(change!, 'sectPr'), 'docGrid');
    expect(priorDocGrid).toBeTruthy();
    expect(
      priorDocGrid?.getAttributeNS(WORD_NAMESPACE, 'type') ??
        priorDocGrid?.getAttribute('w:type') ??
        priorDocGrid?.getAttribute('type'),
    ).toBe('lines');
    expect(
      priorDocGrid?.getAttributeNS(WORD_NAMESPACE, 'linePitch') ??
        priorDocGrid?.getAttribute('w:linePitch') ??
        priorDocGrid?.getAttribute('linePitch'),
    ).toBe('360');
  });

  test('reject restores prior docGrid and drops the pending change', async () => {
    const source = await sectionDocxWithDocGridChange({
      prior: { type: 'lines', linePitch: 18 },
      current: { type: 'linesAndChars', linePitch: 24 },
    });
    const imported = await importOfficeFile(
      new File([source], 'section-docgrid-reject.docx'),
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
      const section = html.body.querySelector('section[data-document-section]');
      expect(section?.getAttribute('data-change-kind')).toBeNull();
      expect(section?.dataset.sectionDocumentGridType).toBe('lines');
      expect(section?.dataset.sectionDocumentGridLinePitch).toBe('18');
    } finally {
      editor.destroy();
    }
  });

  test('live section docGrid edits become reviewable when track changes is on', () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions({
        isTracking: () => true,
      }),
      content: [
        '<section data-document-section="true" data-section-id="section-1"',
        ' data-section-orientation="portrait"',
        ' data-section-document-grid-type="lines"',
        ' data-section-document-grid-line-pitch="18">',
        '<p>Body</p>',
        '</section>',
      ].join(''),
    });
    try {
      const active = activeDocumentSection(editor);
      expect(active).not.toBeNull();
      if (!active) throw new Error('Expected an active document section.');
      expect(
        editor.commands.updateActiveDocumentSection({
          ...active.layout,
          documentGrid: { type: 'linesAndChars', linePitch: 24 },
        }),
      ).toBe(true);
      const changes = collectDocumentChanges(editor.state.doc);
      expect(changes).toHaveLength(1);
      expect(changes[0]?.kind).toBe('section-formatting');
      const html = new DOMParser().parseFromString(
        editor.getHTML(),
        'text/html',
      );
      const section = html.body.querySelector('section[data-document-section]');
      expect(section?.getAttribute('data-change-kind')).toBe(
        'section-formatting',
      );
      expect(
        parseDocumentSectionFormatting(
          section?.getAttribute('data-change-before'),
        ),
      ).toMatchObject({ documentGrid: { type: 'lines', linePitch: 18 } });
      expect(section?.dataset.sectionDocumentGridType).toBe('linesAndChars');
      expect(section?.dataset.sectionDocumentGridLinePitch).toBe('24');
    } finally {
      editor.destroy();
    }
  });

  test('imports lnNumType-only w:sectPrChange as a reviewable section-formatting change', async () => {
    const source = await sectionDocxWithLnNumTypeChange({
      prior: { countBy: 1, start: 1, restart: 'newPage' },
      current: { countBy: 2, start: 5, distance: 720, restart: 'newSection' },
    });
    const imported = await importOfficeFile(
      new File([source], 'section-lnnumtype-formatting.docx'),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const html = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    const section = html.body.querySelector('section[data-document-section]');
    expect(section?.getAttribute('data-change-kind')).toBe(
      'section-formatting',
    );
    expect(
      section?.getAttribute('data-section-property-revision-omml'),
    ).toBeNull();
    expect(
      parseDocumentSectionFormatting(
        section?.getAttribute('data-change-before'),
      ),
    ).toEqual({
      lnNumType: { countBy: 1, start: 1, restart: 'newPage' },
    });
    expect(section?.dataset.sectionLnNumCountBy).toBe('2');
    expect(section?.dataset.sectionLnNumStart).toBe('5');
    expect(section?.dataset.sectionLnNumDistance).toBe('720');
    expect(section?.dataset.sectionLnNumRestart).toBe('newSection');
  });

  test('pending lnNumType section-formatting change round-trips as native w:sectPrChange', async () => {
    const source = await sectionDocxWithLnNumTypeChange({
      prior: { countBy: 1, start: 1, restart: 'newPage' },
      current: { countBy: 2, start: 5, distance: 720, restart: 'newSection' },
    });
    const imported = await importOfficeFile(
      new File([source], 'section-lnnumtype-roundtrip.docx'),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const output = await JSZip.loadAsync(
      await (await createArtifactBlob(imported)).arrayBuffer(),
    );
    const exported = await xmlEntry(output, 'word/document.xml');
    const section = descendants(exported, 'sectPr').find(
      (element) => element.parentElement?.localName !== 'sectPrChange',
    );
    const change = directChild(section!, 'sectPrChange');
    expect(change).toBeTruthy();
    const priorLnNum = directChild(directChild(change!, 'sectPr'), 'lnNumType');
    expect(priorLnNum).toBeTruthy();
    expect(
      priorLnNum?.getAttributeNS(WORD_NAMESPACE, 'countBy') ??
        priorLnNum?.getAttribute('w:countBy') ??
        priorLnNum?.getAttribute('countBy'),
    ).toBe('1');
    expect(
      priorLnNum?.getAttributeNS(WORD_NAMESPACE, 'start') ??
        priorLnNum?.getAttribute('w:start') ??
        priorLnNum?.getAttribute('start'),
    ).toBe('1');
    expect(
      priorLnNum?.getAttributeNS(WORD_NAMESPACE, 'restart') ??
        priorLnNum?.getAttribute('w:restart') ??
        priorLnNum?.getAttribute('restart'),
    ).toBe('newPage');
    const currentLnNum = directChild(section!, 'lnNumType');
    expect(
      currentLnNum?.getAttributeNS(WORD_NAMESPACE, 'countBy') ??
        currentLnNum?.getAttribute('w:countBy') ??
        currentLnNum?.getAttribute('countBy'),
    ).toBe('2');
    expect(
      currentLnNum?.getAttributeNS(WORD_NAMESPACE, 'distance') ??
        currentLnNum?.getAttribute('w:distance') ??
        currentLnNum?.getAttribute('distance'),
    ).toBe('720');
  });

  test('reject restores prior lnNumType and drops the pending change', async () => {
    const source = await sectionDocxWithLnNumTypeChange({
      prior: { countBy: 1, start: 1, restart: 'newPage' },
      current: { countBy: 2, start: 5, distance: 720, restart: 'newSection' },
    });
    const imported = await importOfficeFile(
      new File([source], 'section-lnnumtype-reject.docx'),
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
      const section = html.body.querySelector('section[data-document-section]');
      expect(section?.getAttribute('data-change-kind')).toBeNull();
      expect(section?.dataset.sectionLnNumCountBy).toBe('1');
      expect(section?.dataset.sectionLnNumStart).toBe('1');
      expect(section?.dataset.sectionLnNumRestart).toBe('newPage');
    } finally {
      editor.destroy();
    }
  });

  test('live section lnNumType edits become reviewable when track changes is on', () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions({
        isTracking: () => true,
      }),
      content: [
        '<section data-document-section="true" data-section-id="section-1"',
        ' data-section-orientation="portrait"',
        ' data-section-ln-num-count-by="1"',
        ' data-section-ln-num-start="1"',
        ' data-section-ln-num-restart="newPage">',
        '<p>Body</p>',
        '</section>',
      ].join(''),
    });
    try {
      const active = activeDocumentSection(editor);
      expect(active).not.toBeNull();
      if (!active) throw new Error('Expected an active document section.');
      expect(
        editor.commands.updateActiveDocumentSection({
          ...active.layout,
          lnNumType: {
            countBy: 2,
            start: 5,
            distance: 720,
            restart: 'continuous',
          },
        }),
      ).toBe(true);
      const changes = collectDocumentChanges(editor.state.doc);
      expect(changes).toHaveLength(1);
      expect(changes[0]?.kind).toBe('section-formatting');
      const html = new DOMParser().parseFromString(
        editor.getHTML(),
        'text/html',
      );
      const section = html.body.querySelector('section[data-document-section]');
      expect(section?.getAttribute('data-change-kind')).toBe(
        'section-formatting',
      );
      expect(
        parseDocumentSectionFormatting(
          section?.getAttribute('data-change-before'),
        ),
      ).toMatchObject({
        lnNumType: { countBy: 1, start: 1, restart: 'newPage' },
      });
      expect(section?.dataset.sectionLnNumCountBy).toBe('2');
      expect(section?.dataset.sectionLnNumStart).toBe('5');
      expect(section?.dataset.sectionLnNumDistance).toBe('720');
      expect(section?.dataset.sectionLnNumRestart).toBe('continuous');
    } finally {
      editor.destroy();
    }
  });

  test('imports pgNumType-only w:sectPrChange as a reviewable section-formatting change', async () => {
    const source = await sectionDocxWithPgNumTypeChange({
      prior: { fmt: 'decimal', start: 1 },
      current: { fmt: 'upperRoman', start: 5 },
    });
    const imported = await importOfficeFile(
      new File([source], 'section-pgnumtype-formatting.docx'),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const html = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    const section = html.body.querySelector('section[data-document-section]');
    expect(section?.getAttribute('data-change-kind')).toBe(
      'section-formatting',
    );
    expect(
      section?.getAttribute('data-section-property-revision-omml'),
    ).toBeNull();
    expect(
      parseDocumentSectionFormatting(
        section?.getAttribute('data-change-before'),
      ),
    ).toEqual({
      pgNumType: { fmt: 'decimal', start: 1 },
    });
    expect(section?.dataset.sectionPgNumFmt).toBe('upperRoman');
    expect(section?.dataset.sectionPgNumStart).toBe('5');
    expect(section?.dataset.sectionPageNumberStart).toBe('5');
  });

  test('pending pgNumType section-formatting change round-trips as native w:sectPrChange', async () => {
    const source = await sectionDocxWithPgNumTypeChange({
      prior: { fmt: 'decimal', start: 1 },
      current: { fmt: 'upperRoman', start: 5 },
    });
    const imported = await importOfficeFile(
      new File([source], 'section-pgnumtype-roundtrip.docx'),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const output = await JSZip.loadAsync(
      await (await createArtifactBlob(imported)).arrayBuffer(),
    );
    const exported = await xmlEntry(output, 'word/document.xml');
    const section = descendants(exported, 'sectPr').find(
      (element) => element.parentElement?.localName !== 'sectPrChange',
    );
    const change = directChild(section!, 'sectPrChange');
    expect(change).toBeTruthy();
    const priorPgNum = directChild(directChild(change!, 'sectPr'), 'pgNumType');
    expect(priorPgNum).toBeTruthy();
    expect(
      priorPgNum?.getAttributeNS(WORD_NAMESPACE, 'fmt') ??
        priorPgNum?.getAttribute('w:fmt') ??
        priorPgNum?.getAttribute('fmt'),
    ).toBe('decimal');
    expect(
      priorPgNum?.getAttributeNS(WORD_NAMESPACE, 'start') ??
        priorPgNum?.getAttribute('w:start') ??
        priorPgNum?.getAttribute('start'),
    ).toBe('1');
    const currentPgNum = directChild(section!, 'pgNumType');
    expect(
      currentPgNum?.getAttributeNS(WORD_NAMESPACE, 'fmt') ??
        currentPgNum?.getAttribute('w:fmt') ??
        currentPgNum?.getAttribute('fmt'),
    ).toBe('upperRoman');
    expect(
      currentPgNum?.getAttributeNS(WORD_NAMESPACE, 'start') ??
        currentPgNum?.getAttribute('w:start') ??
        currentPgNum?.getAttribute('start'),
    ).toBe('5');
  });

  test('reject restores prior pgNumType and drops the pending change', async () => {
    const source = await sectionDocxWithPgNumTypeChange({
      prior: { fmt: 'decimal', start: 1 },
      current: { fmt: 'upperRoman', start: 5 },
    });
    const imported = await importOfficeFile(
      new File([source], 'section-pgnumtype-reject.docx'),
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
      const section = html.body.querySelector('section[data-document-section]');
      expect(section?.getAttribute('data-change-kind')).toBeNull();
      expect(section?.dataset.sectionPgNumFmt).toBe('decimal');
      expect(section?.dataset.sectionPgNumStart).toBe('1');
      expect(section?.dataset.sectionPageNumberStart).toBe('1');
    } finally {
      editor.destroy();
    }
  });

  test('live section pgNumType edits become reviewable when track changes is on', () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions({
        isTracking: () => true,
      }),
      content: [
        '<section data-document-section="true" data-section-id="section-1"',
        ' data-section-orientation="portrait"',
        ' data-section-pg-num-fmt="decimal"',
        ' data-section-pg-num-start="1"',
        ' data-section-page-number-start="1">',
        '<p>Body</p>',
        '</section>',
      ].join(''),
    });
    try {
      const active = activeDocumentSection(editor);
      expect(active).not.toBeNull();
      if (!active) throw new Error('Expected an active document section.');
      expect(
        editor.commands.updateActiveDocumentSection({
          ...active.layout,
          pgNumType: { fmt: 'lowerRoman', start: 3 },
          pageNumberStart: 3,
        }),
      ).toBe(true);
      const changes = collectDocumentChanges(editor.state.doc);
      expect(changes).toHaveLength(1);
      expect(changes[0]?.kind).toBe('section-formatting');
      const html = new DOMParser().parseFromString(
        editor.getHTML(),
        'text/html',
      );
      const section = html.body.querySelector('section[data-document-section]');
      expect(section?.getAttribute('data-change-kind')).toBe(
        'section-formatting',
      );
      expect(
        parseDocumentSectionFormatting(
          section?.getAttribute('data-change-before'),
        ),
      ).toMatchObject({
        pgNumType: { fmt: 'decimal', start: 1 },
      });
      expect(section?.dataset.sectionPgNumFmt).toBe('lowerRoman');
      expect(section?.dataset.sectionPgNumStart).toBe('3');
      expect(section?.dataset.sectionPageNumberStart).toBe('3');
    } finally {
      editor.destroy();
    }
  });

  test('imports formProt-only w:sectPrChange as a reviewable section-formatting change', async () => {
    const source = await sectionDocxWithFormProtChange({
      prior: true,
      current: false,
    });
    const imported = await importOfficeFile(
      new File([source], 'section-formprot-formatting.docx'),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const html = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    const section = html.body.querySelector('section[data-document-section]');
    expect(section?.getAttribute('data-change-kind')).toBe(
      'section-formatting',
    );
    expect(
      section?.getAttribute('data-section-property-revision-omml'),
    ).toBeNull();
    expect(
      parseDocumentSectionFormatting(
        section?.getAttribute('data-change-before'),
      ),
    ).toEqual({ formProt: true });
    expect(section?.dataset.sectionFormProt).toBe('false');
  });

  test('pending formProt section-formatting change round-trips as native w:sectPrChange', async () => {
    const source = await sectionDocxWithFormProtChange({
      prior: true,
      current: false,
    });
    const imported = await importOfficeFile(
      new File([source], 'section-formprot-roundtrip.docx'),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const output = await JSZip.loadAsync(
      await (await createArtifactBlob(imported)).arrayBuffer(),
    );
    const exported = await xmlEntry(output, 'word/document.xml');
    const section = descendants(exported, 'sectPr').find(
      (element) => element.parentElement?.localName !== 'sectPrChange',
    );
    const change = directChild(section!, 'sectPrChange');
    expect(change).toBeTruthy();
    const priorFormProt = directChild(
      directChild(change!, 'sectPr'),
      'formProt',
    );
    expect(priorFormProt).toBeTruthy();
    expect(
      priorFormProt?.getAttributeNS(WORD_NAMESPACE, 'val') ??
        priorFormProt?.getAttribute('w:val') ??
        priorFormProt?.getAttribute('val'),
    ).toBeNull();
    const currentFormProt = directChild(section!, 'formProt');
    expect(currentFormProt).toBeTruthy();
    expect(
      currentFormProt?.getAttributeNS(WORD_NAMESPACE, 'val') ??
        currentFormProt?.getAttribute('w:val') ??
        currentFormProt?.getAttribute('val'),
    ).toBe('0');
  });

  test('reject restores prior formProt and drops the pending change', async () => {
    const source = await sectionDocxWithFormProtChange({
      prior: true,
      current: false,
    });
    const imported = await importOfficeFile(
      new File([source], 'section-formprot-reject.docx'),
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
      const section = html.body.querySelector('section[data-document-section]');
      expect(section?.getAttribute('data-change-kind')).toBeNull();
      expect(section?.dataset.sectionFormProt).toBe('true');
    } finally {
      editor.destroy();
    }
  });

  test('live section formProt edits become reviewable when track changes is on', () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions({
        isTracking: () => true,
      }),
      content: [
        '<section data-document-section="true" data-section-id="section-1"',
        ' data-section-orientation="portrait"',
        ' data-section-form-prot="false">',
        '<p>Body</p>',
        '</section>',
      ].join(''),
    });
    try {
      const active = activeDocumentSection(editor);
      expect(active).not.toBeNull();
      if (!active) throw new Error('Expected an active document section.');
      expect(
        editor.commands.updateActiveDocumentSection({
          ...active.layout,
          formProt: true,
        }),
      ).toBe(true);
      const changes = collectDocumentChanges(editor.state.doc);
      expect(changes).toHaveLength(1);
      expect(changes[0]?.kind).toBe('section-formatting');
      const html = new DOMParser().parseFromString(
        editor.getHTML(),
        'text/html',
      );
      const section = html.body.querySelector('section[data-document-section]');
      expect(section?.getAttribute('data-change-kind')).toBe(
        'section-formatting',
      );
      expect(
        parseDocumentSectionFormatting(
          section?.getAttribute('data-change-before'),
        ),
      ).toMatchObject({
        formProt: false,
      });
      expect(section?.dataset.sectionFormProt).toBe('true');
    } finally {
      editor.destroy();
    }
  });

  test('imports vAlign-only w:sectPrChange as a reviewable section-formatting change', async () => {
    const source = await sectionDocxWithVerticalAlignChange({
      prior: 'top',
      current: 'center',
    });
    const imported = await importOfficeFile(
      new File([source], 'section-valign-formatting.docx'),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const html = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    const section = html.body.querySelector('section[data-document-section]');
    expect(section?.getAttribute('data-change-kind')).toBe(
      'section-formatting',
    );
    expect(
      section?.getAttribute('data-section-property-revision-omml'),
    ).toBeNull();
    expect(
      parseDocumentSectionFormatting(
        section?.getAttribute('data-change-before'),
      ),
    ).toEqual({ verticalAlign: 'top' });
    expect(section?.dataset.sectionVerticalAlign).toBe('center');
  });

  test('pending vAlign section-formatting change round-trips as native w:sectPrChange', async () => {
    const source = await sectionDocxWithVerticalAlignChange({
      prior: 'top',
      current: 'center',
    });
    const imported = await importOfficeFile(
      new File([source], 'section-valign-roundtrip.docx'),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const output = await JSZip.loadAsync(
      await (await createArtifactBlob(imported)).arrayBuffer(),
    );
    const exported = await xmlEntry(output, 'word/document.xml');
    const section = descendants(exported, 'sectPr').find(
      (element) => element.parentElement?.localName !== 'sectPrChange',
    );
    const change = directChild(section!, 'sectPrChange');
    expect(change).toBeTruthy();
    const priorVAlign = directChild(directChild(change!, 'sectPr'), 'vAlign');
    expect(priorVAlign).toBeTruthy();
    expect(
      priorVAlign?.getAttributeNS(WORD_NAMESPACE, 'val') ??
        priorVAlign?.getAttribute('w:val') ??
        priorVAlign?.getAttribute('val'),
    ).toBe('top');
    const currentVAlign = directChild(section!, 'vAlign');
    expect(currentVAlign).toBeTruthy();
    expect(
      currentVAlign?.getAttributeNS(WORD_NAMESPACE, 'val') ??
        currentVAlign?.getAttribute('w:val') ??
        currentVAlign?.getAttribute('val'),
    ).toBe('center');
  });

  test('reject restores prior vAlign and drops the pending change', async () => {
    const source = await sectionDocxWithVerticalAlignChange({
      prior: 'top',
      current: 'center',
    });
    const imported = await importOfficeFile(
      new File([source], 'section-valign-reject.docx'),
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
      const section = html.body.querySelector('section[data-document-section]');
      expect(section?.getAttribute('data-change-kind')).toBeNull();
      expect(section?.dataset.sectionVerticalAlign).toBe('top');
    } finally {
      editor.destroy();
    }
  });

  test('live section vAlign edits become reviewable when track changes is on', () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions({
        isTracking: () => true,
      }),
      content: [
        '<section data-document-section="true" data-section-id="section-1"',
        ' data-section-orientation="portrait"',
        ' data-section-vertical-align="top">',
        '<p>Body</p>',
        '</section>',
      ].join(''),
    });
    try {
      const active = activeDocumentSection(editor);
      expect(active).not.toBeNull();
      if (!active) throw new Error('Expected an active document section.');
      expect(
        editor.commands.updateActiveDocumentSection({
          ...active.layout,
          verticalAlign: 'bottom',
        }),
      ).toBe(true);
      const changes = collectDocumentChanges(editor.state.doc);
      expect(changes).toHaveLength(1);
      expect(changes[0]?.kind).toBe('section-formatting');
      const html = new DOMParser().parseFromString(
        editor.getHTML(),
        'text/html',
      );
      const section = html.body.querySelector('section[data-document-section]');
      expect(section?.getAttribute('data-change-kind')).toBe(
        'section-formatting',
      );
      expect(
        parseDocumentSectionFormatting(
          section?.getAttribute('data-change-before'),
        ),
      ).toMatchObject({
        verticalAlign: 'top',
      });
      expect(section?.dataset.sectionVerticalAlign).toBe('bottom');
    } finally {
      editor.destroy();
    }
  });

    test('imports rtlGutter-only w:sectPrChange as a reviewable section-formatting change', async () => {
    const source = await sectionDocxWithRtlGutterChange({
      prior: true,
      current: false,
    });
    const imported = await importOfficeFile(
      new File([source], 'section-rtlgutter-formatting.docx'),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const html = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    const section = html.body.querySelector('section[data-document-section]');
    expect(section?.getAttribute('data-change-kind')).toBe(
      'section-formatting',
    );
    expect(
      section?.getAttribute('data-section-property-revision-omml'),
    ).toBeNull();
    expect(
      parseDocumentSectionFormatting(
        section?.getAttribute('data-change-before'),
      ),
    ).toEqual({ rtlGutter: true });
  });

  test('pending rtlGutter section-formatting change round-trips as native w:sectPrChange', async () => {
    const source = await sectionDocxWithRtlGutterChange({
      prior: true,
      current: false,
    });
    const imported = await importOfficeFile(
      new File([source], 'section-rtlgutter-roundtrip.docx'),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const output = await JSZip.loadAsync(
      await (await createArtifactBlob(imported)).arrayBuffer(),
    );
    const exported = await xmlEntry(output, 'word/document.xml');
    const section = descendants(exported, 'sectPr').find(
      (element) => element.parentElement?.localName !== 'sectPrChange',
    );
    const change = directChild(section!, 'sectPrChange');
    expect(change).toBeTruthy();
    expect(
      directChild(directChild(change!, 'sectPr'), 'rtlGutter'),
    ).toBeTruthy();
  });

  test('live section rtlGutter edits become reviewable when track changes is on', () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions({
        isTracking: () => true,
      }),
      content: [
        '<section data-document-section="true" data-section-id="section-1"',
        ' data-section-orientation="portrait"',
        ' data-section-page-margins=\'{"top":1440,"right":1440,"bottom":1440,"left":1440,"header":708,"footer":708,"gutter":0}\'>',
        '<p>Body</p>',
        '</section>',
      ].join(''),
    });
    try {
      const active = activeDocumentSection(editor);
      expect(active).not.toBeNull();
      if (!active) throw new Error('Expected an active document section.');
      expect(
        editor.commands.updateActiveDocumentSection({
          ...active.layout,
          pageMargins: {
            top: 1440,
            right: 1440,
            bottom: 1440,
            left: 1440,
            header: 708,
            footer: 708,
            gutter: 0,
            gutterOnRight: true,
          },
        }),
      ).toBe(true);
      const changes = collectDocumentChanges(editor.state.doc);
      expect(changes).toHaveLength(1);
      expect(changes[0]?.kind).toBe('section-formatting');
      const html = new DOMParser().parseFromString(
        editor.getHTML(),
        'text/html',
      );
      const section = html.body.querySelector('section[data-document-section]');
      expect(section?.getAttribute('data-change-kind')).toBe(
        'section-formatting',
      );
      expect(
        parseDocumentSectionFormatting(
          section?.getAttribute('data-change-before'),
        ),
      ).toMatchObject({ rtlGutter: false });
    } finally {
      editor.destroy();
    }
  });

  test('imports titlePg-only w:sectPrChange as a reviewable section-formatting change', async () => {
    const source = await sectionDocxWithTitlePageChange({
      prior: true,
      current: false,
    });
    const imported = await importOfficeFile(
      new File([source], 'section-titlepg-formatting.docx'),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const html = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    const section = html.body.querySelector('section[data-document-section]');
    expect(section?.getAttribute('data-change-kind')).toBe(
      'section-formatting',
    );
    expect(
      section?.getAttribute('data-section-property-revision-omml'),
    ).toBeNull();
    expect(
      parseDocumentSectionFormatting(
        section?.getAttribute('data-change-before'),
      ),
    ).toEqual({ differentFirstPage: true });
  });

  test('pending titlePg section-formatting change round-trips as native w:sectPrChange', async () => {
    const source = await sectionDocxWithTitlePageChange({
      prior: true,
      current: false,
    });
    const imported = await importOfficeFile(
      new File([source], 'section-titlepg-roundtrip.docx'),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const output = await JSZip.loadAsync(
      await (await createArtifactBlob(imported)).arrayBuffer(),
    );
    const exported = await xmlEntry(output, 'word/document.xml');
    const section = descendants(exported, 'sectPr').find(
      (element) => element.parentElement?.localName !== 'sectPrChange',
    );
    const change = directChild(section!, 'sectPrChange');
    expect(change).toBeTruthy();
    expect(directChild(directChild(change!, 'sectPr'), 'titlePg')).toBeTruthy();
  });

  test('live section differentFirstPage edits become reviewable when track changes is on', () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions({
        isTracking: () => true,
      }),
      content: [
        '<section data-document-section="true" data-section-id="section-1"',
        ' data-section-orientation="portrait"',
        ' data-section-column-count="1">',
        '<p>Body</p>',
        '</section>',
      ].join(''),
    });
    try {
      const active = activeDocumentSection(editor);
      expect(active).not.toBeNull();
      if (!active) throw new Error('Expected an active document section.');
      expect(
        editor.commands.updateActiveDocumentSection({
          ...active.layout,
          pageChrome: {
            ...active.layout.pageChrome!,
            differentFirstPage: true,
          },
        }),
      ).toBe(true);
      const changes = collectDocumentChanges(editor.state.doc);
      expect(changes).toHaveLength(1);
      expect(changes[0]?.kind).toBe('section-formatting');
      const html = new DOMParser().parseFromString(
        editor.getHTML(),
        'text/html',
      );
      const section = html.body.querySelector('section[data-document-section]');
      expect(section?.getAttribute('data-change-kind')).toBe(
        'section-formatting',
      );
      expect(
        parseDocumentSectionFormatting(
          section?.getAttribute('data-change-before'),
        ),
      ).toMatchObject({ differentFirstPage: false });
      const chrome = JSON.parse(
        section?.getAttribute('data-section-page-chrome') ?? '{}',
      );
      expect(chrome.differentFirstPage).toBe(true);
    } finally {
      editor.destroy();
    }
  });

  test('imports unequal-width cols-only w:sectPrChange as a reviewable section-formatting change', async () => {
    const source = await sectionDocxWithUnequalColumnsChange({
      prior: {
        count: 2,
        spacing: 12.7,
        separator: false,
        custom: [
          { widthPercent: 40, spacing: 12.7 },
          { widthPercent: 60, spacing: 0 },
        ],
      },
      current: {
        count: 2,
        spacing: 12,
        separator: true,
        custom: [
          { widthPercent: 50, spacing: 12 },
          { widthPercent: 50, spacing: 0 },
        ],
      },
    });
    const imported = await importOfficeFile(
      new File([source], 'section-unequal-columns-formatting.docx'),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const html = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    const section = html.body.querySelector('section[data-document-section]');
    expect(section?.getAttribute('data-change-kind')).toBe(
      'section-formatting',
    );
    expect(section?.dataset.sectionPropertyRevisionOmml).toBeFalsy();
    expect(
      parseDocumentSectionFormatting(
        section?.getAttribute('data-change-before'),
      ),
    ).toEqual({
      columns: {
        count: 2,
        spacing: 12.7,
        separator: false,
        custom: [
          { widthPercent: 40, spacing: 12.7 },
          { widthPercent: 60, spacing: 0 },
        ],
      },
    });

    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: imported.content.html,
    });
    try {
      const change = collectDocumentChanges(editor.state.doc)[0];
      expect(change?.kind).toBe('section-formatting');
      expect(editor.commands.rejectDocumentChange(change?.id ?? '')).toBe(true);
      const rejected = new DOMParser().parseFromString(
        editor.getHTML(),
        'text/html',
      );
      const rejectedSection = rejected.body.querySelector(
        'section[data-document-section]',
      );
      expect(rejectedSection?.getAttribute('data-change-kind')).toBeNull();
      expect(rejectedSection?.dataset.sectionColumnCount).toBe('2');
      expect(
        JSON.parse(rejectedSection?.dataset.sectionColumnLayout ?? '{}').custom,
      ).toEqual([
        { widthPercent: 40, spacing: 12.7 },
        { widthPercent: 60, spacing: 0 },
      ]);
    } finally {
      editor.destroy();
    }
  });

  test('pending unequal-width cols section-formatting change round-trips as native w:sectPrChange', async () => {
    const source = await sectionDocxWithUnequalColumnsChange({
      prior: {
        count: 2,
        spacing: 12.7,
        separator: false,
        custom: [
          { widthPercent: 40, spacing: 12.7 },
          { widthPercent: 60, spacing: 0 },
        ],
      },
      current: {
        count: 2,
        spacing: 12,
        separator: false,
        custom: [
          { widthPercent: 50, spacing: 12 },
          { widthPercent: 50, spacing: 0 },
        ],
      },
    });
    const imported = await importOfficeFile(
      new File([source], 'section-unequal-columns-round-trip.docx'),
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
      descendants(exported, 'sectPr').find(
        (element) => element.parentElement?.localName !== 'sectPrChange',
      )!,
      'sectPrChange',
    );
    expect(change).toBeTruthy();
    const priorCols = directChild(directChild(change!, 'sectPr'), 'cols');
    expect(priorCols).toBeTruthy();
    expect(
      priorCols?.getAttributeNS(WORD_NAMESPACE, 'equalWidth') ??
        priorCols?.getAttribute('w:equalWidth'),
    ).toBe('0');
    const priorColumnWidths = Array.from(priorCols?.children ?? []).map(
      (column) =>
        column.getAttributeNS(WORD_NAMESPACE, 'w') ??
        column.getAttribute('w:w') ??
        column.getAttribute('w'),
    );
    expect(priorColumnWidths).toEqual(['40', '60']);
  });

  test('imports equal-width cols-only w:sectPrChange as a reviewable section-formatting change', async () => {
    const source = await sectionDocxWithColumnsChange({
      prior: { count: 1, spacing: 12, separator: false },
      current: { count: 2, spacing: 12.7, separator: false },
    });
    const imported = await importOfficeFile(
      new File([source], 'section-columns-formatting.docx'),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const html = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    const section = html.body.querySelector('section[data-document-section]');
    expect(section?.getAttribute('data-change-kind')).toBe(
      'section-formatting',
    );
    expect(section?.dataset.sectionPropertyRevisionOmml).toBeFalsy();
    expect(
      parseDocumentSectionFormatting(
        section?.getAttribute('data-change-before'),
      ),
    ).toEqual({
      columns: { count: 1, spacing: 12, separator: false },
    });

    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: imported.content.html,
    });
    try {
      const change = collectDocumentChanges(editor.state.doc)[0];
      expect(change?.kind).toBe('section-formatting');
      expect(editor.commands.rejectDocumentChange(change?.id ?? '')).toBe(true);
      const rejected = new DOMParser().parseFromString(
        editor.getHTML(),
        'text/html',
      );
      const rejectedSection = rejected.body.querySelector(
        'section[data-document-section]',
      );
      expect(rejectedSection?.getAttribute('data-change-kind')).toBeNull();
      expect(rejectedSection?.dataset.sectionColumnCount).toBe('1');
    } finally {
      editor.destroy();
    }
  });

  test('live section column edits become reviewable when track changes is on', () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions({
        isTracking: () => true,
      }),
      content: [
        '<section data-document-section="true" data-section-id="section-1"',
        ' data-section-orientation="portrait"',
        ' data-section-column-count="1">',
        '<p>Body</p>',
        '</section>',
      ].join(''),
    });
    try {
      const active = activeDocumentSection(editor);
      expect(active).not.toBeNull();
      if (!active) throw new Error('Expected an active document section.');
      expect(
        editor.commands.updateActiveDocumentSection({
          ...active.layout,
          columns: { count: 2, spacing: 12, separator: true },
        }),
      ).toBe(true);
      const changes = collectDocumentChanges(editor.state.doc);
      expect(changes).toHaveLength(1);
      expect(changes[0]?.kind).toBe('section-formatting');
      const html = new DOMParser().parseFromString(
        editor.getHTML(),
        'text/html',
      );
      const section = html.body.querySelector('section[data-document-section]');
      expect(section?.getAttribute('data-change-kind')).toBe(
        'section-formatting',
      );
      expect(
        parseDocumentSectionFormatting(
          section?.getAttribute('data-change-before'),
        )?.columns,
      ).toEqual({ count: 1, spacing: 12, separator: false });
      expect(section?.dataset.sectionColumnCount).toBe('2');
      expect(section?.dataset.sectionColumnSeparator).toBe('true');
    } finally {
      editor.destroy();
    }
  });
});

async function sectionDocxWithDocGridChange(options: {
  prior: { type: 'default' | 'lines' | 'linesAndChars' | 'snapToChars'; linePitch: number };
  current: { type: 'default' | 'lines' | 'linesAndChars' | 'snapToChars'; linePitch: number };
}): Promise<ArrayBuffer> {
  const artifact = createArtifact('blank-document');
  if (artifact.content.type !== 'document') {
    throw new Error('Expected a document artifact.');
  }
  const seed = await createArtifactBlob(artifact);
  const archive = await JSZip.loadAsync(await seed.arrayBuffer());
  const document = await xmlEntry(archive, 'word/document.xml');
  const section = descendants(document, 'sectPr').find(
    (element) => element.parentElement?.localName !== 'sectPrChange',
  );
  if (!section) throw new Error('Expected body sectPr.');
  for (const existing of Array.from(section.children).filter(
    (child) =>
      child.localName === 'docGrid' || child.localName === 'sectPrChange',
  )) {
    existing.remove();
  }
  const current = document.createElementNS(WORD_NAMESPACE, 'w:docGrid');
  current.setAttributeNS(WORD_NAMESPACE, 'w:type', options.current.type);
  current.setAttributeNS(
    WORD_NAMESPACE,
    'w:linePitch',
    String(Math.max(1, Math.round(options.current.linePitch * 20))),
  );
  section.append(current);
  const change = document.createElementNS(WORD_NAMESPACE, 'w:sectPrChange');
  change.setAttributeNS(WORD_NAMESPACE, 'w:id', '33');
  change.setAttributeNS(WORD_NAMESPACE, 'w:author', 'Reviewer');
  change.setAttributeNS(WORD_NAMESPACE, 'w:date', '2026-09-10T00:00:00Z');
  const prior = document.createElementNS(WORD_NAMESPACE, 'w:sectPr');
  const priorGrid = document.createElementNS(WORD_NAMESPACE, 'w:docGrid');
  priorGrid.setAttributeNS(WORD_NAMESPACE, 'w:type', options.prior.type);
  priorGrid.setAttributeNS(
    WORD_NAMESPACE,
    'w:linePitch',
    String(Math.max(1, Math.round(options.prior.linePitch * 20))),
  );
  prior.append(priorGrid);
  change.append(prior);
  section.append(change);
  archive.file(
    'word/document.xml',
    new XMLSerializer().serializeToString(document),
  );
  return archive.generateAsync({ type: 'arraybuffer' });
}

async function sectionDocxWithPgNumTypeChange(options: {
  prior: { fmt?: string; start?: number };
  current: { fmt?: string; start?: number };
}): Promise<ArrayBuffer> {
  const artifact = createArtifact('blank-document');
  if (artifact.content.type !== 'document') {
    throw new Error('Expected a document artifact.');
  }
  const seed = await createArtifactBlob(artifact);
  const archive = await JSZip.loadAsync(await seed.arrayBuffer());
  const document = await xmlEntry(archive, 'word/document.xml');
  const section = descendants(document, 'sectPr').find(
    (element) => element.parentElement?.localName !== 'sectPrChange',
  );
  if (!section) throw new Error('Expected body sectPr.');
  for (const existing of Array.from(section.children).filter(
    (child) =>
      child.localName === 'pgNumType' || child.localName === 'sectPrChange',
  )) {
    existing.remove();
  }
  const current = document.createElementNS(WORD_NAMESPACE, 'w:pgNumType');
  applyPgNumTypeAttributes(current, options.current);
  section.append(current);
  const change = document.createElementNS(WORD_NAMESPACE, 'w:sectPrChange');
  change.setAttributeNS(WORD_NAMESPACE, 'w:id', '35');
  change.setAttributeNS(WORD_NAMESPACE, 'w:author', 'Reviewer');
  change.setAttributeNS(WORD_NAMESPACE, 'w:date', '2026-09-10T00:00:00Z');
  const prior = document.createElementNS(WORD_NAMESPACE, 'w:sectPr');
  const priorPg = document.createElementNS(WORD_NAMESPACE, 'w:pgNumType');
  applyPgNumTypeAttributes(priorPg, options.prior);
  prior.append(priorPg);
  change.append(prior);
  section.append(change);
  archive.file(
    'word/document.xml',
    new XMLSerializer().serializeToString(document),
  );
  return archive.generateAsync({ type: 'arraybuffer' });
}

async function sectionDocxWithFormProtChange(options: {
  prior: boolean;
  current: boolean;
}): Promise<ArrayBuffer> {
  const artifact = createArtifact('blank-document');
  if (artifact.content.type !== 'document') {
    throw new Error('Expected a document artifact.');
  }
  const seed = await createArtifactBlob(artifact);
  const archive = await JSZip.loadAsync(await seed.arrayBuffer());
  const document = await xmlEntry(archive, 'word/document.xml');
  const section = descendants(document, 'sectPr').find(
    (element) => element.parentElement?.localName !== 'sectPrChange',
  );
  if (!section) throw new Error('Expected body sectPr.');
  for (const existing of Array.from(section.children).filter(
    (child) =>
      child.localName === 'formProt' || child.localName === 'sectPrChange',
  )) {
    existing.remove();
  }
  const current = document.createElementNS(WORD_NAMESPACE, 'w:formProt');
  if (!options.current) {
    current.setAttributeNS(WORD_NAMESPACE, 'w:val', '0');
  }
  section.append(current);
  const change = document.createElementNS(WORD_NAMESPACE, 'w:sectPrChange');
  change.setAttributeNS(WORD_NAMESPACE, 'w:id', '36');
  change.setAttributeNS(WORD_NAMESPACE, 'w:author', 'Reviewer');
  change.setAttributeNS(WORD_NAMESPACE, 'w:date', '2026-09-10T00:00:00Z');
  const prior = document.createElementNS(WORD_NAMESPACE, 'w:sectPr');
  const priorFormProt = document.createElementNS(WORD_NAMESPACE, 'w:formProt');
  if (!options.prior) {
    priorFormProt.setAttributeNS(WORD_NAMESPACE, 'w:val', '0');
  }
  prior.append(priorFormProt);
  change.append(prior);
  section.append(change);
  archive.file(
    'word/document.xml',
    new XMLSerializer().serializeToString(document),
  );
  return archive.generateAsync({ type: 'arraybuffer' });
}

async function sectionDocxWithVerticalAlignChange(options: {
  prior: 'top' | 'center' | 'both' | 'bottom';
  current: 'top' | 'center' | 'both' | 'bottom';
}): Promise<ArrayBuffer> {
  const artifact = createArtifact('blank-document');
  if (artifact.content.type !== 'document') {
    throw new Error('Expected a document artifact.');
  }
  const seed = await createArtifactBlob(artifact);
  const archive = await JSZip.loadAsync(await seed.arrayBuffer());
  const document = await xmlEntry(archive, 'word/document.xml');
  const section = descendants(document, 'sectPr').find(
    (element) => element.parentElement?.localName !== 'sectPrChange',
  );
  if (!section) throw new Error('Expected body sectPr.');
  for (const existing of Array.from(section.children).filter(
    (child) =>
      child.localName === 'vAlign' || child.localName === 'sectPrChange',
  )) {
    existing.remove();
  }
  const current = document.createElementNS(WORD_NAMESPACE, 'w:vAlign');
  current.setAttributeNS(WORD_NAMESPACE, 'w:val', options.current);
  section.append(current);
  const change = document.createElementNS(WORD_NAMESPACE, 'w:sectPrChange');
  change.setAttributeNS(WORD_NAMESPACE, 'w:id', '37');
  change.setAttributeNS(WORD_NAMESPACE, 'w:author', 'Reviewer');
  change.setAttributeNS(WORD_NAMESPACE, 'w:date', '2026-09-10T00:00:00Z');
  const prior = document.createElementNS(WORD_NAMESPACE, 'w:sectPr');
  const priorVAlign = document.createElementNS(WORD_NAMESPACE, 'w:vAlign');
  priorVAlign.setAttributeNS(WORD_NAMESPACE, 'w:val', options.prior);
  prior.append(priorVAlign);
  change.append(prior);
  section.append(change);
  archive.file(
    'word/document.xml',
    new XMLSerializer().serializeToString(document),
  );
  return archive.generateAsync({ type: 'arraybuffer' });
}

function applyPgNumTypeAttributes(
  element: Element,
  value: { fmt?: string; start?: number },
): void {
  if (value.fmt !== undefined) {
    element.setAttributeNS(WORD_NAMESPACE, 'w:fmt', value.fmt);
  }
  if (value.start !== undefined) {
    element.setAttributeNS(WORD_NAMESPACE, 'w:start', String(value.start));
  }
}

async function sectionDocxWithLnNumTypeChange(options: {
  prior: {
    countBy?: number;
    start?: number;
    distance?: number;
    restart?: 'newPage' | 'newSection' | 'continuous';
  };
  current: {
    countBy?: number;
    start?: number;
    distance?: number;
    restart?: 'newPage' | 'newSection' | 'continuous';
  };
}): Promise<ArrayBuffer> {
  const artifact = createArtifact('blank-document');
  if (artifact.content.type !== 'document') {
    throw new Error('Expected a document artifact.');
  }
  const seed = await createArtifactBlob(artifact);
  const archive = await JSZip.loadAsync(await seed.arrayBuffer());
  const document = await xmlEntry(archive, 'word/document.xml');
  const section = descendants(document, 'sectPr').find(
    (element) => element.parentElement?.localName !== 'sectPrChange',
  );
  if (!section) throw new Error('Expected body sectPr.');
  for (const existing of Array.from(section.children).filter(
    (child) =>
      child.localName === 'lnNumType' || child.localName === 'sectPrChange',
  )) {
    existing.remove();
  }
  const current = document.createElementNS(WORD_NAMESPACE, 'w:lnNumType');
  applyLnNumTypeAttributes(current, options.current);
  section.append(current);
  const change = document.createElementNS(WORD_NAMESPACE, 'w:sectPrChange');
  change.setAttributeNS(WORD_NAMESPACE, 'w:id', '34');
  change.setAttributeNS(WORD_NAMESPACE, 'w:author', 'Reviewer');
  change.setAttributeNS(WORD_NAMESPACE, 'w:date', '2026-09-10T00:00:00Z');
  const prior = document.createElementNS(WORD_NAMESPACE, 'w:sectPr');
  const priorLn = document.createElementNS(WORD_NAMESPACE, 'w:lnNumType');
  applyLnNumTypeAttributes(priorLn, options.prior);
  prior.append(priorLn);
  change.append(prior);
  section.append(change);
  archive.file(
    'word/document.xml',
    new XMLSerializer().serializeToString(document),
  );
  return archive.generateAsync({ type: 'arraybuffer' });
}

function applyLnNumTypeAttributes(
  element: Element,
  value: {
    countBy?: number;
    start?: number;
    distance?: number;
    restart?: 'newPage' | 'newSection' | 'continuous';
  },
): void {
  if (value.countBy !== undefined) {
    element.setAttributeNS(WORD_NAMESPACE, 'w:countBy', String(value.countBy));
  }
  if (value.start !== undefined) {
    element.setAttributeNS(WORD_NAMESPACE, 'w:start', String(value.start));
  }
  if (value.distance !== undefined) {
    element.setAttributeNS(
      WORD_NAMESPACE,
      'w:distance',
      String(value.distance),
    );
  }
  if (value.restart !== undefined) {
    element.setAttributeNS(WORD_NAMESPACE, 'w:restart', value.restart);
  }
}

async function sectionDocxWithRtlGutterChange(options: {
  prior: boolean;
  current: boolean;
}): Promise<ArrayBuffer> {
  const artifact = createArtifact('blank-document');
  if (artifact.content.type !== 'document') {
    throw new Error('Expected a document artifact.');
  }
  const seed = await createArtifactBlob(artifact);
  const archive = await JSZip.loadAsync(await seed.arrayBuffer());
  const document = await xmlEntry(archive, 'word/document.xml');
  const section = descendants(document, 'sectPr').find(
    (element) => element.parentElement?.localName !== 'sectPrChange',
  );
  if (!section) throw new Error('Expected body sectPr.');
  for (const existing of Array.from(section.children).filter(
    (child) =>
      child.localName === 'rtlGutter' || child.localName === 'sectPrChange',
  )) {
    existing.remove();
  }
  if (options.current) {
    section.append(document.createElementNS(WORD_NAMESPACE, 'w:rtlGutter'));
  }
  const change = document.createElementNS(WORD_NAMESPACE, 'w:sectPrChange');
  change.setAttributeNS(WORD_NAMESPACE, 'w:id', '32');
  change.setAttributeNS(WORD_NAMESPACE, 'w:author', 'Reviewer');
  change.setAttributeNS(WORD_NAMESPACE, 'w:date', '2026-09-09T00:00:00Z');
  const prior = document.createElementNS(WORD_NAMESPACE, 'w:sectPr');
  if (options.prior) {
    prior.append(document.createElementNS(WORD_NAMESPACE, 'w:rtlGutter'));
  } else {
    const off = document.createElementNS(WORD_NAMESPACE, 'w:rtlGutter');
    off.setAttributeNS(WORD_NAMESPACE, 'w:val', '0');
    prior.append(off);
  }
  change.append(prior);
  section.append(change);
  archive.file(
    'word/document.xml',
    new XMLSerializer().serializeToString(document),
  );
  return archive.generateAsync({ type: 'arraybuffer' });
}

async function sectionDocxWithTitlePageChange(options: {
  prior: boolean;
  current: boolean;
}): Promise<ArrayBuffer> {
  const artifact = createArtifact('blank-document');
  if (artifact.content.type !== 'document') {
    throw new Error('Expected a document artifact.');
  }
  const seed = await createArtifactBlob(artifact);
  const archive = await JSZip.loadAsync(await seed.arrayBuffer());
  const document = await xmlEntry(archive, 'word/document.xml');
  const section = descendants(document, 'sectPr').find(
    (element) => element.parentElement?.localName !== 'sectPrChange',
  );
  if (!section) throw new Error('Expected body sectPr.');
  for (const existing of Array.from(section.children).filter(
    (child) =>
      child.localName === 'titlePg' || child.localName === 'sectPrChange',
  )) {
    existing.remove();
  }
  if (options.current) {
    section.append(document.createElementNS(WORD_NAMESPACE, 'w:titlePg'));
  }
  const change = document.createElementNS(WORD_NAMESPACE, 'w:sectPrChange');
  change.setAttributeNS(WORD_NAMESPACE, 'w:id', '31');
  change.setAttributeNS(WORD_NAMESPACE, 'w:author', 'Reviewer');
  change.setAttributeNS(WORD_NAMESPACE, 'w:date', '2026-09-09T00:00:00Z');
  const prior = document.createElementNS(WORD_NAMESPACE, 'w:sectPr');
  if (options.prior) {
    prior.append(document.createElementNS(WORD_NAMESPACE, 'w:titlePg'));
  } else {
    const off = document.createElementNS(WORD_NAMESPACE, 'w:titlePg');
    off.setAttributeNS(WORD_NAMESPACE, 'w:val', '0');
    prior.append(off);
  }
  change.append(prior);
  section.append(change);
  archive.file(
    'word/document.xml',
    new XMLSerializer().serializeToString(document),
  );
  return archive.generateAsync({ type: 'arraybuffer' });
}

async function sectionDocxWithOrientationChange(options: {
  prior: 'portrait' | 'landscape';
  current?: 'portrait' | 'landscape';
}): Promise<ArrayBuffer> {
  const artifact = createArtifact('blank-document');
  if (artifact.content.type !== 'document') {
    throw new Error('Expected a document artifact.');
  }
  const seed = await createArtifactBlob(artifact);
  const archive = await JSZip.loadAsync(await seed.arrayBuffer());
  const document = await xmlEntry(archive, 'word/document.xml');
  const section = descendants(document, 'sectPr').find(
    (element) => element.parentElement?.localName !== 'sectPrChange',
  );
  if (!section) throw new Error('Expected body sectPr.');
  const currentOrientation = options.current ?? 'portrait';
  let pageSize = directChild(section, 'pgSz');
  if (!pageSize) {
    pageSize = document.createElementNS(WORD_NAMESPACE, 'w:pgSz');
    section.insertBefore(pageSize, section.firstChild);
  }
  pageSize.setAttributeNS(WORD_NAMESPACE, 'w:orient', currentOrientation);
  for (const existing of Array.from(section.children).filter(
    (child) => child.localName === 'sectPrChange',
  )) {
    existing.remove();
  }
  const change = document.createElementNS(WORD_NAMESPACE, 'w:sectPrChange');
  change.setAttributeNS(WORD_NAMESPACE, 'w:id', '41');
  change.setAttributeNS(WORD_NAMESPACE, 'w:author', 'Reviewer');
  change.setAttributeNS(WORD_NAMESPACE, 'w:date', '2026-09-08T00:00:00Z');
  const prior = document.createElementNS(WORD_NAMESPACE, 'w:sectPr');
  const priorSize = document.createElementNS(WORD_NAMESPACE, 'w:pgSz');
  priorSize.setAttributeNS(WORD_NAMESPACE, 'w:orient', options.prior);
  prior.append(priorSize);
  change.append(prior);
  section.append(change);
  archive.file(
    'word/document.xml',
    new XMLSerializer().serializeToString(document),
  );
  return archive.generateAsync({ type: 'arraybuffer' });
}

async function sectionDocxWithMarginChange(options: {
  prior: {
    top: number;
    right: number;
    bottom: number;
    left: number;
    header: number;
    footer: number;
    gutter: number;
  };
  current: {
    top: number;
    right: number;
    bottom: number;
    left: number;
    header: number;
    footer: number;
    gutter: number;
  };
}): Promise<ArrayBuffer> {
  const artifact = createArtifact('blank-document');
  if (artifact.content.type !== 'document') {
    throw new Error('Expected a document artifact.');
  }
  const seed = await createArtifactBlob(artifact);
  const archive = await JSZip.loadAsync(await seed.arrayBuffer());
  const document = await xmlEntry(archive, 'word/document.xml');
  const section = descendants(document, 'sectPr').find(
    (element) => element.parentElement?.localName !== 'sectPrChange',
  );
  if (!section) throw new Error('Expected body sectPr.');
  for (const existing of Array.from(section.children).filter(
    (child) =>
      child.localName === 'pgMar' || child.localName === 'sectPrChange',
  )) {
    existing.remove();
  }
  const current = document.createElementNS(WORD_NAMESPACE, 'w:pgMar');
  for (const [key, value] of Object.entries(options.current)) {
    current.setAttributeNS(WORD_NAMESPACE, `w:${key}`, String(value));
  }
  section.append(current);
  const change = document.createElementNS(WORD_NAMESPACE, 'w:sectPrChange');
  change.setAttributeNS(WORD_NAMESPACE, 'w:id', '42');
  change.setAttributeNS(WORD_NAMESPACE, 'w:author', 'Reviewer');
  change.setAttributeNS(WORD_NAMESPACE, 'w:date', '2026-09-08T00:00:00Z');
  const prior = document.createElementNS(WORD_NAMESPACE, 'w:sectPr');
  const priorMargins = document.createElementNS(WORD_NAMESPACE, 'w:pgMar');
  for (const [key, value] of Object.entries(options.prior)) {
    priorMargins.setAttributeNS(WORD_NAMESPACE, `w:${key}`, String(value));
  }
  prior.append(priorMargins);
  change.append(prior);
  section.append(change);
  archive.file(
    'word/document.xml',
    new XMLSerializer().serializeToString(document),
  );
  return archive.generateAsync({ type: 'arraybuffer' });
}

async function sectionDocxWithPageGeometryChange(options: {
  prior: {
    width: number;
    height: number;
    orientation: 'portrait' | 'landscape';
  };
  current: {
    width: number;
    height: number;
    orientation: 'portrait' | 'landscape';
  };
}): Promise<ArrayBuffer> {
  const artifact = createArtifact('blank-document');
  if (artifact.content.type !== 'document') {
    throw new Error('Expected a document artifact.');
  }
  const seed = await createArtifactBlob(artifact);
  const archive = await JSZip.loadAsync(await seed.arrayBuffer());
  const document = await xmlEntry(archive, 'word/document.xml');
  const section = descendants(document, 'sectPr').find(
    (element) => element.parentElement?.localName !== 'sectPrChange',
  );
  if (!section) throw new Error('Expected body sectPr.');
  for (const existing of Array.from(section.children).filter(
    (child) => child.localName === 'pgSz' || child.localName === 'sectPrChange',
  )) {
    existing.remove();
  }
  const current = document.createElementNS(WORD_NAMESPACE, 'w:pgSz');
  current.setAttributeNS(WORD_NAMESPACE, 'w:w', String(options.current.width));
  current.setAttributeNS(WORD_NAMESPACE, 'w:h', String(options.current.height));
  current.setAttributeNS(
    WORD_NAMESPACE,
    'w:orient',
    options.current.orientation,
  );
  section.insertBefore(current, section.firstChild);
  const change = document.createElementNS(WORD_NAMESPACE, 'w:sectPrChange');
  change.setAttributeNS(WORD_NAMESPACE, 'w:id', '43');
  change.setAttributeNS(WORD_NAMESPACE, 'w:author', 'Reviewer');
  change.setAttributeNS(WORD_NAMESPACE, 'w:date', '2026-09-08T00:00:00Z');
  const prior = document.createElementNS(WORD_NAMESPACE, 'w:sectPr');
  const priorSize = document.createElementNS(WORD_NAMESPACE, 'w:pgSz');
  priorSize.setAttributeNS(WORD_NAMESPACE, 'w:w', String(options.prior.width));
  priorSize.setAttributeNS(WORD_NAMESPACE, 'w:h', String(options.prior.height));
  priorSize.setAttributeNS(
    WORD_NAMESPACE,
    'w:orient',
    options.prior.orientation,
  );
  prior.append(priorSize);
  change.append(prior);
  section.append(change);
  archive.file(
    'word/document.xml',
    new XMLSerializer().serializeToString(document),
  );
  return archive.generateAsync({ type: 'arraybuffer' });
}

async function sectionDocxWithPaperSourceChange(options: {
  prior: { first: number; other: number };
  current: { first: number; other: number };
}): Promise<ArrayBuffer> {
  const artifact = createArtifact('blank-document');
  if (artifact.content.type !== 'document') {
    throw new Error('Expected a document artifact.');
  }
  const seed = await createArtifactBlob(artifact);
  const archive = await JSZip.loadAsync(await seed.arrayBuffer());
  const document = await xmlEntry(archive, 'word/document.xml');
  const section = descendants(document, 'sectPr').find(
    (element) => element.parentElement?.localName !== 'sectPrChange',
  );
  if (!section) throw new Error('Expected body sectPr.');
  for (const existing of Array.from(section.children).filter(
    (child) =>
      child.localName === 'paperSrc' || child.localName === 'sectPrChange',
  )) {
    existing.remove();
  }
  const current = document.createElementNS(WORD_NAMESPACE, 'w:paperSrc');
  current.setAttributeNS(
    WORD_NAMESPACE,
    'w:first',
    String(options.current.first),
  );
  current.setAttributeNS(
    WORD_NAMESPACE,
    'w:other',
    String(options.current.other),
  );
  section.append(current);
  const change = document.createElementNS(WORD_NAMESPACE, 'w:sectPrChange');
  change.setAttributeNS(WORD_NAMESPACE, 'w:id', '44');
  change.setAttributeNS(WORD_NAMESPACE, 'w:author', 'Reviewer');
  change.setAttributeNS(WORD_NAMESPACE, 'w:date', '2026-09-08T00:00:00Z');
  const prior = document.createElementNS(WORD_NAMESPACE, 'w:sectPr');
  const priorSource = document.createElementNS(WORD_NAMESPACE, 'w:paperSrc');
  priorSource.setAttributeNS(
    WORD_NAMESPACE,
    'w:first',
    String(options.prior.first),
  );
  priorSource.setAttributeNS(
    WORD_NAMESPACE,
    'w:other',
    String(options.prior.other),
  );
  prior.append(priorSource);
  change.append(prior);
  section.append(change);
  archive.file(
    'word/document.xml',
    new XMLSerializer().serializeToString(document),
  );
  return archive.generateAsync({ type: 'arraybuffer' });
}

async function sectionDocxWithColumnsChange(options: {
  prior: { count: number; spacing: number; separator: boolean };
  current: { count: number; spacing: number; separator: boolean };
}): Promise<ArrayBuffer> {
  const artifact = createArtifact('blank-document');
  if (artifact.content.type !== 'document') {
    throw new Error('Expected a document artifact.');
  }
  const seed = await createArtifactBlob(artifact);
  const archive = await JSZip.loadAsync(await seed.arrayBuffer());
  const document = await xmlEntry(archive, 'word/document.xml');
  const section = descendants(document, 'sectPr').find(
    (element) => element.parentElement?.localName !== 'sectPrChange',
  );
  if (!section) throw new Error('Expected body sectPr.');
  for (const existing of Array.from(section.children).filter(
    (child) => child.localName === 'cols' || child.localName === 'sectPrChange',
  )) {
    existing.remove();
  }
  section.append(createEqualWidthCols(document, options.current));
  const change = document.createElementNS(WORD_NAMESPACE, 'w:sectPrChange');
  change.setAttributeNS(WORD_NAMESPACE, 'w:id', '45');
  change.setAttributeNS(WORD_NAMESPACE, 'w:author', 'Reviewer');
  change.setAttributeNS(WORD_NAMESPACE, 'w:date', '2026-09-08T00:00:00Z');
  const prior = document.createElementNS(WORD_NAMESPACE, 'w:sectPr');
  prior.append(createEqualWidthCols(document, options.prior));
  change.append(prior);
  section.append(change);
  archive.file(
    'word/document.xml',
    new XMLSerializer().serializeToString(document),
  );
  return archive.generateAsync({ type: 'arraybuffer' });
}

async function sectionDocxWithUnequalColumnsChange(options: {
  prior: {
    count: number;
    spacing: number;
    separator: boolean;
    custom: Array<{ widthPercent: number; spacing: number }>;
  };
  current: {
    count: number;
    spacing: number;
    separator: boolean;
    custom: Array<{ widthPercent: number; spacing: number }>;
  };
}): Promise<ArrayBuffer> {
  const artifact = createArtifact('blank-document');
  if (artifact.content.type !== 'document') {
    throw new Error('Expected a document artifact.');
  }
  const seed = await createArtifactBlob(artifact);
  const archive = await JSZip.loadAsync(await seed.arrayBuffer());
  const document = await xmlEntry(archive, 'word/document.xml');
  const section = descendants(document, 'sectPr').find(
    (element) => element.parentElement?.localName !== 'sectPrChange',
  );
  if (!section) throw new Error('Expected body sectPr.');
  for (const existing of Array.from(section.children).filter(
    (child) => child.localName === 'cols' || child.localName === 'sectPrChange',
  )) {
    existing.remove();
  }
  section.append(createUnequalWidthCols(document, options.current));
  const change = document.createElementNS(WORD_NAMESPACE, 'w:sectPrChange');
  change.setAttributeNS(WORD_NAMESPACE, 'w:id', '46');
  change.setAttributeNS(WORD_NAMESPACE, 'w:author', 'Reviewer');
  change.setAttributeNS(WORD_NAMESPACE, 'w:date', '2026-09-08T00:00:00Z');
  const prior = document.createElementNS(WORD_NAMESPACE, 'w:sectPr');
  prior.append(createUnequalWidthCols(document, options.prior));
  change.append(prior);
  section.append(change);
  archive.file(
    'word/document.xml',
    new XMLSerializer().serializeToString(document),
  );
  return archive.generateAsync({ type: 'arraybuffer' });
}

function createUnequalWidthCols(
  document: Document,
  columns: {
    count: number;
    spacing: number;
    separator: boolean;
    custom: Array<{ widthPercent: number; spacing: number }>;
  },
): Element {
  const cols = document.createElementNS(WORD_NAMESPACE, 'w:cols');
  cols.setAttributeNS(WORD_NAMESPACE, 'w:num', String(columns.count));
  cols.setAttributeNS(WORD_NAMESPACE, 'w:equalWidth', '0');
  if (columns.separator) {
    cols.setAttributeNS(WORD_NAMESPACE, 'w:sep', '1');
  }
  for (const [index, column] of columns.custom.entries()) {
    const col = document.createElementNS(WORD_NAMESPACE, 'w:col');
    col.setAttributeNS(WORD_NAMESPACE, 'w:w', String(Math.round(column.widthPercent)));
    if (index < columns.custom.length - 1) {
      col.setAttributeNS(
        WORD_NAMESPACE,
        'w:space',
        String(Math.round((column.spacing * 1440) / 25.4)),
      );
    }
    cols.append(col);
  }
  return cols;
}

function createEqualWidthCols(
  document: Document,
  columns: { count: number; spacing: number; separator: boolean },
): Element {
  const cols = document.createElementNS(WORD_NAMESPACE, 'w:cols');
  cols.setAttributeNS(WORD_NAMESPACE, 'w:num', String(columns.count));
  cols.setAttributeNS(
    WORD_NAMESPACE,
    'w:space',
    String(Math.round((columns.spacing * 1440) / 25.4)),
  );
  if (columns.separator) {
    cols.setAttributeNS(WORD_NAMESPACE, 'w:sep', '1');
  }
  return cols;
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
