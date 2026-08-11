import { describe, expect, test } from '@rstest/core';
import JSZip from 'jszip';
import {
  artifactExtension,
  createArtifact,
  createArtifactBlob,
  defaultPptxRuntimeUrl,
  importOfficeFile,
  OFFICE_FILE_ACCEPT,
  officeKindForFile,
} from '../src/core';

describe('office core', () => {
  test('creates complete blank artifacts', () => {
    const document = createArtifact('blank-document');
    const markdown = createArtifact('blank-markdown');
    const spreadsheet = createArtifact('blank-spreadsheet');
    const presentation = createArtifact('blank-presentation');

    expect(document.kind).toBe('document');
    expect(document.content).toEqual({
      type: 'document',
      pageSize: 'a4',
      html: '<p></p>',
    });
    expect(markdown.kind).toBe('markdown');
    expect(markdown.content).toEqual({ type: 'markdown', markdown: '' });
    expect(spreadsheet.kind).toBe('spreadsheet');
    expect(spreadsheet.content.type).toBe('spreadsheet');
    expect(presentation.kind).toBe('presentation');
    expect(presentation.content.type).toBe('presentation');
    if (presentation.content.type !== 'presentation')
      throw new Error('Expected a presentation artifact.');
    expect(presentation.content.slides[0]?.elements).toMatchObject([
      {
        text: '',
        placeholder: {
          key: 'title',
          type: 'title',
          prompt: '单击添加标题',
        },
      },
      {
        text: '',
        placeholder: {
          key: 'subtitle',
          type: 'subtitle',
          prompt: '添加副标题',
        },
      },
    ]);
    expect(artifactExtension(presentation.kind)).toBe('pptx');
  });

  test('detects supported files without reading their contents', () => {
    expect(officeKindForFile(new File([], 'proposal.docx'))).toBe('document');
    expect(officeKindForFile(new File([], 'readme.md'))).toBe('markdown');
    expect(officeKindForFile(new File([], 'notes.markdown'))).toBe('markdown');
    expect(officeKindForFile(new File([], 'forecast.xlsx'))).toBe(
      'spreadsheet',
    );
    expect(officeKindForFile(new File([], 'deck.pptx'))).toBe('presentation');
    expect(officeKindForFile(new File([], 'contract.pdf'))).toBe('pdf');
    expect(officeKindForFile(new File([], 'archive.zip'))).toBeNull();
    expect(OFFICE_FILE_ACCEPT).toContain('.docx');
    expect(OFFICE_FILE_ACCEPT).toContain('.md');
    expect(OFFICE_FILE_ACCEPT).toContain('.pdf');
  });

  test('imports and exports Markdown without converting its source', async () => {
    const source = '# Release notes\n\n- Fast\n- Controlled\n';
    const artifact = await importOfficeFile(
      new File([source], 'release-notes.md', { type: 'text/markdown' }),
    );

    expect(artifact.kind).toBe('markdown');
    expect(artifact.title).toBe('release-notes');
    expect(artifact.content).toEqual({ type: 'markdown', markdown: source });
    expect(artifactExtension(artifact.kind)).toBe('md');

    const output = await createArtifactBlob(artifact);
    expect(output.type).toContain('text/markdown');
    expect(await output.text()).toBe(source);
  });

  test('materializes imported PDF bytes before the source file can disappear', async () => {
    const bytes = new Uint8Array([
      0x25, 0x50, 0x44, 0x46, 0x2d, 0x31, 0x2e, 0x37, 0x0a,
    ]);
    const source = new File([bytes], 'contract.pdf', {
      lastModified: 1_765_756_800_000,
      type: 'application/pdf',
    });
    let sourceReads = 0;
    let finishSourceRead: ((bytes: ArrayBuffer) => void) | undefined;
    source.arrayBuffer = () => {
      sourceReads += 1;
      return new Promise<ArrayBuffer>((resolve) => {
        finishSourceRead = resolve;
      });
    };

    const pendingImport = importOfficeFile(source);

    expect(sourceReads).toBe(1);
    finishSourceRead?.(bytes.slice().buffer);
    const artifact = await pendingImport;

    expect(sourceReads).toBe(1);
    expect(artifact).toMatchObject({
      kind: 'pdf',
      title: 'contract',
      content: { type: 'pdf' },
      source: {
        name: 'contract.pdf',
        contentType: 'application/pdf',
        size: bytes.byteLength,
        updatedAt: 1_765_756_800_000,
      },
    });
    const output = await createArtifactBlob(artifact);
    expect(output).not.toBe(source);
    expect(output.type).toBe('application/pdf');
    expect(new Uint8Array(await output.arrayBuffer())).toEqual(bytes);
    expect(sourceReads).toBe(1);
  });

  test('imports browser documents into the versioned structured model', async () => {
    const artifact = await importOfficeFile(
      new File(['<h1>A3S Office</h1><p>Structured import.</p>'], 'brief.html', {
        type: 'text/html',
      }),
    );

    expect(artifact.content.type).toBe('document');
    if (artifact.content.type !== 'document')
      throw new Error('Expected an imported document artifact.');
    expect(artifact.content.model).toMatchObject({
      schema: 'a3s.office.document',
      version: 1,
      revision: 1,
      root: { type: 'doc' },
    });
    expect(artifact.content.model?.root.content?.[0]).toMatchObject({
      type: 'documentSection',
    });
  });

  test('exports a blank spreadsheet as a real workbook blob', async () => {
    const blob = await createArtifactBlob(createArtifact('blank-spreadsheet'));

    expect(blob.type).toBe(
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    );
    expect(blob.size).toBeGreaterThan(1_000);
  });

  test('publishes a colocated default PowerPoint runtime URL', () => {
    expect(new URL(defaultPptxRuntimeUrl).pathname).toMatch(
      /\/pptxgen\.bundle\.js$/,
    );
  });

  test('keeps document typography and paragraph layout in DOCX export', async () => {
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document')
      throw new Error('Expected a document artifact.');
    artifact.content.html = [
      '<p dir="rtl" style="text-align: justify; line-height: 18pt; margin-top: 12pt;',
      ' margin-right: 18px; margin-bottom: 6pt; margin-left: 24px;',
      ' text-indent: -12px;"',
      ' data-office-indent-level="1" data-office-indent-right="18"',
      ' data-office-indent-first-line="-12"',
      ' data-office-space-before="12" data-office-space-after="6"',
      ' data-office-line-rule="exact"',
      ' data-office-keep-lines="true" data-office-keep-with-next="true"',
      ' data-office-page-break-before="true"',
      ' data-office-widow-control="false"',
      ' data-office-tab-stops=\'[{"position":48,"alignment":"left",',
      '"leader":"none"},{"position":144,"alignment":"right",',
      '"leader":"dot"}]\'>',
      '<span style="font-family: Arial; font-size: 12pt; color: #336699;',
      ' background-color: #ffff00;">A3S</span>',
      '<span data-document-tab="true"></span><span>Office</span>',
      '</p>',
    ].join('');

    const blob = await createArtifactBlob(artifact);
    const archive = await JSZip.loadAsync(await blob.arrayBuffer());
    const xml = await archive.file('word/document.xml')?.async('string');

    expect(xml).toBeDefined();
    expect(xml).not.toContain('<w:docGrid');
    expect(xml).toContain('<w:rFonts w:ascii="Arial"');
    expect(xml).toContain('w:hAnsi="Arial"');
    expect(xml).toContain('<w:sz w:val="24"/>');
    expect(xml).toContain('<w:color w:val="336699"/>');
    expect(xml).toContain('w:fill="FFFF00"');
    expect(xml).toContain('<w:jc w:val="both"/>');
    expect(xml).toContain('<w:bidi/>');
    expect(xml).toContain('<w:rtl/>');
    expect(xml).toContain(
      '<w:spacing w:after="120" w:before="240" w:line="360" w:lineRule="exact"/>',
    );
    expect(xml).toContain(
      '<w:ind w:left="360" w:right="270" w:hanging="180"/>',
    );
    expect(xml).toContain('<w:keepNext/>');
    expect(xml).toContain('<w:keepLines/>');
    expect(xml).toContain('<w:pageBreakBefore/>');
    expect(xml).toContain('<w:widowControl w:val="false"/>');
    expect(xml).toContain('<w:tabs>');
    expect(xml).toContain('<w:tab w:val="left" w:pos="720"/>');
    expect(xml).toContain('<w:tab w:val="right" w:pos="2160" w:leader="dot"/>');
    expect(xml).toContain('<w:tab/>');

    const imported = await importOfficeFile(
      new File([blob], 'paragraph-layout.docx', { type: blob.type }),
    );
    expect(imported.content.type).toBe('document');
    if (imported.content.type !== 'document')
      throw new Error('Expected an imported document artifact.');
    expect(imported.content.html).toContain('data-office-indent-level="1"');
    expect(imported.content.html).toContain('data-office-indent-right="18"');
    expect(imported.content.html).toContain(
      'data-office-indent-first-line="-12"',
    );
    expect(imported.content.html).toContain('data-office-space-before="12"');
    expect(imported.content.html).toContain('data-office-space-after="6"');
    expect(imported.content.html).toContain('data-office-line-rule="exact"');
    expect(imported.content.html).toContain('dir="rtl"');
    expect(imported.content.html).toContain('font-family: Arial');
    expect(imported.content.html).toContain('font-size: 12pt');
    expect(imported.content.html).toContain('color: #336699');
    expect(imported.content.html).toContain('background-color: #ffff00');
    expect(imported.content.html).toContain('data-office-keep-lines="true"');
    expect(imported.content.html).toContain(
      'data-office-keep-with-next="true"',
    );
    expect(imported.content.html).toContain(
      'data-office-page-break-before="true"',
    );
    expect(imported.content.html).toContain(
      'data-office-widow-control="false"',
    );
    expect(imported.content.html).toContain('data-office-tab-stops=');
    expect(imported.content.html).toContain('data-document-tab="true"');
    expect(imported.content.model).toMatchObject({
      schema: 'a3s.office.document',
      version: 1,
      revision: 1,
    });
  });

  test('round-trips Word automatic line multiples without exporting the WPS render metric', async () => {
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    artifact.content.html = [
      '<p style="line-height: 1.5;',
      ' --work-document-office-auto-line-height: 1.725"',
      ' data-office-line-rule="auto"',
      ' data-office-auto-line-height="1.725">',
      '<span style="font-family: Arial; font-size: 16pt">',
      'Automatic line spacing</span></p>',
    ].join('');

    const blob = await createArtifactBlob(artifact);
    const archive = await JSZip.loadAsync(await blob.arrayBuffer());
    const xml = await archive.file('word/document.xml')?.async('string');

    expect(xml).toContain(
      '<w:spacing w:after="120" w:line="360" w:lineRule="auto"/>',
    );

    const imported = await importOfficeFile(
      new File([blob], 'automatic-line-spacing.docx', { type: blob.type }),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    expect(imported.content.html).toContain('line-height: 1.5');
    expect(imported.content.html).toContain(
      'data-office-auto-line-height="1.725"',
    );
  });

  test('round-trips the Word document grid and run grid override', async () => {
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    artifact.content.html = [
      '<section data-document-section="true" data-section-id="grid-section"',
      ' data-section-document-grid-type="lines"',
      ' data-section-document-grid-line-pitch="18">',
      '<p data-office-line-rule="auto"',
      ' data-office-auto-line-height="1.15" style="line-height: 1">',
      '<span style="font-family: Microsoft YaHei; font-size: 10.5pt"',
      ' data-office-word-line-height-factor="1.7143"',
      ' data-office-word-snap-to-grid="false">Grid off</span>',
      '<span style="font-family: Arial; font-size: 10.5pt"',
      ' data-office-word-line-height-factor="1.15"',
      ' data-office-word-snap-to-grid="true">Grid on</span>',
      '</p></section>',
      '<section data-document-section="true" data-section-id="grid-section-2"',
      ' data-section-document-grid-type="linesAndChars"',
      ' data-section-document-grid-line-pitch="24">',
      '<p><span style="font-family: Arial; font-size: 10.5pt">',
      'Second grid</span></p></section>',
    ].join('');

    const blob = await createArtifactBlob(artifact);
    const archive = await JSZip.loadAsync(await blob.arrayBuffer());
    const xml = await archive.file('word/document.xml')?.async('string');

    expect(xml).toContain('<w:docGrid w:type="lines" w:linePitch="360"/>');
    expect(xml).toContain(
      '<w:docGrid w:type="linesAndChars" w:linePitch="480"/>',
    );
    expect(xml).toContain('<w:snapToGrid w:val="false"/>');
    expect(xml).toContain('<w:snapToGrid/>');
    const snapToGrid = Array.from(
      (xml ?? '').matchAll(/<w:snapToGrid(?:\s[^>]*)?\/>/g),
      ([element]) => element,
    );
    expect(snapToGrid).toHaveLength(2);
    const runProperties = Array.from(
      (xml ?? '').matchAll(/<w:rPr(?:\s[^>]*)?>[\s\S]*?<\/w:rPr>/g),
      ([element]) => element,
    );
    expect(
      runProperties.filter((element) => element.includes('<w:snapToGrid')),
    ).toHaveLength(2);

    const imported = await importOfficeFile(
      new File([blob], 'document-grid.docx', { type: blob.type }),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    expect(imported.content.html).toContain(
      'data-section-document-grid-type="lines"',
    );
    expect(imported.content.html).toContain(
      'data-section-document-grid-line-pitch="18"',
    );
    expect(imported.content.html).toContain(
      'data-section-document-grid-type="linesAndChars"',
    );
    expect(imported.content.html).toContain(
      'data-section-document-grid-line-pitch="24"',
    );
    expect(imported.content.html).toContain(
      'data-office-word-snap-to-grid="false"',
    );
    expect(imported.content.html).toContain(
      'data-office-word-snap-to-grid="true"',
    );
  });

  test('round-trips structured DOCX lists, nesting, starts, and RTL items', async () => {
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document')
      throw new Error('Expected a document artifact.');
    artifact.content.html = [
      '<ol start="3">',
      '<li><p>Third item</p>',
      '<ol><li><p>Nested item</p></li></ol>',
      '</li>',
      '<li><p>Fourth item</p></li>',
      '<li dir="rtl"><p>مرحبا A3S שלום</p></li>',
      '</ol>',
      '<ul data-office-bullet-style="square" style="list-style-type: square">',
      '<li><p>Bullet item</p></li></ul>',
    ].join('');

    const blob = await createArtifactBlob(artifact);
    const archive = await JSZip.loadAsync(await blob.arrayBuffer());
    const documentXml = await archive
      .file('word/document.xml')
      ?.async('string');
    const numberingXml = await archive
      .file('word/numbering.xml')
      ?.async('string');

    expect(documentXml).toBeDefined();
    expect(numberingXml).toBeDefined();
    const paragraphs = Array.from(
      (documentXml ?? '').matchAll(/<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/g),
      ([paragraph]) => paragraph,
    ).filter((paragraph) => paragraph.includes('<w:t'));
    expect(
      paragraphs.map((paragraph) =>
        Array.from(
          paragraph.matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g),
          (match) => match[1] ?? '',
        ).join(''),
      ),
    ).toEqual([
      'Third item',
      'Nested item',
      'Fourth item',
      'مرحبا A3S שלום',
      'Bullet item',
    ]);
    expect(
      paragraphs.filter((paragraph) => paragraph.includes('<w:numPr>')),
    ).toHaveLength(5);
    const rtl = paragraphs.find((paragraph) => paragraph.includes('مرحبا'));
    expect(rtl).toContain('<w:bidi/>');
    expect(rtl).toContain('<w:rtl/>');
    expect(numberingXml).toContain('<w:start w:val="3"/>');
    expect(numberingXml).toContain('<w:lvl w:ilvl="1"');
    expect(numberingXml).toContain('<w:lvlText w:val="■"/>');

    const imported = await importOfficeFile(
      new File([blob], 'structured-lists.docx', { type: blob.type }),
    );
    expect(imported.content.type).toBe('document');
    if (imported.content.type !== 'document')
      throw new Error('Expected an imported document artifact.');
    const importedDocument = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    const importedOrderedLists = Array.from(
      importedDocument.body.querySelectorAll('ol'),
    );
    expect(importedOrderedLists[0]).toHaveAttribute('start', '3');
    expect(importedOrderedLists[0]).toHaveAttribute('data-office-numbering-id');
    expect(importedOrderedLists[0]).toHaveAttribute(
      'data-office-abstract-numbering-id',
    );
    expect(importedOrderedLists[0]).toHaveAttribute(
      'data-office-numbering-level',
      '0',
    );
    expect(importedOrderedLists[1]).toHaveAttribute('type', 'a');
    expect(importedOrderedLists[1]).toHaveAttribute(
      'data-office-numbering-level',
      '1',
    );
    expect(imported.content.html).toContain(
      'data-office-bullet-style="square"',
    );
    expect(imported.content.html).toContain('<li dir="rtl">');
    expect(imported.content.html).toContain('Nested item');
    expect(imported.content.html).toContain('مرحبا A3S שלום');
  });

  test('keeps structured numbering in DOCX table cells, notes, and page chrome', async () => {
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document')
      throw new Error('Expected a document artifact.');
    artifact.content.html = [
      '<p>Body note',
      '<sup data-document-note-reference="true" data-note-kind="footnote" ',
      'data-note-id="note-1" data-note-number="1">1</sup></p>',
      '<table><tbody><tr><td>',
      '<ol start="2"><li><p>Cell second</p></li></ol>',
      '</td></tr></tbody></table>',
      '<aside data-document-note="true" data-note-kind="footnote" ',
      'data-note-id="note-1" data-note-number="1">',
      '<ol start="4"><li><p>Note fourth</p></li></ol>',
      '</aside>',
    ].join('');
    artifact.content.pageChrome = {
      differentFirstPage: false,
      differentOddEvenPages: false,
      default: {
        headerHtml: '<ol start="6"><li><p>Header sixth</p></li></ol>',
        footerHtml: '',
        showPageNumber: false,
      },
      first: { headerHtml: '', footerHtml: '', showPageNumber: false },
      even: { headerHtml: '', footerHtml: '', showPageNumber: false },
    };

    const blob = await createArtifactBlob(artifact);
    const archive = await JSZip.loadAsync(await blob.arrayBuffer());
    const documentXml =
      (await archive.file('word/document.xml')?.async('string')) ?? '';
    const footnotesXml =
      (await archive.file('word/footnotes.xml')?.async('string')) ?? '';
    const numberingXml =
      (await archive.file('word/numbering.xml')?.async('string')) ?? '';
    const headerXml = (
      await Promise.all(
        Object.keys(archive.files)
          .filter((path) => /^word\/header\d+\.xml$/.test(path))
          .map(async (path) => archive.file(path)?.async('string')),
      )
    ).find((xml) => xml?.includes('Header sixth'));

    expect(paragraphXmlContaining(documentXml, 'Cell second')).toContain(
      '<w:numPr>',
    );
    expect(paragraphXmlContaining(footnotesXml, 'Note fourth')).toContain(
      '<w:numPr>',
    );
    expect(paragraphXmlContaining(headerXml ?? '', 'Header sixth')).toContain(
      '<w:numPr>',
    );
    expect(numberingXml).toContain('<w:start w:val="2"/>');
    expect(numberingXml).toContain('<w:start w:val="4"/>');
    expect(numberingXml).toContain('<w:start w:val="6"/>');
  });

  test('round-trips DOCX table row pagination properties', async () => {
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document')
      throw new Error('Expected a document artifact.');
    artifact.content.html = [
      '<table><tbody>',
      '<tr data-office-repeat-header="true" data-office-cant-split="true">',
      '<th><p>项目</p></th><th><p>负责人</p></th>',
      '</tr>',
      '<tr data-office-cant-split="false">',
      '<td><p>A3S Office</p></td><td><p>A3S Lab</p></td>',
      '</tr>',
      '</tbody></table>',
    ].join('');

    const blob = await createArtifactBlob(artifact);
    const archive = await JSZip.loadAsync(await blob.arrayBuffer());
    const xml = await archive.file('word/document.xml')?.async('string');

    expect(xml).toContain('<w:tblHeader/>');
    expect(xml).toContain('<w:cantSplit/>');
    expect(xml).toContain('<w:cantSplit w:val="false"/>');

    const imported = await importOfficeFile(
      new File([blob], 'table-pagination.docx', { type: blob.type }),
    );
    expect(imported.content.type).toBe('document');
    if (imported.content.type !== 'document')
      throw new Error('Expected an imported document artifact.');
    expect(imported.content.html).toContain('data-office-repeat-header="true"');
    expect(imported.content.html).toContain('data-office-cant-split="true"');
    expect(imported.content.html).toContain('data-office-cant-split="false"');
  });

  test('round-trips DOCX table layout, exact column widths, and row heights', async () => {
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document')
      throw new Error('Expected a document artifact.');
    artifact.content.html = [
      '<table data-office-table-layout="fixed"><tbody>',
      '<tr data-office-row-height="48" data-office-row-height-rule="exact">',
      '<td colwidth="120"><p>Left</p></td>',
      '<td colwidth="180"><p>Right</p></td>',
      '</tr>',
      '</tbody></table>',
    ].join('');

    const blob = await createArtifactBlob(artifact);
    const archive = await JSZip.loadAsync(await blob.arrayBuffer());
    const xml =
      (await archive.file('word/document.xml')?.async('string')) ?? '';

    expect(xml).toContain('<w:tblLayout w:type="fixed"/>');
    expect(xml).toContain('<w:gridCol w:w="1800"/>');
    expect(xml).toContain('<w:gridCol w:w="2700"/>');
    expect(xml).toMatch(/<w:trHeight[^>]*w:val="720"[^>]*w:hRule="exact"/);

    const imported = await importOfficeFile(
      new File([blob], 'table-sizing.docx', { type: blob.type }),
    );
    if (imported.content.type !== 'document')
      throw new Error('Expected an imported document artifact.');
    expect(imported.content.html).toContain('data-office-table-layout="fixed"');
    expect(imported.content.html).toContain('colwidth="120"');
    expect(imported.content.html).toContain('colwidth="180"');
    expect(imported.content.html).toContain('data-office-row-height="48"');
    expect(imported.content.html).toContain(
      'data-office-row-height-rule="exact"',
    );
  });

  test('round-trips percentage column preferences with pixel grid fallbacks', async () => {
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    artifact.content.html = [
      '<table data-office-table-layout="fixed" data-office-table-width-type="percent" data-office-table-width="100"><tbody><tr>',
      '<td colwidth="120" data-office-column-widths-percent="40"><p>Left</p></td>',
      '<td colwidth="180" data-office-column-widths-percent="60"><p>Right</p></td>',
      '</tr></tbody></table>',
    ].join('');

    const blob = await createArtifactBlob(artifact);
    const archive = await JSZip.loadAsync(await blob.arrayBuffer());
    const xml =
      (await archive.file('word/document.xml')?.async('string')) ?? '';
    expect(xml).toContain('<w:gridCol w:w="1800"/>');
    expect(xml).toContain('<w:gridCol w:w="2700"/>');
    expect(xml).toMatch(/<w:tcW\b(?=[^>]*w:type="pct")(?=[^>]*w:w="40%")/);
    expect(xml).toMatch(/<w:tcW\b(?=[^>]*w:type="pct")(?=[^>]*w:w="60%")/);

    const imported = await importOfficeFile(
      new File([blob], 'percentage-columns.docx', { type: blob.type }),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    expect(imported.content.html).toContain(
      'data-office-column-widths-percent="40"',
    );
    expect(imported.content.html).toContain(
      'data-office-column-widths-percent="60"',
    );
    expect(imported.content.html).toContain('colwidth="120"');
    expect(imported.content.html).toContain('colwidth="180"');
  });

  test('round-trips independent DOCX table width, position, and cell margins', async () => {
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    artifact.content.html = [
      '<table data-office-table-layout="autofit" ',
      'data-office-table-width-type="percent" ',
      'data-office-table-width="62.5" ',
      'data-office-table-alignment="center" ',
      'data-office-table-indent="8" ',
      'data-office-table-cell-margin-top="0" ',
      'data-office-table-cell-margin-right="14.4" ',
      'data-office-table-cell-margin-bottom="4.8" ',
      'data-office-table-cell-margin-left="9.6"><tbody><tr>',
      '<td colwidth="160" data-office-cell-margin-top="8" ',
      'data-office-cell-margin-right="16"><p>Geometry</p></td>',
      '</tr></tbody></table>',
    ].join('');

    const blob = await createArtifactBlob(artifact);
    const archive = await JSZip.loadAsync(await blob.arrayBuffer());
    const xml =
      (await archive.file('word/document.xml')?.async('string')) ?? '';

    expect(xml).toMatch(/<w:tblW\b(?=[^>]*w:type="pct")(?=[^>]*w:w="62\.5%")/);
    expect(xml).toContain('<w:tblLayout w:type="autofit"/>');
    expect(xml).toContain('<w:jc w:val="center"/>');
    expect(xml).toMatch(/<w:tblInd\b(?=[^>]*w:type="dxa")(?=[^>]*w:w="120")/);
    expect(xml).toMatch(
      /<w:tblCellMar>[\s\S]*?<w:top\b[^>]*w:w="0"[\s\S]*?<w:left\b[^>]*w:w="144"[\s\S]*?<w:bottom\b[^>]*w:w="72"[\s\S]*?<w:right\b[^>]*w:w="216"/,
    );
    expect(xml).toMatch(
      /<w:tcMar>[\s\S]*?<w:top\b[^>]*w:w="120"[\s\S]*?<w:right\b[^>]*w:w="240"/,
    );

    const imported = await importOfficeFile(
      new File([blob], 'table-geometry.docx', { type: blob.type }),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    expect(imported.content.html).toContain(
      'data-office-table-layout="autofit"',
    );
    expect(imported.content.html).toContain(
      'data-office-table-width-type="percent"',
    );
    expect(imported.content.html).toContain('data-office-table-width="62.5"');
    expect(imported.content.html).toContain(
      'data-office-table-alignment="center"',
    );
    expect(imported.content.html).toContain(
      'data-office-table-cell-margin-left="9.6"',
    );
    expect(imported.content.html).toContain(
      'data-office-cell-margin-right="16"',
    );
    expect(imported.compatibility.issues).toContainEqual(
      expect.objectContaining({
        code: 'docx.tables',
        message: expect.stringContaining(
          'preferred auto, percentage, or pixel width',
        ),
      }),
    );
  });

  test('materializes DOCX export from the structured model instead of stale HTML', async () => {
    const artifact = await importOfficeFile(
      new File(['<p>HTML cache</p>'], 'model-source.html', {
        type: 'text/html',
      }),
    );
    if (
      artifact.content.type !== 'document' ||
      !artifact.content.model?.root.content?.[0]?.content?.[0]?.content?.[0]
    ) {
      throw new Error('Expected an imported structured document.');
    }
    const model = structuredClone(artifact.content.model);
    const text = model.root.content?.[0]?.content?.[0]?.content?.[0];
    if (!text) throw new Error('Expected an imported text node.');
    text.text = 'Model-owned export';
    model.revision += 1;
    artifact.content = { ...artifact.content, model };

    const blob = await createArtifactBlob(artifact);
    const archive = await JSZip.loadAsync(await blob.arrayBuffer());
    const xml = await archive.file('word/document.xml')?.async('string');

    expect(xml).toContain('Model-owned export');
    expect(xml).not.toContain('HTML cache');
  });
});

function paragraphXmlContaining(xml: string, text: string): string {
  return (
    Array.from(
      xml.matchAll(/<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/g),
      ([paragraph]) => paragraph,
    ).find((paragraph) => paragraph.includes(text)) ?? ''
  );
}
