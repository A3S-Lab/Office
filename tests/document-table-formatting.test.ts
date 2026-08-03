import { Editor } from '@tiptap/core';
import { describe, expect, test } from '@rstest/core';
import {
  createArtifact,
  createArtifactBlob,
  importOfficeFile,
} from '../src/core';
import { createWorkDocumentExtensions } from '../src/internal/features/work/work-document-extensions';
import { activeDocumentTableStyle } from '../src/internal/features/work/work-document-table-cell-formatting';

describe('document table formatting', () => {
  test('formats every cell from a whole-table node selection', () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: [
        '<section data-document-section="true">',
        '<table><tbody>',
        '<tr><th><p>Title</p></th><th><p>Owner</p></th></tr>',
        '<tr><td><p>Plan</p></td><td><p>A3S</p></td></tr>',
        '</tbody></table>',
        '</section>',
      ].join(''),
    });
    editor.commands.setNodeSelection(firstTablePosition(editor));

    expect(editor.isActive('table')).toBe(true);
    expect(editor.commands.applyDocumentTableStyle('blueStripe')).toBe(true);
    expect(
      tableCellAttributes(editor).every(
        ({ borderColor, borderStyle }) =>
          borderColor === '#9fbad0' && borderStyle === 'solid',
      ),
    ).toBe(true);
    expect(
      editor.commands.setDocumentTableCellFormat({
        backgroundColor: '#fff2cc',
      }),
    ).toBe(true);
    expect(
      tableCellAttributes(editor).every(
        ({ backgroundColor }) => backgroundColor === '#fff2cc',
      ),
    ).toBe(true);
    expect(editor.commands.setDocumentTableHorizontalAlignment('center')).toBe(
      true,
    );
    expect(editor.getHTML().match(/text-align: center/g)).toHaveLength(4);

    editor.destroy();
  });

  test('keeps structured cell presentation in editable HTML', () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: formattedTableHtml(),
    });

    const cells = tableCellAttributes(editor);
    expect(cells[0]).toMatchObject({
      backgroundColor: '#d9eaf7',
      verticalAlign: 'middle',
      borderColor: '#4472c4',
      borderStyle: 'double',
      borderWidth: 2,
      margins: null,
    });
    expect(cells[1]).toMatchObject({
      backgroundColor: '#fff2cc',
      verticalAlign: 'bottom',
      borderColor: '#70ad47',
      borderStyle: 'dashed',
      borderWidth: 1,
    });

    const html = editor.getHTML();
    expect(html).toContain('data-office-cell-fill="#d9eaf7"');
    expect(html).toContain('data-office-cell-vertical-align="middle"');
    expect(html).toContain('data-office-cell-border-style="double"');
    expect(html).toContain('border: 2px double #4472c4');
    editor.destroy();
  });

  test('keeps partial cell-margin overrides in editable HTML', () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: [
        '<table><tbody><tr><td ',
        'data-office-cell-margin-top="8" ',
        'data-office-cell-margin-right="16" ',
        'style="padding-top: 8px; padding-right: 16px">',
        '<p>Inset</p></td></tr></tbody></table>',
      ].join(''),
    });
    editor.commands.setTextSelection(firstTablePosition(editor) + 4);

    expect(tableCellAttributes(editor)[0]?.margins).toEqual({
      top: 8,
      right: 16,
    });
    expect(
      editor.commands.setDocumentTableCellFormat({
        margins: { top: 4, right: 6, bottom: 4, left: 6 },
      }),
    ).toBe(true);
    expect(tableCellAttributes(editor)[0]?.margins).toEqual({
      top: 4,
      right: 6,
      bottom: 4,
      left: 6,
    });
    expect(editor.getHTML()).toContain('data-office-cell-margin-left="6"');
    const rendered = document.createElement('div');
    rendered.innerHTML = editor.getHTML();
    expect(rendered.querySelector<HTMLElement>('td')?.style.paddingLeft).toBe(
      '6px',
    );

    editor.destroy();
  });

  test('parses independent inline CSS borders without proprietary attributes', () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: [
        '<table><tbody><tr><td style="',
        'border-top: 2px double #4472c4;',
        'border-right: 1px dotted #c00000;',
        'border-bottom: 1px dashed #70ad47;',
        'border-left: 0.5px solid #000000;',
        '"><p>Mixed borders</p></td></tr></tbody></table>',
      ].join(''),
    });

    expect(tableCellAttributes(editor)[0]?.borders).toEqual({
      top: { color: '#4472c4', style: 'double', width: 2 },
      right: { color: '#c00000', style: 'dotted', width: 1 },
      bottom: { color: '#70ad47', style: 'dashed', width: 1 },
      left: { color: '#000000', style: 'solid', width: 0.5 },
    });
    editor.destroy();
  });

  test('round-trips common cell shading, alignment, and uniform borders through DOCX', async () => {
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    artifact.content.html = formattedTableHtml();

    const blob = await createArtifactBlob(artifact);
    const imported = await importOfficeFile(
      new File([blob], 'table-formatting.docx', { type: blob.type }),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: imported.content.html,
    });

    expect(tableCellAttributes(editor)[0]).toMatchObject({
      backgroundColor: '#d9eaf7',
      verticalAlign: 'middle',
      borderColor: '#4472c4',
      borderStyle: 'double',
      borderWidth: 2,
    });
    expect(tableCellAttributes(editor)[1]).toMatchObject({
      backgroundColor: '#fff2cc',
      verticalAlign: 'bottom',
      borderColor: '#70ad47',
      borderStyle: 'dashed',
      borderWidth: 1,
    });
    editor.destroy();
  });

  test('applies outside and inside borders without flattening each cell edge', async () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: [
        '<section data-document-section="true">',
        '<table><tbody>',
        '<tr><th><p>Title</p></th><th><p>Owner</p></th></tr>',
        '<tr><td><p>Plan</p></td><td><p>A3S</p></td></tr>',
        '</tbody></table>',
        '</section>',
      ].join(''),
    });
    editor.commands.setNodeSelection(firstTablePosition(editor));

    expect(
      editor.commands.setDocumentTableBorders('all', {
        color: '#000000',
        style: 'none',
        width: 0,
      }),
    ).toBe(true);
    expect(
      editor.commands.setDocumentTableBorders('outside', {
        color: '#4472c4',
        style: 'double',
        width: 2,
      }),
    ).toBe(true);
    expect(
      editor.commands.setDocumentTableBorders('insideHorizontal', {
        color: '#70ad47',
        style: 'dashed',
        width: 1,
      }),
    ).toBe(true);

    const [topLeft, topRight, bottomLeft, bottomRight] =
      tableCellAttributes(editor);
    const none = { color: '#000000', style: 'none', width: 0 };
    const outside = { color: '#4472c4', style: 'double', width: 2 };
    const inside = { color: '#70ad47', style: 'dashed', width: 1 };
    expect(topLeft?.borders).toEqual({
      top: outside,
      right: none,
      bottom: inside,
      left: outside,
    });
    expect(topRight?.borders).toEqual({
      top: outside,
      right: outside,
      bottom: inside,
      left: none,
    });
    expect(bottomLeft?.borders).toEqual({
      top: inside,
      right: none,
      bottom: outside,
      left: outside,
    });
    expect(bottomRight?.borders).toEqual({
      top: inside,
      right: outside,
      bottom: outside,
      left: none,
    });

    const html = editor.getHTML();
    expect(html).toContain('data-office-cell-border-top-style="double"');
    expect(html).toContain('data-office-cell-border-bottom-style="dashed"');
    const rendered = document.createElement('div');
    rendered.innerHTML = html;
    const renderedTopLeft = rendered.querySelector('th');
    expect(renderedTopLeft).not.toBeNull();
    expect((renderedTopLeft as HTMLElement).style.borderTopStyle).toBe(
      'double',
    );
    expect((renderedTopLeft as HTMLElement).style.borderTopWidth).toBe('2px');
    expect((renderedTopLeft as HTMLElement).style.borderBottomStyle).toBe(
      'dashed',
    );
    expect((renderedTopLeft as HTMLElement).style.borderBottomWidth).toBe(
      '1px',
    );

    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    artifact.content.html = html;
    const blob = await createArtifactBlob(artifact);
    const imported = await importOfficeFile(
      new File([blob], 'per-edge-borders.docx', { type: blob.type }),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const reopened = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: imported.content.html,
    });
    expect(tableCellAttributes(reopened).map(({ borders }) => borders)).toEqual(
      [topLeft, topRight, bottomLeft, bottomRight].map(
        ({ borders }) => borders,
      ),
    );

    reopened.destroy();
    editor.destroy();
  });

  test('applies side targets only to the corresponding selection boundary', () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: [
        '<section data-document-section="true">',
        '<table><tbody>',
        '<tr><td><p>A</p></td><td><p>B</p></td></tr>',
        '<tr><td><p>C</p></td><td><p>D</p></td></tr>',
        '</tbody></table>',
        '</section>',
      ].join(''),
    });
    editor.commands.setNodeSelection(firstTablePosition(editor));
    const none = { color: '#000000', style: 'none', width: 0 } as const;
    const top = { color: '#7030a0', style: 'dotted', width: 1 } as const;
    const right = { color: '#c00000', style: 'solid', width: 2 } as const;

    expect(activeDocumentTableStyle(editor.state)).toBe('grid');
    expect(editor.commands.setDocumentTableBorders('bottom', right)).toBe(true);
    expect(activeDocumentTableStyle(editor.state)).toBeNull();
    expect(editor.commands.setDocumentTableBorders('all', none)).toBe(true);
    expect(editor.commands.setDocumentTableBorders('top', top)).toBe(true);
    expect(editor.commands.setDocumentTableBorders('right', right)).toBe(true);

    const [topLeft, topRight, bottomLeft, bottomRight] =
      tableCellAttributes(editor);
    expect(topLeft?.borders).toEqual({
      top,
      right: none,
      bottom: none,
      left: none,
    });
    expect(topRight?.borders).toEqual({
      top,
      right,
      bottom: none,
      left: none,
    });
    expect(bottomLeft?.borders).toEqual({
      top: none,
      right: none,
      bottom: none,
      left: none,
    });
    expect(bottomRight?.borders).toEqual({
      top: none,
      right,
      bottom: none,
      left: none,
    });

    editor.destroy();
  });
});

function formattedTableHtml(): string {
  return [
    '<table><tbody><tr>',
    '<th data-office-cell-fill="#d9eaf7" ',
    'data-office-cell-vertical-align="middle" ',
    'data-office-cell-border-color="#4472c4" ',
    'data-office-cell-border-style="double" ',
    'data-office-cell-border-width="2" ',
    'style="background-color: #d9eaf7; vertical-align: middle; border: 2px double #4472c4;">',
    '<p style="text-align: center;">Header</p></th>',
    '<td data-office-cell-fill="#fff2cc" ',
    'data-office-cell-vertical-align="bottom" ',
    'data-office-cell-border-color="#70ad47" ',
    'data-office-cell-border-style="dashed" ',
    'data-office-cell-border-width="1" ',
    'style="background-color: #fff2cc; vertical-align: bottom; border: 1px dashed #70ad47;">',
    '<p style="text-align: right;">Value</p></td>',
    '</tr></tbody></table>',
  ].join('');
}

function tableCellAttributes(editor: Editor): Record<string, unknown>[] {
  const attributes: Record<string, unknown>[] = [];
  editor.state.doc.descendants((node) => {
    if (node.type.name === 'tableCell' || node.type.name === 'tableHeader') {
      attributes.push(node.attrs);
      return false;
    }
    return true;
  });
  return attributes;
}

function firstTablePosition(editor: Editor): number {
  let position: number | null = null;
  editor.state.doc.descendants((node, offset) => {
    if (position === null && node.type.name === 'table') position = offset;
    return position === null;
  });
  if (position === null) throw new Error('Expected a table node.');
  return position;
}
