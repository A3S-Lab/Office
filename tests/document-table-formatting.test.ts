import { Editor } from '@tiptap/core';
import { describe, expect, test } from '@rstest/core';
import {
  createArtifact,
  createArtifactBlob,
  importOfficeFile,
} from '../src/core';
import { createWorkDocumentExtensions } from '../src/internal/features/work/work-document-extensions';

describe('document table formatting', () => {
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
