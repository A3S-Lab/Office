import { Editor } from '@tiptap/core';
import type { Node as ProseMirrorNode } from '@tiptap/pm/model';
import { describe, expect, test } from '@rstest/core';
import { createWorkDocumentExtensions } from '../src/internal/features/work/work-document-extensions';
import {
  normalizeDocumentTableRowIdentity,
  type WorkDocumentTableRowIdentity,
} from '../src/internal/features/work/work-document-table-row-identity';

describe('document table-row identity', () => {
  test('retains format-only versions and rotates text and structure versions', () => {
    const editor = rowEditor(tableHtml('Alpha', '1A2B3C4D', '1A2B3C4E'));
    const initial = rowIdentities(editor)[0];
    expect(initial).toEqual({
      rowId: '1A2B3C4D',
      rowTextId: '1A2B3C4E',
    });

    const alpha = textPosition(editor, 'Alpha');
    editor.commands.setTextSelection({ from: alpha, to: alpha + 5 });
    expect(editor.commands.toggleBold()).toBe(true);
    expect(rowIdentities(editor)[0]).toEqual(initial);

    expect(editor.commands.insertContentAt(alpha + 1, '!')).toBe(true);
    const textEdited = rowIdentities(editor)[0];
    expect(textEdited?.rowId).toBe(initial?.rowId);
    expect(textEdited?.rowTextId).not.toBe(initial?.rowTextId);

    editor.commands.setTextSelection(textPosition(editor, 'A!lpha'));
    expect(editor.commands.addColumnAfter()).toBe(true);
    const structurallyEdited = rowIdentities(editor)[0];
    expect(structurallyEdited?.rowId).toBe(initial?.rowId);
    expect(structurallyEdited?.rowTextId).not.toBe(textEdited?.rowTextId);
    editor.destroy();
  });

  test('repairs copied rows without assigning identities to ordinary rows', () => {
    const duplicate = rowHtml('First', '2A2B3C4D', '2A2B3C4E');
    const editor = rowEditor(
      `<table><tbody>${duplicate}${duplicate}${rowHtml('Ordinary')}</tbody></table>`,
    );
    const rows = tableRows(editor);
    const identities = rows.flatMap(({ node }) => {
      const identity = normalizeDocumentTableRowIdentity({
        rowId: node.attrs.rowId,
        rowTextId: node.attrs.rowTextId,
      });
      return identity ? [identity] : [];
    });
    expect(identities).toHaveLength(2);
    expect(new Set(identities.map((identity) => identity.rowId)).size).toBe(2);
    expect(identities[0]).toEqual({
      rowId: '2A2B3C4D',
      rowTextId: '2A2B3C4E',
    });
    expect(rows[2].node.attrs.rowId).toBeNull();
    expect(rows[2].node.attrs.rowTextId).toBeNull();
    editor.destroy();
  });

  test('keeps identities attached when rows move', () => {
    const editor = rowEditor(
      `<table><tbody>${rowHtml('First', '3A2B3C4D', '3A2B3C4E')}${rowHtml(
        'Second',
        '4A2B3C4D',
        '4A2B3C4E',
      )}</tbody></table>`,
    );
    const [first, second] = tableRows(editor);
    const transaction = editor.state.tr
      .delete(second.position, second.position + second.node.nodeSize)
      .insert(first.position, second.node);
    editor.view.dispatch(transaction);
    expect(rowIdentityByText(editor, 'Second')).toEqual({
      rowId: '4A2B3C4D',
      rowTextId: '4A2B3C4E',
    });
    expect(rowIdentityByText(editor, 'First')).toEqual({
      rowId: '3A2B3C4D',
      rowTextId: '3A2B3C4E',
    });
    editor.destroy();
  });
});

function rowEditor(content: string): Editor {
  return new Editor({ extensions: createWorkDocumentExtensions(), content });
}

function tableHtml(text: string, rowId: string, rowTextId: string): string {
  return `<table><tbody>${rowHtml(text, rowId, rowTextId)}</tbody></table>`;
}

function rowHtml(text: string, rowId = '', rowTextId = ''): string {
  const identity =
    rowId && rowTextId
      ? ` data-office-row-id="${rowId}" data-office-row-text-id="${rowTextId}"`
      : '';
  return `<tr${identity}><td><p>${text}</p></td></tr>`;
}

function tableRows(editor: Editor): Array<{
  node: ProseMirrorNode;
  position: number;
}> {
  const rows: Array<{ node: ProseMirrorNode; position: number }> = [];
  editor.state.doc.descendants((node, position) => {
    if (node.type.name === 'tableRow') rows.push({ node, position });
  });
  return rows;
}

function rowIdentities(editor: Editor): WorkDocumentTableRowIdentity[] {
  return tableRows(editor).flatMap(({ node }) => {
    const identity = normalizeDocumentTableRowIdentity({
      rowId: node.attrs.rowId,
      rowTextId: node.attrs.rowTextId,
    });
    return identity ? [identity] : [];
  });
}

function rowIdentityByText(
  editor: Editor,
  text: string,
): WorkDocumentTableRowIdentity | null {
  const row = tableRows(editor).find(({ node }) => node.textContent === text);
  return row
    ? normalizeDocumentTableRowIdentity({
        rowId: row.node.attrs.rowId,
        rowTextId: row.node.attrs.rowTextId,
      })
    : null;
}

function textPosition(editor: Editor, text: string): number {
  let result = -1;
  editor.state.doc.descendants((node, position) => {
    if (result < 0 && node.isText && node.text === text) result = position;
  });
  if (result < 0) throw new Error(`Missing text node: ${text}`);
  return result;
}
