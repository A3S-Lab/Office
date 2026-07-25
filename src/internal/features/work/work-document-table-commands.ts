import { type CommandProps, Extension } from '@tiptap/core';
import { createTable } from '@tiptap/extension-table';
import type {
  Node as ProseMirrorNode,
  NodeType,
  ResolvedPos,
} from '@tiptap/pm/model';
import { TextSelection } from '@tiptap/pm/state';

export interface DocumentTableDimensions {
  rows: number;
  columns: number;
}

export interface InsertDocumentTableOptions {
  headerRow?: boolean;
  restoreFocus?: boolean;
}

declare module '@tiptap/core' {
  interface Commands<ReturnType> {
    documentTableCommands: {
      insertDocumentTable: (
        dimensions: DocumentTableDimensions,
        options?: InsertDocumentTableOptions,
      ) => ReturnType;
    };
  }
}

export const DocumentTableCommands = Extension.create({
  name: 'documentTableCommands',

  addCommands() {
    return {
      insertDocumentTable:
        (dimensions, options = {}) =>
        (props) =>
          insertDocumentTable(props, dimensions, options),
    };
  },
});

function insertDocumentTable(
  { chain, dispatch, editor, state }: CommandProps,
  dimensions: DocumentTableDimensions,
  options: InsertDocumentTableOptions,
): boolean {
  const rows = positiveInteger(dimensions.rows);
  const columns = positiveInteger(dimensions.columns);
  if (rows === null || columns === null) return false;

  const headerRow = options.headerRow !== false;
  const selection = state.selection;
  if (!selection.empty) {
    const createdTable = createTable(state.schema, rows, columns, headerRow);
    const table = headerRow
      ? withRepeatingFirstRow(createdTable)
      : createdTable;
    const insertionPosition = blockInsertionPosition(selection.$to, table.type);
    if (insertionPosition !== null) {
      const transaction = state.tr.insert(insertionPosition, table);
      transaction
        .setSelection(
          TextSelection.near(transaction.doc.resolve(insertionPosition + 1)),
        )
        .scrollIntoView();
      dispatch?.(transaction);
      if (dispatch && options.restoreFocus !== false) editor.commands.focus();
      return true;
    }
  }

  let commandChain = chain();
  if (options.restoreFocus !== false) commandChain = commandChain.focus();
  if (!selection.empty)
    commandChain = commandChain.setTextSelection(selection.to);
  commandChain = commandChain.insertTable({
    rows,
    cols: columns,
    withHeaderRow: headerRow,
  });
  if (headerRow) {
    commandChain = commandChain.updateAttributes('tableRow', {
      repeatHeader: true,
    });
  }
  return commandChain.run();
}

function positiveInteger(value: number): number | null {
  if (!Number.isSafeInteger(value) || value < 1) return null;
  return value;
}

function blockInsertionPosition(
  position: ResolvedPos,
  tableType: NodeType,
): number | null {
  for (let depth = position.depth; depth > 0; depth -= 1) {
    const parentDepth = depth - 1;
    const parent = position.node(parentDepth);
    const index = position.indexAfter(parentDepth);
    if (parent.canReplaceWith(index, index, tableType)) {
      return position.after(depth);
    }
  }
  return null;
}

function withRepeatingFirstRow(table: ProseMirrorNode): ProseMirrorNode {
  const firstRow = table.firstChild;
  if (!firstRow) return table;
  const repeatedRow = firstRow.type.create(
    { ...firstRow.attrs, repeatHeader: true },
    firstRow.content,
    firstRow.marks,
  );
  return table.copy(table.content.replaceChild(0, repeatedRow));
}
