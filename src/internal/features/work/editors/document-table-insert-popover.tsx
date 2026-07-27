import type { Editor } from '@tiptap/core';
import { OfficeTableInsertPopover } from './office-table-insert-popover';

export function DocumentTableInsertPopover({ editor }: { editor: Editor }) {
  return (
    <OfficeTableInsertPopover
      className="work-document-table-insert-popover"
      label="插入表格"
      onInsert={(dimensions) => editor.commands.insertDocumentTable(dimensions)}
    />
  );
}
