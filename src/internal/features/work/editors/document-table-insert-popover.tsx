import type { Editor } from '@tiptap/core';
import { officeMessage } from '../../../i18n/office-locale';
import { OfficeTableInsertPopover } from './office-table-insert-popover';
import { useOfficeMessages } from './office-messages-context';

export function DocumentTableInsertPopover({ editor }: { editor: Editor }) {
  const messages = useOfficeMessages();
  return (
    <OfficeTableInsertPopover
      className="work-document-table-insert-popover"
      label={officeMessage(messages, 'document.tableInsert.label')}
      onInsert={(dimensions) => editor.commands.insertDocumentTable(dimensions)}
    />
  );
}
