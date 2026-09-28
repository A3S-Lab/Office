import type { Editor } from '@tiptap/core';
import { Pilcrow } from 'lucide-react';
import { useCallback, useRef, useSyncExternalStore } from 'react';
import { Popover } from '../../../design-system/primitives';
import { officeMessage } from '../../../i18n/office-locale';
import { documentParagraphPagination } from '../work-document-paragraph-formatting';
import { OfficeCheckbox } from './office-controls';
import { useOfficeMessages } from './office-messages-context';

export function DocumentPaginationPopover({ editor }: { editor: Editor }) {
  const messages = useOfficeMessages();
  const panelRef = useRef<HTMLElement>(null);
  const subscribe = useCallback(
    (notify: () => void) => {
      editor.on('transaction', notify);
      return () => editor.off('transaction', notify);
    },
    [editor],
  );
  useSyncExternalStore(
    subscribe,
    () => documentPaginationSnapshot(editor),
    () => documentPaginationSnapshot(editor),
  );
  const pagination = documentParagraphPagination(editor);
  const attributes = editor.isActive('heading')
    ? editor.getAttributes('heading')
    : editor.getAttributes('paragraph');
  const customized = [
    attributes.keepLines,
    attributes.keepWithNext,
    attributes.pageBreakBefore,
    attributes.widowControl,
  ].some((value) => typeof value === 'boolean');
  const update = (key: keyof typeof pagination, checked: boolean): void => {
    editor.commands.setDocumentParagraphPagination(
      { [key]: checked },
      { restoreFocus: false },
    );
  };
  const clear = (): void => {
    if (!customized) return;
    panelRef.current
      ?.querySelector<HTMLInputElement>('input[type="checkbox"]:not(:disabled)')
      ?.focus({ preventScroll: true });
    editor.commands.clearDocumentParagraphPagination({ restoreFocus: false });
  };

  return (
    <Popover
      label={officeMessage(messages, 'document.pagination.label')}
      panelLabel={officeMessage(messages, 'document.pagination.panelLabel')}
      panelRole="dialog"
      portal
      panelRef={panelRef}
      focusFirstOnOpen
      className="work-document-pagination-popover"
      panelClassName="work-document-pagination-panel"
      trigger={(triggerProps, { open }) => (
        <button
          {...triggerProps}
          className={`with-label${customized || open ? ' active' : ''}`}
          title={officeMessage(
            messages,
            customized
              ? 'document.pagination.titleCustom'
              : 'document.pagination.label',
          )}
        >
          <Pilcrow size={19} />
          <span>{officeMessage(messages, 'document.pagination.label')}</span>
        </button>
      )}
    >
      <fieldset>
        <legend>
          {officeMessage(messages, 'document.pagination.legend')}
        </legend>
        <OfficeCheckbox
          ariaLabel={officeMessage(messages, 'document.pagination.keepLines')}
          checked={pagination.keepLines}
          onCheckedChange={(checked) => update('keepLines', checked)}
        >
          {officeMessage(messages, 'document.pagination.keepLines')}
        </OfficeCheckbox>
        <OfficeCheckbox
          ariaLabel={officeMessage(
            messages,
            'document.pagination.keepWithNext',
          )}
          checked={pagination.keepWithNext}
          onCheckedChange={(checked) => update('keepWithNext', checked)}
        >
          {officeMessage(messages, 'document.pagination.keepWithNext')}
        </OfficeCheckbox>
        <OfficeCheckbox
          ariaLabel={officeMessage(
            messages,
            'document.pagination.pageBreakBefore',
          )}
          checked={pagination.pageBreakBefore}
          onCheckedChange={(checked) => update('pageBreakBefore', checked)}
        >
          {officeMessage(messages, 'document.pagination.pageBreakBefore')}
        </OfficeCheckbox>
        <OfficeCheckbox
          ariaLabel={officeMessage(
            messages,
            'document.pagination.widowControl',
          )}
          checked={pagination.widowControl}
          onCheckedChange={(checked) => update('widowControl', checked)}
        >
          {officeMessage(messages, 'document.pagination.widowControl')}
        </OfficeCheckbox>
        <button
          type="button"
          className="work-document-pagination-reset"
          aria-label={officeMessage(messages, 'document.pagination.resetAria')}
          disabled={!customized}
          onClick={clear}
        >
          {officeMessage(messages, 'document.pagination.reset')}
        </button>
      </fieldset>
    </Popover>
  );
}

function documentPaginationSnapshot(editor: Editor): string {
  const attributes = editor.isActive('heading')
    ? editor.getAttributes('heading')
    : editor.getAttributes('paragraph');
  const pagination = documentParagraphPagination(editor);
  return JSON.stringify({
    node: editor.isActive('heading') ? 'heading' : 'paragraph',
    pagination,
    direct: {
      keepLines: attributes.keepLines ?? null,
      keepWithNext: attributes.keepWithNext ?? null,
      pageBreakBefore: attributes.pageBreakBefore ?? null,
      widowControl: attributes.widowControl ?? null,
    },
  });
}
