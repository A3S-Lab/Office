import type { Editor } from '@tiptap/core';
import { Pilcrow } from 'lucide-react';
import { useCallback, useRef, useSyncExternalStore } from 'react';
import { Popover } from '../../../design-system/primitives';
import { documentParagraphPagination } from '../work-document-paragraph-formatting';
import { OfficeCheckbox } from './office-controls';

export function DocumentPaginationPopover({ editor }: { editor: Editor }) {
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
      label="段落分页"
      panelLabel="段落分页选项"
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
          aria-pressed={customized}
        >
          <Pilcrow size={19} />
          <span>段落分页</span>
        </button>
      )}
    >
      <fieldset>
        <legend>段落分页</legend>
        <OfficeCheckbox
          ariaLabel="段落不跨页"
          checked={pagination.keepLines}
          onCheckedChange={(checked) => update('keepLines', checked)}
        >
          段落不跨页
        </OfficeCheckbox>
        <OfficeCheckbox
          ariaLabel="与下一段同页"
          checked={pagination.keepWithNext}
          onCheckedChange={(checked) => update('keepWithNext', checked)}
        >
          与下一段同页
        </OfficeCheckbox>
        <OfficeCheckbox
          ariaLabel="段前另起一页"
          checked={pagination.pageBreakBefore}
          onCheckedChange={(checked) => update('pageBreakBefore', checked)}
        >
          段前另起一页
        </OfficeCheckbox>
        <OfficeCheckbox
          ariaLabel="避免页首、页尾单行"
          checked={pagination.widowControl}
          onCheckedChange={(checked) => update('widowControl', checked)}
        >
          避免页首、页尾单行
        </OfficeCheckbox>
        <button
          type="button"
          className="work-document-pagination-reset"
          aria-label="恢复默认分页规则"
          disabled={!customized}
          onClick={clear}
        >
          恢复默认
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
