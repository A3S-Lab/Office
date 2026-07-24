import type { Editor } from '@tiptap/core';
import { Pilcrow } from 'lucide-react';
import { Popover } from '../../../design-system/primitives';
import {
  documentParagraphPagination,
  setDocumentParagraphPagination,
} from '../work-document-paragraph-formatting';
import { OfficeCheckbox } from './office-controls';

export function DocumentPaginationPopover({ editor }: { editor: Editor }) {
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
    const current = documentParagraphPagination(editor);
    setDocumentParagraphPagination(
      editor,
      { ...current, [key]: checked },
      { restoreFocus: false },
    );
  };

  return (
    <Popover
      label="段落分页"
      panelLabel="段落分页选项"
      panelRole="dialog"
      portal
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
      </fieldset>
    </Popover>
  );
}
