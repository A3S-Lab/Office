import type { Editor } from '@tiptap/core';
import {
  BookOpen,
  Image as ImageIcon,
  Link2,
  ListOrdered,
  ListTree,
  RefreshCw,
  Table2,
  Tags,
} from 'lucide-react';
import type { ReactNode } from 'react';
import type { WorkDocumentCaptionKind } from '../work-document-captions';
import type { WorkDocumentNoteKind } from '../work-document-notes';
import {
  documentHasIndex,
  selectedDocumentIndexDraft,
} from '../work-document-index-nodes';
import { documentHasTableOfContents } from '../work-document-table-of-contents-node';
import { getDocumentCommandDefinition } from './document-command-catalog';
import { documentHasRefreshableFields } from './document-editor-support';
import {
  WorkOfficeRibbonButton,
  WorkOfficeRibbonGroup,
} from './work-office-chrome';

export interface DocumentReferencesRibbonProps {
  editor: Editor;
  citationsOpen: boolean;
  citationSourceCount: number;
  onInsertNote: (kind: WorkDocumentNoteKind) => void;
  onInsertCaption: (kind: WorkDocumentCaptionKind) => void;
  onInsertCrossReference: () => void;
  onOpenTableOfContents: () => void;
  onRefreshTableOfContents: () => void;
  onOpenIndexEntry: () => void;
  onOpenIndex: () => void;
  onRefreshIndex: () => void;
  onToggleCitations: () => void;
  onRefreshFields: () => void;
}

export function DocumentReferencesRibbon({
  editor,
  citationsOpen,
  citationSourceCount,
  onInsertNote,
  onInsertCaption,
  onInsertCrossReference,
  onOpenTableOfContents,
  onRefreshTableOfContents,
  onOpenIndexEntry,
  onOpenIndex,
  onRefreshIndex,
  onToggleCitations,
  onRefreshFields,
}: DocumentReferencesRibbonProps) {
  const hasTableOfContents = documentHasTableOfContents(editor);
  const hasIndex = documentHasIndex(editor);
  const canMarkIndexEntry = Boolean(selectedDocumentIndexDraft(editor));
  const hasRefreshableFields = documentHasRefreshableFields(editor);
  const refreshFieldsCommand = getDocumentCommandDefinition('refreshFields');

  return (
    <>
      <WorkOfficeRibbonGroup label="目录" priority="high">
        <ReferencesButton
          label="插入或自定义目录"
          onClick={onOpenTableOfContents}
        >
          <ListTree size={19} />
        </ReferencesButton>
        <ReferencesButton
          label="更新目录"
          disabled={!hasTableOfContents}
          title={
            hasTableOfContents
              ? '根据当前标题和页码更新目录'
              : '文档中没有可更新的目录'
          }
          onClick={onRefreshTableOfContents}
        >
          <RefreshCw size={19} />
        </ReferencesButton>
      </WorkOfficeRibbonGroup>
      <WorkOfficeRibbonGroup label="脚注" priority="high">
        <ReferencesButton
          label="插入脚注"
          onClick={() => onInsertNote('footnote')}
        >
          <span className="work-ribbon-glyph">¹</span>
        </ReferencesButton>
        <ReferencesButton
          label="插入尾注"
          onClick={() => onInsertNote('endnote')}
        >
          <span className="work-ribbon-glyph">ⅰ</span>
        </ReferencesButton>
      </WorkOfficeRibbonGroup>
      <WorkOfficeRibbonGroup label="题注">
        <ReferencesButton
          label="插入图片题注"
          onClick={() => onInsertCaption('figure')}
        >
          <ImageIcon size={19} />
        </ReferencesButton>
        <ReferencesButton
          label="插入表格题注"
          onClick={() => onInsertCaption('table')}
        >
          <Table2 size={19} />
        </ReferencesButton>
        <ReferencesButton label="插入交叉引用" onClick={onInsertCrossReference}>
          <Link2 size={19} />
        </ReferencesButton>
      </WorkOfficeRibbonGroup>
      <WorkOfficeRibbonGroup label="引文和书目" priority="high">
        <ReferencesButton
          label={`文献库${citationSourceCount ? `（${citationSourceCount}）` : ''}`}
          active={citationsOpen}
          onClick={onToggleCitations}
        >
          <BookOpen size={19} />
        </ReferencesButton>
      </WorkOfficeRibbonGroup>
      <WorkOfficeRibbonGroup label="索引" priority="high">
        <ReferencesButton
          label="标记索引项"
          disabled={!canMarkIndexEntry}
          title={
            canMarkIndexEntry
              ? '把当前选中文字标记为索引项'
              : '请先选择正文文字或已有索引项'
          }
          onClick={onOpenIndexEntry}
        >
          <Tags size={19} />
        </ReferencesButton>
        <ReferencesButton label="插入或自定义索引" onClick={onOpenIndex}>
          <ListOrdered size={19} />
        </ReferencesButton>
        <ReferencesButton
          label="更新索引"
          disabled={!hasIndex}
          title={
            hasIndex ? '根据当前索引项和页码更新索引' : '文档中没有可更新的索引'
          }
          onClick={onRefreshIndex}
        >
          <RefreshCw size={19} />
        </ReferencesButton>
      </WorkOfficeRibbonGroup>
      <WorkOfficeRibbonGroup label="更新" priority="low">
        <ReferencesButton
          label="更新页码和日期"
          shortcut={refreshFieldsCommand.shortcut?.label}
          ariaKeyShortcuts={refreshFieldsCommand.shortcut?.aria}
          disabled={!hasRefreshableFields}
          title={
            hasRefreshableFields
              ? `更新页码和日期（${refreshFieldsCommand.shortcut?.label}）`
              : '文档中没有可更新的页码或日期'
          }
          onClick={onRefreshFields}
        >
          <RefreshCw size={19} />
        </ReferencesButton>
      </WorkOfficeRibbonGroup>
    </>
  );
}

function ReferencesButton({
  label,
  title,
  shortcut,
  ariaKeyShortcuts,
  active = false,
  disabled = false,
  onClick,
  children,
}: {
  label: string;
  title?: string;
  shortcut?: string;
  ariaKeyShortcuts?: string;
  active?: boolean;
  disabled?: boolean;
  onClick: () => void;
  children: ReactNode;
}) {
  return (
    <WorkOfficeRibbonButton
      label={label}
      visibleLabel={label.replace(/（\d+）$/, '')}
      title={title ?? (shortcut ? `${label}（${shortcut}）` : label)}
      aria-keyshortcuts={ariaKeyShortcuts}
      active={active}
      displayLabel
      disabled={disabled}
      onClick={onClick}
    >
      {children}
    </WorkOfficeRibbonButton>
  );
}
