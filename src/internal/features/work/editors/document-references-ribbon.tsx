import type { Editor } from '@tiptap/core';
import {
  BookOpen,
  Image as ImageIcon,
  Link2,
  ListOrdered,
  ListTree,
  Lock,
  LockOpen,
  RefreshCw,
  Table2,
  Tags,
  Unlink2,
} from 'lucide-react';
import type { ReactNode } from 'react';
import { officeMessage } from '../../../i18n/office-locale';
import type { WorkDocumentCaptionKind } from '../work-document-captions';
import { documentFieldLockTargets } from '../work-document-field-node';
import {
  documentHasIndex,
  selectedDocumentIndexDraft,
} from '../work-document-index-nodes';
import type { WorkDocumentNoteKind } from '../work-document-notes';
import { documentHasTableOfContents } from '../work-document-table-of-contents-node';
import { getDocumentCommandDefinition } from './document-command-catalog';
import { documentCommandLabel } from './document-command-i18n';
import { documentHasRefreshableFields } from './document-editor-support';
import { selectedDocumentFieldIds } from './document-field-code-overrides';
import { useOfficeMessages } from './office-messages-context';
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
  onUnlinkFields: () => void;
  onLockFields: () => void;
  onUnlockFields: () => void;
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
  onUnlinkFields,
  onLockFields,
  onUnlockFields,
}: DocumentReferencesRibbonProps) {
  const messages = useOfficeMessages();
  const hasTableOfContents = documentHasTableOfContents(editor);
  const hasIndex = documentHasIndex(editor);
  const canMarkIndexEntry = Boolean(selectedDocumentIndexDraft(editor));
  const hasRefreshableFields = documentHasRefreshableFields(editor);
  const fieldLockTargets = documentFieldLockTargets(editor.state);
  const canUnlinkFields = selectedDocumentFieldIds(editor).length > 0;
  const canLockFields = fieldLockTargets.some((field) => !field.locked);
  const canUnlockFields = fieldLockTargets.some((field) => field.locked);
  const refreshFieldsCommand = getDocumentCommandDefinition('refreshFields');
  const unlinkFieldsCommand = getDocumentCommandDefinition('unlinkFields');
  const lockFieldsCommand = getDocumentCommandDefinition('lockFields');
  const unlockFieldsCommand = getDocumentCommandDefinition('unlockFields');

  return (
    <>
      <WorkOfficeRibbonGroup
        label={officeMessage(messages, 'document.references.group.toc')}
        priority="high"
      >
        <ReferencesButton
          label={documentCommandLabel('tableOfContents', messages)}
          onClick={onOpenTableOfContents}
        >
          <ListTree size={19} />
        </ReferencesButton>
        <ReferencesButton
          label={documentCommandLabel('refreshTableOfContents', messages)}
          disabled={!hasTableOfContents}
          title={officeMessage(
            messages,
            hasTableOfContents
              ? 'document.references.refreshToc.enabled'
              : 'document.references.refreshToc.disabled',
          )}
          onClick={onRefreshTableOfContents}
        >
          <RefreshCw size={19} />
        </ReferencesButton>
      </WorkOfficeRibbonGroup>
      <WorkOfficeRibbonGroup
        label={officeMessage(messages, 'document.references.group.notes')}
        priority="high"
      >
        <ReferencesButton
          label={officeMessage(
            messages,
            'document.references.insertFootnote',
          )}
          onClick={() => onInsertNote('footnote')}
        >
          <span className="work-ribbon-glyph">¹</span>
        </ReferencesButton>
        <ReferencesButton
          label={officeMessage(messages, 'document.references.insertEndnote')}
          onClick={() => onInsertNote('endnote')}
        >
          <span className="work-ribbon-glyph">ⅰ</span>
        </ReferencesButton>
      </WorkOfficeRibbonGroup>
      <WorkOfficeRibbonGroup
        label={officeMessage(messages, 'document.references.group.captions')}
      >
        <ReferencesButton
          label={officeMessage(
            messages,
            'document.references.insertFigureCaption',
          )}
          onClick={() => onInsertCaption('figure')}
        >
          <ImageIcon size={19} />
        </ReferencesButton>
        <ReferencesButton
          label={officeMessage(
            messages,
            'document.references.insertTableCaption',
          )}
          onClick={() => onInsertCaption('table')}
        >
          <Table2 size={19} />
        </ReferencesButton>
        <ReferencesButton
          label={officeMessage(
            messages,
            'document.references.insertCrossReference',
          )}
          onClick={onInsertCrossReference}
        >
          <Link2 size={19} />
        </ReferencesButton>
      </WorkOfficeRibbonGroup>
      <WorkOfficeRibbonGroup
        label={officeMessage(messages, 'document.references.group.citations')}
        priority="high"
      >
        <ReferencesButton
          label={
            citationSourceCount
              ? officeMessage(
                  messages,
                  'document.references.bibliographyWithCount',
                  { count: String(citationSourceCount) },
                )
              : officeMessage(messages, 'document.references.bibliography')
          }
          visibleLabel={officeMessage(
            messages,
            'document.references.bibliography',
          )}
          active={citationsOpen}
          onClick={onToggleCitations}
        >
          <BookOpen size={19} />
        </ReferencesButton>
      </WorkOfficeRibbonGroup>
      <WorkOfficeRibbonGroup
        label={officeMessage(messages, 'document.references.group.index')}
        priority="high"
      >
        <ReferencesButton
          label={documentCommandLabel('markIndexEntry', messages)}
          disabled={!canMarkIndexEntry}
          title={officeMessage(
            messages,
            canMarkIndexEntry
              ? 'document.references.markIndex.enabled'
              : 'document.references.markIndex.disabled',
          )}
          onClick={onOpenIndexEntry}
        >
          <Tags size={19} />
        </ReferencesButton>
        <ReferencesButton
          label={documentCommandLabel('index', messages)}
          onClick={onOpenIndex}
        >
          <ListOrdered size={19} />
        </ReferencesButton>
        <ReferencesButton
          label={documentCommandLabel('refreshIndex', messages)}
          disabled={!hasIndex}
          title={officeMessage(
            messages,
            hasIndex
              ? 'document.references.refreshIndex.enabled'
              : 'document.references.refreshIndex.disabled',
          )}
          onClick={onRefreshIndex}
        >
          <RefreshCw size={19} />
        </ReferencesButton>
      </WorkOfficeRibbonGroup>
      <WorkOfficeRibbonGroup
        label={officeMessage(messages, 'document.references.group.update')}
        priority="low"
      >
        <ReferencesButton
          label={documentCommandLabel('refreshFields', messages)}
          shortcut={refreshFieldsCommand.shortcut?.label}
          ariaKeyShortcuts={refreshFieldsCommand.shortcut?.aria}
          disabled={!hasRefreshableFields}
          title={
            hasRefreshableFields
              ? officeMessage(messages, 'document.fields.refreshTitle', {
                  shortcut: refreshFieldsCommand.shortcut?.label ?? '',
                })
              : officeMessage(messages, 'document.fields.refreshEmpty')
          }
          onClick={onRefreshFields}
        >
          <RefreshCw size={19} />
        </ReferencesButton>
        <ReferencesButton
          label={documentCommandLabel('unlinkFields', messages)}
          shortcut={unlinkFieldsCommand.shortcut?.label}
          ariaKeyShortcuts={unlinkFieldsCommand.shortcut?.aria}
          disabled={!canUnlinkFields}
          title={
            canUnlinkFields
              ? officeMessage(messages, 'document.fields.unlinkTitle', {
                  shortcut: unlinkFieldsCommand.shortcut?.label ?? '',
                })
              : officeMessage(messages, 'document.fields.unlinkEmpty')
          }
          onClick={onUnlinkFields}
        >
          <Unlink2 size={19} />
        </ReferencesButton>
        <ReferencesButton
          label={documentCommandLabel('lockFields', messages)}
          shortcut={lockFieldsCommand.shortcut?.label}
          ariaKeyShortcuts={lockFieldsCommand.shortcut?.aria}
          disabled={!canLockFields}
          title={
            canLockFields
              ? officeMessage(messages, 'document.fields.lockTitle', {
                  shortcut: lockFieldsCommand.shortcut?.label ?? '',
                })
              : officeMessage(messages, 'document.fields.lockEmpty')
          }
          onClick={onLockFields}
        >
          <Lock size={19} />
        </ReferencesButton>
        <ReferencesButton
          label={documentCommandLabel('unlockFields', messages)}
          shortcut={unlockFieldsCommand.shortcut?.label}
          ariaKeyShortcuts={unlockFieldsCommand.shortcut?.aria}
          disabled={!canUnlockFields}
          title={
            canUnlockFields
              ? officeMessage(messages, 'document.fields.unlockTitle', {
                  shortcut: unlockFieldsCommand.shortcut?.label ?? '',
                })
              : officeMessage(messages, 'document.fields.unlockEmpty')
          }
          onClick={onUnlockFields}
        >
          <LockOpen size={19} />
        </ReferencesButton>
      </WorkOfficeRibbonGroup>
    </>
  );
}

function ReferencesButton({
  label,
  visibleLabel,
  title,
  shortcut,
  ariaKeyShortcuts,
  active = false,
  disabled = false,
  onClick,
  children,
}: {
  label: string;
  visibleLabel?: string;
  title?: string;
  shortcut?: string;
  ariaKeyShortcuts?: string;
  active?: boolean;
  disabled?: boolean;
  onClick: () => void;
  children: ReactNode;
}) {
  const messages = useOfficeMessages();
  return (
    <WorkOfficeRibbonButton
      label={label}
      visibleLabel={visibleLabel ?? label}
      title={
        title ??
        (shortcut
          ? officeMessage(messages, 'document.ribbon.shortcutTitle', {
              label,
              shortcut,
            })
          : label)
      }
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
