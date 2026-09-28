import type { Editor } from '@tiptap/core';
import { NodeSelection } from '@tiptap/pm/state';
import {
  Fragment,
  type KeyboardEvent,
  type MutableRefObject,
  type ReactNode,
  useCallback,
  useId,
  useRef,
  useState,
} from 'react';
import { Button, Dialog } from '../../../design-system/primitives';
import type { WorkDocumentBookmarkReferenceTarget } from '../work-document-bookmark-references';
import { editorDocumentBookmarkReferenceTargets } from '../work-document-bookmarks';
import { editorDocumentCaptionTargets } from '../work-document-caption-nodes';
import type {
  WorkDocumentCaptionKind,
  WorkDocumentCaptionTarget,
} from '../work-document-captions';
import {
  DOCUMENT_CONTENT_CONTROL_DEFAULTS,
  documentContentControlProperties,
  normalizeDocumentContentControlProperties,
  type WorkDocumentContentControlProperties,
} from '../work-document-content-control';
import type {
  WorkDocumentFieldContextResolver,
  WorkDocumentFieldDraft,
  WorkDocumentFieldKind,
} from '../work-document-fields';
import {
  documentFieldDisplay,
  documentFieldDraftFromAttributes,
  documentFieldInstruction,
  documentFieldOptionsFromDraft,
} from '../work-document-fields';
import {
  DEFAULT_DOCUMENT_INDEX_OPTIONS,
  type WorkDocumentIndexEntryDraft,
  type WorkDocumentIndexOptions,
} from '../work-document-index';
import {
  selectedDocumentIndexDraft,
  selectedDocumentIndexEntry,
  selectedDocumentIndexOptions,
} from '../work-document-index-nodes';
import type { WorkDocumentNoteKind } from '../work-document-notes';
import {
  DEFAULT_DOCUMENT_TABLE_OF_CONTENTS_OPTIONS,
  type WorkDocumentTableOfContentsOptions,
} from '../work-document-table-of-contents';
import { selectedDocumentTableOfContentsOptions } from '../work-document-table-of-contents-node';
import type { WorkDocumentContent } from '../work-types';
import { DocumentContentControlDialog } from './document-content-control-dialog';
import {
  DocumentFieldDialog,
  type DocumentFieldTargetOption,
} from './document-field-dialog';
import { DocumentIndexDialog } from './document-index-dialog';
import { DocumentIndexEntryDialog } from './document-index-entry-dialog';
import { DocumentTableOfContentsDialog } from './document-table-of-contents-dialog';
import { OfficeTextField, useOfficeDialog } from './office-controls';
import { readOfficeFileAsDataUrl } from './office-file-data';
import { officeMessage, resolveOfficeMessages } from '../../../i18n/office-locale';
import { useOfficeMessages } from './office-messages-context';

type DocumentInsertDialog =
  | {
      kind: 'caption';
      captionKind: WorkDocumentCaptionKind;
      title: string;
    }
  | {
      kind: 'crossReference';
      selectedKey: string;
      targets: WorkDocumentCrossReferenceTarget[];
    }
  | {
      kind: 'tableOfContents';
      editing: boolean;
      options: WorkDocumentTableOfContentsOptions;
    }
  | {
      kind: 'indexEntry';
      editing: boolean;
      value: WorkDocumentIndexEntryDraft;
    }
  | {
      kind: 'index';
      editing: boolean;
      options: WorkDocumentIndexOptions;
    }
  | {
      kind: 'contentControl';
      editing: boolean;
      properties: WorkDocumentContentControlProperties;
    }
  | {
      kind: 'field';
      editing: boolean;
      draft: WorkDocumentFieldDraft;
      targets: DocumentFieldTargetOption[];
    };

type WorkDocumentCrossReferenceTarget =
  | (WorkDocumentCaptionTarget & { type: 'caption' })
  | WorkDocumentBookmarkReferenceTarget;

export interface DocumentInsertCommands {
  dialog: ReactNode;
  insertCaption: (kind: WorkDocumentCaptionKind) => void;
  insertCrossReference: () => void;
  insertField: (kind: WorkDocumentFieldKind) => void;
  openField: () => void;
  insertImage: (file: File) => Promise<void>;
  insertNote: (kind: WorkDocumentNoteKind) => boolean;
  insertTextBox: () => boolean;
  insertConnector: () => boolean;
  openContentControl: () => void;
  openIndexEntry: () => void;
  openIndex: () => void;
  openTableOfContents: () => void;
  refreshFields: () => boolean;
  refreshIndex: () => boolean;
  refreshTableOfContents: () => boolean;
}

export function useDocumentInsertCommands({
  contentRef,
  editor,
  resolveFieldContext,
}: {
  contentRef: MutableRefObject<WorkDocumentContent>;
  editor: Editor | null;
  resolveFieldContext?: WorkDocumentFieldContextResolver | null;
}): DocumentInsertCommands {
  const messages = useOfficeMessages();
  const officeDialog = useOfficeDialog();
  const [insertDialog, setInsertDialog] = useState<DocumentInsertDialog | null>(
    null,
  );
  const invokerRef = useRef<HTMLElement | null>(null);
  const captionFieldId = useId();
  const referenceInputRefs = useRef(new Map<string, HTMLInputElement>());

  const rememberInvoker = useCallback(() => {
    invokerRef.current =
      document.activeElement instanceof HTMLElement
        ? document.activeElement
        : null;
  }, []);

  const insertImage = useCallback(
    async (file: File) => {
      if (!editor) return;
      if (
        ![
          'image/bmp',
          'image/gif',
          'image/jpeg',
          'image/png',
          'image/webp',
        ].includes(file.type)
      ) {
        await officeDialog.notice({
          title: officeMessage(messages, 'document.insert.imageUnsupportedTitle'),
          description: officeMessage(messages, 'document.insert.imageUnsupportedDesc'),
        });
        return;
      }
      if (file.size > 8 * 1024 * 1024) {
        await officeDialog.notice({
          title: officeMessage(messages, 'document.insert.imageTooLargeTitle'),
          description: officeMessage(messages, 'document.insert.imageTooLargeDesc'),
        });
        return;
      }
      try {
        const src = await readOfficeFileAsDataUrl(file);
        if (!editor.isDestroyed) {
          editor
            .chain()
            .focus()
            .setImage({ src, alt: file.name, title: file.name })
            .run();
        }
      } catch {
        await officeDialog.notice({
          title: officeMessage(messages, 'document.insert.imageReadFailedTitle'),
          description: officeMessage(messages, 'document.insert.imageReadFailedDesc'),
        });
      }
    },
    [editor, officeDialog],
  );

  const insertNote = useCallback(
    (kind: WorkDocumentNoteKind) =>
      editor?.chain().focus().insertDocumentNote(kind).run() ?? false,
    [editor],
  );

  const insertTextBox = useCallback(() => {
    if (!editor || editor.isDestroyed) return false;
    return editor.chain().focus().insertDocumentTextBox().run();
  }, [editor]);

  const insertConnector = useCallback(() => {
    if (!editor || editor.isDestroyed) return false;
    return editor.chain().focus().insertDocumentConnector().run();
  }, [editor]);

  const openContentControl = useCallback(() => {
    if (!editor || editor.isDestroyed) return;
    rememberInvoker();
    const active = editor.isActive('documentContentControl');
    setInsertDialog({
      kind: 'contentControl',
      editing: active,
      properties: active
        ? documentContentControlProperties(editor)
        : normalizeDocumentContentControlProperties(
            DOCUMENT_CONTENT_CONTROL_DEFAULTS,
          ),
    });
  }, [editor, rememberInvoker]);

  const insertCaption = useCallback(
    (kind: WorkDocumentCaptionKind) => {
      if (!editor) return;
      rememberInvoker();
      setInsertDialog({ kind: 'caption', captionKind: kind, title: '' });
    },
    [editor, rememberInvoker],
  );

  const insertCrossReference = useCallback(() => {
    if (!editor) return;
    const targets: WorkDocumentCrossReferenceTarget[] = [
      ...editorDocumentCaptionTargets(editor).map((target) => ({
        ...target,
        type: 'caption' as const,
      })),
      ...editorDocumentBookmarkReferenceTargets(editor),
    ];
    if (!targets.length) {
      void officeDialog.notice({
        title: officeMessage(messages, 'document.insert.noReferenceTargetsTitle'),
        description: officeMessage(messages, 'document.insert.noReferenceTargetsDesc'),
      });
      return;
    }
    rememberInvoker();
    setInsertDialog({
      kind: 'crossReference',
      selectedKey: targets[0] ? referenceTargetKey(targets[0]) : '',
      targets,
    });
  }, [editor, officeDialog, rememberInvoker]);

  const insertField = useCallback(
    (kind: WorkDocumentFieldKind) => {
      if (!editor) return;
      if (!editor.chain().focus().insertDocumentField(kind).run()) return;
      editor.commands.refreshDocumentFields(contentRef.current, {
        resolveContext: resolveFieldContext ?? undefined,
        addToHistory: false,
      });
    },
    [contentRef, editor, resolveFieldContext],
  );

  const openField = useCallback(() => {
    if (!editor) return;
    const selectedNode =
      editor.state.selection instanceof NodeSelection
        ? editor.state.selection.node
        : null;
    const selectedDraft =
      selectedNode?.type.name === 'documentField'
        ? documentFieldDraftFromAttributes(selectedNode.attrs)
        : null;
    const draft = selectedDraft ?? defaultDocumentFieldDraft();
    const targets = editorDocumentBookmarkReferenceTargets(editor).map(
      ({ id, name }) => ({ id, name }),
    );
    if (
      draft.kind === 'pageReference' &&
      draft.targetName &&
      !targets.some((target) => target.id === draft.targetId)
    ) {
      targets.unshift({ id: draft.targetId, name: draft.targetName });
    }
    rememberInvoker();
    setInsertDialog({
      kind: 'field',
      editing: Boolean(selectedDraft),
      draft,
      targets,
    });
  }, [editor, rememberInvoker]);

  const openTableOfContents = useCallback(() => {
    if (!editor) return;
    rememberInvoker();
    const selected = selectedDocumentTableOfContentsOptions(editor);
    setInsertDialog({
      kind: 'tableOfContents',
      editing: Boolean(selected),
      options: selected ?? DEFAULT_DOCUMENT_TABLE_OF_CONTENTS_OPTIONS,
    });
  }, [editor, rememberInvoker]);

  const openIndexEntry = useCallback(() => {
    if (!editor) return;
    const value = selectedDocumentIndexDraft(editor);
    if (!value) {
      void officeDialog.notice({
        title: officeMessage(messages, 'document.insert.selectIndexTextTitle'),
        description: officeMessage(messages, 'document.insert.selectIndexTextDesc'),
      });
      return;
    }
    rememberInvoker();
    setInsertDialog({
      kind: 'indexEntry',
      editing: Boolean(selectedDocumentIndexEntry(editor)),
      value,
    });
  }, [editor, officeDialog, rememberInvoker]);

  const openIndex = useCallback(() => {
    if (!editor) return;
    rememberInvoker();
    const selected = selectedDocumentIndexOptions(editor);
    setInsertDialog({
      kind: 'index',
      editing: Boolean(selected),
      options: selected ?? DEFAULT_DOCUMENT_INDEX_OPTIONS,
    });
  }, [editor, rememberInvoker]);

  const refreshFields = useCallback(
    () =>
      editor?.commands.refreshDocumentFields(contentRef.current, {
        resolveContext: resolveFieldContext ?? undefined,
      }) ?? false,
    [contentRef, editor, resolveFieldContext],
  );
  const refreshTableOfContents = useCallback(
    () =>
      editor?.commands.refreshDocumentTablesOfContents({
        resolveContext: resolveFieldContext ?? undefined,
      }) ?? false,
    [editor, resolveFieldContext],
  );
  const refreshIndex = useCallback(
    () =>
      editor?.commands.refreshDocumentIndexes({
        resolveContext: resolveFieldContext ?? undefined,
      }) ?? false,
    [editor, resolveFieldContext],
  );
  const submitCaption = () => {
    if (!editor || insertDialog?.kind !== 'caption') return;
    const inserted = editor
      .chain()
      .focus()
      .insertDocumentCaption(
        insertDialog.captionKind,
        insertDialog.title.trim(),
      )
      .run();
    if (!inserted) return;
    invokerRef.current = editor.view.dom;
    setInsertDialog(null);
  };
  const submitCrossReference = () => {
    if (!editor || insertDialog?.kind !== 'crossReference') return;
    const target = insertDialog.targets.find(
      (candidate) => referenceTargetKey(candidate) === insertDialog.selectedKey,
    );
    if (!target) return;
    const inserted = editor
      .chain()
      .focus()
      .insertDocumentCrossReference(target)
      .run();
    if (!inserted) return;
    invokerRef.current = editor.view.dom;
    setInsertDialog(null);
  };
  const submitPageReference = () => {
    if (!editor || insertDialog?.kind !== 'crossReference') return;
    const target = insertDialog.targets.find(
      (candidate) => referenceTargetKey(candidate) === insertDialog.selectedKey,
    );
    if (!target || target.type !== 'bookmark') return;
    const inserted = editor
      .chain()
      .focus()
      .insertDocumentField('pageReference', {
        targetId: target.id,
        targetName: target.name,
      })
      .run();
    if (!inserted) return;
    editor.commands.refreshDocumentFields(contentRef.current, {
      resolveContext: resolveFieldContext ?? undefined,
      addToHistory: false,
    });
    invokerRef.current = editor.view.dom;
    setInsertDialog(null);
  };
  const submitTableOfContents = () => {
    if (!editor || insertDialog?.kind !== 'tableOfContents') return;
    const buildOptions = {
      resolveContext: resolveFieldContext ?? undefined,
    };
    const applied = insertDialog.editing
      ? editor.commands.updateDocumentTableOfContents(
          insertDialog.options,
          buildOptions,
        )
      : editor.commands.insertDocumentTableOfContents(
          insertDialog.options,
          buildOptions,
        );
    if (!applied) return;
    invokerRef.current = editor.view.dom;
    setInsertDialog(null);
  };
  const submitIndexEntry = () => {
    if (!editor || insertDialog?.kind !== 'indexEntry') return;
    const applied = insertDialog.editing
      ? editor.commands.updateDocumentIndexEntry(insertDialog.value)
      : editor.commands.markDocumentIndexEntry(insertDialog.value);
    if (!applied) return;
    invokerRef.current = editor.view.dom;
    setInsertDialog(null);
  };
  const submitIndex = () => {
    if (!editor || insertDialog?.kind !== 'index') return;
    const buildOptions = {
      resolveContext: resolveFieldContext ?? undefined,
    };
    const applied = insertDialog.editing
      ? editor.commands.updateDocumentIndex(insertDialog.options, buildOptions)
      : editor.commands.insertDocumentIndex(insertDialog.options, buildOptions);
    if (!applied) return;
    invokerRef.current = editor.view.dom;
    setInsertDialog(null);
  };
  const submitContentControl = () => {
    if (!editor || insertDialog?.kind !== 'contentControl') return;
    const applied = insertDialog.editing
      ? editor
          .chain()
          .focus()
          .setDocumentContentControlProperties(insertDialog.properties)
          .run()
      : editor
          .chain()
          .focus()
          .insertDocumentContentControl(insertDialog.properties)
          .run();
    if (!applied) return;
    invokerRef.current = editor.view.dom;
    setInsertDialog(null);
  };
  const submitField = () => {
    if (!editor || insertDialog?.kind !== 'field') return;
    const applied = insertDialog.editing
      ? editor.chain().focus().updateDocumentField(insertDialog.draft).run()
      : editor
          .chain()
          .focus()
          .insertDocumentField(
            insertDialog.draft.kind,
            documentFieldOptionsFromDraft(insertDialog.draft),
          )
          .run();
    if (!applied) return;
    editor.commands.refreshDocumentFields(contentRef.current, {
      resolveContext: resolveFieldContext ?? undefined,
      addToHistory: false,
    });
    invokerRef.current = editor.view.dom;
    setInsertDialog(null);
  };
  const moveReferenceSelection = (
    event: KeyboardEvent<HTMLInputElement>,
    currentIndex: number,
  ) => {
    if (insertDialog?.kind !== 'crossReference') return;
    let nextIndex = currentIndex;
    if (event.key === 'ArrowDown' || event.key === 'ArrowRight') {
      nextIndex = Math.min(insertDialog.targets.length - 1, currentIndex + 1);
    } else if (event.key === 'ArrowUp' || event.key === 'ArrowLeft') {
      nextIndex = Math.max(0, currentIndex - 1);
    } else if (event.key === 'Home') {
      nextIndex = 0;
    } else if (event.key === 'End') {
      nextIndex = insertDialog.targets.length - 1;
    } else {
      return;
    }
    event.preventDefault();
    event.stopPropagation();
    const target = insertDialog.targets[nextIndex];
    if (!target) return;
    const targetKey = referenceTargetKey(target);
    setInsertDialog({ ...insertDialog, selectedKey: targetKey });
    requestAnimationFrame(() =>
      referenceInputRefs.current.get(targetKey)?.focus({ preventScroll: true }),
    );
  };

  const dialog = (
    <Fragment>
      {officeDialog.dialog}
      {insertDialog?.kind === 'field' && (
        <DocumentFieldDialog
          editing={insertDialog.editing}
          draft={insertDialog.draft}
          preview={documentFieldPreview(insertDialog.draft)}
          targets={insertDialog.targets}
          restoreFocusTarget={() => invokerRef.current}
          onCancel={() => setInsertDialog(null)}
          onChange={(draft) => setInsertDialog({ ...insertDialog, draft })}
          onSubmit={submitField}
        />
      )}
      {insertDialog?.kind === 'caption' && (
        <Dialog
          title={
            insertDialog.captionKind === 'figure'
              ? officeMessage(messages, 'document.insert.captionFigureTitle')
              : officeMessage(messages, 'document.insert.captionTableTitle')
          }
          description={officeMessage(messages, 'document.insert.captionDesc')}
          className="work-document-insert-dialog"
          restoreFocusTarget={() => invokerRef.current}
          onClose={() => setInsertDialog(null)}
          footer={
            <>
              <Button tone="quiet" onClick={() => setInsertDialog(null)}>
                {officeMessage(messages, 'office.dialog.cancel')}
              </Button>
              <Button tone="primary" onClick={submitCaption}>
                {officeMessage(messages, 'document.insert.captionSubmit')}
              </Button>
            </>
          }
        >
          <label
            className="work-document-dialog-field"
            htmlFor={captionFieldId}
          >
            <span>{officeMessage(messages, 'document.insert.captionLabel')}</span>
            <OfficeTextField
              id={captionFieldId}
              data-autofocus
              aria-label={officeMessage(messages, 'document.insert.captionLabel')}
              value={insertDialog.title}
              maxLength={240}
              placeholder={officeMessage(messages, 'document.insert.captionPlaceholder')}
              onChange={(event) =>
                setInsertDialog({
                  ...insertDialog,
                  title: event.target.value,
                })
              }
              onKeyDown={(event) => {
                if (event.key !== 'Enter') return;
                event.preventDefault();
                submitCaption();
              }}
            />
          </label>
        </Dialog>
      )}
      {insertDialog?.kind === 'crossReference' && (
        <Dialog
          title={officeMessage(messages, 'document.insert.crossRefTitle')}
          description={officeMessage(messages, 'document.insert.crossRefDesc')}
          className="work-document-reference-dialog"
          restoreFocusTarget={() => invokerRef.current}
          onClose={() => setInsertDialog(null)}
          footer={
            <>
              <Button tone="quiet" onClick={() => setInsertDialog(null)}>
                {officeMessage(messages, 'office.dialog.cancel')}
              </Button>
              <Button
                tone="secondary"
                disabled={
                  !insertDialog.selectedKey ||
                  insertDialog.targets.find(
                    (candidate) =>
                      referenceTargetKey(candidate) ===
                      insertDialog.selectedKey,
                  )?.type !== 'bookmark'
                }
                onClick={submitPageReference}
              >
                {officeMessage(messages, 'document.insert.crossRefInsertPage')}
              </Button>
              <Button
                tone="primary"
                disabled={!insertDialog.selectedKey}
                onClick={submitCrossReference}
              >
                {officeMessage(messages, 'document.insert.crossRefInsert')}
              </Button>
            </>
          }
        >
          <div
            className="work-document-reference-list"
            role="radiogroup"
            aria-label={officeMessage(messages, 'document.insert.crossRefTargetsAria')}
          >
            {insertDialog.targets.map((target, index) => {
              const targetKey = referenceTargetKey(target);
              return (
                <label key={targetKey} onDoubleClick={submitCrossReference}>
                  <input
                    ref={(element) => {
                      if (element)
                        referenceInputRefs.current.set(targetKey, element);
                      else referenceInputRefs.current.delete(targetKey);
                    }}
                    type="radio"
                    name="document-cross-reference"
                    value={targetKey}
                    checked={targetKey === insertDialog.selectedKey}
                    data-autofocus={
                      targetKey === insertDialog.selectedKey ? '' : undefined
                    }
                    tabIndex={targetKey === insertDialog.selectedKey ? 0 : -1}
                    onChange={() =>
                      setInsertDialog({
                        ...insertDialog,
                        selectedKey: targetKey,
                      })
                    }
                    onKeyDown={(event) => moveReferenceSelection(event, index)}
                  />
                  <strong>
                    {target.type === 'bookmark'
                      ? officeMessage(messages, 'document.insert.bookmarkLabel', {
                          name: target.name,
                        })
                      : target.display}
                  </strong>
                  <span>
                    {target.title ||
                      (target.type === 'bookmark'
                        ? officeMessage(messages, 'document.insert.emptyBookmark')
                        : officeMessage(messages, 'document.insert.emptyCaption'))}
                  </span>
                </label>
              );
            })}
          </div>
        </Dialog>
      )}
      {insertDialog?.kind === 'tableOfContents' && (
        <DocumentTableOfContentsDialog
          editing={insertDialog.editing}
          options={insertDialog.options}
          restoreFocusTarget={() => invokerRef.current}
          onCancel={() => setInsertDialog(null)}
          onOptionsChange={(options) =>
            setInsertDialog({ ...insertDialog, options })
          }
          onSubmit={submitTableOfContents}
        />
      )}
      {insertDialog?.kind === 'indexEntry' && (
        <DocumentIndexEntryDialog
          editing={insertDialog.editing}
          value={insertDialog.value}
          restoreFocusTarget={() => invokerRef.current}
          onCancel={() => setInsertDialog(null)}
          onChange={(value) => setInsertDialog({ ...insertDialog, value })}
          onSubmit={submitIndexEntry}
        />
      )}
      {insertDialog?.kind === 'index' && (
        <DocumentIndexDialog
          editing={insertDialog.editing}
          options={insertDialog.options}
          restoreFocusTarget={() => invokerRef.current}
          onCancel={() => setInsertDialog(null)}
          onOptionsChange={(options) =>
            setInsertDialog({ ...insertDialog, options })
          }
          onSubmit={submitIndex}
        />
      )}
      {insertDialog?.kind === 'contentControl' && (
        <DocumentContentControlDialog
          editing={insertDialog.editing}
          properties={insertDialog.properties}
          restoreFocusTarget={() => invokerRef.current}
          onCancel={() => setInsertDialog(null)}
          onChange={(properties) =>
            setInsertDialog({ ...insertDialog, properties })
          }
          onSubmit={submitContentControl}
        />
      )}
    </Fragment>
  );

  return {
    dialog,
    insertCaption,
    insertCrossReference,
    insertField,
    openField,
    insertImage,
    insertNote,
    insertTextBox,
    insertConnector,
    openContentControl,
    openIndexEntry,
    openIndex,
    openTableOfContents,
    refreshFields,
    refreshIndex,
    refreshTableOfContents,
  };
}

function referenceTargetKey(target: WorkDocumentCrossReferenceTarget): string {
  return `${target.type}:${target.id}`;
}

function defaultDocumentFieldDraft(): WorkDocumentFieldDraft {
  return {
    kind: 'page',
    format: { kind: 'numeric', value: 'arabic' },
    targetId: '',
    targetName: '',
    hyperlink: true,
    mergeFormat: false,
  };
}

function documentFieldPreview(draft: WorkDocumentFieldDraft): string {
  const messages = resolveOfficeMessages();
  if (draft.kind === 'pageReference' && !draft.targetName) {
    return officeMessage(messages, 'document.insert.fieldSelectTarget');
  }
  if (draft.kind === 'mergeField' && !draft.targetName.trim()) {
    return officeMessage(messages, 'document.insert.fieldEnterMergeName');
  }
  return documentFieldDisplay(
    draft.kind,
    {
      pageNumber: 1,
      totalPages: 2,
      sectionNumber: 1,
      sectionPages: 1,
      wordCount: 12,
      characterCount: 24,
      referencePageNumber: 1,
      now: new Date(2026, 8, 6, 14, 5, 9),
    },
    documentFieldInstruction(draft.kind, documentFieldOptionsFromDraft(draft)),
  );
}
