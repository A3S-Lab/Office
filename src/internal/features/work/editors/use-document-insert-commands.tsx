import type { Editor } from '@tiptap/core';
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
import type { WorkDocumentFieldKind } from '../work-document-fields';
import type { WorkDocumentNoteKind } from '../work-document-notes';
import type { WorkDocumentContent } from '../work-types';
import { OfficeTextField, useOfficeDialog } from './office-controls';
import { readOfficeFileAsDataUrl } from './office-file-data';

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
    };

type WorkDocumentCrossReferenceTarget =
  | (WorkDocumentCaptionTarget & { type: 'caption' })
  | WorkDocumentBookmarkReferenceTarget;

export interface DocumentInsertCommands {
  dialog: ReactNode;
  insertCaption: (kind: WorkDocumentCaptionKind) => void;
  insertCrossReference: () => void;
  insertField: (kind: WorkDocumentFieldKind) => void;
  insertImage: (file: File) => Promise<void>;
  insertNote: (kind: WorkDocumentNoteKind) => boolean;
  refreshFields: () => boolean;
}

export function useDocumentInsertCommands({
  contentRef,
  editor,
}: {
  contentRef: MutableRefObject<WorkDocumentContent>;
  editor: Editor | null;
}): DocumentInsertCommands {
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
          title: '无法插入图片',
          description: '请选择 PNG、JPEG、GIF、WebP 或 BMP 图片。',
        });
        return;
      }
      if (file.size > 8 * 1024 * 1024) {
        await officeDialog.notice({
          title: '图片过大',
          description: '单张图片不能超过 8 MiB。',
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
          title: '无法读取图片',
          description: '文件可能已经移动或损坏，请重新选择。',
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
        title: '还没有可引用目标',
        description: '请先插入图片或表格题注，或在正文中添加书签。',
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
      editor.commands.refreshDocumentFields(contentRef.current);
    },
    [contentRef, editor],
  );

  const refreshFields = useCallback(
    () => editor?.commands.refreshDocumentFields(contentRef.current) ?? false,
    [contentRef, editor],
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
      {insertDialog?.kind === 'caption' && (
        <Dialog
          title={
            insertDialog.captionKind === 'figure'
              ? '插入图片题注'
              : '插入表格题注'
          }
          description="题注编号会自动生成，并可用于交叉引用。"
          className="work-document-insert-dialog"
          restoreFocusTarget={() => invokerRef.current}
          onClose={() => setInsertDialog(null)}
          footer={
            <>
              <Button tone="quiet" onClick={() => setInsertDialog(null)}>
                取消
              </Button>
              <Button tone="primary" onClick={submitCaption}>
                插入题注
              </Button>
            </>
          }
        >
          <label
            className="work-document-dialog-field"
            htmlFor={captionFieldId}
          >
            <span>题注文字</span>
            <OfficeTextField
              id={captionFieldId}
              data-autofocus
              aria-label="题注文字"
              value={insertDialog.title}
              maxLength={240}
              placeholder="例如：系统架构"
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
          title="插入交叉引用"
          description="选择正文中已有的图片、表格题注或书签。"
          className="work-document-reference-dialog"
          restoreFocusTarget={() => invokerRef.current}
          onClose={() => setInsertDialog(null)}
          footer={
            <>
              <Button tone="quiet" onClick={() => setInsertDialog(null)}>
                取消
              </Button>
              <Button
                tone="primary"
                disabled={!insertDialog.selectedKey}
                onClick={submitCrossReference}
              >
                插入引用
              </Button>
            </>
          }
        >
          <div
            className="work-document-reference-list"
            role="radiogroup"
            aria-label="可引用目标"
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
                      ? `书签 ${target.name}`
                      : target.display}
                  </strong>
                  <span>
                    {target.title ||
                      (target.type === 'bookmark'
                        ? '（空书签）'
                        : '（无题注文字）')}
                  </span>
                </label>
              );
            })}
          </div>
        </Dialog>
      )}
    </Fragment>
  );

  return {
    dialog,
    insertCaption,
    insertCrossReference,
    insertField,
    insertImage,
    insertNote,
    refreshFields,
  };
}

function referenceTargetKey(target: WorkDocumentCrossReferenceTarget): string {
  return `${target.type}:${target.id}`;
}
