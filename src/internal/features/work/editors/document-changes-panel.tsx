import type { Editor } from '@tiptap/core';
import { Check, CheckCheck, Undo2, XCircle } from 'lucide-react';
import { useLayoutEffect, useRef } from 'react';
import { Button, CollectionState } from '../../../design-system/primitives';
import type { WorkDocumentChange } from '../work-document-changes';
import { DocumentTaskPane } from './document-task-pane';
import { useOfficeDialog } from './office-controls';

type DocumentChangeDecision = 'accept' | 'reject';

interface PendingDocumentChangeFocus {
  changeId: string | null;
  decision: DocumentChangeDecision;
}

export function DocumentChangesPanel({
  editor,
  changes,
  onClose,
}: {
  editor: Editor;
  changes: WorkDocumentChange[];
  onClose: () => void;
}) {
  const officeDialog = useOfficeDialog();
  const decisionButtonRefs = useRef(new Map<string, HTMLButtonElement>());
  const pendingFocusRef = useRef<PendingDocumentChangeFocus | null>(null);

  useLayoutEffect(() => {
    const pending = pendingFocusRef.current;
    if (!pending) return;
    pendingFocusRef.current = null;
    const nextButton = pending.changeId
      ? decisionButtonRefs.current.get(
          documentChangeDecisionKey(pending.changeId, pending.decision),
        )
      : undefined;
    if (nextButton) {
      nextButton.focus({ preventScroll: true });
      nextButton.scrollIntoView({ block: 'nearest' });
    } else if (!editor.isDestroyed) {
      editor.view.dom.focus({ preventScroll: true });
    }
  }, [changes, editor]);

  const decideChange = (
    change: WorkDocumentChange,
    index: number,
    decision: DocumentChangeDecision,
  ) => {
    const nextChange = changes[index + 1] ?? changes[index - 1] ?? null;
    pendingFocusRef.current = {
      changeId: nextChange?.id ?? null,
      decision,
    };
    const handled =
      decision === 'accept'
        ? editor.commands.acceptDocumentChange(change.id)
        : editor.commands.rejectDocumentChange(change.id);
    if (!handled) pendingFocusRef.current = null;
  };
  const acceptAll = async () => {
    const confirmed = await officeDialog.confirm({
      title: '接受全部修订？',
      description: `将确认当前 ${changes.length} 项修订。`,
      confirmLabel: '全部接受',
    });
    if (confirmed && !editor.isDestroyed) {
      editor.commands.acceptAllDocumentChanges();
      requestAnimationFrame(() => {
        if (!editor.isDestroyed) editor.view.dom.focus();
      });
    }
  };
  const rejectAll = async () => {
    const confirmed = await officeDialog.confirm({
      title: '拒绝全部修订？',
      description: `将撤销当前 ${changes.length} 项修订。`,
      confirmLabel: '全部拒绝',
      confirmTone: 'danger',
    });
    if (confirmed && !editor.isDestroyed) {
      editor.commands.rejectAllDocumentChanges();
      requestAnimationFrame(() => {
        if (!editor.isDestroyed) editor.view.dom.focus();
      });
    }
  };

  return (
    <>
      <DocumentTaskPane
        className="work-document-changes-panel"
        title="修订审阅"
        description={
          changes.length ? `${changes.length} 项待处理` : '没有待处理的修订'
        }
        closeLabel="关闭修订审阅"
        onClose={onClose}
      >
        {changes.length > 0 && (
          <div className="work-document-changes-bulk-actions">
            <Button tone="quiet" onClick={() => void acceptAll()}>
              <CheckCheck size={13} />
              全部接受
            </Button>
            <Button tone="quiet" onClick={() => void rejectAll()}>
              <Undo2 size={13} />
              全部拒绝
            </Button>
          </div>
        )}
        <div className="work-document-change-list work-document-task-pane-body">
          {changes.map((change, index) => (
            <article
              className={change.kind}
              key={`${change.kind}-${change.id}`}
            >
              <button
                type="button"
                className="work-document-change-summary"
                aria-label={`定位修订 ${index + 1}`}
                onClick={() =>
                  editor
                    .chain()
                    .focus()
                    .setTextSelection({
                      from: Math.min(
                        change.from,
                        editor.state.doc.content.size,
                      ),
                      to: Math.min(change.to, editor.state.doc.content.size),
                    })
                    .run()
                }
              >
                <span>{change.kind === 'insertion' ? '插入' : '删除'}</span>
                <strong>{change.text.trim() || '（空白字符）'}</strong>
                <small>
                  {change.author}
                  {change.date ? ` · ${formatChangeDate(change.date)}` : ''}
                </small>
              </button>
              <div>
                <Button
                  ref={(element) => {
                    const key = documentChangeDecisionKey(change.id, 'accept');
                    if (element) decisionButtonRefs.current.set(key, element);
                    else decisionButtonRefs.current.delete(key);
                  }}
                  tone="quiet"
                  aria-label={`接受修订 ${index + 1}`}
                  onClick={() => decideChange(change, index, 'accept')}
                >
                  <Check size={13} />
                  接受
                </Button>
                <Button
                  ref={(element) => {
                    const key = documentChangeDecisionKey(change.id, 'reject');
                    if (element) decisionButtonRefs.current.set(key, element);
                    else decisionButtonRefs.current.delete(key);
                  }}
                  tone="quiet"
                  aria-label={`拒绝修订 ${index + 1}`}
                  onClick={() => decideChange(change, index, 'reject')}
                >
                  <XCircle size={13} />
                  拒绝
                </Button>
              </div>
            </article>
          ))}
          {!changes.length && (
            <CollectionState
              className="work-document-changes-empty"
              role="status"
            >
              开启修订后，改动会显示在这里。
            </CollectionState>
          )}
        </div>
      </DocumentTaskPane>
      {officeDialog.dialog}
    </>
  );
}

function documentChangeDecisionKey(
  changeId: string,
  decision: DocumentChangeDecision,
): string {
  return `${changeId}:${decision}`;
}

function formatChangeDate(value: string): string {
  const time = Date.parse(value);
  if (!Number.isFinite(time)) return value;
  return new Intl.DateTimeFormat('zh-CN', {
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
  }).format(time);
}
