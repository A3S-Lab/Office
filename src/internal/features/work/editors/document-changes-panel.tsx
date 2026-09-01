import type { Editor } from '@tiptap/core';
import { Check, CheckCheck, FileDiff, Undo2, XCircle } from 'lucide-react';
import {
  type KeyboardEvent,
  useLayoutEffect,
  useMemo,
  useRef,
  useState,
} from 'react';
import { Button, CollectionState } from '../../../design-system/primitives';
import type { WorkDocumentChange } from '../work-document-changes';
import type { WorkDocumentChangeDecision } from '../work-types';
import {
  DOCUMENT_NAVIGATION_COLLECTION_WINDOW_LIMIT,
  type DocumentNavigationWindowSpacerEntry,
  useDocumentNavigationWindow,
} from './document-navigation-window';
import { DocumentTaskPane } from './document-task-pane';
import { useOfficeDialog } from './office-controls';

type DocumentChangeDecision = 'accept' | 'reject';

const DOCUMENT_CHANGE_ITEM_HEIGHT = 84;
const DOCUMENT_CHANGE_ITEM_GAP = 7;
const DOCUMENT_CHANGE_LIST_PADDING_TOP = 10;

interface PendingDocumentChangeFocus {
  changeKey: string | null;
  decision: DocumentChangeDecision;
}

export function DocumentChangesPanel({
  editor,
  changes,
  decisions = [],
  trackChanges,
  suggestionOnly = false,
  onDecideChanges,
  onTrackChangesChange,
  onClose,
}: {
  editor: Editor;
  changes: WorkDocumentChange[];
  decisions?: WorkDocumentChangeDecision[];
  trackChanges: boolean;
  suggestionOnly?: boolean;
  onDecideChanges?: (
    changes: readonly WorkDocumentChange[],
    decision: DocumentChangeDecision,
  ) => boolean;
  onTrackChangesChange: (enabled: boolean) => void;
  onClose: () => void;
}) {
  const officeDialog = useOfficeDialog();
  const decisionButtonRefs = useRef(new Map<string, HTMLButtonElement>());
  const pendingFocusRef = useRef<PendingDocumentChangeFocus | null>(null);
  const changeKeys = useMemo(
    () => changes.map(documentChangeWindowKey),
    [changes],
  );
  const [rovingChangeKey, setRovingChangeKey] = useState<string | null>(
    () => changeKeys[0] ?? null,
  );
  const effectiveRovingChangeKey =
    rovingChangeKey && changeKeys.includes(rovingChangeKey)
      ? rovingChangeKey
      : (changeKeys[0] ?? null);
  const changeWindow = useDocumentNavigationWindow<HTMLOListElement>({
    estimatedItemHeight: DOCUMENT_CHANGE_ITEM_HEIGHT,
    itemGap: DOCUMENT_CHANGE_ITEM_GAP,
    keys: changeKeys,
    listPaddingTop: DOCUMENT_CHANGE_LIST_PADDING_TOP,
    onRovingKeyChange: setRovingChangeKey,
    rovingKey: effectiveRovingChangeKey,
  });

  useLayoutEffect(() => {
    const pending = pendingFocusRef.current;
    if (!pending) return;
    const nextButton = pending.changeKey
      ? decisionButtonRefs.current.get(
          documentChangeDecisionKey(pending.changeKey, pending.decision),
        )
      : undefined;
    if (nextButton) {
      pendingFocusRef.current = null;
      nextButton.focus({ preventScroll: true });
      nextButton.scrollIntoView({ block: 'nearest' });
    } else if (!pending.changeKey && !editor.isDestroyed) {
      pendingFocusRef.current = null;
      editor.view.dom.focus({ preventScroll: true });
    }
  }, [changeWindow.mountedCount, changes, editor]);

  const decideChange = (
    change: WorkDocumentChange,
    index: number,
    decision: DocumentChangeDecision,
  ) => {
    const nextChange = changes[index + 1] ?? changes[index - 1] ?? null;
    const currentChangeKey = documentChangeWindowKey(change);
    const nextChangeKey = nextChange
      ? documentChangeWindowKey(nextChange)
      : null;
    pendingFocusRef.current = {
      changeKey: nextChangeKey,
      decision,
    };
    setRovingChangeKey(nextChangeKey);
    const handled = onDecideChanges
      ? onDecideChanges([change], decision)
      : decision === 'accept'
        ? editor.commands.acceptDocumentChange(change.id)
        : editor.commands.rejectDocumentChange(change.id);
    if (!handled) {
      pendingFocusRef.current = null;
      setRovingChangeKey(currentChangeKey);
    }
  };
  const handleSummaryKeyDown = (
    event: KeyboardEvent<HTMLButtonElement>,
    index: number,
  ) => {
    const nextIndex = documentChangeKeyboardDestination(
      event.key,
      index,
      changes.length,
    );
    if (nextIndex === null) return;
    event.preventDefault();
    changeWindow.focusAt(nextIndex);
  };
  const acceptAll = async () => {
    const confirmed = await officeDialog.confirm({
      title: '接受全部修订？',
      description: `将确认当前 ${changes.length} 项修订。`,
      confirmLabel: '全部接受',
    });
    if (confirmed && !editor.isDestroyed) {
      if (onDecideChanges) onDecideChanges(changes, 'accept');
      else editor.commands.acceptAllDocumentChanges();
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
      if (onDecideChanges) onDecideChanges(changes, 'reject');
      else editor.commands.rejectAllDocumentChanges();
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
          changes.length
            ? `${changes.length} 项待处理 · ${decisions.length} 项已决定`
            : decisions.length
              ? `没有待处理的修订 · ${decisions.length} 项已决定`
              : '没有待处理的修订'
        }
        closeLabel="关闭修订审阅"
        onClose={onClose}
      >
        {changes.length > 0 && !suggestionOnly && (
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
        {changes.length > 0 ? (
          <ol
            ref={changeWindow.viewportRef}
            className="work-document-change-list work-document-task-pane-body"
            aria-label="待处理修订"
            data-document-change-count={changes.length}
            data-document-change-mounted-count={changeWindow.mountedCount}
            data-document-change-window-end={changeWindow.range.end}
            data-document-change-window-limit={
              DOCUMENT_NAVIGATION_COLLECTION_WINDOW_LIMIT
            }
            data-document-change-window-start={changeWindow.range.start}
            data-document-change-windowed={
              changeWindow.range.windowed ? 'true' : 'false'
            }
            onScroll={changeWindow.onScroll}
          >
            {changeWindow.entries.map((entry) => {
              if (entry.kind === 'spacer') {
                return (
                  <DocumentChangeWindowSpacer
                    entry={entry}
                    key={`spacer-${entry.start}-${entry.end}`}
                  />
                );
              }
              const change = changes[entry.index];
              if (!change) return null;
              const changeKey = documentChangeWindowKey(change);
              return (
                <li
                  aria-posinset={entry.index + 1}
                  aria-setsize={changes.length}
                  className={`work-document-change-item ${change.kind}`}
                  data-document-change-item={entry.index + 1}
                  key={changeKey}
                  onFocusCapture={() => changeWindow.onItemFocus(entry.index)}
                >
                  <button
                    ref={(element) =>
                      changeWindow.registerItem(changeKey, element)
                    }
                    type="button"
                    className="work-document-change-summary"
                    tabIndex={changeWindow.rovingIndex === entry.index ? 0 : -1}
                    aria-label={`定位修订 ${entry.index + 1}`}
                    onKeyDown={(event) =>
                      handleSummaryKeyDown(event, entry.index)
                    }
                    onClick={() =>
                      editor
                        .chain()
                        .focus()
                        .setTextSelection({
                          from: Math.min(
                            change.from,
                            editor.state.doc.content.size,
                          ),
                          to: Math.min(
                            change.to,
                            editor.state.doc.content.size,
                          ),
                        })
                        .run()
                    }
                  >
                    <span>{documentChangeKindLabel(change.kind)}</span>
                    <strong>{change.text.trim() || '（空白字符）'}</strong>
                    <small>
                      {change.author}
                      {change.date ? ` · ${formatChangeDate(change.date)}` : ''}
                    </small>
                  </button>
                  {!suggestionOnly && (
                    <div>
                      <Button
                        ref={(element) => {
                          const key = documentChangeDecisionKey(
                            changeKey,
                            'accept',
                          );
                          if (element)
                            decisionButtonRefs.current.set(key, element);
                          else decisionButtonRefs.current.delete(key);
                        }}
                        tone="quiet"
                        aria-label={`接受修订 ${entry.index + 1}`}
                        onClick={() =>
                          decideChange(change, entry.index, 'accept')
                        }
                      >
                        <Check size={13} />
                        接受
                      </Button>
                      <Button
                        ref={(element) => {
                          const key = documentChangeDecisionKey(
                            changeKey,
                            'reject',
                          );
                          if (element)
                            decisionButtonRefs.current.set(key, element);
                          else decisionButtonRefs.current.delete(key);
                        }}
                        tone="quiet"
                        aria-label={`拒绝修订 ${entry.index + 1}`}
                        onClick={() =>
                          decideChange(change, entry.index, 'reject')
                        }
                      >
                        <XCircle size={13} />
                        拒绝
                      </Button>
                    </div>
                  )}
                </li>
              );
            })}
          </ol>
        ) : (
          <div className="work-document-change-list work-document-task-pane-body">
            <CollectionState
              className="work-document-changes-empty"
              icon={<FileDiff />}
              actions={
                suggestionOnly ? undefined : (
                  <Button
                    size="compact"
                    tone={trackChanges ? 'quiet' : 'primary'}
                    onClick={() => onTrackChangesChange(!trackChanges)}
                  >
                    {trackChanges ? '停止记录' : '开启修订'}
                  </Button>
                )
              }
              tone={trackChanges ? 'info' : 'neutral'}
              role="status"
            >
              {suggestionOnly
                ? '建议模式会自动记录带身份的新改动。'
                : trackChanges
                  ? '正在记录新的改动。'
                  : '当前没有记录新的改动。'}
            </CollectionState>
          </div>
        )}
        {decisions.length > 0 && (
          <DocumentChangeDecisionHistory decisions={decisions} />
        )}
      </DocumentTaskPane>
      {officeDialog.dialog}
    </>
  );
}

function DocumentChangeDecisionHistory({
  decisions,
}: {
  decisions: readonly WorkDocumentChangeDecision[];
}) {
  const visible = decisions.slice(-20).reverse();
  return (
    <section
      className="work-document-change-decisions"
      aria-label="修订决定记录"
    >
      <header>
        <strong>决定记录</strong>
        <span>{decisions.length} 项</span>
      </header>
      <ol>
        {visible.map((decision) => (
          <li
            data-document-change-decision={decision.decision}
            data-document-change-kind={decision.changeKind}
            key={decision.id}
          >
            <span>{decision.decision === 'accept' ? '已接受' : '已拒绝'}</span>
            <strong>{decision.text.trim() || '（空白字符）'}</strong>
            <small>
              {decision.suggestedBy} → {decision.decidedBy}
              {decision.decidedAt
                ? ` · ${formatChangeDate(decision.decidedAt)}`
                : ''}
            </small>
          </li>
        ))}
      </ol>
    </section>
  );
}

function documentChangeDecisionKey(
  changeKey: string,
  decision: DocumentChangeDecision,
): string {
  return `${changeKey}:${decision}`;
}

function documentChangeKindLabel(kind: WorkDocumentChange['kind']): string {
  if (kind === 'insertion') return '插入';
  if (kind === 'formatting') return '格式';
  if (kind === 'paragraph-formatting') return '段落格式';
  if (kind === 'numbering') return '编号格式';
  return '删除';
}

function documentChangeWindowKey(change: WorkDocumentChange): string {
  return `${change.kind}:${change.id}`;
}

function documentChangeKeyboardDestination(
  key: string,
  index: number,
  itemCount: number,
): number | null {
  if (!itemCount) return null;
  if (key === 'ArrowDown') return Math.min(itemCount - 1, index + 1);
  if (key === 'ArrowUp') return Math.max(0, index - 1);
  if (key === 'PageDown') return Math.min(itemCount - 1, index + 8);
  if (key === 'PageUp') return Math.max(0, index - 8);
  if (key === 'Home') return 0;
  if (key === 'End') return itemCount - 1;
  return null;
}

function DocumentChangeWindowSpacer({
  entry,
}: {
  entry: DocumentNavigationWindowSpacerEntry;
}) {
  return (
    <li
      aria-hidden="true"
      className="work-document-change-window-spacer"
      data-document-change-spacer={entry.position}
      data-document-change-spacer-end={entry.end}
      data-document-change-spacer-start={entry.start + 1}
      role="presentation"
      style={{ height: `${entry.height}px` }}
    />
  );
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
