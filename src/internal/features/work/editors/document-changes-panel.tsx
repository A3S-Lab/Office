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
import { officeMessage } from '../../../i18n/office-locale';
import type {
  OfficeMessageCatalog,
  OfficeMessageKey,
} from '../../../i18n/office-messages';
import type { WorkDocumentChange } from '../work-document-changes';
import type { WorkDocumentChangeDecision } from '../work-types';
import {
  DOCUMENT_NAVIGATION_COLLECTION_WINDOW_LIMIT,
  type DocumentNavigationWindowSpacerEntry,
  useDocumentNavigationWindow,
} from './document-navigation-window';
import { DocumentTaskPane } from './document-task-pane';
import { useOfficeDialog } from './office-controls';
import { useOfficeMessages } from './office-messages-context';

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
  const messages = useOfficeMessages();
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
      title: officeMessage(
        messages,
        'document.changes.acceptAllConfirm.title',
      ),
      description: officeMessage(
        messages,
        'document.changes.acceptAllConfirm.description',
        { count: String(changes.length) },
      ),
      confirmLabel: officeMessage(
        messages,
        'document.changes.acceptAllConfirm.confirm',
      ),
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
      title: officeMessage(
        messages,
        'document.changes.rejectAllConfirm.title',
      ),
      description: officeMessage(
        messages,
        'document.changes.rejectAllConfirm.description',
        { count: String(changes.length) },
      ),
      confirmLabel: officeMessage(
        messages,
        'document.changes.rejectAllConfirm.confirm',
      ),
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

  const description = changes.length
    ? officeMessage(messages, 'document.changes.description.pending', {
        pending: String(changes.length),
        decided: String(decisions.length),
      })
    : decisions.length
      ? officeMessage(messages, 'document.changes.description.decidedOnly', {
          decided: String(decisions.length),
        })
      : officeMessage(messages, 'document.changes.description.empty');

  return (
    <>
      <DocumentTaskPane
        className="work-document-changes-panel"
        title={officeMessage(messages, 'document.changes.title')}
        description={description}
        closeLabel={officeMessage(messages, 'document.changes.close')}
        onClose={onClose}
      >
        {changes.length > 0 && !suggestionOnly && (
          <div className="work-document-changes-bulk-actions">
            <Button tone="quiet" onClick={() => void acceptAll()}>
              <CheckCheck size={13} />
              {officeMessage(messages, 'document.changes.acceptAll')}
            </Button>
            <Button tone="quiet" onClick={() => void rejectAll()}>
              <Undo2 size={13} />
              {officeMessage(messages, 'document.changes.rejectAll')}
            </Button>
          </div>
        )}
        {changes.length > 0 ? (
          <ol
            ref={changeWindow.viewportRef}
            className="work-document-change-list work-document-task-pane-body"
            aria-label={officeMessage(messages, 'document.changes.listAria')}
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
              const indexLabel = String(entry.index + 1);
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
                    aria-label={officeMessage(
                      messages,
                      'document.changes.locateAria',
                      { n: indexLabel },
                    )}
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
                    <span>
                      {documentChangeKindLabel(messages, change.kind)}
                    </span>
                    <strong>
                      {change.text.trim() ||
                        officeMessage(messages, 'document.changes.blankText')}
                    </strong>
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
                        aria-label={officeMessage(
                          messages,
                          'document.changes.acceptAria',
                          { n: indexLabel },
                        )}
                        onClick={() =>
                          decideChange(change, entry.index, 'accept')
                        }
                      >
                        <Check size={13} />
                        {officeMessage(messages, 'document.changes.acceptShort')}
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
                        aria-label={officeMessage(
                          messages,
                          'document.changes.rejectAria',
                          { n: indexLabel },
                        )}
                        onClick={() =>
                          decideChange(change, entry.index, 'reject')
                        }
                      >
                        <XCircle size={13} />
                        {officeMessage(messages, 'document.changes.rejectShort')}
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
                    {officeMessage(
                      messages,
                      trackChanges
                        ? 'document.changes.track.stop'
                        : 'document.changes.track.start',
                    )}
                  </Button>
                )
              }
              tone={trackChanges ? 'info' : 'neutral'}
              role="status"
            >
              {officeMessage(
                messages,
                suggestionOnly
                  ? 'document.changes.empty.suggestion'
                  : trackChanges
                    ? 'document.changes.empty.recording'
                    : 'document.changes.empty.idle',
              )}
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
  const messages = useOfficeMessages();
  const visible = decisions.slice(-20).reverse();
  return (
    <section
      className="work-document-change-decisions"
      aria-label={officeMessage(messages, 'document.changes.history.aria')}
    >
      <header>
        <strong>
          {officeMessage(messages, 'document.changes.history.title')}
        </strong>
        <span>
          {officeMessage(messages, 'document.changes.history.count', {
            count: String(decisions.length),
          })}
        </span>
      </header>
      <ol>
        {visible.map((decision) => (
          <li
            data-document-change-decision={decision.decision}
            data-document-change-kind={decision.changeKind}
            key={decision.id}
          >
            <span>
              {officeMessage(
                messages,
                decision.decision === 'accept'
                  ? 'document.changes.history.accepted'
                  : 'document.changes.history.rejected',
              )}
            </span>
            <strong>
              {decision.text.trim() ||
                officeMessage(messages, 'document.changes.blankText')}
            </strong>
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

function documentChangeKindLabel(
  messages: OfficeMessageCatalog,
  kind: WorkDocumentChange['kind'],
): string {
  const keys: Record<WorkDocumentChange['kind'], OfficeMessageKey> = {
    insertion: 'document.changes.kind.insertion',
    deletion: 'document.changes.kind.deletion',
    formatting: 'document.changes.kind.formatting',
    'paragraph-formatting': 'document.changes.kind.paragraphFormatting',
    'table-formatting': 'document.changes.kind.tableFormatting',
    'row-formatting': 'document.changes.kind.rowFormatting',
    'cell-formatting': 'document.changes.kind.cellFormatting',
    'section-formatting': 'document.changes.kind.sectionFormatting',
    'paragraph-break': 'document.changes.kind.paragraphBreak',
    numbering: 'document.changes.kind.numbering',
    move: 'document.changes.kind.move',
  };
  return officeMessage(messages, keys[kind]);
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
