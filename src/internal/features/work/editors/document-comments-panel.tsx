import type { Editor } from '@tiptap/core';
import {
  CheckCircle2,
  MessageSquareReply,
  RotateCcw,
  Trash2,
  X,
} from 'lucide-react';
import {
  Fragment,
  type KeyboardEvent,
  type RefObject,
  useCallback,
  useEffect,
  useLayoutEffect,
  useRef,
  useState,
} from 'react';
import {
  Button,
  CollectionState,
  IconButton,
} from '../../../design-system/primitives';
import {
  documentCommentDraftRange,
  type WorkDocumentCommentRange,
  type WorkDocumentCommentView,
} from '../work-document-comments';
import {
  DocumentCommentComposer,
  type DocumentCommentDraft,
} from './document-comment-composer';
import { OfficeTextArea, useOfficeDialog } from './office-controls';
import { useOfficeTaskPaneModal } from './office-task-pane';

interface CommentTrackItem {
  id: string;
  kind: 'comment' | 'draft';
  commentId?: string;
  cardTop: number;
  startX: number;
  startY: number;
  endX: number;
  endY: number;
}

interface CommentTrackLayout {
  width: number;
  height: number;
  trackHeight: number;
  items: CommentTrackItem[];
}

const emptyLayout: CommentTrackLayout = {
  width: 1,
  height: 1,
  trackHeight: 1,
  items: [],
};

export function DocumentCommentsPanel({
  editor,
  comments,
  draft,
  surfaceRef,
  onReply,
  onToggleResolved,
  onDelete,
  onCancelDraft,
  onSubmitDraft,
  onClose,
  onDirtyChange,
}: {
  editor: Editor;
  comments: WorkDocumentCommentView[];
  draft: DocumentCommentDraft | null;
  surfaceRef: RefObject<HTMLDivElement | null>;
  onReply: (id: string, text: string) => void;
  onToggleResolved: (id: string) => void;
  onDelete: (id: string) => void;
  onCancelDraft: () => void;
  onSubmitDraft: (text: string) => string | null;
  onClose: () => void;
  onDirtyChange?: (dirty: boolean) => void;
}) {
  const [drafts, setDrafts] = useState<Record<string, string>>({});
  const [activeCommentId, setActiveCommentId] = useState<string | null>(
    () =>
      comments.find((comment) => !comment.resolved)?.id ??
      comments[0]?.id ??
      null,
  );
  const [draftDirty, setDraftDirty] = useState(false);
  const [layout, setLayout] = useState<CommentTrackLayout>(emptyLayout);
  const panelRef = useRef<HTMLElement>(null);
  const trackRef = useRef<HTMLDivElement>(null);
  const cardRefs = useRef(new Map<string, HTMLElement>());
  const completedDraftIdRef = useRef<string | null>(null);
  const frameRef = useRef(0);
  const officeDialog = useOfficeDialog();
  const modal = useOfficeTaskPaneModal();
  const modalAttributes = modal
    ? ({ role: 'dialog', 'aria-modal': true } as const)
    : {};
  const unresolved = comments.filter((comment) => !comment.resolved).length;
  const repliesDirty = Object.values(drafts).some((value) =>
    Boolean(value.trim()),
  );
  const commentsDirty = draftDirty || repliesDirty;

  useEffect(() => {
    onDirtyChange?.(commentsDirty);
  }, [commentsDirty, onDirtyChange]);

  useEffect(
    () => () => {
      onDirtyChange?.(false);
    },
    [onDirtyChange],
  );

  useEffect(() => setDraftDirty(false), [draft?.id]);

  const submitReply = (commentId: string) => {
    const text = drafts[commentId]?.trim();
    if (!text) return;
    onReply(commentId, text);
    setDrafts((current) => ({ ...current, [commentId]: '' }));
  };
  const cancelDraft = async () => {
    if (draftDirty) {
      const confirmed = await officeDialog.confirm({
        title: '放弃未完成的批注？',
        description: '未添加的批注不会保留。',
        confirmLabel: '放弃内容',
        confirmTone: 'danger',
        restoreFocusTarget: () => {
          const draftCard = draft
            ? cardRefs.current.get(`draft:${draft.id}`)
            : null;
          return (
            draftCard?.querySelector<HTMLElement>('[aria-label="批注内容"]') ??
            editor.view.dom
          );
        },
      });
      if (!confirmed) return;
    }
    setDraftDirty(false);
    onCancelDraft();
  };
  const handlePanelKeyDown = (event: KeyboardEvent<HTMLElement>) => {
    if (event.key !== 'Escape' || event.defaultPrevented) return;
    event.preventDefault();
    event.stopPropagation();
    onClose();
  };
  const deleteComment = async (commentId: string, index: number) => {
    const replyDirty = Boolean(drafts[commentId]?.trim());
    const confirmed = await officeDialog.confirm({
      title: '删除批注？',
      description: replyDirty
        ? '批注、已有回复和未发送的回复都将删除。'
        : '批注及其回复将被删除。',
      confirmLabel: '删除',
      confirmTone: 'danger',
      restoreFocusTarget: replyDirty
        ? () =>
            cardRefs.current
              .get(commentId)
              ?.querySelector<HTMLElement>(
                '.work-document-comment-reply textarea',
              ) ?? editor.view.dom
        : undefined,
    });
    if (!confirmed) return;
    const nextComment = comments[index + 1] ?? comments[index - 1];
    setDrafts((current) => {
      const next = { ...current };
      delete next[commentId];
      return next;
    });
    onDelete(commentId);
    requestAnimationFrame(() => {
      const nextTarget = nextComment
        ? cardRefs.current
            .get(nextComment.id)
            ?.querySelector<HTMLElement>('.work-document-comment-anchor')
        : panelRef.current?.querySelector<HTMLElement>('.ds-icon-button.close');
      if (nextTarget) nextTarget.focus({ preventScroll: true });
      else if (!editor.isDestroyed)
        editor.view.dom.focus({ preventScroll: true });
    });
  };

  const measure = useCallback(() => {
    if (editor.isDestroyed) return;
    const surface = surfaceRef.current;
    const panel = panelRef.current;
    const track = trackRef.current;
    if (!surface || !panel || !track) return;

    const surfaceRect = surface.getBoundingClientRect();
    const panelRect = panel.getBoundingClientRect();
    const trackRect = track.getBoundingClientRect();
    const scroll = surface.closest<HTMLElement>('.work-document-scroll');
    const scrollRect = scroll?.getBoundingClientRect();
    const panelHeader = panel.firstElementChild;
    const panelHeaderRect =
      panelHeader instanceof HTMLElement
        ? panelHeader.getBoundingClientRect()
        : undefined;
    const trackOffset = trackRect.top - surfaceRect.top;
    const visibleTrackTop = Math.max(
      8,
      Math.max(
        scrollRect?.top ?? trackRect.top,
        panelHeaderRect?.bottom ?? trackRect.top,
      ) -
        trackRect.top +
        8,
    );
    const visibleTrackBottom =
      Math.min(scrollRect?.bottom ?? trackRect.bottom, panelRect.bottom) -
      trackRect.top -
      8;
    const commentsById = new Map(
      comments.map((comment) => [comment.id, comment] as const),
    );
    const marks = [
      ...editor.view.dom.querySelectorAll<HTMLElement>(
        '[data-document-comment][data-comment-id]',
      ),
    ];

    for (const mark of marks) {
      const id = mark.dataset.commentId ?? '';
      mark.classList.toggle('is-active-comment', id === activeCommentId);
      mark.classList.toggle(
        'is-resolved-comment',
        Boolean(commentsById.get(id)?.resolved),
      );
    }

    const entries: Array<{
      id: string;
      kind: CommentTrackItem['kind'];
      commentId?: string;
      from: number;
      anchorRect?: CommentAnchorRect;
    }> = comments.map((comment) => {
      const anchorRect = marks
        .filter((mark) => mark.dataset.commentId === comment.id)
        .flatMap((mark) => [...mark.getClientRects()])
        .at(-1);
      return {
        id: comment.id,
        kind: 'comment',
        commentId: comment.id,
        from: comment.from,
        anchorRect: anchorRect
          ? {
              right: anchorRect.right,
              top: anchorRect.top,
              height: anchorRect.height,
            }
          : undefined,
      };
    });
    if (draft) {
      const range = documentCommentDraftRange(editor) ?? draft;
      entries.push({
        id: `draft:${draft.id}`,
        kind: 'draft',
        from: range.from,
        anchorRect: draftCommentAnchorRect(editor, range),
      });
    }
    entries.sort(
      (left, right) =>
        left.from - right.from || left.kind.localeCompare(right.kind),
    );

    let nextCardTop = 8;
    const items = entries.map((entry) => {
      const { anchorRect } = entry;
      const cardHeight =
        cardRefs.current.get(entry.id)?.getBoundingClientRect().height ??
        (entry.kind === 'draft' ? 164 : 112);
      const preferredTop = anchorRect
        ? anchorRect.top - trackRect.top - 18
        : nextCardTop;
      const minimumTop = Math.max(8, nextCardTop);
      let cardTop = Math.max(minimumTop, preferredTop);
      if (entry.kind === 'draft' && scrollRect) {
        const maximumVisibleTop = Math.max(
          visibleTrackTop,
          visibleTrackBottom - cardHeight,
        );
        if (minimumTop <= maximumVisibleTop) {
          cardTop = Math.min(
            maximumVisibleTop,
            Math.max(minimumTop, visibleTrackTop, preferredTop),
          );
        }
      }
      nextCardTop = cardTop + cardHeight + 10;
      const endX = panelRect.left - surfaceRect.left + 1;
      const endY = trackOffset + cardTop + 24;
      const startX = anchorRect
        ? Math.min(anchorRect.right - surfaceRect.left + 3, endX - 24)
        : endX - 24;
      const startY = anchorRect
        ? anchorRect.top - surfaceRect.top + Math.min(anchorRect.height / 2, 12)
        : endY;
      return {
        id: entry.id,
        kind: entry.kind,
        commentId: entry.commentId,
        cardTop,
        startX,
        startY,
        endX,
        endY,
      };
    });
    const nextLayout = {
      width: Math.max(1, surfaceRect.width),
      height: Math.max(1, surfaceRect.height, trackOffset + nextCardTop),
      trackHeight: Math.max(1, nextCardTop),
      items,
    };
    setLayout((current) =>
      sameCommentTrackLayout(current, nextLayout) ? current : nextLayout,
    );
  }, [activeCommentId, comments, draft, editor, surfaceRef]);

  const scheduleMeasure = useCallback(() => {
    cancelAnimationFrame(frameRef.current);
    frameRef.current = requestAnimationFrame(measure);
  }, [measure]);

  useLayoutEffect(() => {
    measure();
    return () => cancelAnimationFrame(frameRef.current);
  }, [measure]);

  useEffect(() => {
    if (draft) {
      completedDraftIdRef.current = draft.id;
      if (activeCommentId !== null) setActiveCommentId(null);
      return;
    }
    const completedDraftId = completedDraftIdRef.current;
    if (
      completedDraftId &&
      comments.some((comment) => comment.id === completedDraftId)
    ) {
      completedDraftIdRef.current = null;
      if (activeCommentId !== completedDraftId)
        setActiveCommentId(completedDraftId);
      return;
    }
    if (
      activeCommentId &&
      comments.some((comment) => comment.id === activeCommentId)
    )
      return;
    setActiveCommentId(
      comments.find((comment) => !comment.resolved)?.id ??
        comments[0]?.id ??
        null,
    );
  }, [activeCommentId, comments, draft]);

  useEffect(() => {
    const surface = surfaceRef.current;
    const panel = panelRef.current;
    if (!surface || !panel) return;
    const observer =
      typeof ResizeObserver === 'undefined'
        ? null
        : new ResizeObserver(scheduleMeasure);
    observer?.observe(surface);
    observer?.observe(panel);
    observer?.observe(editor.view.dom);
    for (const card of cardRefs.current.values()) observer?.observe(card);
    const mutationObserver =
      typeof MutationObserver === 'undefined'
        ? null
        : new MutationObserver(scheduleMeasure);
    mutationObserver?.observe(editor.view.dom, {
      attributes: true,
      childList: true,
      characterData: true,
      subtree: true,
    });
    const scroll = surface.closest<HTMLElement>('.work-document-scroll');
    scroll?.addEventListener('scroll', scheduleMeasure, { passive: true });
    window.addEventListener('resize', scheduleMeasure);
    return () => {
      observer?.disconnect();
      mutationObserver?.disconnect();
      scroll?.removeEventListener('scroll', scheduleMeasure);
      window.removeEventListener('resize', scheduleMeasure);
    };
  }, [editor, scheduleMeasure, surfaceRef]);

  useEffect(() => {
    if (editor.isDestroyed) return;
    const editorDom = editor.view.dom;
    const openAnchoredComment = (event: Event) => {
      const target = event.target;
      if (!(target instanceof Element)) return;
      const anchor = target.closest<HTMLElement>(
        '[data-document-comment][data-comment-id]',
      );
      const id = anchor?.dataset.commentId;
      if (id && comments.some((comment) => comment.id === id)) {
        setActiveCommentId(id);
      }
    };
    editorDom.addEventListener('click', openAnchoredComment);
    return () => editorDom.removeEventListener('click', openAnchoredComment);
  }, [comments, editor]);

  return (
    <Fragment>
      <svg
        className="work-document-comment-connectors"
        viewBox={`0 0 ${layout.width} ${layout.height}`}
        preserveAspectRatio="none"
        aria-hidden="true"
      >
        {layout.items.map((item) => {
          const comment = comments.find(
            (candidate) => candidate.id === item.commentId,
          );
          const bendX = Math.max(item.startX + 16, item.endX - 28);
          const className = [
            comment?.resolved ? 'resolved' : '',
            activeCommentId === item.commentId ? 'active' : '',
            item.kind === 'draft' ? 'draft active' : '',
          ]
            .filter(Boolean)
            .join(' ');
          return (
            <g className={className} key={item.id}>
              <path
                d={`M ${item.startX} ${item.startY} L ${bendX} ${item.startY} L ${item.endX} ${item.endY}`}
              />
              <circle cx={item.startX} cy={item.startY} r="2.5" />
            </g>
          );
        })}
      </svg>
      <aside
        {...modalAttributes}
        ref={panelRef}
        className="work-document-comments-panel"
        aria-label="批注审阅"
        onKeyDown={handlePanelKeyDown}
      >
        <header>
          <div>
            <strong>批注</strong>
            <span>
              {draft
                ? '正在添加批注'
                : comments.length
                  ? `${unresolved} 条待处理 · 共 ${comments.length} 条`
                  : '没有批注'}
            </span>
          </div>
          <IconButton className="close" label="关闭批注审阅" onClick={onClose}>
            <X size={14} />
          </IconButton>
        </header>
        <div
          ref={trackRef}
          className="work-document-comment-track"
          style={{ minHeight: `${layout.trackHeight}px` }}
        >
          {draft && (
            <DocumentCommentComposer
              key={draft.id}
              ref={(element) => {
                const id = `draft:${draft.id}`;
                if (element) cardRefs.current.set(id, element);
                else cardRefs.current.delete(id);
              }}
              draft={draft}
              top={
                layout.items.find((item) => item.kind === 'draft')?.cardTop ?? 8
              }
              onCancel={() => void cancelDraft()}
              onDirtyChange={setDraftDirty}
              onSubmit={onSubmitDraft}
            />
          )}
          {comments.map((comment, index) => {
            const item = layout.items.find(
              (candidate) => candidate.id === comment.id,
            );
            const active = activeCommentId === comment.id;
            return (
              <article
                ref={(element) => {
                  if (element) cardRefs.current.set(comment.id, element);
                  else cardRefs.current.delete(comment.id);
                }}
                className={`${comment.resolved ? 'resolved' : ''}${active ? ' active' : ''}`}
                data-comment-id={comment.id}
                key={comment.id}
                style={{ top: `${item?.cardTop ?? 8}px` }}
                onFocusCapture={() => setActiveCommentId(comment.id)}
              >
                <button
                  type="button"
                  className="work-document-comment-anchor"
                  aria-label={`定位批注 ${index + 1}`}
                  onClick={() => {
                    setActiveCommentId(comment.id);
                    editor
                      .chain()
                      .focus()
                      .setTextSelection({
                        from: Math.min(
                          comment.from,
                          editor.state.doc.content.size,
                        ),
                        to: Math.min(comment.to, editor.state.doc.content.size),
                      })
                      .run();
                  }}
                >
                  <span className="work-document-comment-avatar">
                    {commentAuthorInitials(comment.author)}
                  </span>
                  <span className="work-document-comment-meta">
                    <strong>{comment.author}</strong>
                    <time dateTime={comment.date}>
                      {formatCommentDate(comment.date)}
                    </time>
                  </span>
                  <span className="work-document-comment-quote">
                    {comment.anchorText.trim() || '（空白字符）'}
                  </span>
                </button>
                <section className="work-document-comment-thread">
                  <p>{comment.text}</p>
                  {comment.replies?.map((reply) => (
                    <article className="reply" key={reply.id}>
                      <header>
                        <strong>{reply.author}</strong>
                        <time dateTime={reply.date}>
                          {formatCommentDate(reply.date)}
                        </time>
                      </header>
                      <p>{reply.text}</p>
                    </article>
                  ))}
                </section>
                <div className="work-document-comment-reply">
                  <OfficeTextArea
                    aria-label={`回复批注 ${index + 1}`}
                    value={drafts[comment.id] ?? ''}
                    placeholder="回复此批注…"
                    onChange={(event) =>
                      setDrafts((current) => ({
                        ...current,
                        [comment.id]: event.target.value,
                      }))
                    }
                    onKeyDown={(event) => {
                      if (
                        event.key !== 'Enter' ||
                        !(event.metaKey || event.ctrlKey)
                      )
                        return;
                      event.preventDefault();
                      submitReply(comment.id);
                    }}
                  />
                  <Button
                    size="compact"
                    tone="primary"
                    aria-label={`发送回复 ${index + 1}`}
                    disabled={!drafts[comment.id]?.trim()}
                    onClick={() => submitReply(comment.id)}
                  >
                    <MessageSquareReply size={13} />
                    回复
                  </Button>
                </div>
                <footer>
                  <Button
                    size="compact"
                    tone="quiet"
                    aria-label={`${comment.resolved ? '重新打开' : '解决'}批注 ${index + 1}`}
                    onClick={() => onToggleResolved(comment.id)}
                  >
                    {comment.resolved ? (
                      <RotateCcw size={13} />
                    ) : (
                      <CheckCircle2 size={13} />
                    )}
                    {comment.resolved ? '重新打开' : '解决'}
                  </Button>
                  <Button
                    size="compact"
                    tone="danger"
                    aria-label={`删除批注 ${index + 1}`}
                    onClick={() => void deleteComment(comment.id, index)}
                  >
                    <Trash2 size={13} />
                    删除
                  </Button>
                </footer>
              </article>
            );
          })}
          {!comments.length && !draft && (
            <CollectionState
              className="work-document-comments-empty"
              role="status"
            >
              选择文字并添加批注。
            </CollectionState>
          )}
        </div>
      </aside>
      {officeDialog.dialog}
    </Fragment>
  );
}

interface CommentAnchorRect {
  right: number;
  top: number;
  height: number;
}

function draftCommentAnchorRect(
  editor: Editor,
  range: WorkDocumentCommentRange,
): CommentAnchorRect | undefined {
  try {
    const position = Math.max(
      0,
      Math.min(editor.state.doc.content.size, range.to),
    );
    const coordinates = editor.view.coordsAtPos(position);
    return {
      right: Math.max(coordinates.left, coordinates.right),
      top: coordinates.top,
      height: Math.max(1, coordinates.bottom - coordinates.top),
    };
  } catch {
    return undefined;
  }
}

function sameCommentTrackLayout(
  current: CommentTrackLayout,
  next: CommentTrackLayout,
): boolean {
  if (
    Math.abs(current.width - next.width) > 0.5 ||
    Math.abs(current.height - next.height) > 0.5 ||
    Math.abs(current.trackHeight - next.trackHeight) > 0.5 ||
    current.items.length !== next.items.length
  )
    return false;
  return current.items.every((item, index) => {
    const candidate = next.items[index];
    return (
      candidate?.id === item.id &&
      candidate.kind === item.kind &&
      candidate.commentId === item.commentId &&
      Math.abs(candidate.cardTop - item.cardTop) <= 0.5 &&
      Math.abs(candidate.startX - item.startX) <= 0.5 &&
      Math.abs(candidate.startY - item.startY) <= 0.5 &&
      Math.abs(candidate.endX - item.endX) <= 0.5 &&
      Math.abs(candidate.endY - item.endY) <= 0.5
    );
  });
}

function commentAuthorInitials(author: string): string {
  const words = author.trim().split(/\s+/).filter(Boolean);
  if (!words.length) return '审';
  if (words.length === 1) return Array.from(words[0]).slice(0, 1).join('');
  return words
    .slice(0, 2)
    .map((word) => Array.from(word)[0])
    .join('')
    .toLocaleUpperCase();
}

function formatCommentDate(value: string): string {
  const time = Date.parse(value);
  if (!Number.isFinite(time)) return value;
  return new Intl.DateTimeFormat('zh-CN', {
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
  }).format(time);
}
