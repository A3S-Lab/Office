import type { Editor } from '@tiptap/core';
import {
  CheckCircle2,
  MessageSquareReply,
  RotateCcw,
  Trash2,
  Unlink2,
  X,
} from 'lucide-react';
import {
  Fragment,
  type KeyboardEvent,
  type RefObject,
  useCallback,
  useEffect,
  useLayoutEffect,
  useMemo,
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
import {
  DOCUMENT_COMMENT_WINDOW_LIMIT,
  documentCommentKeyboardDestination,
  focusDocumentCommentItem,
  scrollDocumentCommentItemIntoView,
  useDocumentCommentWindow,
} from './document-comment-window';
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
  detached?: boolean;
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

interface PendingCommentFocus {
  commentId: string | null;
}

export function DocumentCommentsPanel({
  editor,
  comments,
  canDelete = canDeleteAnyComment,
  draft,
  draftAuthor = '我',
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
  canDelete?: (id: string) => boolean;
  draft: DocumentCommentDraft | null;
  draftAuthor?: string;
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
  const trackRef = useRef<HTMLOListElement>(null);
  const cardRefs = useRef(new Map<string, HTMLElement>());
  const anchorRefs = useRef(new Map<string, HTMLButtonElement>());
  const measuredCardHeightsRef = useRef(new Map<string, number>());
  const layoutRef = useRef<CommentTrackLayout>(emptyLayout);
  const completedDraftIdRef = useRef<string | null>(null);
  const pendingFocusRef = useRef<PendingCommentFocus | null>(null);
  const pendingRevealCommentIdRef = useRef<string | null>(null);
  const frameRef = useRef(0);
  const pendingFocusFrameRef = useRef(0);
  const windowFrameRef = useRef(0);
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
  const commentsById = useMemo(
    () => new Map(comments.map((comment) => [comment.id, comment] as const)),
    [comments],
  );
  const commentKeys = useMemo(
    () => comments.map((comment) => comment.id),
    [comments],
  );
  const commentIndexById = useMemo(
    () => new Map(commentKeys.map((id, index) => [id, index] as const)),
    [commentKeys],
  );
  const effectiveActiveCommentId =
    activeCommentId && commentsById.has(activeCommentId)
      ? activeCommentId
      : (comments.find((comment) => !comment.resolved)?.id ??
        comments[0]?.id ??
        null);
  const dirtyReplyKeys = useMemo(
    () =>
      Object.entries(drafts)
        .filter(([, value]) => Boolean(value.trim()))
        .map(([id]) => id),
    [drafts],
  );
  const commentWindow = useDocumentCommentWindow({
    keys: commentKeys,
    onRovingKeyChange: setActiveCommentId,
    pinnedKeys: dirtyReplyKeys,
    rovingKey: effectiveActiveCommentId,
  });
  const mountedCommentIndicesKey = commentWindow.mountedIndices.join(',');
  const mountedCommentIds = useMemo(
    () =>
      new Set(
        commentWindow.mountedIndices.flatMap((index) => {
          const id = commentKeys[index];
          return id ? [id] : [];
        }),
      ),
    [commentKeys, commentWindow.mountedIndices],
  );
  const layoutItemsById = useMemo(
    () => new Map(layout.items.map((item) => [item.id, item] as const)),
    [layout.items],
  );
  const visibleConnectorItems = useMemo(
    () =>
      layout.items.filter(
        (item) =>
          item.kind === 'draft' ||
          (mountedCommentIds.has(item.id) && !item.detached),
      ),
    [layout.items, mountedCommentIds],
  );

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
    if (!canDelete(commentId)) return;
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
    pendingFocusRef.current = { commentId: nextComment?.id ?? null };
    setActiveCommentId(nextComment?.id ?? null);
    setDrafts((current) => {
      const next = { ...current };
      delete next[commentId];
      return next;
    });
    onDelete(commentId);
    if (nextComment) selectDocumentComment(editor, nextComment);
  };
  const handleCommentAnchorKeyDown = (
    event: KeyboardEvent<HTMLButtonElement>,
    index: number,
  ) => {
    const nextIndex = documentCommentKeyboardDestination(
      event.key,
      index,
      comments.length,
    );
    if (nextIndex === null) return;
    event.preventDefault();
    const nextComment = comments[nextIndex];
    if (nextComment) selectDocumentComment(editor, nextComment);
    commentWindow.focusAt(nextIndex);
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
    const marks = [
      ...editor.view.dom.querySelectorAll<HTMLElement>(
        '[data-document-comment][data-comment-id]',
      ),
    ];
    const anchorRectsById = new Map<string, CommentAnchorRect>();

    for (const mark of marks) {
      const id = mark.dataset.commentId ?? '';
      mark.classList.toggle(
        'is-active-comment',
        id === effectiveActiveCommentId,
      );
      mark.classList.toggle(
        'is-resolved-comment',
        Boolean(commentsById.get(id)?.resolved),
      );
      const anchorRect = [...mark.getClientRects()].at(-1);
      if (id && anchorRect) {
        anchorRectsById.set(id, {
          right: anchorRect.right,
          top: anchorRect.top,
          height: anchorRect.height,
        });
      }
    }

    const entries: Array<{
      id: string;
      kind: CommentTrackItem['kind'];
      commentId?: string;
      from: number;
      detached?: boolean;
      anchorRect?: CommentAnchorRect;
    }> = comments.map((comment) => ({
      id: comment.id,
      kind: 'comment',
      commentId: comment.id,
      from: comment.from ?? Number.MAX_SAFE_INTEGER,
      detached: comment.detached,
      anchorRect: anchorRectsById.get(comment.id),
    }));
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
      const card = cardRefs.current.get(entry.id);
      const measuredCardHeight = card?.getBoundingClientRect().height ?? 0;
      let cardHeight = 164;
      if (entry.kind === 'comment') {
        const comment = commentsById.get(entry.id);
        const measurementKey = comment
          ? commentCardMeasurementKey(
              comment,
              comment.id === effectiveActiveCommentId,
            )
          : entry.id;
        if (measuredCardHeight > 0) {
          measuredCardHeightsRef.current.set(
            measurementKey,
            measuredCardHeight,
          );
        }
        cardHeight =
          measuredCardHeightsRef.current.get(measurementKey) ??
          (comment
            ? estimateCommentCardHeight(
                comment,
                comment.id === effectiveActiveCommentId,
              )
            : 112);
      } else if (measuredCardHeight > 0) {
        cardHeight = measuredCardHeight;
      }
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
        detached: entry.detached,
      };
    });
    const nextLayout = {
      width: Math.max(1, surfaceRect.width),
      height: Math.max(1, surfaceRect.height, trackOffset + nextCardTop),
      trackHeight: Math.max(1, nextCardTop),
      items,
    };
    layoutRef.current = nextLayout;
    setLayout((current) =>
      sameCommentTrackLayout(current, nextLayout) ? current : nextLayout,
    );
  }, [
    comments,
    commentsById,
    draft,
    editor,
    effectiveActiveCommentId,
    surfaceRef,
  ]);

  const scheduleMeasure = useCallback(() => {
    cancelAnimationFrame(frameRef.current);
    frameRef.current = requestAnimationFrame(measure);
  }, [measure]);

  const updateWindowFromViewport = useCallback(() => {
    const surface = surfaceRef.current;
    const panel = panelRef.current;
    const track = trackRef.current;
    if (!surface || !panel || !track) return;
    const focusedCard = panel.ownerDocument.activeElement?.closest<HTMLElement>(
      '.work-document-comment-card[data-comment-id]',
    );
    const focusedCommentId = focusedCard?.dataset.commentId;
    const focusedIndex = focusedCommentId
      ? commentIndexById.get(focusedCommentId)
      : undefined;
    if (focusedIndex !== undefined) {
      commentWindow.onViewportAnchorChange(focusedIndex);
      return;
    }
    const scroll = surface.closest<HTMLElement>('.work-document-scroll');
    const scrollRect = scroll?.getBoundingClientRect();
    const panelRect = panel.getBoundingClientRect();
    const trackRect = track.getBoundingClientRect();
    const panelHeader = panel.firstElementChild;
    const panelHeaderRect =
      panelHeader instanceof HTMLElement
        ? panelHeader.getBoundingClientRect()
        : undefined;
    const visibleTop =
      Math.max(
        scrollRect?.top ?? trackRect.top,
        panelRect.top,
        panelHeaderRect?.bottom ?? panelRect.top,
      ) - trackRect.top;
    const visibleComments = layoutRef.current.items.filter(
      (item) => item.kind === 'comment',
    );
    let anchor = visibleComments[0];
    for (const item of visibleComments) {
      if (item.cardTop > visibleTop + 8) break;
      anchor = item;
    }
    const index = anchor ? commentIndexById.get(anchor.id) : undefined;
    if (index !== undefined) commentWindow.onViewportAnchorChange(index);
  }, [commentIndexById, commentWindow.onViewportAnchorChange, surfaceRef]);

  const scheduleWindowUpdate = useCallback(() => {
    cancelAnimationFrame(windowFrameRef.current);
    windowFrameRef.current = requestAnimationFrame(updateWindowFromViewport);
  }, [updateWindowFromViewport]);

  useLayoutEffect(() => {
    measure();
    return () => {
      cancelAnimationFrame(frameRef.current);
      cancelAnimationFrame(windowFrameRef.current);
    };
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

  useLayoutEffect(() => {
    const pending = pendingFocusRef.current;
    if (!pending || pendingFocusFrameRef.current) return;
    pendingFocusFrameRef.current = requestAnimationFrame(() => {
      pendingFocusFrameRef.current = 0;
      const current = pendingFocusRef.current;
      if (!current) return;
      const target = current.commentId
        ? anchorRefs.current.get(current.commentId)
        : panelRef.current?.querySelector<HTMLElement>('.ds-icon-button.close');
      if (!target) return;
      pendingFocusRef.current = null;
      focusDocumentCommentItem(target);
    });
  }, [comments, effectiveActiveCommentId, mountedCommentIndicesKey]);

  useEffect(() => () => cancelAnimationFrame(pendingFocusFrameRef.current), []);

  useLayoutEffect(() => {
    const commentId = pendingRevealCommentIdRef.current;
    if (!commentId) return;
    const card = cardRefs.current.get(commentId);
    if (!card) return;
    pendingRevealCommentIdRef.current = null;
    scrollDocumentCommentItemIntoView(card);
  }, [effectiveActiveCommentId, mountedCommentIndicesKey]);

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
    const handleScroll = draft ? scheduleMeasure : scheduleWindowUpdate;
    scroll?.addEventListener('scroll', handleScroll, { passive: true });
    panel.addEventListener('scroll', handleScroll, { passive: true });
    window.addEventListener('resize', scheduleMeasure);
    return () => {
      observer?.disconnect();
      mutationObserver?.disconnect();
      scroll?.removeEventListener('scroll', handleScroll);
      panel.removeEventListener('scroll', handleScroll);
      window.removeEventListener('resize', scheduleMeasure);
    };
  }, [
    draft,
    editor,
    mountedCommentIndicesKey,
    scheduleMeasure,
    scheduleWindowUpdate,
    surfaceRef,
  ]);

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
        pendingRevealCommentIdRef.current = id;
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
        {visibleConnectorItems.map((item) => {
          const comment = item.commentId
            ? commentsById.get(item.commentId)
            : undefined;
          const bendX = Math.max(item.startX + 16, item.endX - 28);
          const className = [
            comment?.resolved ? 'resolved' : '',
            effectiveActiveCommentId === item.commentId ? 'active' : '',
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
        <ol
          ref={trackRef}
          className="work-document-comment-track"
          aria-label="文档批注"
          data-document-comment-count={comments.length}
          data-document-comment-mounted-count={
            commentWindow.mountedIndices.length
          }
          data-document-comment-track-height={Math.round(layout.trackHeight)}
          data-document-comment-window-end={commentWindow.range.end}
          data-document-comment-window-limit={DOCUMENT_COMMENT_WINDOW_LIMIT}
          data-document-comment-window-start={commentWindow.range.start}
          data-document-comment-windowed={
            commentWindow.range.windowed ? 'true' : 'false'
          }
          style={{ minHeight: `${layout.trackHeight}px` }}
        >
          {draft && (
            <li className="work-document-comment-draft-item" key={draft.id}>
              <DocumentCommentComposer
                author={draftAuthor}
                ref={(element) => {
                  const id = `draft:${draft.id}`;
                  if (element) cardRefs.current.set(id, element);
                  else cardRefs.current.delete(id);
                }}
                draft={draft}
                top={layoutItemsById.get(`draft:${draft.id}`)?.cardTop ?? 8}
                onCancel={() => void cancelDraft()}
                onDirtyChange={setDraftDirty}
                onSubmit={onSubmitDraft}
              />
            </li>
          )}
          {commentWindow.mountedIndices.map((index) => {
            const comment = comments[index];
            if (!comment) return null;
            const item = layoutItemsById.get(comment.id);
            const active = effectiveActiveCommentId === comment.id;
            return (
              <li
                aria-posinset={index + 1}
                aria-setsize={comments.length}
                ref={(element) => {
                  if (element) cardRefs.current.set(comment.id, element);
                  else cardRefs.current.delete(comment.id);
                }}
                className={`work-document-comment-card${comment.resolved ? ' resolved' : ''}${comment.detached ? ' detached' : ''}${active ? ' active' : ''}`}
                data-comment-id={comment.id}
                data-document-comment-detached={
                  comment.detached ? 'true' : 'false'
                }
                data-document-comment-card-top={Math.round(item?.cardTop ?? 8)}
                data-document-comment-item={index + 1}
                key={comment.id}
                style={{ top: `${item?.cardTop ?? 8}px` }}
                onFocusCapture={() => commentWindow.onItemFocus(index)}
              >
                <button
                  ref={(element) => {
                    commentWindow.registerItem(comment.id, element);
                    if (element) anchorRefs.current.set(comment.id, element);
                    else anchorRefs.current.delete(comment.id);
                  }}
                  type="button"
                  className="work-document-comment-anchor"
                  aria-label={`${comment.detached ? '查看已脱离正文的' : '定位'}批注 ${index + 1}`}
                  aria-current={active ? 'location' : undefined}
                  tabIndex={commentWindow.rovingIndex === index ? 0 : -1}
                  onKeyDown={(event) =>
                    handleCommentAnchorKeyDown(event, index)
                  }
                  onClick={() => {
                    setActiveCommentId(comment.id);
                    const selection = documentCommentSelection(editor, comment);
                    if (selection) {
                      editor.chain().focus().setTextSelection(selection).run();
                    }
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
                    {comment.detached ? (
                      <>
                        <Unlink2 size={10} aria-hidden="true" />
                        原文锚点已删除
                      </>
                    ) : (
                      comment.anchorText.trim() || '（空白字符）'
                    )}
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
                    disabled={!canDelete(comment.id)}
                    title={
                      canDelete(comment.id)
                        ? '删除批注'
                        : '批注模式只能删除自己创建的批注'
                    }
                    onClick={() => void deleteComment(comment.id, index)}
                  >
                    <Trash2 size={13} />
                    删除
                  </Button>
                </footer>
              </li>
            );
          })}
          {!comments.length && !draft && (
            <li className="work-document-comment-empty-item">
              <CollectionState
                className="work-document-comments-empty"
                role="status"
              >
                选择文字并添加批注。
              </CollectionState>
            </li>
          )}
        </ol>
      </aside>
      {officeDialog.dialog}
    </Fragment>
  );
}

function canDeleteAnyComment(): boolean {
  return true;
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
      candidate.detached === item.detached &&
      Math.abs(candidate.cardTop - item.cardTop) <= 0.5 &&
      Math.abs(candidate.startX - item.startX) <= 0.5 &&
      Math.abs(candidate.startY - item.startY) <= 0.5 &&
      Math.abs(candidate.endX - item.endX) <= 0.5 &&
      Math.abs(candidate.endY - item.endY) <= 0.5
    );
  });
}

function commentCardMeasurementKey(
  comment: WorkDocumentCommentView,
  active: boolean,
): string {
  const replies = (comment.replies ?? [])
    .map((reply) => `${reply.id}:${reply.text.length}`)
    .join('|');
  return `${comment.id}:${active ? 'active' : 'idle'}:${comment.text.length}:${replies}`;
}

function estimateCommentCardHeight(
  comment: WorkDocumentCommentView,
  active: boolean,
): number {
  const bodyLines = estimatedWrappedLines(comment.text, 34);
  const replyHeight = (comment.replies ?? []).reduce(
    (height, reply) => height + 32 + estimatedWrappedLines(reply.text, 31) * 14,
    0,
  );
  return 78 + bodyLines * 15 + replyHeight + (active ? 92 : 0);
}

function selectDocumentComment(
  editor: Editor,
  comment: WorkDocumentCommentView,
): void {
  if (editor.isDestroyed) return;
  const selection = documentCommentSelection(editor, comment);
  if (selection) editor.commands.setTextSelection(selection);
}

function documentCommentSelection(
  editor: Editor,
  comment: WorkDocumentCommentView,
): { from: number; to: number } | null {
  if (comment.detached || comment.from === null || comment.to === null) {
    return null;
  }
  const maximum = Math.max(1, editor.state.doc.content.size - 1);
  return {
    from: Math.min(comment.from, maximum),
    to: Math.min(comment.to, maximum),
  };
}

function estimatedWrappedLines(
  value: string,
  charactersPerLine: number,
): number {
  return Math.max(
    1,
    value
      .split('\n')
      .reduce(
        (lines, line) =>
          lines + Math.max(1, Math.ceil(line.length / charactersPerLine)),
        0,
      ),
  );
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
