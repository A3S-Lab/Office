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
import type { WorkDocumentCommentView } from '../work-document-comments';
import { OfficeTextArea } from './office-controls';

interface CommentTrackItem {
  id: string;
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
  surfaceRef,
  onReply,
  onToggleResolved,
  onDelete,
  onClose,
}: {
  editor: Editor;
  comments: WorkDocumentCommentView[];
  surfaceRef: RefObject<HTMLDivElement | null>;
  onReply: (id: string, text: string) => void;
  onToggleResolved: (id: string) => void;
  onDelete: (id: string) => void;
  onClose: () => void;
}) {
  const [drafts, setDrafts] = useState<Record<string, string>>({});
  const [activeCommentId, setActiveCommentId] = useState<string | null>(
    () =>
      comments.find((comment) => !comment.resolved)?.id ??
      comments[0]?.id ??
      null,
  );
  const [layout, setLayout] = useState<CommentTrackLayout>(emptyLayout);
  const panelRef = useRef<HTMLElement>(null);
  const trackRef = useRef<HTMLDivElement>(null);
  const cardRefs = useRef(new Map<string, HTMLElement>());
  const frameRef = useRef(0);
  const unresolved = comments.filter((comment) => !comment.resolved).length;

  const measure = useCallback(() => {
    const surface = surfaceRef.current;
    const panel = panelRef.current;
    const track = trackRef.current;
    if (!surface || !panel || !track) return;

    const surfaceRect = surface.getBoundingClientRect();
    const panelRect = panel.getBoundingClientRect();
    const trackRect = track.getBoundingClientRect();
    const trackOffset = trackRect.top - surfaceRect.top;
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

    let nextCardTop = 8;
    const items = comments.map((comment) => {
      const anchorRects = marks
        .filter((mark) => mark.dataset.commentId === comment.id)
        .flatMap((mark) => [...mark.getClientRects()]);
      const anchorRect = anchorRects.at(-1);
      const cardHeight =
        cardRefs.current.get(comment.id)?.getBoundingClientRect().height ?? 112;
      const preferredTop = anchorRect
        ? anchorRect.top - trackRect.top - 18
        : nextCardTop;
      const cardTop = Math.max(8, preferredTop, nextCardTop);
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
        id: comment.id,
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
  }, [activeCommentId, comments, editor, surfaceRef]);

  const scheduleMeasure = useCallback(() => {
    cancelAnimationFrame(frameRef.current);
    frameRef.current = requestAnimationFrame(measure);
  }, [measure]);

  useLayoutEffect(() => {
    scheduleMeasure();
    return () => cancelAnimationFrame(frameRef.current);
  }, [scheduleMeasure]);

  useEffect(() => {
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
  }, [activeCommentId, comments]);

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
    window.addEventListener('resize', scheduleMeasure);
    return () => {
      observer?.disconnect();
      mutationObserver?.disconnect();
      window.removeEventListener('resize', scheduleMeasure);
    };
  }, [editor, scheduleMeasure, surfaceRef]);

  useEffect(() => {
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
    editor.view.dom.addEventListener('click', openAnchoredComment);
    return () =>
      editor.view.dom.removeEventListener('click', openAnchoredComment);
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
            (candidate) => candidate.id === item.id,
          );
          const bendX = Math.max(item.startX + 16, item.endX - 28);
          return (
            <g
              className={`${comment?.resolved ? 'resolved' : ''}${activeCommentId === item.id ? ' active' : ''}`}
              key={item.id}
            >
              <path
                d={`M ${item.startX} ${item.startY} L ${bendX} ${item.startY} L ${item.endX} ${item.endY}`}
              />
              <circle cx={item.startX} cy={item.startY} r="2.5" />
            </g>
          );
        })}
      </svg>
      <aside
        ref={panelRef}
        className="work-document-comments-panel"
        aria-label="批注审阅"
      >
        <header>
          <div>
            <strong>批注</strong>
            <span>
              {comments.length
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
                onPointerEnter={() => setActiveCommentId(comment.id)}
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
                  />
                  <Button
                    tone="quiet"
                    aria-label={`发送回复 ${index + 1}`}
                    disabled={!drafts[comment.id]?.trim()}
                    onClick={() => {
                      const text = drafts[comment.id]?.trim();
                      if (!text) return;
                      onReply(comment.id, text);
                      setDrafts((current) => ({
                        ...current,
                        [comment.id]: '',
                      }));
                    }}
                  >
                    <MessageSquareReply size={13} />
                    回复
                  </Button>
                </div>
                <footer>
                  <Button
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
                    tone="quiet"
                    aria-label={`删除批注 ${index + 1}`}
                    onClick={() => onDelete(comment.id)}
                  >
                    <Trash2 size={13} />
                    删除
                  </Button>
                </footer>
              </article>
            );
          })}
          {!comments.length && (
            <CollectionState
              className="work-document-comments-empty"
              role="status"
            >
              选择文字并添加批注。
            </CollectionState>
          )}
        </div>
      </aside>
    </Fragment>
  );
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
