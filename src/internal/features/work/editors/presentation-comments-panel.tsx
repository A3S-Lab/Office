import { MapPin, Trash2, X } from 'lucide-react';
import { officeMessage, resolveOfficeMessages } from '../../../i18n/office-locale';
import { useRef } from 'react';
import {
  Button,
  CollectionState,
  IconButton,
} from '../../../design-system/primitives';
import { useDialogFocusScope } from '../../../design-system/primitives/overlay/dialog-focus-scope';
import type { WorkSlide } from '../work-types';
import { CommittedOfficeTextArea } from './office-controls';
import { useOfficeMessages } from './office-messages-context';
import type { PresentationEditorCommands } from './presentation-command-types';
import {
  handleOfficeTaskPaneKeyDown,
  useOfficeTaskPaneModal,
} from './office-task-pane';

const PRESENTATION_COMMENTS_PANE_MODAL_QUERY = '(max-width: 640px)';

export type PresentationCommentsPanelCommands = Pick<
  PresentationEditorCommands,
  | 'closeComments'
  | 'deletePresentationComment'
  | 'locatePresentationComment'
  | 'updatePresentationComment'
>;

interface PresentationCommentView {
  slideId: string;
  slideName: string;
  slideNumber: number;
  commentIndex: number;
  id: string;
  author: string;
  date: string;
  text: string;
}

export function PresentationCommentsPanel({
  slides,
  activeCommentId,
  commands,
  restoreFocusTarget,
}: {
  slides: WorkSlide[];
  activeCommentId: string | null;
  commands: PresentationCommentsPanelCommands;
  restoreFocusTarget?: () => HTMLElement | null;
}) {
  const panelRef = useRef<HTMLElement>(null);
  const closeRef = useRef<HTMLButtonElement>(null);
  const modal = useOfficeTaskPaneModal(PRESENTATION_COMMENTS_PANE_MODAL_QUERY);
  const modalAttributes = modal
    ? ({ role: 'dialog', 'aria-modal': true } as const)
    : {};
  const focusScope = useDialogFocusScope<HTMLElement>({
    active: modal,
    initialFocus: () => closeRef.current,
    getActiveScope: () => panelRef.current,
    restoreFocusTarget,
  });
  const messages = useOfficeMessages();
  const comments = presentationCommentViews(slides, messages);
  return (
    <section
      {...modalAttributes}
      ref={panelRef}
      className="work-presentation-comments-panel"
      aria-label={officeMessage(messages, 'presentation.comments.panelAria')}
      onKeyDown={(event) => {
        focusScope.handleKeyDown(event);
        if (!event.defaultPrevented)
          handleOfficeTaskPaneKeyDown(event, commands.closeComments);
      }}
    >
      <header>
        <div>
          <strong>{officeMessage(messages, 'presentation.comments.title')}</strong>
          <span>
            {comments.length ? officeMessage(messages, 'presentation.comments.count', { count: String(comments.length) }) : officeMessage(messages, 'presentation.comments.empty')}
          </span>
        </div>
        <IconButton
          ref={closeRef}
          className="close"
          label={officeMessage(messages, 'presentation.comments.close')}
          onClick={commands.closeComments}
        >
          <X size={14} />
        </IconButton>
      </header>
      <div className="work-presentation-comment-list">
        {comments.map((comment, index) => (
          <article
            className={comment.id === activeCommentId ? 'active' : ''}
            key={`${comment.slideId}:${comment.id}`}
          >
            <button
              type="button"
              className="work-presentation-comment-location"
              aria-label={officeMessage(messages, 'presentation.comments.locateAria', { index: String(index + 1) })}
              onClick={() =>
                commands.locatePresentationComment(comment.slideId, comment.id)
              }
            >
              <MapPin size={12} />
              <span>
                {officeMessage(messages, 'presentation.comments.slideMeta', { number: String(comment.slideNumber), name: comment.slideName })}
              </span>
            </button>
            <header>
              <strong>{comment.author}</strong>
              <time dateTime={comment.date}>
                {formatCommentDate(comment.date)}
              </time>
            </header>
            <CommittedOfficeTextArea
              aria-label={officeMessage(messages, 'presentation.comments.editAria', { index: String(index + 1) })}
              value={comment.text}
              formatValue={(text) => text}
              parseValue={(draft) => draft}
              onFocus={() =>
                commands.locatePresentationComment(comment.slideId, comment.id)
              }
              onValueCommit={(text) =>
                commands.updatePresentationComment(
                  comment.slideId,
                  comment.id,
                  text,
                )
              }
            />
            <footer>
              <span>{officeMessage(messages, 'presentation.comments.itemLabel', { index: String(comment.commentIndex + 1) })}</span>
              <Button
                tone="quiet"
                aria-label={officeMessage(messages, 'presentation.comments.deleteAria', { index: String(index + 1) })}
                onClick={() =>
                  commands.deletePresentationComment(
                    comment.slideId,
                    comment.id,
                  )
                }
              >
                <Trash2 size={12} />
                {officeMessage(messages, 'presentation.comments.delete')}
              </Button>
            </footer>
          </article>
        ))}
        {!comments.length && (
          <CollectionState
            className="work-presentation-comments-empty"
            role="status"
          >
            {officeMessage(messages, 'presentation.comments.hint')}
          </CollectionState>
        )}
      </div>
    </section>
  );
}

export function presentationCommentCount(slides: readonly WorkSlide[]): number {
  return slides.reduce(
    (total, slide) => total + (slide.comments?.length ?? 0),
    0,
  );
}

function presentationCommentViews(
  slides: readonly WorkSlide[],
  messages = resolveOfficeMessages(),
): PresentationCommentView[] {
  return slides.flatMap((slide, slideIndex) =>
    (slide.comments ?? []).map((comment, commentIndex) => ({
      slideId: slide.id,
      slideName: slide.name,
      slideNumber: slideIndex + 1,
      commentIndex,
      id: comment.id,
      author:
        comment.author ||
        officeMessage(messages, 'presentation.comments.unknownAuthor'),
      date: comment.date,
      text: comment.text,
    })),
  );
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
