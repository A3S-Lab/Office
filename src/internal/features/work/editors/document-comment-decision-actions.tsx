import { Check, ListChecks, Undo2 } from 'lucide-react';
import { Button } from '../../../design-system/primitives';
import {
  type OfficeMessageCatalog,
  officeMessage,
} from '../../../i18n/office-locale';
import type { WorkDocumentCommentView } from '../work-document-comments';

export function DocumentCommentDecisionActions({
  comment,
  indexLabel,
  messages,
  onAccept,
  onProcess,
  onWithdraw,
}: {
  comment: WorkDocumentCommentView;
  indexLabel: string;
  messages: OfficeMessageCatalog;
  onAccept?: (id: string) => void;
  onProcess?: (id: string) => void;
  onWithdraw?: (id: string) => void;
}) {
  if (comment.resolved) {
    return (
      <>
        <span className="work-document-comment-decision">
          {officeMessage(
            messages,
            comment.decision === 'processed'
              ? 'document.comment.decision.processed'
              : 'document.comment.decision.accepted',
          )}
        </span>
        <Button
          size="compact"
          tone="quiet"
          aria-label={officeMessage(messages, 'document.comment.withdrawAria', {
            n: indexLabel,
          })}
          title={officeMessage(messages, 'document.comment.withdrawTitle')}
          onClick={() => onWithdraw?.(comment.id)}
        >
          <Undo2 size={13} />
          {officeMessage(messages, 'document.comment.withdraw')}
        </Button>
      </>
    );
  }

  return (
    <>
      <Button
        size="compact"
        tone="quiet"
        aria-label={officeMessage(messages, 'document.comment.acceptAria', {
          n: indexLabel,
        })}
        title={officeMessage(messages, 'document.comment.acceptTitle')}
        onClick={() => onAccept?.(comment.id)}
      >
        <Check size={13} />
        {officeMessage(messages, 'document.comment.accept')}
      </Button>
      <Button
        size="compact"
        tone="quiet"
        aria-label={officeMessage(messages, 'document.comment.processAria', {
          n: indexLabel,
        })}
        title={officeMessage(messages, 'document.comment.processTitle')}
        onClick={() => onProcess?.(comment.id)}
      >
        <ListChecks size={13} />
        {officeMessage(messages, 'document.comment.process')}
      </Button>
    </>
  );
}
