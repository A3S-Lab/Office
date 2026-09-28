import { Button, Dialog } from '../../../design-system/primitives';
import { officeMessage } from '../../../i18n/office-locale';
import type { OfficeMessageKey } from '../../../i18n/office-messages';
import type { DocumentTextStatistics } from './document-editor-support';
import { useOfficeMessages } from './office-messages-context';

const STATISTIC_ROWS = [
  ['document.statistics.pages', 'pageCount'],
  ['document.statistics.words', 'wordCount'],
  ['document.statistics.charsNoSpaces', 'characterCountWithoutSpaces'],
  ['document.statistics.charsWithSpaces', 'characterCountWithSpaces'],
  ['document.statistics.paragraphs', 'paragraphCount'],
] as const satisfies ReadonlyArray<
  readonly [OfficeMessageKey, keyof DocumentTextStatistics | 'pageCount']
>;

export function DocumentStatisticsDialog({
  pageCount,
  restoreFocusTarget,
  statistics,
  onClose,
}: {
  pageCount: number;
  restoreFocusTarget: () => HTMLElement | null;
  statistics: DocumentTextStatistics;
  onClose: () => void;
}) {
  const messages = useOfficeMessages();
  const values = {
    pageCount,
    ...statistics,
  };

  return (
    <Dialog
      title={officeMessage(messages, 'document.statistics.title')}
      description={officeMessage(messages, 'document.statistics.description')}
      className="work-document-statistics-dialog"
      focusKey="document-statistics"
      restoreFocusTarget={restoreFocusTarget}
      onClose={onClose}
      footer={
        <Button tone="primary" onClick={onClose}>
          {officeMessage(messages, 'document.statistics.confirm')}
        </Button>
      }
    >
      <dl
        className="work-document-statistics"
        aria-label={officeMessage(messages, 'document.statistics.detailsAria')}
      >
        {STATISTIC_ROWS.map(([key, field]) => {
          const label = officeMessage(messages, key);
          return (
            <div key={key}>
              <dt>{label}</dt>
              <dd>{values[field]}</dd>
            </div>
          );
        })}
      </dl>
    </Dialog>
  );
}
