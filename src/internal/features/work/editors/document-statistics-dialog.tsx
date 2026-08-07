import { Button, Dialog } from '../../../design-system/primitives';
import type { DocumentTextStatistics } from './document-editor-support';

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
  const rows = [
    ['页数', pageCount],
    ['字数', statistics.wordCount],
    ['字符数（不计空格）', statistics.characterCountWithoutSpaces],
    ['字符数（计空格）', statistics.characterCountWithSpaces],
    ['段落数', statistics.paragraphCount],
  ] as const;

  return (
    <Dialog
      title="字数统计"
      description="当前文档的文本统计。"
      className="work-document-statistics-dialog"
      focusKey="document-statistics"
      restoreFocusTarget={restoreFocusTarget}
      onClose={onClose}
      footer={
        <Button tone="primary" onClick={onClose}>
          确定
        </Button>
      }
    >
      <dl className="work-document-statistics" aria-label="字数统计详情">
        {rows.map(([label, value]) => (
          <div key={label}>
            <dt>{label}</dt>
            <dd>{value}</dd>
          </div>
        ))}
      </dl>
    </Dialog>
  );
}
