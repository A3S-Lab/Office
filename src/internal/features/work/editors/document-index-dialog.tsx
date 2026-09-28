import { Button, Dialog } from '../../../design-system/primitives';
import { officeMessage } from '../../../i18n/office-locale';
import type { OfficeMessageCatalog } from '../../../i18n/office-messages';
import type {
  WorkDocumentIndexFormat,
  WorkDocumentIndexLeader,
  WorkDocumentIndexOptions,
} from '../work-document-index';
import {
  OfficeCheckbox,
  OfficeSelect,
  type OfficeSelectOption,
} from './office-controls';
import { useOfficeMessages } from './office-messages-context';

export interface DocumentIndexDialogProps {
  editing: boolean;
  options: WorkDocumentIndexOptions;
  restoreFocusTarget: () => HTMLElement | null;
  onCancel: () => void;
  onOptionsChange: (options: WorkDocumentIndexOptions) => void;
  onSubmit: () => void;
}

export function DocumentIndexDialog({
  editing,
  options,
  restoreFocusTarget,
  onCancel,
  onOptionsChange,
  onSubmit,
}: DocumentIndexDialogProps) {
  const messages = useOfficeMessages();
  const update = (patch: Partial<WorkDocumentIndexOptions>) =>
    onOptionsChange({ ...options, ...patch });

  return (
    <Dialog
      title={officeMessage(
        messages,
        editing ? 'document.index.title.edit' : 'document.index.title.insert',
      )}
      description={officeMessage(messages, 'document.index.description')}
      className="work-document-index-dialog"
      restoreFocusTarget={restoreFocusTarget}
      onClose={onCancel}
      footer={
        <>
          <Button tone="quiet" onClick={onCancel}>
            {officeMessage(messages, 'document.index.cancel')}
          </Button>
          <Button tone="primary" onClick={onSubmit}>
            {officeMessage(
              messages,
              editing ? 'document.index.apply' : 'document.index.insert',
            )}
          </Button>
        </>
      }
    >
      <div className="work-document-index-dialog-grid">
        <div className="work-document-index-dialog-field">
          <span>{officeMessage(messages, 'document.index.columns')}</span>
          <OfficeSelect
            initialFocus
            ariaLabel={officeMessage(messages, 'document.index.columnsAria')}
            value={String(options.columns)}
            options={indexColumnOptions(messages)}
            onValueChange={(columns) => update({ columns: Number(columns) })}
          />
        </div>
        <div className="work-document-index-dialog-field">
          <span>{officeMessage(messages, 'document.index.layout')}</span>
          <OfficeSelect<WorkDocumentIndexFormat>
            ariaLabel={officeMessage(messages, 'document.index.layoutAria')}
            value={options.format}
            options={indexFormatOptions(messages)}
            onValueChange={(format) => update({ format })}
          />
        </div>
      </div>
      <div className="work-document-index-dialog-options">
        <OfficeCheckbox
          ariaLabel={officeMessage(messages, 'document.index.rightAlignAria')}
          checked={options.rightAlignPageNumbers}
          onCheckedChange={(rightAlignPageNumbers) =>
            update({ rightAlignPageNumbers })
          }
        >
          {officeMessage(messages, 'document.index.rightAlign')}
        </OfficeCheckbox>
        <div className="work-document-index-dialog-field">
          <span>{officeMessage(messages, 'document.index.leader')}</span>
          <OfficeSelect<WorkDocumentIndexLeader>
            ariaLabel={officeMessage(messages, 'document.index.leaderAria')}
            value={options.leader}
            options={indexLeaderOptions(messages)}
            disabled={!options.rightAlignPageNumbers}
            onValueChange={(leader) => update({ leader })}
          />
        </div>
      </div>
      <fieldset
        className="work-document-index-dialog-preview"
        aria-label={officeMessage(messages, 'document.index.previewAria')}
        data-index-format={options.format}
        data-index-leader={options.leader}
        data-index-right-align-page-numbers={String(
          options.rightAlignPageNumbers,
        )}
      >
        <legend>
          {officeMessage(messages, 'document.index.preview.legend')}
        </legend>
        <span>
          Architecture <i aria-hidden="true" /> 1, 3
        </span>
        <span className="nested">
          Runtime <i aria-hidden="true" /> 5
        </span>
      </fieldset>
    </Dialog>
  );
}

function indexColumnOptions(
  messages: OfficeMessageCatalog,
): readonly OfficeSelectOption[] {
  return [
    {
      value: '1',
      label: officeMessage(messages, 'document.pageLayout.columns.count1'),
    },
    {
      value: '2',
      label: officeMessage(messages, 'document.pageLayout.columns.count2'),
    },
    {
      value: '3',
      label: officeMessage(messages, 'document.pageLayout.columns.count3'),
    },
    {
      value: '4',
      label: officeMessage(messages, 'document.pageLayout.columns.count4'),
    },
  ];
}

function indexFormatOptions(
  messages: OfficeMessageCatalog,
): readonly OfficeSelectOption<WorkDocumentIndexFormat>[] {
  return [
    {
      value: 'indented',
      label: officeMessage(messages, 'document.index.layout.indented'),
    },
    {
      value: 'run-in',
      label: officeMessage(messages, 'document.index.layout.runIn'),
    },
  ];
}

function indexLeaderOptions(
  messages: OfficeMessageCatalog,
): readonly OfficeSelectOption<WorkDocumentIndexLeader>[] {
  return [
    {
      value: 'dot',
      label: officeMessage(messages, 'document.toc.leader.dot'),
    },
    {
      value: 'dash',
      label: officeMessage(messages, 'document.toc.leader.dash'),
    },
    {
      value: 'underline',
      label: officeMessage(messages, 'document.toc.leader.underline'),
    },
    {
      value: 'none',
      label: officeMessage(messages, 'document.toc.leader.none'),
    },
  ];
}
