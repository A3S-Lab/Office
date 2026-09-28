import { Button, Dialog } from '../../../design-system/primitives';
import { officeMessage } from '../../../i18n/office-locale';
import type { OfficeMessageCatalog } from '../../../i18n/office-messages';
import type {
  WorkDocumentTableOfContentsLeader,
  WorkDocumentTableOfContentsOptions,
} from '../work-document-table-of-contents';
import {
  OfficeCheckbox,
  OfficeSelect,
  type OfficeSelectOption,
} from './office-controls';
import { useOfficeMessages } from './office-messages-context';

export interface DocumentTableOfContentsDialogProps {
  editing: boolean;
  options: WorkDocumentTableOfContentsOptions;
  restoreFocusTarget: () => HTMLElement | null;
  onCancel: () => void;
  onOptionsChange: (options: WorkDocumentTableOfContentsOptions) => void;
  onSubmit: () => void;
}

export function DocumentTableOfContentsDialog({
  editing,
  options,
  restoreFocusTarget,
  onCancel,
  onOptionsChange,
  onSubmit,
}: DocumentTableOfContentsDialogProps) {
  const messages = useOfficeMessages();
  const updateOptions = (update: Partial<WorkDocumentTableOfContentsOptions>) =>
    onOptionsChange({ ...options, ...update });
  const levelOptions = tableOfContentsLevelOptions(messages);

  return (
    <Dialog
      title={officeMessage(
        messages,
        editing ? 'document.toc.title.edit' : 'document.toc.title.insert',
      )}
      description={officeMessage(messages, 'document.toc.description')}
      className="work-document-table-of-contents-dialog"
      restoreFocusTarget={restoreFocusTarget}
      onClose={onCancel}
      footer={
        <>
          <Button tone="quiet" onClick={onCancel}>
            {officeMessage(messages, 'document.toc.cancel')}
          </Button>
          <Button tone="primary" onClick={onSubmit}>
            {officeMessage(
              messages,
              editing ? 'document.toc.apply' : 'document.toc.insert',
            )}
          </Button>
        </>
      }
    >
      <div className="work-document-table-of-contents-dialog-grid">
        <div className="work-document-table-of-contents-dialog-field">
          <span>{officeMessage(messages, 'document.toc.startLevel')}</span>
          <OfficeSelect
            initialFocus
            ariaLabel={officeMessage(messages, 'document.toc.startLevelAria')}
            value={String(options.minLevel)}
            options={levelOptions}
            onValueChange={(value) => {
              const minLevel = Number(value);
              updateOptions({
                minLevel,
                maxLevel: Math.max(minLevel, options.maxLevel),
              });
            }}
          />
        </div>
        <div className="work-document-table-of-contents-dialog-field">
          <span>{officeMessage(messages, 'document.toc.endLevel')}</span>
          <OfficeSelect
            ariaLabel={officeMessage(messages, 'document.toc.endLevelAria')}
            value={String(options.maxLevel)}
            options={levelOptions.map((option) => ({
              ...option,
              disabled: Number(option.value) < options.minLevel,
            }))}
            onValueChange={(value) =>
              updateOptions({ maxLevel: Number(value) })
            }
          />
        </div>
      </div>
      <div className="work-document-table-of-contents-dialog-options">
        <OfficeCheckbox
          ariaLabel={officeMessage(messages, 'document.toc.hyperlinksAria')}
          checked={options.hyperlinks}
          onCheckedChange={(hyperlinks) => updateOptions({ hyperlinks })}
        >
          {officeMessage(messages, 'document.toc.hyperlinks')}
        </OfficeCheckbox>
        <OfficeCheckbox
          ariaLabel={officeMessage(messages, 'document.toc.showPagesAria')}
          checked={options.showPageNumbers}
          onCheckedChange={(showPageNumbers) =>
            updateOptions({
              showPageNumbers,
              rightAlignPageNumbers:
                showPageNumbers && options.rightAlignPageNumbers,
            })
          }
        >
          {officeMessage(messages, 'document.toc.showPages')}
        </OfficeCheckbox>
        <OfficeCheckbox
          ariaLabel={officeMessage(messages, 'document.toc.rightAlignAria')}
          checked={options.rightAlignPageNumbers}
          disabled={!options.showPageNumbers}
          onCheckedChange={(rightAlignPageNumbers) =>
            updateOptions({ rightAlignPageNumbers })
          }
        >
          {officeMessage(messages, 'document.toc.rightAlign')}
        </OfficeCheckbox>
      </div>
      <div className="work-document-table-of-contents-dialog-leader">
        <span>{officeMessage(messages, 'document.toc.leader')}</span>
        <OfficeSelect<WorkDocumentTableOfContentsLeader>
          ariaLabel={officeMessage(messages, 'document.toc.leaderAria')}
          value={options.leader}
          options={tableOfContentsLeaderOptions(messages)}
          disabled={!options.showPageNumbers || !options.rightAlignPageNumbers}
          onValueChange={(leader) => updateOptions({ leader })}
        />
      </div>
      <fieldset
        className="work-document-table-of-contents-dialog-preview"
        aria-label={officeMessage(messages, 'document.toc.previewAria')}
        data-toc-leader={options.leader}
        data-toc-right-align-page-numbers={String(
          options.rightAlignPageNumbers,
        )}
      >
        <legend>
          {officeMessage(messages, 'document.toc.preview.legend')}
        </legend>
        <span>
          {officeMessage(messages, 'document.toc.preview.item1')}{' '}
          <i aria-hidden="true" />
          {options.showPageNumbers ? '1' : ''}
        </span>
        <span className="nested">
          {officeMessage(messages, 'document.toc.preview.item2')}{' '}
          <i aria-hidden="true" />
          {options.showPageNumbers ? '3' : ''}
        </span>
      </fieldset>
    </Dialog>
  );
}

function tableOfContentsLevelOptions(
  messages: OfficeMessageCatalog,
): readonly OfficeSelectOption[] {
  return Array.from({ length: 9 }, (_, index) => ({
    value: String(index + 1),
    label: officeMessage(messages, 'document.toc.headingLevel', {
      n: String(index + 1),
    }),
  }));
}

function tableOfContentsLeaderOptions(
  messages: OfficeMessageCatalog,
): readonly OfficeSelectOption<WorkDocumentTableOfContentsLeader>[] {
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
