import { Copy, LayoutTemplate, Plus, Trash2, X } from 'lucide-react';
import { Button, IconButton } from '../../../design-system/primitives';
import { officeMessage } from '../../../i18n/office-locale';
import type {
  WorkPresentationContent,
  WorkPresentationLayout,
  WorkPresentationMaster,
  WorkSlide,
} from '../work-types';
import {
  CommittedOfficeTextField,
  OfficeCheckbox,
  OfficeColorPicker,
  OfficeSelect,
} from './office-controls';
import { useOfficeMessages } from './office-messages-context';
import type {
  PresentationEditorCanCommands,
  PresentationEditorCommands,
} from './presentation-command-types';
import type { PresentationDesignMode } from './presentation-editor-types';
import { handleOfficeTaskPaneKeyDown } from './office-task-pane';

export type PresentationDesignPanelCommands = Pick<
  PresentationEditorCommands,
  | 'addDesignPlaceholder'
  | 'applyPresentationLayout'
  | 'closeDesign'
  | 'createPresentationLayout'
  | 'deletePresentationLayout'
  | 'editDesign'
  | 'renamePresentationLayout'
  | 'renamePresentationMaster'
  | 'setPresentationLayoutBackground'
  | 'setPresentationMasterBackground'
  | 'togglePresentationLayoutBackground'
>;

export function PresentationDesignPanel({
  can,
  commands,
  content,
  slide,
  layout,
  master,
  mode,
}: {
  can: Pick<PresentationEditorCanCommands, 'deletePresentationLayout'>;
  commands: PresentationDesignPanelCommands;
  content: WorkPresentationContent;
  slide: WorkSlide;
  layout: WorkPresentationLayout;
  master: WorkPresentationMaster;
  mode: PresentationDesignMode;
}) {
  const messages = useOfficeMessages();
  return (
    <section
      className="work-presentation-design-panel"
      aria-label={officeMessage(messages, 'presentation.design.panelAria')}
      onKeyDown={(event) =>
        handleOfficeTaskPaneKeyDown(event, commands.closeDesign)
      }
    >
      <header>
        <div>
          <LayoutTemplate size={15} />
          <strong>
            {officeMessage(messages, 'presentation.design.title')}
          </strong>
          <span>
            {officeMessage(messages, 'presentation.design.masterCount', {
              count: String(content.masters?.length ?? 0),
            })}{' '}
            ·{' '}
            {officeMessage(messages, 'presentation.design.layoutCount', {
              count: String(content.layouts?.length ?? 0),
            })}
          </span>
        </div>
        <IconButton
          className="close"
          label={officeMessage(messages, 'presentation.design.close')}
          onClick={commands.closeDesign}
        >
          <X size={14} />
        </IconButton>
      </header>

      <div className="work-presentation-design-controls">
        <div className="work-office-field">
          <span>
            {officeMessage(messages, 'presentation.design.currentLayout')}
          </span>
          <OfficeSelect
            ariaLabel={officeMessage(
              messages,
              'presentation.design.layoutAria',
            )}
            value={layout.id}
            options={(content.layouts ?? []).map((candidate) => ({
              value: candidate.id,
              label: candidate.name,
            }))}
            onValueChange={commands.applyPresentationLayout}
          />
        </div>
        <OfficeCheckbox
          className="toggle"
          ariaLabel={officeMessage(
            messages,
            'presentation.design.useLayoutBackgroundAria',
          )}
          checked={slide.useLayoutBackground === true}
          onCheckedChange={commands.togglePresentationLayoutBackground}
        >
          {officeMessage(messages, 'presentation.design.useLayoutBackground')}
        </OfficeCheckbox>
        <Button
          size="compact"
          tone={mode === 'layout' ? 'primary' : 'secondary'}
          aria-pressed={mode === 'layout'}
          onClick={() => commands.editDesign('layout')}
        >
          {officeMessage(messages, 'presentation.design.editLayout')}
        </Button>
        <Button
          size="compact"
          tone={mode === 'master' ? 'primary' : 'secondary'}
          aria-pressed={mode === 'master'}
          onClick={() => commands.editDesign('master')}
        >
          {officeMessage(messages, 'presentation.design.editMaster')}
        </Button>
        <Button
          size="compact"
          aria-label={officeMessage(
            messages,
            'presentation.design.newLayoutAria',
          )}
          onClick={() => commands.createPresentationLayout(false)}
        >
          <Plus size={13} />
          {officeMessage(messages, 'presentation.design.newLayout')}
        </Button>
        <Button
          size="compact"
          aria-label={officeMessage(
            messages,
            'presentation.design.copyLayoutAria',
          )}
          onClick={() => commands.createPresentationLayout(true)}
        >
          <Copy size={13} />
          {officeMessage(messages, 'presentation.design.copyLayout')}
        </Button>
        <Button
          size="compact"
          tone="danger"
          aria-label={officeMessage(
            messages,
            'presentation.design.deleteLayoutAria',
          )}
          disabled={!can.deletePresentationLayout()}
          onClick={commands.deletePresentationLayout}
        >
          <Trash2 size={13} />
          {officeMessage(messages, 'presentation.design.deleteLayout')}
        </Button>
      </div>

      {mode === 'layout' ? (
        <div
          className="work-presentation-design-editing"
          data-design-mode="layout"
        >
          <strong>
            {officeMessage(messages, 'presentation.design.editingLayout')}
          </strong>
          <div className="work-office-field">
            <span>{officeMessage(messages, 'presentation.design.name')}</span>
            <CommittedOfficeTextField
              aria-label={officeMessage(
                messages,
                'presentation.design.layoutNameAria',
              )}
              value={layout.name}
              maxLength={255}
              formatValue={(name) => name}
              parseValue={normalizePresentationDesignName}
              onValueCommit={commands.renamePresentationLayout}
            />
          </div>
          <OfficeColorPicker
            compact
            className="work-color-tool"
            ariaLabel={officeMessage(
              messages,
              'presentation.design.layoutBgAria',
            )}
            value={layout.background ?? master.background}
            onValueChange={commands.setPresentationLayoutBackground}
          />
          <OfficeCheckbox
            className="toggle"
            ariaLabel={officeMessage(
              messages,
              'presentation.design.layoutUseMasterBgAria',
            )}
            checked={!layout.background}
            onCheckedChange={(checked) =>
              commands.setPresentationLayoutBackground(
                checked ? undefined : master.background,
              )
            }
          >
            {officeMessage(messages, 'presentation.design.useMasterBackground')}
          </OfficeCheckbox>
          <PlaceholderButtons onAdd={commands.addDesignPlaceholder} />
          <Button
            size="compact"
            tone="quiet"
            onClick={() => commands.editDesign('slide')}
          >
            {officeMessage(messages, 'presentation.design.returnToSlide')}
          </Button>
        </div>
      ) : null}

      {mode === 'master' ? (
        <div
          className="work-presentation-design-editing"
          data-design-mode="master"
        >
          <strong>
            {officeMessage(messages, 'presentation.design.editingMaster')}
          </strong>
          <div className="work-office-field">
            <span>{officeMessage(messages, 'presentation.design.name')}</span>
            <CommittedOfficeTextField
              aria-label={officeMessage(
                messages,
                'presentation.design.masterNameAria',
              )}
              value={master.name}
              maxLength={255}
              formatValue={(name) => name}
              parseValue={normalizePresentationDesignName}
              onValueCommit={commands.renamePresentationMaster}
            />
          </div>
          <OfficeColorPicker
            compact
            className="work-color-tool"
            ariaLabel={officeMessage(
              messages,
              'presentation.design.masterBgAria',
            )}
            value={master.background}
            onValueChange={commands.setPresentationMasterBackground}
          />
          <PlaceholderButtons onAdd={commands.addDesignPlaceholder} />
          <Button
            size="compact"
            tone="quiet"
            onClick={() => commands.editDesign('slide')}
          >
            {officeMessage(messages, 'presentation.design.returnToSlide')}
          </Button>
        </div>
      ) : null}
    </section>
  );
}

function normalizePresentationDesignName(draft: string): string | null {
  const name = draft.trim().slice(0, 255);
  return name || null;
}

function PlaceholderButtons({
  onAdd,
}: {
  onAdd: (type: 'title' | 'body') => void;
}) {
  const messages = useOfficeMessages();
  return (
    <div className="work-presentation-placeholder-actions">
      <Button
        size="compact"
        aria-label={officeMessage(
          messages,
          'presentation.design.addTitlePlaceholderAria',
        )}
        onClick={() => onAdd('title')}
      >
        {officeMessage(messages, 'presentation.design.addTitlePlaceholder')}
      </Button>
      <Button
        size="compact"
        aria-label={officeMessage(
          messages,
          'presentation.design.addContentPlaceholderAria',
        )}
        onClick={() => onAdd('body')}
      >
        {officeMessage(messages, 'presentation.design.addContentPlaceholder')}
      </Button>
    </div>
  );
}
