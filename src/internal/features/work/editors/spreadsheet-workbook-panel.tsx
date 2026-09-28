import type { Selection } from '@fortune-sheet/core';
import { Plus, Trash2, X } from 'lucide-react';
import { useEffect, useRef, useState } from 'react';
import { officeMessage } from '../../../i18n/office-locale';
import {
  Button,
  CollectionState,
  IconButton,
  InlineNotice,
} from '../../../design-system/primitives';
import { useDialogFocusScope } from '../../../design-system/primitives/overlay/dialog-focus-scope';
import { isValidSpreadsheetDefinedName } from '../work-spreadsheet-ranges';
import { createWorkId } from '../work-templates';
import type {
  WorkSpreadsheetContent,
  WorkSpreadsheetNamedRange,
} from '../work-types';
import { OfficeSelect, OfficeTextField } from './office-controls';
import { useOfficeMessages } from './office-messages-context';
import { SpreadsheetChartPanel } from './spreadsheet-chart-panel';
import type {
  SpreadsheetEditorCanCommands,
  SpreadsheetEditorCommands,
} from './spreadsheet-command-controller';
import { SpreadsheetConditionalFormatPanel } from './spreadsheet-conditional-format-panel';
import { SpreadsheetFormulaPanel } from './spreadsheet-formula-panel';
import {
  handleOfficeTaskPaneKeyDown,
  useOfficeTaskPaneEscape,
  useOfficeTaskPaneModal,
} from './office-task-pane';
import { useOfficeDraft } from './use-office-draft';
import { SpreadsheetPivotPanel } from './spreadsheet-pivot-panel';
import { SpreadsheetPrintSettingsPanel } from './spreadsheet-print-settings-panel';
import { SpreadsheetProtectionPanel } from './spreadsheet-protection-panel';

export type SpreadsheetWorkbookPanelView =
  | 'names'
  | 'print-area'
  | 'conditional-formatting'
  | 'protection'
  | 'charts'
  | 'formulas'
  | 'pivots';

interface SpreadsheetWorkbookPanelProps {
  id?: string;
  content: WorkSpreadsheetContent;
  view: SpreadsheetWorkbookPanelView;
  activeSheetId: string;
  selection?: Selection;
  can: Pick<SpreadsheetEditorCanCommands, 'recalculateFormula'>;
  commands: Pick<
    SpreadsheetEditorCommands,
    'recalculateFormula' | 'setSpreadsheetContent'
  >;
  restoreFocusTarget?: () => HTMLElement | null;
  onClose: () => void;
}

export function SpreadsheetWorkbookPanel({
  id,
  content,
  view,
  activeSheetId,
  selection,
  can,
  commands,
  restoreFocusTarget,
  onClose,
}: SpreadsheetWorkbookPanelProps) {
  const messages = useOfficeMessages();
  const title = panelTitle(view, messages);
  const panelRef = useRef<HTMLElement>(null);
  const closeRef = useRef<HTMLButtonElement>(null);
  const modal = useOfficeTaskPaneModal();
  const modalAttributes = modal
    ? ({ role: 'dialog', 'aria-modal': true } as const)
    : {};
  const focusScope = useDialogFocusScope<HTMLElement>({
    active: modal,
    initialFocus: () => closeRef.current,
    getActiveScope: () => panelRef.current,
    restoreFocusTarget,
  });
  useOfficeTaskPaneEscape(true, onClose);
  return (
    <section
      {...modalAttributes}
      ref={panelRef}
      id={id}
      className="work-spreadsheet-workbook-panel"
      aria-label={title.label}
      data-view={view}
      onKeyDown={(event) => {
        focusScope.handleKeyDown(event);
        if (!event.defaultPrevented)
          handleOfficeTaskPaneKeyDown(event, onClose);
      }}
    >
      <header>
        <div>
          <strong>{title.heading}</strong>
          <span>{title.description}</span>
        </div>
        <IconButton
          ref={closeRef}
          label={officeMessage(messages, 'spreadsheet.workbook.closeAria', { heading: title.heading })}
          onClick={onClose}
        >
          <X size={14} />
        </IconButton>
      </header>
      <section
        key={view}
        className="work-spreadsheet-workbook-panel-body"
        aria-label={officeMessage(messages, 'spreadsheet.workbook.contentAria', { heading: title.heading })}
      >
        {view === 'names' ? (
          <NamedRangeManager
            content={content}
            onChange={commands.setSpreadsheetContent}
          />
        ) : view === 'formulas' ? (
          <SpreadsheetFormulaPanel
            content={content}
            canRecalculateSelection={can.recalculateFormula('selection')}
            canRecalculateWorkbook={can.recalculateFormula('workbook')}
            onChange={commands.setSpreadsheetContent}
            onRecalculate={commands.recalculateFormula}
          />
        ) : view === 'charts' ? (
          <SpreadsheetChartPanel
            content={content}
            activeSheetId={activeSheetId}
            selection={selection}
            onChange={commands.setSpreadsheetContent}
          />
        ) : view === 'pivots' ? (
          <SpreadsheetPivotPanel
            content={content}
            activeSheetId={activeSheetId}
            selection={selection}
            onChange={commands.setSpreadsheetContent}
          />
        ) : view === 'conditional-formatting' ? (
          <SpreadsheetConditionalFormatPanel
            content={content}
            onChange={commands.setSpreadsheetContent}
          />
        ) : view === 'protection' ? (
          <SpreadsheetProtectionPanel
            content={content}
            onChange={commands.setSpreadsheetContent}
          />
        ) : (
          <SpreadsheetPrintSettingsPanel
            content={content}
            onChange={commands.setSpreadsheetContent}
          />
        )}
      </section>
    </section>
  );
}

function panelTitle(
  view: SpreadsheetWorkbookPanelView,
  messages: ReturnType<typeof useOfficeMessages>,
) {
  if (view === 'names') {
    return {
      label: officeMessage(messages, 'spreadsheet.workbook.nav.names'),
      heading: officeMessage(messages, 'spreadsheet.workbook.nav.namesHeading'),
      description: officeMessage(messages, 'spreadsheet.workbook.nav.namesDesc'),
    };
  }
  if (view === 'conditional-formatting') {
    return {
      label: officeMessage(messages, 'spreadsheet.workbook.nav.cf'),
      heading: officeMessage(messages, 'spreadsheet.workbook.nav.cfHeading'),
      description: officeMessage(messages, 'spreadsheet.workbook.nav.cfDesc'),
    };
  }
  if (view === 'formulas') {
    return {
      label: officeMessage(messages, 'spreadsheet.workbook.nav.formula'),
      heading: officeMessage(messages, 'spreadsheet.workbook.nav.formulaHeading'),
      description: officeMessage(messages, 'spreadsheet.workbook.nav.formulaDesc'),
    };
  }
  if (view === 'charts') {
    return {
      label: officeMessage(messages, 'spreadsheet.workbook.nav.charts'),
      heading: officeMessage(messages, 'spreadsheet.workbook.nav.chartsHeading'),
      description: officeMessage(messages, 'spreadsheet.workbook.nav.chartsDesc'),
    };
  }
  if (view === 'pivots') {
    return {
      label: officeMessage(messages, 'spreadsheet.workbook.nav.pivot'),
      heading: officeMessage(messages, 'spreadsheet.workbook.nav.pivotHeading'),
      description: officeMessage(messages, 'spreadsheet.workbook.nav.pivotDesc'),
    };
  }
  if (view === 'protection') {
    return {
      label: officeMessage(messages, 'spreadsheet.workbook.nav.protection'),
      heading: officeMessage(messages, 'spreadsheet.workbook.nav.protectionHeading'),
      description: officeMessage(messages, 'spreadsheet.workbook.nav.protectionDesc'),
    };
  }
  return {
    label: officeMessage(messages, 'spreadsheet.workbook.nav.print'),
    heading: officeMessage(messages, 'spreadsheet.workbook.nav.printHeading'),
    description: officeMessage(messages, 'spreadsheet.workbook.nav.printDesc'),
  };
}

interface NamedRangeDraft {
  id?: string;
  name: string;
  reference: string;
  scopeSheetId: string;
  comment: string;
}

function NamedRangeManager({
  content,
  onChange,
}: {
  content: WorkSpreadsheetContent;
  onChange: (content: WorkSpreadsheetContent) => void;
}) {
  const messages = useOfficeMessages();
  const ranges = content.namedRanges ?? [];
  const [selectedId, setSelectedId] = useState<string | null>(
    ranges[0]?.id ?? null,
  );
  const {
    cancelDraft: resetDraft,
    dirty,
    draft,
    replaceDraft,
    setDraft,
    syncDraft,
  } = useOfficeDraft<NamedRangeDraft>(() => rangeDraft(ranges[0]));
  const [error, setError] = useState('');

  useEffect(() => {
    const selected = ranges.find((range) => range.id === selectedId);
    if (selected) syncDraft(rangeDraft(selected));
    else if (selectedId) {
      setSelectedId(ranges[0]?.id ?? null);
      replaceDraft(rangeDraft(ranges[0]));
    }
  }, [content.namedRanges, replaceDraft, syncDraft]);

  const selectRange = (range: WorkSpreadsheetNamedRange) => {
    if (range.id === selectedId) return;
    if (dirty) {
      setError(officeMessage(messages, 'spreadsheet.workbook.names.error.unsaved'));
      return;
    }
    setSelectedId(range.id);
    replaceDraft(rangeDraft(range));
    setError('');
  };
  const startNewRange = () => {
    if (dirty) {
      setError(officeMessage(messages, 'spreadsheet.workbook.names.error.unsaved'));
      return;
    }
    setSelectedId(null);
    replaceDraft(rangeDraft());
    setError('');
  };
  const saveRange = () => {
    if (draft.id && !dirty) return;
    const name = draft.name.trim();
    const reference = draft.reference.trim().replace(/^=/, '');
    if (!isValidSpreadsheetDefinedName(name)) {
      setError(officeMessage(messages, 'spreadsheet.workbook.names.error.invalidName'));
      return;
    }
    if (!reference) {
      setError(officeMessage(messages, 'spreadsheet.workbook.names.error.refRequired'));
      return;
    }
    const scopeSheetId = draft.scopeSheetId || undefined;
    if (
      ranges.some(
        (range) =>
          range.id !== draft.id &&
          range.name.toLowerCase() === name.toLowerCase() &&
          (range.scopeSheetId ?? '') === (scopeSheetId ?? ''),
      )
    ) {
      setError(officeMessage(messages, 'spreadsheet.workbook.names.error.duplicate'));
      return;
    }
    const saved: WorkSpreadsheetNamedRange = {
      id: draft.id ?? createWorkId('name'),
      name,
      reference,
      scopeSheetId,
      comment: draft.comment.trim() || undefined,
    };
    const next = draft.id
      ? ranges.map((range) => (range.id === draft.id ? saved : range))
      : [...ranges, saved];
    onChange({ ...content, namedRanges: next });
    setSelectedId(saved.id);
    replaceDraft(rangeDraft(saved));
    setError('');
  };
  const cancelDraft = () => {
    resetDraft();
    setError('');
  };
  const deleteRange = () => {
    if (!draft.id) {
      startNewRange();
      return;
    }
    const next = ranges.filter((range) => range.id !== draft.id);
    onChange({ ...content, namedRanges: next.length ? next : undefined });
    const selected = next[0];
    setSelectedId(selected?.id ?? null);
    replaceDraft(rangeDraft(selected));
    setError('');
  };

  return (
    <fieldset
      className="work-spreadsheet-name-manager"
      data-office-escape-consumer={dirty || undefined}
      onKeyDown={(event) => {
        if (event.key !== 'Escape' || event.defaultPrevented || !dirty) return;
        event.preventDefault();
        event.stopPropagation();
        cancelDraft();
      }}
    >
      <legend className="sr-only">{officeMessage(messages, 'spreadsheet.workbook.names.legend')}</legend>
      <aside aria-label={officeMessage(messages, 'spreadsheet.workbook.names.listAria')}>
        <Button className="create" tone="secondary" onClick={startNewRange}>
          <Plus size={13} />
          {officeMessage(messages, 'spreadsheet.workbook.names.new')}
        </Button>
        <div className="work-spreadsheet-name-list">
          {ranges.map((range) => (
            <button
              type="button"
              className={range.id === selectedId ? 'active' : ''}
              key={range.id}
              onClick={() => selectRange(range)}
            >
              <strong>{range.name}</strong>
              <span>
                {range.scopeSheetId
                  ? (content.sheets.find(
                      (sheet) => sheet.id === range.scopeSheetId,
                    )?.name ?? officeMessage(messages, 'spreadsheet.workbook.names.scopeSheet'))
                  : officeMessage(messages, 'spreadsheet.workbook.names.scopeWorkbook')}
              </span>
            </button>
          ))}
          {!ranges.length && (
            <CollectionState
              className="work-office-collection-empty"
              role="status"
            >
              {officeMessage(messages, 'spreadsheet.workbook.names.empty')}
            </CollectionState>
          )}
        </div>
      </aside>
      <form
        onSubmit={(event) => {
          event.preventDefault();
          saveRange();
        }}
      >
        <div className="work-office-field">
          <span>{officeMessage(messages, 'spreadsheet.workbook.names.name')}</span>
          <OfficeTextField
            aria-label={officeMessage(messages, 'spreadsheet.workbook.names.nameAria')}
            value={draft.name}
            maxLength={255}
            placeholder={officeMessage(messages, 'spreadsheet.workbook.names.namePlaceholder')}
            onChange={(event) =>
              setDraft({ ...draft, name: event.target.value })
            }
          />
        </div>
        <div className="work-office-field">
          <span>{officeMessage(messages, 'spreadsheet.workbook.names.scope')}</span>
          <OfficeSelect
            ariaLabel={officeMessage(messages, 'spreadsheet.workbook.names.scopeAria')}
            value={draft.scopeSheetId}
            options={[
              { value: '', label: officeMessage(messages, 'spreadsheet.workbook.names.scopeWorkbook') },
              ...content.sheets.flatMap((sheet) =>
                sheet.id ? [{ value: sheet.id, label: sheet.name }] : [],
              ),
            ]}
            onValueChange={(scopeSheetId) =>
              setDraft({ ...draft, scopeSheetId })
            }
          />
        </div>
        <div className="work-office-field reference">
          <span>{officeMessage(messages, 'spreadsheet.workbook.names.ref')}</span>
          <OfficeTextField
            aria-label={officeMessage(messages, 'spreadsheet.workbook.names.refAria')}
            value={draft.reference}
            placeholder={officeMessage(messages, 'spreadsheet.workbook.names.refPlaceholder')}
            onChange={(event) =>
              setDraft({ ...draft, reference: event.target.value })
            }
          />
        </div>
        <div className="work-office-field comment">
          <span>{officeMessage(messages, 'spreadsheet.workbook.names.comment')}</span>
          <OfficeTextField
            aria-label={officeMessage(messages, 'spreadsheet.workbook.names.commentAria')}
            value={draft.comment}
            maxLength={255}
            placeholder={officeMessage(messages, 'spreadsheet.workbook.names.commentPlaceholder')}
            onChange={(event) =>
              setDraft({ ...draft, comment: event.target.value })
            }
          />
        </div>
        <div className="actions">
          {error && (
            <InlineNotice
              className="work-office-form-error"
              tone="danger"
              role="alert"
            >
              {error}
            </InlineNotice>
          )}
          <Button tone="danger" disabled={!draft.id} onClick={deleteRange}>
            <Trash2 size={13} />
            {officeMessage(messages, 'spreadsheet.workbook.names.delete')}
          </Button>
          <Button tone="secondary" disabled={!dirty} onClick={cancelDraft}>
            {officeMessage(messages, 'spreadsheet.workbook.names.cancel')}
          </Button>
          <Button
            type="submit"
            tone="primary"
            disabled={Boolean(draft.id) && !dirty}
          >
            {officeMessage(messages, 'spreadsheet.workbook.names.save')}
          </Button>
        </div>
      </form>
    </fieldset>
  );
}

function rangeDraft(range?: WorkSpreadsheetNamedRange): NamedRangeDraft {
  return {
    id: range?.id,
    name: range?.name ?? '',
    reference: range?.reference ?? '',
    scopeSheetId: range?.scopeSheetId ?? '',
    comment: range?.comment ?? '',
  };
}
