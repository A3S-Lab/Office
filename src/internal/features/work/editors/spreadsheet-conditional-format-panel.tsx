import type { Sheet } from '@fortune-sheet/core';
import { ArrowDown, ArrowUp, Plus, Trash2 } from 'lucide-react';
import { useEffect, useState } from 'react';
import { officeMessage } from '../../../i18n/office-locale';
import {
  Button,
  CollectionState,
  InlineNotice,
  StateView,
} from '../../../design-system/primitives';
import {
  defaultSpreadsheetConditionalIconThresholds,
  SPREADSHEET_CONDITIONAL_ICON_SETS,
  spreadsheetConditionalIconSetLabel,
  type SpreadsheetConditionalIconSetName,
} from '../work-spreadsheet-conditional-icons';
import { defaultSpreadsheetColorScaleThresholds } from '../work-spreadsheet-conditional-values';
import { formatSpreadsheetCellRanges } from '../work-spreadsheet-ranges';
import type { WorkSpreadsheetContent } from '../work-types';
import type { FortuneConditionalFormatRule } from '../work-xlsx-conditional-format';
import {
  OfficeCheckbox,
  OfficeColorPicker,
  OfficeNumberField,
  OfficeSelect,
  OfficeTextField,
} from './office-controls';
import { useOfficeMessages } from './office-messages-context';
import { SpreadsheetConditionalComparisonFields } from './spreadsheet-conditional-comparison-fields';
import {
  buildConditionalRule,
  type ConditionalRuleDraft,
  conditionalRuleDraftForRule,
  conditionalRuleLabel,
  conditionalThresholdDrafts,
  conditionalToolbarRuleSummary,
  isManagedConditionalRule,
  managedConditionalRuleEntries,
  newConditionalRuleDraft,
  type SpreadsheetConditionalThresholdDraft,
  sheetConditionalRules,
  withConditionalRules,
} from './spreadsheet-conditional-format-model';
import { SpreadsheetConditionalThresholdFields } from './spreadsheet-conditional-threshold-fields';
import { useOfficeDraft } from './use-office-draft';
import { MAX_SPREADSHEET_LOCAL_FORMULA_LENGTH } from '../work-spreadsheet-local-formula';

export { managedConditionalFormatCount } from './spreadsheet-conditional-format-model';

interface SpreadsheetConditionalFormatPanelProps {
  content: WorkSpreadsheetContent;
  onChange: (content: WorkSpreadsheetContent) => void;
}

interface ConditionalRuleLocation {
  sheetId: string;
  index: number;
}

export function SpreadsheetConditionalFormatPanel({
  content,
  onChange,
}: SpreadsheetConditionalFormatPanelProps) {
  const messages = useOfficeMessages();
  const sheets = content.sheets.filter(
    (sheet): sheet is Sheet & { id: string } => Boolean(sheet.id),
  );
  const activeSheetId =
    sheets.find((sheet) => sheet.status === 1)?.id ?? sheets[0]?.id ?? '';
  const entries = managedConditionalRuleEntries(sheets);
  const [selection, setSelection] = useState<ConditionalRuleLocation | null>(
    entries[0]
      ? { sheetId: entries[0].sheet.id, index: entries[0].index }
      : null,
  );
  const {
    cancelDraft: resetDraft,
    dirty,
    draft,
    replaceDraft,
    setDraft,
    syncDraft,
  } = useOfficeDraft<ConditionalRuleDraft>(() =>
    entries[0]
      ? conditionalRuleDraftForRule(entries[0].sheet.id, entries[0].rule)
      : newConditionalRuleDraft(activeSheetId),
  );
  const [error, setError] = useState('');
  const selectedSheet = selection
    ? sheets.find((sheet) => sheet.id === selection.sheetId)
    : undefined;

  useEffect(() => {
    if (!selection) {
      const next = newConditionalRuleDraft(activeSheetId);
      syncDraft(next);
      return;
    }
    const selectedSheet = content.sheets.find(
      (sheet) => sheet.id === selection.sheetId,
    );
    const rule = selectedSheet
      ? sheetConditionalRules(selectedSheet)[selection.index]
      : undefined;
    if (isManagedConditionalRule(rule)) {
      const next = conditionalRuleDraftForRule(selection.sheetId, rule);
      syncDraft(next);
      return;
    }
    const next = newConditionalRuleDraft(activeSheetId);
    setSelection(null);
    replaceDraft(next);
  }, [activeSheetId, content.sheets, replaceDraft, syncDraft]);

  const startNew = () => {
    if (dirty) {
      setError(officeMessage(messages, 'spreadsheet.cf.error.unsaved'));
      return;
    }
    const next = newConditionalRuleDraft(activeSheetId);
    setSelection(null);
    replaceDraft(next);
    setError('');
  };
  const selectRule = (
    location: ConditionalRuleLocation,
    rule: FortuneConditionalFormatRule,
  ) => {
    if (
      selection?.sheetId === location.sheetId &&
      selection.index === location.index
    ) {
      return;
    }
    if (dirty) {
      setError(officeMessage(messages, 'spreadsheet.cf.error.unsaved'));
      return;
    }
    const next = conditionalRuleDraftForRule(location.sheetId, rule);
    setSelection(location);
    replaceDraft(next);
    setError('');
  };
  const setIconSet = (iconSet: SpreadsheetConditionalIconSetName) => {
    setDraft({
      ...draft,
      iconSet,
      iconThresholds: conditionalThresholdDrafts(
        defaultSpreadsheetConditionalIconThresholds(iconSet),
      ),
    });
  };
  const updateThreshold = (
    field: 'scaleThresholds' | 'barThresholds' | 'iconThresholds',
    index: number,
    patch: Partial<SpreadsheetConditionalThresholdDraft>,
  ) => {
    const thresholds = draft[field].map((threshold, thresholdIndex) =>
      thresholdIndex === index ? { ...threshold, ...patch } : threshold,
    );
    setDraft({ ...draft, [field]: thresholds });
  };
  const saveRule = () => {
    if (selection && !dirty) return;
    const result = buildConditionalRule(draft);
    if ('error' in result) {
      setError(result.error);
      return;
    }
    const rule = result.rule;
    let savedIndex = -1;
    const selectedRule = selection;
    const nextSheets = content.sheets.map((sheet) => {
      let rules = sheetConditionalRules(sheet);
      if (selectedRule && selectedRule.sheetId === sheet.id) {
        rules = [...rules];
        if (selectedRule.sheetId === draft.sheetId) {
          rules[selectedRule.index] = rule;
          savedIndex = selectedRule.index;
        } else {
          rules.splice(selectedRule.index, 1);
        }
      }
      if (
        sheet.id === draft.sheetId &&
        selectedRule?.sheetId !== draft.sheetId
      ) {
        rules = [...rules, rule];
        savedIndex = rules.length - 1;
      }
      return withConditionalRules(sheet, rules);
    });
    if (savedIndex < 0) {
      setError(officeMessage(messages, 'spreadsheet.cf.error.invalidSheet'));
      return;
    }
    onChange({ ...content, sheets: nextSheets });
    const savedDraft = conditionalRuleDraftForRule(draft.sheetId, rule);
    setSelection({ sheetId: draft.sheetId, index: savedIndex });
    replaceDraft(savedDraft);
    setError('');
  };
  const cancelDraft = () => {
    resetDraft();
    setError('');
  };
  const deleteRule = () => {
    if (!selection) return;
    const nextSheets = content.sheets.map((sheet) => {
      if (sheet.id !== selection.sheetId) return sheet;
      const rules = sheetConditionalRules(sheet).filter(
        (_, index) => index !== selection.index,
      );
      return withConditionalRules(sheet, rules);
    });
    onChange({ ...content, sheets: nextSheets });
    const next = newConditionalRuleDraft(activeSheetId);
    setSelection(null);
    replaceDraft(next);
    setError('');
  };
  const moveRule = (offset: -1 | 1) => {
    if (!selection || dirty) return;
    let movedIndex = selection.index;
    const nextSheets = content.sheets.map((sheet) => {
      if (sheet.id !== selection.sheetId) return sheet;
      const rules = [...sheetConditionalRules(sheet)];
      const target = selection.index + offset;
      if (target < 0 || target >= rules.length) return sheet;
      [rules[selection.index], rules[target]] = [
        rules[target],
        rules[selection.index],
      ];
      movedIndex = target;
      return withConditionalRules(sheet, rules);
    });
    if (movedIndex === selection.index) return;
    onChange({ ...content, sheets: nextSheets });
    setSelection({ ...selection, index: movedIndex });
    setError('');
  };

  const selectedRuleCount = selectedSheet
    ? sheetConditionalRules(selectedSheet).length
    : 0;

  if (!sheets.length) {
    return (
      <StateView
        className="work-office-panel-empty work-spreadsheet-conditional-empty"
        size="compact"
        title={officeMessage(messages, 'spreadsheet.cf.emptySheets.title')}
      />
    );
  }
  return (
    <fieldset
      className="work-spreadsheet-conditional-manager"
      data-office-escape-consumer={dirty || undefined}
      onKeyDown={(event) => {
        if (event.key !== 'Escape' || event.defaultPrevented || !dirty) return;
        event.preventDefault();
        event.stopPropagation();
        cancelDraft();
      }}
    >
      <legend className="sr-only">{officeMessage(messages, 'spreadsheet.cf.legend')}</legend>
      <aside aria-label={officeMessage(messages, 'spreadsheet.cf.rulesAria')}>
        <Button className="create" tone="secondary" onClick={startNew}>
          <Plus size={13} />
          {officeMessage(messages, 'spreadsheet.cf.newRule')}
        </Button>
        <div className="work-spreadsheet-conditional-list">
          {entries.map(({ sheet, rule, index }) => {
            const selected =
              selection?.sheetId === sheet.id && selection.index === index;
            return (
              <button
                type="button"
                className={selected ? 'active' : ''}
                key={`${sheet.id}-${index}`}
                onClick={() => selectRule({ sheetId: sheet.id, index }, rule)}
              >
                <strong>{conditionalRuleLabel(rule)}</strong>
                <span>{sheet.name}</span>
                <small>{formatSpreadsheetCellRanges(rule.cellrange)}</small>
              </button>
            );
          })}
          {!entries.length && (
            <CollectionState
              className="work-office-collection-empty"
              role="status"
            >
              {officeMessage(messages, 'spreadsheet.cf.emptyRules')}
            </CollectionState>
          )}
        </div>
      </aside>
      <form
        onSubmit={(event) => {
          event.preventDefault();
          saveRule();
        }}
      >
        <div className="work-office-field">
          <span>{officeMessage(messages, 'spreadsheet.cf.sheet')}</span>
          <OfficeSelect
            ariaLabel={officeMessage(messages, 'spreadsheet.cf.sheetAria')}
            value={draft.sheetId}
            options={sheets.map((sheet) => ({
              value: sheet.id,
              label: sheet.name,
            }))}
            onValueChange={(sheetId) => setDraft({ ...draft, sheetId })}
          />
        </div>
        <div className="work-office-field">
          <span>{officeMessage(messages, 'spreadsheet.cf.ruleType')}</span>
          <OfficeSelect
            ariaLabel={officeMessage(messages, 'spreadsheet.cf.ruleTypeAria')}
            value={draft.type}
            disabled={draft.type === 'toolbarRule'}
            options={[
              { value: 'cellComparison', label: officeMessage(messages, 'spreadsheet.cf.type.cellComparison') },
              { value: 'colorGradation', label: officeMessage(messages, 'spreadsheet.cf.type.colorGradation') },
              { value: 'dataBar', label: officeMessage(messages, 'spreadsheet.cf.type.dataBar') },
              { value: 'icons', label: officeMessage(messages, 'spreadsheet.cf.type.icons') },
              { value: 'formula', label: officeMessage(messages, 'spreadsheet.cf.type.formula') },
              ...(draft.type === 'toolbarRule'
                ? [{ value: 'toolbarRule' as const, label: officeMessage(messages, 'spreadsheet.cf.type.toolbarRule') }]
                : []),
            ]}
            onValueChange={(type) =>
              setDraft({ ...draft, type: type as ConditionalRuleDraft['type'] })
            }
          />
        </div>
        <div className="work-office-field reference">
          <span>{officeMessage(messages, 'spreadsheet.cf.range')}</span>
          <OfficeTextField
            aria-label={officeMessage(messages, 'spreadsheet.cf.rangeAria')}
            value={draft.reference}
            placeholder="A2:A20"
            onChange={(event) =>
              setDraft({ ...draft, reference: event.target.value })
            }
          />
        </div>
        <OfficeCheckbox
          className="toggle"
          ariaLabel={officeMessage(messages, 'spreadsheet.cf.stopIfTrue')}
          checked={draft.stopIfTrue}
          onCheckedChange={(stopIfTrue) => setDraft({ ...draft, stopIfTrue })}
        >
          {officeMessage(messages, 'spreadsheet.cf.stopIfTrue')}
        </OfficeCheckbox>
        {draft.type === 'toolbarRule' ? (
          <div className="work-office-field reference">
            <span>{officeMessage(messages, 'spreadsheet.cf.summary')}</span>
            <OfficeTextField
              aria-label={officeMessage(messages, 'spreadsheet.cf.summaryAria')}
              readOnly
              value={conditionalToolbarRuleSummary(draft.preservedRule)}
            />
          </div>
        ) : draft.type === 'cellComparison' ? (
          <SpreadsheetConditionalComparisonFields
            draft={draft}
            onChange={(patch) => setDraft({ ...draft, ...patch })}
          />
        ) : draft.type === 'formula' ? (
          <>
            <div className="work-office-field reference">
              <span>{officeMessage(messages, 'spreadsheet.cf.formula')}</span>
              <OfficeTextField
                aria-label={officeMessage(messages, 'spreadsheet.cf.formulaAria')}
                value={draft.formula}
                maxLength={MAX_SPREADSHEET_LOCAL_FORMULA_LENGTH + 1}
                placeholder="=A2>0"
                onChange={(event) =>
                  setDraft({ ...draft, formula: event.target.value })
                }
              />
              <small>
                {officeMessage(messages, 'spreadsheet.cf.formula.anchorNote')}{' '}
                {Array.from(draft.formula.replace(/^=/, '')).length}/
                {MAX_SPREADSHEET_LOCAL_FORMULA_LENGTH}
              </small>
            </div>
            <p className="work-office-field-hint">
              {officeMessage(messages, 'spreadsheet.cf.formula.supportNote')}
            </p>
            <OfficeCheckbox
              className="toggle"
              ariaLabel={officeMessage(messages, 'spreadsheet.cf.formula.useTextColorAria')}
              checked={draft.formulaUseTextColor}
              onCheckedChange={(formulaUseTextColor) =>
                setDraft({ ...draft, formulaUseTextColor })
              }
            >
              {officeMessage(messages, 'spreadsheet.cf.formula.useTextColor')}
            </OfficeCheckbox>
            {draft.formulaUseTextColor && (
              <ColorField
                label={officeMessage(messages, 'spreadsheet.cf.formula.textColor')}
                value={draft.formulaTextColor}
                onChange={(formulaTextColor) =>
                  setDraft({ ...draft, formulaTextColor })
                }
              />
            )}
            <OfficeCheckbox
              className="toggle"
              ariaLabel={officeMessage(messages, 'spreadsheet.cf.formula.useFillColorAria')}
              checked={draft.formulaUseCellColor}
              onCheckedChange={(formulaUseCellColor) =>
                setDraft({ ...draft, formulaUseCellColor })
              }
            >
              {officeMessage(messages, 'spreadsheet.cf.formula.useFillColor')}
            </OfficeCheckbox>
            {draft.formulaUseCellColor && (
              <ColorField
                label={officeMessage(messages, 'spreadsheet.cf.formula.fillColor')}
                value={draft.formulaCellColor}
                onChange={(formulaCellColor) =>
                  setDraft({ ...draft, formulaCellColor })
                }
              />
            )}
          </>
        ) : draft.type === 'colorGradation' ? (
          <>
            <div className="work-office-field">
              <span>{officeMessage(messages, 'spreadsheet.cf.scale.steps')}</span>
              <OfficeSelect
                ariaLabel={officeMessage(messages, 'spreadsheet.cf.scale.stepsAria')}
                value={draft.scaleSize}
                options={[
                  { value: '2', label: officeMessage(messages, 'spreadsheet.cf.scale.two') },
                  { value: '3', label: officeMessage(messages, 'spreadsheet.cf.scale.three') },
                ]}
                onValueChange={(scaleSize) =>
                  setDraft({
                    ...draft,
                    scaleSize: scaleSize as ConditionalRuleDraft['scaleSize'],
                    scaleThresholds: conditionalThresholdDrafts(
                      defaultSpreadsheetColorScaleThresholds(Number(scaleSize)),
                    ),
                  })
                }
              />
            </div>
            <ColorField
              label={officeMessage(messages, 'spreadsheet.cf.scale.minColor')}
              value={draft.minimumColor}
              onChange={(minimumColor) => setDraft({ ...draft, minimumColor })}
            />
            {draft.scaleSize === '3' && (
              <ColorField
                label={officeMessage(messages, 'spreadsheet.cf.scale.midColor')}
                value={draft.midpointColor}
                onChange={(midpointColor) =>
                  setDraft({ ...draft, midpointColor })
                }
              />
            )}
            <ColorField
              label={officeMessage(messages, 'spreadsheet.cf.scale.maxColor')}
              value={draft.maximumColor}
              onChange={(maximumColor) => setDraft({ ...draft, maximumColor })}
            />
            <SpreadsheetConditionalThresholdFields
              label={officeMessage(messages, 'spreadsheet.cf.scale.label')}
              thresholds={draft.scaleThresholds}
              onChange={(index, patch) =>
                updateThreshold('scaleThresholds', index, patch)
              }
            />
          </>
        ) : draft.type === 'dataBar' ? (
          <>
            <ColorField
              label={officeMessage(messages, 'spreadsheet.cf.bar.color')}
              value={draft.barColor}
              onChange={(barColor) => setDraft({ ...draft, barColor })}
            />
            <OfficeCheckbox
              className="toggle"
              ariaLabel={officeMessage(messages, 'spreadsheet.cf.bar.showValue')}
              checked={draft.barShowValue}
              onCheckedChange={(barShowValue) =>
                setDraft({ ...draft, barShowValue })
              }
            >
              {officeMessage(messages, 'spreadsheet.cf.bar.showValue')}
            </OfficeCheckbox>
            <div className="work-office-field">
              <span>{officeMessage(messages, 'spreadsheet.cf.bar.minLen')}</span>
              <OfficeNumberField
                min={0}
                max={100}
                ariaLabel={officeMessage(messages, 'spreadsheet.cf.bar.minLenAria')}
                value={draft.barMinLength}
                onValueChange={(barMinLength) =>
                  setDraft({ ...draft, barMinLength })
                }
              />
            </div>
            <div className="work-office-field">
              <span>{officeMessage(messages, 'spreadsheet.cf.bar.maxLen')}</span>
              <OfficeNumberField
                min={0}
                max={100}
                ariaLabel={officeMessage(messages, 'spreadsheet.cf.bar.maxLenAria')}
                value={draft.barMaxLength}
                onValueChange={(barMaxLength) =>
                  setDraft({ ...draft, barMaxLength })
                }
              />
            </div>
            <SpreadsheetConditionalThresholdFields
              label={officeMessage(messages, 'spreadsheet.cf.type.dataBar')}
              thresholds={draft.barThresholds}
              onChange={(index, patch) =>
                updateThreshold('barThresholds', index, patch)
              }
            />
          </>
        ) : (
          <>
            <div className="work-office-field">
              <span>
                {officeMessage(messages, 'spreadsheet.cf.type.icons')}
              </span>
              <OfficeSelect
                ariaLabel={officeMessage(messages, 'spreadsheet.cf.type.icons')}
                value={draft.iconSet}
                options={SPREADSHEET_CONDITIONAL_ICON_SETS.map((iconSet) => ({
                  value: iconSet.name,
                  label: spreadsheetConditionalIconSetLabel(iconSet.name),
                }))}
                onValueChange={(iconSet) =>
                  setIconSet(iconSet as SpreadsheetConditionalIconSetName)
                }
              />
            </div>
            <OfficeCheckbox
              className="toggle"
              ariaLabel={officeMessage(messages, 'spreadsheet.cf.icon.reverse')}
              checked={draft.iconReverse}
              onCheckedChange={(iconReverse) =>
                setDraft({ ...draft, iconReverse })
              }
            >
              {officeMessage(messages, 'spreadsheet.cf.icon.reverse')}
            </OfficeCheckbox>
            <OfficeCheckbox
              className="toggle"
              ariaLabel={officeMessage(messages, 'spreadsheet.cf.icon.showValue')}
              checked={draft.iconShowValue}
              onCheckedChange={(iconShowValue) =>
                setDraft({ ...draft, iconShowValue })
              }
            >
              {officeMessage(messages, 'spreadsheet.cf.icon.showValue')}
            </OfficeCheckbox>
            <SpreadsheetConditionalThresholdFields
              label={officeMessage(messages, 'spreadsheet.cf.icon.set')}
              thresholds={draft.iconThresholds}
              startIndex={1}
              showEquality
              onChange={(index, patch) =>
                updateThreshold('iconThresholds', index, patch)
              }
            />
          </>
        )}
        {draft.type === 'toolbarRule' && (
          <p>{officeMessage(messages, 'spreadsheet.cf.toolbar.readonly')}</p>
        )}
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
          <Button
            tone="secondary"
            disabled={dirty || !selection || selection.index <= 0}
            aria-label={officeMessage(messages, 'spreadsheet.cf.moveUp')}
            onClick={() => moveRule(-1)}
          >
            <ArrowUp size={13} />
            {officeMessage(messages, 'spreadsheet.cf.moveUp')}
          </Button>
          <Button
            tone="secondary"
            disabled={
              dirty || !selection || selection.index >= selectedRuleCount - 1
            }
            aria-label={officeMessage(messages, 'spreadsheet.cf.moveDown')}
            onClick={() => moveRule(1)}
          >
            <ArrowDown size={13} />
            {officeMessage(messages, 'spreadsheet.cf.moveDown')}
          </Button>
          <Button tone="danger" disabled={!selection} onClick={deleteRule}>
            <Trash2 size={13} />
            {officeMessage(messages, 'spreadsheet.cf.delete')}
          </Button>
          <Button tone="secondary" disabled={!dirty} onClick={cancelDraft}>
            {officeMessage(messages, 'spreadsheet.cf.cancel')}
          </Button>
          <Button
            type="submit"
            tone="primary"
            disabled={Boolean(selection) && !dirty}
          >
            {officeMessage(messages, 'spreadsheet.cf.save')}
          </Button>
        </div>
      </form>
    </fieldset>
  );
}

function ColorField({
  label,
  value,
  onChange,
}: {
  label: string;
  value: string;
  onChange: (value: string) => void;
}) {
  return (
    <div className="work-office-field color">
      <span>{label}</span>
      <OfficeColorPicker
        ariaLabel={label}
        value={value}
        onValueChange={onChange}
      />
    </div>
  );
}
