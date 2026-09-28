import { Calculator, RefreshCw, Save } from 'lucide-react';
import { useEffect, useMemo, useState } from 'react';
import { officeMessage } from '../../../i18n/office-locale';
import { Button, CollectionState } from '../../../design-system/primitives';
import { spreadsheetFormulaAnalysis } from '../work-spreadsheet-formula-analysis';
import { effectiveSpreadsheetCalculationSettings } from '../work-spreadsheet-formulas';
import type {
  WorkSpreadsheetCalculationSettings,
  WorkSpreadsheetContent,
} from '../work-types';
import {
  CommittedOfficeNumberField,
  OfficeCheckbox,
  OfficeSelect,
} from './office-controls';
import { useOfficeMessages } from './office-messages-context';
import { useOfficeDraft } from './use-office-draft';

interface SpreadsheetFormulaPanelProps {
  content: WorkSpreadsheetContent;
  canRecalculateSelection: boolean;
  canRecalculateWorkbook: boolean;
  onChange: (content: WorkSpreadsheetContent) => void;
  onRecalculate: (scope: 'workbook' | 'selection') => boolean;
}

export function SpreadsheetFormulaPanel({
  content,
  canRecalculateSelection,
  canRecalculateWorkbook,
  onChange,
  onRecalculate,
}: SpreadsheetFormulaPanelProps) {
  const messages = useOfficeMessages();
  const {
    cancelDraft,
    dirty,
    draft: settings,
    replaceDraft,
    setDraft: setSettings,
    syncDraft,
  } = useOfficeDraft(() =>
    effectiveSpreadsheetCalculationSettings(content.calculation),
  );
  const [status, setStatus] = useState('');
  const sourceSettings = useMemo(
    () => effectiveSpreadsheetCalculationSettings(content.calculation),
    [content.calculation],
  );
  const { summary, diagnostics } = useMemo(
    () => spreadsheetFormulaAnalysis(content),
    [content],
  );

  useEffect(() => {
    syncDraft(sourceSettings);
  }, [content.calculation, sourceSettings, syncDraft]);

  const update = <Key extends keyof WorkSpreadsheetCalculationSettings>(
    key: Key,
    value: WorkSpreadsheetCalculationSettings[Key],
  ) => {
    setSettings((current) => ({ ...current, [key]: value }));
    setStatus('');
  };
  const save = () => {
    if (!dirty) return;
    const saved = effectiveSpreadsheetCalculationSettings(settings);
    onChange({
      ...content,
      calculation: saved,
    });
    replaceDraft(saved);
    setStatus(officeMessage(messages, 'spreadsheet.formula.status.saved'));
  };
  const cancel = () => {
    cancelDraft();
    setStatus('');
  };
  const recalculate = (scope: 'workbook' | 'selection') => {
    const started = onRecalculate(scope);
    setStatus(
      started
        ? scope === 'workbook'
          ? officeMessage(messages, 'spreadsheet.formula.status.recalcWorkbook')
          : officeMessage(messages, 'spreadsheet.formula.status.recalcSelection')
        : officeMessage(messages, 'spreadsheet.formula.status.notReady'),
    );
  };

  return (
    <fieldset
      className="work-spreadsheet-formula-manager"
      data-office-escape-consumer={dirty || undefined}
      onKeyDown={(event) => {
        if (event.key !== 'Escape' || event.defaultPrevented || !dirty) return;
        event.preventDefault();
        event.stopPropagation();
        cancel();
      }}
    >
      <legend className="sr-only">{officeMessage(messages, 'spreadsheet.formula.legend')}</legend>
      <aside aria-label={officeMessage(messages, 'spreadsheet.formula.statsAria')}>
        <FormulaStat label={officeMessage(messages, 'spreadsheet.formula.stat.formulaCells')} value={summary.formulaCells} />
        <FormulaStat
          label={officeMessage(messages, 'spreadsheet.formula.stat.cachedErrors')}
          value={summary.cachedErrorCells}
          tone={summary.cachedErrorCells ? 'error' : undefined}
        />
        <FormulaStat label={officeMessage(messages, 'spreadsheet.formula.stat.arrayRanges')} value={summary.arrayRanges} />
        <FormulaStat label={officeMessage(messages, 'spreadsheet.formula.stat.dynamicArrays')} value={summary.dynamicArrayRanges} />
        <FormulaStat label={officeMessage(messages, 'spreadsheet.formula.stat.dataTables')} value={summary.dataTableRanges} />
        <FormulaStat
          label={officeMessage(messages, 'spreadsheet.formula.stat.compat')}
          value={diagnostics.filter((item) => item.severity !== 'info').length}
          tone={
            diagnostics.some((item) => item.severity === 'error')
              ? 'error'
              : undefined
          }
        />
      </aside>
      <form
        onSubmit={(event) => {
          event.preventDefault();
          save();
        }}
      >
        <div className="work-spreadsheet-calculation-fields">
          <div className="work-office-field">
            <span>{officeMessage(messages, 'spreadsheet.formula.calcMode')}</span>
            <OfficeSelect
              ariaLabel={officeMessage(messages, 'spreadsheet.formula.calcMode')}
              value={settings.mode}
              options={[
                { value: 'automatic', label: officeMessage(messages, 'spreadsheet.formula.mode.automatic') },
                {
                  value: 'automatic-except-data-tables',
                  label: officeMessage(messages, 'spreadsheet.formula.mode.automaticExceptTables'),
                },
                { value: 'manual', label: officeMessage(messages, 'spreadsheet.formula.mode.manual') },
              ]}
              onValueChange={(mode) =>
                update(
                  'mode',
                  mode as WorkSpreadsheetCalculationSettings['mode'],
                )
              }
            />
          </div>
          <div className="work-office-field">
            <span>{officeMessage(messages, 'spreadsheet.formula.maxIterations')}</span>
            <CommittedOfficeNumberField
              ariaLabel={officeMessage(messages, 'spreadsheet.formula.maxIterations')}
              min={1}
              max={10_000}
              step={1}
              disabled={!settings.iterativeCalculation}
              value={settings.maximumIterations}
              normalizeValue={(value) =>
                normalizeCalculationInteger(value, 1, 10_000)
              }
              onValueCommit={(value) => update('maximumIterations', value)}
            />
          </div>
          <div className="work-office-field">
            <span>{officeMessage(messages, 'spreadsheet.formula.maxChange')}</span>
            <CommittedOfficeNumberField
              ariaLabel={officeMessage(messages, 'spreadsheet.formula.maxChange')}
              min={0.000000000001}
              step={0.000001}
              disabled={!settings.iterativeCalculation}
              value={settings.maximumChange}
              normalizeValue={normalizeCalculationChange}
              onValueCommit={(value) => update('maximumChange', value)}
            />
          </div>
          <OfficeCheckbox
            className="toggle"
            ariaLabel={officeMessage(messages, 'spreadsheet.formula.fullCalcOnLoad')}
            checked={settings.fullCalculationOnLoad}
            onCheckedChange={(checked) =>
              update('fullCalculationOnLoad', checked)
            }
          >
            {officeMessage(messages, 'spreadsheet.formula.fullCalcOnLoad')}
          </OfficeCheckbox>
          <OfficeCheckbox
            className="toggle"
            ariaLabel={officeMessage(messages, 'spreadsheet.formula.forceFullCalc')}
            checked={settings.forceFullCalculation}
            onCheckedChange={(checked) =>
              update('forceFullCalculation', checked)
            }
          >
            {officeMessage(messages, 'spreadsheet.formula.forceFullCalc')}
          </OfficeCheckbox>
          <OfficeCheckbox
            className="toggle"
            ariaLabel={officeMessage(messages, 'spreadsheet.formula.iterate')}
            checked={settings.iterativeCalculation}
            onCheckedChange={(checked) =>
              update('iterativeCalculation', checked)
            }
          >
            {officeMessage(messages, 'spreadsheet.formula.iterate')}
          </OfficeCheckbox>
          <OfficeCheckbox
            className="toggle"
            ariaLabel={officeMessage(messages, 'spreadsheet.formula.fullPrecision')}
            checked={settings.fullPrecision}
            onCheckedChange={(checked) => update('fullPrecision', checked)}
          >
            {officeMessage(messages, 'spreadsheet.formula.fullPrecision')}
          </OfficeCheckbox>
          <div className="actions">
            {status && <span className="status">{status}</span>}
            <Button
              tone="secondary"
              disabled={!canRecalculateSelection}
              onClick={() => recalculate('selection')}
            >
              <RefreshCw size={12} />
              {officeMessage(messages, 'spreadsheet.formula.recalcSelection')}
            </Button>
            <Button
              tone="secondary"
              disabled={!canRecalculateWorkbook}
              onClick={() => recalculate('workbook')}
            >
              <Calculator size={12} />
              {officeMessage(messages, 'spreadsheet.formula.recalcWorkbook')}
            </Button>
            <Button tone="secondary" disabled={!dirty} onClick={cancel}>
              {officeMessage(messages, 'spreadsheet.formula.cancel')}
            </Button>
            <Button type="submit" tone="primary" disabled={!dirty}>
              <Save size={12} />
              {officeMessage(messages, 'spreadsheet.formula.save')}
            </Button>
          </div>
        </div>
        <section
          className="work-spreadsheet-formula-diagnostics"
          aria-label={officeMessage(messages, 'spreadsheet.formula.diagnosticsAria')}
        >
          <header>
            <strong>{officeMessage(messages, 'spreadsheet.formula.diagnosticsTitle')}</strong>
            <span>
              {diagnostics.length
                ? officeMessage(messages, 'spreadsheet.formula.diagnosticsCount', {
                    n: String(diagnostics.length),
                  })
                : officeMessage(messages, 'spreadsheet.formula.diagnosticsNone')}
            </span>
          </header>
          <div>
            {diagnostics.map((diagnostic) => (
              <article className={diagnostic.severity} key={diagnostic.code}>
                <strong>{diagnostic.title}</strong>
                <p>{diagnostic.message}</p>
                {diagnostic.locations.length > 0 && (
                  <small>
                    {diagnostic.locations.slice(0, 4).join('、')}
                    {diagnostic.locations.length > 4
                      ? officeMessage(messages, 'spreadsheet.formula.diagnosticsMore', {
                          n: String(diagnostic.locations.length),
                        })
                      : ''}
                  </small>
                )}
              </article>
            ))}
            {!diagnostics.length && (
              <CollectionState
                className="work-office-collection-empty"
                role="status"
              >
                {officeMessage(messages, 'spreadsheet.formula.diagnosticsEmpty')}
              </CollectionState>
            )}
          </div>
        </section>
      </form>
    </fieldset>
  );
}

function normalizeCalculationInteger(
  value: string,
  minimum: number,
  maximum: number,
): number | null {
  if (!value.trim()) return null;
  const number = Number(value);
  return Number.isFinite(number)
    ? Math.min(maximum, Math.max(minimum, Math.round(number)))
    : null;
}

function normalizeCalculationChange(value: string): number | null {
  if (!value.trim()) return null;
  const number = Number(value);
  return Number.isFinite(number) && number > 0
    ? Math.max(0.000000000001, number)
    : null;
}

function FormulaStat({
  label,
  value,
  tone,
}: {
  label: string;
  value: number;
  tone?: 'error';
}) {
  return (
    <div className={tone}>
      <strong>{value}</strong>
      <span>{label}</span>
    </div>
  );
}
