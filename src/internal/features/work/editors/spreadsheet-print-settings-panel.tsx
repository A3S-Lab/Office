import { useEffect, useState } from 'react';
import { officeMessage } from '../../../i18n/office-locale';
import { Button, InlineNotice } from '../../../design-system/primitives';
import {
  formatSpreadsheetColumnPageBreaks,
  formatSpreadsheetRowPageBreaks,
  parseSpreadsheetColumnPageBreaks,
  parseSpreadsheetRowPageBreaks,
} from '../work-spreadsheet-page-breaks';
import {
  type EffectiveSpreadsheetPageSetup,
  effectiveSpreadsheetPageSetup,
} from '../work-spreadsheet-page-setup';
import {
  normalizeSpreadsheetPrintArea,
  normalizeSpreadsheetPrintTitleColumns,
  normalizeSpreadsheetPrintTitleRows,
} from '../work-spreadsheet-ranges';
import type { WorkSpreadsheetContent } from '../work-types';
import {
  CommittedOfficeNumberField,
  OfficeCheckbox,
  OfficeSelect,
  OfficeTextField,
} from './office-controls';
import { useOfficeMessages } from './office-messages-context';
import { SpreadsheetHeaderFooterFields } from './spreadsheet-header-footer-fields';
import { useOfficeDraft } from './use-office-draft';

interface SpreadsheetPrintSettingsDraft {
  reference: string;
  titleRows: string;
  titleColumns: string;
  rowPageBreaks: string;
  columnPageBreaks: string;
  pageSetup: EffectiveSpreadsheetPageSetup;
}

export function spreadsheetPrintSettingCount(
  content: WorkSpreadsheetContent,
): number {
  return new Set([
    ...(content.printAreas ?? []).map((area) => area.sheetId),
    ...(content.printTitles ?? []).map((titles) => titles.sheetId),
    ...(content.pageBreaks ?? []).map((pageBreaks) => pageBreaks.sheetId),
    ...(content.pageSetups ?? []).map((pageSetup) => pageSetup.sheetId),
  ]).size;
}

export function SpreadsheetPrintSettingsPanel({
  content,
  onChange,
}: {
  content: WorkSpreadsheetContent;
  onChange: (content: WorkSpreadsheetContent) => void;
}) {
  const messages = useOfficeMessages();
  const availableSheets = content.sheets.filter(
    (sheet): sheet is typeof sheet & { id: string } => Boolean(sheet.id),
  );
  const initialSheet =
    availableSheets.find((sheet) => sheet.status === 1)?.id ??
    availableSheets[0]?.id ??
    '';
  const [sheetId, setSheetId] = useState(initialSheet);
  const savedArea = content.printAreas?.find(
    (area) => area.sheetId === sheetId,
  );
  const savedTitles = content.printTitles?.find(
    (titles) => titles.sheetId === sheetId,
  );
  const savedPageBreaks = content.pageBreaks?.find(
    (pageBreaks) => pageBreaks.sheetId === sheetId,
  );
  const savedPageSetup = content.pageSetups?.find(
    (pageSetup) => pageSetup.sheetId === sheetId,
  );
  const {
    cancelDraft: resetDraft,
    dirty,
    draft,
    replaceDraft,
    setDraft,
    syncDraft,
  } = useOfficeDraft<SpreadsheetPrintSettingsDraft>(() =>
    spreadsheetPrintSettingsDraft(content, sheetId),
  );
  const {
    columnPageBreaks,
    pageSetup,
    reference,
    rowPageBreaks,
    titleColumns,
    titleRows,
  } = draft;
  const setReference = (value: string) =>
    setDraft((current) => ({ ...current, reference: value }));
  const setTitleRows = (value: string) =>
    setDraft((current) => ({ ...current, titleRows: value }));
  const setTitleColumns = (value: string) =>
    setDraft((current) => ({ ...current, titleColumns: value }));
  const setRowPageBreaks = (value: string) =>
    setDraft((current) => ({ ...current, rowPageBreaks: value }));
  const setColumnPageBreaks = (value: string) =>
    setDraft((current) => ({ ...current, columnPageBreaks: value }));
  const setPageSetup = (value: EffectiveSpreadsheetPageSetup) =>
    setDraft((current) => ({ ...current, pageSetup: value }));
  const [error, setError] = useState('');

  useEffect(() => {
    syncDraft(spreadsheetPrintSettingsDraft(content, sheetId));
  }, [content, sheetId, syncDraft]);

  const changeSheet = (nextSheetId: string) => {
    if (nextSheetId === sheetId) return;
    if (dirty) {
      setError(officeMessage(messages, 'spreadsheet.print.error.unsaved'));
      return;
    }
    setSheetId(nextSheetId);
    replaceDraft(spreadsheetPrintSettingsDraft(content, nextSheetId));
    setError('');
  };

  const saveSettings = () => {
    if (!dirty) return;
    const sheet = availableSheets.find((item) => item.id === sheetId);
    const maximumRow = spreadsheetMaximumRow(sheet);
    const maximumColumn = spreadsheetMaximumColumn(sheet);
    const normalized = reference.trim()
      ? normalizeSpreadsheetPrintArea(reference)
      : null;
    const normalizedRows = titleRows.trim()
      ? normalizeSpreadsheetPrintTitleRows(titleRows)
      : null;
    const normalizedColumns = titleColumns.trim()
      ? normalizeSpreadsheetPrintTitleColumns(titleColumns)
      : null;
    const parsedRowPageBreaks = rowPageBreaks.trim()
      ? parseSpreadsheetRowPageBreaks(rowPageBreaks, maximumRow)
      : [];
    const parsedColumnPageBreaks = columnPageBreaks.trim()
      ? parseSpreadsheetColumnPageBreaks(columnPageBreaks, maximumColumn)
      : [];
    if (reference.trim() && !normalized) {
      setError(officeMessage(messages, 'spreadsheet.print.error.invalidRange'));
      return;
    }
    if (titleRows.trim() && !normalizedRows) {
      setError(officeMessage(messages, 'spreadsheet.print.error.titleRows'));
      return;
    }
    if (titleColumns.trim() && !normalizedColumns) {
      setError(officeMessage(messages, 'spreadsheet.print.error.titleCols'));
      return;
    }
    if (!parsedRowPageBreaks) {
      setError(officeMessage(messages, 'spreadsheet.print.error.hBreaks', { n: String(maximumRow + 1) }));
      return;
    }
    if (!parsedColumnPageBreaks) {
      setError(officeMessage(messages, 'spreadsheet.print.error.vBreaks'));
      return;
    }
    if (!validPageSetup(pageSetup)) {
      setError(officeMessage(messages, 'spreadsheet.print.error.outOfRange'));
      return;
    }

    const nextAreas = (content.printAreas ?? []).filter(
      (area) => area.sheetId !== sheetId,
    );
    if (normalized) nextAreas.push({ sheetId, reference: normalized });
    const nextTitles = (content.printTitles ?? []).filter(
      (titles) => titles.sheetId !== sheetId,
    );
    if (normalizedRows || normalizedColumns) {
      nextTitles.push({
        sheetId,
        rows: normalizedRows ?? undefined,
        columns: normalizedColumns ?? undefined,
      });
    }
    const nextPageBreaks = (content.pageBreaks ?? []).filter(
      (pageBreaks) => pageBreaks.sheetId !== sheetId,
    );
    if (parsedRowPageBreaks.length || parsedColumnPageBreaks.length) {
      nextPageBreaks.push({
        sheetId,
        rows: parsedRowPageBreaks.length ? parsedRowPageBreaks : undefined,
        columns: parsedColumnPageBreaks.length
          ? parsedColumnPageBreaks
          : undefined,
      });
    }
    const nextPageSetups = (content.pageSetups ?? []).filter(
      (item) => item.sheetId !== sheetId,
    );
    nextPageSetups.push({ sheetId, ...pageSetup });
    onChange({
      ...content,
      printAreas: nextAreas.length ? nextAreas : undefined,
      printTitles: nextTitles.length ? nextTitles : undefined,
      pageBreaks: nextPageBreaks.length ? nextPageBreaks : undefined,
      pageSetups: nextPageSetups,
    });
    replaceDraft({
      ...draft,
      reference: normalized ?? '',
      titleRows: normalizedRows ?? '',
      titleColumns: normalizedColumns ?? '',
      rowPageBreaks: formatSpreadsheetRowPageBreaks(parsedRowPageBreaks),
      columnPageBreaks: formatSpreadsheetColumnPageBreaks(
        parsedColumnPageBreaks,
      ),
    });
    setError('');
  };

  const clearSettings = () => {
    const nextAreas = (content.printAreas ?? []).filter(
      (area) => area.sheetId !== sheetId,
    );
    const nextTitles = (content.printTitles ?? []).filter(
      (titles) => titles.sheetId !== sheetId,
    );
    const nextPageBreaks = (content.pageBreaks ?? []).filter(
      (pageBreaks) => pageBreaks.sheetId !== sheetId,
    );
    const nextPageSetups = (content.pageSetups ?? []).filter(
      (pageSetup) => pageSetup.sheetId !== sheetId,
    );
    onChange({
      ...content,
      printAreas: nextAreas.length ? nextAreas : undefined,
      printTitles: nextTitles.length ? nextTitles : undefined,
      pageBreaks: nextPageBreaks.length ? nextPageBreaks : undefined,
      pageSetups: nextPageSetups.length ? nextPageSetups : undefined,
    });
    replaceDraft(spreadsheetPrintSettingsDraft(content, sheetId, true));
    setError('');
  };

  const cancelDraft = () => {
    resetDraft();
    setError('');
  };

  return (
    <form
      className="work-spreadsheet-print-area-form"
      data-office-escape-consumer={dirty || undefined}
      onKeyDown={(event) => {
        if (event.key !== 'Escape' || event.defaultPrevented || !dirty) return;
        event.preventDefault();
        event.stopPropagation();
        cancelDraft();
      }}
      onSubmit={(event) => {
        event.preventDefault();
        saveSettings();
      }}
    >
      <div className="work-office-field">
        <span>{officeMessage(messages, 'spreadsheet.print.sheet')}</span>
        <OfficeSelect
          ariaLabel={officeMessage(messages, 'spreadsheet.print.sheetAria')}
          value={sheetId}
          options={availableSheets.map((sheet) => ({
            value: sheet.id,
            label: sheet.name,
          }))}
          onValueChange={changeSheet}
        />
      </div>
      <div className="work-office-field reference">
        <span>{officeMessage(messages, 'spreadsheet.print.printArea')}</span>
        <OfficeTextField
          aria-label={officeMessage(messages, 'spreadsheet.print.printArea')}
          value={reference}
          placeholder="$A$1:$J$40"
          onChange={(event) => setReference(event.target.value)}
        />
      </div>
      <div className="work-office-field reference">
        <span>{officeMessage(messages, 'spreadsheet.print.titleRows')}</span>
        <OfficeTextField
          aria-label={officeMessage(messages, 'spreadsheet.print.titleRows')}
          value={titleRows}
          placeholder="$1:$2"
          onChange={(event) => setTitleRows(event.target.value)}
        />
      </div>
      <div className="work-office-field reference">
        <span>{officeMessage(messages, 'spreadsheet.print.titleCols')}</span>
        <OfficeTextField
          aria-label={officeMessage(messages, 'spreadsheet.print.titleCols')}
          value={titleColumns}
          placeholder="$A:$C"
          onChange={(event) => setTitleColumns(event.target.value)}
        />
      </div>
      <div className="work-office-field reference">
        <span>{officeMessage(messages, 'spreadsheet.print.hBreaks')}</span>
        <OfficeTextField
          aria-label={officeMessage(messages, 'spreadsheet.print.hBreaks')}
          value={rowPageBreaks}
          placeholder="20, 35"
          onChange={(event) => setRowPageBreaks(event.target.value)}
        />
      </div>
      <div className="work-office-field reference">
        <span>{officeMessage(messages, 'spreadsheet.print.vBreaks')}</span>
        <OfficeTextField
          aria-label={officeMessage(messages, 'spreadsheet.print.vBreaks')}
          value={columnPageBreaks}
          placeholder="E, K"
          onChange={(event) => setColumnPageBreaks(event.target.value)}
        />
      </div>
      <fieldset className="work-spreadsheet-page-setup-fields">
        <legend>{officeMessage(messages, 'spreadsheet.print.pageSetupLegend')}</legend>
        <div className="work-office-field">
          <span>{officeMessage(messages, 'spreadsheet.print.paperSize')}</span>
          <OfficeSelect
            ariaLabel={officeMessage(messages, 'spreadsheet.print.paperSize')}
            value={pageSetup.paperSize}
            options={[
              { value: 'a3', label: 'A3' },
              { value: 'a4', label: 'A4' },
              { value: 'a5', label: 'A5' },
              { value: 'letter', label: 'Letter' },
              { value: 'legal', label: 'Legal' },
              { value: 'tabloid', label: 'Tabloid' },
            ]}
            onValueChange={(paperSize) =>
              setPageSetup({
                ...pageSetup,
                paperSize:
                  paperSize as EffectiveSpreadsheetPageSetup['paperSize'],
              })
            }
          />
        </div>
        <div className="work-office-field">
          <span>{officeMessage(messages, 'spreadsheet.print.orientation')}</span>
          <OfficeSelect
            ariaLabel={officeMessage(messages, 'spreadsheet.print.orientation')}
            value={pageSetup.orientation}
            options={[
              { value: 'landscape', label: officeMessage(messages, 'spreadsheet.print.landscape') },
              { value: 'portrait', label: officeMessage(messages, 'spreadsheet.print.portrait') },
            ]}
            onValueChange={(orientation) =>
              setPageSetup({
                ...pageSetup,
                orientation:
                  orientation as EffectiveSpreadsheetPageSetup['orientation'],
              })
            }
          />
        </div>
        <div className="work-office-field">
          <span>{officeMessage(messages, 'spreadsheet.print.scalingMode')}</span>
          <OfficeSelect
            ariaLabel={officeMessage(messages, 'spreadsheet.print.scalingMode')}
            value={pageSetup.fitToPage ? 'fit' : 'scale'}
            options={[
              { value: 'scale', label: officeMessage(messages, 'spreadsheet.print.scaleBy') },
              { value: 'fit', label: officeMessage(messages, 'spreadsheet.print.fitToPages') },
            ]}
            onValueChange={(mode) =>
              setPageSetup({ ...pageSetup, fitToPage: mode === 'fit' })
            }
          />
        </div>
        <div className="work-office-field">
          <span>{officeMessage(messages, 'spreadsheet.print.scalePercent')}</span>
          <CommittedOfficeNumberField
            ariaLabel={officeMessage(messages, 'spreadsheet.print.scalePercentAria')}
            min={10}
            max={400}
            disabled={pageSetup.fitToPage}
            value={pageSetup.scale}
            normalizeValue={(value) => normalizePrintInteger(value, 10, 400)}
            onValueCommit={(scale) => setPageSetup({ ...pageSetup, scale })}
          />
        </div>
        <div className="work-office-field">
          <span>{officeMessage(messages, 'spreadsheet.print.fitWidth')}</span>
          <CommittedOfficeNumberField
            ariaLabel={officeMessage(messages, 'spreadsheet.print.fitWidthAria')}
            min={0}
            max={32767}
            disabled={!pageSetup.fitToPage}
            value={pageSetup.fitToWidth}
            normalizeValue={(value) => normalizePrintInteger(value, 0, 32_767)}
            onValueCommit={(fitToWidth) =>
              setPageSetup({ ...pageSetup, fitToWidth })
            }
          />
        </div>
        <div className="work-office-field">
          <span>{officeMessage(messages, 'spreadsheet.print.fitHeight')}</span>
          <CommittedOfficeNumberField
            ariaLabel={officeMessage(messages, 'spreadsheet.print.fitHeightAria')}
            min={0}
            max={32767}
            disabled={!pageSetup.fitToPage}
            value={pageSetup.fitToHeight}
            normalizeValue={(value) => normalizePrintInteger(value, 0, 32_767)}
            onValueCommit={(fitToHeight) =>
              setPageSetup({ ...pageSetup, fitToHeight })
            }
          />
        </div>
        <PageMarginField
          label={officeMessage(messages, 'spreadsheet.print.marginTop')}
          value={pageSetup.margins.top}
          onChange={(top) =>
            setPageSetup({
              ...pageSetup,
              margins: { ...pageSetup.margins, top },
            })
          }
        />
        <PageMarginField
          label={officeMessage(messages, 'spreadsheet.print.marginRight')}
          value={pageSetup.margins.right}
          onChange={(right) =>
            setPageSetup({
              ...pageSetup,
              margins: { ...pageSetup.margins, right },
            })
          }
        />
        <PageMarginField
          label={officeMessage(messages, 'spreadsheet.print.marginBottom')}
          value={pageSetup.margins.bottom}
          onChange={(bottom) =>
            setPageSetup({
              ...pageSetup,
              margins: { ...pageSetup.margins, bottom },
            })
          }
        />
        <PageMarginField
          label={officeMessage(messages, 'spreadsheet.print.marginLeft')}
          value={pageSetup.margins.left}
          onChange={(left) =>
            setPageSetup({
              ...pageSetup,
              margins: { ...pageSetup.margins, left },
            })
          }
        />
        <PageMarginField
          label={officeMessage(messages, 'spreadsheet.print.marginHeader')}
          value={pageSetup.margins.header}
          onChange={(header) =>
            setPageSetup({
              ...pageSetup,
              margins: { ...pageSetup.margins, header },
            })
          }
        />
        <PageMarginField
          label={officeMessage(messages, 'spreadsheet.print.marginFooter')}
          value={pageSetup.margins.footer}
          onChange={(footer) =>
            setPageSetup({
              ...pageSetup,
              margins: { ...pageSetup.margins, footer },
            })
          }
        />
        <OfficeCheckbox
          className="toggle"
          ariaLabel={officeMessage(messages, 'spreadsheet.print.centerH')}
          checked={pageSetup.horizontalCentered}
          onCheckedChange={(horizontalCentered) =>
            setPageSetup({ ...pageSetup, horizontalCentered })
          }
        >
          {officeMessage(messages, 'spreadsheet.print.centerH')}
        </OfficeCheckbox>
        <OfficeCheckbox
          className="toggle"
          ariaLabel={officeMessage(messages, 'spreadsheet.print.centerV')}
          checked={pageSetup.verticalCentered}
          onCheckedChange={(verticalCentered) =>
            setPageSetup({ ...pageSetup, verticalCentered })
          }
        >
          {officeMessage(messages, 'spreadsheet.print.centerV')}
        </OfficeCheckbox>
      </fieldset>
      <SpreadsheetHeaderFooterFields
        pageSetup={pageSetup}
        onChange={setPageSetup}
      />
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
          disabled={
            dirty ||
            (!savedArea && !savedTitles && !savedPageBreaks && !savedPageSetup)
          }
          onClick={clearSettings}
        >
          {officeMessage(messages, 'spreadsheet.print.clear')}
        </Button>
        <Button tone="secondary" disabled={!dirty} onClick={cancelDraft}>
          {officeMessage(messages, 'spreadsheet.print.cancelChanges')}
        </Button>
        <Button type="submit" tone="primary" disabled={!sheetId || !dirty}>
          {officeMessage(messages, 'spreadsheet.print.save')}
        </Button>
      </div>
    </form>
  );
}

function spreadsheetPrintSettingsDraft(
  content: WorkSpreadsheetContent,
  sheetId: string,
  empty = false,
): SpreadsheetPrintSettingsDraft {
  const area = empty
    ? undefined
    : content.printAreas?.find((item) => item.sheetId === sheetId);
  const titles = empty
    ? undefined
    : content.printTitles?.find((item) => item.sheetId === sheetId);
  const pageBreaks = empty
    ? undefined
    : content.pageBreaks?.find((item) => item.sheetId === sheetId);
  const pageSetup = empty
    ? undefined
    : content.pageSetups?.find((item) => item.sheetId === sheetId);
  return {
    reference: area?.reference ?? '',
    titleRows: titles?.rows ?? '',
    titleColumns: titles?.columns ?? '',
    rowPageBreaks: formatSpreadsheetRowPageBreaks(pageBreaks?.rows),
    columnPageBreaks: formatSpreadsheetColumnPageBreaks(pageBreaks?.columns),
    pageSetup: effectiveSpreadsheetPageSetup(pageSetup),
  };
}

function spreadsheetMaximumRow(
  sheet: WorkSpreadsheetContent['sheets'][number] | undefined,
): number {
  return Math.max(0, (sheet?.row ?? 1) - 1, (sheet?.data?.length ?? 1) - 1);
}

function spreadsheetMaximumColumn(
  sheet: WorkSpreadsheetContent['sheets'][number] | undefined,
): number {
  let maximum = Math.max(0, (sheet?.column ?? 1) - 1);
  for (const row of sheet?.data ?? [])
    maximum = Math.max(maximum, row.length - 1);
  return maximum;
}

function PageMarginField({
  label,
  value,
  onChange,
}: {
  label: string;
  value: number;
  onChange: (value: number) => void;
}) {
  return (
    <div className="work-office-field">
      <span>{label}</span>
      <CommittedOfficeNumberField
        ariaLabel={label}
        min={0}
        max={100}
        step={0.01}
        value={value}
        normalizeValue={(nextValue) => normalizePrintDecimal(nextValue, 0, 100)}
        onValueCommit={onChange}
      />
    </div>
  );
}

function normalizePrintInteger(
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

function normalizePrintDecimal(
  value: string,
  minimum: number,
  maximum: number,
): number | null {
  if (!value.trim()) return null;
  const number = Number(value);
  return Number.isFinite(number)
    ? Math.min(maximum, Math.max(minimum, Math.round(number * 100) / 100))
    : null;
}

function validPageSetup(pageSetup: EffectiveSpreadsheetPageSetup): boolean {
  const validInteger = (value: number, minimum: number, maximum: number) =>
    Number.isInteger(value) && value >= minimum && value <= maximum;
  return (
    validInteger(pageSetup.scale, 10, 400) &&
    validInteger(pageSetup.fitToWidth, 0, 32_767) &&
    validInteger(pageSetup.fitToHeight, 0, 32_767) &&
    validInteger(pageSetup.pageNumberStart, 1, 32_767) &&
    Object.values(pageSetup.margins).every(
      (margin) => Number.isFinite(margin) && margin >= 0 && margin <= 100,
    )
  );
}
