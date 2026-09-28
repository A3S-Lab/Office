import type { Cell } from '@fortune-sheet/core';
import {
  AArrowDown,
  AArrowUp,
  BadgeJapaneseYen,
  Bold,
  DecimalsArrowLeft,
  DecimalsArrowRight,
  Italic,
  Percent,
  Settings2,
  Strikethrough,
} from 'lucide-react';
import type { ReactNode } from 'react';
import { officeMessage } from '../../../i18n/office-locale';
import {
  OfficeColorPicker,
  OfficeSelect,
  type OfficeSelectOption,
} from './office-controls';
import { useOfficeMessages } from './office-messages-context';
import { SpreadsheetBorderRibbon } from './spreadsheet-border-ribbon';
import { spreadsheetCommandCatalog } from './spreadsheet-command-catalog';
import type {
  SpreadsheetEditorCanCommands,
  SpreadsheetEditorCommands,
} from './spreadsheet-command-controller';
import { SpreadsheetDateTimeMenu } from './spreadsheet-date-time-ribbon';
import {
  spreadsheetFontFamilyOptions,
  spreadsheetFontSizeOptions,
} from './spreadsheet-editor-support';
import { DEFAULT_SPREADSHEET_FONT_SIZE } from './spreadsheet-font-size';
import {
  type SpreadsheetNumberFormatPreset,
  spreadsheetNumberFormatCode,
  spreadsheetNumberFormatPreset,
  spreadsheetNumberFormatPresetLabels,
  spreadsheetNumberFormatValue,
} from './spreadsheet-number-format';
import { SpreadsheetUnderlineRibbon } from './spreadsheet-underline-ribbon';
import {
  WorkOfficeRibbonButton,
  WorkOfficeRibbonGroup,
} from './work-office-chrome';

interface SpreadsheetHomeFormatRibbonProps {
  can: SpreadsheetEditorCanCommands;
  commands: SpreadsheetEditorCommands;
  toolbarCell: Cell | null | undefined;
}

export function SpreadsheetFontRibbonGroup({
  can,
  commands,
  toolbarCell,
}: SpreadsheetHomeFormatRibbonProps) {
  const messages = useOfficeMessages();
  const fontFamily =
    typeof toolbarCell?.ff === 'string' ? toolbarCell.ff : 'Aptos';
  const fontSize = Number(toolbarCell?.fs ?? DEFAULT_SPREADSHEET_FONT_SIZE);
  const bold = Number(toolbarCell?.bl) === 1;
  const italic = Number(toolbarCell?.it) === 1;
  const strike = Number(toolbarCell?.cl) === 1;
  const textColor =
    typeof toolbarCell?.fc === 'string' ? toolbarCell.fc : '#172033';
  const fillColor =
    typeof toolbarCell?.bg === 'string' ? toolbarCell.bg : '#ffffff';
  const growFont = spreadsheetCommandCatalog.growFont;
  const shrinkFont = spreadsheetCommandCatalog.shrinkFont;

  return (
    <WorkOfficeRibbonGroup
      label={officeMessage(messages, 'spreadsheet.ribbon.font')}
      priority="high"
    >
      <OfficeSelect
        className="work-spreadsheet-font-family"
        ariaLabel={officeMessage(messages, 'spreadsheet.ribbon.fontAria')}
        value={fontFamily}
        disabled={!can.setCellFormat('ff', fontFamily)}
        options={spreadsheetFontFamilyOptions(
          typeof toolbarCell?.ff === 'string' ? toolbarCell.ff : undefined,
        )}
        onValueChange={(value) => commands.setCellFormat('ff', value)}
      />
      <OfficeSelect
        className="work-spreadsheet-font-size"
        ariaLabel={officeMessage(messages, 'spreadsheet.ribbon.fontSizeAria')}
        value={String(fontSize)}
        disabled={!can.setCellFormat('fs', fontSize)}
        options={spreadsheetFontSizeOptions(toolbarCell?.fs)}
        onValueChange={(value) => commands.setCellFormat('fs', Number(value))}
      />
      <WorkOfficeRibbonButton
        label={growFont.label}
        title={officeMessage(messages, 'spreadsheet.ribbon.labelWithShortcut', {
          label: growFont.label,
          shortcut: growFont.shortcut.label,
        })}
        aria-keyshortcuts={growFont.shortcut.aria}
        displayLabel={false}
        disabled={!can.adjustFontSize('grow')}
        onClick={() => commands.adjustFontSize('grow')}
      >
        <AArrowUp size={15} />
      </WorkOfficeRibbonButton>
      <WorkOfficeRibbonButton
        label={shrinkFont.label}
        title={officeMessage(messages, 'spreadsheet.ribbon.labelWithShortcut', {
          label: shrinkFont.label,
          shortcut: shrinkFont.shortcut.label,
        })}
        aria-keyshortcuts={shrinkFont.shortcut.aria}
        displayLabel={false}
        disabled={!can.adjustFontSize('shrink')}
        onClick={() => commands.adjustFontSize('shrink')}
      >
        <AArrowDown size={15} />
      </WorkOfficeRibbonButton>
      <SpreadsheetFontToggle
        active={bold}
        attribute="bl"
        can={can}
        commands={commands}
        command="bold"
        icon={<Bold size={15} />}
      />
      <SpreadsheetFontToggle
        active={italic}
        attribute="it"
        can={can}
        commands={commands}
        command="italic"
        icon={<Italic size={15} />}
      />
      <SpreadsheetUnderlineRibbon
        can={can}
        commands={commands}
        value={toolbarCell?.un}
      />
      <SpreadsheetFontToggle
        active={strike}
        attribute="cl"
        can={can}
        commands={commands}
        command="strike"
        icon={<Strikethrough size={15} />}
      />
      <OfficeColorPicker
        compact
        className="work-color-tool work-spreadsheet-font-color"
        ariaLabel={officeMessage(messages, 'spreadsheet.ribbon.textColorAria')}
        value={textColor}
        disabled={!can.setCellFormat('fc', textColor)}
        onValueChange={(value) => commands.setCellFormat('fc', value)}
        resetAction={{
          kind: 'automatic',
          label: officeMessage(messages, 'spreadsheet.ribbon.autoColor'),
          onSelect: () => commands.setCellFormat('fc', undefined),
        }}
      />
      <OfficeColorPicker
        compact
        className="work-color-tool work-spreadsheet-fill-color"
        ariaLabel={officeMessage(messages, 'spreadsheet.ribbon.fillColorAria')}
        value={fillColor}
        disabled={!can.setCellFormat('bg', fillColor)}
        onValueChange={(value) => commands.setCellFormat('bg', value)}
        resetAction={{
          kind: 'none',
          label: officeMessage(messages, 'spreadsheet.ribbon.noFill'),
          onSelect: () => commands.setCellFormat('bg', undefined),
        }}
      />
      <SpreadsheetBorderRibbon can={can} commands={commands} />
    </WorkOfficeRibbonGroup>
  );
}

export function SpreadsheetNumberRibbonGroup({
  can,
  commands,
  toolbarCell,
}: SpreadsheetHomeFormatRibbonProps) {
  const messages = useOfficeMessages();
  const numberFormat = toolbarCell?.ct?.fa?.trim() || 'General';
  const numberFormatPreset = spreadsheetNumberFormatPreset(numberFormat);
  const currentNumberFormatValue = spreadsheetNumberFormatValue(
    numberFormat,
    toolbarCell,
  );
  const currencyDefinition = spreadsheetCommandCatalog.numberFormatCurrency;
  const percentDefinition = spreadsheetCommandCatalog.numberFormatPercent;
  const decreaseDefinition = spreadsheetCommandCatalog.decreaseDecimalPlaces;
  const increaseDefinition = spreadsheetCommandCatalog.increaseDecimalPlaces;
  const formatCellsDefinition = spreadsheetCommandCatalog.formatCells;
  const commonGroup = officeMessage(
    messages,
    'spreadsheet.numberFormat.group.common',
  );
  const dateTimeGroup = officeMessage(
    messages,
    'spreadsheet.numberFormat.group.dateTime',
  );
  const otherGroup = officeMessage(
    messages,
    'spreadsheet.numberFormat.group.other',
  );
  const spreadsheetNumberFormatOptions: readonly OfficeSelectOption<SpreadsheetNumberFormatPreset>[] =
    [
      {
        value: 'general',
        label: spreadsheetCommandCatalog.numberFormatGeneral.label,
        group: commonGroup,
        meta: spreadsheetCommandCatalog.numberFormatGeneral.shortcut.label,
      },
      {
        value: 'number',
        label: spreadsheetCommandCatalog.numberFormatNumber.label,
        group: commonGroup,
        meta: spreadsheetCommandCatalog.numberFormatNumber.shortcut.label,
      },
      {
        value: 'currency',
        label: spreadsheetCommandCatalog.numberFormatCurrency.label,
        group: commonGroup,
        meta: spreadsheetCommandCatalog.numberFormatCurrency.shortcut.label,
      },
      {
        value: 'accounting',
        label: spreadsheetCommandCatalog.numberFormatAccounting.label,
        group: commonGroup,
      },
      {
        value: 'percent',
        label: spreadsheetCommandCatalog.numberFormatPercent.label,
        group: commonGroup,
        meta: spreadsheetCommandCatalog.numberFormatPercent.shortcut.label,
      },
      {
        value: 'date',
        label: spreadsheetCommandCatalog.numberFormatDate.label,
        group: dateTimeGroup,
        meta: spreadsheetCommandCatalog.numberFormatDate.shortcut.label,
      },
      {
        value: 'time',
        label: spreadsheetCommandCatalog.numberFormatTime.label,
        group: dateTimeGroup,
        meta: spreadsheetCommandCatalog.numberFormatTime.shortcut.label,
      },
      {
        value: 'scientific',
        label: spreadsheetCommandCatalog.numberFormatScientific.label,
        group: otherGroup,
        meta: spreadsheetCommandCatalog.numberFormatScientific.shortcut.label,
      },
      {
        value: 'fraction',
        label: spreadsheetCommandCatalog.numberFormatFraction.label,
        group: otherGroup,
      },
      {
        value: 'text',
        label: spreadsheetCommandCatalog.numberFormatText.label,
        group: otherGroup,
      },
      {
        value: 'custom',
        label: spreadsheetNumberFormatPresetLabels.custom,
        group: otherGroup,
        disabled: true,
      },
    ];

  return (
    <WorkOfficeRibbonGroup
      label={officeMessage(messages, 'spreadsheet.ribbon.number')}
      priority="high"
    >
      <div className="work-spreadsheet-number-format-stack">
        <OfficeSelect
          className="work-spreadsheet-number-format"
          ariaLabel={officeMessage(
            messages,
            'spreadsheet.ribbon.numberFormatAria',
          )}
          value={numberFormatPreset}
          disabled={!can.setCellFormat('ct', currentNumberFormatValue)}
          options={spreadsheetNumberFormatOptions}
          onValueChange={(preset) => {
            if (preset === 'custom') return;
            commands.setCellFormat(
              'ct',
              spreadsheetNumberFormatValue(
                spreadsheetNumberFormatCode(preset),
                toolbarCell,
              ),
            );
          }}
        />
        <SpreadsheetDateTimeMenu can={can} commands={commands} />
      </div>
      <WorkOfficeRibbonButton
        label={officeMessage(messages, 'spreadsheet.ribbon.formatSuffix', {
          label: currencyDefinition.label,
        })}
        title={officeMessage(
          messages,
          'spreadsheet.ribbon.formatSuffixWithShortcut',
          {
            label: currencyDefinition.label,
            shortcut: currencyDefinition.shortcut.label,
          },
        )}
        aria-keyshortcuts={currencyDefinition.shortcut.aria}
        displayLabel={false}
        active={numberFormatPreset === 'currency'}
        disabled={
          !can.setCellFormat(
            'ct',
            spreadsheetNumberFormatValue(
              spreadsheetNumberFormatCode('currency'),
              toolbarCell,
            ),
          )
        }
        onClick={() =>
          commands.setCellFormat(
            'ct',
            spreadsheetNumberFormatValue(
              spreadsheetNumberFormatCode('currency'),
              toolbarCell,
            ),
          )
        }
      >
        <BadgeJapaneseYen size={15} />
      </WorkOfficeRibbonButton>
      <WorkOfficeRibbonButton
        label={officeMessage(messages, 'spreadsheet.ribbon.formatSuffix', {
          label: percentDefinition.label,
        })}
        title={officeMessage(
          messages,
          'spreadsheet.ribbon.formatSuffixWithShortcut',
          {
            label: percentDefinition.label,
            shortcut: percentDefinition.shortcut.label,
          },
        )}
        aria-keyshortcuts={percentDefinition.shortcut.aria}
        displayLabel={false}
        active={numberFormatPreset === 'percent'}
        disabled={
          !can.setCellFormat(
            'ct',
            spreadsheetNumberFormatValue(
              spreadsheetNumberFormatCode('percent'),
              toolbarCell,
            ),
          )
        }
        onClick={() =>
          commands.setCellFormat(
            'ct',
            spreadsheetNumberFormatValue(
              spreadsheetNumberFormatCode('percent'),
              toolbarCell,
            ),
          )
        }
      >
        <Percent size={15} />
      </WorkOfficeRibbonButton>
      <WorkOfficeRibbonButton
        label={decreaseDefinition.label}
        title={decreaseDefinition.label}
        displayLabel={false}
        disabled={!can.adjustDecimalPlaces('decrease')}
        onClick={() => commands.adjustDecimalPlaces('decrease')}
      >
        <DecimalsArrowLeft size={16} />
      </WorkOfficeRibbonButton>
      <WorkOfficeRibbonButton
        label={increaseDefinition.label}
        title={increaseDefinition.label}
        displayLabel={false}
        disabled={!can.adjustDecimalPlaces('increase')}
        onClick={() => commands.adjustDecimalPlaces('increase')}
      >
        <DecimalsArrowRight size={16} />
      </WorkOfficeRibbonButton>
      <WorkOfficeRibbonButton
        label={formatCellsDefinition.label}
        title={officeMessage(messages, 'spreadsheet.ribbon.labelWithShortcut', {
          label: formatCellsDefinition.label,
          shortcut: formatCellsDefinition.shortcut.label,
        })}
        aria-keyshortcuts={formatCellsDefinition.shortcut.aria}
        displayLabel={false}
        disabled={!can.openFormatCells()}
        onClick={() => commands.openFormatCells()}
      >
        <Settings2 size={15} />
      </WorkOfficeRibbonButton>
    </WorkOfficeRibbonGroup>
  );
}

function SpreadsheetFontToggle({
  active,
  attribute,
  can,
  command,
  commands,
  icon,
}: Pick<SpreadsheetHomeFormatRibbonProps, 'can' | 'commands'> & {
  active: boolean;
  attribute: 'bl' | 'cl' | 'it';
  command: 'bold' | 'italic' | 'strike';
  icon: ReactNode;
}) {
  const messages = useOfficeMessages();
  const definition = spreadsheetCommandCatalog[command];
  return (
    <WorkOfficeRibbonButton
      data-spreadsheet-rich-text-format="true"
      label={definition.label}
      title={officeMessage(messages, 'spreadsheet.ribbon.labelWithShortcut', {
        label: definition.label,
        shortcut: definition.shortcut.label,
      })}
      aria-keyshortcuts={definition.shortcut.aria}
      displayLabel={false}
      active={active}
      disabled={!can.toggleCellFormat(attribute)}
      onClick={() => commands.toggleCellFormat(attribute)}
    >
      {icon}
    </WorkOfficeRibbonButton>
  );
}
