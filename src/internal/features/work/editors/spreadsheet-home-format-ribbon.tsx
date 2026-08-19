import type { Cell } from '@fortune-sheet/core';
import {
  Bold,
  DecimalsArrowLeft,
  DecimalsArrowRight,
  Italic,
  Percent,
  Strikethrough,
  Underline,
} from 'lucide-react';
import type { ReactNode } from 'react';
import { SpreadsheetBorderRibbon } from './spreadsheet-border-ribbon';
import { OfficeColorPicker, OfficeSelect } from './office-controls';
import { spreadsheetCommandCatalog } from './spreadsheet-command-catalog';
import type {
  SpreadsheetEditorCanCommands,
  SpreadsheetEditorCommands,
} from './spreadsheet-command-controller';
import {
  spreadsheetFontFamilyOptions,
  spreadsheetFontSizeOptions,
} from './spreadsheet-editor-support';
import {
  adjustSpreadsheetNumberFormat,
  type SpreadsheetNumberFormatPreset,
  spreadsheetNumberFormatCode,
  spreadsheetNumberFormatPreset,
  spreadsheetNumberFormatValue,
} from './spreadsheet-number-format';
import {
  WorkOfficeRibbonButton,
  WorkOfficeRibbonGroup,
} from './work-office-chrome';

const spreadsheetNumberFormatOptions: readonly {
  value: SpreadsheetNumberFormatPreset;
  label: string;
  disabled?: boolean;
}[] = [
  { value: 'general', label: '常规' },
  { value: 'number', label: '数字' },
  { value: 'percent', label: '百分比' },
  { value: 'custom', label: '自定义', disabled: true },
];

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
  const fontFamily =
    typeof toolbarCell?.ff === 'string' ? toolbarCell.ff : 'Aptos';
  const fontSize = Number(toolbarCell?.fs ?? 10);
  const bold = Number(toolbarCell?.bl) === 1;
  const italic = Number(toolbarCell?.it) === 1;
  const underline = Number(toolbarCell?.un) === 1;
  const strike = Number(toolbarCell?.cl) === 1;
  const textColor =
    typeof toolbarCell?.fc === 'string' ? toolbarCell.fc : '#172033';
  const fillColor =
    typeof toolbarCell?.bg === 'string' ? toolbarCell.bg : '#ffffff';

  return (
    <WorkOfficeRibbonGroup label="字体" priority="high">
      <OfficeSelect
        className="work-spreadsheet-font-family"
        ariaLabel="字体"
        value={fontFamily}
        disabled={!can.setCellFormat('ff', fontFamily)}
        options={spreadsheetFontFamilyOptions(
          typeof toolbarCell?.ff === 'string' ? toolbarCell.ff : undefined,
        )}
        onValueChange={(value) => commands.setCellFormat('ff', value)}
      />
      <OfficeSelect
        ariaLabel="字号"
        value={String(fontSize)}
        disabled={!can.setCellFormat('fs', fontSize)}
        options={spreadsheetFontSizeOptions(toolbarCell?.fs)}
        onValueChange={(value) => commands.setCellFormat('fs', Number(value))}
      />
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
      <SpreadsheetFontToggle
        active={underline}
        attribute="un"
        can={can}
        commands={commands}
        command="underline"
        icon={<Underline size={15} />}
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
        className="work-color-tool"
        ariaLabel="文字颜色"
        value={textColor}
        disabled={!can.setCellFormat('fc', textColor)}
        onValueChange={(value) => commands.setCellFormat('fc', value)}
      />
      <OfficeColorPicker
        compact
        className="work-color-tool work-spreadsheet-fill-color"
        ariaLabel="填充颜色"
        value={fillColor}
        disabled={!can.setCellFormat('bg', fillColor)}
        onValueChange={(value) => commands.setCellFormat('bg', value)}
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
  const numberFormat = toolbarCell?.ct?.fa?.trim() || 'General';
  const numberFormatPreset = spreadsheetNumberFormatPreset(numberFormat);
  const decreasedNumberFormat = adjustSpreadsheetNumberFormat(numberFormat, -1);
  const increasedNumberFormat = adjustSpreadsheetNumberFormat(numberFormat, 1);
  const currentNumberFormatValue = spreadsheetNumberFormatValue(
    numberFormat,
    toolbarCell,
  );
  const decreasedNumberFormatValue = spreadsheetNumberFormatValue(
    decreasedNumberFormat,
    toolbarCell,
  );
  const increasedNumberFormatValue = spreadsheetNumberFormatValue(
    increasedNumberFormat,
    toolbarCell,
  );

  return (
    <WorkOfficeRibbonGroup label="数字" priority="high">
      <OfficeSelect
        className="work-spreadsheet-number-format"
        ariaLabel="数字格式"
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
      <WorkOfficeRibbonButton
        label="百分比格式"
        title="百分比格式"
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
        label="减少小数位"
        title="减少小数位"
        displayLabel={false}
        disabled={
          decreasedNumberFormat === numberFormat ||
          !can.setCellFormat('ct', decreasedNumberFormatValue)
        }
        onClick={() => commands.setCellFormat('ct', decreasedNumberFormatValue)}
      >
        <DecimalsArrowLeft size={16} />
      </WorkOfficeRibbonButton>
      <WorkOfficeRibbonButton
        label="增加小数位"
        title="增加小数位"
        displayLabel={false}
        disabled={
          increasedNumberFormat === numberFormat ||
          !can.setCellFormat('ct', increasedNumberFormatValue)
        }
        onClick={() => commands.setCellFormat('ct', increasedNumberFormatValue)}
      >
        <DecimalsArrowRight size={16} />
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
  attribute: 'bl' | 'cl' | 'it' | 'un';
  command: 'bold' | 'italic' | 'strike' | 'underline';
  icon: ReactNode;
}) {
  const definition = spreadsheetCommandCatalog[command];
  const value = active ? 0 : 1;
  return (
    <WorkOfficeRibbonButton
      label={definition.label}
      title={`${definition.label}（${definition.shortcut.label}）`}
      aria-keyshortcuts={definition.shortcut.aria}
      displayLabel={false}
      active={active}
      disabled={!can.setCellFormat(attribute, value)}
      onClick={() => commands.setCellFormat(attribute, value)}
    >
      {icon}
    </WorkOfficeRibbonButton>
  );
}
