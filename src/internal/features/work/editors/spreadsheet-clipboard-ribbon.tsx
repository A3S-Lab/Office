import {
  ChevronDown,
  ClipboardPaste,
  Copy,
  Hash,
  Paintbrush,
  Scissors,
  Sigma,
  TableProperties,
} from 'lucide-react';
import type { ReactNode } from 'react';
import { Popover } from '../../../design-system/primitives';
import { moveOfficeMenuFocus } from './office-menu-keyboard';
import { spreadsheetCommandCatalog } from './spreadsheet-command-catalog';
import type {
  SpreadsheetEditorCanCommands,
  SpreadsheetEditorCommands,
} from './spreadsheet-command-controller';
import type { SpreadsheetFormatPainterMode } from './spreadsheet-format-painter';
import type { SpreadsheetPasteContent } from './spreadsheet-paste-special';
import {
  WorkOfficeRibbonButton,
  WorkOfficeRibbonGroup,
} from './work-office-chrome';

export function SpreadsheetClipboardRibbonGroup({
  can,
  commands,
  formatPainterMode,
}: {
  can: SpreadsheetEditorCanCommands;
  commands: SpreadsheetEditorCommands;
  formatPainterMode: SpreadsheetFormatPainterMode | null;
}) {
  return (
    <WorkOfficeRibbonGroup label="剪贴板" priority="high">
      <SpreadsheetPasteMenu can={can} commands={commands} />
      <WorkOfficeRibbonButton
        label={spreadsheetCommandCatalog.cut.label}
        title={`${spreadsheetCommandCatalog.cut.label}（${spreadsheetCommandCatalog.cut.shortcut.label}）`}
        aria-keyshortcuts={spreadsheetCommandCatalog.cut.shortcut.aria}
        disabled={!can.cutSelection()}
        onClick={commands.cutSelection}
      >
        <Scissors size={19} />
      </WorkOfficeRibbonButton>
      <WorkOfficeRibbonButton
        label={spreadsheetCommandCatalog.copy.label}
        title={`${spreadsheetCommandCatalog.copy.label}（${spreadsheetCommandCatalog.copy.shortcut.label}）`}
        aria-keyshortcuts={spreadsheetCommandCatalog.copy.shortcut.aria}
        disabled={!can.copySelection()}
        onClick={commands.copySelection}
      >
        <Copy size={19} />
      </WorkOfficeRibbonButton>
      <WorkOfficeRibbonButton
        label={spreadsheetCommandCatalog.formatPainter.label}
        title={spreadsheetFormatPainterTitle(formatPainterMode)}
        active={formatPainterMode !== null}
        badge={formatPainterMode === 'locked' ? '连续' : undefined}
        disabled={
          formatPainterMode === null
            ? !can.activateFormatPainter('once')
            : !can.cancelFormatPainter()
        }
        onClick={(event) => {
          if (event.detail > 1) return;
          if (formatPainterMode === null) {
            commands.activateFormatPainter('once');
          } else {
            commands.cancelFormatPainter();
          }
        }}
        onDoubleClick={() => commands.activateFormatPainter('locked')}
      >
        <Paintbrush size={19} />
      </WorkOfficeRibbonButton>
    </WorkOfficeRibbonGroup>
  );
}

function SpreadsheetPasteMenu({
  can,
  commands,
}: {
  can: SpreadsheetEditorCanCommands;
  commands: SpreadsheetEditorCommands;
}) {
  const items: readonly {
    content: SpreadsheetPasteContent;
    id: string;
    label: string;
    icon: ReactNode;
  }[] = [
    {
      content: 'all',
      id: `${spreadsheetCommandCatalog.paste.id}.all`,
      label: '全部',
      icon: <ClipboardPaste size={16} />,
    },
    {
      content: 'values',
      id: `${spreadsheetCommandCatalog.paste.id}.values`,
      label: '值',
      icon: <Hash size={16} />,
    },
    {
      content: 'formulas',
      id: `${spreadsheetCommandCatalog.paste.id}.formulas`,
      label: '公式',
      icon: <Sigma size={16} />,
    },
    {
      content: 'formats',
      id: `${spreadsheetCommandCatalog.paste.id}.formats`,
      label: '格式',
      icon: <Paintbrush size={16} />,
    },
  ];
  const primaryDisabled = !can.pasteSelection();
  const menuDisabled =
    !can.openPasteSpecial() &&
    items.every(({ content }) => !can.pasteSpecial(content));

  return (
    <Popover
      label="更多粘贴方式"
      panelLabel="粘贴选项"
      panelRole="menu"
      portal
      className="work-spreadsheet-ribbon-split-root"
      panelClassName="work-office-context-menu work-spreadsheet-ribbon-menu work-spreadsheet-paste-menu"
      disabled={menuDisabled}
      focusFirstOnOpen
      onPanelKeyDown={moveOfficeMenuFocus}
      trigger={(triggerProps, { open }) => (
        <>
          <button
            type="button"
            className="with-label work-spreadsheet-ribbon-split-primary"
            aria-label={spreadsheetCommandCatalog.paste.label}
            aria-keyshortcuts={spreadsheetCommandCatalog.paste.shortcut.aria}
            title={`${spreadsheetCommandCatalog.paste.label}（${spreadsheetCommandCatalog.paste.shortcut.label}）`}
            disabled={primaryDisabled}
            onClick={commands.pasteSelection}
          >
            <ClipboardPaste size={19} />
            <span>{spreadsheetCommandCatalog.paste.label}</span>
          </button>
          <button
            {...triggerProps}
            className={`work-spreadsheet-ribbon-split-disclosure${open ? ' active' : ''}`}
            title="更多粘贴方式"
          >
            <ChevronDown size={13} aria-hidden="true" />
          </button>
        </>
      )}
    >
      {(close) => (
        <>
          {items.map(({ content, id, label, icon }) => (
            <button
              key={id}
              type="button"
              role="menuitem"
              tabIndex={-1}
              disabled={!can.pasteSpecial(content)}
              onClick={() => {
                close();
                commands.pasteSpecial(content);
              }}
            >
              <span className="work-spreadsheet-ribbon-menu-item-icon">
                {icon}
              </span>
              <span>{label}</span>
            </button>
          ))}
          <hr className="work-spreadsheet-ribbon-menu-separator" />
          <button
            type="button"
            role="menuitem"
            tabIndex={-1}
            aria-keyshortcuts={
              spreadsheetCommandCatalog.pasteSpecial.shortcut.aria
            }
            disabled={!can.openPasteSpecial()}
            onClick={() => {
              close();
              commands.openPasteSpecial();
            }}
          >
            <span className="work-spreadsheet-ribbon-menu-item-icon">
              <TableProperties size={16} />
            </span>
            <span>{spreadsheetCommandCatalog.pasteSpecial.label}…</span>
            <kbd>{spreadsheetCommandCatalog.pasteSpecial.shortcut.label}</kbd>
          </button>
        </>
      )}
    </Popover>
  );
}

function spreadsheetFormatPainterTitle(
  mode: SpreadsheetFormatPainterMode | null,
): string {
  if (mode === 'locked') {
    return '格式刷已锁定（再次点击或按 Escape 退出）';
  }
  if (mode === 'once') {
    return '格式刷已开启（选择目标区域，按 Escape 退出）';
  }
  return '格式刷（单击应用一次，双击锁定连续应用）';
}
