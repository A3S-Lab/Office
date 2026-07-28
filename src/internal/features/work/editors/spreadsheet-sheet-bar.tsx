import {
  ArrowLeft,
  ArrowRight,
  Check,
  ChevronDown,
  CopyPlus,
  EyeOff,
  Layers2,
  Palette,
  Pencil,
  Plus,
  Trash2,
} from 'lucide-react';
import {
  type CSSProperties,
  type KeyboardEvent,
  type ReactNode,
  useEffect,
  useRef,
  useState,
} from 'react';
import { Popover } from '../../../design-system/primitives';
import type { WorkSpreadsheetSheet } from '../work-types';
import {
  isSpreadsheetSheetHidden,
  type SpreadsheetSheetMoveDirection,
} from './spreadsheet-sheet-model';

const spreadsheetSheetColors = [
  { color: '#4f7de8', label: '蓝色' },
  { color: '#35a37a', label: '绿色' },
  { color: '#d59a36', label: '橙色' },
  { color: '#e06c53', label: '红色' },
  { color: '#8b6ad8', label: '紫色' },
] as const;

export interface SpreadsheetSheetBarProps {
  activeSheetId: string;
  editable: boolean;
  sheets: readonly WorkSpreadsheetSheet[];
  onActivate(sheetId: string): void;
  onCreate(): void;
  onDelete(sheetId: string): void;
  onDuplicate(sheetId: string): void;
  onHide(sheetId: string): void;
  onMove(sheetId: string, direction: SpreadsheetSheetMoveDirection): void;
  onRename(sheetId: string, name: string): void;
  onSetColor(sheetId: string, color: string | null): void;
  onShow(sheetId: string): void;
}

export function SpreadsheetSheetBar({
  activeSheetId,
  editable,
  sheets,
  onActivate,
  onCreate,
  onDelete,
  onDuplicate,
  onHide,
  onMove,
  onRename,
  onSetColor,
  onShow,
}: SpreadsheetSheetBarProps) {
  const orderedSheets = [...sheets].sort(
    (left, right) =>
      (left.order ?? sheets.indexOf(left)) -
      (right.order ?? sheets.indexOf(right)),
  );
  const visibleSheets = orderedSheets.filter(
    (sheet) => !isSpreadsheetSheetHidden(sheet),
  );
  const [renamingSheetId, setRenamingSheetId] = useState<string | null>(null);
  const [renameDraft, setRenameDraft] = useState('');
  const renameInputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    if (!renamingSheetId) return;
    const frame = requestAnimationFrame(() => {
      renameInputRef.current?.focus({ preventScroll: true });
      renameInputRef.current?.select();
    });
    return () => cancelAnimationFrame(frame);
  }, [renamingSheetId]);

  const beginRename = (sheet: WorkSpreadsheetSheet) => {
    setRenameDraft(sheet.name);
    setRenamingSheetId(sheet.id ?? null);
  };
  const finishRename = (sheet: WorkSpreadsheetSheet, commit: boolean) => {
    if (commit && sheet.id && renameDraft.trim() !== sheet.name) {
      onRename(sheet.id, renameDraft);
    }
    setRenamingSheetId(null);
  };

  return (
    <nav className="work-spreadsheet-sheet-bar" aria-label="工作表">
      <div className="work-spreadsheet-sheet-tools">
        <button
          type="button"
          aria-label="新建工作表"
          title="新建工作表"
          disabled={!editable}
          onClick={onCreate}
        >
          <Plus size={15} />
        </button>
        <Popover
          label="工作表列表"
          panelLabel="工作表列表"
          panelRole="menu"
          placement="top-start"
          portal
          panelClassName="work-spreadsheet-sheet-popover work-spreadsheet-sheet-list"
          focusFirstOnOpen
          onPanelKeyDown={moveSpreadsheetSheetMenuFocus}
          trigger={(triggerProps) => (
            <button {...triggerProps} title="工作表列表">
              <Layers2 size={14} />
            </button>
          )}
        >
          {(close) =>
            orderedSheets.map((sheet) => (
              <button
                key={sheet.id ?? sheet.name}
                type="button"
                role="menuitemradio"
                aria-checked={sheet.id === activeSheetId}
                onClick={() => {
                  close();
                  if (!sheet.id) return;
                  if (isSpreadsheetSheetHidden(sheet)) onShow(sheet.id);
                  else onActivate(sheet.id);
                }}
              >
                {sheet.id === activeSheetId ? (
                  <Check size={14} />
                ) : isSpreadsheetSheetHidden(sheet) ? (
                  <EyeOff size={14} />
                ) : (
                  <span aria-hidden="true" />
                )}
                <span>{sheet.name}</span>
              </button>
            ))
          }
        </Popover>
      </div>

      <div className="work-spreadsheet-sheet-tabs" role="tablist">
        {visibleSheets.map((sheet, index) => {
          const sheetId = sheet.id ?? '';
          const active = sheetId === activeSheetId;
          const colorStyle = sheet.color
            ? ({ '--work-sheet-color': sheet.color } as CSSProperties)
            : undefined;
          return (
            <div
              key={sheetId || sheet.name}
              className={`work-spreadsheet-sheet-tab${active ? ' active' : ''}`}
              style={colorStyle}
            >
              {renamingSheetId === sheetId ? (
                <input
                  ref={renameInputRef}
                  aria-label={`重命名${sheet.name}`}
                  data-office-shortcuts="ignore"
                  value={renameDraft}
                  maxLength={31}
                  onChange={(event) =>
                    setRenameDraft(event.currentTarget.value)
                  }
                  onBlur={() => finishRename(sheet, true)}
                  onKeyDown={(event) => {
                    if (event.key === 'Enter') {
                      event.preventDefault();
                      finishRename(sheet, true);
                    } else if (event.key === 'Escape') {
                      event.preventDefault();
                      finishRename(sheet, false);
                    }
                  }}
                />
              ) : (
                <button
                  type="button"
                  role="tab"
                  aria-selected={active}
                  tabIndex={active ? 0 : -1}
                  title={sheet.name}
                  onClick={() => sheetId && onActivate(sheetId)}
                  onDoubleClick={() => editable && beginRename(sheet)}
                >
                  <span>{sheet.name}</span>
                </button>
              )}
              {editable && renamingSheetId !== sheetId && (
                <SpreadsheetSheetMenu
                  sheet={sheet}
                  canDelete={sheets.length > 1}
                  canHide={visibleSheets.length > 1}
                  canMoveLeft={index > 0}
                  canMoveRight={index < visibleSheets.length - 1}
                  onDelete={onDelete}
                  onDuplicate={onDuplicate}
                  onHide={onHide}
                  onMove={onMove}
                  onRename={beginRename}
                  onSetColor={onSetColor}
                />
              )}
            </div>
          );
        })}
      </div>
    </nav>
  );
}

function SpreadsheetSheetMenu({
  sheet,
  canDelete,
  canHide,
  canMoveLeft,
  canMoveRight,
  onDelete,
  onDuplicate,
  onHide,
  onMove,
  onRename,
  onSetColor,
}: {
  sheet: WorkSpreadsheetSheet;
  canDelete: boolean;
  canHide: boolean;
  canMoveLeft: boolean;
  canMoveRight: boolean;
  onDelete(sheetId: string): void;
  onDuplicate(sheetId: string): void;
  onHide(sheetId: string): void;
  onMove(sheetId: string, direction: SpreadsheetSheetMoveDirection): void;
  onRename(sheet: WorkSpreadsheetSheet): void;
  onSetColor(sheetId: string, color: string | null): void;
}) {
  const sheetId = sheet.id ?? '';
  return (
    <Popover
      label={`${sheet.name}选项`}
      panelLabel={`${sheet.name}工作表操作`}
      panelRole="menu"
      placement="top-start"
      portal
      panelClassName="work-spreadsheet-sheet-popover"
      focusFirstOnOpen
      onPanelKeyDown={moveSpreadsheetSheetMenuFocus}
      trigger={(triggerProps, { open }) => (
        <button
          {...triggerProps}
          className="work-spreadsheet-sheet-options"
          title={`${sheet.name}选项`}
        >
          <ChevronDown
            size={12}
            className={open ? 'open' : undefined}
            aria-hidden="true"
          />
        </button>
      )}
    >
      {(close) => (
        <>
          <SheetMenuButton
            icon={<Pencil size={14} />}
            label="重命名"
            onClick={() => {
              close();
              onRename(sheet);
            }}
          />
          <SheetMenuButton
            icon={<CopyPlus size={14} />}
            label="复制工作表"
            onClick={() => {
              close();
              onDuplicate(sheetId);
            }}
          />
          <SheetMenuButton
            icon={<EyeOff size={14} />}
            label="隐藏工作表"
            disabled={!canHide}
            onClick={() => {
              close();
              onHide(sheetId);
            }}
          />
          <fieldset className="work-spreadsheet-sheet-color-row">
            <legend>
              <Palette size={14} aria-hidden="true" />
              <span>标签颜色</span>
            </legend>
            <div>
              {spreadsheetSheetColors.map(({ color, label }) => (
                <button
                  key={color}
                  type="button"
                  aria-label={`${label}标签`}
                  title={`${label}标签`}
                  className={sheet.color === color ? 'active' : undefined}
                  style={{ backgroundColor: color }}
                  onClick={() => {
                    close();
                    onSetColor(sheetId, color);
                  }}
                />
              ))}
              <button
                type="button"
                aria-label="清除标签颜色"
                title="清除标签颜色"
                className="clear"
                onClick={() => {
                  close();
                  onSetColor(sheetId, null);
                }}
              />
            </div>
          </fieldset>
          <hr />
          <SheetMenuButton
            icon={<ArrowLeft size={14} />}
            label="向左移动"
            disabled={!canMoveLeft}
            onClick={() => {
              close();
              onMove(sheetId, -1);
            }}
          />
          <SheetMenuButton
            icon={<ArrowRight size={14} />}
            label="向右移动"
            disabled={!canMoveRight}
            onClick={() => {
              close();
              onMove(sheetId, 1);
            }}
          />
          <hr />
          <SheetMenuButton
            icon={<Trash2 size={14} />}
            label="删除工作表"
            danger
            disabled={!canDelete}
            onClick={() => {
              close();
              onDelete(sheetId);
            }}
          />
        </>
      )}
    </Popover>
  );
}

function SheetMenuButton({
  icon,
  label,
  danger = false,
  disabled = false,
  onClick,
}: {
  icon: ReactNode;
  label: string;
  danger?: boolean;
  disabled?: boolean;
  onClick(): void;
}) {
  return (
    <button
      type="button"
      role="menuitem"
      className={danger ? 'danger' : undefined}
      disabled={disabled}
      onClick={onClick}
    >
      {icon}
      <span>{label}</span>
    </button>
  );
}

function moveSpreadsheetSheetMenuFocus(event: KeyboardEvent<HTMLElement>) {
  if (!['ArrowDown', 'ArrowUp', 'Home', 'End'].includes(event.key)) return;
  const buttons = [
    ...event.currentTarget.querySelectorAll<HTMLButtonElement>(
      'button:not(:disabled)',
    ),
  ];
  if (!buttons.length) return;
  event.preventDefault();
  const current = buttons.indexOf(document.activeElement as HTMLButtonElement);
  const next =
    event.key === 'Home'
      ? 0
      : event.key === 'End'
        ? buttons.length - 1
        : event.key === 'ArrowDown'
          ? (current + 1 + buttons.length) % buttons.length
          : (current - 1 + buttons.length) % buttons.length;
  buttons[next]?.focus();
}
