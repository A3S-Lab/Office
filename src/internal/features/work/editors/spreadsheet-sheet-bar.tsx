import {
  ArrowLeft,
  ArrowRight,
  Check,
  ChevronDown,
  ChevronLeft,
  ChevronRight,
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
  type ReactNode,
  useEffect,
  useRef,
  useState,
} from 'react';
import { Popover } from '../../../design-system/primitives';
import type { WorkSpreadsheetSheet } from '../work-types';
import { moveOfficeMenuFocus } from './office-menu-keyboard';
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
  const listedSheets = editable ? orderedSheets : visibleSheets;
  const [renamingSheetId, setRenamingSheetId] = useState<string | null>(null);
  const [renameDraft, setRenameDraft] = useState('');
  const renameInputRef = useRef<HTMLInputElement>(null);
  const sheetTabRefs = useRef(new Map<string, HTMLButtonElement>());

  useEffect(() => {
    if (!renamingSheetId) return;
    const frame = requestAnimationFrame(() => {
      renameInputRef.current?.focus({ preventScroll: true });
      renameInputRef.current?.select();
    });
    return () => cancelAnimationFrame(frame);
  }, [renamingSheetId]);

  useEffect(() => {
    const frame = requestAnimationFrame(() => {
      scrollSpreadsheetSheetTabIntoView(
        sheetTabRefs.current.get(activeSheetId),
      );
    });
    return () => cancelAnimationFrame(frame);
  }, [activeSheetId]);

  const beginRename = (sheet: WorkSpreadsheetSheet) => {
    setRenameDraft(sheet.name);
    setRenamingSheetId(sheet.id ?? null);
  };
  const finishRename = (
    sheet: WorkSpreadsheetSheet,
    commit: boolean,
    restoreTabFocus = false,
  ) => {
    if (commit && sheet.id && renameDraft.trim() !== sheet.name) {
      onRename(sheet.id, renameDraft);
    }
    setRenamingSheetId(null);
    if (restoreTabFocus && sheet.id) {
      requestAnimationFrame(() => {
        const tab = sheetTabRefs.current.get(sheet.id ?? '');
        tab?.focus({ preventScroll: true });
        scrollSpreadsheetSheetTabIntoView(tab);
      });
    }
  };
  const activateSheetFromTabKeyboard = (
    sheetId: string,
    key: string,
  ): boolean => {
    const currentIndex = visibleSheets.findIndex(
      (candidate) => candidate.id === sheetId,
    );
    if (currentIndex < 0) return false;
    const targetIndex =
      key === 'Home'
        ? 0
        : key === 'End'
          ? visibleSheets.length - 1
          : key === 'ArrowRight'
            ? (currentIndex + 1) % visibleSheets.length
            : key === 'ArrowLeft'
              ? (currentIndex - 1 + visibleSheets.length) % visibleSheets.length
              : -1;
    const targetId = visibleSheets[targetIndex]?.id;
    if (!targetId || targetId === sheetId) return false;
    onActivate(targetId);
    requestAnimationFrame(() => {
      requestAnimationFrame(() => {
        const target = sheetTabRefs.current.get(targetId);
        target?.focus({ preventScroll: true });
        scrollSpreadsheetSheetTabIntoView(target);
      });
    });
    return true;
  };
  const activateAdjacentSheet = (direction: -1 | 1) => {
    if (visibleSheets.length < 2) return;
    const activeIndex = visibleSheets.findIndex(
      (sheet) => sheet.id === activeSheetId,
    );
    const nextIndex =
      (Math.max(0, activeIndex) + direction + visibleSheets.length) %
      visibleSheets.length;
    const targetId = visibleSheets[nextIndex]?.id;
    if (targetId) onActivate(targetId);
  };

  return (
    <nav
      className="work-spreadsheet-sheet-bar"
      aria-label="工作表"
      data-editable={editable ? 'true' : 'false'}
    >
      <div className="work-spreadsheet-sheet-tools">
        <button
          type="button"
          className="work-spreadsheet-sheet-navigation"
          aria-label="上一个工作表"
          aria-keyshortcuts="Control+PageUp Meta+PageUp"
          title="上一个工作表（Ctrl/⌘+PageUp）"
          disabled={visibleSheets.length < 2}
          onClick={() => activateAdjacentSheet(-1)}
        >
          <ChevronLeft size={14} />
        </button>
        <button
          type="button"
          className="work-spreadsheet-sheet-navigation"
          aria-label="下一个工作表"
          aria-keyshortcuts="Control+PageDown Meta+PageDown"
          title="下一个工作表（Ctrl/⌘+PageDown）"
          disabled={visibleSheets.length < 2}
          onClick={() => activateAdjacentSheet(1)}
        >
          <ChevronRight size={14} />
        </button>
        <Popover
          label="工作表列表"
          panelLabel="工作表列表"
          panelRole="menu"
          placement="top-start"
          portal
          panelClassName="work-office-context-menu work-spreadsheet-sheet-popover work-spreadsheet-sheet-list"
          focusFirstOnOpen
          onPanelKeyDown={moveOfficeMenuFocus}
          trigger={(triggerProps) => (
            <button {...triggerProps} title="查看全部工作表">
              <Layers2 size={14} />
            </button>
          )}
        >
          {(close) =>
            listedSheets.map((sheet) => (
              <button
                key={sheet.id ?? sheet.name}
                type="button"
                role="menuitemradio"
                aria-checked={sheet.id === activeSheetId}
                tabIndex={-1}
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
        {editable && (
          <button
            type="button"
            className="work-spreadsheet-sheet-add"
            aria-label="新建工作表"
            aria-keyshortcuts="Shift+F11 Alt+Shift+F1"
            title="新建工作表（Shift+F11）"
            onClick={onCreate}
          >
            <Plus size={15} />
          </button>
        )}
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
                      event.stopPropagation();
                      finishRename(sheet, true, true);
                    } else if (event.key === 'Escape') {
                      event.preventDefault();
                      event.stopPropagation();
                      finishRename(sheet, false, true);
                    }
                  }}
                />
              ) : (
                <button
                  ref={(node) => {
                    if (node) sheetTabRefs.current.set(sheetId, node);
                    else sheetTabRefs.current.delete(sheetId);
                  }}
                  type="button"
                  role="tab"
                  aria-selected={active}
                  tabIndex={active ? 0 : -1}
                  title={sheet.name}
                  onClick={() => sheetId && onActivate(sheetId)}
                  onDoubleClick={() => editable && beginRename(sheet)}
                  onContextMenu={(event) => {
                    if (!editable) return;
                    event.preventDefault();
                    openSpreadsheetSheetMenu(event.currentTarget);
                  }}
                  onKeyDown={(event) => {
                    if (
                      editable &&
                      (event.key === 'ContextMenu' ||
                        (event.key === 'F10' && event.shiftKey))
                    ) {
                      event.preventDefault();
                      event.stopPropagation();
                      openSpreadsheetSheetMenu(event.currentTarget);
                      return;
                    }
                    if (!activateSheetFromTabKeyboard(sheetId, event.key)) {
                      return;
                    }
                    event.preventDefault();
                    event.stopPropagation();
                  }}
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

function scrollSpreadsheetSheetTabIntoView(
  target: HTMLButtonElement | undefined,
): void {
  target?.scrollIntoView?.({
    behavior: 'auto',
    block: 'nearest',
    inline: 'nearest',
  });
}

function openSpreadsheetSheetMenu(sheetTab: HTMLButtonElement): void {
  const trigger = sheetTab.parentElement?.querySelector<HTMLButtonElement>(
    '.work-spreadsheet-sheet-options',
  );
  if (!trigger) return;
  trigger.focus({ preventScroll: true });
  trigger.click();
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
      panelClassName="work-office-context-menu work-spreadsheet-sheet-popover"
      focusFirstOnOpen
      onPanelKeyDown={moveOfficeMenuFocus}
      trigger={(triggerProps, { open }) => (
        <button
          {...triggerProps}
          className="work-spreadsheet-sheet-options"
          aria-keyshortcuts="Shift+F10"
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
          <fieldset
            className="work-spreadsheet-sheet-color-row"
            aria-label="标签颜色"
          >
            <legend>
              <Palette size={14} aria-hidden="true" />
              <span>标签颜色</span>
            </legend>
            <div>
              {spreadsheetSheetColors.map(({ color, label }) => (
                <button
                  key={color}
                  type="button"
                  role="menuitemradio"
                  aria-checked={sheet.color === color}
                  tabIndex={-1}
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
                role="menuitemradio"
                aria-checked={!sheet.color}
                tabIndex={-1}
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
      tabIndex={-1}
      className={danger ? 'danger' : undefined}
      disabled={disabled}
      onClick={onClick}
    >
      {icon}
      <span>{label}</span>
    </button>
  );
}
