import type { SpreadsheetSortKey } from './spreadsheet-sort';
import {
  parseSpreadsheetSortAppearanceTargetValue,
  spreadsheetSortAppearanceTargetLabel,
  spreadsheetSortAppearanceTargets,
  spreadsheetSortAppearanceTargetsEqual,
  spreadsheetSortAppearanceTargetValue,
  type SpreadsheetSortAppearanceColumn,
  type SpreadsheetSortAppearanceKind,
  type SpreadsheetSortAppearanceTarget,
} from './spreadsheet-sort-appearance';
import {
  spreadsheetSortCustomListsEqual,
  type SpreadsheetSortCustomList,
} from './spreadsheet-sort-custom-list';

const CREATE_CUSTOM_LIST_ORDER = 'create-custom-list';
const CUSTOM_LIST_ORDER_PREFIX = 'custom-list:';

export function SpreadsheetSortOrderControls({
  appearanceColumn,
  customLists,
  level,
  sortKey,
  onBeginCustomListEdit,
  onChange,
}: {
  appearanceColumn: SpreadsheetSortAppearanceColumn | undefined;
  customLists: readonly SpreadsheetSortCustomList[];
  level: number;
  sortKey: SpreadsheetSortKey;
  onBeginCustomListEdit: (entries?: readonly string[]) => void;
  onChange: (key: SpreadsheetSortKey) => void;
}) {
  const sortOn = sortKey.sortOn ?? 'values';
  const appearanceTarget = spreadsheetSortKeyAppearanceTarget(sortKey);
  const appearancePosition = sortKey.position === 'bottom' ? 'bottom' : 'top';
  const appearanceTargets =
    sortOn === 'values'
      ? []
      : spreadsheetSortAppearanceTargets(appearanceColumn, sortOn);

  return (
    <div
      className="work-spreadsheet-sort-order-controls"
      data-appearance={sortOn === 'values' ? undefined : 'true'}
    >
      <label>
        <span>排序依据</span>
        <select
          aria-label={`排序条件 ${level} 排序依据`}
          value={sortOn}
          onChange={(event) => {
            const next = event.currentTarget.value;
            if (next === 'values') {
              onChange({ column: sortKey.column, direction: 'ascending' });
              return;
            }
            if (!isSpreadsheetSortAppearanceKind(next)) return;
            const target = spreadsheetSortAppearanceTargets(
              appearanceColumn,
              next,
            )[0];
            if (target) {
              onChange(
                spreadsheetSortAppearanceKey(sortKey.column, target, 'top'),
              );
            }
          }}
        >
          <option value="values">值</option>
          <option
            value="cell-color"
            disabled={
              !spreadsheetSortAppearanceTargets(appearanceColumn, 'cell-color')
                .length
            }
          >
            单元格颜色
          </option>
          <option
            value="font-color"
            disabled={
              !spreadsheetSortAppearanceTargets(appearanceColumn, 'font-color')
                .length
            }
          >
            字体颜色
          </option>
          <option
            value="icon"
            disabled={
              !spreadsheetSortAppearanceTargets(appearanceColumn, 'icon').length
            }
          >
            条件格式图标
          </option>
        </select>
      </label>

      {sortOn === 'values' ? (
        <SpreadsheetSortValueOrder
          customLists={customLists}
          level={level}
          sortKey={sortKey}
          onBeginCustomListEdit={onBeginCustomListEdit}
          onChange={onChange}
        />
      ) : (
        <>
          <label>
            <span>次序</span>
            <select
              aria-label={`排序条件 ${level} 目标外观`}
              value={
                appearanceTarget
                  ? spreadsheetSortAppearanceTargetValue(appearanceTarget)
                  : ''
              }
              onChange={(event) => {
                const target = parseSpreadsheetSortAppearanceTargetValue(
                  event.currentTarget.value,
                );
                if (!target || target.kind !== sortOn) return;
                onChange(
                  spreadsheetSortAppearanceKey(
                    sortKey.column,
                    target,
                    appearancePosition,
                  ),
                );
              }}
            >
              {appearanceTargets.map((target) => (
                <option
                  key={spreadsheetSortAppearanceTargetValue(target)}
                  value={spreadsheetSortAppearanceTargetValue(target)}
                >
                  {spreadsheetSortAppearanceTargetLabel(target)}
                </option>
              ))}
            </select>
          </label>
          <label>
            <span>位置</span>
            <select
              aria-label={`排序条件 ${level} 位置`}
              value={appearancePosition}
              onChange={(event) => {
                if (!appearanceTarget) return;
                const position = event.currentTarget.value;
                if (position !== 'top' && position !== 'bottom') return;
                onChange(
                  spreadsheetSortAppearanceKey(
                    sortKey.column,
                    appearanceTarget,
                    position,
                  ),
                );
              }}
            >
              <option value="top">置于顶端</option>
              <option value="bottom">置于底端</option>
            </select>
          </label>
          {appearanceTarget ? (
            <div className="work-spreadsheet-sort-appearance-preview">
              {appearanceTarget.kind === 'icon' ? (
                <span aria-hidden="true">
                  {spreadsheetSortAppearanceTargetLabel(appearanceTarget).split(
                    ' ',
                  )[0] ?? ''}
                </span>
              ) : (
                <span
                  aria-hidden="true"
                  className="work-spreadsheet-sort-color-swatch"
                  data-empty={
                    appearanceTarget.color === null ? 'true' : undefined
                  }
                  style={
                    appearanceTarget.color
                      ? { backgroundColor: appearanceTarget.color }
                      : undefined
                  }
                />
              )}
              <span>
                {spreadsheetSortAppearanceTargetLabel(appearanceTarget)}，
                {appearancePosition === 'top' ? '置于顶端' : '置于底端'}
              </span>
            </div>
          ) : null}
        </>
      )}
    </div>
  );
}

function SpreadsheetSortValueOrder({
  customLists,
  level,
  sortKey,
  onBeginCustomListEdit,
  onChange,
}: {
  customLists: readonly SpreadsheetSortCustomList[];
  level: number;
  sortKey: SpreadsheetSortKey;
  onBeginCustomListEdit: (entries?: readonly string[]) => void;
  onChange: (key: SpreadsheetSortKey) => void;
}) {
  return (
    <label>
      <span>次序</span>
      <select
        aria-label={`排序条件 ${level} 次序`}
        value={spreadsheetSortOrderValue(sortKey, customLists)}
        onChange={(event) => {
          const order = event.currentTarget.value;
          if (order === CREATE_CUSTOM_LIST_ORDER) {
            onBeginCustomListEdit(sortKey.customList);
            return;
          }
          if (order === 'ascending' || order === 'descending') {
            onChange({ column: sortKey.column, direction: order });
            return;
          }
          const customListIndex = Number(
            order.slice(CUSTOM_LIST_ORDER_PREFIX.length),
          );
          const customList = customLists[customListIndex];
          if (!customList) return;
          onChange({
            column: sortKey.column,
            customList: [...customList.entries],
          });
        }}
      >
        <option value="ascending">升序（A 到 Z）</option>
        <option value="descending">降序（Z 到 A）</option>
        <optgroup label="内置序列">
          {customLists.map((customList, customListIndex) =>
            customList.source === 'built-in' ? (
              <option
                key={`built-in:${customListIndex}`}
                value={`${CUSTOM_LIST_ORDER_PREFIX}${customListIndex}`}
              >
                {customList.label}
              </option>
            ) : null,
          )}
        </optgroup>
        {customLists.some((list) => list.source === 'session') ? (
          <optgroup label="本次会话的序列">
            {customLists.map((customList, customListIndex) =>
              customList.source === 'session' ? (
                <option
                  key={`session:${customListIndex}`}
                  value={`${CUSTOM_LIST_ORDER_PREFIX}${customListIndex}`}
                >
                  {customList.label}
                </option>
              ) : null,
            )}
          </optgroup>
        ) : null}
        <option value={CREATE_CUSTOM_LIST_ORDER}>新建自定义序列…</option>
      </select>
    </label>
  );
}

export function spreadsheetSortAppearanceKey(
  column: number,
  target: SpreadsheetSortAppearanceTarget,
  position: 'top' | 'bottom',
): SpreadsheetSortKey {
  return target.kind === 'icon'
    ? { column, sortOn: 'icon', icon: { ...target.icon }, position }
    : { column, sortOn: target.kind, color: target.color, position };
}

export function nextSpreadsheetSortKey(
  keys: readonly SpreadsheetSortKey[],
  columns: readonly { column: number }[],
  appearanceColumns: readonly SpreadsheetSortAppearanceColumn[],
): SpreadsheetSortKey | null {
  const usedColumns = new Set(keys.map((key) => key.column));
  const unused = columns.find((column) => !usedColumns.has(column.column));
  if (unused) return { column: unused.column, direction: 'ascending' };

  const defaultKinds: readonly SpreadsheetSortAppearanceKind[] = [
    'cell-color',
    'font-color',
    'icon',
  ];
  for (const key of [...keys].reverse()) {
    const column = appearanceColumns.find(
      (candidate) => candidate.column === key.column,
    );
    const currentTarget = spreadsheetSortKeyAppearanceTarget(key);
    const kinds = currentTarget
      ? [
          currentTarget.kind,
          ...defaultKinds.filter((kind) => kind !== currentTarget.kind),
        ]
      : defaultKinds;
    for (const kind of kinds) {
      const target = spreadsheetSortAppearanceTargets(column, kind).find(
        (candidate) =>
          !keys.some((existing) => {
            const existingTarget = spreadsheetSortKeyAppearanceTarget(existing);
            return (
              existing.column === key.column &&
              existingTarget !== null &&
              spreadsheetSortAppearanceTargetsEqual(existingTarget, candidate)
            );
          }),
      );
      if (target) {
        return spreadsheetSortAppearanceKey(
          key.column,
          target,
          currentTarget?.kind === kind && key.position === 'bottom'
            ? 'bottom'
            : 'top',
        );
      }
    }
  }
  return null;
}

function spreadsheetSortKeyAppearanceTarget(
  key: SpreadsheetSortKey,
): SpreadsheetSortAppearanceTarget | null {
  if (key.sortOn === 'cell-color' || key.sortOn === 'font-color') {
    return { kind: key.sortOn, color: key.color };
  }
  return key.sortOn === 'icon' ? { kind: 'icon', icon: { ...key.icon } } : null;
}

function spreadsheetSortOrderValue(
  key: SpreadsheetSortKey,
  customLists: readonly SpreadsheetSortCustomList[],
): string {
  if (key.customList === undefined) return key.direction ?? 'ascending';
  const index = customLists.findIndex((customList) =>
    spreadsheetSortCustomListsEqual(customList.entries, key.customList ?? []),
  );
  return index < 0
    ? CREATE_CUSTOM_LIST_ORDER
    : `${CUSTOM_LIST_ORDER_PREFIX}${index}`;
}

function isSpreadsheetSortAppearanceKind(
  value: string,
): value is SpreadsheetSortAppearanceKind {
  return value === 'cell-color' || value === 'font-color' || value === 'icon';
}
