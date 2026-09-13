import type {
  SpreadsheetSortKey,
  SpreadsheetSortOrientation,
} from './spreadsheet-sort';
import {
  parseSpreadsheetSortAppearanceTargetValue,
  spreadsheetSortAppearanceTargetLabel,
  spreadsheetSortAppearanceTargets,
  spreadsheetSortAppearanceTargetsEqual,
  spreadsheetSortAppearanceTargetValue,
  type SpreadsheetSortAppearanceField,
  type SpreadsheetSortAppearanceKind,
  type SpreadsheetSortAppearanceTarget,
} from './spreadsheet-sort-appearance';
import {
  spreadsheetSortCustomListsEqual,
  type SpreadsheetSortCustomList,
} from './spreadsheet-sort-custom-list';
import { OfficeSelect, type OfficeSelectOption } from './office-controls';

const CREATE_CUSTOM_LIST_ORDER = 'create-custom-list';
const CUSTOM_LIST_ORDER_PREFIX = 'custom-list:';

export function SpreadsheetSortOrderControls({
  appearanceField,
  customLists,
  level,
  orientation,
  sortKey,
  onBeginCustomListEdit,
  onChange,
}: {
  appearanceField: SpreadsheetSortAppearanceField | undefined;
  customLists: readonly SpreadsheetSortCustomList[];
  level: number;
  orientation: SpreadsheetSortOrientation;
  sortKey: SpreadsheetSortKey;
  onBeginCustomListEdit: (entries?: readonly string[]) => void;
  onChange: (key: SpreadsheetSortKey) => void;
}) {
  const sortOn = sortKey.sortOn ?? 'values';
  const appearanceTarget = spreadsheetSortKeyAppearanceTarget(sortKey);
  const appearancePosition = sortKey.position === 'last' ? 'last' : 'first';
  const appearanceTargets =
    sortOn === 'values'
      ? []
      : spreadsheetSortAppearanceTargets(appearanceField, sortOn);

  const sortOnOptions: readonly OfficeSelectOption<string>[] = [
    { value: 'values', label: '值' },
    {
      value: 'cell-color',
      label: '单元格颜色',
      disabled:
        !spreadsheetSortAppearanceTargets(appearanceField, 'cell-color').length,
    },
    {
      value: 'font-color',
      label: '字体颜色',
      disabled:
        !spreadsheetSortAppearanceTargets(appearanceField, 'font-color').length,
    },
    {
      value: 'icon',
      label: '条件格式图标',
      disabled: !spreadsheetSortAppearanceTargets(appearanceField, 'icon')
        .length,
    },
  ];

  const appearanceTargetOptions: readonly OfficeSelectOption<string>[] =
    appearanceTargets.map((target) => ({
      value: spreadsheetSortAppearanceTargetValue(target),
      label: spreadsheetSortAppearanceTargetLabel(target),
    }));

  const positionOptions: readonly OfficeSelectOption<'first' | 'last'>[] = [
    {
      value: 'first',
      label: spreadsheetSortPositionLabel(orientation, 'first'),
    },
    {
      value: 'last',
      label: spreadsheetSortPositionLabel(orientation, 'last'),
    },
  ];

  return (
    <div
      className="work-spreadsheet-sort-order-controls"
      data-appearance={sortOn === 'values' ? undefined : 'true'}
    >
      <label>
        <span>排序依据</span>
        <OfficeSelect
          ariaLabel={`排序条件 ${level} 排序依据`}
          value={sortOn}
          options={sortOnOptions}
          onValueChange={(next) => {
            if (next === 'values') {
              onChange({ index: sortKey.index, direction: 'ascending' });
              return;
            }
            if (!isSpreadsheetSortAppearanceKind(next)) return;
            const target = spreadsheetSortAppearanceTargets(
              appearanceField,
              next,
            )[0];
            if (target) {
              onChange(
                spreadsheetSortAppearanceKey(sortKey.index, target, 'first'),
              );
            }
          }}
        />
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
            <OfficeSelect
              ariaLabel={`排序条件 ${level} 目标外观`}
              value={
                appearanceTarget
                  ? spreadsheetSortAppearanceTargetValue(appearanceTarget)
                  : (appearanceTargetOptions[0]?.value ?? '')
              }
              options={appearanceTargetOptions}
              disabled={appearanceTargetOptions.length === 0}
              onValueChange={(raw) => {
                const target = parseSpreadsheetSortAppearanceTargetValue(raw);
                if (!target || target.kind !== sortOn) return;
                onChange(
                  spreadsheetSortAppearanceKey(
                    sortKey.index,
                    target,
                    appearancePosition,
                  ),
                );
              }}
            />
          </label>
          <label>
            <span>位置</span>
            <OfficeSelect<'first' | 'last'>
              ariaLabel={`排序条件 ${level} 位置`}
              value={appearancePosition}
              options={positionOptions}
              disabled={!appearanceTarget}
              onValueChange={(position) => {
                if (!appearanceTarget) return;
                onChange(
                  spreadsheetSortAppearanceKey(
                    sortKey.index,
                    appearanceTarget,
                    position,
                  ),
                );
              }}
            />
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
                {spreadsheetSortPositionLabel(orientation, appearancePosition)}
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
  const orderOptions: OfficeSelectOption<string>[] = [
    { value: 'ascending', label: '升序（A 到 Z）' },
    { value: 'descending', label: '降序（Z 到 A）' },
  ];
  customLists.forEach((customList, customListIndex) => {
    if (customList.source === 'built-in') {
      orderOptions.push({
        value: `${CUSTOM_LIST_ORDER_PREFIX}${customListIndex}`,
        label: customList.label,
        group: '内置序列',
      });
    }
  });
  customLists.forEach((customList, customListIndex) => {
    if (customList.source === 'stored') {
      orderOptions.push({
        value: `${CUSTOM_LIST_ORDER_PREFIX}${customListIndex}`,
        label: customList.label,
        group: '已保存的序列',
      });
    }
  });
  customLists.forEach((customList, customListIndex) => {
    if (customList.source === 'session') {
      orderOptions.push({
        value: `${CUSTOM_LIST_ORDER_PREFIX}${customListIndex}`,
        label: customList.label,
        group: '本次会话的序列',
      });
    }
  });
  orderOptions.push({
    value: CREATE_CUSTOM_LIST_ORDER,
    label: '新建自定义序列…',
  });

  return (
    <label>
      <span>次序</span>
      <OfficeSelect
        ariaLabel={`排序条件 ${level} 次序`}
        value={spreadsheetSortOrderValue(sortKey, customLists)}
        options={orderOptions}
        onValueChange={(order) => {
          if (order === CREATE_CUSTOM_LIST_ORDER) {
            onBeginCustomListEdit(sortKey.customList);
            return;
          }
          if (order === 'ascending' || order === 'descending') {
            onChange({ index: sortKey.index, direction: order });
            return;
          }
          const customListIndex = Number(
            order.slice(CUSTOM_LIST_ORDER_PREFIX.length),
          );
          const customList = customLists[customListIndex];
          if (!customList) return;
          onChange({
            index: sortKey.index,
            customList: [...customList.entries],
          });
        }}
      />
    </label>
  );
}

export function spreadsheetSortAppearanceKey(
  index: number,
  target: SpreadsheetSortAppearanceTarget,
  position: 'first' | 'last',
): SpreadsheetSortKey {
  return target.kind === 'icon'
    ? { index, sortOn: 'icon', icon: { ...target.icon }, position }
    : { index, sortOn: target.kind, color: target.color, position };
}

export function nextSpreadsheetSortKey(
  keys: readonly SpreadsheetSortKey[],
  fields: readonly { index: number }[],
  appearanceFields: readonly SpreadsheetSortAppearanceField[],
): SpreadsheetSortKey | null {
  const usedIndices = new Set(keys.map((key) => key.index));
  const unused = fields.find((field) => !usedIndices.has(field.index));
  if (unused) return { index: unused.index, direction: 'ascending' };

  const defaultKinds: readonly SpreadsheetSortAppearanceKind[] = [
    'cell-color',
    'font-color',
    'icon',
  ];
  for (const key of [...keys].reverse()) {
    const field = appearanceFields.find(
      (candidate) => candidate.index === key.index,
    );
    const currentTarget = spreadsheetSortKeyAppearanceTarget(key);
    const kinds = currentTarget
      ? [
          currentTarget.kind,
          ...defaultKinds.filter((kind) => kind !== currentTarget.kind),
        ]
      : defaultKinds;
    for (const kind of kinds) {
      const target = spreadsheetSortAppearanceTargets(field, kind).find(
        (candidate) =>
          !keys.some((existing) => {
            const existingTarget = spreadsheetSortKeyAppearanceTarget(existing);
            return (
              existing.index === key.index &&
              existingTarget !== null &&
              spreadsheetSortAppearanceTargetsEqual(existingTarget, candidate)
            );
          }),
      );
      if (target) {
        return spreadsheetSortAppearanceKey(
          key.index,
          target,
          currentTarget?.kind === kind && key.position === 'last'
            ? 'last'
            : 'first',
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

function spreadsheetSortPositionLabel(
  orientation: SpreadsheetSortOrientation,
  position: 'first' | 'last',
): string {
  if (orientation === 'left-to-right') {
    return position === 'first' ? '置于左侧' : '置于右侧';
  }
  return position === 'first' ? '置于顶端' : '置于底端';
}

function isSpreadsheetSortAppearanceKind(
  value: string,
): value is SpreadsheetSortAppearanceKind {
  return value === 'cell-color' || value === 'font-color' || value === 'icon';
}
