import { LocateFixed, UsersRound } from 'lucide-react';
import { useMemo } from 'react';
import type {
  WorkOfficeCollaborationParticipant,
  WorkOfficeCollaborationPresenceActivity,
} from '../../../collaboration/office-collaboration-presence';
import { Popover } from '../../../design-system/primitives';
import { useOfficeCollaborationParticipantNavigation } from './office-collaboration-presence-context';
import {
  OfficePresenceAvatar,
  useOfficePresenceSnapshot,
} from './office-collaboration-presence-ui';

export function WorkOfficeCollaborationParticipants({
  variant = 'status',
}: {
  variant?: 'status' | 'toolbar';
}) {
  const snapshot = useOfficePresenceSnapshot();
  const navigateToParticipant = useOfficeCollaborationParticipantNavigation();
  const participants = useMemo(
    () => orderedParticipants(snapshot?.participants ?? []),
    [snapshot],
  );

  if (!participants.length) return null;

  const count = participants.length;
  const summary = count === 1 ? '仅你在线' : `${count} 位协作者`;
  const visibleParticipants = participants.slice(0, 3);
  const overflow = Math.max(0, count - visibleParticipants.length);

  return (
    <Popover
      label={`查看协作者，${summary}`}
      panelLabel="协作者"
      panelRole="dialog"
      className="work-office-collaboration"
      panelClassName="work-office-collaboration-popover"
      placement={variant === 'status' ? 'top-end' : 'bottom-end'}
      portal
      focusFirstOnOpen={participants.some(
        (participant) => !participant.local && Boolean(participant.location),
      )}
      trigger={(triggerProps, { open }) => (
        <button
          {...triggerProps}
          className="work-office-collaboration-trigger"
          data-collaboration-count={count}
          data-variant={variant}
          title={`协作者：${summary}`}
        >
          <span className="work-office-collaboration-stack" aria-hidden="true">
            {visibleParticipants.map((participant) => (
              <OfficePresenceAvatar
                key={participant.presenceId}
                participant={participant}
                compact
              />
            ))}
            {overflow > 0 && (
              <span className="work-office-collaboration-overflow">
                +{overflow}
              </span>
            )}
          </span>
          <span className="work-office-collaboration-summary">
            {variant === 'status' ? summary : count}
          </span>
          <span className="sr-only" aria-live="polite">
            {open ? '协作者列表已打开' : `${summary}在线`}
          </span>
        </button>
      )}
    >
      {(close) => (
        <>
          <header className="work-office-collaboration-heading">
            <span className="work-office-collaboration-heading-icon">
              <UsersRound size={16} aria-hidden="true" />
            </span>
            <span>
              <strong>协作者</strong>
              <small>{count} 个在线会话</small>
            </span>
          </header>
          <ul className="work-office-collaboration-list">
            {participants.map((participant) => {
              const locationLabel = presenceLocationLabel(participant);
              const navigable =
                !participant.local &&
                Boolean(participant.location) &&
                Boolean(navigateToParticipant);
              const content = (
                <>
                  <OfficePresenceAvatar participant={participant} />
                  <span className="work-office-collaboration-person">
                    <span className="work-office-collaboration-name">
                      <strong>{participant.actor.name}</strong>
                      {participant.local && <small>你</small>}
                      {participant.actor.kind === 'agent' && (
                        <small>Agent</small>
                      )}
                      {participant.actor.kind === 'system' && (
                        <small>系统</small>
                      )}
                    </span>
                    <span className="work-office-collaboration-detail">
                      {presenceActivityLabel(participant.activity)} ·{' '}
                      {presenceModeLabel(participant.mode)} · {locationLabel}
                    </span>
                  </span>
                  {navigable && (
                    <LocateFixed
                      className="work-office-collaboration-locate"
                      size={14}
                      aria-hidden="true"
                    />
                  )}
                </>
              );
              return (
                <li
                  key={participant.presenceId}
                  data-collaboration-participant={participant.actor.id}
                  data-activity={participant.activity}
                  data-local={participant.local ? 'true' : undefined}
                  data-navigable={navigable ? 'true' : undefined}
                >
                  {navigable ? (
                    <button
                      type="button"
                      className="work-office-collaboration-participant"
                      aria-label={`跳转到 ${participant.actor.name} 的位置，${locationLabel}`}
                      onClick={() => {
                        close();
                        requestAnimationFrame(() =>
                          navigateToParticipant?.(participant),
                        );
                      }}
                    >
                      {content}
                    </button>
                  ) : (
                    <span className="work-office-collaboration-participant">
                      {content}
                    </span>
                  )}
                </li>
              );
            })}
          </ul>
        </>
      )}
    </Popover>
  );
}

function orderedParticipants(
  participants: readonly WorkOfficeCollaborationParticipant[],
): readonly WorkOfficeCollaborationParticipant[] {
  return [...participants].sort(
    (left, right) =>
      Number(right.local) - Number(left.local) ||
      presenceActivityOrder(left.activity) -
        presenceActivityOrder(right.activity) ||
      left.actor.name.localeCompare(right.actor.name) ||
      left.clientId - right.clientId,
  );
}

function presenceActivityOrder(
  activity: WorkOfficeCollaborationPresenceActivity,
): number {
  return activity === 'active' ? 0 : activity === 'idle' ? 1 : 2;
}

function presenceActivityLabel(
  activity: WorkOfficeCollaborationPresenceActivity,
): string {
  return activity === 'active'
    ? '正在使用'
    : activity === 'idle'
      ? '空闲'
      : '离开';
}

function presenceModeLabel(
  mode: WorkOfficeCollaborationParticipant['mode'],
): string {
  if (mode === 'edit') return '可编辑';
  if (mode === 'suggest') return '建议';
  if (mode === 'comment') return '评论';
  return '只读';
}

function presenceLocationLabel(
  participant: WorkOfficeCollaborationParticipant,
): string {
  const location = participant.location;
  if (!location) return '未共享位置';
  switch (location.kind) {
    case 'document':
    case 'markdown': {
      const selected = Math.abs(location.head - location.anchor);
      return selected > 0
        ? `已选择 ${selected} 个位置`
        : `位置 ${location.head + 1}`;
    }
    case 'spreadsheet': {
      const active = location.activeCell ?? {
        row: location.ranges[0]?.startRow ?? 0,
        column: location.ranges[0]?.startColumn ?? 0,
      };
      return `R${active.row + 1}C${active.column + 1}`;
    }
    case 'presentation':
      return location.elementIds.length > 0
        ? `幻灯片中的 ${location.elementIds.length} 个对象`
        : '当前幻灯片';
    case 'pdf':
      return `第 ${location.pageIndex + 1} 页`;
  }
}
