import { Bot, UsersRound } from 'lucide-react';
import { type CSSProperties, useEffect, useMemo, useState } from 'react';
import type {
  WorkOfficeCollaborationParticipant,
  WorkOfficeCollaborationPresence,
  WorkOfficeCollaborationPresenceActivity,
  WorkOfficeCollaborationPresenceSnapshot,
} from '../../../collaboration/office-collaboration-presence';
import { Popover } from '../../../design-system/primitives';
import { useOfficeCollaborationPresence } from './office-collaboration-presence-context';

const PRESENCE_COLORS = [
  '#1d4ed8',
  '#6d28d9',
  '#be123c',
  '#047857',
  '#a16207',
  '#0369a1',
  '#9d174d',
  '#4338ca',
] as const;

type PresenceAvatarStyle = CSSProperties & {
  '--work-office-presence-color': string;
};

export function WorkOfficeCollaborationParticipants({
  variant = 'status',
}: {
  variant?: 'status' | 'toolbar';
}) {
  const presence = useOfficeCollaborationPresence();
  const snapshot = useOfficePresenceSnapshot(presence);
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
              <PresenceAvatar
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
        {participants.map((participant) => (
          <li
            key={participant.presenceId}
            data-collaboration-participant={participant.actor.id}
            data-activity={participant.activity}
            data-local={participant.local ? 'true' : undefined}
          >
            <PresenceAvatar participant={participant} />
            <span className="work-office-collaboration-person">
              <span className="work-office-collaboration-name">
                <strong>{participant.actor.name}</strong>
                {participant.local && <small>你</small>}
                {participant.actor.kind === 'agent' && <small>Agent</small>}
                {participant.actor.kind === 'system' && <small>系统</small>}
              </span>
              <span className="work-office-collaboration-detail">
                {presenceActivityLabel(participant.activity)} ·{' '}
                {presenceModeLabel(participant.mode)} ·{' '}
                {presenceLocationLabel(participant)}
              </span>
            </span>
          </li>
        ))}
      </ul>
    </Popover>
  );
}

function useOfficePresenceSnapshot(
  presence: WorkOfficeCollaborationPresence | undefined,
): WorkOfficeCollaborationPresenceSnapshot | undefined {
  const [state, setState] = useState<{
    owner?: WorkOfficeCollaborationPresence;
    snapshot?: WorkOfficeCollaborationPresenceSnapshot;
  }>(() => ({ owner: presence, snapshot: presence?.snapshot() }));

  useEffect(() => {
    if (!presence) {
      setState({ owner: undefined, snapshot: undefined });
      return;
    }
    setState({ owner: presence, snapshot: presence.snapshot() });
    return presence.subscribe((snapshot) =>
      setState({ owner: presence, snapshot }),
    );
  }, [presence]);

  return state.owner === presence ? state.snapshot : presence?.snapshot();
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

function PresenceAvatar({
  compact = false,
  participant,
}: {
  compact?: boolean;
  participant: WorkOfficeCollaborationParticipant;
}) {
  const avatarUrl = safeAvatarUrl(participant.actor.avatarUrl);
  const style: PresenceAvatarStyle = {
    '--work-office-presence-color': presenceColor(participant),
  };
  return (
    <span
      className="work-office-collaboration-avatar"
      data-activity={participant.activity}
      data-compact={compact ? 'true' : undefined}
      style={style}
      aria-hidden="true"
    >
      {participant.actor.kind === 'agent' ? (
        <Bot size={compact ? 10 : 12} />
      ) : (
        <span>{participantInitial(participant.actor.name)}</span>
      )}
      {avatarUrl && (
        <img
          src={avatarUrl}
          alt=""
          loading="lazy"
          referrerPolicy="no-referrer"
          onError={(event) => {
            event.currentTarget.hidden = true;
          }}
        />
      )}
    </span>
  );
}

function participantInitial(name: string): string {
  return Array.from(name.trim())[0]?.toLocaleUpperCase() ?? '?';
}

function presenceColor(
  participant: WorkOfficeCollaborationParticipant,
): string {
  const requested = participant.actor.color;
  if (requested && safeCssColor(requested)) return requested;
  let hash = 0;
  for (const character of participant.actor.id) {
    hash = (hash * 31 + (character.codePointAt(0) ?? 0)) >>> 0;
  }
  return PRESENCE_COLORS[hash % PRESENCE_COLORS.length] ?? PRESENCE_COLORS[0];
}

function safeCssColor(value: string): boolean {
  if (/[;{}]|url\(|var\(/i.test(value)) return false;
  if (/^#(?:[\da-f]{3}|[\da-f]{4}|[\da-f]{6}|[\da-f]{8})$/i.test(value)) {
    return true;
  }
  return typeof CSS !== 'undefined' && CSS.supports('color', value);
}

function safeAvatarUrl(value: string | undefined): string | undefined {
  if (!value || typeof URL === 'undefined') return undefined;
  try {
    const url = new URL(value, 'https://a3s.invalid/');
    if (
      url.protocol === 'https:' ||
      url.protocol === 'http:' ||
      url.protocol === 'blob:' ||
      (url.protocol === 'data:' && value.startsWith('data:image/'))
    ) {
      return value;
    }
  } catch {
    return undefined;
  }
  return undefined;
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
