import { LocateFixed, UsersRound } from 'lucide-react';
import { officeMessage, resolveOfficeMessages } from '../../../i18n/office-locale';
import { useOfficeMessages } from './office-messages-context';
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
  const messages = useOfficeMessages();
  const snapshot = useOfficePresenceSnapshot();
  const navigateToParticipant = useOfficeCollaborationParticipantNavigation();
  const participants = useMemo(
    () => orderedParticipants(snapshot?.participants ?? []),
    [snapshot],
  );

  if (!participants.length) return null;

  const count = participants.length;
  const summary = count === 1 ? officeMessage(messages, 'office.collab.youOnly') : officeMessage(messages, 'office.collab.count', { count: String(count) });
  const visibleParticipants = participants.slice(0, 3);
  const overflow = Math.max(0, count - visibleParticipants.length);

  return (
    <Popover
      label={officeMessage(messages, 'office.collab.viewAria', { summary })}
      panelLabel={officeMessage(messages, 'office.collab.panel')}
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
          title={officeMessage(messages, 'office.collab.title', { summary })}
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
            {open ? officeMessage(messages, 'office.collab.listOpen') : officeMessage(messages, 'office.collab.onlineSummary', { summary })}
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
              <strong>{officeMessage(messages, 'office.collab.heading')}</strong>
              <small>{officeMessage(messages, 'office.collab.sessions', { count: String(count) })}</small>
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
                      {participant.local && <small>{officeMessage(messages, 'office.collab.you')}</small>}
                      {participant.actor.kind === 'agent' && (
                        <small>Agent</small>
                      )}
                      {participant.actor.kind === 'system' && (
                        <small>{officeMessage(messages, 'office.collab.system')}</small>
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
                      aria-label={officeMessage(messages, 'office.collab.jumpAria', { name: participant.actor.name, location: locationLabel })}
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
    ? officeMessage(resolveOfficeMessages(), 'office.collab.using')
    : activity === 'idle'
      ? officeMessage(resolveOfficeMessages(), 'office.collab.idle')
      : officeMessage(resolveOfficeMessages(), 'office.collab.away');
}

function presenceModeLabel(
  mode: WorkOfficeCollaborationParticipant['mode'],
): string {
  if (mode === 'edit') return officeMessage(resolveOfficeMessages(), 'office.collab.mode.edit');
  if (mode === 'suggest') return officeMessage(resolveOfficeMessages(), 'office.collab.mode.suggest');
  if (mode === 'comment') return officeMessage(resolveOfficeMessages(), 'office.collab.mode.comment');
  return officeMessage(resolveOfficeMessages(), 'office.collab.mode.read');
}

function presenceLocationLabel(
  participant: WorkOfficeCollaborationParticipant,
): string {
  const location = participant.location;
  if (!location) return officeMessage(resolveOfficeMessages(), 'office.collab.location.none');
  switch (location.kind) {
    case 'document':
    case 'markdown': {
      const selected = Math.abs(location.head - location.anchor);
      return selected > 0
        ? officeMessage(resolveOfficeMessages(), 'office.collab.location.selected', { count: String(selected) })
        : officeMessage(resolveOfficeMessages(), 'office.collab.location.index', { index: String(location.head + 1) });
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
        ? officeMessage(resolveOfficeMessages(), 'office.collab.location.slideObjects', { count: String(location.elementIds.length) })
        : officeMessage(resolveOfficeMessages(), 'office.collab.location.slide');
    case 'pdf':
      return officeMessage(resolveOfficeMessages(), 'office.collab.location.page', { page: String(location.pageIndex + 1) });
  }
}
