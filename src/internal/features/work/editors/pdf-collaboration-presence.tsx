import {
  OfficePresenceAvatar,
  officePresenceColorStyle,
  useOfficeRemoteParticipants,
} from './office-collaboration-presence-ui';

export function PdfCollaborationPresenceLayer({
  pageIndex,
}: {
  pageIndex: number;
}) {
  const participants = useOfficeRemoteParticipants().filter(
    (participant) =>
      participant.location?.kind === 'pdf' &&
      participant.location.pageIndex === pageIndex,
  );
  if (!participants.length) return null;

  return (
    <div
      className="work-pdf-remote-presence-layer"
      data-pdf-presence-page={pageIndex}
      aria-hidden="true"
    >
      {participants.map((participant) => {
        const location = participant.location;
        if (location?.kind !== 'pdf') return null;
        return (
          <span
            className="work-pdf-remote-participant"
            data-activity={participant.activity}
            data-annotation-id={location.annotationId}
            data-participant-id={participant.actor.id}
            key={participant.presenceId}
            style={officePresenceColorStyle(participant)}
          >
            <OfficePresenceAvatar compact participant={participant} />
            <span>
              <strong>{participant.actor.name}</strong>
              <small>
                {location.annotationId ? '正在查看批注' : '正在查看此页'}
              </small>
            </span>
          </span>
        );
      })}
    </div>
  );
}
