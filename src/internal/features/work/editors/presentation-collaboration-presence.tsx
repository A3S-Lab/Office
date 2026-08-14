import type { WorkSlideElement } from '../work-types';
import {
  officePresenceColorStyle,
  useOfficeRemoteParticipants,
} from './office-collaboration-presence-ui';

export function PresentationCollaborationPresenceLayer({
  elements,
  slideId,
}: {
  elements: readonly WorkSlideElement[];
  slideId: string;
}) {
  const participants = useOfficeRemoteParticipants().filter(
    (participant) =>
      participant.location?.kind === 'presentation' &&
      participant.location.slideId === slideId,
  );
  if (!participants.length) return null;

  return (
    <div className="work-presentation-remote-presence-layer" aria-hidden="true">
      {participants.flatMap((participant, participantIndex) => {
        const location = participant.location;
        if (location?.kind !== 'presentation') return [];
        const selected = location.elementIds.flatMap((elementId) => {
          const element = elements.find(
            (candidate) => candidate.id === elementId,
          );
          return element ? [element] : [];
        });
        if (!selected.length) {
          return [
            <span
              className="work-presentation-remote-slide-marker"
              data-activity={participant.activity}
              data-participant-id={participant.actor.id}
              key={`${participant.presenceId}:slide`}
              style={{
                ...officePresenceColorStyle(participant),
                top: 8 + participantIndex * 22,
              }}
            >
              {participant.actor.name}
            </span>,
          ];
        }
        return selected.map((element, index) => (
          <span
            className="work-presentation-remote-element"
            data-activity={participant.activity}
            data-participant-id={participant.actor.id}
            data-remote-slide-element-id={element.id}
            key={`${participant.presenceId}:${element.id}`}
            style={{
              ...officePresenceColorStyle(participant),
              left: `${element.x}%`,
              top: `${element.y}%`,
              width: `${element.width}%`,
              height: `${element.height}%`,
              transform: element.rotation
                ? `rotate(${element.rotation}deg)`
                : undefined,
            }}
          >
            {index === 0 && <span>{participant.actor.name}</span>}
          </span>
        ));
      })}
    </div>
  );
}
