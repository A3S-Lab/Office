import type { KeyboardEvent } from 'react';
import type { WorkSlideElement } from '../work-types';
import {
  type PresentationObjectListItem,
  presentationObjectListItems,
} from './presentation-object-list';

export interface PresentationObjectListProps {
  elements: readonly WorkSlideElement[];
  selectedElementIds: readonly string[];
  listLabel: string;
  emptyLabel: string;
  onSelectElement: (elementId: string, additive: boolean) => void;
}

export function PresentationObjectList({
  elements,
  selectedElementIds,
  listLabel,
  emptyLabel,
  onSelectElement,
}: PresentationObjectListProps) {
  const items = presentationObjectListItems(elements);
  const selectedId =
    selectedElementIds.find((id) => items.some((item) => item.id === id)) ??
    items[0]?.id ??
    null;

  if (!items.length) {
    return (
      <div
        className="work-presentation-object-list"
        data-presentation-object-list
        data-empty="true"
      >
        <p className="work-presentation-object-list-empty">{emptyLabel}</p>
      </div>
    );
  }

  const move = (event: KeyboardEvent<HTMLButtonElement>, nextId: string) => {
    event.preventDefault();
    onSelectElement(nextId, event.shiftKey);
  };

  return (
    <div
      className="work-presentation-object-list"
      data-presentation-object-list
      role="listbox"
      aria-label={listLabel}
      aria-multiselectable="true"
    >
      {items.map((item, index) => {
        const selected = selectedElementIds.includes(item.id);
        const tabIndex =
          selectedId === item.id || (!selectedId && index === 0) ? 0 : -1;
        return (
          <button
            key={item.id}
            type="button"
            role="option"
            className="work-presentation-object-list-item"
            data-presentation-object-id={item.id}
            data-presentation-object-type={item.type}
            aria-selected={selected}
            tabIndex={tabIndex}
            onClick={(event) =>
              onSelectElement(item.id, event.shiftKey || event.metaKey)
            }
            onKeyDown={(event) => {
              if (event.key === 'ArrowDown' || event.key === 'ArrowRight') {
                const next = items[(index + 1) % items.length];
                if (next) move(event, next.id);
              } else if (event.key === 'ArrowUp' || event.key === 'ArrowLeft') {
                const next = items[(index - 1 + items.length) % items.length];
                if (next) move(event, next.id);
              } else if (event.key === 'Home') {
                const next = items[0];
                if (next) move(event, next.id);
              } else if (event.key === 'End') {
                const next = items[items.length - 1];
                if (next) move(event, next.id);
              } else if (event.key === ' ' || event.key === 'Enter') {
                event.preventDefault();
                onSelectElement(item.id, event.shiftKey);
              }
            }}
          >
            <span className="work-presentation-object-list-type">
              {item.type}
            </span>
            <span className="work-presentation-object-list-label">
              {item.label}
            </span>
          </button>
        );
      })}
    </div>
  );
}

export type { PresentationObjectListItem };
