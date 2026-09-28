import { CopyCheck } from 'lucide-react';
import { useEffect, useRef, useState } from 'react';
import { officeMessage } from '../../../i18n/office-locale';
import { createWorkSlideTransition } from '../work-presentation-transition';
import type {
  WorkSlideTransition,
  WorkSlideTransitionDirection,
  WorkSlideTransitionSpeed,
  WorkSlideTransitionType,
} from '../work-types';
import {
  OfficeCheckbox,
  OfficeNumberField,
  OfficeSelect,
} from './office-controls';
import { useOfficeMessages } from './office-messages-context';
import {
  WorkOfficeRibbonButton,
  WorkOfficeRibbonGroup,
} from './work-office-chrome';

export function PresentationTransitionPanel({
  slideId,
  transition,
  editable,
  canApplyToAll,
  onChange,
  onApplyToAll,
}: {
  slideId: string;
  transition: WorkSlideTransition | undefined;
  editable: boolean;
  canApplyToAll: (transition: WorkSlideTransition | undefined) => boolean;
  onChange: (transition: WorkSlideTransition | undefined) => void;
  onApplyToAll: (transition: WorkSlideTransition | undefined) => void;
}) {
  const messages = useOfficeMessages();
  const update = (patch: Partial<WorkSlideTransition>) => {
    if (transition) onChange({ ...transition, ...patch });
  };
  const selectedAdvanceAfterDraft = presentationAdvanceAfterDraft(
    transition?.advanceAfterMs,
  );
  const selectedAdvanceAfterDraftRef = useRef({
    slideId,
    value: selectedAdvanceAfterDraft,
  });
  const [advanceAfterDraft, setAdvanceAfterDraft] = useState(
    selectedAdvanceAfterDraft,
  );
  useEffect(() => {
    const previous = selectedAdvanceAfterDraftRef.current;
    selectedAdvanceAfterDraftRef.current = {
      slideId,
      value: selectedAdvanceAfterDraft,
    };
    setAdvanceAfterDraft((draft) =>
      previous.slideId !== slideId || draft === previous.value
        ? selectedAdvanceAfterDraft
        : draft,
    );
  }, [selectedAdvanceAfterDraft, slideId]);
  const commitAdvanceAfter = (value: string): void => {
    if (!transition || transition.advanceAfterMs === undefined) {
      setAdvanceAfterDraft('');
      return;
    }
    const advanceAfterMs = normalizedPresentationAdvanceAfterMs(
      value,
      transition.advanceAfterMs,
    );
    setAdvanceAfterDraft(presentationAdvanceAfterDraft(advanceAfterMs));
    if (advanceAfterMs === transition.advanceAfterMs) return;
    update({ advanceAfterMs });
  };
  const transitionWithAdvanceDraft =
    transition?.advanceAfterMs === undefined
      ? transition
      : {
          ...transition,
          advanceAfterMs: normalizedPresentationAdvanceAfterMs(
            advanceAfterDraft,
            transition.advanceAfterMs,
          ),
        };
  const applyToAll = () => {
    if (transitionWithAdvanceDraft?.advanceAfterMs !== undefined) {
      setAdvanceAfterDraft(
        presentationAdvanceAfterDraft(
          transitionWithAdvanceDraft.advanceAfterMs,
        ),
      );
    }
    onApplyToAll(transitionWithAdvanceDraft);
  };
  return (
    <>
      <WorkOfficeRibbonGroup label={officeMessage(messages, 'presentation.transition.group')}>
        <fieldset className="work-presentation-transition-options">
          <legend className="sr-only">{officeMessage(messages, 'presentation.transition.settingsAria')}</legend>
          <div className="work-office-field effect">
            <span>{officeMessage(messages, 'presentation.transition.effect')}</span>
            <OfficeSelect
              ariaLabel={officeMessage(messages, 'presentation.transition.effectAria')}
              disabled={!editable}
              value={transition?.type ?? 'none'}
              options={[
                { value: 'none', label: officeMessage(messages, 'presentation.transition.none') },
                { value: 'fade', label: officeMessage(messages, 'presentation.transition.fade') },
                { value: 'push', label: officeMessage(messages, 'presentation.transition.push') },
                { value: 'wipe', label: officeMessage(messages, 'presentation.transition.wipe') },
                { value: 'split', label: officeMessage(messages, 'presentation.transition.split') },
                { value: 'cut', label: officeMessage(messages, 'presentation.transition.cut') },
              ]}
              onValueChange={(type) => {
                onChange(
                  type === 'none'
                    ? undefined
                    : createWorkSlideTransition(
                        type as WorkSlideTransitionType,
                        transition,
                      ),
                );
              }}
            />
          </div>
          {(transition?.type === 'push' || transition?.type === 'wipe') && (
            <div className="work-office-field direction">
              <span>{officeMessage(messages, 'presentation.transition.direction')}</span>
              <OfficeSelect
                ariaLabel={officeMessage(messages, 'presentation.transition.directionAria')}
                disabled={!editable}
                value={transition.direction ?? 'left'}
                options={[
                  { value: 'left', label: officeMessage(messages, 'presentation.transition.left') },
                  { value: 'right', label: officeMessage(messages, 'presentation.transition.right') },
                  { value: 'up', label: officeMessage(messages, 'presentation.transition.up') },
                  { value: 'down', label: officeMessage(messages, 'presentation.transition.down') },
                ]}
                onValueChange={(direction) =>
                  update({
                    direction: direction as WorkSlideTransitionDirection,
                  })
                }
              />
            </div>
          )}
          {transition?.type === 'split' && (
            <>
              <div className="work-office-field direction">
                <span>{officeMessage(messages, 'presentation.transition.direction')}</span>
                <OfficeSelect
                  ariaLabel={officeMessage(messages, 'presentation.transition.directionAria')}
                  disabled={!editable}
                  value={transition.direction ?? 'out'}
                  options={[
                    { value: 'out', label: officeMessage(messages, 'presentation.transition.out') },
                    { value: 'in', label: officeMessage(messages, 'presentation.transition.in') },
                  ]}
                  onValueChange={(direction) =>
                    update({
                      direction: direction as WorkSlideTransitionDirection,
                    })
                  }
                />
              </div>
              <div className="work-office-field orientation">
                <span>{officeMessage(messages, 'presentation.transition.splitMode')}</span>
                <OfficeSelect
                  ariaLabel={officeMessage(messages, 'presentation.transition.splitModeAria')}
                  disabled={!editable}
                  value={transition.orientation ?? 'horizontal'}
                  options={[
                    { value: 'horizontal', label: officeMessage(messages, 'presentation.transition.horizontal') },
                    { value: 'vertical', label: officeMessage(messages, 'presentation.transition.vertical') },
                  ]}
                  onValueChange={(orientation) =>
                    update({
                      orientation: orientation as 'horizontal' | 'vertical',
                    })
                  }
                />
              </div>
            </>
          )}
          <div className="work-office-field speed">
            <span>{officeMessage(messages, 'presentation.transition.speed')}</span>
            <OfficeSelect
              ariaLabel={officeMessage(messages, 'presentation.transition.speedAria')}
              disabled={!editable || !transition}
              value={transition?.speed ?? 'medium'}
              options={[
                { value: 'fast', label: officeMessage(messages, 'presentation.transition.fast') },
                { value: 'medium', label: officeMessage(messages, 'presentation.transition.medium') },
                { value: 'slow', label: officeMessage(messages, 'presentation.transition.slow') },
              ]}
              onValueChange={(speed) =>
                update({ speed: speed as WorkSlideTransitionSpeed })
              }
            />
          </div>
        </fieldset>
      </WorkOfficeRibbonGroup>
      <WorkOfficeRibbonGroup label={officeMessage(messages, 'presentation.transition.advanceGroup')}>
        <div className="work-presentation-transition-timing">
          <OfficeCheckbox
            className="toggle"
            ariaLabel={officeMessage(messages, 'presentation.transition.onClickAria')}
            disabled={!editable || !transition}
            checked={transition?.advanceOnClick ?? true}
            onCheckedChange={(advanceOnClick) => update({ advanceOnClick })}
          >
            {officeMessage(messages, 'presentation.transition.onClick')}
          </OfficeCheckbox>
          <OfficeCheckbox
            className="toggle"
            ariaLabel={officeMessage(messages, 'presentation.transition.autoAria')}
            disabled={!editable || !transition}
            checked={transition?.advanceAfterMs !== undefined}
            onCheckedChange={(checked) => {
              const advanceAfterMs = checked ? 5000 : undefined;
              const draft = presentationAdvanceAfterDraft(advanceAfterMs);
              selectedAdvanceAfterDraftRef.current = {
                slideId,
                value: draft,
              };
              setAdvanceAfterDraft(draft);
              update({ advanceAfterMs });
            }}
          >
            {officeMessage(messages, 'presentation.transition.auto')}
          </OfficeCheckbox>
          <div className="work-office-field seconds">
            <span>{officeMessage(messages, 'presentation.transition.seconds')}</span>
            <OfficeNumberField
              ariaLabel={officeMessage(messages, 'presentation.transition.secondsAria')}
              min={0.25}
              max={3600}
              step={0.25}
              disabled={
                !editable ||
                !transition ||
                transition.advanceAfterMs === undefined
              }
              value={advanceAfterDraft}
              escapeConsumer={
                advanceAfterDraft !==
                presentationAdvanceAfterDraft(transition?.advanceAfterMs)
              }
              onValueChange={setAdvanceAfterDraft}
              onCommit={commitAdvanceAfter}
              onCancel={() =>
                setAdvanceAfterDraft(
                  presentationAdvanceAfterDraft(transition?.advanceAfterMs),
                )
              }
            />
          </div>
        </div>
      </WorkOfficeRibbonGroup>
      <WorkOfficeRibbonGroup label={officeMessage(messages, 'presentation.transition.applyGroup')}>
        <WorkOfficeRibbonButton
          label={officeMessage(messages, 'presentation.transition.applyAll')}
          visibleLabel={officeMessage(messages, 'presentation.transition.applyAllVisible')}
          disabled={!editable || !canApplyToAll(transitionWithAdvanceDraft)}
          onClick={applyToAll}
        >
          <CopyCheck size={19} />
        </WorkOfficeRibbonButton>
      </WorkOfficeRibbonGroup>
    </>
  );
}

function presentationAdvanceAfterDraft(
  advanceAfterMs: number | undefined,
): string {
  return advanceAfterMs === undefined ? '' : String(advanceAfterMs / 1000);
}

function normalizedPresentationAdvanceAfterMs(
  value: string,
  current: number,
): number {
  if (!value.trim()) return current;
  const seconds = Number(value);
  if (!Number.isFinite(seconds)) return current;
  const steppedSeconds = Math.round(seconds * 4) / 4;
  return Math.round(Math.min(3600, Math.max(0.25, steppedSeconds)) * 1000);
}
