import { ArrowDown, ArrowUp, Play } from 'lucide-react';
import { useEffect, useState } from 'react';
import { officeMessage } from '../../../i18n/office-locale';
import {
  WORK_SLIDE_ANIMATION_MAX_DELAY_MS,
  WORK_SLIDE_ANIMATION_MAX_DURATION_MS,
  WORK_SLIDE_ANIMATION_MIN_DURATION_MS,
} from '../work-presentation-animation-constraints';
import type {
  WorkSlideAnimation,
  WorkSlideAnimationClass,
  WorkSlideAnimationDirection,
  WorkSlideAnimationEffect,
  WorkSlideAnimationTrigger,
} from '../work-types';
import { OfficeNumberField, OfficeSelect } from './office-controls';
import { useOfficeMessages } from './office-messages-context';
import {
  WorkOfficeRibbonButton,
  WorkOfficeRibbonGroup,
} from './work-office-chrome';

export function PresentationAnimationPanel({
  animations,
  canMove,
  canPreview,
  canUpdate,
  editable,
  elementId,
  onMove,
  onPreview,
  onSetEffect,
  onUpdate,
}: {
  animations: Record<WorkSlideAnimationClass, WorkSlideAnimation | undefined>;
  canMove: (
    animationClass: WorkSlideAnimationClass,
    direction: -1 | 1,
  ) => boolean;
  canPreview: boolean;
  canUpdate: (
    animationClass: WorkSlideAnimationClass,
    patch: Partial<WorkSlideAnimation>,
  ) => boolean;
  editable: boolean;
  elementId: string | undefined;
  onMove: (animationClass: WorkSlideAnimationClass, direction: -1 | 1) => void;
  onPreview: () => void;
  onSetEffect: (
    animationClass: WorkSlideAnimationClass,
    effect: WorkSlideAnimationEffect | undefined,
  ) => void;
  onUpdate: (
    animationClass: WorkSlideAnimationClass,
    patch: Partial<WorkSlideAnimation>,
  ) => void;
}) {
  const messages = useOfficeMessages();
  const defaultClass = animations.entrance
    ? 'entrance'
    : animations.exit
      ? 'exit'
      : 'entrance';
  const [selection, setSelection] = useState<{
    animationClass: WorkSlideAnimationClass;
    elementId: string | undefined;
  }>({ animationClass: defaultClass, elementId });
  const animationClass =
    selection.elementId === elementId ? selection.animationClass : defaultClass;
  const animation = animations[animationClass];
  const flyEffect =
    animation?.effect === 'fly-in' || animation?.effect === 'fly-out';
  const effectOptions =
    animationClass === 'entrance'
      ? [
          { value: 'none', label: officeMessage(messages, 'presentation.animation.none') },
          { value: 'appear', label: officeMessage(messages, 'presentation.animation.appear') },
          { value: 'fade', label: officeMessage(messages, 'presentation.animation.fade') },
          { value: 'fly-in', label: officeMessage(messages, 'presentation.animation.flyIn') },
          { value: 'zoom', label: officeMessage(messages, 'presentation.animation.zoom') },
        ]
      : [
          { value: 'none', label: officeMessage(messages, 'presentation.animation.none') },
          { value: 'disappear', label: officeMessage(messages, 'presentation.animation.disappear') },
          { value: 'fade-out', label: officeMessage(messages, 'presentation.animation.fadeOut') },
          { value: 'fly-out', label: officeMessage(messages, 'presentation.animation.flyOut') },
          { value: 'zoom-out', label: officeMessage(messages, 'presentation.animation.zoomOut') },
        ];
  return (
    <>
      <WorkOfficeRibbonGroup label={officeMessage(messages, 'presentation.animation.group')}>
        <fieldset
          className="work-presentation-animation-options"
          data-animation-class={animationClass}
        >
          <legend className="sr-only">{officeMessage(messages, 'presentation.animation.settingsAria')}</legend>
          <div className="work-office-field animation-class">
            <span>{officeMessage(messages, 'presentation.animation.type')}</span>
            <OfficeSelect
              ariaLabel={officeMessage(messages, 'presentation.animation.typeAria')}
              disabled={!editable}
              value={animationClass}
              options={[
                { value: 'entrance', label: officeMessage(messages, 'presentation.animation.entrance'), meta: officeMessage(messages, 'presentation.animation.entranceMeta') },
                { value: 'exit', label: officeMessage(messages, 'presentation.animation.exit'), meta: officeMessage(messages, 'presentation.animation.exitMeta') },
              ]}
              onValueChange={(nextClass) =>
                setSelection({ animationClass: nextClass, elementId })
              }
            />
          </div>
          <div className="work-office-field effect">
            <span>{officeMessage(messages, 'presentation.animation.effect')}</span>
            <OfficeSelect
              ariaLabel={officeMessage(messages, 'presentation.animation.effectAria')}
              disabled={!editable}
              value={animation?.effect ?? 'none'}
              options={effectOptions}
              onValueChange={(effect) =>
                onSetEffect(
                  animationClass,
                  effect === 'none'
                    ? undefined
                    : (effect as WorkSlideAnimationEffect),
                )
              }
            />
          </div>
          <div className="work-office-field trigger">
            <span>{officeMessage(messages, 'presentation.animation.start')}</span>
            <OfficeSelect
              ariaLabel={officeMessage(messages, 'presentation.animation.startAria')}
              disabled={!animation}
              value={animation?.trigger ?? 'on-click'}
              options={[
                {
                  value: 'on-click',
                  label: officeMessage(messages, 'presentation.animation.onClick'),
                  disabled:
                    Boolean(animation) &&
                    !canUpdate(animationClass, { trigger: 'on-click' }),
                },
                {
                  value: 'with-previous',
                  label: officeMessage(messages, 'presentation.animation.withPrevious'),
                  disabled:
                    Boolean(animation) &&
                    !canUpdate(animationClass, {
                      trigger: 'with-previous',
                    }),
                },
                {
                  value: 'after-previous',
                  label: officeMessage(messages, 'presentation.animation.afterPrevious'),
                  disabled:
                    Boolean(animation) &&
                    !canUpdate(animationClass, {
                      trigger: 'after-previous',
                    }),
                },
              ]}
              onValueChange={(trigger) =>
                onUpdate(animationClass, {
                  trigger: trigger as WorkSlideAnimationTrigger,
                })
              }
            />
          </div>
          {flyEffect && (
            <div className="work-office-field direction">
              <span>{officeMessage(messages, 'presentation.animation.direction')}</span>
              <OfficeSelect
                ariaLabel={officeMessage(messages, 'presentation.animation.directionAria')}
                value={animation.direction ?? 'left'}
                options={[
                  {
                    value: 'left',
                    label: animationClass === 'entrance' ? officeMessage(messages, 'presentation.animation.fromLeft') : officeMessage(messages, 'presentation.animation.toLeft'),
                  },
                  {
                    value: 'right',
                    label: animationClass === 'entrance' ? officeMessage(messages, 'presentation.animation.fromRight') : officeMessage(messages, 'presentation.animation.toRight'),
                  },
                  {
                    value: 'up',
                    label: animationClass === 'entrance' ? officeMessage(messages, 'presentation.animation.fromTop') : officeMessage(messages, 'presentation.animation.toTop'),
                  },
                  {
                    value: 'down',
                    label: animationClass === 'entrance' ? officeMessage(messages, 'presentation.animation.fromBottom') : officeMessage(messages, 'presentation.animation.toBottom'),
                  },
                ]}
                onValueChange={(direction) =>
                  onUpdate(animationClass, {
                    direction: direction as WorkSlideAnimationDirection,
                  })
                }
              />
            </div>
          )}
        </fieldset>
      </WorkOfficeRibbonGroup>
      <WorkOfficeRibbonGroup label={officeMessage(messages, 'presentation.animation.timingGroup')}>
        <div className="work-presentation-animation-timing">
          <AnimationTimingField
            animationId={animation?.id}
            ariaLabel={officeMessage(messages, 'presentation.animation.durationAria')}
            disabled={!animation}
            label={officeMessage(messages, 'presentation.animation.duration')}
            maximumMs={WORK_SLIDE_ANIMATION_MAX_DURATION_MS}
            minimumMs={WORK_SLIDE_ANIMATION_MIN_DURATION_MS}
            valueMs={animation?.durationMs ?? 500}
            onCommit={(durationMs) => {
              if (!canUpdate(animationClass, { durationMs })) return false;
              onUpdate(animationClass, { durationMs });
              return true;
            }}
          />
          <AnimationTimingField
            animationId={animation?.id}
            ariaLabel={officeMessage(messages, 'presentation.animation.delayAria')}
            disabled={!animation}
            label={officeMessage(messages, 'presentation.animation.delay')}
            maximumMs={WORK_SLIDE_ANIMATION_MAX_DELAY_MS}
            minimumMs={0}
            valueMs={animation?.delayMs ?? 0}
            onCommit={(delayMs) => {
              if (!canUpdate(animationClass, { delayMs })) return false;
              onUpdate(animationClass, { delayMs });
              return true;
            }}
          />
        </div>
      </WorkOfficeRibbonGroup>
      <WorkOfficeRibbonGroup label={officeMessage(messages, 'presentation.animation.orderGroup')}>
        <WorkOfficeRibbonButton
          label={officeMessage(messages, 'presentation.animation.moveEarlier')}
          visibleLabel={officeMessage(messages, 'presentation.animation.moveEarlierVisible')}
          disabled={!canMove(animationClass, -1)}
          onClick={() => onMove(animationClass, -1)}
        >
          <ArrowUp size={19} />
        </WorkOfficeRibbonButton>
        <WorkOfficeRibbonButton
          label={officeMessage(messages, 'presentation.animation.moveLater')}
          visibleLabel={officeMessage(messages, 'presentation.animation.moveLaterVisible')}
          disabled={!canMove(animationClass, 1)}
          onClick={() => onMove(animationClass, 1)}
        >
          <ArrowDown size={19} />
        </WorkOfficeRibbonButton>
      </WorkOfficeRibbonGroup>
      <WorkOfficeRibbonGroup label={officeMessage(messages, 'presentation.animation.previewGroup')}>
        <WorkOfficeRibbonButton
          label={officeMessage(messages, 'presentation.animation.preview')}
          visibleLabel={officeMessage(messages, 'presentation.animation.previewVisible')}
          disabled={!canPreview}
          onClick={onPreview}
        >
          <Play size={19} />
        </WorkOfficeRibbonButton>
      </WorkOfficeRibbonGroup>
    </>
  );
}

function AnimationTimingField({
  animationId,
  ariaLabel,
  disabled,
  label,
  maximumMs,
  minimumMs,
  onCommit,
  valueMs,
}: {
  animationId: string | undefined;
  ariaLabel: string;
  disabled: boolean;
  label: string;
  maximumMs: number;
  minimumMs: number;
  onCommit: (valueMs: number) => boolean | void;
  valueMs: number;
}) {
  const canonical = animationSecondsDraft(valueMs);
  const [draft, setDraft] = useState(canonical);
  useEffect(() => setDraft(canonical), [animationId, canonical]);
  const commit = (value: string) => {
    const next = normalizedAnimationMilliseconds(
      value,
      valueMs,
      minimumMs,
      maximumMs,
    );
    if (next === valueMs) {
      setDraft(animationSecondsDraft(next));
      return;
    }
    const accepted = onCommit(next);
    setDraft(animationSecondsDraft(accepted === false ? valueMs : next));
  };
  return (
    <div className="work-office-field seconds">
      <span>{label}</span>
      <OfficeNumberField
        ariaLabel={ariaLabel}
        disabled={disabled}
        min={minimumMs / 1000}
        max={maximumMs / 1000}
        step={0.1}
        value={draft}
        escapeConsumer={draft !== canonical}
        onValueChange={setDraft}
        onCommit={commit}
        onCancel={() => setDraft(canonical)}
      />
    </div>
  );
}

function animationSecondsDraft(valueMs: number): string {
  return String(valueMs / 1000);
}

function normalizedAnimationMilliseconds(
  value: string,
  current: number,
  minimum: number,
  maximum: number,
): number {
  if (!value.trim()) return current;
  const seconds = Number(value);
  if (!Number.isFinite(seconds)) return current;
  const stepped = Math.round(seconds * 10) / 10;
  return Math.round(
    Math.min(maximum / 1000, Math.max(minimum / 1000, stepped)) * 1000,
  );
}
