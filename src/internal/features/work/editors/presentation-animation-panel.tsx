import { ArrowDown, ArrowUp, Play } from 'lucide-react';
import { useEffect, useState } from 'react';
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
          { value: 'none', label: '无' },
          { value: 'appear', label: '出现' },
          { value: 'fade', label: '淡入' },
          { value: 'fly-in', label: '飞入' },
          { value: 'zoom', label: '缩放' },
        ]
      : [
          { value: 'none', label: '无' },
          { value: 'disappear', label: '消失' },
          { value: 'fade-out', label: '淡出' },
          { value: 'fly-out', label: '飞出' },
          { value: 'zoom-out', label: '缩小' },
        ];
  return (
    <>
      <WorkOfficeRibbonGroup label="对象动画">
        <fieldset
          className="work-presentation-animation-options"
          data-animation-class={animationClass}
        >
          <legend className="sr-only">对象动画设置</legend>
          <div className="work-office-field animation-class">
            <span>类型</span>
            <OfficeSelect
              ariaLabel="对象动画类型"
              disabled={!editable}
              value={animationClass}
              options={[
                { value: 'entrance', label: '进入', meta: '让对象出现' },
                { value: 'exit', label: '退出', meta: '让对象消失' },
              ]}
              onValueChange={(nextClass) =>
                setSelection({ animationClass: nextClass, elementId })
              }
            />
          </div>
          <div className="work-office-field effect">
            <span>效果</span>
            <OfficeSelect
              ariaLabel="对象动画效果"
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
            <span>开始</span>
            <OfficeSelect
              ariaLabel="对象动画触发方式"
              disabled={!animation}
              value={animation?.trigger ?? 'on-click'}
              options={[
                {
                  value: 'on-click',
                  label: '单击时',
                  disabled:
                    Boolean(animation) &&
                    !canUpdate(animationClass, { trigger: 'on-click' }),
                },
                {
                  value: 'with-previous',
                  label: '与上一动画同时',
                  disabled:
                    Boolean(animation) &&
                    !canUpdate(animationClass, {
                      trigger: 'with-previous',
                    }),
                },
                {
                  value: 'after-previous',
                  label: '上一动画之后',
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
              <span>方向</span>
              <OfficeSelect
                ariaLabel="对象动画方向"
                value={animation.direction ?? 'left'}
                options={[
                  {
                    value: 'left',
                    label: animationClass === 'entrance' ? '从左侧' : '向左侧',
                  },
                  {
                    value: 'right',
                    label: animationClass === 'entrance' ? '从右侧' : '向右侧',
                  },
                  {
                    value: 'up',
                    label: animationClass === 'entrance' ? '从上方' : '向上方',
                  },
                  {
                    value: 'down',
                    label: animationClass === 'entrance' ? '从下方' : '向下方',
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
      <WorkOfficeRibbonGroup label="动画计时">
        <div className="work-presentation-animation-timing">
          <AnimationTimingField
            animationId={animation?.id}
            ariaLabel="对象动画持续秒数"
            disabled={!animation}
            label="持续"
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
            ariaLabel="对象动画延迟秒数"
            disabled={!animation}
            label="延迟"
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
      <WorkOfficeRibbonGroup label="动画顺序">
        <WorkOfficeRibbonButton
          label="提前对象动画"
          visibleLabel="提前"
          disabled={!canMove(animationClass, -1)}
          onClick={() => onMove(animationClass, -1)}
        >
          <ArrowUp size={19} />
        </WorkOfficeRibbonButton>
        <WorkOfficeRibbonButton
          label="推后对象动画"
          visibleLabel="推后"
          disabled={!canMove(animationClass, 1)}
          onClick={() => onMove(animationClass, 1)}
        >
          <ArrowDown size={19} />
        </WorkOfficeRibbonButton>
      </WorkOfficeRibbonGroup>
      <WorkOfficeRibbonGroup label="动画预览">
        <WorkOfficeRibbonButton
          label="预览当前幻灯片动画"
          visibleLabel="预览"
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
