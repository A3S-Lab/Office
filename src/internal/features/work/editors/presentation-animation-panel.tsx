import { ArrowDown, ArrowUp, Play } from 'lucide-react';
import { useEffect, useState } from 'react';
import {
  WORK_SLIDE_ANIMATION_MAX_DELAY_MS,
  WORK_SLIDE_ANIMATION_MAX_DURATION_MS,
  WORK_SLIDE_ANIMATION_MIN_DURATION_MS,
} from '../work-presentation-animation-constraints';
import type {
  WorkSlideAnimation,
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
  animation,
  canMove,
  canPreview,
  editable,
  onMove,
  onPreview,
  onSetEffect,
  onUpdate,
}: {
  animation: WorkSlideAnimation | undefined;
  canMove: (direction: -1 | 1) => boolean;
  canPreview: boolean;
  editable: boolean;
  onMove: (direction: -1 | 1) => void;
  onPreview: () => void;
  onSetEffect: (effect: WorkSlideAnimationEffect | undefined) => void;
  onUpdate: (patch: Partial<WorkSlideAnimation>) => void;
}) {
  return (
    <>
      <WorkOfficeRibbonGroup label="入场动画">
        <fieldset className="work-presentation-animation-options">
          <legend className="sr-only">对象入场动画设置</legend>
          <div className="work-office-field effect">
            <span>效果</span>
            <OfficeSelect
              ariaLabel="对象入场动画效果"
              disabled={!editable}
              value={animation?.effect ?? 'none'}
              options={[
                { value: 'none', label: '无' },
                { value: 'appear', label: '出现' },
                { value: 'fade', label: '淡入' },
                { value: 'fly-in', label: '飞入' },
                { value: 'zoom', label: '缩放' },
              ]}
              onValueChange={(effect) =>
                onSetEffect(
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
              ariaLabel="对象入场动画触发方式"
              disabled={!animation}
              value={animation?.trigger ?? 'on-click'}
              options={[
                { value: 'on-click', label: '单击时' },
                { value: 'with-previous', label: '与上一动画同时' },
                { value: 'after-previous', label: '上一动画之后' },
              ]}
              onValueChange={(trigger) =>
                onUpdate({ trigger: trigger as WorkSlideAnimationTrigger })
              }
            />
          </div>
          {animation?.effect === 'fly-in' && (
            <div className="work-office-field direction">
              <span>方向</span>
              <OfficeSelect
                ariaLabel="对象飞入方向"
                value={animation.direction ?? 'left'}
                options={[
                  { value: 'left', label: '从左侧' },
                  { value: 'right', label: '从右侧' },
                  { value: 'up', label: '从上方' },
                  { value: 'down', label: '从下方' },
                ]}
                onValueChange={(direction) =>
                  onUpdate({
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
            ariaLabel="对象入场动画持续秒数"
            disabled={!animation}
            label="持续"
            maximumMs={WORK_SLIDE_ANIMATION_MAX_DURATION_MS}
            minimumMs={WORK_SLIDE_ANIMATION_MIN_DURATION_MS}
            valueMs={animation?.durationMs ?? 500}
            onCommit={(durationMs) => onUpdate({ durationMs })}
          />
          <AnimationTimingField
            animationId={animation?.id}
            ariaLabel="对象入场动画延迟秒数"
            disabled={!animation}
            label="延迟"
            maximumMs={WORK_SLIDE_ANIMATION_MAX_DELAY_MS}
            minimumMs={0}
            valueMs={animation?.delayMs ?? 0}
            onCommit={(delayMs) => onUpdate({ delayMs })}
          />
        </div>
      </WorkOfficeRibbonGroup>
      <WorkOfficeRibbonGroup label="动画顺序">
        <WorkOfficeRibbonButton
          label="提前对象动画"
          visibleLabel="提前"
          disabled={!canMove(-1)}
          onClick={() => onMove(-1)}
        >
          <ArrowUp size={19} />
        </WorkOfficeRibbonButton>
        <WorkOfficeRibbonButton
          label="推后对象动画"
          visibleLabel="推后"
          disabled={!canMove(1)}
          onClick={() => onMove(1)}
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
  onCommit: (valueMs: number) => void;
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
    setDraft(animationSecondsDraft(next));
    if (next !== valueMs) onCommit(next);
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
