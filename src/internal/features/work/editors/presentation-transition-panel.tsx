import { CopyCheck } from 'lucide-react';
import { useEffect, useRef, useState } from 'react';
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
      <WorkOfficeRibbonGroup label="切换效果">
        <fieldset className="work-presentation-transition-options">
          <legend className="sr-only">幻灯片切换设置</legend>
          <div className="work-office-field effect">
            <span>效果</span>
            <OfficeSelect
              ariaLabel="幻灯片切换效果"
              disabled={!editable}
              value={transition?.type ?? 'none'}
              options={[
                { value: 'none', label: '无' },
                { value: 'fade', label: '淡化' },
                { value: 'push', label: '推进' },
                { value: 'wipe', label: '擦除' },
                { value: 'split', label: '分割' },
                { value: 'cut', label: '切换' },
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
              <span>方向</span>
              <OfficeSelect
                ariaLabel="切换方向"
                disabled={!editable}
                value={transition.direction ?? 'left'}
                options={[
                  { value: 'left', label: '向左' },
                  { value: 'right', label: '向右' },
                  { value: 'up', label: '向上' },
                  { value: 'down', label: '向下' },
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
                <span>方向</span>
                <OfficeSelect
                  ariaLabel="切换方向"
                  disabled={!editable}
                  value={transition.direction ?? 'out'}
                  options={[
                    { value: 'out', label: '向外' },
                    { value: 'in', label: '向内' },
                  ]}
                  onValueChange={(direction) =>
                    update({
                      direction: direction as WorkSlideTransitionDirection,
                    })
                  }
                />
              </div>
              <div className="work-office-field orientation">
                <span>分割方式</span>
                <OfficeSelect
                  ariaLabel="分割方式"
                  disabled={!editable}
                  value={transition.orientation ?? 'horizontal'}
                  options={[
                    { value: 'horizontal', label: '水平' },
                    { value: 'vertical', label: '垂直' },
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
            <span>速度</span>
            <OfficeSelect
              ariaLabel="切换速度"
              disabled={!editable || !transition}
              value={transition?.speed ?? 'medium'}
              options={[
                { value: 'fast', label: '快速' },
                { value: 'medium', label: '中速' },
                { value: 'slow', label: '慢速' },
              ]}
              onValueChange={(speed) =>
                update({ speed: speed as WorkSlideTransitionSpeed })
              }
            />
          </div>
        </fieldset>
      </WorkOfficeRibbonGroup>
      <WorkOfficeRibbonGroup label="换片方式">
        <div className="work-presentation-transition-timing">
          <OfficeCheckbox
            className="toggle"
            ariaLabel="单击鼠标后换片"
            disabled={!editable || !transition}
            checked={transition?.advanceOnClick ?? true}
            onCheckedChange={(advanceOnClick) => update({ advanceOnClick })}
          >
            单击鼠标后
          </OfficeCheckbox>
          <OfficeCheckbox
            className="toggle"
            ariaLabel="自动换片"
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
            自动换片
          </OfficeCheckbox>
          <div className="work-office-field seconds">
            <span>秒数</span>
            <OfficeNumberField
              ariaLabel="自动换片秒数"
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
      <WorkOfficeRibbonGroup label="应用">
        <WorkOfficeRibbonButton
          label="应用切换效果到全部幻灯片"
          visibleLabel="应用到全部"
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
