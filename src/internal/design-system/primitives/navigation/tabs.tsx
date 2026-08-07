import type { ReactNode, Ref } from 'react';
import { useTabNavigation } from './tab-navigation';

export type TabItem<T extends string> = {
  id: T;
  label: string;
  compactLabel?: string;
  icon?: ReactNode;
  badge?: ReactNode;
  disabled?: boolean;
  tabId?: string;
  panelId?: string;
};

export function Tabs<T extends string>({
  ariaLabel,
  value,
  items,
  onChange,
  onTabDoubleClick,
  variant = 'segment',
  size = 'standard',
  className = '',
  containerRef,
}: {
  ariaLabel: string;
  value: T;
  items: readonly TabItem<T>[];
  onChange: (value: T) => void;
  onTabDoubleClick?: (value: T) => void;
  variant?: 'segment' | 'line';
  size?: 'standard' | 'compact';
  className?: string;
  containerRef?: Ref<HTMLDivElement>;
}) {
  const tabNavigation = useTabNavigation({ items, onChange });

  return (
    <div
      ref={containerRef}
      className={`ds-tabs ${variant} ${size}${className ? ` ${className}` : ''}`}
      role="tablist"
      aria-label={ariaLabel}
    >
      {items.map((item) => {
        const selected = item.id === value;
        return (
          <button
            key={item.id}
            ref={(element) => {
              tabNavigation.setTabElement(item.id, element);
            }}
            type="button"
            role="tab"
            id={item.tabId}
            data-tab-id={item.id}
            aria-label={item.label}
            aria-selected={selected}
            aria-controls={item.panelId}
            tabIndex={selected ? 0 : -1}
            disabled={item.disabled}
            onClick={() => onChange(item.id)}
            onDoubleClick={() => onTabDoubleClick?.(item.id)}
            onKeyDown={(event) =>
              tabNavigation.handleTabKeyDown(event, item.id)
            }
          >
            {item.icon && (
              <span className="ds-tabs-icon" aria-hidden="true">
                {item.icon}
              </span>
            )}
            <span
              className={`ds-tabs-label${item.compactLabel ? ' has-compact' : ''}`}
            >
              {item.label}
            </span>
            {item.compactLabel && (
              <span className="ds-tabs-label-compact" aria-hidden="true">
                {item.compactLabel}
              </span>
            )}
            {item.badge !== undefined && (
              <>
                {' '}
                <span className="ds-tabs-badge">{item.badge}</span>
              </>
            )}
          </button>
        );
      })}
    </div>
  );
}
