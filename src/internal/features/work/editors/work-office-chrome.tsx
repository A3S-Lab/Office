import {
  ChevronDown,
  ChevronLeft,
  ChevronRight,
  Eye,
  Minus,
  Plus,
} from 'lucide-react';
import {
  type ButtonHTMLAttributes,
  Fragment,
  type ReactNode,
  useCallback,
  useId,
  useLayoutEffect,
  useRef,
  useState,
} from 'react';
import { Popover, Tabs } from '../../../design-system/primitives';
import { OfficeSlider } from './office-controls';
import { moveOfficeMenuFocus } from './office-menu-keyboard';
import {
  calculateRibbonOverflow,
  calculateRibbonScrollTarget,
  type WorkOfficeRibbonItemGeometry,
} from './work-office-ribbon-overflow';

export interface WorkOfficeRibbonTab<T extends string> {
  id: T;
  label: string;
  compactLabel?: string;
}

export interface WorkOfficeFileAction {
  id: string;
  label: string;
  icon?: ReactNode;
  shortcut?: string;
  disabled?: boolean;
  separatorBefore?: boolean;
  onSelect: () => void | Promise<void>;
}

export function WorkOfficeRibbon<T extends string>({
  ariaLabel,
  tabs,
  defaultTab,
  activeTab,
  onTabChange,
  panels,
  fileActions,
  className = '',
  toolbarClassName = '',
}: {
  ariaLabel: string;
  tabs: readonly WorkOfficeRibbonTab<T>[];
  defaultTab: T;
  activeTab?: T;
  onTabChange?: (tab: T) => void;
  panels: Record<T, ReactNode>;
  fileActions?: readonly WorkOfficeFileAction[];
  className?: string;
  toolbarClassName?: string;
}) {
  const reactId = useId().replaceAll(':', '');
  const [internalTab, setInternalTab] = useState(defaultTab);
  const [ribbonOverflow, setRibbonOverflow] = useState({
    backward: false,
    forward: false,
  });
  const [tabOverflow, setTabOverflow] = useState({
    backward: false,
    forward: false,
  });
  const tabListRef = useRef<HTMLDivElement>(null);
  const toolbarRef = useRef<HTMLDivElement>(null);
  const selectedTab = activeTab ?? internalTab;
  const selectedLabel =
    tabs.find((tab) => tab.id === selectedTab)?.label ?? tabs[0]?.label ?? '';
  const selectedPanel = panels[selectedTab];
  const hasSelectedPanel =
    selectedPanel !== null && selectedPanel !== undefined;
  const hasRibbonOverflow = ribbonOverflow.backward || ribbonOverflow.forward;
  const updateRibbonOverflow = useCallback(() => {
    const toolbar = toolbarRef.current;
    if (!toolbar) return;
    const { backward, forward } = calculateRibbonOverflow({
      clientWidth: toolbar.clientWidth,
      items: ribbonItemGeometry(toolbar),
      scrollLeft: toolbar.scrollLeft,
    });
    if (!backward && !forward && toolbar.scrollLeft > 0) {
      toolbar.scrollLeft = 0;
    }
    setRibbonOverflow((current) =>
      current.backward === backward && current.forward === forward
        ? current
        : { backward, forward },
    );
  }, []);
  const scrollRibbon = (direction: -1 | 1) => {
    const toolbar = toolbarRef.current;
    if (!toolbar) return;
    toolbar.scrollLeft = calculateRibbonScrollTarget({
      clientWidth: toolbar.clientWidth,
      direction,
      items: ribbonItemGeometry(toolbar),
      scrollLeft: toolbar.scrollLeft,
    });
    updateRibbonOverflow();
  };
  const updateTabOverflow = useCallback(() => {
    const tabList = tabListRef.current;
    if (!tabList) return;
    const { backward, forward } = calculateRibbonOverflow({
      clientWidth: tabList.clientWidth,
      items: ribbonItemGeometry(tabList),
      scrollLeft: tabList.scrollLeft,
    });
    if (!backward && !forward && tabList.scrollLeft > 0) {
      tabList.scrollLeft = 0;
    }
    setTabOverflow((current) =>
      current.backward === backward && current.forward === forward
        ? current
        : { backward, forward },
    );
  }, []);
  const scrollTabs = (direction: -1 | 1) => {
    const tabList = tabListRef.current;
    if (!tabList) return;
    tabList.scrollLeft = calculateRibbonScrollTarget({
      clientWidth: tabList.clientWidth,
      direction,
      items: ribbonItemGeometry(tabList),
      scrollLeft: tabList.scrollLeft,
    });
    updateTabOverflow();
  };
  const selectTab = (tab: T) => {
    if (activeTab === undefined) setInternalTab(tab);
    onTabChange?.(tab);
  };

  useLayoutEffect(() => {
    const toolbar = toolbarRef.current;
    if (!toolbar) return;
    toolbar.scrollLeft = 0;
    updateRibbonOverflow();
  }, [selectedTab, hasSelectedPanel, updateRibbonOverflow]);

  useLayoutEffect(() => {
    const toolbar = toolbarRef.current;
    if (!toolbar) return;
    let frame = requestAnimationFrame(updateRibbonOverflow);
    const scheduleOverflowUpdate = () => {
      cancelAnimationFrame(frame);
      frame = requestAnimationFrame(updateRibbonOverflow);
    };
    const resizeObserver =
      typeof ResizeObserver === 'undefined'
        ? null
        : new ResizeObserver(scheduleOverflowUpdate);
    const observeChild = (node: Node) => {
      if (node instanceof Element) resizeObserver?.observe(node);
    };
    resizeObserver?.observe(toolbar);
    for (const child of toolbar.children) observeChild(child);
    const mutationObserver =
      typeof MutationObserver === 'undefined'
        ? null
        : new MutationObserver((records) => {
            for (const record of records) {
              for (const node of record.addedNodes) observeChild(node);
            }
            scheduleOverflowUpdate();
          });
    mutationObserver?.observe(toolbar, {
      attributes: true,
      characterData: true,
      childList: true,
      subtree: true,
    });
    window.addEventListener('resize', scheduleOverflowUpdate);
    return () => {
      cancelAnimationFrame(frame);
      resizeObserver?.disconnect();
      mutationObserver?.disconnect();
      window.removeEventListener('resize', scheduleOverflowUpdate);
    };
  }, [hasSelectedPanel, updateRibbonOverflow]);

  useLayoutEffect(() => {
    const tabList = tabListRef.current;
    if (!tabList) return;
    let frame = requestAnimationFrame(() => {
      scrollSelectedRibbonTabIntoView(tabList, selectedTab);
      updateTabOverflow();
    });
    const scheduleOverflowUpdate = () => {
      cancelAnimationFrame(frame);
      frame = requestAnimationFrame(updateTabOverflow);
    };
    const resizeObserver =
      typeof ResizeObserver === 'undefined'
        ? null
        : new ResizeObserver(scheduleOverflowUpdate);
    const observeChild = (node: Node) => {
      if (node instanceof Element) resizeObserver?.observe(node);
    };
    resizeObserver?.observe(tabList);
    for (const child of tabList.children) observeChild(child);
    const mutationObserver =
      typeof MutationObserver === 'undefined'
        ? null
        : new MutationObserver((records) => {
            for (const record of records) {
              for (const node of record.addedNodes) observeChild(node);
            }
            scheduleOverflowUpdate();
          });
    mutationObserver?.observe(tabList, {
      attributes: true,
      characterData: true,
      childList: true,
      subtree: true,
    });
    tabList.addEventListener('scroll', updateTabOverflow);
    window.addEventListener('resize', scheduleOverflowUpdate);
    return () => {
      cancelAnimationFrame(frame);
      resizeObserver?.disconnect();
      mutationObserver?.disconnect();
      tabList.removeEventListener('scroll', updateTabOverflow);
      window.removeEventListener('resize', scheduleOverflowUpdate);
    };
  }, [selectedTab, tabs, updateTabOverflow]);

  return (
    <section
      className={`work-office-ribbon ${className}`.trim()}
      aria-label={ariaLabel}
    >
      <div className="work-office-ribbon-tabs-row">
        {fileActions?.length ? (
          <WorkOfficeFileMenu actions={fileActions} />
        ) : null}
        <div className="work-office-ribbon-tab-strip">
          <Tabs
            ariaLabel={ariaLabel}
            value={selectedTab}
            variant="line"
            size="compact"
            className="work-office-ribbon-tabs"
            containerRef={tabListRef}
            items={tabs.map((tab) => ({
              ...tab,
              tabId: `${reactId}-tab-${tab.id}`,
              panelId: `${reactId}-panel`,
            }))}
            onChange={selectTab}
          />
          {(tabOverflow.backward || tabOverflow.forward) && (
            <nav
              className="work-office-ribbon-tab-navigation"
              aria-label="功能区标签翻页"
            >
              <button
                type="button"
                aria-label="向左查看更多功能区标签"
                disabled={!tabOverflow.backward}
                onClick={() => scrollTabs(-1)}
              >
                <ChevronLeft size={14} />
              </button>
              <button
                type="button"
                aria-label="向右查看更多功能区标签"
                disabled={!tabOverflow.forward}
                onClick={() => scrollTabs(1)}
              >
                <ChevronRight size={14} />
              </button>
            </nav>
          )}
        </div>
      </div>
      <div
        id={`${reactId}-panel`}
        className="work-office-ribbon-panel"
        data-empty={hasSelectedPanel ? undefined : 'true'}
        role="tabpanel"
        aria-labelledby={`${reactId}-tab-${selectedTab}`}
      >
        {hasSelectedPanel && (
          <>
            {hasRibbonOverflow && (
              <button
                type="button"
                className="work-office-ribbon-scroll previous"
                aria-label={`向左查看更多${selectedLabel}工具`}
                disabled={!ribbonOverflow.backward}
                onClick={() => scrollRibbon(-1)}
              >
                <ChevronLeft size={15} />
              </button>
            )}
            <div
              ref={toolbarRef}
              className={`work-office-toolbar ${toolbarClassName}`.trim()}
              data-has-overflow={hasRibbonOverflow ? 'true' : undefined}
              role="toolbar"
              aria-label={`${selectedLabel}工具栏`}
              onScroll={updateRibbonOverflow}
            >
              {selectedPanel}
            </div>
            {hasRibbonOverflow && (
              <button
                type="button"
                className="work-office-ribbon-scroll next"
                aria-label={`向右查看更多${selectedLabel}工具`}
                disabled={!ribbonOverflow.forward}
                onClick={() => scrollRibbon(1)}
              >
                <ChevronRight size={15} />
              </button>
            )}
          </>
        )}
      </div>
    </section>
  );
}

function ribbonItemGeometry(
  toolbar: HTMLDivElement,
): WorkOfficeRibbonItemGeometry[] {
  return Array.from(toolbar.children).flatMap((element) =>
    element instanceof HTMLElement ? [ribbonItemBounds(toolbar, element)] : [],
  );
}

function ribbonItemBounds(
  toolbar: HTMLDivElement,
  element: HTMLElement,
): WorkOfficeRibbonItemGeometry {
  const left =
    element.offsetParent === toolbar
      ? element.offsetLeft
      : element.offsetLeft - toolbar.offsetLeft;
  return { left, right: left + element.offsetWidth };
}

function scrollSelectedRibbonTabIntoView(
  tabList: HTMLDivElement,
  selectedTab: string,
): void {
  const selected = Array.from(tabList.children).find(
    (element): element is HTMLElement =>
      element instanceof HTMLElement && element.dataset.tabId === selectedTab,
  );
  if (!selected || tabList.clientWidth <= 0) return;
  const inset = 8;
  const visibleLeft = tabList.scrollLeft + inset;
  const visibleRight = tabList.scrollLeft + tabList.clientWidth - inset;
  const selectedLeft = selected.offsetLeft;
  const selectedRight = selectedLeft + selected.offsetWidth;
  if (selectedLeft < visibleLeft) {
    tabList.scrollLeft = Math.max(0, selectedLeft - inset);
  } else if (selectedRight > visibleRight) {
    tabList.scrollLeft = Math.max(
      0,
      selectedRight - tabList.clientWidth + inset,
    );
  }
}

export function WorkOfficePreviewBar({
  ariaLabel,
  label,
  detail,
  fileActions,
  className = '',
}: {
  ariaLabel: string;
  label: string;
  detail?: string;
  fileActions?: readonly WorkOfficeFileAction[];
  className?: string;
}) {
  return (
    <section
      className={`work-office-preview-bar ${className}`.trim()}
      aria-label={ariaLabel}
    >
      {fileActions?.length ? (
        <WorkOfficeFileMenu actions={fileActions} />
      ) : null}
      <div className="work-office-preview-summary">
        <Eye size={14} aria-hidden="true" />
        <strong>{label}</strong>
        {detail && <span>{detail}</span>}
      </div>
    </section>
  );
}

function WorkOfficeFileMenu({
  actions,
}: {
  actions: readonly WorkOfficeFileAction[];
}) {
  const [open, setOpen] = useState(false);
  const menuRef = useRef<HTMLElement>(null);
  const focusEdgeRef = useRef<'first' | 'last'>('first');

  const focusRequestedAction = () =>
    requestAnimationFrame(() => {
      const buttons = [
        ...(menuRef.current?.querySelectorAll<HTMLButtonElement>(
          'button:not(:disabled)',
        ) ?? []),
      ];
      const button =
        focusEdgeRef.current === 'last' ? buttons.at(-1) : buttons[0];
      button?.focus();
    });

  return (
    <Popover
      label="文件"
      panelLabel="文件菜单"
      panelRole="menu"
      portal
      className="work-office-file-menu"
      panelClassName="work-office-file-popover"
      open={open}
      panelRef={menuRef}
      onPanelKeyDown={moveOfficeMenuFocus}
      onOpenChange={(nextOpen) => {
        setOpen(nextOpen);
        if (nextOpen) focusRequestedAction();
      }}
      trigger={(triggerProps, { open }) => (
        <button
          {...triggerProps}
          className="work-office-file-trigger"
          onKeyDown={(event) => {
            if (event.key !== 'ArrowDown' && event.key !== 'ArrowUp') return;
            event.preventDefault();
            focusEdgeRef.current = event.key === 'ArrowUp' ? 'last' : 'first';
            if (!open) event.currentTarget.click();
            else focusRequestedAction();
          }}
        >
          <span>文件</span>
          <ChevronDown size={12} aria-hidden="true" />
        </button>
      )}
    >
      {(close) => (
        <>
          {actions.map((action) => (
            <Fragment key={action.id}>
              {action.separatorBefore && (
                <hr className="work-office-file-separator" />
              )}
              <button
                type="button"
                role="menuitem"
                tabIndex={-1}
                disabled={action.disabled}
                onClick={() => {
                  close();
                  void action.onSelect();
                }}
              >
                <span
                  className="work-office-file-action-icon"
                  aria-hidden="true"
                >
                  {action.icon}
                </span>
                <span>{action.label}</span>
                {action.shortcut && <kbd>{action.shortcut}</kbd>}
              </button>
            </Fragment>
          ))}
        </>
      )}
    </Popover>
  );
}

export function WorkOfficeRibbonGroup({
  label,
  children,
}: {
  label: string;
  children: ReactNode;
}) {
  return (
    <section className="work-office-ribbon-group" aria-label={label}>
      <div>{children}</div>
      <span aria-hidden="true">{label}</span>
    </section>
  );
}

export function WorkOfficeRibbonButton({
  label,
  visibleLabel = label,
  badge,
  active,
  displayLabel = true,
  className = '',
  title = label,
  children,
  ...props
}: Omit<ButtonHTMLAttributes<HTMLButtonElement>, 'aria-label' | 'children'> & {
  label: string;
  visibleLabel?: string;
  badge?: ReactNode;
  active?: boolean;
  displayLabel?: boolean;
  children: ReactNode;
}) {
  return (
    <button
      type="button"
      aria-label={label}
      aria-pressed={active}
      title={title}
      className={`${displayLabel ? 'with-label' : ''} ${active ? 'active' : ''} ${className}`.trim()}
      {...props}
    >
      {children}
      {displayLabel && <span>{visibleLabel}</span>}
      {badge !== undefined && (
        <small className="work-office-ribbon-badge">{badge}</small>
      )}
    </button>
  );
}

export function WorkOfficeStatusBar({
  children,
  controls,
  className = '',
}: {
  children: ReactNode;
  controls?: ReactNode;
  className?: string;
}) {
  return (
    <footer className={`work-office-status ${className}`.trim()}>
      <div className="work-office-status-info">{children}</div>
      {controls && <div className="work-office-status-view">{controls}</div>}
    </footer>
  );
}

export function WorkOfficeZoomControls({
  zoom,
  minimum = 50,
  maximum = 200,
  step = 10,
  decreaseLabel,
  increaseLabel,
  outputLabel,
  sliderLabel,
  onChange,
}: {
  zoom: number;
  minimum?: number;
  maximum?: number;
  step?: number;
  decreaseLabel: string;
  increaseLabel: string;
  outputLabel: string;
  sliderLabel: string;
  onChange: (zoom: number) => void;
}) {
  const clamp = (value: number) =>
    Math.min(maximum, Math.max(minimum, Math.round(value)));
  return (
    <>
      <button
        type="button"
        aria-label={decreaseLabel}
        title={decreaseLabel}
        disabled={zoom <= minimum}
        onClick={() => onChange(clamp(zoom - step))}
      >
        <Minus size={13} />
      </button>
      <output aria-label={outputLabel}>{zoom}%</output>
      <OfficeSlider
        className="work-office-status-slider"
        ariaLabel={sliderLabel}
        min={minimum}
        max={maximum}
        step={step}
        value={zoom}
        onValueChange={(value) => onChange(clamp(value))}
      />
      <button
        type="button"
        aria-label={increaseLabel}
        title={increaseLabel}
        disabled={zoom >= maximum}
        onClick={() => onChange(clamp(zoom + step))}
      >
        <Plus size={13} />
      </button>
    </>
  );
}
