import {
  ArrowUpRight,
  ChevronDown,
  ChevronLeft,
  ChevronRight,
  ChevronUp,
  Eye,
  File as FileIcon,
  Minus,
  Plus,
} from 'lucide-react';
import {
  type ButtonHTMLAttributes,
  Fragment,
  type KeyboardEvent,
  type ReactNode,
  useCallback,
  useEffect,
  useId,
  useLayoutEffect,
  useRef,
  useState,
} from 'react';
import { Popover, Tabs } from '../../../design-system/primitives';
import { WorkOfficeCollaborationParticipants } from './office-collaboration-participants';
import { OfficeSlider } from './office-controls';
import { moveOfficeMenuFocus } from './office-menu-keyboard';
import {
  calculateRibbonDensity,
  calculateRibbonOverflow,
  calculateRibbonScrollTarget,
  type WorkOfficeRibbonDensity,
  type WorkOfficeRibbonItemGeometry,
} from './work-office-ribbon-overflow';

export interface WorkOfficeRibbonTab<T extends string> {
  id: T;
  label: string;
  compactLabel?: string;
  contextual?: boolean;
}

export interface WorkOfficeFileAction {
  id: string;
  label: string;
  /** Uses a neutral file glyph when omitted so the menu never reserves a blank icon cell. */
  icon?: ReactNode;
  shortcut?: string;
  disabled?: boolean;
  /** Visually distinguishes an irreversible action without changing its behavior. */
  danger?: boolean;
  separatorBefore?: boolean;
  onSelect: () => void | Promise<void>;
}

export interface WorkOfficeQuickAccessAction {
  id: string;
  label: string;
  icon: ReactNode;
  shortcut?: string;
  ariaKeyShortcuts?: string;
  disabled?: boolean;
  onSelect: () => void | Promise<void>;
}

interface WorkOfficeRibbonDialogLauncher {
  label: string;
  shortcut?: string;
  ariaKeyShortcuts?: string;
  onClick: () => void;
}

export function WorkOfficeRibbon<T extends string>({
  ariaLabel,
  tabs,
  defaultTab,
  activeTab,
  onTabChange,
  panels,
  fileActions,
  quickAccessActions,
  adaptive = false,
  collapsible = false,
  collapsed: controlledCollapsed,
  defaultCollapsed = false,
  onCollapsedChange,
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
  quickAccessActions?: readonly WorkOfficeQuickAccessAction[];
  adaptive?: boolean;
  collapsible?: boolean;
  collapsed?: boolean;
  defaultCollapsed?: boolean;
  onCollapsedChange?: (collapsed: boolean) => void;
  className?: string;
  toolbarClassName?: string;
}) {
  const reactId = useId().replaceAll(':', '');
  const [internalTab, setInternalTab] = useState(defaultTab);
  const [internalCollapsed, setInternalCollapsed] = useState(defaultCollapsed);
  const [temporarilyExpanded, setTemporarilyExpanded] = useState(false);
  const [ribbonDensity, setRibbonDensity] =
    useState<WorkOfficeRibbonDensity>('comfortable');
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
  const ribbonRef = useRef<HTMLElement>(null);
  const selectedTab = activeTab ?? internalTab;
  const selectedLabel =
    tabs.find((tab) => tab.id === selectedTab)?.label ?? tabs[0]?.label ?? '';
  const contextualTabActive = Boolean(
    tabs.find((tab) => tab.id === selectedTab)?.contextual,
  );
  const selectedPanel = panels[selectedTab];
  const hasSelectedPanel =
    selectedPanel !== null && selectedPanel !== undefined;
  const collapsed = collapsible && (controlledCollapsed ?? internalCollapsed);
  const panelExpanded = hasSelectedPanel && (!collapsed || temporarilyExpanded);
  const hasRibbonOverflow = ribbonOverflow.backward || ribbonOverflow.forward;
  const updateRibbonOverflow = useCallback(() => {
    const toolbar = toolbarRef.current;
    if (!toolbar) return;
    const ribbonWidth = ribbonRef.current?.clientWidth || toolbar.clientWidth;
    const nextDensity = adaptive
      ? calculateRibbonDensity(ribbonWidth)
      : 'comfortable';
    setRibbonDensity((current) =>
      current === nextDensity ? current : nextDensity,
    );
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
  }, [adaptive]);
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
    if (collapsed) setTemporarilyExpanded(true);
    onTabChange?.(tab);
  };
  const changeCollapsed = useCallback(
    (nextCollapsed: boolean) => {
      if (controlledCollapsed === undefined) {
        setInternalCollapsed(nextCollapsed);
      }
      setTemporarilyExpanded(false);
      onCollapsedChange?.(nextCollapsed);
    },
    [controlledCollapsed, onCollapsedChange],
  );

  useEffect(() => {
    if (!collapsed || !temporarilyExpanded) return;
    const closeTemporaryPanel = (event: PointerEvent) => {
      if (
        event.target instanceof Node &&
        ribbonRef.current?.contains(event.target)
      ) {
        return;
      }
      setTemporarilyExpanded(false);
    };
    document.addEventListener('pointerdown', closeTemporaryPanel, true);
    return () =>
      document.removeEventListener('pointerdown', closeTemporaryPanel, true);
  }, [collapsed, temporarilyExpanded]);

  useLayoutEffect(() => {
    const toolbar = toolbarRef.current;
    if (!toolbar) return;
    toolbar.scrollLeft = 0;
    updateRibbonOverflow();
  }, [selectedTab, panelExpanded, updateRibbonOverflow]);

  useLayoutEffect(() => {
    const toolbar = toolbarRef.current;
    if (!toolbar || !panelExpanded) return;
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
  }, [panelExpanded, updateRibbonOverflow]);

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
      ref={ribbonRef}
      className={`work-office-ribbon ${className}`.trim()}
      aria-label={ariaLabel}
      data-collapsed={collapsed ? 'true' : undefined}
      data-contextual-active={contextualTabActive ? 'true' : undefined}
    >
      <div className="work-office-ribbon-tabs-row">
        {fileActions?.length ? (
          <WorkOfficeFileMenu actions={fileActions} />
        ) : null}
        {quickAccessActions?.length ? (
          <WorkOfficeQuickAccessToolbar actions={quickAccessActions} />
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
            onTabDoubleClick={(tab) => {
              if (collapsible && tab === selectedTab) {
                changeCollapsed(!collapsed);
              }
            }}
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
        {collapsible && (
          <button
            type="button"
            className="work-office-ribbon-collapse"
            aria-label={collapsed ? '展开功能区' : '折叠功能区'}
            title={collapsed ? '展开功能区' : '折叠功能区'}
            aria-controls={`${reactId}-panel`}
            aria-expanded={panelExpanded}
            onClick={() => changeCollapsed(!collapsed)}
          >
            {collapsed ? (
              <ChevronDown size={14} aria-hidden="true" />
            ) : (
              <ChevronUp size={14} aria-hidden="true" />
            )}
          </button>
        )}
      </div>
      <div
        id={`${reactId}-panel`}
        className="work-office-ribbon-panel"
        data-empty={hasSelectedPanel ? undefined : 'true'}
        hidden={!panelExpanded}
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
              data-density={ribbonDensity}
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

function WorkOfficeQuickAccessToolbar({
  actions,
}: {
  actions: readonly WorkOfficeQuickAccessAction[];
}) {
  return (
    <div
      className="work-office-quick-access"
      role="toolbar"
      aria-label="快速访问工具栏"
    >
      {actions.map((action) => (
        <button
          key={action.id}
          type="button"
          aria-label={action.label}
          aria-keyshortcuts={action.ariaKeyShortcuts}
          title={
            action.shortcut
              ? `${action.label}（${action.shortcut}）`
              : action.label
          }
          disabled={action.disabled}
          onClick={() => void action.onSelect()}
        >
          <span aria-hidden="true">{action.icon}</span>
        </button>
      ))}
    </div>
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
      <WorkOfficeCollaborationParticipants variant="toolbar" />
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
                className={action.danger ? 'danger' : undefined}
                data-danger={action.danger ? 'true' : undefined}
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
                  {action.icon ?? <FileIcon size={16} />}
                </span>
                <span className="work-office-file-action-label">
                  {action.label}
                </span>
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
  priority = 'normal',
  dialogLauncher,
  children,
}: {
  label: string;
  priority?: 'high' | 'normal' | 'low';
  dialogLauncher?: WorkOfficeRibbonDialogLauncher;
  children: ReactNode;
}) {
  return (
    <section
      className="work-office-ribbon-group"
      aria-label={label}
      data-priority={priority}
      data-has-dialog-launcher={dialogLauncher ? 'true' : undefined}
    >
      <div>{children}</div>
      <span aria-hidden="true">{label}</span>
      {dialogLauncher && (
        <button
          type="button"
          className="work-office-ribbon-dialog-launcher"
          aria-label={dialogLauncher.label}
          aria-keyshortcuts={dialogLauncher.ariaKeyShortcuts}
          title={
            dialogLauncher.shortcut
              ? `${dialogLauncher.label}（${dialogLauncher.shortcut}）`
              : dialogLauncher.label
          }
          onMouseDown={(event) => event.preventDefault()}
          onClick={dialogLauncher.onClick}
        >
          <ArrowUpRight size={9} aria-hidden="true" />
        </button>
      )}
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
  ariaLabel = '编辑器状态栏',
  children,
  controls,
  controlsLabel = '视图与缩放',
  className = '',
}: {
  ariaLabel?: string;
  children: ReactNode;
  controls?: ReactNode;
  controlsLabel?: string;
  className?: string;
}) {
  return (
    <section
      className={`work-office-status ${className}`.trim()}
      aria-label={ariaLabel}
    >
      <div className="work-office-status-info">{children}</div>
      <WorkOfficeCollaborationParticipants />
      {controls && (
        <div
          className="work-office-status-view"
          role="toolbar"
          aria-label={controlsLabel}
          onKeyDown={moveOfficeToolbarFocus}
        >
          {controls}
        </div>
      )}
    </section>
  );
}

function moveOfficeToolbarFocus(event: KeyboardEvent<HTMLDivElement>): void {
  if (
    !['ArrowLeft', 'ArrowRight', 'Home', 'End'].includes(event.key) ||
    (event.target instanceof HTMLElement &&
      event.target.getAttribute('role') === 'slider')
  ) {
    return;
  }
  const controls = [
    ...event.currentTarget.querySelectorAll<HTMLElement>(
      'button:not(:disabled), [role="slider"]:not([aria-disabled="true"])',
    ),
  ];
  if (!controls.length) return;
  const currentIndex = controls.indexOf(document.activeElement as HTMLElement);
  const nextIndex =
    event.key === 'Home'
      ? 0
      : event.key === 'End'
        ? controls.length - 1
        : event.key === 'ArrowRight'
          ? (currentIndex + 1 + controls.length) % controls.length
          : (currentIndex - 1 + controls.length) % controls.length;
  event.preventDefault();
  controls[nextIndex]?.focus();
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
