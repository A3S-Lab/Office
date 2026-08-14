import type { Editor } from '@tiptap/core';
import { Bot } from 'lucide-react';
import {
  type CSSProperties,
  type RefObject,
  useEffect,
  useLayoutEffect,
  useMemo,
  useRef,
  useState,
} from 'react';
import type {
  WorkOfficeCollaborationParticipant,
  WorkOfficeCollaborationPresence,
  WorkOfficeCollaborationPresenceLocation,
  WorkOfficeCollaborationPresenceSnapshot,
} from '../../../collaboration/office-collaboration-presence';
import { useOfficeCollaborationPresence } from './office-collaboration-presence-context';

const PRESENCE_COLORS = [
  '#1d4ed8',
  '#6d28d9',
  '#be123c',
  '#047857',
  '#a16207',
  '#0369a1',
  '#9d174d',
  '#4338ca',
] as const;
const MAX_PROJECTED_PARTICIPANTS = 24;
const MAX_TEXT_SELECTION_RECTS = 64;

export type OfficePresenceColorStyle = CSSProperties & {
  '--work-office-presence-color': string;
};

export function useOfficePresenceSnapshot():
  | WorkOfficeCollaborationPresenceSnapshot
  | undefined {
  const presence = useOfficeCollaborationPresence();
  const [state, setState] = useState<{
    owner?: WorkOfficeCollaborationPresence;
    snapshot?: WorkOfficeCollaborationPresenceSnapshot;
  }>(() => ({ owner: presence, snapshot: presence?.snapshot() }));

  useEffect(() => {
    if (!presence) {
      setState({ owner: undefined, snapshot: undefined });
      return;
    }
    setState({ owner: presence, snapshot: presence.snapshot() });
    return presence.subscribe((snapshot) =>
      setState({ owner: presence, snapshot }),
    );
  }, [presence]);

  return state.owner === presence ? state.snapshot : presence?.snapshot();
}

export function useOfficeRemoteParticipants(): readonly WorkOfficeCollaborationParticipant[] {
  const snapshot = useOfficePresenceSnapshot();
  return useMemo(
    () =>
      (snapshot?.participants ?? [])
        .filter((participant) => !participant.local && participant.location)
        .slice(0, MAX_PROJECTED_PARTICIPANTS),
    [snapshot],
  );
}

export function useOfficePublishPresenceLocation(
  location: WorkOfficeCollaborationPresenceLocation | null,
): void {
  const presence = useOfficeCollaborationPresence();
  const locationRef = useRef(location);
  const publishedKeyRef = useRef<string | null>(null);
  locationRef.current = location;
  const locationKey = presenceLocationKey(location);

  useEffect(() => {
    if (!presence) return;
    const next = locationRef.current;
    const current = presence.local().location ?? null;
    if (presenceLocationKey(current) !== locationKey) {
      presence.update({ location: next });
    }
    publishedKeyRef.current = locationKey;
  }, [locationKey, presence]);

  useEffect(
    () => () => {
      if (!presence || publishedKeyRef.current === null) return;
      try {
        const current = presence.local().location ?? null;
        if (presenceLocationKey(current) === publishedKeyRef.current) {
          presence.update({ location: null });
        }
      } catch {
        // The host may destroy its Presence owner before React unmounts.
      }
    },
    [presence],
  );
}

export function OfficePresenceAvatar({
  compact = false,
  participant,
}: {
  compact?: boolean;
  participant: WorkOfficeCollaborationParticipant;
}) {
  const avatarUrl = safeAvatarUrl(participant.actor.avatarUrl);
  return (
    <span
      className="work-office-collaboration-avatar"
      data-activity={participant.activity}
      data-compact={compact ? 'true' : undefined}
      style={officePresenceColorStyle(participant)}
      aria-hidden="true"
    >
      {participant.actor.kind === 'agent' ? (
        <Bot size={compact ? 10 : 12} />
      ) : (
        <span>{participantInitial(participant.actor.name)}</span>
      )}
      {avatarUrl && (
        <img
          src={avatarUrl}
          alt=""
          loading="lazy"
          referrerPolicy="no-referrer"
          onError={(event) => {
            event.currentTarget.hidden = true;
          }}
        />
      )}
    </span>
  );
}

export function officePresenceColorStyle(
  participant: WorkOfficeCollaborationParticipant,
): OfficePresenceColorStyle {
  return {
    '--work-office-presence-color': officePresenceColor(participant),
  };
}

export function officePresenceColor(
  participant: WorkOfficeCollaborationParticipant,
): string {
  const requested = participant.actor.color;
  if (requested && safeCssColor(requested)) return requested;
  let hash = 0;
  for (const character of participant.actor.id) {
    hash = (hash * 31 + (character.codePointAt(0) ?? 0)) >>> 0;
  }
  return PRESENCE_COLORS[hash % PRESENCE_COLORS.length] ?? PRESENCE_COLORS[0];
}

interface OfficeTextPresenceGeometry {
  participant: WorkOfficeCollaborationParticipant;
  caret: OfficePresenceRect;
  ranges: readonly OfficePresenceRect[];
}

interface OfficePresenceRect {
  height: number;
  left: number;
  top: number;
  width: number;
}

export function OfficeTiptapPresenceLayer({
  containerRef,
  editor,
  kind,
  markdownSurface,
}: {
  containerRef: RefObject<HTMLElement | null>;
  editor: Editor | null;
  kind: 'document' | 'markdown';
  markdownSurface?: 'visual';
}) {
  const remoteParticipants = useOfficeRemoteParticipants();
  const [geometry, setGeometry] = useState<
    readonly OfficeTextPresenceGeometry[]
  >([]);

  useLayoutEffect(() => {
    if (!editor || editor.isDestroyed) {
      setGeometry([]);
      return;
    }
    let frame: number | null = null;
    const measure = () => {
      frame = null;
      setGeometry(
        measureTiptapPresence(
          editor,
          containerRef.current,
          remoteParticipants,
          kind,
          markdownSurface,
        ),
      );
    };
    const schedule = () => {
      if (frame !== null) cancelAnimationFrame(frame);
      frame = requestAnimationFrame(measure);
    };
    measure();
    editor.on('transaction', schedule);
    window.addEventListener('resize', schedule);
    window.addEventListener('scroll', schedule, true);
    const observer =
      typeof ResizeObserver === 'undefined'
        ? null
        : new ResizeObserver(schedule);
    if (containerRef.current) observer?.observe(containerRef.current);
    observer?.observe(editor.view.dom);
    return () => {
      if (frame !== null) cancelAnimationFrame(frame);
      editor.off('transaction', schedule);
      window.removeEventListener('resize', schedule);
      window.removeEventListener('scroll', schedule, true);
      observer?.disconnect();
    };
  }, [containerRef, editor, kind, markdownSurface, remoteParticipants]);

  if (!geometry.length) return null;
  return (
    <div
      className="work-office-remote-presence-layer"
      data-presence-surface={kind}
      aria-hidden="true"
    >
      {geometry.flatMap(({ caret, participant, ranges }) => {
        const style = officePresenceColorStyle(participant);
        return [
          ...ranges.map((range, index) => (
            <span
              className="work-office-remote-selection"
              data-activity={participant.activity}
              data-participant-id={participant.actor.id}
              key={`${participant.presenceId}:range:${index}`}
              style={{ ...style, ...range }}
            />
          )),
          <span
            className="work-office-remote-caret"
            data-activity={participant.activity}
            data-participant-id={participant.actor.id}
            key={`${participant.presenceId}:caret`}
            style={{ ...style, ...caret }}
          >
            <span>{participant.actor.name}</span>
          </span>,
        ];
      })}
    </div>
  );
}

export function MarkdownSourcePresenceLayer({
  markdown,
  sourceRef,
}: {
  markdown: string;
  sourceRef: RefObject<HTMLTextAreaElement | null>;
}) {
  const remoteParticipants = useOfficeRemoteParticipants();
  const [viewport, setViewport] = useState<{
    height: number;
    left: number;
    scrollLeft: number;
    scrollTop: number;
    scrollWidth: number;
    top: number;
    width: number;
  } | null>(null);
  const sourceParticipants = remoteParticipants.filter((participant) => {
    const location = participant.location;
    return (
      location?.kind === 'markdown' &&
      (location.surface ?? 'source') === 'source'
    );
  });

  useLayoutEffect(() => {
    const source = sourceRef.current;
    if (!source) {
      setViewport(null);
      return;
    }
    let frame: number | null = null;
    const measure = () => {
      frame = null;
      setViewport({
        height: source.clientHeight,
        left: source.offsetLeft,
        scrollLeft: source.scrollLeft,
        scrollTop: source.scrollTop,
        scrollWidth: Math.max(source.clientWidth, source.scrollWidth),
        top: source.offsetTop,
        width: source.clientWidth,
      });
    };
    const schedule = () => {
      if (frame !== null) cancelAnimationFrame(frame);
      frame = requestAnimationFrame(measure);
    };
    measure();
    source.addEventListener('scroll', schedule);
    window.addEventListener('resize', schedule);
    const observer =
      typeof ResizeObserver === 'undefined'
        ? null
        : new ResizeObserver(schedule);
    observer?.observe(source);
    return () => {
      if (frame !== null) cancelAnimationFrame(frame);
      source.removeEventListener('scroll', schedule);
      window.removeEventListener('resize', schedule);
      observer?.disconnect();
    };
  }, [markdown, sourceRef]);

  if (!viewport || !sourceParticipants.length) return null;
  return (
    <div
      className="work-markdown-source-presence-layer"
      aria-hidden="true"
      style={{
        height: viewport.height,
        left: viewport.left,
        top: viewport.top,
        width: viewport.width,
      }}
    >
      {sourceParticipants.map((participant) => {
        const location = participant.location;
        if (location?.kind !== 'markdown') return null;
        if (
          !textPositionExists(location.anchor, markdown.length) ||
          !textPositionExists(location.head, markdown.length)
        ) {
          return null;
        }
        const { anchor, head } = location;
        const from = Math.min(anchor, head);
        const to = Math.max(anchor, head);
        const collapsed = from === to;
        return (
          <pre
            className="work-markdown-source-presence-mirror"
            data-activity={participant.activity}
            data-participant-id={participant.actor.id}
            key={participant.presenceId}
            style={{
              ...officePresenceColorStyle(participant),
              transform: `translate(${-viewport.scrollLeft}px, ${-viewport.scrollTop}px)`,
              width: viewport.scrollWidth,
            }}
          >
            {markdown.slice(0, from)}
            <span
              className="work-markdown-source-remote-range"
              data-collapsed={collapsed ? 'true' : undefined}
              data-participant-name={participant.actor.name}
            >
              {collapsed ? '\u200b' : markdown.slice(from, to)}
            </span>
            {markdown.slice(to)}
          </pre>
        );
      })}
    </div>
  );
}

function measureTiptapPresence(
  editor: Editor,
  container: HTMLElement | null,
  participants: readonly WorkOfficeCollaborationParticipant[],
  kind: 'document' | 'markdown',
  markdownSurface?: 'visual',
): readonly OfficeTextPresenceGeometry[] {
  if (!container?.isConnected || editor.isDestroyed) return [];
  const containerBounds = container.getBoundingClientRect();
  if (containerBounds.width <= 0 || containerBounds.height <= 0) return [];
  const scaleX = safeElementScale(containerBounds.width, container.offsetWidth);
  const scaleY = safeElementScale(
    containerBounds.height,
    container.offsetHeight,
  );
  const maximum = editor.state.doc.content.size;
  const result: OfficeTextPresenceGeometry[] = [];
  for (const participant of participants) {
    const location = participant.location;
    if (location?.kind !== kind) continue;
    if (
      kind === 'markdown' &&
      location.kind === 'markdown' &&
      (location.surface ?? 'source') !== markdownSurface
    ) {
      continue;
    }
    if (
      !textPositionExists(location.anchor, maximum) ||
      !textPositionExists(location.head, maximum)
    ) {
      continue;
    }
    const { anchor, head } = location;
    const from = Math.min(anchor, head);
    const to = Math.max(anchor, head);
    try {
      const caretBounds = editor.view.coordsAtPos(head);
      const caret = localPresenceRect(
        {
          height: Math.max(1, caretBounds.bottom - caretBounds.top),
          left: caretBounds.left,
          top: caretBounds.top,
          width: 2,
        },
        container,
        containerBounds,
        scaleX,
        scaleY,
      );
      const ranges =
        from === to
          ? []
          : tiptapSelectionRects(editor, from, to)
              .slice(0, MAX_TEXT_SELECTION_RECTS)
              .map((bounds) =>
                localPresenceRect(
                  bounds,
                  container,
                  containerBounds,
                  scaleX,
                  scaleY,
                ),
              )
              .filter((rect) => rect.width > 0 && rect.height > 0);
      result.push({ caret, participant, ranges });
    } catch {
      // Stale or out-of-range remote locations are ignored until refreshed.
    }
  }
  return result;
}

function tiptapSelectionRects(
  editor: Editor,
  from: number,
  to: number,
): readonly DOMRect[] {
  const start = editor.view.domAtPos(from);
  const end = editor.view.domAtPos(to);
  const range = document.createRange();
  range.setStart(start.node, start.offset);
  range.setEnd(end.node, end.offset);
  return Array.from(range.getClientRects());
}

function localPresenceRect(
  bounds: Pick<DOMRect, 'height' | 'left' | 'top' | 'width'>,
  container: HTMLElement,
  containerBounds: DOMRect,
  scaleX: number,
  scaleY: number,
): OfficePresenceRect {
  return {
    height: Math.max(1, bounds.height / scaleY),
    left: (bounds.left - containerBounds.left) / scaleX + container.scrollLeft,
    top: (bounds.top - containerBounds.top) / scaleY + container.scrollTop,
    width: Math.max(1, bounds.width / scaleX),
  };
}

function safeElementScale(rendered: number, layout: number): number {
  const scale = layout > 0 ? rendered / layout : 1;
  return Number.isFinite(scale) && scale > 0 ? scale : 1;
}

function textPositionExists(value: number, maximum: number): boolean {
  return Number.isSafeInteger(value) && value >= 0 && value <= maximum;
}

function presenceLocationKey(
  location: WorkOfficeCollaborationPresenceLocation | null,
): string {
  return location ? JSON.stringify(location) : '';
}

function participantInitial(name: string): string {
  return Array.from(name.trim())[0]?.toLocaleUpperCase() ?? '?';
}

function safeCssColor(value: string): boolean {
  if (/[;{}]|url\(|var\(/i.test(value)) return false;
  if (/^#(?:[\da-f]{3}|[\da-f]{4}|[\da-f]{6}|[\da-f]{8})$/i.test(value)) {
    return true;
  }
  return typeof CSS !== 'undefined' && CSS.supports('color', value);
}

function safeAvatarUrl(value: string | undefined): string | undefined {
  if (!value || typeof URL === 'undefined') return undefined;
  try {
    const url = new URL(value, 'https://a3s.invalid/');
    if (
      url.protocol === 'https:' ||
      url.protocol === 'http:' ||
      url.protocol === 'blob:' ||
      (url.protocol === 'data:' && value.startsWith('data:image/'))
    ) {
      return value;
    }
  } catch {
    return undefined;
  }
  return undefined;
}
