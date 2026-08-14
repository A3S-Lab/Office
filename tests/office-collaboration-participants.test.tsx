import { expect, test } from '@rstest/core';
import {
  fireEvent,
  render,
  screen,
  waitFor,
  within,
} from '@testing-library/react';
import {
  Awareness,
  applyAwarenessUpdate,
  encodeAwarenessUpdate,
} from 'y-protocols/awareness';
import * as Y from 'yjs';
import {
  createOfficeCollaborationPresence,
  createOfficeCollaborationSession,
  type OfficeCollaborationParticipant,
} from '../src/core';
import { WorkOfficeCollaborationParticipants } from '../src/internal/features/work/editors/office-collaboration-participants';
import {
  assertOfficeCollaborationPresencePairing,
  OfficeCollaborationPresenceProvider,
  useOfficeCollaborationLocationNavigator,
} from '../src/internal/features/work/editors/office-collaboration-presence-context';
import { WorkOfficeStatusBar } from '../src/internal/features/work/editors/work-office-chrome';

test('projects live human and agent participants into an accessible roster', async () => {
  const fixture = collaborationPresenceFixture();
  const view = render(
    <OfficeCollaborationPresenceProvider presence={fixture.localPresence}>
      <WorkOfficeStatusBar controls={<button type="button">页面视图</button>}>
        已保存
      </WorkOfficeStatusBar>
    </OfficeCollaborationPresenceProvider>,
  );

  try {
    const trigger = screen.getByRole('button', {
      name: '查看协作者，2 位协作者',
    });
    expect(trigger).toHaveAttribute('data-collaboration-count', '2');
    expect(trigger).toHaveTextContent('2 位协作者');
    expect(trigger.closest('.work-office-status')).not.toBeNull();

    trigger.focus();
    fireEvent.click(trigger);
    const roster = screen.getByRole('dialog', { name: '协作者' });
    expect(within(roster).getByText('2 个在线会话')).toBeVisible();
    expect(within(roster).getByText('Grace')).toBeVisible();
    expect(within(roster).getByText('你')).toBeVisible();
    expect(within(roster).getByText('A3S Agent')).toBeVisible();
    expect(within(roster).getByText('Agent')).toBeVisible();
    const remoteParticipant = roster.querySelector(
      '[data-collaboration-participant="agent"]',
    );
    expect(remoteParticipant).toHaveTextContent('已选择 6 个位置');
    expect(remoteParticipant).toHaveAttribute('data-activity', 'active');

    fixture.remotePresence.update({ activity: 'away' });
    relayAwareness(fixture.remoteAwareness, fixture.localAwareness);
    await waitFor(() =>
      expect(within(roster).getByText(/离开 · 建议/)).toBeVisible(),
    );

    fireEvent.keyDown(roster, { key: 'Escape' });
    await waitFor(() => expect(trigger).toHaveFocus());
    expect(screen.queryByRole('dialog', { name: '协作者' })).toBeNull();
  } finally {
    view.unmount();
    fixture.destroy();
  }
});

test('renders the compact toolbar projection without a status bar dependency', () => {
  const fixture = collaborationPresenceFixture();
  const view = render(
    <OfficeCollaborationPresenceProvider presence={fixture.localPresence}>
      <header role="toolbar" aria-label="PDF 工具栏">
        <WorkOfficeCollaborationParticipants variant="toolbar" />
      </header>
    </OfficeCollaborationPresenceProvider>,
  );

  try {
    const trigger = screen.getByRole('button', {
      name: '查看协作者，2 位协作者',
    });
    expect(trigger).toHaveAttribute('data-variant', 'toolbar');
    expect(trigger).toHaveTextContent('2');
  } finally {
    view.unmount();
    fixture.destroy();
  }
});

test('navigates from a remote roster row and leaves focus in the editor', async () => {
  const fixture = collaborationPresenceFixture();
  const calls: OfficeCollaborationParticipant[] = [];
  const view = render(
    <OfficeCollaborationPresenceProvider presence={fixture.localPresence}>
      <ParticipantNavigationProbe
        onNavigate={(participant) => {
          calls.push(participant);
          screen.getByRole('textbox', { name: '文档画布' }).focus();
          return true;
        }}
      />
      <div>
        <textarea aria-label="文档画布" />
        <WorkOfficeStatusBar>已保存</WorkOfficeStatusBar>
      </div>
    </OfficeCollaborationPresenceProvider>,
  );

  try {
    fireEvent.click(
      screen.getByRole('button', { name: '查看协作者，2 位协作者' }),
    );
    const locate = await screen.findByRole('button', {
      name: '跳转到 A3S Agent 的位置，已选择 6 个位置',
    });
    expect(locate).toHaveFocus();
    fireEvent.click(locate);

    await waitFor(() => expect(calls).toHaveLength(1));
    expect(calls[0]?.actor.id).toBe('agent');
    expect(screen.getByRole('textbox', { name: '文档画布' })).toHaveFocus();
    expect(screen.queryByRole('dialog', { name: '协作者' })).toBeNull();
  } finally {
    view.unmount();
    fixture.destroy();
  }
});

test('fails closed when presence is not paired with the exact editor session', () => {
  const fixture = collaborationPresenceFixture();
  const otherArtifact = createOfficeCollaborationSession({
    actor: { id: 'grace', name: 'Grace', color: '#047857' },
    artifactId: 'another-document',
    kind: 'document',
  });
  const otherClient = createOfficeCollaborationSession({
    actor: { id: 'grace', name: 'Grace', color: '#047857' },
    artifactId: 'presence-ui-document',
    kind: 'document',
  });

  try {
    expect(() =>
      assertOfficeCollaborationPresencePairing({
        expectedKind: 'document',
        presence: fixture.localPresence,
      }),
    ).toThrow(/must be paired/);
    expect(() =>
      assertOfficeCollaborationPresencePairing({
        expectedKind: 'markdown',
        presence: fixture.localPresence,
        session: fixture.localSession,
      }),
    ).toThrow(/both target 'markdown'/);
    expect(() =>
      assertOfficeCollaborationPresencePairing({
        expectedKind: 'document',
        presence: fixture.localPresence,
        session: otherArtifact,
      }),
    ).toThrow(/another collaboration artifact/);
    expect(() =>
      assertOfficeCollaborationPresencePairing({
        expectedKind: 'document',
        presence: fixture.localPresence,
        session: otherClient,
      }),
    ).toThrow(/does not belong to the supplied collaboration session/);
  } finally {
    otherArtifact.destroy();
    otherClient.destroy();
    fixture.destroy();
  }
});

function collaborationPresenceFixture() {
  const localDocument = new Y.Doc();
  const remoteDocument = new Y.Doc();
  const localAwareness = new Awareness(localDocument);
  const remoteAwareness = new Awareness(remoteDocument);
  const localSession = createOfficeCollaborationSession({
    actor: { id: 'grace', name: 'Grace', color: '#047857' },
    artifactId: 'presence-ui-document',
    awareness: localAwareness,
    document: localDocument,
    kind: 'document',
  });
  const remoteSession = createOfficeCollaborationSession({
    actor: {
      id: 'agent',
      name: 'A3S Agent',
      color: '#6d28d9',
      kind: 'agent',
    },
    artifactId: 'presence-ui-document',
    awareness: remoteAwareness,
    document: remoteDocument,
    kind: 'document',
    mode: 'suggest',
  });
  const localPresence = createOfficeCollaborationPresence(localSession);
  const remotePresence = createOfficeCollaborationPresence(remoteSession, {
    location: { kind: 'document', anchor: 3, head: 9 },
  });
  relayAwareness(remoteAwareness, localAwareness);

  return {
    localAwareness,
    localPresence,
    localSession,
    remoteAwareness,
    remotePresence,
    destroy() {
      localPresence.destroy();
      remotePresence.destroy();
      localSession.destroy();
      remoteSession.destroy();
      localAwareness.destroy();
      remoteAwareness.destroy();
      localDocument.destroy();
      remoteDocument.destroy();
    },
  };
}

function ParticipantNavigationProbe({
  onNavigate,
}: {
  onNavigate: (participant: OfficeCollaborationParticipant) => boolean;
}) {
  useOfficeCollaborationLocationNavigator(onNavigate);
  return null;
}

function relayAwareness(source: Awareness, target: Awareness): void {
  applyAwarenessUpdate(
    target,
    encodeAwarenessUpdate(source, [source.clientID]),
    'test-transport',
  );
}
