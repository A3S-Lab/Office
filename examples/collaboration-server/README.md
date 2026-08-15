# A3S Office collaboration server

This directory is a runnable reference backend for A3S Office real-time
collaboration. It combines:

- `a3s-boot` 0.2 WebSocket gateways, dependency injection, controllers, rooms,
  lifecycle hooks, and ACL configuration;
- the `a3s-office` native Yrs replica for crash-safe Yjs v1 persistence;
- short-lived HMAC-signed room tickets and an explicit WebSocket Origin allowlist;
- document synchronization, ephemeral Yjs Awareness, reconnect repair, and
  polling of the service-owned durable replica.

It supports the same `document`, `markdown`, `spreadsheet`, `presentation`, and
`pdf` collaboration envelopes as the browser package. The server never stores
source PDF bytes in Yjs; it persists only the PDF collaboration overlay.

## Architecture

```text
host backend ── POST /api/collaboration/tickets ──> signed 5-minute ticket
                                                        │
browser Y.Doc + Awareness ── WebSocket ──> A3S Boot room│
native replica ── host JSONL bridge ─────>       │      │
                                                ├─ Yrs durable replica
                                                ├─ Awareness relay (memory only)
                                                └─ service-owned event poller
```

Every document update is validated against the ticket room and Yjs client ID,
persisted before broadcast, and assigned a deterministic delivery operation ID.
Duplicate delivery is idempotent. Only `edit` tickets may send `sync-step-2` or
`update`; `view`, `comment`, and `suggest` tickets can synchronize down and
publish presence but cannot mutate canonical content.

## Run locally

From the Office repository root:

```bash
export A3S_OFFICE_TICKET_SECRET="$(openssl rand -hex 32)"
export A3S_OFFICE_ADMIN_TOKEN="$(openssl rand -hex 24)"
cargo run -p a3s-office-collaboration-server
```

The default ACL file is
[`collaboration-server.acl`](./collaboration-server.acl). Pass another file as
the first argument when deploying:

```bash
cargo run -p a3s-office-collaboration-server -- /etc/a3s/office-collaboration.acl
```

The service exposes:

| Endpoint | Purpose |
| --- | --- |
| `GET /api/collaboration/healthz` | Health and protocol version |
| `POST /api/collaboration/tickets` | Host-only short-lived ticket issuer |
| `WS /collaboration/{kind}/{artifactId}?ticket=...` | Authenticated room transport |

Issue a local ticket from a trusted host process:

```bash
curl --fail http://127.0.0.1:8787/api/collaboration/tickets \
  --header "Authorization: Bearer ${A3S_OFFICE_ADMIN_TOKEN}" \
  --header 'Content-Type: application/json' \
  --data '{
    "artifactId": "quarterly-plan",
    "artifactKind": "document",
    "actorId": "user-42",
    "actorKind": "human",
    "mode": "edit"
  }'
```

The response follows the repository API envelope. Pass
`data.webSocketUrl` to the browser adapter. Do not expose the admin token in
browser code; a real host endpoint derives the ticket claims from its existing
authenticated application session.

## Connect an editor

[`client.ts`](./client.ts) is a complete browser adapter. It converts the
structured Office transport envelope to JSON with bounded base64 payloads,
relays Awareness, removes stale presence after disconnect, and performs a
two-way state-vector handshake after every reconnect.

```tsx
import * as Y from 'yjs';
import { Awareness } from 'y-protocols/awareness';
import {
  createOfficeCollaborationPresence,
  createOfficeCollaborationSession,
  initializeOfficeDocumentCollaboration,
} from '@a3s-lab/office/core';
import { DocumentEditor } from '@a3s-lab/office/react';
import { connectA3sBootCollaborationRoom } from './client';

const document = new Y.Doc();
const awareness = new Awareness(document);
const session = createOfficeCollaborationSession({
  artifactId: 'quarterly-plan',
  kind: 'document',
  document,
  awareness,
  actor: { id: 'user-42', name: 'Mina', kind: 'human' },
  mode: 'edit',
});

if (!session.metadata()?.initialized) {
  initializeOfficeDocumentCollaboration(session, initialContent);
}

const presence = createOfficeCollaborationPresence(session);
const room = connectA3sBootCollaborationRoom({
  webSocketUrl: ticketResponse.data.webSocketUrl,
  session,
  awareness,
  onStatus: (status, detail) => reportConnection(status, detail),
});

root.render(
  <DocumentEditor
    content={initialContent}
    collaboration={session}
    presence={presence}
    onChange={reportSnapshot}
  />,
);

// Component unmount:
room.destroy();
presence.destroy();
session.destroy();
awareness.destroy();
document.destroy();
```

## Wire events

All Boot WebSocket frames are JSON objects with `{ "event", "data" }`.

| Event | Direction | Durable |
| --- | --- | --- |
| `collaboration.hello` | client to server | No |
| `collaboration.ready` | server to client | No |
| `collaboration.document` | both | Yes, for received updates |
| `collaboration.ack` | server to sender | Receipt only |
| `collaboration.awareness` | both | No |
| `collaboration.peer-left` | server to room | No |
| `collaboration.error` | server to sender | No |

The document event carries the Office protocol identity, message type, sender
Yjs client ID, optional trusted origin, and `payloadBase64`. Decoded payloads are
standard Yjs v1 state vectors or updates; the browser adapter restores the
`Uint8Array` expected by `createOfficeCollaborationTransportBinding`.

## Connect a native agent

Issue the agent an `actorKind: "agent"` ticket and relay the same
`collaboration.document` envelope over the WebSocket. The transport-neutral
`a3s-office collab session` command writes outbound envelopes as JSONL and
accepts inbound `receive` records on stdin, so the product host can bridge it
to this room without giving the CLI an application credential.

The host owns that process bridge and reconnect lifecycle. Do not point an
agent directly at the server's internal `data_dir`: each native participant
keeps its own actor-scoped replica, while the service keeps the durable room
replica. This preserves actor attribution and independent Yjs client IDs.

## Production boundary

The reference implementation is complete for one service process. For multiple
replicas, use sticky room routing or add an A3S Boot Redis/NATS transport for
cross-process broadcast and a shared durable store with one writer/lock policy.
Keep the following boundaries unchanged:

- mint tickets only after application authentication and document authorization;
- bind actor, artifact, kind, namespace, mode, expiration, and Origin;
- persist document updates before acknowledging or broadcasting them;
- never persist Awareness as document state;
- retain the deterministic operation ID during retries;
- terminate TLS at the service or a trusted reverse proxy and use `wss://`;
- cap frame size, connection count, ticket issuance, and room fan-out.

## Verify

```bash
cargo test -p a3s-office-collaboration-server
bun run collaboration-server:typecheck
```

The Rust integration test opens two authenticated Boot gateway connections,
persists a browser-generated Yjs update, verifies room broadcast and restart-safe
Yrs state, and confirms a view-only ticket cannot publish content.
