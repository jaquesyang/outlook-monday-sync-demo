# Outlook ↔ monday.com Calendar Sync — Design Spec

- **Date**: 2026-05-05
- **Status**: Approved (brainstorming)
- **Owner**: jaquesyang
- **Repo**: `/Users/jaquesyang/Documents/monday.com/outlook_calendar_sync`

## 1. Goal

Build an Outlook add-in that bidirectionally syncs Outlook calendar events with monday.com items. A user installs the add-in (sideload), connects both monday and Microsoft Graph once, picks which monday boards to sync, and from then on changes flow in real time in both directions with last-write-wins conflict resolution.

## 2. Decisions Locked During Brainstorming

| Area | Decision |
|------|----------|
| Sync direction | Bidirectional |
| monday auth | OAuth 2.0 (multi-tenant) |
| monday data scope | User picks one or more boards in settings |
| Outlook auth | Microsoft Graph API + OAuth (Auth Code + PKCE), with Office.js SSO for identity |
| Stack | Next.js 16 App Router on Vercel, single repo |
| Database | Supabase Postgres (via Prisma) |
| Sync trigger | monday webhooks + Microsoft Graph subscriptions (real-time, both sides) |
| Conflict resolution | Last-write-wins with etag-based echo suppression |
| monday→Outlook column source | Date column **or** Timeline column, user picks per board |
| Outlook→monday for new events | New Outlook events that have no mapping flow into a user-configured default board + group |
| Distribution | Sideload only (manifest.xml). AppSource and tenant-wide deploy out of scope for v1 |
| Locales | zh-CN and en-US |

## 3. Architecture

```
Outlook (Web/Desktop)
└─ Add-in Taskpane (iframe → /taskpane Next page)
     └─ Office.js: identity (SSO ID token) + current item context

         │ HTTPS
         ▼
Next.js 16 App Router on Vercel (single repo)

  Routes:
    /taskpane                React UI: status, boards, settings, log
    /api/auth/microsoft/*    Microsoft Graph OAuth (popup, PKCE)
    /api/auth/monday/*       monday OAuth (popup)
    /api/webhooks/monday     monday webhook receiver
    /api/webhooks/graph      Graph subscription receiver
    /api/cron/refresh        hourly: refresh tokens, renew Graph subs, reconcile
    /api/sync/run            manual reconciliation
    /api/sync/one            single-event push from ribbon button
    /api/boards              list monday boards / toggle subscription
    /api/calendars           list user's Outlook calendars
    /api/status              taskpane status payload

  Libraries:
    lib/sync       sync engine (bidirectional, LWW, echo suppression)
    lib/monday     monday GraphQL client + webhook helpers
    lib/graph      Microsoft Graph client + subscription helpers
    lib/db         Prisma client + per-user scoped helpers
    lib/crypto     AES-256-GCM token encryption
    lib/auth       Office SSO ID-token validation + session

         │
         ▼
Supabase Postgres
  users, ms_accounts, monday_accounts,
  board_subscriptions, event_mappings, sync_log
```

Important properties:

- The taskpane never calls monday/Graph directly. All third-party calls go through `/api/*` so secrets stay server-side.
- Webhook handlers acknowledge with `200` immediately, then enqueue work for async processing (Vercel Queues beta — fall back to in-function `waitUntil` if Queues unavailable).
- A single Vercel cron (`schedule: '7 * * * *'`) handles token refresh, Graph subscription renewal (Graph subs expire ≤3 days), and reconciliation against missed webhooks.

## 4. Data Model

```sql
users (
  id              uuid pk,
  ms_tenant_id    text,
  ms_user_id      text,
  ms_user_email   text,
  created_at      timestamptz,
  unique (ms_tenant_id, ms_user_id)
);

ms_accounts (
  user_id              uuid pk fk users,
  access_token_enc     bytea,           -- AES-GCM ciphertext
  refresh_token_enc    bytea,
  expires_at           timestamptz,
  graph_subscription_id          text,
  graph_subscription_expires_at  timestamptz,
  selected_calendar_id text                -- user pick, default = primary
);

monday_accounts (
  user_id           uuid pk fk users,
  monday_user_id    bigint,
  monday_account_id bigint,
  access_token_enc  bytea,
  refresh_token_enc bytea,
  expires_at        timestamptz,
  default_board_id  bigint,                 -- where Outlook-originated events land
  default_group_id  text
);

board_subscriptions (
  id                 uuid pk,
  user_id            uuid fk users,
  monday_board_id    bigint,
  date_column_id     text,
  date_column_kind   text check (date_column_kind in ('date','timeline')),
  monday_webhook_id  bigint,
  active             boolean default true,
  unique(user_id, monday_board_id)
);

event_mappings (
  id                 uuid pk,
  user_id            uuid fk users,
  monday_item_id     bigint,
  monday_board_id    bigint,
  graph_event_id     text,
  graph_calendar_id  text,
  monday_etag        text,                  -- monday updated_at last seen
  graph_etag         text,                  -- Graph @odata.etag last seen
  origin             text check (origin in ('monday','outlook')),
  last_synced_at     timestamptz,
  deleted_at         timestamptz,
  unique (user_id, monday_item_id),
  unique (user_id, graph_event_id)
);

sync_log (
  id          uuid pk,
  user_id     uuid fk users,
  direction   text,
  mapping_id  uuid,
  action      text,                          -- create|update|delete|skip-echo|skip-conflict|error
  message     text,
  occurred_at timestamptz default now()
);
```

Notes:

- Token columns are AES-256-GCM ciphertext. Key is `TOKEN_ENC_KEY` env var (32 random bytes, base64). Never logged.
- `event_mappings` is the source of truth for "are these two paired?". On every webhook we look up by the foreign id; if absent, treat as create.
- `monday_etag` / `graph_etag` are the latest values we wrote ourselves OR observed externally — used for echo suppression.
- Soft delete via `deleted_at`. Reconciliation cron purges rows older than 30 days.

## 5. Sync Engine

### 5.1 Field mapping

| monday item                    | Outlook event                              |
|-------------------------------|--------------------------------------------|
| `name`                        | `subject`                                  |
| Date column (selected)         | `start.date`/`end.date` (all-day)          |
| Timeline column (`from`,`to`)  | `start.dateTime`,`end.dateTime`            |
| Latest item update             | `body.content` (HTML)                      |
| Item URL `…/pulses/<id>`       | Append link to `body`                      |
| Owner / Person column          | (v1: not synced) — `attendees` later       |

Outlook → monday writes the inverse: `subject` → `name`, `start/end` → the column the user selected for that board.

### 5.2 Five canonical events

| Trigger        | Behavior                                                                                                                                                                                              |
|----------------|------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| monday create  | 1) Look up mapping. 2) None → create event in `selected_calendar_id`. 3) Write `event_mappings` with both etags.                                                                                       |
| monday update  | 1) Find mapping. 2) If `monday_updated_at == mapping.monday_etag` → echo, skip. 3) Compare lastModified, monday newer → PATCH event. 4) Refresh both etags.                                            |
| monday delete  | 1) Find mapping → DELETE event. 2) Set `mapping.deleted_at`.                                                                                                                                          |
| outlook create | 1) Look up mapping. 2) None → create item in `default_board_id` / `default_group_id`, set the date/timeline column. 3) Write `event_mappings` (origin = 'outlook').                                    |
| outlook update / delete | 1) Find mapping. 2) Echo check via `graph_etag`. 3) LWW: whichever side has later `lastModified` wins; loser is overwritten. Refresh etags.                                                  |

### 5.3 Echo suppression (critical)

After every write to the opposite side:

1. Capture the new etag the API returned.
2. Persist it on `event_mappings` **before** the inevitable webhook from that side arrives.
3. When the webhook fires with the same etag → recognize as our own write, skip.

This is more reliable than timestamp comparison and survives clock skew.

### 5.4 Last-write-wins

Triggered only when **both** etags drift from what we last stored AND the change isn't an echo. Rule: compare `lastModified`, later wins, write to loser, refresh etags. Log `skip-conflict` if timestamps tie within 1 second and pick monday by tiebreak.

### 5.5 Rate limiting & retry

- monday: 5000 req / 5 min per account. Use a token-bucket per `monday_account_id`.
- Graph: per-app and per-mailbox throttling, follow `Retry-After`.
- All clients: exponential backoff with jitter; 5 retries on 429/5xx.
- Webhook bodies are deduped by `monday.event.changeId` / Graph `subscriptionId+resourceData.id+changeType` for 24h via in-memory LRU + DB stamp.

## 6. Authentication & Security

### 6.1 Three OAuth flows

1. **Office.js SSO (identity only)**: `Office.auth.getAccessToken({ allowSignInPrompt: true })` returns an ID token. Backend validates `aud` = our App ID and `iss` = Microsoft, upserts `users` row.
2. **Microsoft Graph (Calendars.ReadWrite + offline_access)**: Auth Code + PKCE in a popup window, callback to `/api/auth/microsoft/callback`. Tokens encrypted into `ms_accounts`. Subscribe `/me/events` on success.
3. **monday OAuth**: standard authorization code flow, popup → `/api/auth/monday/callback`. Tokens encrypted into `monday_accounts`. Webhook registered per board when user enables that board.

### 6.2 Security controls

- Token columns AES-256-GCM, key from env, never logged.
- Webhooks verified: monday `Authorization` header signing token; Graph `clientState` + `validationToken` echo on subscription handshake.
- OAuth `state` is encrypted JSON `{user_id, nonce}`; rejected on callback if missing or stale.
- Office SSO ID-token: validate `aud`, `iss`, `exp`, signature against Microsoft JWKS.
- Per-user data isolation: every Prisma query goes through `db.forUser(uid)` which appends `where userId = uid`. No bare client exported from `lib/db`.
- Rate-limit per user on `/api/sync/*` and webhooks.
- Outlook Add-in HTTPS enforced by Vercel by default.
- `manifest.xml` `<AppDomains>` and `<WebApplicationInfo>` lock to our domain.

### 6.3 Environment variables

```
MS_GRAPH_CLIENT_ID
MS_GRAPH_CLIENT_SECRET
MS_GRAPH_TENANT             # "common" for multi-tenant
MONDAY_CLIENT_ID
MONDAY_CLIENT_SECRET
MONDAY_SIGNING_SECRET       # monday webhook verification
TOKEN_ENC_KEY               # 32 bytes base64 — encrypts token columns
SUPABASE_URL
SUPABASE_SERVICE_KEY
DATABASE_URL                # Supabase Postgres connection string for Prisma
APP_BASE_URL                # https://outlook-monday-sync.vercel.app
```

## 7. Taskpane UI

Right-side panel (320–400px). Single Next.js route `/taskpane`, four tabs.

### 7.1 Tabs

- **Status (default)** — three connection cards (Microsoft Graph, monday, Webhooks) with green/yellow/red, "Reconnect" button when red, last 5 sync_log entries, "Sync now" button → `POST /api/sync/run`.
- **Boards** — list monday boards with toggles. Enabling expands a row: choose which Date or Timeline column drives the sync. Save registers the monday webhook; failures roll back the toggle.
- **Settings** — Outlook calendar dropdown (`/me/calendars`), monday default board + group dropdowns, two "Disconnect" buttons, plus a destructive "Wipe all mappings" that detaches all pairs without deleting source data.
- **Log** — `sync_log` reverse chronological with filters: errors-only, by mapping, time range.

### 7.2 Ribbon command

manifest registers a "Sync this event" button on the event ribbon. Clicking → `/api/sync/one` with the current event id. Handy for "I need this in monday right now" without waiting for the webhook hop.

### 7.3 i18n

`next-intl` with zh-CN and en-US. manifest has `DefaultLocale="en-US"` and `<Override Locale="zh-CN">` blocks; Outlook auto-picks user locale.

## 8. Project Structure

```
outlook_calendar_sync/
├── app/
│   ├── (marketing)/page.tsx
│   ├── taskpane/
│   │   ├── layout.tsx
│   │   ├── page.tsx
│   │   ├── _components/
│   │   │   ├── StatusTab.tsx
│   │   │   ├── BoardsTab.tsx
│   │   │   ├── SettingsTab.tsx
│   │   │   └── LogTab.tsx
│   │   └── _hooks/useOfficeAuth.ts
│   ├── commands/page.tsx
│   └── api/
│       ├── auth/
│       │   ├── microsoft/route.ts
│       │   ├── microsoft/callback/route.ts
│       │   ├── monday/route.ts
│       │   └── monday/callback/route.ts
│       ├── webhooks/
│       │   ├── monday/route.ts
│       │   └── graph/route.ts
│       ├── cron/refresh/route.ts
│       ├── sync/
│       │   ├── run/route.ts
│       │   └── one/route.ts
│       ├── boards/route.ts
│       ├── calendars/route.ts
│       └── status/route.ts
├── lib/
│   ├── monday/{client,boards,items,webhooks}.ts
│   ├── graph/{client,events,calendars,subscriptions}.ts
│   ├── sync/{engine,mapping,fields,conflict}.ts
│   ├── db/{client,users,mappings,boards}.ts
│   ├── crypto/token.ts
│   └── auth/{office-sso,session}.ts
├── prisma/
│   ├── schema.prisma
│   └── migrations/
├── public/icons/                 # 16/32/64/80/128 PNG
├── manifest/manifest.xml
├── vercel.ts
├── package.json
├── tsconfig.json
├── next.config.ts
└── README.md
```

`vercel.ts` minimum:

```ts
import { type VercelConfig } from '@vercel/config/v1';

export const config: VercelConfig = {
  framework: 'nextjs',
  crons: [{ path: '/api/cron/refresh', schedule: '7 * * * *' }],
};
```

## 9. Phased Implementation

| Phase | Scope | Acceptance |
|-------|-------|------------|
| **P1 Skeleton** | Next.js 16 project, manifest.xml, taskpane "hello" placeholder, sideload-able into Outlook web. | Sideload succeeds; "hello" rendered in Outlook on the web. |
| **P2 Auth** | Office SSO, Microsoft Graph OAuth, monday OAuth. Supabase + Prisma migrations. Token encryption. Settings tab shows "Connected" for both. | All three token records created; status endpoint returns green for Graph + monday. |
| **P3 monday→Outlook** | Boards tab. Pick board + column. Register monday webhook. Receive webhook → create/update/delete Outlook event. Persist event_mappings. | Edit a date in monday — within 30s the Outlook calendar reflects it. |
| **P4 Outlook→monday + LWW** | Graph subscription. Outlook→monday for create/update/delete. Etag-based echo suppression. last-write-wins. | Concurrent edits on both sides converge correctly without infinite loops. |
| **P5 Polish** | Log tab. Ribbon "Sync this event" button. Cron token refresh + Graph sub renewal. Error retry / dead-letter. zh-CN + en-US. | Acceptance script (in repo `docs/acceptance.md`) passes end-to-end. |

## 10. Testing Strategy

- `lib/sync/*` is pure functions — vitest unit tests with mocked monday/Graph clients.
- API routes — supertest + msw for outbound mocks, local Supabase Docker for DB.
- Manifest validated with `office-addin-manifest validate` in CI.
- Smoke test on every PR: deploy preview to Vercel, reload manifest in Outlook on the web, run a scripted edit on a fixture board.

## 11. Out of Scope for v1 (YAGNI)

- Recurring (master/instance) events
- Multi-attendee bidirectional sync
- File attachments
- AppSource compliance (privacy page, security questionnaire)
- Tenant-wide M365 admin deployment
- Locales beyond zh-CN / en-US

## 12. Open Risks / To Watch

- **Graph subscription renewal cadence**: Graph caps `events` subscriptions at 4230 minutes (~3 days). Cron must renew well before expiry. If a renewal fails, status card goes yellow then red and the user is prompted in the taskpane.
- **monday rate limits during back-fill**: First-time enable of a busy board can blow the 5000/5min budget. Implement back-pressure: enable triggers an asynchronous initial import that paginates with sleeps.
- **Office SSO availability**: Some tenants disable third-party SSO. Fall back path: if `Office.auth.getAccessToken` errors with 13003/13007, taskpane shows a "Sign in with Microsoft" button that does a full OAuth popup and stores a session cookie.
- **Vercel Queues is beta**: If unavailable / unstable, fall back to processing inline in the webhook function with `waitUntil` and idempotent retries.
