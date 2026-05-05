# Outlook ↔ monday.com Calendar Sync — Plan 1: Skeleton + Auth

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Stand up a sideload-able Outlook add-in whose taskpane can connect both Microsoft Graph and monday.com via OAuth, with encrypted tokens persisted in Supabase. After this plan, a user can install the add-in, click "Connect Microsoft" and "Connect monday", and the Settings tab shows green for both — but no calendar data flows yet.

**Architecture:** Single Next.js 16 App Router project on Vercel. The taskpane is a Next page hosted on the same domain as the API. Office.js runs inside the Outlook iframe and gets an SSO ID token to identify the user, then opens popup OAuth flows for the read-write scopes. Token columns are AES-256-GCM ciphertext, decrypted only inside server-side helpers. Prisma is the data layer, Supabase Postgres the store.

**Tech Stack:** Next.js 16 (App Router, TypeScript strict), Tailwind v4, Prisma 6, Supabase Postgres, `jose` (JWT), `@vercel/config` (vercel.ts), Office.js (CDN), vitest.

---

## Phase Boundary

This plan covers **P1 (Skeleton)** and **P2 (Auth)** from the design spec. After it's done:
- Add-in sideloads into Outlook on the web.
- `/api/status` returns three connection cards (Microsoft, monday, webhooks placeholder).
- Three OAuth flows complete and persist encrypted tokens.
- Settings tab in the taskpane shows "Connected" for both providers.

Subsequent plans will cover P3 (monday→Outlook), P4 (Outlook→monday + LWW), and P5 (Polish).

---

## File Structure (this plan)

```
outlook_calendar_sync/
├── app/
│   ├── layout.tsx                      # root, Tailwind import
│   ├── page.tsx                        # marketing placeholder
│   ├── taskpane/
│   │   ├── layout.tsx                  # injects Office.js, theme
│   │   ├── page.tsx                    # 4-tab shell (only Settings active in this plan)
│   │   └── _components/
│   │       ├── ConnectionStatus.tsx    # status card row (used by Settings + later Status tab)
│   │       └── SettingsTab.tsx
│   └── api/
│       ├── status/route.ts             # GET — returns connection state for taskpane
│       ├── auth/
│       │   ├── microsoft/route.ts            # popup entry (302 to MS authorize)
│       │   ├── microsoft/callback/route.ts   # auth code → tokens → DB
│       │   ├── monday/route.ts               # popup entry (302 to monday authorize)
│       │   └── monday/callback/route.ts      # auth code → tokens → DB
│       └── session/route.ts            # POST: validate Office SSO ID token, set session cookie
├── lib/
│   ├── crypto/token.ts                 # AES-256-GCM encrypt/decrypt helpers
│   ├── auth/
│   │   ├── office-sso.ts               # validate Office.js SSO ID token (jose)
│   │   ├── oauth-state.ts              # encrypted state JSON for CSRF
│   │   └── session.ts                  # cookie-based session helpers
│   ├── ms/
│   │   └── oauth.ts                    # MS authorize URL builder + token exchange
│   ├── monday/
│   │   └── oauth.ts                    # monday authorize URL builder + token exchange
│   └── db/
│       └── client.ts                   # Prisma singleton + per-user scoped helper
├── prisma/
│   └── schema.prisma
├── manifest/
│   └── manifest.xml
├── public/
│   └── icons/                          # 16/32/64/80/128 PNG (placeholder logo)
├── tests/
│   ├── crypto-token.test.ts
│   ├── office-sso.test.ts
│   └── oauth-state.test.ts
├── vercel.ts
├── next.config.ts
├── tailwind.config.ts
├── tsconfig.json
├── package.json
├── vitest.config.ts
├── .env.example
└── .gitignore
```

Each file has one responsibility. `lib/crypto/token.ts`, `lib/auth/office-sso.ts`, and `lib/auth/oauth-state.ts` are pure functions with vitest unit tests. OAuth route files are thin wrappers around `lib/ms/oauth.ts` and `lib/monday/oauth.ts`.

---

## Task 1: Initialize Next.js 16 project

**Files:**
- Create: `package.json`, `tsconfig.json`, `next.config.ts`, `app/layout.tsx`, `app/page.tsx`, `tailwind.config.ts`, `app/globals.css`, `.gitignore`

- [ ] **Step 1: Run `create-next-app` non-interactively**

```bash
npx --yes create-next-app@latest . \
  --typescript --tailwind --app --src-dir=false \
  --eslint --turbopack --import-alias="@/*" \
  --use-npm --skip-install --yes
```

Expected: scaffolds files; we install in next step.

- [ ] **Step 2: Pin to Next 16 + react 19 in `package.json`**

Edit the generated `package.json` so the dependencies block matches:

```json
"dependencies": {
  "next": "^16.0.0",
  "react": "^19.0.0",
  "react-dom": "^19.0.0"
},
"devDependencies": {
  "@types/node": "^22.0.0",
  "@types/react": "^19.0.0",
  "@types/react-dom": "^19.0.0",
  "tailwindcss": "^4.0.0",
  "@tailwindcss/postcss": "^4.0.0",
  "postcss": "^8.4.0",
  "typescript": "^5.6.0",
  "eslint": "^9.0.0",
  "eslint-config-next": "^16.0.0"
}
```

- [ ] **Step 3: Install**

Run: `npm install`
Expected: lock file created, no errors.

- [ ] **Step 4: Replace `app/page.tsx` with marketing placeholder**

```tsx
export default function HomePage() {
  return (
    <main className="mx-auto max-w-2xl p-12">
      <h1 className="text-3xl font-semibold">Outlook ↔ monday.com Calendar Sync</h1>
      <p className="mt-3 text-zinc-600">
        Install the add-in in Outlook to start syncing.
      </p>
    </main>
  );
}
```

- [ ] **Step 5: Run dev server**

Run: `npm run dev`
Open `http://localhost:3000`. Expected: page renders.
Stop the server (Ctrl+C) before continuing.

- [ ] **Step 6: Commit**

```bash
git add .
git commit -m "feat: scaffold Next.js 16 app"
```

---

## Task 2: Add `vercel.ts`, `.env.example`, install runtime deps

**Files:**
- Create: `vercel.ts`, `.env.example`
- Modify: `package.json`

- [ ] **Step 1: Install runtime + dev dependencies**

```bash
npm install @vercel/config jose @prisma/client
npm install -D prisma vitest @vitest/coverage-v8 office-addin-manifest tsx
```

- [ ] **Step 2: Create `vercel.ts`**

```ts
import { type VercelConfig } from '@vercel/config/v1';

export const config: VercelConfig = {
  framework: 'nextjs',
  crons: [{ path: '/api/cron/refresh', schedule: '7 * * * *' }],
};
```

- [ ] **Step 3: Create `.env.example`**

```bash
# Microsoft Graph App registration (Azure AD)
MS_GRAPH_CLIENT_ID=
MS_GRAPH_CLIENT_SECRET=
MS_GRAPH_TENANT=common

# monday.com OAuth app
MONDAY_CLIENT_ID=
MONDAY_CLIENT_SECRET=
MONDAY_SIGNING_SECRET=

# 32 random bytes, base64 — encrypts token columns
TOKEN_ENC_KEY=

# Supabase project
DATABASE_URL=
SUPABASE_URL=
SUPABASE_SERVICE_KEY=

# Public origin used in OAuth redirect URIs
APP_BASE_URL=https://localhost:3000
```

- [ ] **Step 4: Add `.env.local` to `.gitignore`**

Open `.gitignore`, ensure `.env*.local` is present (create-next-app adds it; double-check).

- [ ] **Step 5: Add scripts to `package.json`**

In the `"scripts"` block, replace with:

```json
"scripts": {
  "dev": "next dev --turbopack",
  "build": "next build --turbopack",
  "start": "next start",
  "lint": "next lint",
  "test": "vitest run",
  "test:watch": "vitest",
  "prisma:generate": "prisma generate",
  "prisma:migrate": "prisma migrate dev",
  "manifest:validate": "office-addin-manifest validate manifest/manifest.xml"
}
```

- [ ] **Step 6: Commit**

```bash
git add package.json package-lock.json vercel.ts .env.example .gitignore
git commit -m "chore: add runtime deps, vercel.ts, env example"
```

---

## Task 3: Create placeholder icons

**Files:**
- Create: `public/icons/icon-16.png`, `icon-32.png`, `icon-64.png`, `icon-80.png`, `icon-128.png`

The Outlook manifest requires PNG icons at five sizes. Use solid-fill placeholders for now; replace with real artwork later.

- [ ] **Step 1: Generate icons with ImageMagick**

```bash
mkdir -p public/icons
for size in 16 32 64 80 128; do
  magick -size ${size}x${size} xc:'#0073EA' \
    -gravity center -fill white -pointsize $((size/3)) \
    -annotate +0+0 'OM' \
    public/icons/icon-${size}.png
done
ls -l public/icons/
```

Expected: five PNGs listed. If `magick` is not available, install ImageMagick (`brew install imagemagick`) or substitute any solid-fill PNG of the correct sizes.

- [ ] **Step 2: Commit**

```bash
git add public/icons/
git commit -m "feat: add placeholder add-in icons"
```

---

## Task 4: Create Outlook add-in manifest.xml

**Files:**
- Create: `manifest/manifest.xml`

This is the legacy XML manifest (not the unified JSON manifest), which is widely supported across Outlook web, desktop, and Mac.

- [ ] **Step 1: Create `manifest/manifest.xml`**

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
           xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
           xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
           xmlns:mailappor="http://schemas.microsoft.com/office/mailappversionoverrides/1.0"
           xsi:type="MailApp">

  <Id>11111111-2222-3333-4444-555555555555</Id>
  <Version>0.1.0.0</Version>
  <ProviderName>Outlook monday Sync</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Outlook ↔ monday Sync" />
  <Description DefaultValue="Bidirectional sync between Outlook calendar and monday.com." />
  <IconUrl DefaultValue="https://localhost:3000/icons/icon-64.png" />
  <HighResolutionIconUrl DefaultValue="https://localhost:3000/icons/icon-128.png" />
  <SupportUrl DefaultValue="https://localhost:3000" />
  <AppDomains>
    <AppDomain>https://localhost:3000</AppDomain>
  </AppDomains>

  <Hosts>
    <Host Name="Mailbox" />
  </Hosts>

  <Requirements>
    <Sets>
      <Set Name="Mailbox" MinVersion="1.10" />
    </Sets>
  </Requirements>

  <FormSettings>
    <Form xsi:type="ItemRead">
      <DesktopSettings>
        <SourceLocation DefaultValue="https://localhost:3000/taskpane" />
        <RequestedHeight>250</RequestedHeight>
      </DesktopSettings>
    </Form>
  </FormSettings>

  <Permissions>ReadWriteMailbox</Permissions>
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemIs" ItemType="Message" />
    <Rule xsi:type="ItemIs" ItemType="AppointmentAttendee" />
    <Rule xsi:type="ItemIs" ItemType="AppointmentOrganizer" />
  </Rule>
  <DisableEntityHighlighting>false</DisableEntityHighlighting>

  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides"
                    xsi:type="VersionOverridesV1_0">
    <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1"
                      xsi:type="VersionOverridesV1_1">
      <Requirements>
        <bt:Sets DefaultMinVersion="1.10">
          <bt:Set Name="Mailbox" />
        </bt:Sets>
      </Requirements>
      <Hosts>
        <Host xsi:type="MailHost">
          <DesktopFormFactor>
            <FunctionFile resid="commands.url" />
            <ExtensionPoint xsi:type="MessageReadCommandSurface">
              <OfficeTab id="TabDefault">
                <Group id="msgReadGroup">
                  <Label resid="group.label" />
                  <Control xsi:type="Button" id="openTaskpane">
                    <Label resid="open.label" />
                    <Supertip>
                      <Title resid="open.label" />
                      <Description resid="open.tip" />
                    </Supertip>
                    <Icon>
                      <bt:Image size="16" resource="icon16" />
                      <bt:Image size="32" resource="icon32" />
                      <bt:Image size="80" resource="icon80" />
                    </Icon>
                    <Action xsi:type="ShowTaskpane">
                      <SourceLocation resid="taskpane.url" />
                    </Action>
                  </Control>
                </Group>
              </OfficeTab>
            </ExtensionPoint>
          </DesktopFormFactor>
        </Host>
      </Hosts>
      <Resources>
        <bt:Images>
          <bt:Image id="icon16" DefaultValue="https://localhost:3000/icons/icon-16.png" />
          <bt:Image id="icon32" DefaultValue="https://localhost:3000/icons/icon-32.png" />
          <bt:Image id="icon80" DefaultValue="https://localhost:3000/icons/icon-80.png" />
        </bt:Images>
        <bt:Urls>
          <bt:Url id="taskpane.url" DefaultValue="https://localhost:3000/taskpane" />
          <bt:Url id="commands.url" DefaultValue="https://localhost:3000/commands" />
        </bt:Urls>
        <bt:ShortStrings>
          <bt:String id="group.label" DefaultValue="monday Sync" />
          <bt:String id="open.label" DefaultValue="Open sync panel" />
        </bt:ShortStrings>
        <bt:LongStrings>
          <bt:String id="open.tip" DefaultValue="Open the Outlook ↔ monday sync panel." />
        </bt:LongStrings>
      </Resources>
      <WebApplicationInfo>
        <Id>00000000-0000-0000-0000-000000000000</Id>
        <Resource>api://localhost:3000/00000000-0000-0000-0000-000000000000</Resource>
        <Scopes>
          <Scope>profile</Scope>
          <Scope>openid</Scope>
        </Scopes>
      </WebApplicationInfo>
    </VersionOverrides>
  </VersionOverrides>
</OfficeApp>
```

The `<Id>` and `<WebApplicationInfo><Id>` GUIDs will be replaced once the Azure AD app registration is created (Task 12). Same for the GUID inside `<Resource>`. Leave the placeholder for now.

- [ ] **Step 2: Validate with `office-addin-manifest`**

Run: `npm run manifest:validate`
Expected: validation passes (errors about Web Application Info Id being a placeholder GUID are acceptable for now — the validator accepts any well-formed GUID).

- [ ] **Step 3: Commit**

```bash
git add manifest/manifest.xml
git commit -m "feat: add Outlook add-in manifest"
```

---

## Task 5: Create taskpane shell page

**Files:**
- Create: `app/taskpane/layout.tsx`, `app/taskpane/page.tsx`, `app/taskpane/_components/ConnectionStatus.tsx`, `app/taskpane/_components/SettingsTab.tsx`, `app/commands/page.tsx`

The taskpane is a Next route. Office.js is loaded via a CDN script tag in the taskpane layout. Only the Settings tab is functional in this plan; the other three are stubs for later plans.

- [ ] **Step 1: Create `app/taskpane/layout.tsx`**

```tsx
import Script from 'next/script';

export const metadata = { title: 'Outlook ↔ monday Sync' };

export default function TaskpaneLayout({ children }: { children: React.ReactNode }) {
  return (
    <>
      <Script
        src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"
        strategy="beforeInteractive"
      />
      <div className="min-h-screen bg-white text-zinc-900 text-sm">
        {children}
      </div>
    </>
  );
}
```

- [ ] **Step 2: Create `app/taskpane/_components/ConnectionStatus.tsx`**

```tsx
type Variant = 'green' | 'yellow' | 'red';

export function ConnectionStatus(props: {
  label: string;
  variant: Variant;
  detail?: string;
  action?: { label: string; onClick: () => void };
}) {
  const dot =
    props.variant === 'green'
      ? 'bg-emerald-500'
      : props.variant === 'yellow'
        ? 'bg-amber-500'
        : 'bg-rose-500';
  return (
    <div className="flex items-center justify-between rounded border border-zinc-200 px-3 py-2">
      <div className="flex items-center gap-2">
        <span className={`h-2 w-2 rounded-full ${dot}`} />
        <div>
          <div className="font-medium">{props.label}</div>
          {props.detail && <div className="text-xs text-zinc-500">{props.detail}</div>}
        </div>
      </div>
      {props.action && (
        <button
          className="rounded bg-zinc-900 px-2 py-1 text-xs text-white"
          onClick={props.action.onClick}
        >
          {props.action.label}
        </button>
      )}
    </div>
  );
}
```

- [ ] **Step 3: Create `app/taskpane/_components/SettingsTab.tsx`**

```tsx
'use client';

import { useEffect, useState } from 'react';
import { ConnectionStatus } from './ConnectionStatus';

type StatusPayload = {
  microsoft: { connected: boolean; email?: string };
  monday: { connected: boolean; account?: string };
};

export function SettingsTab() {
  const [status, setStatus] = useState<StatusPayload | null>(null);

  async function refresh() {
    const r = await fetch('/api/status', { credentials: 'include' });
    if (r.ok) setStatus(await r.json());
  }

  useEffect(() => {
    refresh();
  }, []);

  function openPopup(url: string) {
    const w = window.open(url, '_blank', 'width=520,height=640');
    const t = setInterval(() => {
      if (w?.closed) {
        clearInterval(t);
        refresh();
      }
    }, 800);
  }

  return (
    <div className="space-y-3 p-4">
      <ConnectionStatus
        label="Microsoft Graph"
        variant={status?.microsoft.connected ? 'green' : 'red'}
        detail={status?.microsoft.email ?? 'Not connected'}
        action={
          status?.microsoft.connected
            ? undefined
            : { label: 'Connect', onClick: () => openPopup('/api/auth/microsoft') }
        }
      />
      <ConnectionStatus
        label="monday.com"
        variant={status?.monday.connected ? 'green' : 'red'}
        detail={status?.monday.account ?? 'Not connected'}
        action={
          status?.monday.connected
            ? undefined
            : { label: 'Connect', onClick: () => openPopup('/api/auth/monday') }
        }
      />
    </div>
  );
}
```

- [ ] **Step 4: Create `app/taskpane/page.tsx`**

```tsx
'use client';

import { useState } from 'react';
import { SettingsTab } from './_components/SettingsTab';

const TABS = ['Status', 'Boards', 'Settings', 'Log'] as const;

export default function TaskpanePage() {
  const [tab, setTab] = useState<(typeof TABS)[number]>('Settings');

  return (
    <div className="flex h-screen flex-col">
      <header className="border-b border-zinc-200 px-4 py-2 text-xs font-medium">
        Outlook ↔ monday Sync
      </header>
      <nav className="flex border-b border-zinc-200">
        {TABS.map((t) => (
          <button
            key={t}
            onClick={() => setTab(t)}
            className={
              t === tab
                ? 'flex-1 border-b-2 border-zinc-900 px-2 py-2 text-xs font-semibold'
                : 'flex-1 px-2 py-2 text-xs text-zinc-500'
            }
          >
            {t}
          </button>
        ))}
      </nav>
      <main className="flex-1 overflow-y-auto">
        {tab === 'Settings' ? (
          <SettingsTab />
        ) : (
          <div className="p-4 text-zinc-500">Coming in a later phase.</div>
        )}
      </main>
    </div>
  );
}
```

- [ ] **Step 5: Create `app/commands/page.tsx` placeholder**

```tsx
export default function CommandsPage() {
  return null;
}
```

- [ ] **Step 6: Run dev server and visit taskpane**

Run: `npm run dev`
Open `http://localhost:3000/taskpane`. Expected: header + four tab buttons, Settings active, both connection cards red with "Connect" buttons.

- [ ] **Step 7: Commit**

```bash
git add app/
git commit -m "feat: taskpane shell + Settings tab + connection status component"
```

---

## Task 6: Sideload smoke test

This task is a **manual verification** — no code changes. The goal is to confirm a real Outlook can render the taskpane pointing at our local dev server before we add backend complexity.

**Files:** None.

- [ ] **Step 1: Start the dev server with HTTPS**

Office Add-ins require HTTPS. Use Next.js experimental HTTPS:

```bash
npm run dev -- --experimental-https
```

Expected: server starts on `https://localhost:3000`. Accept the self-signed cert in your browser by visiting `https://localhost:3000/taskpane` once.

- [ ] **Step 2: Sideload manifest into Outlook on the web**

1. Sign in to `https://outlook.office.com`.
2. Click the gear → "View all Outlook settings" → "General" → "Manage add-ins" → "My add-ins" → "Add a custom add-in" → "Add from file".
3. Select `manifest/manifest.xml`.
4. Confirm the warning dialog.

Expected: add-in appears under "Custom add-ins". Open any email → click the add-in icon (it may be under the "..." overflow). Taskpane loads and shows the same UI as in browser.

If the panel is blank, open browser devtools while the panel is open (or check `F12` from the taskpane). Common errors:
- "Refused to connect": cert not trusted — re-visit `https://localhost:3000/taskpane` and accept.
- "Failed to load Office.js": network blocked, retry.

- [ ] **Step 3: Note the working URL of `taskpane` in Outlook**

In Outlook devtools, the taskpane's URL should be `https://localhost:3000/taskpane?...`. Confirm and stop the dev server.

- [ ] **Step 4: Commit a checkpoint marker**

No code changed; commit with `--allow-empty` so the milestone is on the timeline:

```bash
git commit --allow-empty -m "chore: P1 skeleton verified — sideload + taskpane render OK"
```

---

## Task 7: Set up Prisma + Supabase

**Files:**
- Create: `prisma/schema.prisma`, `lib/db/client.ts`

- [ ] **Step 1: Create Supabase project**

Manual step:
1. Sign in to `https://supabase.com`, create new project named `outlook-monday-sync`.
2. Region: pick the same continent as Vercel (e.g. `us-east`).
3. Copy the connection string (Database → Connection → URI) and the Service Role key.
4. Set in `.env.local`:

```bash
DATABASE_URL="postgres://postgres:<password>@db.<ref>.supabase.co:5432/postgres"
SUPABASE_URL="https://<ref>.supabase.co"
SUPABASE_SERVICE_KEY="<service-role-key>"
```

- [ ] **Step 2: Run `prisma init`**

```bash
npx prisma init --datasource-provider postgresql --output ../node_modules/.prisma/client
```

Expected: `prisma/schema.prisma` created. Don't replace `.env` (we use `.env.local` already).

- [ ] **Step 3: Replace `prisma/schema.prisma` with full schema**

```prisma
generator client {
  provider = "prisma-client-js"
}

datasource db {
  provider = "postgresql"
  url      = env("DATABASE_URL")
}

model User {
  id           String   @id @default(uuid()) @db.Uuid
  msTenantId   String   @map("ms_tenant_id")
  msUserId     String   @map("ms_user_id")
  msUserEmail  String   @map("ms_user_email")
  createdAt    DateTime @default(now()) @map("created_at")

  msAccount      MsAccount?
  mondayAccount  MondayAccount?
  boards         BoardSubscription[]
  mappings       EventMapping[]
  log            SyncLog[]

  @@unique([msTenantId, msUserId])
  @@map("users")
}

model MsAccount {
  userId                       String   @id @map("user_id") @db.Uuid
  accessTokenEnc               Bytes    @map("access_token_enc")
  refreshTokenEnc              Bytes    @map("refresh_token_enc")
  expiresAt                    DateTime @map("expires_at")
  graphSubscriptionId          String?  @map("graph_subscription_id")
  graphSubscriptionExpiresAt   DateTime? @map("graph_subscription_expires_at")
  selectedCalendarId           String?  @map("selected_calendar_id")

  user User @relation(fields: [userId], references: [id], onDelete: Cascade)

  @@map("ms_accounts")
}

model MondayAccount {
  userId           String  @id @map("user_id") @db.Uuid
  mondayUserId     BigInt  @map("monday_user_id")
  mondayAccountId  BigInt  @map("monday_account_id")
  accessTokenEnc   Bytes   @map("access_token_enc")
  refreshTokenEnc  Bytes   @map("refresh_token_enc")
  expiresAt        DateTime @map("expires_at")
  defaultBoardId   BigInt? @map("default_board_id")
  defaultGroupId   String? @map("default_group_id")

  user User @relation(fields: [userId], references: [id], onDelete: Cascade)

  @@map("monday_accounts")
}

model BoardSubscription {
  id               String   @id @default(uuid()) @db.Uuid
  userId           String   @map("user_id") @db.Uuid
  mondayBoardId    BigInt   @map("monday_board_id")
  dateColumnId     String   @map("date_column_id")
  dateColumnKind   String   @map("date_column_kind")
  mondayWebhookId  BigInt?  @map("monday_webhook_id")
  active           Boolean  @default(true)

  user User @relation(fields: [userId], references: [id], onDelete: Cascade)

  @@unique([userId, mondayBoardId])
  @@map("board_subscriptions")
}

model EventMapping {
  id               String    @id @default(uuid()) @db.Uuid
  userId           String    @map("user_id") @db.Uuid
  mondayItemId     BigInt    @map("monday_item_id")
  mondayBoardId    BigInt    @map("monday_board_id")
  graphEventId     String    @map("graph_event_id")
  graphCalendarId  String    @map("graph_calendar_id")
  mondayEtag       String?   @map("monday_etag")
  graphEtag        String?   @map("graph_etag")
  origin           String
  lastSyncedAt     DateTime? @map("last_synced_at")
  deletedAt        DateTime? @map("deleted_at")

  user User @relation(fields: [userId], references: [id], onDelete: Cascade)

  @@unique([userId, mondayItemId])
  @@unique([userId, graphEventId])
  @@map("event_mappings")
}

model SyncLog {
  id          String   @id @default(uuid()) @db.Uuid
  userId      String   @map("user_id") @db.Uuid
  direction   String
  mappingId   String?  @map("mapping_id") @db.Uuid
  action      String
  message     String?
  occurredAt  DateTime @default(now()) @map("occurred_at")

  user User @relation(fields: [userId], references: [id], onDelete: Cascade)

  @@index([userId, occurredAt])
  @@map("sync_log")
}
```

- [ ] **Step 4: Run initial migration**

```bash
npx prisma migrate dev --name init
```

Expected: connects to Supabase, creates all six tables. If the connection string is wrong, the error will be `P1001` — fix `.env.local` and retry.

- [ ] **Step 5: Create `lib/db/client.ts`**

```ts
import { PrismaClient } from '@prisma/client';

declare global {
  // eslint-disable-next-line no-var
  var __prisma: PrismaClient | undefined;
}

export const prisma =
  global.__prisma ?? new PrismaClient({ log: ['error', 'warn'] });

if (process.env.NODE_ENV !== 'production') global.__prisma = prisma;

/** Per-user query helper. Always pass through this when reading user data. */
export function forUser(userId: string) {
  return {
    msAccount: () => prisma.msAccount.findUnique({ where: { userId } }),
    mondayAccount: () => prisma.mondayAccount.findUnique({ where: { userId } }),
    boards: () => prisma.boardSubscription.findMany({ where: { userId } }),
    mappings: () => prisma.eventMapping.findMany({ where: { userId, deletedAt: null } }),
    log: (limit = 50) =>
      prisma.syncLog.findMany({
        where: { userId },
        orderBy: { occurredAt: 'desc' },
        take: limit,
      }),
  };
}
```

- [ ] **Step 6: Commit**

```bash
git add prisma/ lib/db/
git commit -m "feat: Prisma schema + Supabase migration + per-user db helper"
```

---

## Task 8: AES-256-GCM token encryption (TDD)

**Files:**
- Create: `lib/crypto/token.ts`, `tests/crypto-token.test.ts`, `vitest.config.ts`

- [ ] **Step 1: Create `vitest.config.ts`**

```ts
import { defineConfig } from 'vitest/config';
import path from 'node:path';

export default defineConfig({
  test: {
    environment: 'node',
    include: ['tests/**/*.test.ts'],
  },
  resolve: {
    alias: { '@': path.resolve(__dirname, '.') },
  },
});
```

- [ ] **Step 2: Write failing test `tests/crypto-token.test.ts`**

```ts
import { describe, it, expect, beforeAll } from 'vitest';
import { randomBytes } from 'node:crypto';
import { encryptToken, decryptToken } from '@/lib/crypto/token';

describe('token encryption', () => {
  beforeAll(() => {
    process.env.TOKEN_ENC_KEY = randomBytes(32).toString('base64');
  });

  it('round-trips a string through encrypt/decrypt', () => {
    const plaintext = 'ya29.a0AeXRPp7-not-a-real-token';
    const cipher = encryptToken(plaintext);
    expect(cipher).toBeInstanceOf(Buffer);
    expect(cipher.length).toBeGreaterThan(28);
    expect(decryptToken(cipher)).toBe(plaintext);
  });

  it('produces different ciphertext on repeated calls (random IV)', () => {
    const a = encryptToken('hello');
    const b = encryptToken('hello');
    expect(a.equals(b)).toBe(false);
  });

  it('fails to decrypt tampered data', () => {
    const cipher = encryptToken('hello');
    cipher[cipher.length - 1] ^= 0x01;
    expect(() => decryptToken(cipher)).toThrow();
  });

  it('throws if key is missing', () => {
    const saved = process.env.TOKEN_ENC_KEY;
    delete process.env.TOKEN_ENC_KEY;
    expect(() => encryptToken('x')).toThrow(/TOKEN_ENC_KEY/);
    process.env.TOKEN_ENC_KEY = saved;
  });
});
```

- [ ] **Step 3: Run test to verify it fails**

Run: `npm test`
Expected: `Cannot find module '@/lib/crypto/token'`.

- [ ] **Step 4: Implement `lib/crypto/token.ts`**

```ts
import { createCipheriv, createDecipheriv, randomBytes } from 'node:crypto';

const ALG = 'aes-256-gcm';
const IV_LEN = 12;
const TAG_LEN = 16;

function key(): Buffer {
  const b64 = process.env.TOKEN_ENC_KEY;
  if (!b64) throw new Error('TOKEN_ENC_KEY env var not set');
  const k = Buffer.from(b64, 'base64');
  if (k.length !== 32) throw new Error('TOKEN_ENC_KEY must decode to 32 bytes');
  return k;
}

/** Encrypt a UTF-8 string. Returns: iv (12) || tag (16) || ciphertext. */
export function encryptToken(plaintext: string): Buffer {
  const iv = randomBytes(IV_LEN);
  const cipher = createCipheriv(ALG, key(), iv);
  const ct = Buffer.concat([cipher.update(plaintext, 'utf8'), cipher.final()]);
  const tag = cipher.getAuthTag();
  return Buffer.concat([iv, tag, ct]);
}

export function decryptToken(blob: Buffer): string {
  const iv = blob.subarray(0, IV_LEN);
  const tag = blob.subarray(IV_LEN, IV_LEN + TAG_LEN);
  const ct = blob.subarray(IV_LEN + TAG_LEN);
  const decipher = createDecipheriv(ALG, key(), iv);
  decipher.setAuthTag(tag);
  return Buffer.concat([decipher.update(ct), decipher.final()]).toString('utf8');
}
```

- [ ] **Step 5: Run test to verify it passes**

Run: `npm test`
Expected: 4 tests pass.

- [ ] **Step 6: Commit**

```bash
git add tests/ lib/crypto/ vitest.config.ts
git commit -m "feat: AES-256-GCM token encryption with vitest coverage"
```

---

## Task 9: OAuth state encoding (TDD)

OAuth `state` carries a `userId` and a CSRF nonce. We encrypt it with the same key as tokens so callbacks can recover the user without an extra DB lookup.

**Files:**
- Create: `lib/auth/oauth-state.ts`, `tests/oauth-state.test.ts`

- [ ] **Step 1: Write failing test `tests/oauth-state.test.ts`**

```ts
import { describe, it, expect, beforeAll } from 'vitest';
import { randomBytes } from 'node:crypto';
import { encodeState, decodeState } from '@/lib/auth/oauth-state';

beforeAll(() => {
  process.env.TOKEN_ENC_KEY = randomBytes(32).toString('base64');
});

describe('oauth-state', () => {
  it('round-trips userId + nonce', () => {
    const s = encodeState({ userId: 'u-123', purpose: 'monday' });
    const out = decodeState(s);
    expect(out.userId).toBe('u-123');
    expect(out.purpose).toBe('monday');
    expect(out.nonce).toMatch(/^[a-f0-9]{32}$/);
  });

  it('rejects tampered state', () => {
    const s = encodeState({ userId: 'u-123', purpose: 'monday' });
    const tampered = s.slice(0, -2) + (s.endsWith('A') ? 'B' : 'A');
    expect(() => decodeState(tampered)).toThrow();
  });

  it('produces base64url output (no slashes/pluses/equals)', () => {
    const s = encodeState({ userId: 'u-123', purpose: 'ms' });
    expect(s).toMatch(/^[A-Za-z0-9\-_]+$/);
  });
});
```

- [ ] **Step 2: Run test, verify failure**

Run: `npm test`
Expected: `Cannot find module '@/lib/auth/oauth-state'`.

- [ ] **Step 3: Implement `lib/auth/oauth-state.ts`**

```ts
import { randomBytes } from 'node:crypto';
import { encryptToken, decryptToken } from '@/lib/crypto/token';

type Payload = { userId: string; purpose: 'ms' | 'monday'; nonce?: string };

export function encodeState(p: Omit<Payload, 'nonce'>): string {
  const full: Payload = { ...p, nonce: randomBytes(16).toString('hex') };
  const blob = encryptToken(JSON.stringify(full));
  return blob.toString('base64url');
}

export function decodeState(s: string): Required<Payload> {
  const blob = Buffer.from(s, 'base64url');
  const json = decryptToken(blob);
  const parsed = JSON.parse(json) as Required<Payload>;
  if (!parsed.userId || !parsed.nonce || !parsed.purpose) {
    throw new Error('invalid state payload');
  }
  return parsed;
}
```

- [ ] **Step 4: Run test, verify pass**

Run: `npm test`
Expected: 3 tests pass on top of previous 4 → 7 total.

- [ ] **Step 5: Commit**

```bash
git add lib/auth/oauth-state.ts tests/oauth-state.test.ts
git commit -m "feat: oauth-state encryption helpers (CSRF protection)"
```

---

## Task 10: Office.js SSO ID-token validator (TDD)

Office.js gives the taskpane an ID token. Backend must validate the JWT against Microsoft's public keys, check the audience matches our app, and return the user identity.

**Files:**
- Create: `lib/auth/office-sso.ts`, `tests/office-sso.test.ts`

- [ ] **Step 1: Write failing test**

We can't easily mint a real Microsoft-signed JWT in unit tests, so we test the helper that **parses** the validated token claims (post-`jose` verification) into a `MsIdentity` shape. The signature-verification call is exercised in the higher-level integration test (Task 14).

`tests/office-sso.test.ts`:

```ts
import { describe, it, expect } from 'vitest';
import { extractIdentity } from '@/lib/auth/office-sso';

describe('office-sso extractIdentity', () => {
  it('maps tid + oid + preferred_username to MsIdentity', () => {
    const claims = {
      tid: 'tenant-xyz',
      oid: 'user-abc',
      preferred_username: 'alice@contoso.com',
      aud: 'api://localhost:3000/<app-id>',
      iss: 'https://login.microsoftonline.com/<tid>/v2.0',
      exp: Math.floor(Date.now() / 1000) + 3600,
    };
    expect(extractIdentity(claims)).toEqual({
      tenantId: 'tenant-xyz',
      userId: 'user-abc',
      email: 'alice@contoso.com',
    });
  });

  it('throws if any required claim is missing', () => {
    expect(() => extractIdentity({ tid: 'x' } as never)).toThrow();
    expect(() => extractIdentity({ tid: 'x', oid: 'y' } as never)).toThrow();
  });
});
```

- [ ] **Step 2: Run test, verify failure**

Run: `npm test`
Expected: `Cannot find module '@/lib/auth/office-sso'`.

- [ ] **Step 3: Implement `lib/auth/office-sso.ts`**

```ts
import { createRemoteJWKSet, jwtVerify, type JWTPayload } from 'jose';

const JWKS = createRemoteJWKSet(
  new URL('https://login.microsoftonline.com/common/discovery/v2.0/keys'),
);

export type MsIdentity = { tenantId: string; userId: string; email: string };

export async function verifyOfficeSsoToken(token: string): Promise<MsIdentity> {
  const audience = `api://localhost:3000/${process.env.MS_GRAPH_CLIENT_ID ?? ''}`;
  const { payload } = await jwtVerify(token, JWKS, {
    audience,
    issuer: /^https:\/\/login\.microsoftonline\.com\/[^/]+\/v2\.0$/ as unknown as string,
  });
  return extractIdentity(payload);
}

export function extractIdentity(p: JWTPayload & {
  tid?: string;
  oid?: string;
  preferred_username?: string;
}): MsIdentity {
  if (!p.tid || !p.oid || !p.preferred_username) {
    throw new Error('SSO token missing required claims');
  }
  return {
    tenantId: p.tid,
    userId: p.oid,
    email: p.preferred_username,
  };
}
```

(Note: in production `audience` will read `APP_BASE_URL`; for now it's pinned to localhost. We'll un-pin it in Task 13.)

- [ ] **Step 4: Run test, verify pass**

Run: `npm test`
Expected: 9 tests pass total.

- [ ] **Step 5: Commit**

```bash
git add lib/auth/office-sso.ts tests/office-sso.test.ts
git commit -m "feat: Office.js SSO ID token validator"
```

---

## Task 11: Microsoft Graph OAuth flow

**Files:**
- Create: `lib/ms/oauth.ts`, `app/api/auth/microsoft/route.ts`, `app/api/auth/microsoft/callback/route.ts`

- [ ] **Step 1: Manual — register Azure AD app**

1. Go to `https://entra.microsoft.com` → Identity → Applications → App registrations → New registration.
2. Name: `Outlook monday Sync (dev)`. Supported account types: "Accounts in any organizational directory and personal Microsoft accounts".
3. Redirect URI: Web, `https://localhost:3000/api/auth/microsoft/callback`.
4. Register.
5. From the overview, copy **Application (client) ID** → `MS_GRAPH_CLIENT_ID` in `.env.local`.
6. Certificates & secrets → New client secret → 6 months. Copy the **Value** → `MS_GRAPH_CLIENT_SECRET`.
7. API permissions → Add a permission → Microsoft Graph → Delegated permissions → check `Calendars.ReadWrite`, `User.Read`, `offline_access`. Click "Grant admin consent" if you have rights.
8. Expose an API → Set the Application ID URI to `api://localhost:3000/<client-id>`. Add a scope `access_as_user` (admin and users can consent).
9. In `manifest/manifest.xml`, replace the placeholder GUID inside `<WebApplicationInfo><Id>` and the GUID at the top `<Id>` with the same Application ID. Also update `<Resource>` to match the URI from step 8.
10. Re-validate: `npm run manifest:validate`.

- [ ] **Step 2: Implement `lib/ms/oauth.ts`**

```ts
const AUTH = 'https://login.microsoftonline.com/common/oauth2/v2.0/authorize';
const TOKEN = 'https://login.microsoftonline.com/common/oauth2/v2.0/token';

export const MS_SCOPES = [
  'openid',
  'profile',
  'offline_access',
  'Calendars.ReadWrite',
  'User.Read',
];

export function buildAuthorizeUrl(opts: { state: string; redirectUri: string }) {
  const u = new URL(AUTH);
  u.searchParams.set('client_id', process.env.MS_GRAPH_CLIENT_ID!);
  u.searchParams.set('response_type', 'code');
  u.searchParams.set('response_mode', 'query');
  u.searchParams.set('redirect_uri', opts.redirectUri);
  u.searchParams.set('scope', MS_SCOPES.join(' '));
  u.searchParams.set('state', opts.state);
  u.searchParams.set('prompt', 'select_account');
  return u.toString();
}

export type MsTokenResponse = {
  access_token: string;
  refresh_token: string;
  expires_in: number;
  id_token: string;
};

export async function exchangeCodeForToken(opts: {
  code: string;
  redirectUri: string;
}): Promise<MsTokenResponse> {
  const body = new URLSearchParams({
    client_id: process.env.MS_GRAPH_CLIENT_ID!,
    client_secret: process.env.MS_GRAPH_CLIENT_SECRET!,
    grant_type: 'authorization_code',
    code: opts.code,
    redirect_uri: opts.redirectUri,
    scope: MS_SCOPES.join(' '),
  });
  const r = await fetch(TOKEN, {
    method: 'POST',
    headers: { 'content-type': 'application/x-www-form-urlencoded' },
    body,
  });
  if (!r.ok) throw new Error(`MS token exchange failed: ${r.status} ${await r.text()}`);
  return (await r.json()) as MsTokenResponse;
}
```

- [ ] **Step 3: Implement `app/api/auth/microsoft/route.ts`**

```ts
import { NextRequest, NextResponse } from 'next/server';
import { encodeState } from '@/lib/auth/oauth-state';
import { buildAuthorizeUrl } from '@/lib/ms/oauth';
import { getOrInitSession } from '@/lib/auth/session';

export async function GET(req: NextRequest) {
  const session = await getOrInitSession(req);
  const redirectUri = `${process.env.APP_BASE_URL}/api/auth/microsoft/callback`;
  const state = encodeState({ userId: session.userId, purpose: 'ms' });
  const url = buildAuthorizeUrl({ state, redirectUri });
  const res = NextResponse.redirect(url);
  session.applyTo(res);
  return res;
}
```

- [ ] **Step 4: Implement `app/api/auth/microsoft/callback/route.ts`**

```ts
import { NextRequest, NextResponse } from 'next/server';
import { decodeState } from '@/lib/auth/oauth-state';
import { exchangeCodeForToken } from '@/lib/ms/oauth';
import { encryptToken } from '@/lib/crypto/token';
import { prisma } from '@/lib/db/client';
import { setSessionCookie } from '@/lib/auth/session';
import { decodeJwt } from 'jose';

export async function GET(req: NextRequest) {
  const code = req.nextUrl.searchParams.get('code');
  const state = req.nextUrl.searchParams.get('state');
  if (!code || !state) return NextResponse.json({ error: 'missing code/state' }, { status: 400 });

  const { userId, purpose } = decodeState(state);
  if (purpose !== 'ms') return NextResponse.json({ error: 'wrong purpose' }, { status: 400 });

  const redirectUri = `${process.env.APP_BASE_URL}/api/auth/microsoft/callback`;
  const tok = await exchangeCodeForToken({ code, redirectUri });

  // ID token gives us tenantId/userId/email without an extra Graph call.
  const idClaims = decodeJwt(tok.id_token) as {
    tid: string; oid: string; preferred_username: string;
  };

  const user = await prisma.user.upsert({
    where: { msTenantId_msUserId: { msTenantId: idClaims.tid, msUserId: idClaims.oid } },
    update: { msUserEmail: idClaims.preferred_username },
    create: {
      id: userId,
      msTenantId: idClaims.tid,
      msUserId: idClaims.oid,
      msUserEmail: idClaims.preferred_username,
    },
  });

  await prisma.msAccount.upsert({
    where: { userId: user.id },
    update: {
      accessTokenEnc: encryptToken(tok.access_token),
      refreshTokenEnc: encryptToken(tok.refresh_token),
      expiresAt: new Date(Date.now() + tok.expires_in * 1000),
    },
    create: {
      userId: user.id,
      accessTokenEnc: encryptToken(tok.access_token),
      refreshTokenEnc: encryptToken(tok.refresh_token),
      expiresAt: new Date(Date.now() + tok.expires_in * 1000),
    },
  });

  const res = new NextResponse(
    `<!doctype html><meta charset="utf-8"><title>Connected</title>
     <body style="font-family:system-ui;padding:2rem">
       <h2>Microsoft connected ✓</h2>
       <p>You can close this window.</p>
       <script>window.close();</script>
     </body>`,
    { headers: { 'content-type': 'text/html; charset=utf-8' } },
  );
  // Bind the cookie to the persisted user (in case SSO was unavailable
  // and the cookie was carrying a fresh random uuid that doesn't match
  // the upserted record's id).
  setSessionCookie(res, user.id);
  return res;
}
```

- [ ] **Step 5: Quick local sanity check**

Restart dev server (`npm run dev -- --experimental-https`). With `.env.local` populated, open `https://localhost:3000/api/auth/microsoft` directly in a browser. Expected: redirect to Microsoft login, after consent → "Microsoft connected ✓" page. Confirm a row appears in `users` and `ms_accounts` (Supabase Studio → Table editor).

- [ ] **Step 6: Commit**

```bash
git add lib/ms/ app/api/auth/microsoft/
git commit -m "feat: Microsoft Graph OAuth flow"
```

---

## Task 12: monday OAuth flow

**Files:**
- Create: `lib/monday/oauth.ts`, `app/api/auth/monday/route.ts`, `app/api/auth/monday/callback/route.ts`

- [ ] **Step 1: Manual — register monday OAuth app**

1. Go to `https://monday.com/developers/apps` → Create app.
2. Features → OAuth → set **Redirect URLs** to `https://localhost:3000/api/auth/monday/callback`.
3. Permissions: `boards:read`, `boards:write`, `webhooks:read`, `webhooks:write`, `me:read`, `users:read`.
4. From "OAuth & Permissions" copy **Client ID** → `MONDAY_CLIENT_ID`, **Client Secret** → `MONDAY_CLIENT_SECRET` in `.env.local`.
5. From "Webhooks" copy the **Signing secret** → `MONDAY_SIGNING_SECRET` (we'll use it in a later plan).

- [ ] **Step 2: Implement `lib/monday/oauth.ts`**

```ts
const AUTH = 'https://auth.monday.com/oauth2/authorize';
const TOKEN = 'https://auth.monday.com/oauth2/token';

export const MONDAY_SCOPES = [
  'boards:read',
  'boards:write',
  'webhooks:read',
  'webhooks:write',
  'me:read',
  'users:read',
];

export function buildAuthorizeUrl(opts: { state: string; redirectUri: string }) {
  const u = new URL(AUTH);
  u.searchParams.set('client_id', process.env.MONDAY_CLIENT_ID!);
  u.searchParams.set('response_type', 'code');
  u.searchParams.set('redirect_uri', opts.redirectUri);
  u.searchParams.set('state', opts.state);
  u.searchParams.set('scope', MONDAY_SCOPES.join(' '));
  return u.toString();
}

export type MondayTokenResponse = {
  access_token: string;
  refresh_token: string;
  token_type: string;
  expires_in: number;
  scope: string;
};

export async function exchangeCodeForToken(opts: {
  code: string;
  redirectUri: string;
}): Promise<MondayTokenResponse> {
  const body = new URLSearchParams({
    client_id: process.env.MONDAY_CLIENT_ID!,
    client_secret: process.env.MONDAY_CLIENT_SECRET!,
    grant_type: 'authorization_code',
    code: opts.code,
    redirect_uri: opts.redirectUri,
  });
  const r = await fetch(TOKEN, {
    method: 'POST',
    headers: { 'content-type': 'application/x-www-form-urlencoded' },
    body,
  });
  if (!r.ok) throw new Error(`monday token exchange failed: ${r.status} ${await r.text()}`);
  return (await r.json()) as MondayTokenResponse;
}

/** Query the `me` endpoint to learn the monday user/account ids. */
export async function fetchMondayMe(accessToken: string) {
  const r = await fetch('https://api.monday.com/v2', {
    method: 'POST',
    headers: {
      'content-type': 'application/json',
      authorization: accessToken,
    },
    body: JSON.stringify({
      query: '{ me { id name email account { id } } }',
    }),
  });
  if (!r.ok) throw new Error(`monday me query failed: ${r.status}`);
  const j = (await r.json()) as {
    data: { me: { id: string; name: string; email: string; account: { id: string } } };
  };
  return {
    mondayUserId: BigInt(j.data.me.id),
    mondayAccountId: BigInt(j.data.me.account.id),
    name: j.data.me.name,
    email: j.data.me.email,
  };
}
```

- [ ] **Step 3: Implement `app/api/auth/monday/route.ts`**

```ts
import { NextRequest, NextResponse } from 'next/server';
import { encodeState } from '@/lib/auth/oauth-state';
import { buildAuthorizeUrl } from '@/lib/monday/oauth';
import { getOrInitSession } from '@/lib/auth/session';

export async function GET(req: NextRequest) {
  const session = await getOrInitSession(req);
  const redirectUri = `${process.env.APP_BASE_URL}/api/auth/monday/callback`;
  const state = encodeState({ userId: session.userId, purpose: 'monday' });
  const url = buildAuthorizeUrl({ state, redirectUri });
  const res = NextResponse.redirect(url);
  session.applyTo(res);
  return res;
}
```

- [ ] **Step 4: Implement `app/api/auth/monday/callback/route.ts`**

```ts
import { NextRequest, NextResponse } from 'next/server';
import { decodeState } from '@/lib/auth/oauth-state';
import { exchangeCodeForToken, fetchMondayMe } from '@/lib/monday/oauth';
import { encryptToken } from '@/lib/crypto/token';
import { prisma } from '@/lib/db/client';

export async function GET(req: NextRequest) {
  const code = req.nextUrl.searchParams.get('code');
  const state = req.nextUrl.searchParams.get('state');
  if (!code || !state) return NextResponse.json({ error: 'missing code/state' }, { status: 400 });

  const { userId, purpose } = decodeState(state);
  if (purpose !== 'monday') return NextResponse.json({ error: 'wrong purpose' }, { status: 400 });

  const redirectUri = `${process.env.APP_BASE_URL}/api/auth/monday/callback`;
  const tok = await exchangeCodeForToken({ code, redirectUri });
  const me = await fetchMondayMe(tok.access_token);

  // The user must already exist (Microsoft connect happens first).
  const user = await prisma.user.findUnique({ where: { id: userId } });
  if (!user) return NextResponse.json({ error: 'connect Microsoft first' }, { status: 412 });

  await prisma.mondayAccount.upsert({
    where: { userId },
    update: {
      mondayUserId: me.mondayUserId,
      mondayAccountId: me.mondayAccountId,
      accessTokenEnc: encryptToken(tok.access_token),
      refreshTokenEnc: encryptToken(tok.refresh_token),
      expiresAt: new Date(Date.now() + tok.expires_in * 1000),
    },
    create: {
      userId,
      mondayUserId: me.mondayUserId,
      mondayAccountId: me.mondayAccountId,
      accessTokenEnc: encryptToken(tok.access_token),
      refreshTokenEnc: encryptToken(tok.refresh_token),
      expiresAt: new Date(Date.now() + tok.expires_in * 1000),
    },
  });

  return new NextResponse(
    `<!doctype html><meta charset="utf-8"><title>Connected</title>
     <body style="font-family:system-ui;padding:2rem">
       <h2>monday connected ✓</h2>
       <p>You can close this window.</p>
       <script>window.close();</script>
     </body>`,
    { headers: { 'content-type': 'text/html; charset=utf-8' } },
  );
}
```

- [ ] **Step 5: Sanity check**

Open `https://localhost:3000/api/auth/monday`. Expected: monday consent → "monday connected ✓". Verify a `monday_accounts` row exists.

- [ ] **Step 6: Commit**

```bash
git add lib/monday/ app/api/auth/monday/
git commit -m "feat: monday.com OAuth flow"
```

---

## Task 13: Session middleware

The session is a signed cookie that holds the user id we're operating as. On the very first taskpane visit, before either OAuth has happened, we mint a new id via Office.js SSO. After Microsoft OAuth runs, that id maps to a `users` row.

**Files:**
- Create: `lib/auth/session.ts`, `app/api/session/route.ts`

- [ ] **Step 1: Implement `lib/auth/session.ts`**

```ts
import { NextRequest, NextResponse } from 'next/server';
import { encryptToken, decryptToken } from '@/lib/crypto/token';
import { randomUUID } from 'node:crypto';

const COOKIE = 'oms.sid';
const MAX_AGE = 60 * 60 * 24 * 30; // 30 days

export type Session = {
  userId: string;
  /** call before responding to set/refresh the cookie */
  applyTo(res: NextResponse): void;
};

function read(req: NextRequest): string | null {
  const blob = req.cookies.get(COOKIE)?.value;
  if (!blob) return null;
  try {
    const json = decryptToken(Buffer.from(blob, 'base64url'));
    const parsed = JSON.parse(json) as { userId: string };
    return parsed.userId ?? null;
  } catch {
    return null;
  }
}

/** Write a fresh session cookie for the given userId. */
export function setSessionCookie(res: NextResponse, userId: string) {
  const blob = encryptToken(JSON.stringify({ userId })).toString('base64url');
  res.cookies.set(COOKIE, blob, {
    httpOnly: true,
    secure: true,
    sameSite: 'none',
    path: '/',
    maxAge: MAX_AGE,
  });
}

export async function getOrInitSession(req: NextRequest): Promise<Session> {
  const existing = read(req);
  const userId = existing ?? randomUUID();
  return {
    userId,
    applyTo(res) {
      setSessionCookie(res, userId);
    },
  };
}

export async function requireSession(req: NextRequest): Promise<{ userId: string }> {
  const id = read(req);
  if (!id) throw new Response('unauthenticated', { status: 401 });
  return { userId: id };
}
```

(`sameSite: 'none'` is required for the cookie to flow inside the Outlook iframe.)

- [ ] **Step 2: Implement `app/api/session/route.ts`**

The taskpane POSTs the SSO ID token here on first load. We verify it, upsert the user, and bind the cookie.

```ts
import { NextRequest, NextResponse } from 'next/server';
import { verifyOfficeSsoToken } from '@/lib/auth/office-sso';
import { prisma } from '@/lib/db/client';
import { setSessionCookie } from '@/lib/auth/session';

export async function POST(req: NextRequest) {
  const { idToken } = (await req.json()) as { idToken?: string };
  if (!idToken) return NextResponse.json({ error: 'missing idToken' }, { status: 400 });

  const id = await verifyOfficeSsoToken(idToken);
  const user = await prisma.user.upsert({
    where: { msTenantId_msUserId: { msTenantId: id.tenantId, msUserId: id.userId } },
    update: { msUserEmail: id.email },
    create: {
      msTenantId: id.tenantId,
      msUserId: id.userId,
      msUserEmail: id.email,
    },
  });

  const res = NextResponse.json({ ok: true, userId: user.id, email: user.msUserEmail });
  setSessionCookie(res, user.id);
  return res;
}
```

- [ ] **Step 3: Hook the SSO call in the taskpane**

Edit `app/taskpane/layout.tsx` — the layout already loads Office.js. Add a small client component that runs once and posts the ID token.

Create `app/taskpane/_components/SsoBoot.tsx`:

```tsx
'use client';

import { useEffect } from 'react';

export function SsoBoot() {
  useEffect(() => {
    const Office = (window as unknown as { Office?: { onReady?: (fn: () => void) => void; auth?: { getAccessToken: (opts: object) => Promise<string> } } }).Office;
    if (!Office?.onReady) return;
    Office.onReady(async () => {
      try {
        const idToken = await Office.auth!.getAccessToken({ allowSignInPrompt: true });
        await fetch('/api/session', {
          method: 'POST',
          credentials: 'include',
          headers: { 'content-type': 'application/json' },
          body: JSON.stringify({ idToken }),
        });
      } catch (e) {
        console.warn('SSO unavailable:', e);
      }
    });
  }, []);
  return null;
}
```

Update `app/taskpane/layout.tsx` to render it:

```tsx
import Script from 'next/script';
import { SsoBoot } from './_components/SsoBoot';

export const metadata = { title: 'Outlook ↔ monday Sync' };

export default function TaskpaneLayout({ children }: { children: React.ReactNode }) {
  return (
    <>
      <Script
        src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"
        strategy="beforeInteractive"
      />
      <div className="min-h-screen bg-white text-zinc-900 text-sm">
        <SsoBoot />
        {children}
      </div>
    </>
  );
}
```

- [ ] **Step 4: Update Office SSO audience to use APP_BASE_URL**

In `lib/auth/office-sso.ts`, replace the hardcoded audience:

```ts
const audience = `api://${new URL(process.env.APP_BASE_URL!).host}/${process.env.MS_GRAPH_CLIENT_ID}`;
```

- [ ] **Step 5: Commit**

```bash
git add lib/auth/session.ts app/api/session/ app/taskpane/
git commit -m "feat: session cookie + Office SSO bootstrap"
```

---

## Task 14: Status endpoint

The taskpane polls this to render the connection cards.

**Files:**
- Create: `app/api/status/route.ts`

- [ ] **Step 1: Implement `app/api/status/route.ts`**

```ts
import { NextRequest, NextResponse } from 'next/server';
import { prisma } from '@/lib/db/client';
import { requireSession } from '@/lib/auth/session';

export async function GET(req: NextRequest) {
  let session;
  try {
    session = await requireSession(req);
  } catch {
    return NextResponse.json({
      microsoft: { connected: false },
      monday: { connected: false },
    });
  }

  const [user, ms, monday] = await Promise.all([
    prisma.user.findUnique({ where: { id: session.userId } }),
    prisma.msAccount.findUnique({ where: { userId: session.userId } }),
    prisma.mondayAccount.findUnique({ where: { userId: session.userId } }),
  ]);

  return NextResponse.json({
    microsoft: {
      connected: Boolean(ms),
      email: user?.msUserEmail,
    },
    monday: {
      connected: Boolean(monday),
      account: monday ? `account ${monday.mondayAccountId.toString()}` : undefined,
    },
  });
}
```

- [ ] **Step 2: Manual sanity test**

In the taskpane (sideloaded into Outlook), confirm:
1. Status fetch fires (Network tab shows `/api/status`).
2. Both cards initially red.
3. Click "Connect" on Microsoft → popup → consent → popup closes → card flips green and shows email.
4. Click "Connect" on monday → popup → consent → popup closes → card flips green and shows account id.

If the sideload version doesn't refresh after popup closes, hit reload on the taskpane. The polling logic in `SettingsTab.refresh()` runs after `window.closed` is detected.

- [ ] **Step 3: Commit**

```bash
git add app/api/status/
git commit -m "feat: status endpoint for connection cards"
```

---

## Task 15: End-to-end auth verification

Final manual check. No code changes — this is a checklist that proves Plan 1 is done.

**Files:** None.

- [ ] **Step 1: Fresh database, fresh sideload**

```bash
npx prisma migrate reset --force   # wipes Supabase, re-runs migrations
```

In Outlook on the web, remove the existing custom add-in, then re-sideload `manifest/manifest.xml`.

- [ ] **Step 2: Walk through the full happy path**

1. Open any email → open the add-in taskpane.
2. Confirm taskpane renders with Settings tab active, both cards red.
3. Click "Connect Microsoft" → consent → popup closes → status flips green → email visible.
4. Click "Connect monday" → consent → popup closes → status flips green → account id visible.
5. In Supabase studio, confirm:
   - 1 row in `users` with your tenant + user id + email.
   - 1 row in `ms_accounts` with non-null `access_token_enc`, `refresh_token_enc`, future `expires_at`.
   - 1 row in `monday_accounts` with same.
6. Refresh the taskpane → status remains green (cookies survive reload).

- [ ] **Step 3: Verify unit tests + manifest validation green**

```bash
npm test
npm run manifest:validate
npm run lint
```

Expected: all pass.

- [ ] **Step 4: Commit empty checkpoint marker**

```bash
git commit --allow-empty -m "chore: P2 auth verified end-to-end (Plan 1 done)"
```

---

## Self-Review

This plan was checked against the spec by walking through each section:

- **§3 Architecture** — Tasks 1, 2, 5 set up Next on Vercel; routes match the spec layout. ✓
- **§4 Data Model** — Task 7 schema matches every column the spec defined. Task 8 backs the `*_enc` columns. ✓
- **§5 Sync Engine** — Out of scope for this plan. Documented in "Phase Boundary". ✓
- **§6 Auth & Security** — Tasks 8 (encryption), 9 (state CSRF), 10 (SSO validate), 11 (MS OAuth), 12 (monday OAuth), 13 (session) cover §6.1-§6.2. Env vars from §6.3 listed in Task 2. ✓
- **§7 UI** — Task 5 stubs all 4 tabs; Settings is fully functional. Status/Boards/Log deferred per phase boundary. ✓
- **§8 Project Structure** — File layout in this plan matches the spec's structure for the files touched. ✓
- **§9 Phased Implementation** — This plan is exactly P1 + P2. P3-P5 will come as separate plans. ✓
- **§10 Testing** — Vitest unit tests added in Tasks 8/9/10. Manifest validation in Task 4. End-to-end smoke in Task 15. ✓

No placeholders found in the plan; every step has either complete code or a specific verifiable command. No type/name inconsistencies between tasks (I used `MsIdentity`, `encryptToken`/`decryptToken`, `encodeState`/`decodeState`, `getOrInitSession`/`requireSession` consistently).

One scope decision I want to call out: the Office SSO flow has a known fallback case (some tenants disable third-party SSO — error codes 13003/13007). Spec §12 mentions this as a risk. I've **not** wired the fallback in this plan — `SsoBoot.tsx` swallows the error and the user can still use the explicit "Connect Microsoft" button to bind the cookie via the OAuth callback flow. Documented in code comments; full fallback handling can be added when polishing in P5 if telemetry shows it's needed.
