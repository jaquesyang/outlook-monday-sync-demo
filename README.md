# monday Sync — Outlook Calendar Add-in

Bidirectional sync between Outlook calendar and monday.com boards. Built as an Outlook Office Add-in with a Next.js backend.

## Features

- **OAuth authentication** — Connect Microsoft 365 and monday.com accounts independently
- **Bidirectional sync** — Calendar events from Outlook sync to monday.com items and vice versa
- **Board subscription** — Select monday.com boards and date columns to sync
- **Office SSO bootstrap** — Automatic identity detection via Outlook's SSO token
- **Session cookie auth** — Secure session management for the taskpane
- **Encrypted token storage** — All OAuth tokens are encrypted at rest

## Tech Stack

| Layer | Technology |
|-------|-----------|
| Framework | Next.js 16 (App Router) + React 19 |
| Language | TypeScript |
| Styling | Tailwind CSS 4 |
| ORM | Prisma 5 + PostgreSQL |
| Auth | Jose (JWT), custom session cookies |
| HTTP Client | Native fetch |
| Testing | Vitest |
| Deployment | Vercel |
| Add-in Runtime | Office.js (Outlook) |

## Project Structure

```
├── app/
│   ├── api/
│   │   ├── auth/microsoft/        # Microsoft OAuth flow
│   │   ├── auth/monday/           # monday.com OAuth flow
│   │   ├── session/route.ts       # Session cookie endpoint
│   │   ├── status/route.ts        # Connection status check
│   │   └── sync/route.ts          # Calendar sync operations
│   ├── taskpane/                  # Outlook taskpane UI
│   │   ├── _components/
│   │   │   ├── BoardsTab.tsx      # Board selection UI
│   │   │   ├── SettingsTab.tsx    # Connection settings
│   │   │   └── SsoBoot.tsx        # SSO bootstrap
│   │   ├── layout.tsx
│   │   └── page.tsx
│   └── commands/page.tsx          # Command surface fallback
├── lib/
│   ├── auth/                      # OAuth state, sessions, SSO
│   ├── crypto/token.ts            # Token encryption utilities
│   ├── db/client.ts               # Prisma client singleton
│   ├── env.ts                     # Environment helper (VERCEL_URL aware)
│   ├── monday/                    # monday.com API & OAuth
│   └── ms/                        # Microsoft Graph & OAuth
├── prisma/
│   └── schema.prisma              # Database schema
├── manifest/
│   └── manifest.xml               # Office Add-in manifest
├── scripts/
│   └── start-tunnel.ts            # Dev tunnel + manifest updater
└── tests/                         # Vitest unit tests
```

## Database Schema

```
users                — Microsoft identity (tenant + user)
ms_accounts          — Microsoft Graph tokens & settings
monday_accounts      — monday.com tokens & board defaults
board_subscriptions  — Selected boards + date columns to sync
event_mappings       — Outlook event ↔ monday item links
sync_log             — Sync operation audit trail
```

## Environment Variables

Copy `.env.example` to `.env.local` and fill in:

| Variable | Purpose |
|----------|---------|
| `POSTGRES_URL` | PostgreSQL connection string (pooled via PgBouncer on Supabase) |
| `POSTGRES_URL_NON_POOLING` | Direct connection for Prisma migrations |
| `POSTGRES_PRISMA_URL` | Pooled connection with `pgbouncer=true` (preferred for Prisma Client at runtime) |
| `MS_GRAPH_CLIENT_ID` | Azure AD app registration ID |
| `MS_GRAPH_CLIENT_SECRET` | Azure AD client secret |
| `MONDAY_CLIENT_ID` | monday.com OAuth app ID |
| `MONDAY_CLIENT_SECRET` | monday.com OAuth secret |
| `MONDAY_SIGNING_SECRET` | monday.com webhook signing secret |
| `TOKEN_ENC_KEY` | Base64-encoded 32-byte key for token encryption |
| `APP_BASE_URL` | Public origin for OAuth redirects (local dev only) |

## Development

```bash
# Install dependencies
npm install

# Run database migrations
npx prisma migrate dev

# Start dev server with HTTPS (required for Office Add-in)
npm run dev

# Or start with Cloudflare tunnel + auto manifest update
npm run dev:tunnel

# Validate Office Add-in manifest
npm run manifest:validate

# Run tests
npm test
```

## Deployment

Deployed on Vercel. The build command includes `prisma generate` and `prisma migrate deploy`.

> **Note:** Supabase network restrictions may block Vercel's build IP. If migrations fail during deploy, run the SQL in `scripts/supabase-migrate.sql` manually via Supabase Dashboard → SQL Editor.

## Architecture Notes

- `getAppBaseUrl()` in `lib/env.ts` prefers `VERCEL_URL` for dynamic preview/production URLs, falling back to `APP_BASE_URL` and then `localhost:3000`.
- The manifest supports three Outlook surfaces: **MessageRead**, **AppointmentOrganizer**, and **AppointmentAttendee**.
- `prisma migrate deploy` is wrapped with `timeout 30` in the Vercel build to prevent hanging on unreachable pooler connections.
- Prisma Client prefers `POSTGRES_PRISMA_URL` at runtime (Vercel Supabase integration provides this with `pgbouncer=true`), falling back to `POSTGRES_URL_NON_POOLING` then `POSTGRES_URL`. This avoids PgBouncer prepared-statement conflicts (error 42P05) while keeping connection pooling on serverless.
