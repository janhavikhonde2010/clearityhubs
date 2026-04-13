# Workspace

## Overview

pnpm workspace monorepo using TypeScript. Each package manages its own dependencies.

## Stack

- **Monorepo tool**: pnpm workspaces
- **Node.js version**: 24
- **Package manager**: pnpm
- **TypeScript version**: 5.9
- **API framework**: Express 5
- **Database**: PostgreSQL + Drizzle ORM
- **Validation**: Zod (`zod/v4`), `drizzle-zod`
- **API codegen**: Orval (from OpenAPI spec)
- **Build**: esbuild (CJS bundle)

## Key Commands

- `pnpm run typecheck` — full typecheck across all packages
- `pnpm run build` — typecheck + build all packages
- `pnpm --filter @workspace/api-spec run codegen` — regenerate API hooks and Zod schemas from OpenAPI spec
- `pnpm --filter @workspace/db run push` — push DB schema changes (dev only)
- `pnpm --filter @workspace/api-server run dev` — run API server locally

See the `pnpm-workspace` skill for workspace structure, TypeScript setup, and package details.

## Artifacts

### TWP Analytics Dashboard (`artifacts/twp-dashboard`)
- WhatsApp lead analytics dashboard for The Wise Parrot (TWP)
- Connects to `growth.thewiseparrot.club` APIs
- Pages: Dashboard, Subscribers, Agents, Sequences
- Requires API Token + Phone Number ID (stored in localStorage)
- No database needed — all data comes from external TWP APIs

### API Server (`artifacts/api-server`)
- Express 5 backend
- Routes: `/api/subscribers`, `/api/dashboard/summary`, `/api/dashboard/agent-stats`, `/api/dashboard/sequence-stats`, `/api/dashboard/label-stats`
- Proxies and processes TWP API responses
- In-memory cache (5 min TTL) per credentials pair
- Logic: `artifacts/api-server/src/lib/twp-api.ts`
