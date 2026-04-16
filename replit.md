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

## Computo XLSX Generator (Main App)

A Python FastAPI web app for generating branded Excel files from JSON bundles.

- **Location**: `playbook_industrial/`
- **Entry point**: `playbook_industrial/main.py`
- **Engine modules**: `genera_computo.py`, `validate_bundle.py` (do not modify)
- **Default template**: `playbook_industrial/templates/Computo preliminare-V3.xlsx`
- **Frontend**: `playbook_industrial/static/index.html` (single-page, vanilla JS)
- **Run**: `cd playbook_industrial && python main.py`
- **Port**: `PORT` env var (default 8000)
- **Dependencies**: `playbook_industrial/requirements.txt`

## Key Commands (monorepo)

- `pnpm run typecheck` — full typecheck across all packages
- `pnpm run build` — typecheck + build all packages
- `pnpm --filter @workspace/api-spec run codegen` — regenerate API hooks and Zod schemas from OpenAPI spec
- `pnpm --filter @workspace/db run push` — push DB schema changes (dev only)
- `pnpm --filter @workspace/api-server run dev` — run API server locally

See the `pnpm-workspace` skill for workspace structure, TypeScript setup, and package details.
