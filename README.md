# Blueprint

Project management dashboard for tracking projects, phases, tasks, timelines, resources, and team status.

**Live:** https://atharvak161.github.io/Blueprint/

## Architecture

### Monorepo layout

```
  pnpm workspace + shared dependency catalog
  │
  ├── artifacts/          runnable applications
  │   ├── project-tracker   the dashboard — React 19, Vite, wouter, Tailwind
  │   ├── api-server        the backend
  │   └── mockup-sandbox    isolated UI experiments
  │
  ├── lib/                shared packages, consumed by the artifacts
  │   ├── api-spec          the contract — what endpoints exist and their shape
  │   ├── api-zod           Zod schemas generated from the spec
  │   ├── api-client-react  typed React Query hooks generated from the spec
  │   └── db                Drizzle ORM schema and client
  │
  └── scripts/            workspace tooling, post-merge hooks
```

### The contract chain

```
  lib/api-spec  ──▶  lib/api-zod      runtime validation, both sides
        │        ──▶  lib/api-client-react   typed hooks for the frontend
        │        ──▶  artifacts/api-server   implements against the same spec
        ▼
  one definition, three consumers — a change to the spec breaks the build
  everywhere it is wrong, at compile time rather than in production
```

This is the reason for the monorepo. Frontend and backend share one source of
truth for the API, so the client cannot drift from the server: rename a field
in the spec and TypeScript fails in the server, the hooks and the dashboard
together.

### Data

`lib/db` holds the Drizzle schema and is the only package that talks to
PostgreSQL. `api-server` imports it; nothing in `artifacts/project-tracker`
does — the frontend reaches data exclusively through the generated React Query
hooks in `lib/api-client-react`.

### Dependency management

`pnpm-workspace.yaml` declares a `catalog:` so React, Tailwind, Drizzle, Zod
and the rest are pinned once for the whole workspace. No package picks its own
version of a shared dependency, so two artifacts cannot end up running
different Reacts.

### Why it is shaped this way

**Generate the client, never hand-write it.** A hand-written API client is a
second, silent definition of the contract that drifts the moment someone edits
one side. Generating `api-zod` and `api-client-react` from `api-spec` makes
drift a build error.

**Artifacts are runnable, libs are not.** Anything in `artifacts/` starts and
serves. Anything in `lib/` is imported. Keeping that boundary strict means it
is always obvious where to add a feature and what a change can affect.


## Features

- **Projects** — create, edit, and delete projects with descriptions and date ranges; per-project KPI badges (done / active / overdue / blocked / % complete).
- **Plan view** — group work into phases, add tasks and subtasks, track status and completion percentage inline.
- **Gantt view** — visualise project timelines.
- **Dashboard** — roll-up KPIs and per-phase progress.
- **RAG view** — red/amber/green status reporting across the project.
- **Resources** — manage the people assigned to projects.
- **Holidays** — maintain a holiday calendar that timelines account for.
- **Backup / restore** — export all data to a JSON file and import it back.

## Stack

- React + TypeScript, built with Vite.
- Tailwind CSS with Radix UI / shadcn-style components.
- `wouter` for routing, TanStack Query for data fetching.

The live GitHub Pages build runs entirely in the browser — data is persisted to
`localStorage`, so there is no backend or sign-in. The repo is a pnpm monorepo
that also contains an optional Express + Drizzle API server (`artifacts/api-server`)
for running against a real database, but the deployed app does not use it.

## Usage

Visit the live link above — no setup required.

To run locally:

```bash
pnpm install
pnpm --filter @workspace/project-tracker dev
```
