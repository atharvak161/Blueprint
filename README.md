# Blueprint

Project management dashboard for tracking projects, phases, tasks, timelines, resources, and team status.

**Live:** https://atharvak161.github.io/Blueprint/

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
