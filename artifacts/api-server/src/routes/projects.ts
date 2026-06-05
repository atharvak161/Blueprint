import { Router, type IRouter } from "express";
import { eq, sql } from "drizzle-orm";
import { db, projectsTable, phasesTable, tasksTable } from "@workspace/db";
import {
  CreateProjectBody,
  GetProjectParams,
  UpdateProjectParams,
  UpdateProjectBody,
  DeleteProjectParams,
  GetProjectSummaryParams,
  ListProjectsResponse,
  GetProjectResponse,
  UpdateProjectResponse,
  GetProjectSummaryResponse,
} from "@workspace/api-zod";

const router: IRouter = Router();

router.get("/projects", async (_req, res): Promise<void> => {
  const projects = await db
    .select()
    .from(projectsTable)
    .orderBy(projectsTable.createdAt);
  res.json(ListProjectsResponse.parse(projects.map((p) => ({
    ...p,
    startDate: p.startDate ?? null,
    endDate: p.endDate ?? null,
    description: p.description ?? null,
    createdAt: p.createdAt.toISOString(),
    updatedAt: p.updatedAt.toISOString(),
  }))));
});

router.post("/projects", async (req, res): Promise<void> => {
  const parsed = CreateProjectBody.safeParse(req.body);
  if (!parsed.success) {
    res.status(400).json({ error: parsed.error.message });
    return;
  }
  const [project] = await db.insert(projectsTable).values({
    name: parsed.data.name,
    description: parsed.data.description ?? null,
    startDate: parsed.data.startDate ?? null,
    endDate: parsed.data.endDate ?? null,
  }).returning();
  res.status(201).json(GetProjectResponse.parse({
    ...project,
    startDate: project.startDate ?? null,
    endDate: project.endDate ?? null,
    description: project.description ?? null,
    createdAt: project.createdAt.toISOString(),
    updatedAt: project.updatedAt.toISOString(),
  }));
});

router.get("/projects/:id", async (req, res): Promise<void> => {
  const params = GetProjectParams.safeParse(req.params);
  if (!params.success) {
    res.status(400).json({ error: params.error.message });
    return;
  }
  const [project] = await db.select().from(projectsTable).where(eq(projectsTable.id, params.data.id));
  if (!project) {
    res.status(404).json({ error: "Project not found" });
    return;
  }
  res.json(GetProjectResponse.parse({
    ...project,
    startDate: project.startDate ?? null,
    endDate: project.endDate ?? null,
    description: project.description ?? null,
    createdAt: project.createdAt.toISOString(),
    updatedAt: project.updatedAt.toISOString(),
  }));
});

router.patch("/projects/:id", async (req, res): Promise<void> => {
  const params = UpdateProjectParams.safeParse(req.params);
  if (!params.success) {
    res.status(400).json({ error: params.error.message });
    return;
  }
  const parsed = UpdateProjectBody.safeParse(req.body);
  if (!parsed.success) {
    res.status(400).json({ error: parsed.error.message });
    return;
  }
  const updates: Record<string, unknown> = {};
  if (parsed.data.name !== undefined) updates.name = parsed.data.name;
  if (parsed.data.description !== undefined) updates.description = parsed.data.description;
  if (parsed.data.startDate !== undefined) updates.startDate = parsed.data.startDate;
  if (parsed.data.endDate !== undefined) updates.endDate = parsed.data.endDate;

  const [project] = await db
    .update(projectsTable)
    .set(updates)
    .where(eq(projectsTable.id, params.data.id))
    .returning();
  if (!project) {
    res.status(404).json({ error: "Project not found" });
    return;
  }
  res.json(UpdateProjectResponse.parse({
    ...project,
    startDate: project.startDate ?? null,
    endDate: project.endDate ?? null,
    description: project.description ?? null,
    createdAt: project.createdAt.toISOString(),
    updatedAt: project.updatedAt.toISOString(),
  }));
});

router.delete("/projects/:id", async (req, res): Promise<void> => {
  const params = DeleteProjectParams.safeParse(req.params);
  if (!params.success) {
    res.status(400).json({ error: params.error.message });
    return;
  }
  const [deleted] = await db.delete(projectsTable).where(eq(projectsTable.id, params.data.id)).returning();
  if (!deleted) {
    res.status(404).json({ error: "Project not found" });
    return;
  }
  res.sendStatus(204);
});

router.get("/projects/:id/summary", async (req, res): Promise<void> => {
  const params = GetProjectSummaryParams.safeParse(req.params);
  if (!params.success) {
    res.status(400).json({ error: params.error.message });
    return;
  }
  const projectId = params.data.id;

  const [project] = await db.select({ id: projectsTable.id }).from(projectsTable).where(eq(projectsTable.id, projectId));
  if (!project) {
    res.status(404).json({ error: "Project not found" });
    return;
  }

  const tasks = await db
    .select({ id: tasksTable.id, status: tasksTable.status, pctDone: tasksTable.pctDone, phaseId: tasksTable.phaseId })
    .from(tasksTable)
    .where(eq(tasksTable.projectId, projectId));

  const phases = await db
    .select({ id: phasesTable.id, name: phasesTable.name })
    .from(phasesTable)
    .where(eq(phasesTable.projectId, projectId));

  const total = tasks.length;
  const completed = tasks.filter((t) => t.status === "Completed").length;
  const inProgress = tasks.filter((t) => t.status === "In Progress").length;
  const notStarted = tasks.filter((t) => t.status === "Not Started").length;
  const overdue = tasks.filter((t) => t.status === "Overdue").length;
  const blocked = tasks.filter((t) => t.status === "Blocked").length;
  const pctComplete = total > 0 ? (tasks.reduce((s, t) => s + t.pctDone, 0) / total) : 0;

  const phaseStats = phases.map((phase) => {
    const phaseTasks = tasks.filter((t) => t.phaseId === phase.id);
    const phaseTotal = phaseTasks.length;
    const phaseCompleted = phaseTasks.filter((t) => t.status === "Completed").length;
    return {
      phaseId: phase.id,
      phaseName: phase.name,
      total: phaseTotal,
      completed: phaseCompleted,
      pctComplete: phaseTotal > 0 ? (phaseCompleted / phaseTotal) * 100 : 0,
    };
  });

  res.json(GetProjectSummaryResponse.parse({
    projectId,
    totalTasks: total,
    completed,
    inProgress,
    notStarted,
    overdue,
    blocked,
    pctComplete,
    phaseStats,
  }));
});

export default router;
