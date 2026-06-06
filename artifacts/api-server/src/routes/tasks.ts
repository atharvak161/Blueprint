import { Router, type IRouter } from "express";
import { eq, asc, and } from "drizzle-orm";
import { db, tasksTable, resourcesTable, phasesTable } from "@workspace/db";
import {
  ListTasksParams,
  CreateTaskParams,
  CreateTaskBody,
  UpdateTaskParams,
  UpdateTaskBody,
  DeleteTaskParams,
  ListTasksResponse,
  UpdateTaskResponse,
} from "@workspace/api-zod";

const router: IRouter = Router();

async function taskOut(t: typeof tasksTable.$inferSelect & { resourceName?: string | null }) {
  return {
    id: t.id,
    projectId: t.projectId,
    phaseId: t.phaseId,
    parentTaskId: t.parentTaskId ?? null,
    taskNumber: t.taskNumber,
    taskType: t.taskType,
    description: t.description,
    plannedStart: t.plannedStart ?? null,
    plannedEnd: t.plannedEnd ?? null,
    resourceId: t.resourceId ?? null,
    resourceName: t.resourceName ?? null,
    pctDone: t.pctDone,
    status: t.status,
    notes: t.notes ?? null,
    sortOrder: t.sortOrder,
    createdAt: t.createdAt.toISOString(),
    updatedAt: t.updatedAt.toISOString(),
  };
}

router.get("/projects/:projectId/tasks", async (req, res): Promise<void> => {
  const params = ListTasksParams.safeParse(req.params);
  if (!params.success) {
    res.status(400).json({ error: params.error.message });
    return;
  }
  const tasks = await db
    .select({
      id: tasksTable.id,
      projectId: tasksTable.projectId,
      phaseId: tasksTable.phaseId,
      parentTaskId: tasksTable.parentTaskId,
      taskNumber: tasksTable.taskNumber,
      taskType: tasksTable.taskType,
      description: tasksTable.description,
      plannedStart: tasksTable.plannedStart,
      plannedEnd: tasksTable.plannedEnd,
      resourceId: tasksTable.resourceId,
      resourceName: resourcesTable.name,
      pctDone: tasksTable.pctDone,
      status: tasksTable.status,
      notes: tasksTable.notes,
      sortOrder: tasksTable.sortOrder,
      createdAt: tasksTable.createdAt,
      updatedAt: tasksTable.updatedAt,
    })
    .from(tasksTable)
    .leftJoin(resourcesTable, eq(tasksTable.resourceId, resourcesTable.id))
    .where(eq(tasksTable.projectId, params.data.projectId))
    .orderBy(asc(tasksTable.sortOrder), asc(tasksTable.taskNumber));

  const parsed = await Promise.all(tasks.map(taskOut));
  res.json(ListTasksResponse.parse(parsed));
});

router.post("/projects/:projectId/tasks/renumber", async (req, res): Promise<void> => {
  const projectId = parseInt(req.params.projectId, 10);
  if (isNaN(projectId)) {
    res.status(400).json({ error: "Invalid projectId" });
    return;
  }
  let renumbered = 0;
  await db.transaction(async (tx) => {
    const allTasks = await tx
      .select({ id: tasksTable.id })
      .from(tasksTable)
      .where(eq(tasksTable.projectId, projectId))
      .orderBy(asc(tasksTable.sortOrder), asc(tasksTable.id));

    for (let i = 0; i < allTasks.length; i++) {
      await tx
        .update(tasksTable)
        .set({ taskNumber: i + 1 })
        .where(eq(tasksTable.id, allTasks[i].id));
    }
    renumbered = allTasks.length;
  });
  res.json({ renumbered });
});

router.post("/projects/:projectId/tasks", async (req, res): Promise<void> => {
  const params = CreateTaskParams.safeParse(req.params);
  if (!params.success) {
    res.status(400).json({ error: params.error.message });
    return;
  }
  const parsed = CreateTaskBody.safeParse(req.body);
  if (!parsed.success) {
    res.status(400).json({ error: parsed.error.message });
    return;
  }

  // Verify phaseId belongs to this project
  if (parsed.data.phaseId !== undefined) {
    const [phase] = await db
      .select({ id: phasesTable.id })
      .from(phasesTable)
      .where(and(eq(phasesTable.id, parsed.data.phaseId), eq(phasesTable.projectId, params.data.projectId)));
    if (!phase) {
      res.status(400).json({ error: "phaseId does not belong to this project" });
      return;
    }
  }

  const existingTasks = await db
    .select({ taskNumber: tasksTable.taskNumber })
    .from(tasksTable)
    .where(eq(tasksTable.projectId, params.data.projectId));
  const maxNum = existingTasks.reduce((m, t) => Math.max(m, t.taskNumber), 0);

  const [task] = await db.insert(tasksTable).values({
    projectId: params.data.projectId,
    phaseId: parsed.data.phaseId,
    parentTaskId: parsed.data.parentTaskId ?? null,
    taskNumber: maxNum + 1,
    taskType: parsed.data.taskType,
    description: parsed.data.description,
    plannedStart: parsed.data.plannedStart ?? null,
    plannedEnd: parsed.data.plannedEnd ?? null,
    resourceId: parsed.data.resourceId ?? null,
    pctDone: parsed.data.pctDone ?? 0,
    status: parsed.data.status ?? "Not Started",
    notes: parsed.data.notes ?? null,
    sortOrder: parsed.data.sortOrder ?? maxNum + 1,
  }).returning();

  let resourceName: string | null = null;
  if (task.resourceId) {
    const [r] = await db.select({ name: resourcesTable.name }).from(resourcesTable).where(eq(resourcesTable.id, task.resourceId));
    resourceName = r?.name ?? null;
  }

  res.status(201).json(UpdateTaskResponse.parse(await taskOut({ ...task, resourceName })));
});

router.patch("/tasks/:id", async (req, res): Promise<void> => {
  const params = UpdateTaskParams.safeParse(req.params);
  if (!params.success) {
    res.status(400).json({ error: params.error.message });
    return;
  }
  const parsed = UpdateTaskBody.safeParse(req.body);
  if (!parsed.success) {
    res.status(400).json({ error: parsed.error.message });
    return;
  }
  const updates: Record<string, unknown> = {};
  if (parsed.data.phaseId !== undefined) updates.phaseId = parsed.data.phaseId;
  if (parsed.data.parentTaskId !== undefined) updates.parentTaskId = parsed.data.parentTaskId;
  if (parsed.data.taskType !== undefined) updates.taskType = parsed.data.taskType;
  if (parsed.data.description !== undefined) updates.description = parsed.data.description;
  if (parsed.data.plannedStart !== undefined) updates.plannedStart = parsed.data.plannedStart;
  if (parsed.data.plannedEnd !== undefined) updates.plannedEnd = parsed.data.plannedEnd;
  if (parsed.data.resourceId !== undefined) updates.resourceId = parsed.data.resourceId;
  if (parsed.data.pctDone !== undefined) updates.pctDone = parsed.data.pctDone;
  if (parsed.data.status !== undefined) updates.status = parsed.data.status;
  if (parsed.data.notes !== undefined) updates.notes = parsed.data.notes;
  if (parsed.data.sortOrder !== undefined) updates.sortOrder = parsed.data.sortOrder;

  const [task] = await db.update(tasksTable).set(updates).where(eq(tasksTable.id, params.data.id)).returning();
  if (!task) {
    res.status(404).json({ error: "Task not found" });
    return;
  }
  let resourceName: string | null = null;
  if (task.resourceId) {
    const [r] = await db.select({ name: resourcesTable.name }).from(resourcesTable).where(eq(resourcesTable.id, task.resourceId));
    resourceName = r?.name ?? null;
  }
  res.json(UpdateTaskResponse.parse(await taskOut({ ...task, resourceName })));
});

router.delete("/tasks/:id", async (req, res): Promise<void> => {
  const params = DeleteTaskParams.safeParse(req.params);
  if (!params.success) {
    res.status(400).json({ error: params.error.message });
    return;
  }
  const [deleted] = await db.delete(tasksTable).where(eq(tasksTable.id, params.data.id)).returning();
  if (!deleted) {
    res.status(404).json({ error: "Task not found" });
    return;
  }

  await db.transaction(async (tx) => {
    const remaining = await tx
      .select({ id: tasksTable.id })
      .from(tasksTable)
      .where(eq(tasksTable.projectId, deleted.projectId))
      .orderBy(asc(tasksTable.sortOrder), asc(tasksTable.id));
    for (let i = 0; i < remaining.length; i++) {
      await tx
        .update(tasksTable)
        .set({ taskNumber: i + 1 })
        .where(eq(tasksTable.id, remaining[i].id));
    }
  });

  res.sendStatus(204);
});

export default router;
