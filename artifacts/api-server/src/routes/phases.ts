import { Router, type IRouter } from "express";
import { eq } from "drizzle-orm";
import { db, phasesTable, projectsTable } from "@workspace/db";
import {
  ListPhasesParams,
  CreatePhaseParams,
  CreatePhaseBody,
  UpdatePhaseParams,
  UpdatePhaseBody,
  DeletePhaseParams,
  ListPhasesResponse,
  UpdatePhaseResponse,
} from "@workspace/api-zod";

const router: IRouter = Router();

function phaseOut(p: typeof phasesTable.$inferSelect) {
  return {
    ...p,
    color: p.color ?? null,
    createdAt: p.createdAt.toISOString(),
  };
}

router.get("/projects/:projectId/phases", async (req, res): Promise<void> => {
  const params = ListPhasesParams.safeParse(req.params);
  if (!params.success) {
    res.status(400).json({ error: params.error.message });
    return;
  }
  const phases = await db
    .select()
    .from(phasesTable)
    .where(eq(phasesTable.projectId, params.data.projectId))
    .orderBy(phasesTable.orderIndex);
  res.json(ListPhasesResponse.parse(phases.map(phaseOut)));
});

router.post("/projects/:projectId/phases", async (req, res): Promise<void> => {
  const params = CreatePhaseParams.safeParse(req.params);
  if (!params.success) {
    res.status(400).json({ error: params.error.message });
    return;
  }
  const parsed = CreatePhaseBody.safeParse(req.body);
  if (!parsed.success) {
    res.status(400).json({ error: parsed.error.message });
    return;
  }
  const existing = await db.select({ count: phasesTable.id }).from(phasesTable).where(eq(phasesTable.projectId, params.data.projectId));
  const orderIndex = parsed.data.orderIndex ?? existing.length;
  const [phase] = await db.insert(phasesTable).values({
    projectId: params.data.projectId,
    name: parsed.data.name,
    color: parsed.data.color ?? null,
    orderIndex,
    collapsed: false,
  }).returning();
  res.status(201).json(UpdatePhaseResponse.parse(phaseOut(phase)));
});

router.patch("/phases/:id", async (req, res): Promise<void> => {
  const params = UpdatePhaseParams.safeParse(req.params);
  if (!params.success) {
    res.status(400).json({ error: params.error.message });
    return;
  }
  const parsed = UpdatePhaseBody.safeParse(req.body);
  if (!parsed.success) {
    res.status(400).json({ error: parsed.error.message });
    return;
  }
  const updates: Record<string, unknown> = {};
  if (parsed.data.name !== undefined) updates.name = parsed.data.name;
  if (parsed.data.color !== undefined) updates.color = parsed.data.color;
  if (parsed.data.orderIndex !== undefined) updates.orderIndex = parsed.data.orderIndex;
  if (parsed.data.collapsed !== undefined) updates.collapsed = parsed.data.collapsed;

  const [phase] = await db.update(phasesTable).set(updates).where(eq(phasesTable.id, params.data.id)).returning();
  if (!phase) {
    res.status(404).json({ error: "Phase not found" });
    return;
  }
  res.json(UpdatePhaseResponse.parse(phaseOut(phase)));
});

router.delete("/phases/:id", async (req, res): Promise<void> => {
  const params = DeletePhaseParams.safeParse(req.params);
  if (!params.success) {
    res.status(400).json({ error: params.error.message });
    return;
  }
  const [deleted] = await db.delete(phasesTable).where(eq(phasesTable.id, params.data.id)).returning();
  if (!deleted) {
    res.status(404).json({ error: "Phase not found" });
    return;
  }
  res.sendStatus(204);
});

export default router;
