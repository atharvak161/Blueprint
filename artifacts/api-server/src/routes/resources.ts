import { Router, type IRouter } from "express";
import { eq } from "drizzle-orm";
import { db, resourcesTable } from "@workspace/db";
import {
  CreateResourceBody,
  UpdateResourceParams,
  UpdateResourceBody,
  DeleteResourceParams,
  ListResourcesResponse,
  UpdateResourceResponse,
} from "@workspace/api-zod";

const router: IRouter = Router();

function resourceOut(r: typeof resourcesTable.$inferSelect) {
  return {
    ...r,
    department: r.department ?? null,
    role: r.role ?? null,
    email: r.email ?? null,
    createdAt: r.createdAt.toISOString(),
  };
}

router.get("/resources", async (_req, res): Promise<void> => {
  const resources = await db.select().from(resourcesTable).orderBy(resourcesTable.name);
  res.json(ListResourcesResponse.parse(resources.map(resourceOut)));
});

router.post("/resources", async (req, res): Promise<void> => {
  const parsed = CreateResourceBody.safeParse(req.body);
  if (!parsed.success) {
    res.status(400).json({ error: parsed.error.message });
    return;
  }
  const [resource] = await db.insert(resourcesTable).values({
    name: parsed.data.name,
    department: parsed.data.department ?? null,
    role: parsed.data.role ?? null,
    email: parsed.data.email ?? null,
  }).returning();
  res.status(201).json(UpdateResourceResponse.parse(resourceOut(resource)));
});

router.patch("/resources/:id", async (req, res): Promise<void> => {
  const params = UpdateResourceParams.safeParse(req.params);
  if (!params.success) {
    res.status(400).json({ error: params.error.message });
    return;
  }
  const parsed = UpdateResourceBody.safeParse(req.body);
  if (!parsed.success) {
    res.status(400).json({ error: parsed.error.message });
    return;
  }
  const updates: Record<string, unknown> = {};
  if (parsed.data.name !== undefined) updates.name = parsed.data.name;
  if (parsed.data.department !== undefined) updates.department = parsed.data.department;
  if (parsed.data.role !== undefined) updates.role = parsed.data.role;
  if (parsed.data.email !== undefined) updates.email = parsed.data.email;

  const [resource] = await db.update(resourcesTable).set(updates).where(eq(resourcesTable.id, params.data.id)).returning();
  if (!resource) {
    res.status(404).json({ error: "Resource not found" });
    return;
  }
  res.json(UpdateResourceResponse.parse(resourceOut(resource)));
});

router.delete("/resources/:id", async (req, res): Promise<void> => {
  const params = DeleteResourceParams.safeParse(req.params);
  if (!params.success) {
    res.status(400).json({ error: params.error.message });
    return;
  }
  const [deleted] = await db.delete(resourcesTable).where(eq(resourcesTable.id, params.data.id)).returning();
  if (!deleted) {
    res.status(404).json({ error: "Resource not found" });
    return;
  }
  res.sendStatus(204);
});

export default router;
