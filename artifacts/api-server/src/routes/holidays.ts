import { Router, type IRouter } from "express";
import { eq } from "drizzle-orm";
import { db, holidaysTable, resourcesTable } from "@workspace/db";
import {
  CreateHolidayBody,
  DeleteHolidayParams,
  ListHolidaysResponse,
} from "@workspace/api-zod";

const router: IRouter = Router();

router.get("/holidays", async (_req, res): Promise<void> => {
  const holidays = await db
    .select({
      id: holidaysTable.id,
      name: holidaysTable.name,
      startDate: holidaysTable.startDate,
      endDate: holidaysTable.endDate,
      resourceId: holidaysTable.resourceId,
      resourceName: resourcesTable.name,
      notes: holidaysTable.notes,
      createdAt: holidaysTable.createdAt,
    })
    .from(holidaysTable)
    .leftJoin(resourcesTable, eq(holidaysTable.resourceId, resourcesTable.id))
    .orderBy(holidaysTable.startDate);

  res.json(ListHolidaysResponse.parse(holidays.map((h) => ({
    ...h,
    resourceId: h.resourceId ?? null,
    resourceName: h.resourceName ?? null,
    notes: h.notes ?? null,
    createdAt: h.createdAt.toISOString(),
  }))));
});

router.post("/holidays", async (req, res): Promise<void> => {
  const parsed = CreateHolidayBody.safeParse(req.body);
  if (!parsed.success) {
    res.status(400).json({ error: parsed.error.message });
    return;
  }
  const [holiday] = await db.insert(holidaysTable).values({
    name: parsed.data.name,
    startDate: parsed.data.startDate,
    endDate: parsed.data.endDate,
    resourceId: parsed.data.resourceId ?? null,
    notes: parsed.data.notes ?? null,
  }).returning();

  let resourceName: string | null = null;
  if (holiday.resourceId) {
    const [r] = await db.select({ name: resourcesTable.name }).from(resourcesTable).where(eq(resourcesTable.id, holiday.resourceId));
    resourceName = r?.name ?? null;
  }

  res.status(201).json({
    id: holiday.id,
    name: holiday.name,
    startDate: holiday.startDate,
    endDate: holiday.endDate,
    resourceId: holiday.resourceId ?? null,
    resourceName,
    notes: holiday.notes ?? null,
    createdAt: holiday.createdAt.toISOString(),
  });
});

router.delete("/holidays/:id", async (req, res): Promise<void> => {
  const params = DeleteHolidayParams.safeParse(req.params);
  if (!params.success) {
    res.status(400).json({ error: params.error.message });
    return;
  }
  const [deleted] = await db.delete(holidaysTable).where(eq(holidaysTable.id, params.data.id)).returning();
  if (!deleted) {
    res.status(404).json({ error: "Holiday not found" });
    return;
  }
  res.sendStatus(204);
});

export default router;
