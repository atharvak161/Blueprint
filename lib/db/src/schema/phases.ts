import { pgTable, serial, text, timestamp, integer, boolean } from "drizzle-orm/pg-core";
import { createInsertSchema } from "drizzle-zod";
import { z } from "zod/v4";
import { projectsTable } from "./projects";

export const phasesTable = pgTable("phases", {
  id: serial("id").primaryKey(),
  projectId: integer("project_id").notNull().references(() => projectsTable.id, { onDelete: "cascade" }),
  name: text("name").notNull(),
  color: text("color"),
  orderIndex: integer("order_index").notNull().default(0),
  collapsed: boolean("collapsed").notNull().default(false),
  createdAt: timestamp("created_at", { withTimezone: true }).notNull().defaultNow(),
});

export const insertPhaseSchema = createInsertSchema(phasesTable).omit({ id: true, createdAt: true });
export type InsertPhase = z.infer<typeof insertPhaseSchema>;
export type Phase = typeof phasesTable.$inferSelect;
