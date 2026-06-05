import { pgTable, serial, text, timestamp, integer, real } from "drizzle-orm/pg-core";
import { createInsertSchema } from "drizzle-zod";
import { z } from "zod/v4";
import { projectsTable } from "./projects";
import { phasesTable } from "./phases";
import { resourcesTable } from "./resources";

export const tasksTable = pgTable("tasks", {
  id: serial("id").primaryKey(),
  projectId: integer("project_id").notNull().references(() => projectsTable.id, { onDelete: "cascade" }),
  phaseId: integer("phase_id").notNull().references(() => phasesTable.id, { onDelete: "cascade" }),
  parentTaskId: integer("parent_task_id"),
  taskNumber: integer("task_number").notNull().default(0),
  taskType: text("task_type").notNull().default("task"),
  description: text("description").notNull(),
  plannedStart: text("planned_start"),
  plannedEnd: text("planned_end"),
  resourceId: integer("resource_id").references(() => resourcesTable.id, { onDelete: "set null" }),
  pctDone: real("pct_done").notNull().default(0),
  status: text("status").notNull().default("Not Started"),
  notes: text("notes"),
  sortOrder: integer("sort_order").notNull().default(0),
  createdAt: timestamp("created_at", { withTimezone: true }).notNull().defaultNow(),
  updatedAt: timestamp("updated_at", { withTimezone: true }).notNull().defaultNow().$onUpdate(() => new Date()),
});

export const insertTaskSchema = createInsertSchema(tasksTable).omit({ id: true, createdAt: true, updatedAt: true });
export type InsertTask = z.infer<typeof insertTaskSchema>;
export type Task = typeof tasksTable.$inferSelect;
