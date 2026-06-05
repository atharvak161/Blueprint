import React, { useState, useCallback, useRef, useEffect } from "react";
import { useParams, Link } from "wouter";
import {
  ChevronDown, ChevronRight, Plus, Trash2, Pencil, Check, X,
  AlertCircle, Clock, TrendingUp, RefreshCw
} from "lucide-react";
import {
  useGetProject,
  useGetProjectSummary,
  useListPhases,
  useListTasks,
  useListResources,
  useUpdateProject,
  useCreatePhase,
  useUpdatePhase,
  useDeletePhase,
  useCreateTask,
  useUpdateTask,
  useDeleteTask,
  getGetProjectQueryKey,
  getGetProjectSummaryQueryKey,
  getListPhasesQueryKey,
  getListTasksQueryKey,
} from "@workspace/api-client-react";
import { useQueryClient } from "@tanstack/react-query";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Textarea } from "@/components/ui/textarea";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { Dialog, DialogContent, DialogHeader, DialogTitle } from "@/components/ui/dialog";
import { AlertDialog, AlertDialogAction, AlertDialogCancel, AlertDialogContent, AlertDialogDescription, AlertDialogFooter, AlertDialogHeader, AlertDialogTitle, AlertDialogTrigger } from "@/components/ui/alert-dialog";
import { Skeleton } from "@/components/ui/skeleton";
import { Progress } from "@/components/ui/progress";
import { Badge } from "@/components/ui/badge";
import { StatusBadge, STATUS_CHART_COLORS, formatCompactDate } from "@/lib/constants";
import { BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, Cell, CartesianGrid } from "recharts";
import { useForm } from "react-hook-form";

const STATUSES = ["Not Started", "In Progress", "Completed", "Overdue", "Blocked"] as const;
const PHASE_COLORS = ["#6366f1", "#8b5cf6", "#ec4899", "#f59e0b", "#10b981", "#3b82f6", "#ef4444", "#14b8a6", "#f97316", "#84cc16", "#06b6d4", "#a78bfa"];

type Task = {
  id: number; projectId: number; phaseId: number; parentTaskId: number | null;
  taskNumber: number; taskType: string; description: string;
  plannedStart?: string | null; plannedEnd?: string | null;
  resourceId?: number | null; resourceName?: string | null;
  pctDone: number; status: string; notes?: string | null; sortOrder: number;
  createdAt: string; updatedAt: string;
};

interface TaskFormData {
  description: string; taskType: string; plannedStart: string; plannedEnd: string;
  resourceId: string; pctDone: string; status: string; notes: string;
}

async function renumberTasks(projectId: number) {
  await fetch(`/api/projects/${projectId}/tasks/renumber`, { method: "POST" });
}

// ─── Inline Description Edit ──────────────────────────────────────────────────

function InlineDescEdit({ value, onSave, className = "" }: {
  value: string;
  onSave: (val: string) => void;
  className?: string;
}) {
  const [editing, setEditing] = useState(false);
  const [val, setVal] = useState(value);
  const inputRef = useRef<HTMLInputElement>(null);

  useEffect(() => { if (editing) inputRef.current?.focus(); }, [editing]);
  useEffect(() => { setVal(value); }, [value]);

  function save() {
    const trimmed = val.trim();
    if (trimmed && trimmed !== value) onSave(trimmed);
    setEditing(false);
  }

  if (editing) {
    return (
      <div className="flex items-center gap-1 flex-1 min-w-0">
        <Input
          ref={inputRef}
          value={val}
          onChange={(e) => setVal(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === "Enter") { e.preventDefault(); save(); }
            if (e.key === "Escape") { setVal(value); setEditing(false); }
          }}
          onBlur={save}
          className="h-7 text-sm py-0 px-1.5"
        />
      </div>
    );
  }

  return (
    <span
      className={`cursor-text hover:bg-muted/50 rounded px-0.5 -mx-0.5 transition-colors group/desc ${className}`}
      title="Click to edit"
      onClick={() => setEditing(true)}
    >
      {value}
      <Pencil className="w-2.5 h-2.5 inline ml-1 opacity-0 group-hover/desc:opacity-40 transition-opacity" />
    </span>
  );
}

// ─── Inline Status Select ─────────────────────────────────────────────────────

function InlineStatusSelect({ task, projectId }: { task: Task; projectId: number }) {
  const updateTask = useUpdateTask();
  const queryClient = useQueryClient();
  return (
    <Select
      value={task.status}
      onValueChange={(val) =>
        updateTask.mutate(
          { id: task.id, data: { status: val } },
          { onSuccess: () => { queryClient.invalidateQueries({ queryKey: getListTasksQueryKey(projectId) }); queryClient.invalidateQueries({ queryKey: getGetProjectSummaryQueryKey(projectId) }); } }
        )
      }
    >
      <SelectTrigger className="h-7 w-auto border-0 bg-transparent p-0 focus:ring-0 text-xs" data-testid={`select-status-${task.id}`}>
        <StatusBadge status={task.status} />
      </SelectTrigger>
      <SelectContent>
        {STATUSES.map((s) => (
          <SelectItem key={s} value={s}>
            <StatusBadge status={s} />
          </SelectItem>
        ))}
      </SelectContent>
    </Select>
  );
}

// ─── Inline % Done ────────────────────────────────────────────────────────────

function InlinePctEdit({ task, projectId }: { task: Task; projectId: number }) {
  const [editing, setEditing] = useState(false);
  const [val, setVal] = useState(String(Math.round(task.pctDone)));
  const updateTask = useUpdateTask();
  const queryClient = useQueryClient();

  function save() {
    const n = Math.min(100, Math.max(0, parseInt(val) || 0));
    updateTask.mutate(
      { id: task.id, data: { pctDone: n } },
      { onSuccess: () => { setEditing(false); queryClient.invalidateQueries({ queryKey: getListTasksQueryKey(projectId) }); queryClient.invalidateQueries({ queryKey: getGetProjectSummaryQueryKey(projectId) }); } }
    );
  }

  if (editing) return (
    <div className="flex items-center gap-1">
      <Input value={val} onChange={(e) => setVal(e.target.value)}
        onKeyDown={(e) => { if (e.key === "Enter") save(); if (e.key === "Escape") setEditing(false); }}
        onBlur={save}
        className="h-6 w-14 text-xs p-1" autoFocus />
    </div>
  );
  return (
    <button onClick={() => setEditing(true)} className="text-xs tabular-nums hover:underline hover:text-primary" data-testid={`pct-${task.id}`}>
      {Math.round(task.pctDone)}%
    </button>
  );
}

// ─── Task Row ─────────────────────────────────────────────────────────────────

function TaskRow({ task, depth, projectId, resources, onEdit }: {
  task: Task; depth: number; projectId: number;
  resources: { id: number; name: string }[];
  onEdit: (t: Task) => void;
}) {
  const deleteTask = useDeleteTask();
  const updateTask = useUpdateTask();
  const queryClient = useQueryClient();

  const invalidate = useCallback(() => {
    queryClient.invalidateQueries({ queryKey: getListTasksQueryKey(projectId) });
    queryClient.invalidateQueries({ queryKey: getGetProjectSummaryQueryKey(projectId) });
  }, [queryClient, projectId]);

  function handleDescSave(val: string) {
    updateTask.mutate({ id: task.id, data: { description: val } }, { onSuccess: invalidate });
  }

  return (
    <tr className={`border-b hover:bg-muted/30 transition-colors group ${depth > 0 ? "bg-muted/10" : ""}`} data-testid={`row-task-${task.id}`}>
      <td className="px-3 py-2 text-xs text-muted-foreground tabular-nums font-mono">
        T-{String(task.taskNumber).padStart(3, "0")}
      </td>
      <td className="px-3 py-2">
        <div style={{ paddingLeft: depth * 16 }} className="flex items-center gap-1.5 min-w-0">
          {depth > 0 && <span className="text-muted-foreground text-xs flex-shrink-0">↳</span>}
          <InlineDescEdit value={task.description} onSave={handleDescSave} className="text-sm truncate" />
          {task.taskType === "subtask" && <Badge variant="outline" className="text-xs py-0 h-4 flex-shrink-0">sub</Badge>}
        </div>
      </td>
      <td className="px-3 py-2 whitespace-nowrap text-xs text-muted-foreground">{formatCompactDate(task.plannedStart)}</td>
      <td className="px-3 py-2 whitespace-nowrap text-xs text-muted-foreground">{formatCompactDate(task.plannedEnd)}</td>
      <td className="px-3 py-2 text-xs text-muted-foreground">{task.resourceName ?? "—"}</td>
      <td className="px-3 py-2"><InlinePctEdit task={task} projectId={projectId} /></td>
      <td className="px-3 py-2"><InlineStatusSelect task={task} projectId={projectId} /></td>
      <td className="px-3 py-2">
        <div className="flex items-center gap-1 opacity-0 group-hover:opacity-100 transition-opacity">
          <Button size="icon" variant="ghost" className="h-6 w-6 text-muted-foreground hover:text-foreground" onClick={() => onEdit(task)} data-testid={`button-edit-task-${task.id}`}>
            <Pencil className="w-3 h-3" />
          </Button>
          <AlertDialog>
            <AlertDialogTrigger asChild>
              <Button size="icon" variant="ghost" className="h-6 w-6 text-muted-foreground hover:text-destructive" data-testid={`button-delete-task-${task.id}`}>
                <Trash2 className="w-3 h-3" />
              </Button>
            </AlertDialogTrigger>
            <AlertDialogContent>
              <AlertDialogHeader>
                <AlertDialogTitle>Delete task?</AlertDialogTitle>
                <AlertDialogDescription>Delete "{task.description}"? Task IDs will be renumbered automatically.</AlertDialogDescription>
              </AlertDialogHeader>
              <AlertDialogFooter>
                <AlertDialogCancel>Cancel</AlertDialogCancel>
                <AlertDialogAction
                  className="bg-destructive text-destructive-foreground"
                  onClick={() => deleteTask.mutate({ id: task.id }, { onSuccess: async () => { await renumberTasks(projectId); invalidate(); } })}
                >
                  Delete
                </AlertDialogAction>
              </AlertDialogFooter>
            </AlertDialogContent>
          </AlertDialog>
        </div>
      </td>
    </tr>
  );
}

// ─── Phase Section ────────────────────────────────────────────────────────────

function PhaseSection({ phase, tasks, projectId, resources, onAddTask, onEditTask }: {
  phase: { id: number; name: string; color?: string | null; collapsed: boolean; projectId: number };
  tasks: Task[]; projectId: number;
  resources: { id: number; name: string }[];
  onAddTask: (phaseId: number, parentTaskId?: number) => void;
  onEditTask: (t: Task) => void;
}) {
  const updatePhase = useUpdatePhase();
  const deletePhase = useDeletePhase();
  const queryClient = useQueryClient();
  const [editingName, setEditingName] = useState(false);
  const [nameVal, setNameVal] = useState(phase.name);

  const phaseTasks = tasks.filter((t) => t.phaseId === phase.id && !t.parentTaskId);
  const total = tasks.filter((t) => t.phaseId === phase.id).length;
  const completed = tasks.filter((t) => t.phaseId === phase.id && t.status === "Completed").length;

  function toggleCollapse() {
    updatePhase.mutate(
      { id: phase.id, data: { collapsed: !phase.collapsed } },
      { onSuccess: () => queryClient.invalidateQueries({ queryKey: getListPhasesQueryKey(projectId) }) }
    );
  }

  function saveName() {
    if (!nameVal.trim() || nameVal === phase.name) { setEditingName(false); return; }
    updatePhase.mutate(
      { id: phase.id, data: { name: nameVal.trim() } },
      { onSuccess: () => { setEditingName(false); queryClient.invalidateQueries({ queryKey: getListPhasesQueryKey(projectId) }); } }
    );
  }

  return (
    <div className="border rounded-lg overflow-hidden mb-3" data-testid={`phase-${phase.id}`}>
      <div
        className="flex items-center gap-2 px-3 py-2 select-none"
        style={{ backgroundColor: phase.color ? `${phase.color}18` : undefined, borderLeft: `3px solid ${phase.color || "#6366f1"}` }}
      >
        <button onClick={toggleCollapse} className="text-muted-foreground hover:text-foreground flex-shrink-0">
          {phase.collapsed ? <ChevronRight className="w-4 h-4" /> : <ChevronDown className="w-4 h-4" />}
        </button>

        {editingName ? (
          <div className="flex items-center gap-1 flex-1" onClick={(e) => e.stopPropagation()}>
            <Input
              value={nameVal}
              onChange={(e) => setNameVal(e.target.value)}
              onKeyDown={(e) => { if (e.key === "Enter") saveName(); if (e.key === "Escape") { setNameVal(phase.name); setEditingName(false); } }}
              onBlur={saveName}
              className="h-7 text-sm font-semibold"
              autoFocus
            />
          </div>
        ) : (
          <span className="font-semibold text-sm flex-1 flex items-center gap-2 cursor-pointer" onClick={toggleCollapse}>
            {phase.name}
            <span className="text-xs text-muted-foreground font-normal">{completed}/{total} tasks</span>
          </span>
        )}

        <div className="flex items-center gap-1 ml-auto flex-shrink-0">
          <Button size="icon" variant="ghost" className="h-6 w-6 text-muted-foreground" onClick={(e) => { e.stopPropagation(); setEditingName(true); }} data-testid={`button-edit-phase-${phase.id}`}>
            <Pencil className="w-3 h-3" />
          </Button>
          <Button size="sm" variant="ghost" className="h-6 text-xs text-muted-foreground hover:text-foreground" onClick={(e) => { e.stopPropagation(); onAddTask(phase.id); }} data-testid={`button-add-task-${phase.id}`}>
            <Plus className="w-3 h-3 mr-1" /> Task
          </Button>
          <AlertDialog>
            <AlertDialogTrigger asChild>
              <Button size="icon" variant="ghost" className="h-6 w-6 text-muted-foreground hover:text-destructive" onClick={(e) => e.stopPropagation()} data-testid={`button-delete-phase-${phase.id}`}>
                <Trash2 className="w-3 h-3" />
              </Button>
            </AlertDialogTrigger>
            <AlertDialogContent>
              <AlertDialogHeader><AlertDialogTitle>Delete phase?</AlertDialogTitle><AlertDialogDescription>Delete "{phase.name}" and all its tasks? This cannot be undone.</AlertDialogDescription></AlertDialogHeader>
              <AlertDialogFooter>
                <AlertDialogCancel>Cancel</AlertDialogCancel>
                <AlertDialogAction
                  className="bg-destructive text-destructive-foreground"
                  onClick={() => deletePhase.mutate({ id: phase.id }, { onSuccess: () => { queryClient.invalidateQueries({ queryKey: getListPhasesQueryKey(projectId) }); queryClient.invalidateQueries({ queryKey: getListTasksQueryKey(projectId) }); } })}
                >
                  Delete
                </AlertDialogAction>
              </AlertDialogFooter>
            </AlertDialogContent>
          </AlertDialog>
        </div>
      </div>

      {!phase.collapsed && (
        <table className="w-full">
          <colgroup>
            <col className="w-20" /><col /><col className="w-24" /><col className="w-24" />
            <col className="w-28" /><col className="w-16" /><col className="w-36" /><col className="w-16" />
          </colgroup>
          <thead>
            <tr className="border-b bg-muted/30 text-xs text-muted-foreground">
              <th className="px-3 py-1.5 text-left font-medium">ID</th>
              <th className="px-3 py-1.5 text-left font-medium">Description</th>
              <th className="px-3 py-1.5 text-left font-medium">Start</th>
              <th className="px-3 py-1.5 text-left font-medium">End</th>
              <th className="px-3 py-1.5 text-left font-medium">Resource</th>
              <th className="px-3 py-1.5 text-left font-medium">% Done</th>
              <th className="px-3 py-1.5 text-left font-medium">Status</th>
              <th className="px-3 py-1.5" />
            </tr>
          </thead>
          <tbody>
            {phaseTasks.map((task) => (
              <React.Fragment key={task.id}>
                <TaskRow task={task} depth={0} projectId={projectId} resources={resources} onEdit={onEditTask} />
                {tasks.filter((t) => t.parentTaskId === task.id).map((sub) => (
                  <TaskRow key={sub.id} task={sub} depth={1} projectId={projectId} resources={resources} onEdit={onEditTask} />
                ))}
                <tr>
                  <td colSpan={8} className="px-3 py-1">
                    <button
                      onClick={() => onAddTask(phase.id, task.id)}
                      className="text-xs text-muted-foreground hover:text-foreground flex items-center gap-1 pl-6"
                      data-testid={`button-add-subtask-${task.id}`}
                    >
                      <Plus className="w-3 h-3" /> Add subtask
                    </button>
                  </td>
                </tr>
              </React.Fragment>
            ))}
            {phaseTasks.length === 0 && (
              <tr>
                <td colSpan={8} className="px-3 py-4 text-center text-xs text-muted-foreground">No tasks yet — click "+ Task" to add one</td>
              </tr>
            )}
          </tbody>
        </table>
      )}
    </div>
  );
}

// ─── Gantt View ────────────────────────────────────────────────────────────────

function GanttView({ tasks, phases }: { tasks: Task[]; phases: { id: number; name: string; color?: string | null }[] }) {
  const withDates = tasks.filter((t) => t.plannedStart && t.plannedEnd);
  if (withDates.length === 0) {
    return (
      <div className="flex flex-col items-center justify-center h-48 text-muted-foreground text-sm gap-2">
        <RefreshCw className="w-8 h-8 opacity-20" />
        <span>No tasks with planned dates yet — add start/end dates to tasks to see the Gantt chart.</span>
      </div>
    );
  }

  const allDates = withDates.flatMap((t) => [new Date(t.plannedStart!).getTime(), new Date(t.plannedEnd!).getTime()]);
  const minDate = Math.min(...allDates);
  const maxDate = Math.max(...allDates);
  const totalMs = maxDate - minDate || 1;

  const monthMs = 30 * 24 * 60 * 60 * 1000;
  const months: { label: string; left: number; width: number }[] = [];
  const startMonth = new Date(minDate);
  startMonth.setDate(1);
  let cur = startMonth.getTime();
  while (cur <= maxDate) {
    const d = new Date(cur);
    const next = new Date(d.getFullYear(), d.getMonth() + 1, 1).getTime();
    const left = Math.max(0, ((cur - minDate) / totalMs) * 100);
    const right = Math.min(100, ((next - minDate) / totalMs) * 100);
    months.push({ label: d.toLocaleString("en-GB", { month: "short", year: "2-digit" }), left, width: right - left });
    cur = next;
  }

  return (
    <div className="overflow-x-auto">
      <div className="min-w-[700px]">
        <div className="flex mb-1 pl-44">
          {months.map((m) => (
            <div key={m.label} className="text-[10px] text-muted-foreground text-center border-r last:border-0 px-0.5" style={{ width: `${m.width}%` }}>
              {m.label}
            </div>
          ))}
        </div>

        {phases.map((phase) => {
          const phaseTasks = withDates.filter((t) => t.phaseId === phase.id);
          if (phaseTasks.length === 0) return null;
          return (
            <div key={phase.id} className="mb-4">
              <div className="text-xs font-semibold text-muted-foreground uppercase tracking-wide mb-1 pl-44" style={{ color: phase.color ?? undefined }}>
                {phase.name}
              </div>
              {phaseTasks.map((task) => {
                const start = new Date(task.plannedStart!).getTime();
                const end = new Date(task.plannedEnd!).getTime();
                const left = ((start - minDate) / totalMs) * 100;
                const width = Math.max(((end - start) / totalMs) * 100, 0.5);
                const color = STATUS_CHART_COLORS[task.status as keyof typeof STATUS_CHART_COLORS] || STATUS_CHART_COLORS["Not Started"];
                return (
                  <div key={task.id} className="flex items-center gap-2 py-0.5" data-testid={`gantt-row-${task.id}`}>
                    <div className="w-40 flex-shrink-0 text-xs text-right text-muted-foreground truncate pr-2" title={`T-${String(task.taskNumber).padStart(3,"0")} ${task.description}`}>
                      <span className="font-mono text-[10px] text-muted-foreground/60 mr-1">T-{String(task.taskNumber).padStart(3,"0")}</span>
                      {task.description}
                    </div>
                    <div className="flex-1 relative h-6 bg-muted/20 rounded border border-muted/40">
                      <div
                        className="absolute top-1 h-4 rounded text-[10px] flex items-center px-1.5 text-white font-medium truncate"
                        style={{ left: `${left}%`, width: `${width}%`, backgroundColor: color, minWidth: 4 }}
                        title={`${task.plannedStart} → ${task.plannedEnd} | ${Math.round(task.pctDone)}% done`}
                      >
                        {width > 4 ? Math.round(task.pctDone) + "%" : ""}
                      </div>
                    </div>
                  </div>
                );
              })}
            </div>
          );
        })}

        <div className="flex gap-4 mt-4 flex-wrap pl-44 text-xs pt-2 border-t">
          {Object.entries(STATUS_CHART_COLORS).map(([status, color]) => (
            <span key={status} className="flex items-center gap-1.5">
              <span className="w-3 h-3 rounded" style={{ backgroundColor: color }} />
              {status}
            </span>
          ))}
        </div>
      </div>
    </div>
  );
}

// ─── Dashboard View ────────────────────────────────────────────────────────────

function DashboardView({ projectId }: { projectId: number }) {
  const { data: summary, isLoading } = useGetProjectSummary(projectId, { query: { enabled: true, queryKey: getGetProjectSummaryQueryKey(projectId) } });
  if (isLoading) return <div className="space-y-4"><Skeleton className="h-20 w-full" /><Skeleton className="h-40 w-full" /></div>;
  if (!summary) return null;

  const kpis = [
    { label: "Total Tasks",  value: summary.totalTasks,  color: "text-foreground",       bg: "bg-card" },
    { label: "Completed",    value: summary.completed,   color: "text-emerald-600",      bg: "bg-emerald-50 dark:bg-emerald-950/20" },
    { label: "In Progress",  value: summary.inProgress,  color: "text-blue-600",         bg: "bg-blue-50 dark:bg-blue-950/20" },
    { label: "Overdue",      value: summary.overdue,     color: "text-rose-600",         bg: "bg-rose-50 dark:bg-rose-950/20" },
    { label: "Blocked",      value: summary.blocked,     color: "text-amber-600",        bg: "bg-amber-50 dark:bg-amber-950/20" },
    { label: "Not Started",  value: summary.notStarted,  color: "text-muted-foreground", bg: "bg-card" },
  ];

  const statusBreakdown = [
    { name: "Completed",   value: summary.completed,  color: STATUS_CHART_COLORS.Completed },
    { name: "In Progress", value: summary.inProgress, color: STATUS_CHART_COLORS["In Progress"] },
    { name: "Not Started", value: summary.notStarted, color: STATUS_CHART_COLORS["Not Started"] },
    { name: "Overdue",     value: summary.overdue,    color: STATUS_CHART_COLORS.Overdue },
    { name: "Blocked",     value: summary.blocked,    color: STATUS_CHART_COLORS.Blocked },
  ].filter((d) => d.value > 0);

  return (
    <div className="space-y-6">
      <div>
        <div className="flex items-center justify-between mb-2">
          <span className="text-sm font-medium">Overall Progress</span>
          <span className="text-sm font-semibold tabular-nums">{Math.round(summary.pctComplete)}%</span>
        </div>
        <Progress value={summary.pctComplete} className="h-3" />
      </div>

      <div className="grid grid-cols-2 sm:grid-cols-3 xl:grid-cols-6 gap-3">
        {kpis.map((kpi) => (
          <div key={kpi.label} className={`border rounded-xl p-4 text-center ${kpi.bg}`} data-testid={`kpi-${kpi.label.replace(/ /g, "-").toLowerCase()}`}>
            <div className={`text-3xl font-bold tabular-nums ${kpi.color}`}>{kpi.value}</div>
            <div className="text-xs text-muted-foreground mt-1">{kpi.label}</div>
          </div>
        ))}
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        <div className="border rounded-xl p-4">
          <h3 className="text-sm font-semibold mb-4">Status Breakdown</h3>
          <ResponsiveContainer width="100%" height={200}>
            <BarChart data={statusBreakdown} layout="vertical" margin={{ left: 10, right: 24, bottom: 4 }}>
              <CartesianGrid strokeDasharray="3 3" horizontal={false} />
              <XAxis type="number" tick={{ fontSize: 11 }} />
              <YAxis type="category" dataKey="name" tick={{ fontSize: 11 }} width={80} />
              <Tooltip />
              <Bar dataKey="value" radius={[0, 4, 4, 0]} label={{ position: "right", fontSize: 11 }}>
                {statusBreakdown.map((entry, idx) => <Cell key={idx} fill={entry.color} />)}
              </Bar>
            </BarChart>
          </ResponsiveContainer>
        </div>
        <div className="border rounded-xl p-4">
          <h3 className="text-sm font-semibold mb-4">Progress by Phase</h3>
          {summary.phaseStats.length === 0 ? (
            <p className="text-sm text-muted-foreground">No phases yet</p>
          ) : (
            <div className="space-y-4">
              {summary.phaseStats.map((ps) => (
                <div key={ps.phaseId} data-testid={`phase-stat-${ps.phaseId}`}>
                  <div className="flex justify-between text-xs mb-1.5">
                    <span className="font-medium truncate">{ps.phaseName}</span>
                    <span className="text-muted-foreground tabular-nums ml-2">{ps.completed}/{ps.total} • {Math.round(ps.pctComplete)}%</span>
                  </div>
                  <Progress value={ps.pctComplete} className="h-2" />
                </div>
              ))}
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

// ─── RAG View ─────────────────────────────────────────────────────────────────

function RAGView({ projectId }: { projectId: number }) {
  const { data: summary } = useGetProjectSummary(projectId, { query: { enabled: true, queryKey: getGetProjectSummaryQueryKey(projectId) } });
  if (!summary) return null;

  const ragStatus = summary.overdue > 2 || summary.blocked > 2
    ? "RED"
    : summary.overdue > 0 || summary.blocked > 0
    ? "AMBER"
    : "GREEN";

  const ragConfig = {
    RED:   { label: "Red",   bg: "bg-rose-50 dark:bg-rose-950/20",     border: "border-rose-300",   text: "text-rose-700 dark:text-rose-400",     dot: "bg-rose-500",    desc: "Project has significant issues requiring immediate attention." },
    AMBER: { label: "Amber", bg: "bg-amber-50 dark:bg-amber-950/20",   border: "border-amber-300",  text: "text-amber-700 dark:text-amber-400",   dot: "bg-amber-500",   desc: "Project has some risks that need monitoring and action." },
    GREEN: { label: "Green", bg: "bg-emerald-50 dark:bg-emerald-950/20", border: "border-emerald-300", text: "text-emerald-700 dark:text-emerald-400", dot: "bg-emerald-500", desc: "Project is on track with no significant risks." },
  };

  const cfg = ragConfig[ragStatus];

  const phaseRAGs = summary.phaseStats.map((ps) => {
    const phaseRag = ps.pctComplete === 100 ? "GREEN" : ps.pctComplete === 0 && ps.total > 0 ? "AMBER" : "GREEN";
    return { ...ps, rag: phaseRag as "RED" | "AMBER" | "GREEN" };
  });

  return (
    <div className="space-y-6">
      <div className={`border-2 rounded-xl p-6 ${cfg.bg} ${cfg.border}`}>
        <div className="flex items-center gap-4">
          <span className={`w-8 h-8 rounded-full flex-shrink-0 ${cfg.dot} shadow-lg`} />
          <div>
            <div className={`text-2xl font-bold ${cfg.text}`}>Overall: {cfg.label}</div>
            <div className={`text-sm mt-0.5 ${cfg.text} opacity-80`}>{cfg.desc}</div>
          </div>
        </div>
        <div className="grid grid-cols-2 md:grid-cols-4 gap-4 mt-5">
          {[
            { label: "Total Tasks",  value: summary.totalTasks },
            { label: "Completed",    value: summary.completed },
            { label: "Overdue",      value: summary.overdue },
            { label: "Blocked",      value: summary.blocked },
          ].map((item) => (
            <div key={item.label} className="bg-white/50 dark:bg-black/20 rounded-lg p-3 text-center">
              <div className={`text-2xl font-bold tabular-nums ${cfg.text}`}>{item.value}</div>
              <div className={`text-xs mt-0.5 ${cfg.text} opacity-70`}>{item.label}</div>
            </div>
          ))}
        </div>
      </div>

      <div>
        <h3 className="text-sm font-semibold mb-3">Phase RAG Status</h3>
        <div className="space-y-2">
          {phaseRAGs.map((ps) => {
            const c = ragConfig[ps.rag];
            return (
              <div key={ps.phaseId} className={`border rounded-lg p-3 flex items-center gap-3 ${c.bg} ${c.border}`} data-testid={`rag-phase-${ps.phaseId}`}>
                <span className={`w-3 h-3 rounded-full flex-shrink-0 ${c.dot}`} />
                <div className="flex-1 min-w-0">
                  <div className="flex items-center justify-between gap-2">
                    <span className={`text-sm font-medium truncate ${c.text}`}>{ps.phaseName}</span>
                    <span className={`text-xs tabular-nums ${c.text} opacity-70 flex-shrink-0`}>{ps.completed}/{ps.total} tasks · {Math.round(ps.pctComplete)}%</span>
                  </div>
                  <Progress value={ps.pctComplete} className="h-1.5 mt-1.5" />
                </div>
              </div>
            );
          })}
        </div>
      </div>

      <div className="border rounded-xl p-4">
        <h3 className="text-sm font-semibold mb-3">Risk Summary</h3>
        <div className="space-y-2 text-sm">
          {summary.overdue > 0 && <div className="flex items-center gap-2 text-rose-600"><span className="w-2 h-2 rounded-full bg-rose-500" />{summary.overdue} overdue task{summary.overdue > 1 ? "s" : ""} require immediate attention.</div>}
          {summary.blocked > 0 && <div className="flex items-center gap-2 text-amber-600"><span className="w-2 h-2 rounded-full bg-amber-500" />{summary.blocked} blocked task{summary.blocked > 1 ? "s" : ""} need dependencies resolved.</div>}
          {summary.inProgress > 0 && <div className="flex items-center gap-2 text-blue-600"><span className="w-2 h-2 rounded-full bg-blue-500" />{summary.inProgress} task{summary.inProgress > 1 ? "s" : ""} currently in progress.</div>}
          {summary.overdue === 0 && summary.blocked === 0 && <div className="flex items-center gap-2 text-emerald-600"><span className="w-2 h-2 rounded-full bg-emerald-500" />No overdue or blocked tasks — project is on track.</div>}
        </div>
      </div>
    </div>
  );
}

// ─── Task Modal ────────────────────────────────────────────────────────────────

function TaskModal({ open, onClose, projectId, phaseId, parentTaskId, editTask, phases, resources }: {
  open: boolean; onClose: () => void; projectId: number;
  phaseId?: number; parentTaskId?: number; editTask?: Task | null;
  phases: { id: number; name: string }[];
  resources: { id: number; name: string }[];
}) {
  const createTask = useCreateTask();
  const updateTask = useUpdateTask();
  const queryClient = useQueryClient();

  const defaultPhaseId = phaseId ?? editTask?.phaseId ?? phases[0]?.id;

  const form = useForm<TaskFormData>({
    defaultValues: {
      description: editTask?.description ?? "",
      taskType: editTask?.taskType ?? (parentTaskId ? "subtask" : "task"),
      plannedStart: editTask?.plannedStart ?? "",
      plannedEnd: editTask?.plannedEnd ?? "",
      resourceId: editTask?.resourceId ? String(editTask.resourceId) : "",
      pctDone: editTask?.pctDone != null ? String(Math.round(editTask.pctDone)) : "0",
      status: editTask?.status ?? "Not Started",
      notes: editTask?.notes ?? "",
    },
  });

  useEffect(() => {
    if (open) {
      form.reset({
        description: editTask?.description ?? "",
        taskType: editTask?.taskType ?? (parentTaskId ? "subtask" : "task"),
        plannedStart: editTask?.plannedStart ?? "",
        plannedEnd: editTask?.plannedEnd ?? "",
        resourceId: editTask?.resourceId ? String(editTask.resourceId) : "",
        pctDone: editTask?.pctDone != null ? String(Math.round(editTask.pctDone)) : "0",
        status: editTask?.status ?? "Not Started",
        notes: editTask?.notes ?? "",
      });
    }
  }, [open, editTask]);

  function onSubmit(data: TaskFormData) {
    const payload = {
      phaseId: defaultPhaseId!,
      taskType: data.taskType,
      description: data.description,
      parentTaskId: parentTaskId ?? editTask?.parentTaskId ?? undefined,
      plannedStart: data.plannedStart || undefined,
      plannedEnd: data.plannedEnd || undefined,
      resourceId: data.resourceId ? parseInt(data.resourceId) : undefined,
      pctDone: parseFloat(data.pctDone) || 0,
      status: data.status,
      notes: data.notes || undefined,
    };
    const invalidate = () => {
      queryClient.invalidateQueries({ queryKey: getListTasksQueryKey(projectId) });
      queryClient.invalidateQueries({ queryKey: getGetProjectSummaryQueryKey(projectId) });
      onClose();
    };
    if (editTask) {
      updateTask.mutate({ id: editTask.id, data: payload }, { onSuccess: invalidate });
    } else {
      createTask.mutate({ projectId, data: payload }, { onSuccess: invalidate });
    }
  }

  return (
    <Dialog open={open} onOpenChange={(v) => !v && onClose()}>
      <DialogContent className="max-w-lg">
        <DialogHeader>
          <DialogTitle>{editTask ? "Edit Task" : parentTaskId ? "Add Subtask" : "Add Task"}</DialogTitle>
        </DialogHeader>
        <form onSubmit={form.handleSubmit(onSubmit)} className="space-y-4 mt-1">
          <div>
            <label className="text-sm font-medium">Description *</label>
            <Textarea {...form.register("description", { required: true })} className="mt-1" rows={2} autoFocus placeholder="Describe the task…" />
          </div>
          <div className="grid grid-cols-2 gap-3">
            <div>
              <label className="text-sm font-medium">Phase</label>
              <Select value={String(defaultPhaseId ?? "")} onValueChange={() => {}}>
                <SelectTrigger className="mt-1"><SelectValue placeholder="Phase" /></SelectTrigger>
                <SelectContent>
                  {phases.map((p) => <SelectItem key={p.id} value={String(p.id)}>{p.name}</SelectItem>)}
                </SelectContent>
              </Select>
            </div>
            <div>
              <label className="text-sm font-medium">Type</label>
              <Select value={form.watch("taskType")} onValueChange={(v) => form.setValue("taskType", v)}>
                <SelectTrigger className="mt-1"><SelectValue /></SelectTrigger>
                <SelectContent>
                  <SelectItem value="task">Task</SelectItem>
                  <SelectItem value="subtask">Subtask</SelectItem>
                  <SelectItem value="milestone">Milestone</SelectItem>
                </SelectContent>
              </Select>
            </div>
          </div>
          <div className="grid grid-cols-2 gap-3">
            <div>
              <label className="text-sm font-medium">Start Date</label>
              <Input type="date" {...form.register("plannedStart")} className="mt-1" />
            </div>
            <div>
              <label className="text-sm font-medium">End Date</label>
              <Input type="date" {...form.register("plannedEnd")} className="mt-1" />
            </div>
          </div>
          <div className="grid grid-cols-2 gap-3">
            <div>
              <label className="text-sm font-medium">Resource</label>
              <Select value={form.watch("resourceId")} onValueChange={(v) => form.setValue("resourceId", v)}>
                <SelectTrigger className="mt-1"><SelectValue placeholder="Unassigned" /></SelectTrigger>
                <SelectContent>
                  <SelectItem value="">Unassigned</SelectItem>
                  {resources.map((r) => <SelectItem key={r.id} value={String(r.id)}>{r.name}</SelectItem>)}
                </SelectContent>
              </Select>
            </div>
            <div>
              <label className="text-sm font-medium">Status</label>
              <Select value={form.watch("status")} onValueChange={(v) => form.setValue("status", v)}>
                <SelectTrigger className="mt-1"><SelectValue /></SelectTrigger>
                <SelectContent>
                  {STATUSES.map((s) => <SelectItem key={s} value={s}><StatusBadge status={s} /></SelectItem>)}
                </SelectContent>
              </Select>
            </div>
          </div>
          <div>
            <label className="text-sm font-medium">% Complete</label>
            <Input type="number" min={0} max={100} {...form.register("pctDone")} className="mt-1" />
          </div>
          <div>
            <label className="text-sm font-medium">Notes</label>
            <Textarea {...form.register("notes")} className="mt-1" rows={2} placeholder="Optional notes…" />
          </div>
          <div className="flex justify-end gap-2 pt-1">
            <Button type="button" variant="outline" onClick={onClose}>Cancel</Button>
            <Button type="submit" disabled={createTask.isPending || updateTask.isPending}>
              {editTask ? "Save Changes" : "Add Task"}
            </Button>
          </div>
        </form>
      </DialogContent>
    </Dialog>
  );
}

// ─── Project Header with Inline Edit ─────────────────────────────────────────

function ProjectHeader({ project, projectId }: {
  project: { id: number; name: string; description?: string | null; startDate?: string | null; endDate?: string | null };
  projectId: number;
}) {
  const updateProject = useUpdateProject();
  const queryClient = useQueryClient();
  const [editingName, setEditingName] = useState(false);
  const [editingDesc, setEditingDesc] = useState(false);
  const [nameVal, setNameVal] = useState(project.name);
  const [descVal, setDescVal] = useState(project.description ?? "");

  useEffect(() => { setNameVal(project.name); }, [project.name]);
  useEffect(() => { setDescVal(project.description ?? ""); }, [project.description]);

  function saveName() {
    if (!nameVal.trim() || nameVal === project.name) { setEditingName(false); return; }
    updateProject.mutate(
      { id: projectId, data: { name: nameVal.trim() } },
      { onSuccess: () => { setEditingName(false); queryClient.invalidateQueries({ queryKey: getGetProjectQueryKey(projectId) }); } }
    );
  }

  function saveDesc() {
    if (descVal === (project.description ?? "")) { setEditingDesc(false); return; }
    updateProject.mutate(
      { id: projectId, data: { description: descVal || undefined } },
      { onSuccess: () => { setEditingDesc(false); queryClient.invalidateQueries({ queryKey: getGetProjectQueryKey(projectId) }); } }
    );
  }

  return (
    <div className="flex-1 min-w-0">
      {editingName ? (
        <div className="flex items-center gap-2">
          <Input
            value={nameVal}
            onChange={(e) => setNameVal(e.target.value)}
            onKeyDown={(e) => { if (e.key === "Enter") saveName(); if (e.key === "Escape") { setNameVal(project.name); setEditingName(false); } }}
            onBlur={saveName}
            className="text-2xl font-bold h-9 max-w-lg"
            autoFocus
          />
        </div>
      ) : (
        <h1
          className="text-2xl font-bold tracking-tight truncate cursor-text hover:bg-muted/40 rounded px-0.5 -mx-0.5 inline-flex items-center gap-1.5 group/name"
          onClick={() => setEditingName(true)}
          title="Click to edit project name"
        >
          {project.name}
          <Pencil className="w-3.5 h-3.5 opacity-0 group-hover/name:opacity-40 transition-opacity" />
        </h1>
      )}

      {editingDesc ? (
        <div className="flex items-center gap-2 mt-1">
          <Input
            value={descVal}
            onChange={(e) => setDescVal(e.target.value)}
            onKeyDown={(e) => { if (e.key === "Enter") saveDesc(); if (e.key === "Escape") { setDescVal(project.description ?? ""); setEditingDesc(false); } }}
            onBlur={saveDesc}
            className="text-sm h-7 max-w-lg"
            placeholder="Add project description…"
            autoFocus
          />
        </div>
      ) : (
        <p
          className="text-sm text-muted-foreground mt-0.5 cursor-text hover:text-foreground/70 inline-flex items-center gap-1 group/desc"
          onClick={() => setEditingDesc(true)}
          title="Click to edit description"
        >
          {project.description || <span className="italic">Click to add description…</span>}
          <Pencil className="w-3 h-3 opacity-0 group-hover/desc:opacity-40 transition-opacity" />
        </p>
      )}
    </div>
  );
}

// ─── Main Page ────────────────────────────────────────────────────────────────

export default function ProjectDetailPage() {
  const params = useParams<{ id: string; section?: string }>();
  const projectId = parseInt(params.id, 10);
  const section = params.section ?? "plan";
  const queryClient = useQueryClient();

  const { data: project, isLoading: projectLoading } = useGetProject(projectId, { query: { enabled: !!projectId } });
  const { data: phases = [], isLoading: phasesLoading } = useListPhases(projectId, { query: { enabled: !!projectId } });
  const { data: tasks = [], isLoading: tasksLoading } = useListTasks(projectId, { query: { enabled: !!projectId } });
  const { data: resources = [] } = useListResources();
  const createPhase = useCreatePhase();

  const [taskModal, setTaskModal] = useState<{ open: boolean; phaseId?: number; parentTaskId?: number; editTask?: Task | null }>({ open: false });
  const [newPhaseName, setNewPhaseName] = useState("");
  const [addingPhase, setAddingPhase] = useState(false);

  function openAddTask(phaseId: number, parentTaskId?: number) {
    setTaskModal({ open: true, phaseId, parentTaskId });
  }

  function openEditTask(task: Task) {
    setTaskModal({ open: true, editTask: task });
  }

  function addPhase() {
    if (!newPhaseName.trim()) return;
    createPhase.mutate(
      {
        projectId,
        data: {
          name: newPhaseName.trim(),
          color: PHASE_COLORS[phases.length % PHASE_COLORS.length],
          orderIndex: phases.length,
        },
      },
      {
        onSuccess: () => {
          queryClient.invalidateQueries({ queryKey: getListPhasesQueryKey(projectId) });
          setNewPhaseName("");
          setAddingPhase(false);
        },
      }
    );
  }

  if (projectLoading) {
    return (
      <div className="p-6 space-y-4">
        <Skeleton className="h-8 w-64" />
        <Skeleton className="h-5 w-96" />
        <Skeleton className="h-48 w-full" />
      </div>
    );
  }

  if (!project) {
    return (
      <div className="p-6 text-center">
        <AlertCircle className="w-10 h-10 mx-auto text-muted-foreground mb-3" />
        <p className="text-muted-foreground">Project not found</p>
        <Link href="/" className="text-primary text-sm mt-2 inline-block hover:underline">Back to projects</Link>
      </div>
    );
  }

  return (
    <div className="p-6 max-w-7xl mx-auto">
      <div className="flex items-start gap-3 mb-6">
        <ProjectHeader project={project} projectId={projectId} />
        {(project.startDate || project.endDate) && (
          <div className="text-sm text-muted-foreground whitespace-nowrap pt-1.5 flex-shrink-0">
            {formatCompactDate(project.startDate)} – {formatCompactDate(project.endDate)}
          </div>
        )}
      </div>

      {section === "plan" && (
        <div>
          {phasesLoading || tasksLoading ? (
            <div className="space-y-3">{[1, 2, 3].map((i) => <Skeleton key={i} className="h-24 w-full" />)}</div>
          ) : (
            <>
              {phases.map((phase) => (
                <PhaseSection
                  key={phase.id}
                  phase={phase}
                  tasks={tasks}
                  projectId={projectId}
                  resources={resources}
                  onAddTask={openAddTask}
                  onEditTask={openEditTask}
                />
              ))}
              {phases.length === 0 && (
                <div className="text-center py-16 border-2 border-dashed rounded-xl text-sm text-muted-foreground">
                  No phases yet — add your first phase below to get started
                </div>
              )}
              <div className="mt-4">
                {addingPhase ? (
                  <div className="flex items-center gap-2">
                    <Input
                      value={newPhaseName}
                      onChange={(e) => setNewPhaseName(e.target.value)}
                      onKeyDown={(e) => { if (e.key === "Enter") addPhase(); if (e.key === "Escape") { setAddingPhase(false); setNewPhaseName(""); } }}
                      placeholder="Phase name…"
                      className="max-w-xs"
                      autoFocus
                      data-testid="input-phase-name"
                    />
                    <Button onClick={addPhase} disabled={createPhase.isPending} data-testid="button-add-phase-confirm">Add Phase</Button>
                    <Button variant="outline" onClick={() => { setAddingPhase(false); setNewPhaseName(""); }}>Cancel</Button>
                  </div>
                ) : (
                  <Button variant="outline" onClick={() => setAddingPhase(true)} data-testid="button-add-phase">
                    <Plus className="w-4 h-4 mr-2" /> Add Phase
                  </Button>
                )}
              </div>
            </>
          )}
        </div>
      )}

      {section === "gantt" && (
        <div>
          {tasksLoading ? (
            <Skeleton className="h-48 w-full" />
          ) : (
            <GanttView tasks={tasks} phases={phases} />
          )}
        </div>
      )}

      {section === "dashboard" && (
        <DashboardView projectId={projectId} />
      )}

      {section === "rag" && (
        <RAGView projectId={projectId} />
      )}

      <TaskModal
        open={taskModal.open}
        onClose={() => setTaskModal({ open: false })}
        projectId={projectId}
        phaseId={taskModal.phaseId}
        parentTaskId={taskModal.parentTaskId}
        editTask={taskModal.editTask}
        phases={phases}
        resources={resources}
      />
    </div>
  );
}
