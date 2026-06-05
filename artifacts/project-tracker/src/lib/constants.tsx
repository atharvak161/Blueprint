import { AlertCircle, Clock, CheckCircle2, PlayCircle, XCircle } from "lucide-react";
import React from "react";

export const STATUS_COLORS = {
  "Not Started": "bg-slate-100 text-slate-700 border-slate-200 dark:bg-slate-800 dark:text-slate-300 dark:border-slate-700",
  "In Progress": "bg-blue-100 text-blue-700 border-blue-200 dark:bg-blue-900/30 dark:text-blue-400 dark:border-blue-800",
  "Completed": "bg-emerald-100 text-emerald-700 border-emerald-200 dark:bg-emerald-900/30 dark:text-emerald-400 dark:border-emerald-800",
  "Overdue": "bg-rose-100 text-rose-700 border-rose-200 dark:bg-rose-900/30 dark:text-rose-400 dark:border-rose-800",
  "Blocked": "bg-amber-100 text-amber-700 border-amber-200 dark:bg-amber-900/30 dark:text-amber-400 dark:border-amber-800",
};

export const STATUS_CHART_COLORS = {
  "Not Started": "hsl(215 15% 65%)",
  "In Progress": "hsl(221 83% 53%)",
  "Completed": "hsl(160 84% 39%)",
  "Overdue": "hsl(348 83% 47%)",
  "Blocked": "hsl(38 92% 50%)",
};

export const STATUS_ICONS = {
  "Not Started": Clock,
  "In Progress": PlayCircle,
  "Completed": CheckCircle2,
  "Overdue": AlertCircle,
  "Blocked": XCircle,
};

export function StatusBadge({ status }: { status: string }) {
  const colorClass = STATUS_COLORS[status as keyof typeof STATUS_COLORS] || STATUS_COLORS["Not Started"];
  const Icon = STATUS_ICONS[status as keyof typeof STATUS_ICONS] || STATUS_ICONS["Not Started"];
  
  return (
    <span className={`inline-flex items-center gap-1.5 px-2.5 py-0.5 rounded-full text-xs font-medium border ${colorClass}`}>
      <Icon className="w-3.5 h-3.5" />
      {status}
    </span>
  );
}

export function formatCompactDate(dateStr?: string | null) {
  if (!dateStr) return "-";
  const d = new Date(dateStr);
  return d.toLocaleDateString('en-GB', { day: 'numeric', month: 'short' });
}
