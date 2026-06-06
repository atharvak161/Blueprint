// local-db.ts

const KEYS = {
  projects: 'eph_projects',
  phases: 'eph_phases',
  tasks: 'eph_tasks',
  resources: 'eph_resources',
  holidays: 'eph_holidays',
  ids: 'eph_ids',
};

function load<T>(key: string): T[] {
  try { return JSON.parse(localStorage.getItem(key) || '[]'); } catch { return []; }
}
function save(key: string, data: unknown[]) {
  localStorage.setItem(key, JSON.stringify(data));
}
function nextId(entity: string): number {
  const ids = JSON.parse(localStorage.getItem(KEYS.ids) || '{}');
  const next = (ids[entity] || 0) + 1;
  ids[entity] = next;
  localStorage.setItem(KEYS.ids, JSON.stringify(ids));
  return next;
}
function now() { return new Date().toISOString(); }

// ── Projects ──────────────────────────────────────────────────────────────────
export function listProjects() { return load(KEYS.projects); }
export function createProject(input: { name: string; description?: string; startDate?: string; endDate?: string }) {
  const projects = load<any>(KEYS.projects);
  const p = { id: nextId('projects'), ...input, createdAt: now(), updatedAt: now() };
  projects.push(p); save(KEYS.projects, projects); return p;
}
export function getProject(id: number) { return load<any>(KEYS.projects).find((p: any) => p.id === id) || null; }
export function updateProject(id: number, patch: Record<string, unknown>) {
  const projects = load<any>(KEYS.projects);
  const idx = projects.findIndex((p: any) => p.id === id);
  if (idx === -1) return null;
  projects[idx] = { ...projects[idx], ...patch, updatedAt: now() };
  save(KEYS.projects, projects); return projects[idx];
}
export function deleteProject(id: number) {
  // cascade: delete phases and tasks
  const phases = load<any>(KEYS.phases).filter((p: any) => p.projectId !== id);
  save(KEYS.phases, phases);
  const tasks = load<any>(KEYS.tasks).filter((t: any) => t.projectId !== id);
  save(KEYS.tasks, tasks);
  const projects = load<any>(KEYS.projects).filter((p: any) => p.id !== id);
  save(KEYS.projects, projects);
}

// ── Phases ────────────────────────────────────────────────────────────────────
export function listPhases(projectId: number) {
  return load<any>(KEYS.phases).filter((p: any) => p.projectId === projectId)
    .sort((a: any, b: any) => a.orderIndex - b.orderIndex);
}
export function createPhase(projectId: number, input: { name: string; color?: string; orderIndex?: number }) {
  const phases = load<any>(KEYS.phases);
  const existing = phases.filter((p: any) => p.projectId === projectId);
  const ph = {
    id: nextId('phases'), projectId, ...input,
    orderIndex: input.orderIndex ?? existing.length,
    collapsed: false, createdAt: now(),
  };
  phases.push(ph); save(KEYS.phases, phases); return ph;
}
export function updatePhase(id: number, patch: Record<string, unknown>) {
  const phases = load<any>(KEYS.phases);
  const idx = phases.findIndex((p: any) => p.id === id);
  if (idx === -1) return null;
  phases[idx] = { ...phases[idx], ...patch };
  save(KEYS.phases, phases); return phases[idx];
}
export function deletePhase(id: number) {
  const tasks = load<any>(KEYS.tasks).filter((t: any) => t.phaseId !== id);
  save(KEYS.tasks, tasks);
  const phases = load<any>(KEYS.phases).filter((p: any) => p.id !== id);
  save(KEYS.phases, phases);
}

// ── Tasks ─────────────────────────────────────────────────────────────────────
export function listTasks(projectId: number) {
  const tasks = load<any>(KEYS.tasks).filter((t: any) => t.projectId === projectId);
  const resources = load<any>(KEYS.resources);
  return tasks.map((t: any) => {
    const r = resources.find((r: any) => r.id === t.resourceId);
    return { ...t, resourceName: r ? r.name : null };
  }).sort((a: any, b: any) => a.sortOrder - b.sortOrder);
}
export function createTask(projectId: number, input: {
  phaseId: number; taskType: string; description: string;
  parentTaskId?: number; plannedStart?: string; plannedEnd?: string;
  resourceId?: number; pctDone?: number; status?: string; notes?: string; sortOrder?: number;
}) {
  const tasks = load<any>(KEYS.tasks);
  const projectTasks = tasks.filter((t: any) => t.projectId === projectId);
  const maxNum = projectTasks.reduce((m: number, t: any) => Math.max(m, t.taskNumber), 0);
  const task = {
    id: nextId('tasks'), projectId,
    taskNumber: maxNum + 1,
    pctDone: 0, status: 'Not Started',
    sortOrder: maxNum + 1,
    ...input,
    parentTaskId: input.parentTaskId ?? null,
    plannedStart: input.plannedStart ?? null,
    plannedEnd: input.plannedEnd ?? null,
    resourceId: input.resourceId ?? null,
    notes: input.notes ?? null,
    createdAt: now(), updatedAt: now(),
  };
  tasks.push(task); save(KEYS.tasks, tasks);
  const resources = load<any>(KEYS.resources);
  const r = resources.find((r: any) => r.id === task.resourceId);
  return { ...task, resourceName: r ? r.name : null };
}
export function updateTask(id: number, patch: Record<string, unknown>) {
  const tasks = load<any>(KEYS.tasks);
  const idx = tasks.findIndex((t: any) => t.id === id);
  if (idx === -1) return null;
  tasks[idx] = { ...tasks[idx], ...patch, updatedAt: now() };
  save(KEYS.tasks, tasks);
  const resources = load<any>(KEYS.resources);
  const r = resources.find((r: any) => r.id === tasks[idx].resourceId);
  return { ...tasks[idx], resourceName: r ? r.name : null };
}
export function deleteTask(id: number) {
  let tasks = load<any>(KEYS.tasks);
  const task = tasks.find((t: any) => t.id === id);
  if (!task) return;
  tasks = tasks.filter((t: any) => t.id !== id);
  // Renumber remaining tasks in same project
  const projectTasks = tasks.filter((t: any) => t.projectId === task.projectId)
    .sort((a: any, b: any) => a.sortOrder - b.sortOrder);
  projectTasks.forEach((t: any, i: number) => { t.taskNumber = i + 1; t.sortOrder = i + 1; });
  save(KEYS.tasks, tasks);
}
export function getProjectSummary(projectId: number) {
  const phases = listPhases(projectId);
  const tasks = listTasks(projectId);
  const now_ = new Date();
  const total = tasks.length;
  const completed = tasks.filter((t: any) => t.status === 'Completed').length;
  const inProgress = tasks.filter((t: any) => t.status === 'In Progress').length;
  const notStarted = tasks.filter((t: any) => t.status === 'Not Started').length;
  const blocked = tasks.filter((t: any) => t.status === 'Blocked').length;
  const overdue = tasks.filter((t: any) =>
    t.plannedEnd && new Date(t.plannedEnd) < now_ && t.status !== 'Completed'
  ).length;
  const pctComplete = total > 0 ? Math.round((completed / total) * 100) : 0;
  const phaseStats = phases.map((ph: any) => {
    const pt = tasks.filter((t: any) => t.phaseId === ph.id);
    const pc = pt.filter((t: any) => t.status === 'Completed').length;
    return { phaseId: ph.id, phaseName: ph.name, total: pt.length, completed: pc, pctComplete: pt.length > 0 ? Math.round((pc / pt.length) * 100) : 0 };
  });
  return { projectId, totalTasks: total, completed, inProgress, notStarted, overdue, blocked, pctComplete, phaseStats };
}

// ── Resources ─────────────────────────────────────────────────────────────────
export function listResources() { return load(KEYS.resources); }
export function createResource(input: { name: string; department?: string; role?: string; email?: string }) {
  const resources = load<any>(KEYS.resources);
  const r = { id: nextId('resources'), ...input, createdAt: now() };
  resources.push(r); save(KEYS.resources, resources); return r;
}
export function updateResource(id: number, patch: Record<string, unknown>) {
  const resources = load<any>(KEYS.resources);
  const idx = resources.findIndex((r: any) => r.id === id);
  if (idx === -1) return null;
  resources[idx] = { ...resources[idx], ...patch };
  save(KEYS.resources, resources); return resources[idx];
}
export function deleteResource(id: number) {
  const resources = load<any>(KEYS.resources).filter((r: any) => r.id !== id);
  save(KEYS.resources, resources);
}

// ── Holidays ──────────────────────────────────────────────────────────────────
export function listHolidays() {
  const holidays = load<any>(KEYS.holidays);
  const resources = load<any>(KEYS.resources);
  return holidays.map((h: any) => {
    const r = resources.find((r: any) => r.id === h.resourceId);
    return { ...h, resourceName: r ? r.name : null };
  });
}
export function createHoliday(input: { name: string; startDate: string; endDate: string; resourceId?: number; notes?: string }) {
  const holidays = load<any>(KEYS.holidays);
  const resources = load<any>(KEYS.resources);
  const h = { id: nextId('holidays'), ...input, resourceId: input.resourceId ?? null, notes: input.notes ?? null, createdAt: now() };
  holidays.push(h); save(KEYS.holidays, holidays);
  const r = resources.find((r: any) => r.id === h.resourceId);
  return { ...h, resourceName: r ? r.name : null };
}
export function deleteHoliday(id: number) {
  const holidays = load<any>(KEYS.holidays).filter((h: any) => h.id !== id);
  save(KEYS.holidays, holidays);
}
