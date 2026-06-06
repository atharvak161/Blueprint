import * as db from './local-db';

type Method = 'GET' | 'POST' | 'PATCH' | 'DELETE';

function jsonResp(data: unknown, status = 200): unknown {
  // Return a plain object — the customFetch override returns Promise<T> directly
  if (status === 404) throw Object.assign(new Error('Not found'), { status: 404 });
  if (status === 400) throw Object.assign(new Error((data as any)?.error || 'Bad request'), { status: 400 });
  if (status === 204) return null;
  return data;
}

function parseBody(options?: RequestInit): Record<string, unknown> {
  if (!options?.body) return {};
  try { return JSON.parse(options.body as string); } catch { return {}; }
}

export async function localFetch(input: RequestInfo | URL, options?: RequestInit): Promise<unknown> {
  const url = typeof input === 'string' ? input : input instanceof URL ? input.pathname : (input as Request).url;
  const method = ((options?.method || 'GET').toUpperCase()) as Method;
  const body = parseBody(options);

  // Strip leading /api
  const path = url.replace(/^.*\/api/, '');

  // ── Health ─────────────────────────────────────────────────────────────────
  if (path === '/healthz' && method === 'GET') return jsonResp({ status: 'ok' });

  // ── Projects ───────────────────────────────────────────────────────────────
  if (path === '/projects' && method === 'GET') return jsonResp(db.listProjects());
  if (path === '/projects' && method === 'POST') return jsonResp(db.createProject(body as any));

  const projMatch = path.match(/^\/projects\/(\d+)$/);
  if (projMatch) {
    const id = parseInt(projMatch[1]);
    if (method === 'GET') return jsonResp(db.getProject(id) || jsonResp(null, 404));
    if (method === 'PATCH') return jsonResp(db.updateProject(id, body));
    if (method === 'DELETE') { db.deleteProject(id); return jsonResp(null, 204); }
  }

  const projSummary = path.match(/^\/projects\/(\d+)\/summary$/);
  if (projSummary && method === 'GET') return jsonResp(db.getProjectSummary(parseInt(projSummary[1])));

  // ── Phases ─────────────────────────────────────────────────────────────────
  const phasesForProj = path.match(/^\/projects\/(\d+)\/phases$/);
  if (phasesForProj) {
    const projectId = parseInt(phasesForProj[1]);
    if (method === 'GET') return jsonResp(db.listPhases(projectId));
    if (method === 'POST') return jsonResp(db.createPhase(projectId, body as any));
  }
  const phaseMatch = path.match(/^\/phases\/(\d+)$/);
  if (phaseMatch) {
    const id = parseInt(phaseMatch[1]);
    if (method === 'PATCH') return jsonResp(db.updatePhase(id, body));
    if (method === 'DELETE') { db.deletePhase(id); return jsonResp(null, 204); }
  }

  // ── Tasks ──────────────────────────────────────────────────────────────────
  const tasksForProj = path.match(/^\/projects\/(\d+)\/tasks$/);
  if (tasksForProj) {
    const projectId = parseInt(tasksForProj[1]);
    if (method === 'GET') return jsonResp(db.listTasks(projectId));
    if (method === 'POST') return jsonResp(db.createTask(projectId, body as any));
  }
  const taskMatch = path.match(/^\/tasks\/(\d+)$/);
  if (taskMatch) {
    const id = parseInt(taskMatch[1]);
    if (method === 'PATCH') return jsonResp(db.updateTask(id, body));
    if (method === 'DELETE') { db.deleteTask(id); return jsonResp(null, 204); }
  }
  // Renumber endpoint (used by frontend after delete)
  if (path.match(/^\/projects\/\d+\/tasks\/renumber$/) && method === 'POST') return jsonResp(null, 204);

  // ── Resources ──────────────────────────────────────────────────────────────
  if (path === '/resources' && method === 'GET') return jsonResp(db.listResources());
  if (path === '/resources' && method === 'POST') return jsonResp(db.createResource(body as any));
  const resMatch = path.match(/^\/resources\/(\d+)$/);
  if (resMatch) {
    const id = parseInt(resMatch[1]);
    if (method === 'PATCH') return jsonResp(db.updateResource(id, body));
    if (method === 'DELETE') { db.deleteResource(id); return jsonResp(null, 204); }
  }

  // ── Holidays ───────────────────────────────────────────────────────────────
  if (path === '/holidays' && method === 'GET') return jsonResp(db.listHolidays());
  if (path === '/holidays' && method === 'POST') return jsonResp(db.createHoliday(body as any));
  const holMatch = path.match(/^\/holidays\/(\d+)$/);
  if (holMatch && method === 'DELETE') { db.deleteHoliday(parseInt(holMatch[1])); return jsonResp(null, 204); }

  throw Object.assign(new Error(`localFetch: unhandled ${method} ${path}`), { status: 404 });
}
