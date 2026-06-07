import { useState, useRef } from "react";
import { Link } from "wouter";
import { Plus, Trash2, FolderOpen, Calendar, ChevronRight, Download, Upload } from "lucide-react";
import {
  useListProjects,
  useCreateProject,
  useDeleteProject,
  useGetProjectSummary,
  getListProjectsQueryKey,
  getGetProjectSummaryQueryKey,
} from "@workspace/api-client-react";
import { useQueryClient } from "@tanstack/react-query";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Textarea } from "@/components/ui/textarea";
import { Dialog, DialogContent, DialogHeader, DialogTitle, DialogTrigger } from "@/components/ui/dialog";
import { AlertDialog, AlertDialogAction, AlertDialogCancel, AlertDialogContent, AlertDialogDescription, AlertDialogFooter, AlertDialogHeader, AlertDialogTitle, AlertDialogTrigger } from "@/components/ui/alert-dialog";
import { useForm } from "react-hook-form";
import { Skeleton } from "@/components/ui/skeleton";
import { formatCompactDate } from "@/lib/constants";

function ProjectKPIBadges({ projectId }: { projectId: number }) {
  const { data: summary } = useGetProjectSummary(projectId, { query: { enabled: true, queryKey: getGetProjectSummaryQueryKey(projectId) } });
  if (!summary) return null;
  return (
    <div className="flex items-center gap-3 text-xs text-muted-foreground mt-2">
      <span className="flex items-center gap-1">
        <span className="w-2 h-2 rounded-full bg-emerald-500" />
        {summary.completed} done
      </span>
      <span className="flex items-center gap-1">
        <span className="w-2 h-2 rounded-full bg-blue-500" />
        {summary.inProgress} active
      </span>
      {summary.overdue > 0 && (
        <span className="flex items-center gap-1">
          <span className="w-2 h-2 rounded-full bg-rose-500" />
          {summary.overdue} overdue
        </span>
      )}
      {summary.blocked > 0 && (
        <span className="flex items-center gap-1">
          <span className="w-2 h-2 rounded-full bg-amber-500" />
          {summary.blocked} blocked
        </span>
      )}
      <span className="ml-auto font-medium text-foreground">{Math.round(summary.pctComplete)}% complete</span>
    </div>
  );
}

interface ProjectFormData {
  name: string;
  description: string;
  startDate: string;
  endDate: string;
}

const EXPORT_KEYS = ['eph_projects', 'eph_phases', 'eph_tasks', 'eph_resources', 'eph_holidays', 'eph_ids'];

function exportData() {
  const data: Record<string, unknown> = { exportedAt: new Date().toISOString() };
  EXPORT_KEYS.forEach(k => {
    try { data[k] = JSON.parse(localStorage.getItem(k) || 'null'); } catch { data[k] = null; }
  });
  const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `blueprint-${new Date().toISOString().slice(0, 10)}.json`;
  a.click();
  URL.revokeObjectURL(url);
}

function importData(file: File, onDone: () => void) {
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const data = JSON.parse(e.target?.result as string);
      EXPORT_KEYS.forEach(k => {
        if (data[k] !== undefined && data[k] !== null) {
          localStorage.setItem(k, JSON.stringify(data[k]));
        }
      });
      onDone();
    } catch {
      alert('Invalid backup file — could not import data.');
    }
  };
  reader.readAsText(file);
}

export default function ProjectsPage() {
  const queryClient = useQueryClient();
  const { data: projects, isLoading } = useListProjects();
  const createProject = useCreateProject();
  const deleteProject = useDeleteProject();
  const [open, setOpen] = useState(false);
  const importRef = useRef<HTMLInputElement>(null);

  const form = useForm<ProjectFormData>({
    defaultValues: { name: "", description: "", startDate: "", endDate: "" },
  });

  function onSubmit(data: ProjectFormData) {
    createProject.mutate(
      {
        data: {
          name: data.name,
          description: data.description || undefined,
          startDate: data.startDate || undefined,
          endDate: data.endDate || undefined,
        },
      },
      {
        onSuccess: () => {
          queryClient.invalidateQueries({ queryKey: getListProjectsQueryKey() });
          form.reset();
          setOpen(false);
        },
      }
    );
  }

  return (
    <div className="p-6 max-w-5xl mx-auto">
      <div className="flex items-center justify-between mb-8">
        <div>
          <h1 className="text-2xl font-bold tracking-tight">All Projects</h1>
          <p className="text-sm text-muted-foreground mt-1">
            {projects?.length ?? 0} project{projects?.length !== 1 ? "s" : ""}
          </p>
        </div>
        <div className="flex items-center gap-2">
          <input
            ref={importRef}
            type="file"
            accept=".json"
            className="hidden"
            onChange={(e) => {
              const file = e.target.files?.[0];
              if (!file) return;
              importData(file, () => {
                queryClient.invalidateQueries();
                e.target.value = '';
              });
            }}
          />
          <Button variant="outline" size="sm" onClick={() => importRef.current?.click()}>
            <Upload className="w-4 h-4 mr-2" />
            Import
          </Button>
          <Button variant="outline" size="sm" onClick={exportData}>
            <Download className="w-4 h-4 mr-2" />
            Export
          </Button>
          <Dialog open={open} onOpenChange={setOpen}>
          <DialogTrigger asChild>
            <Button data-testid="button-new-project">
              <Plus className="w-4 h-4 mr-2" />
              New Project
            </Button>
          </DialogTrigger>
          <DialogContent>
            <DialogHeader>
              <DialogTitle>Create Project</DialogTitle>
            </DialogHeader>
            <form onSubmit={form.handleSubmit(onSubmit)} className="space-y-4 mt-2">
              <div>
                <label className="text-sm font-medium">Name *</label>
                <Input
                  data-testid="input-project-name"
                  {...form.register("name", { required: true })}
                  placeholder="Project name"
                  className="mt-1"
                />
              </div>
              <div>
                <label className="text-sm font-medium">Description</label>
                <Textarea
                  data-testid="input-project-description"
                  {...form.register("description")}
                  placeholder="Brief description"
                  className="mt-1"
                  rows={3}
                />
              </div>
              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="text-sm font-medium">Start Date</label>
                  <Input
                    data-testid="input-project-start"
                    type="date"
                    {...form.register("startDate")}
                    className="mt-1"
                  />
                </div>
                <div>
                  <label className="text-sm font-medium">End Date</label>
                  <Input
                    data-testid="input-project-end"
                    type="date"
                    {...form.register("endDate")}
                    className="mt-1"
                  />
                </div>
              </div>
              <div className="flex justify-end gap-2 pt-2">
                <Button type="button" variant="outline" onClick={() => setOpen(false)}>
                  Cancel
                </Button>
                <Button type="submit" disabled={createProject.isPending} data-testid="button-submit-project">
                  {createProject.isPending ? "Creating..." : "Create"}
                </Button>
              </div>
            </form>
          </DialogContent>
        </Dialog>
        </div>
      </div>

      {isLoading ? (
        <div className="space-y-3">
          {[1, 2, 3].map((i) => (
            <Skeleton key={i} className="h-28 w-full rounded-lg" />
          ))}
        </div>
      ) : !projects?.length ? (
        <div className="text-center py-20 border border-dashed rounded-xl">
          <FolderOpen className="w-10 h-10 mx-auto text-muted-foreground mb-3" />
          <p className="text-sm font-medium text-muted-foreground">No projects yet</p>
          <p className="text-xs text-muted-foreground mt-1">Create your first project to get started</p>
        </div>
      ) : (
        <div className="space-y-3">
          {projects.map((project) => (
            <div
              key={project.id}
              data-testid={`card-project-${project.id}`}
              className="border rounded-lg p-4 bg-card hover:border-primary/40 transition-colors group"
            >
              <div className="flex items-start justify-between">
                <Link href={`/projects/${project.id}`} className="flex-1 min-w-0">
                  <div className="flex items-center gap-2">
                    <h2 className="font-semibold text-base group-hover:text-primary transition-colors">
                      {project.name}
                    </h2>
                    <ChevronRight className="w-4 h-4 text-muted-foreground group-hover:text-primary transition-colors" />
                  </div>
                  {project.description && (
                    <p className="text-sm text-muted-foreground mt-0.5 line-clamp-1">{project.description}</p>
                  )}
                  {(project.startDate || project.endDate) && (
                    <div className="flex items-center gap-1 text-xs text-muted-foreground mt-1">
                      <Calendar className="w-3 h-3" />
                      <span>
                        {formatCompactDate(project.startDate)} – {formatCompactDate(project.endDate)}
                      </span>
                    </div>
                  )}
                  <ProjectKPIBadges projectId={project.id} />
                </Link>
                <AlertDialog>
                  <AlertDialogTrigger asChild>
                    <Button
                      variant="ghost"
                      size="icon"
                      className="opacity-0 group-hover:opacity-100 transition-opacity ml-2 text-muted-foreground hover:text-destructive"
                      data-testid={`button-delete-project-${project.id}`}
                    >
                      <Trash2 className="w-4 h-4" />
                    </Button>
                  </AlertDialogTrigger>
                  <AlertDialogContent>
                    <AlertDialogHeader>
                      <AlertDialogTitle>Delete project?</AlertDialogTitle>
                      <AlertDialogDescription>
                        This will permanently delete <strong>{project.name}</strong> and all its phases and tasks.
                      </AlertDialogDescription>
                    </AlertDialogHeader>
                    <AlertDialogFooter>
                      <AlertDialogCancel>Cancel</AlertDialogCancel>
                      <AlertDialogAction
                        onClick={() =>
                          deleteProject.mutate(
                            { id: project.id },
                            { onSuccess: () => queryClient.invalidateQueries({ queryKey: getListProjectsQueryKey() }) }
                          )
                        }
                        className="bg-destructive text-destructive-foreground hover:bg-destructive/90"
                      >
                        Delete
                      </AlertDialogAction>
                    </AlertDialogFooter>
                  </AlertDialogContent>
                </AlertDialog>
              </div>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}
