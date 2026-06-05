import { useState } from "react";
import { Plus, Trash2, User, Mail, Building, Briefcase } from "lucide-react";
import {
  useListResources,
  useCreateResource,
  useDeleteResource,
  useUpdateResource,
  getListResourcesQueryKey,
} from "@workspace/api-client-react";
import { useQueryClient } from "@tanstack/react-query";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Dialog, DialogContent, DialogHeader, DialogTitle, DialogTrigger } from "@/components/ui/dialog";
import { AlertDialog, AlertDialogAction, AlertDialogCancel, AlertDialogContent, AlertDialogDescription, AlertDialogFooter, AlertDialogHeader, AlertDialogTitle, AlertDialogTrigger } from "@/components/ui/alert-dialog";
import { useForm } from "react-hook-form";
import { Skeleton } from "@/components/ui/skeleton";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "@/components/ui/table";

interface ResourceFormData {
  name: string;
  department: string;
  role: string;
  email: string;
}

export default function ResourcesPage() {
  const queryClient = useQueryClient();
  const { data: resources, isLoading } = useListResources();
  const createResource = useCreateResource();
  const deleteResource = useDeleteResource();
  const updateResource = useUpdateResource();
  const [open, setOpen] = useState(false);
  const [editingId, setEditingId] = useState<number | null>(null);

  const form = useForm<ResourceFormData>({
    defaultValues: { name: "", department: "", role: "", email: "" },
  });

  function onSubmit(data: ResourceFormData) {
    const payload = {
      name: data.name,
      department: data.department || undefined,
      role: data.role || undefined,
      email: data.email || undefined,
    };
    if (editingId) {
      updateResource.mutate(
        { id: editingId, data: payload },
        {
          onSuccess: () => {
            queryClient.invalidateQueries({ queryKey: getListResourcesQueryKey() });
            form.reset();
            setOpen(false);
            setEditingId(null);
          },
        }
      );
    } else {
      createResource.mutate(
        { data: payload },
        {
          onSuccess: () => {
            queryClient.invalidateQueries({ queryKey: getListResourcesQueryKey() });
            form.reset();
            setOpen(false);
          },
        }
      );
    }
  }

  function openEdit(r: { id: number; name: string; department?: string | null; role?: string | null; email?: string | null }) {
    setEditingId(r.id);
    form.reset({ name: r.name, department: r.department ?? "", role: r.role ?? "", email: r.email ?? "" });
    setOpen(true);
  }

  return (
    <div className="p-6 max-w-5xl mx-auto">
      <div className="flex items-center justify-between mb-8">
        <div>
          <h1 className="text-2xl font-bold tracking-tight">Resources</h1>
          <p className="text-sm text-muted-foreground mt-1">Team members available for task assignment</p>
        </div>
        <Dialog
          open={open}
          onOpenChange={(v) => {
            setOpen(v);
            if (!v) { setEditingId(null); form.reset(); }
          }}
        >
          <DialogTrigger asChild>
            <Button data-testid="button-new-resource">
              <Plus className="w-4 h-4 mr-2" />
              Add Resource
            </Button>
          </DialogTrigger>
          <DialogContent>
            <DialogHeader>
              <DialogTitle>{editingId ? "Edit Resource" : "Add Resource"}</DialogTitle>
            </DialogHeader>
            <form onSubmit={form.handleSubmit(onSubmit)} className="space-y-4 mt-2">
              <div>
                <label className="text-sm font-medium">Name *</label>
                <Input
                  data-testid="input-resource-name"
                  {...form.register("name", { required: true })}
                  placeholder="Full name"
                  className="mt-1"
                />
              </div>
              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="text-sm font-medium">Department</label>
                  <Input
                    data-testid="input-resource-dept"
                    {...form.register("department")}
                    placeholder="e.g. Engineering"
                    className="mt-1"
                  />
                </div>
                <div>
                  <label className="text-sm font-medium">Role</label>
                  <Input
                    data-testid="input-resource-role"
                    {...form.register("role")}
                    placeholder="e.g. Developer"
                    className="mt-1"
                  />
                </div>
              </div>
              <div>
                <label className="text-sm font-medium">Email</label>
                <Input
                  data-testid="input-resource-email"
                  {...form.register("email")}
                  type="email"
                  placeholder="email@example.com"
                  className="mt-1"
                />
              </div>
              <div className="flex justify-end gap-2 pt-2">
                <Button type="button" variant="outline" onClick={() => { setOpen(false); setEditingId(null); form.reset(); }}>
                  Cancel
                </Button>
                <Button type="submit" disabled={createResource.isPending || updateResource.isPending} data-testid="button-submit-resource">
                  {editingId ? "Save" : "Add"}
                </Button>
              </div>
            </form>
          </DialogContent>
        </Dialog>
      </div>

      {isLoading ? (
        <div className="space-y-2">
          {[1, 2, 3].map((i) => <Skeleton key={i} className="h-12 w-full" />)}
        </div>
      ) : !resources?.length ? (
        <div className="text-center py-20 border border-dashed rounded-xl">
          <User className="w-10 h-10 mx-auto text-muted-foreground mb-3" />
          <p className="text-sm font-medium text-muted-foreground">No resources yet</p>
          <p className="text-xs text-muted-foreground mt-1">Add team members to assign them to tasks</p>
        </div>
      ) : (
        <div className="border rounded-lg overflow-hidden">
          <Table>
            <TableHeader>
              <TableRow>
                <TableHead>Name</TableHead>
                <TableHead>Role</TableHead>
                <TableHead>Department</TableHead>
                <TableHead>Email</TableHead>
                <TableHead className="w-20" />
              </TableRow>
            </TableHeader>
            <TableBody>
              {resources.map((r) => (
                <TableRow
                  key={r.id}
                  data-testid={`row-resource-${r.id}`}
                  className="cursor-pointer hover:bg-muted/50"
                  onClick={() => openEdit(r)}
                >
                  <TableCell>
                    <div className="flex items-center gap-2">
                      <div className="w-7 h-7 rounded-full bg-primary/10 text-primary flex items-center justify-center text-xs font-semibold">
                        {r.name.charAt(0).toUpperCase()}
                      </div>
                      <span className="font-medium">{r.name}</span>
                    </div>
                  </TableCell>
                  <TableCell>
                    <div className="flex items-center gap-1.5 text-sm text-muted-foreground">
                      <Briefcase className="w-3.5 h-3.5" />
                      {r.role || "-"}
                    </div>
                  </TableCell>
                  <TableCell>
                    <div className="flex items-center gap-1.5 text-sm text-muted-foreground">
                      <Building className="w-3.5 h-3.5" />
                      {r.department || "-"}
                    </div>
                  </TableCell>
                  <TableCell>
                    <div className="flex items-center gap-1.5 text-sm text-muted-foreground">
                      <Mail className="w-3.5 h-3.5" />
                      {r.email || "-"}
                    </div>
                  </TableCell>
                  <TableCell onClick={(e) => e.stopPropagation()}>
                    <AlertDialog>
                      <AlertDialogTrigger asChild>
                        <Button variant="ghost" size="icon" className="text-muted-foreground hover:text-destructive" data-testid={`button-delete-resource-${r.id}`}>
                          <Trash2 className="w-4 h-4" />
                        </Button>
                      </AlertDialogTrigger>
                      <AlertDialogContent>
                        <AlertDialogHeader>
                          <AlertDialogTitle>Remove resource?</AlertDialogTitle>
                          <AlertDialogDescription>
                            Remove <strong>{r.name}</strong>? They will be unassigned from any tasks.
                          </AlertDialogDescription>
                        </AlertDialogHeader>
                        <AlertDialogFooter>
                          <AlertDialogCancel>Cancel</AlertDialogCancel>
                          <AlertDialogAction
                            onClick={() => deleteResource.mutate({ id: r.id }, { onSuccess: () => queryClient.invalidateQueries({ queryKey: getListResourcesQueryKey() }) })}
                            className="bg-destructive text-destructive-foreground"
                          >
                            Remove
                          </AlertDialogAction>
                        </AlertDialogFooter>
                      </AlertDialogContent>
                    </AlertDialog>
                  </TableCell>
                </TableRow>
              ))}
            </TableBody>
          </Table>
        </div>
      )}
    </div>
  );
}
