import { useState } from "react";
import { Plus, Trash2, CalendarDays } from "lucide-react";
import {
  useListHolidays,
  useCreateHoliday,
  useDeleteHoliday,
  getListHolidaysQueryKey,
} from "@workspace/api-client-react";
import { useQueryClient } from "@tanstack/react-query";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Dialog, DialogContent, DialogHeader, DialogTitle, DialogTrigger } from "@/components/ui/dialog";
import { AlertDialog, AlertDialogAction, AlertDialogCancel, AlertDialogContent, AlertDialogDescription, AlertDialogFooter, AlertDialogHeader, AlertDialogTitle, AlertDialogTrigger } from "@/components/ui/alert-dialog";
import { useForm } from "react-hook-form";
import { Skeleton } from "@/components/ui/skeleton";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "@/components/ui/table";
import { Badge } from "@/components/ui/badge";

interface HolidayFormData {
  name: string;
  startDate: string;
  endDate: string;
  notes: string;
}

function formatDate(d?: string | null) {
  if (!d) return "-";
  return new Date(d).toLocaleDateString("en-GB", { weekday: "short", day: "numeric", month: "short", year: "numeric" });
}

export default function HolidaysPage() {
  const queryClient = useQueryClient();
  const { data: holidays, isLoading } = useListHolidays();
  const createHoliday = useCreateHoliday();
  const deleteHoliday = useDeleteHoliday();
  const [open, setOpen] = useState(false);

  const form = useForm<HolidayFormData>({
    defaultValues: { name: "", startDate: "", endDate: "", notes: "" },
  });

  function onSubmit(data: HolidayFormData) {
    createHoliday.mutate(
      {
        data: {
          name: data.name,
          startDate: data.startDate,
          endDate: data.endDate,
          notes: data.notes || undefined,
        },
      },
      {
        onSuccess: () => {
          queryClient.invalidateQueries({ queryKey: getListHolidaysQueryKey() });
          form.reset();
          setOpen(false);
        },
      }
    );
  }

  const grouped = holidays?.reduce<Record<string, typeof holidays>>((acc, h) => {
    const year = h.startDate?.slice(0, 4) ?? "Unknown";
    if (!acc[year]) acc[year] = [];
    acc[year].push(h);
    return acc;
  }, {});

  return (
    <div className="p-6 max-w-5xl mx-auto">
      <div className="flex items-center justify-between mb-8">
        <div>
          <h1 className="text-2xl font-bold tracking-tight">UK Bank Holidays</h1>
          <p className="text-sm text-muted-foreground mt-1">
            Pre-seeded 2025–2027 England & Wales bank holidays, plus custom entries
          </p>
        </div>
        <Dialog open={open} onOpenChange={setOpen}>
          <DialogTrigger asChild>
            <Button data-testid="button-new-holiday">
              <Plus className="w-4 h-4 mr-2" />
              Add Holiday
            </Button>
          </DialogTrigger>
          <DialogContent>
            <DialogHeader>
              <DialogTitle>Add Holiday</DialogTitle>
            </DialogHeader>
            <form onSubmit={form.handleSubmit(onSubmit)} className="space-y-4 mt-2">
              <div>
                <label className="text-sm font-medium">Name *</label>
                <Input
                  data-testid="input-holiday-name"
                  {...form.register("name", { required: true })}
                  placeholder="e.g. Company Away Day"
                  className="mt-1"
                />
              </div>
              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="text-sm font-medium">Start Date *</label>
                  <Input
                    data-testid="input-holiday-start"
                    type="date"
                    {...form.register("startDate", { required: true })}
                    className="mt-1"
                  />
                </div>
                <div>
                  <label className="text-sm font-medium">End Date *</label>
                  <Input
                    data-testid="input-holiday-end"
                    type="date"
                    {...form.register("endDate", { required: true })}
                    className="mt-1"
                  />
                </div>
              </div>
              <div>
                <label className="text-sm font-medium">Notes</label>
                <Input
                  data-testid="input-holiday-notes"
                  {...form.register("notes")}
                  placeholder="Optional notes"
                  className="mt-1"
                />
              </div>
              <div className="flex justify-end gap-2 pt-2">
                <Button type="button" variant="outline" onClick={() => setOpen(false)}>
                  Cancel
                </Button>
                <Button type="submit" disabled={createHoliday.isPending} data-testid="button-submit-holiday">
                  Add
                </Button>
              </div>
            </form>
          </DialogContent>
        </Dialog>
      </div>

      {isLoading ? (
        <div className="space-y-2">
          {[1, 2, 3, 4, 5].map((i) => <Skeleton key={i} className="h-12 w-full" />)}
        </div>
      ) : !holidays?.length ? (
        <div className="text-center py-20 border border-dashed rounded-xl">
          <CalendarDays className="w-10 h-10 mx-auto text-muted-foreground mb-3" />
          <p className="text-sm font-medium text-muted-foreground">No holidays</p>
        </div>
      ) : (
        <div className="space-y-6">
          {Object.entries(grouped ?? {}).sort(([a], [b]) => a.localeCompare(b)).map(([year, items]) => (
            <div key={year}>
              <h2 className="text-sm font-semibold text-muted-foreground mb-2 uppercase tracking-wider">{year}</h2>
              <div className="border rounded-lg overflow-hidden">
                <Table>
                  <TableHeader>
                    <TableRow>
                      <TableHead>Holiday</TableHead>
                      <TableHead>Start</TableHead>
                      <TableHead>End</TableHead>
                      <TableHead>Notes</TableHead>
                      <TableHead>Type</TableHead>
                      <TableHead className="w-16" />
                    </TableRow>
                  </TableHeader>
                  <TableBody>
                    {items.sort((a, b) => (a.startDate ?? "").localeCompare(b.startDate ?? "")).map((h) => (
                      <TableRow key={h.id} data-testid={`row-holiday-${h.id}`}>
                        <TableCell className="font-medium">{h.name}</TableCell>
                        <TableCell className="text-sm text-muted-foreground">{formatDate(h.startDate)}</TableCell>
                        <TableCell className="text-sm text-muted-foreground">{formatDate(h.endDate)}</TableCell>
                        <TableCell className="text-sm text-muted-foreground">{h.notes || "-"}</TableCell>
                        <TableCell>
                          {h.resourceId ? (
                            <Badge variant="outline" className="text-xs">{h.resourceName ?? "Personal"}</Badge>
                          ) : (
                            <Badge variant="secondary" className="text-xs">Public</Badge>
                          )}
                        </TableCell>
                        <TableCell>
                          <AlertDialog>
                            <AlertDialogTrigger asChild>
                              <Button variant="ghost" size="icon" className="text-muted-foreground hover:text-destructive" data-testid={`button-delete-holiday-${h.id}`}>
                                <Trash2 className="w-4 h-4" />
                              </Button>
                            </AlertDialogTrigger>
                            <AlertDialogContent>
                              <AlertDialogHeader>
                                <AlertDialogTitle>Remove holiday?</AlertDialogTitle>
                                <AlertDialogDescription>
                                  Remove <strong>{h.name}</strong>?
                                </AlertDialogDescription>
                              </AlertDialogHeader>
                              <AlertDialogFooter>
                                <AlertDialogCancel>Cancel</AlertDialogCancel>
                                <AlertDialogAction
                                  onClick={() => deleteHoliday.mutate({ id: h.id }, { onSuccess: () => queryClient.invalidateQueries({ queryKey: getListHolidaysQueryKey() }) })}
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
            </div>
          ))}
        </div>
      )}
    </div>
  );
}
