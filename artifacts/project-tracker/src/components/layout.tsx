import React from "react";
import { LayoutDashboard, Users, Calendar, Table2, GitBranch, BarChart3, Flag } from "lucide-react";
import { Link, useLocation } from "wouter";
import { useListProjects } from "@workspace/api-client-react";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";

const REPORT_SECTIONS = [
  { section: "plan",      label: "Plan",       icon: Table2 },
  { section: "gantt",     label: "Gantt Chart", icon: GitBranch },
  { section: "dashboard", label: "Dashboard",  icon: BarChart3 },
  { section: "rag",       label: "RAG Status", icon: Flag },
];

const GLOBAL_NAV = [
  { href: "/",          label: "All Projects", icon: LayoutDashboard },
  { href: "/resources", label: "Resources",    icon: Users },
  { href: "/holidays",  label: "Holidays",     icon: Calendar },
];

export function AppLayout({ children }: { children: React.ReactNode }) {
  const [location, navigate] = useLocation();
  const { data: projects = [] } = useListProjects();

  const projectMatch = location.match(/^\/projects\/(\d+)/);
  const projectId = projectMatch ? parseInt(projectMatch[1], 10) : null;
  const sectionMatch = location.match(/^\/projects\/\d+\/(\w+)/);
  const currentSection = sectionMatch ? sectionMatch[1] : "plan";

  return (
    <div className="flex h-screen bg-background">
      <aside className="w-60 border-r bg-sidebar flex-shrink-0 flex flex-col">
        <div className="h-14 flex items-center px-4 border-b flex-shrink-0">
          <div className="flex items-center gap-2">
            <div className="w-8 h-8 bg-primary rounded flex items-center justify-center">
              <div className="w-3 h-3 bg-primary-foreground rotate-45" />
            </div>
            <span className="font-semibold tracking-tight text-sidebar-foreground">Blueprint</span>
          </div>
        </div>

        <nav className="flex-1 overflow-y-auto p-3 space-y-4">
          <div>
            <p className="px-2 mb-1.5 text-[11px] font-semibold uppercase tracking-wider text-muted-foreground/70">
              Project
            </p>
            <Select
              value={projectId ? String(projectId) : "__none__"}
              onValueChange={(val) => {
                if (val && val !== "__none__") {
                  navigate(`/projects/${val}/plan`);
                } else {
                  navigate("/");
                }
              }}
            >
              <SelectTrigger className="h-8 text-sm w-full bg-background">
                <SelectValue placeholder="Select project…" />
              </SelectTrigger>
              <SelectContent>
                <SelectItem value="__none__">
                  <span className="text-muted-foreground">— Select project —</span>
                </SelectItem>
                {projects.map((p) => (
                  <SelectItem key={p.id} value={String(p.id)}>
                    {p.name}
                  </SelectItem>
                ))}
              </SelectContent>
            </Select>
          </div>

          <div>
            <p className="px-2 mb-1.5 text-[11px] font-semibold uppercase tracking-wider text-muted-foreground/70">
              Reports
            </p>
            <div className="space-y-0.5">
              {REPORT_SECTIONS.map(({ section, label, icon: Icon }) => {
                const isActive = projectId !== null && currentSection === section;
                const isDisabled = projectId === null;
                return (
                  <button
                    key={section}
                    disabled={isDisabled}
                    onClick={() => projectId && navigate(`/projects/${projectId}/${section}`)}
                    className={`w-full flex items-center gap-2.5 px-3 py-2 rounded-md text-sm font-medium transition-colors text-left ${
                      isDisabled
                        ? "text-muted-foreground/40 cursor-not-allowed"
                        : isActive
                        ? "bg-sidebar-accent text-sidebar-accent-foreground"
                        : "text-muted-foreground hover:bg-sidebar-accent/50 hover:text-sidebar-foreground cursor-pointer"
                    }`}
                  >
                    <Icon className="w-4 h-4 flex-shrink-0" />
                    {label}
                  </button>
                );
              })}
            </div>
          </div>

          <div className="border-t pt-3">
            <p className="px-2 mb-1.5 text-[11px] font-semibold uppercase tracking-wider text-muted-foreground/70">
              Manage
            </p>
            <div className="space-y-0.5">
              {GLOBAL_NAV.map((item) => {
                const isActive =
                  location === item.href ||
                  (item.href !== "/" && location.startsWith(item.href));
                return (
                  <Link
                    key={item.href}
                    href={item.href}
                    className={`flex items-center gap-2.5 px-3 py-2 rounded-md text-sm font-medium transition-colors ${
                      isActive
                        ? "bg-sidebar-accent text-sidebar-accent-foreground"
                        : "text-muted-foreground hover:bg-sidebar-accent/50 hover:text-sidebar-foreground"
                    }`}
                  >
                    <item.icon className="w-4 h-4 flex-shrink-0" />
                    {item.label}
                  </Link>
                );
              })}
            </div>
          </div>
        </nav>
      </aside>

      <main className="flex-1 overflow-y-auto">
        {children}
      </main>
    </div>
  );
}
