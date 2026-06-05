import { LayoutDashboard, Users, Calendar, Settings } from "lucide-react";
import { Link, useLocation } from "wouter";

export function AppLayout({ children }: { children: React.ReactNode }) {
  const [location] = useLocation();

  const navItems = [
    { href: "/", label: "Projects", icon: LayoutDashboard },
    { href: "/resources", label: "Resources", icon: Users },
    { href: "/holidays", label: "Holidays", icon: Calendar },
  ];

  return (
    <div className="flex h-screen bg-background">
      <aside className="w-64 border-r bg-sidebar flex-shrink-0 flex flex-col">
        <div className="h-14 flex items-center px-4 border-b">
          <div className="flex items-center gap-2">
            <div className="w-8 h-8 bg-primary rounded flex items-center justify-center">
              <div className="w-3 h-3 bg-primary-foreground rotate-45" />
            </div>
            <span className="font-semibold tracking-tight text-sidebar-foreground">ProgTrack</span>
          </div>
        </div>
        <nav className="flex-1 overflow-y-auto p-4 space-y-1">
          {navItems.map((item) => {
            const isActive = location === item.href || (item.href !== "/" && location.startsWith(item.href));
            return (
              <Link 
                key={item.href} 
                href={item.href}
                className={`flex items-center gap-3 px-3 py-2 rounded-md text-sm font-medium transition-colors ${
                  isActive 
                    ? "bg-sidebar-accent text-sidebar-accent-foreground" 
                    : "text-muted-foreground hover:bg-sidebar-accent/50 hover:text-sidebar-foreground"
                }`}
              >
                <item.icon className="w-4 h-4" />
                {item.label}
              </Link>
            );
          })}
        </nav>
      </aside>
      <main className="flex-1 overflow-y-auto">
        {children}
      </main>
    </div>
  );
}
