import { Switch, Route, Router as WouterRouter, Redirect } from "wouter";

import { QueryClient, QueryClientProvider } from "@tanstack/react-query";
import { Toaster } from "@/components/ui/toaster";
import { TooltipProvider } from "@/components/ui/tooltip";
import { AppLayout } from "@/components/layout";
import NotFound from "@/pages/not-found";
import ProjectsPage from "@/pages/projects";
import ProjectDetailPage from "@/pages/project-detail";
import ResourcesPage from "@/pages/resources";
import HolidaysPage from "@/pages/holidays";
import { ErrorBoundary } from "@/components/ErrorBoundary";

const queryClient = new QueryClient({
  defaultOptions: { queries: { staleTime: 30000, retry: 1 } },
});

function Router() {
  return (
    <AppLayout>
      <Switch>
        <Route path="/" component={ProjectsPage} />
        <Route path="/projects/:id/:section" component={ProjectDetailPage} />
        <Route path="/projects/:id">
          {(params) => <Redirect to={`/projects/${params.id}/plan`} />}
        </Route>
        <Route path="/resources" component={ResourcesPage} />
        <Route path="/holidays" component={HolidaysPage} />
        <Route component={NotFound} />
      </Switch>
    </AppLayout>
  );
}

function App() {
  return (
    <QueryClientProvider client={queryClient}>
      <TooltipProvider>
        <ErrorBoundary>
          <WouterRouter base={import.meta.env.BASE_URL.replace(/\/$/, "")}>
            <Router />
          </WouterRouter>
          <Toaster />
        </ErrorBoundary>
      </TooltipProvider>
    </QueryClientProvider>
  );
}

export default App;
