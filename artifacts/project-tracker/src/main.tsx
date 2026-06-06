import { setFetchOverride } from '@workspace/api-client-react';
import { localFetch } from './lib/local-fetch';
import { createRoot } from "react-dom/client";
import App from "./App";
import "./index.css";

// Enable localStorage mode when built for GitHub Pages
if (import.meta.env.VITE_USE_LOCAL_DB === 'true') {
  setFetchOverride(localFetch as any);
}

createRoot(document.getElementById("root")!).render(<App />);
