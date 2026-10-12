import { Component, type ReactNode } from "react";
import { createRoot, type Root } from "react-dom/client";
import App from "./App";
import "./styles.css";
import { registerCommands } from "./services/commands";

declare global {
  interface Window {
    __zeeLog?: (type: string, msg: string) => void;
    __zeeShowError?: (msg: string) => void;
    __zeeRoot?: Root;
  }
}
const log = (type: string, msg: string) => window.__zeeLog?.(type, msg);

/**
 * After a new deploy, Office may still run a cached copy of the old page whose code files no longer
 * exist on the server. Detect that and reload once with a cache-busting parameter.
 */
const STALE_RE = /Failed to fetch dynamically imported module|Importing a module script failed|error loading dynamically imported module|Unable to preload CSS/i;
function reloadFresh(reason: string): boolean {
  try {
    const key = "zeeai.word.staleReload";
    const last = Number(sessionStorage.getItem(key) ?? 0);
    if (Date.now() - last < 30_000) return false; // already tried – don't loop
    sessionStorage.setItem(key, String(Date.now()));
  } catch {
    /* storage blocked – still try once */
  }
  log("stale-build", reason);
  const url = new URL(window.location.href);
  url.searchParams.set("zeev", Date.now().toString(36));
  window.location.replace(url.toString());
  return true;
}
window.addEventListener("vite:preloadError", (e) => {
  if (reloadFresh("vite:preloadError")) e.preventDefault();
});

/** Shows render-time crashes instead of a blank pane. */
class ErrorBoundary extends Component<{ children: ReactNode }, { error: Error | null }> {
  state = { error: null as Error | null };
  static getDerivedStateFromError(error: Error) {
    return { error };
  }
  componentDidCatch(error: Error, info: { componentStack?: string | null }) {
    if (STALE_RE.test(error.message) && reloadFresh(error.message)) return;
    log("react-error", `${error.stack ?? error.message}\n${info.componentStack ?? ""}`);
  }
  render() {
    if (this.state.error) {
      return (
        <div style={{ padding: 12, font: "13px 'Segoe UI', sans-serif" }}>
          <p>
            <b>Zee AI hit an error.</b> Close the pane and open it again. If it keeps happening, run diagnose.bat and send diag-log.txt.
          </p>
          <pre style={{ whiteSpace: "pre-wrap", color: "#b42318", font: "12px Consolas, monospace" }}>{this.state.error.stack ?? this.state.error.message}</pre>
        </div>
      );
    }
    return this.props.children;
  }
}

let rendered = false;
function render(reason: string) {
  if (rendered) return;
  rendered = true;
  const el = document.getElementById("root");
  if (!el) return;
  log("render", `rendering (${reason}); host=${typeof Office !== "undefined" ? String(Office.context?.host) : "no-office"}`);
  // Re-use the root if this module runs again (dev hot reload): calling createRoot twice on the
  // same container logged a React warning in the Excel add-in.
  // No React.StrictMode: its double mount made Fluent UI log "Keyborg instance … disposed incorrectly"
  // and registered Office event handlers twice.
  const root = window.__zeeRoot ?? createRoot(el);
  window.__zeeRoot = root;
  root.render(
    <ErrorBoundary>
      <App />
    </ErrorBoundary>
  );
}

log("main", "main.tsx executing");
// Office.onReady resolves both inside Office and in a plain browser (host = null),
// so the UI can also be previewed at https://localhost:3200/taskpane.html.
if (typeof Office !== "undefined") {
  Office.onReady(() => {
    registerCommands(); // ribbon menu, right-click menu, keyboard shortcuts (shared runtime)
    render("Office.onReady");
  });
  // Safety net: never leave the pane blank if Office.onReady stalls.
  setTimeout(() => {
    if (!rendered) log("warn", "Office.onReady did not resolve within 8s – rendering anyway");
    render("timeout");
  }, 8000);
} else {
  log("warn", "office.js not available");
  render("no-office");
}
