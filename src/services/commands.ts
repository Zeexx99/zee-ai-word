/**
 * Ribbon menu, right-click menu and keyboard shortcuts (Office "ExecuteFunction" commands).
 *
 * The manifest points these at the task pane page (shared runtime), so the handlers below run in
 * the same JavaScript as the pane: they open the pane and ask it to start the action. The action
 * then shows its usual preview – nothing changes in the document without a click.
 */
export type CommandId = "rewrite" | "grammar" | "summarize" | "consistency" | "open";

/** Function names used in manifest.xml and public/shortcuts.json. */
export const COMMAND_FUNCTIONS: Record<string, CommandId> = {
  zeeRewrite: "rewrite",
  zeeGrammar: "grammar",
  zeeSummarize: "summarize",
  zeeConsistency: "consistency",
  zeeOpen: "open",
};

type Listener = (cmd: CommandId) => void;
let listener: Listener | null = null;
const queue: CommandId[] = [];

/** App subscribes once mounted; commands that arrived earlier are replayed. */
export function onCommand(l: Listener): () => void {
  listener = l;
  queue.splice(0).forEach((c) => l(c));
  return () => {
    if (listener === l) listener = null;
  };
}

function dispatch(cmd: CommandId) {
  if (listener) listener(cmd);
  else queue.push(cmd);
}

/** Registers the command handlers with Office (safe to call when unsupported). */
export function registerCommands(): void {
  try {
    if (typeof Office === "undefined" || !Office.actions?.associate) return;
    for (const [fn, cmd] of Object.entries(COMMAND_FUNCTIONS)) {
      Office.actions.associate(fn, async (event?: Office.AddinCommands.Event) => {
        try {
          await Office.addin?.showAsTaskpane?.();
        } catch {
          /* pane already open or shared runtime not supported */
        }
        dispatch(cmd);
        try {
          event?.completed();
        } catch {
          /* keyboard shortcuts pass no event */
        }
      });
    }
  } catch {
    /* older Office: ribbon menu and shortcuts are simply unavailable */
  }
}
