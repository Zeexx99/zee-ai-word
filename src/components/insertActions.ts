import { createContext, useContext } from "react";
import type { ChatMessage } from "../hooks/useChat";
import type { InsertLocation } from "../services/wordService";

/** Actions that chat output can trigger (implemented in App, which owns confirmations + toasts). */
export interface InsertActions {
  /** Insert an assistant message into the document (Replace asks / previews first). */
  insert: (message: ChatMessage, location: InsertLocation) => Promise<void>;
  /** Add grammar findings as Word comments. */
  comment: (message: ChatMessage) => Promise<void>;
  /** Apply a table clean-up proposal cell by cell. */
  applyTable: (message: ChatMessage) => Promise<void>;
  /** Insert a caption below the selected picture, or set its alt text. */
  applyPicture: (message: ChatMessage) => Promise<void>;
  /** Select paragraph n (1-based) in the document – used by [P#] citations. */
  goToParagraph: (n: number) => void;
  /** True inside Word (insert buttons are disabled in the browser preview). */
  canInsert: boolean;
  /** Run an answer again (Regenerate / Retry). */
  retry: (assistantId: string) => void;
  /** Put a prompt back into the chat box to edit and resend it. */
  editPrompt: (text: string) => void;
  /** Copy a diagnostics report for a failed answer. */
  copyDiagnostics: (message: ChatMessage) => void;
  /** A request is running (Regenerate is disabled meanwhile). */
  busy: boolean;
}

export const InsertActionsContext = createContext<InsertActions | null>(null);

export function useInsertActions(): InsertActions {
  const ctx = useContext(InsertActionsContext);
  if (!ctx) throw new Error("InsertActionsContext missing");
  return ctx;
}

/** Open a URL in the user's browser (Office API inside Word, a new tab elsewhere). */
export function openExternal(url: string) {
  try {
    Office.context.ui.openBrowserWindow(url);
  } catch {
    window.open(url, "_blank", "noopener");
  }
}

export async function copyText(text: string): Promise<boolean> {
  try {
    await navigator.clipboard.writeText(text);
    return true;
  } catch {
    // The Clipboard API is often blocked in Office webviews – fall back to execCommand.
    const ta = document.createElement("textarea");
    ta.value = text;
    ta.style.position = "fixed";
    ta.style.opacity = "0";
    document.body.appendChild(ta);
    ta.select();
    let ok = false;
    try {
      ok = document.execCommand("copy");
    } catch {
      ok = false;
    }
    ta.remove();
    return ok;
  }
}
