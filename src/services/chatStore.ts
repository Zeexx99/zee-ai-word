/**
 * Per-document chat history, kept in localStorage (never inside the document).
 *
 * Key: a hash of the document's full path/URL when it has been saved, so reopening the same file
 * brings its chat back without writing anything into the .docx. A new, never-saved document has
 * no path yet; it then gets a small random id stored in its document settings (a harmless value,
 * never a key or a setting).
 */
import { hashString } from "../utils/hash";

const DOC_ID_KEY = "zeeai.word.docId";
const MAX_MESSAGES = 60;
const MAX_PROMPT_CHARS = 12_000;
const MAX_ORIGINAL_CHARS = 20_000;

let cachedDocId: string | null = null;

/** Document title shown in the UI (file name without extension). */
export function getDocTitle(): string {
  try {
    const url = Office.context?.document?.url ?? "";
    if (!url) return "Untitled document";
    const name = decodeURIComponent(url.split(/[\\/]/).pop() ?? url);
    return name.replace(/\.(docx|docm|doc|dotx|rtf)$/i, "") || "Untitled document";
  } catch {
    return "Untitled document";
  }
}

export function getDocId(): string {
  if (cachedDocId) return cachedDocId;
  try {
    const doc = Office.context?.document;
    if (doc) {
      const url = (doc.url ?? "").trim();
      if (url) {
        cachedDocId = `doc-${hashString(url.toLowerCase())}`;
        return cachedDocId;
      }
      const settings = doc.settings;
      if (settings) {
        let id = settings.get(DOC_ID_KEY) as string | null;
        if (!id) {
          id = `new-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
          settings.set(DOC_ID_KEY, id);
          settings.saveAsync();
        }
        cachedDocId = id;
        return id;
      }
    }
  } catch {
    /* not in Office */
  }
  cachedDocId = "preview";
  return cachedDocId;
}

const keyFor = (docId: string) => `zeeai.word.chat.${docId}`;

export function loadChat<T>(docId: string): T[] {
  try {
    const raw = localStorage.getItem(keyFor(docId));
    const list = raw ? (JSON.parse(raw) as unknown) : [];
    return Array.isArray(list) ? (list as T[]) : [];
  } catch {
    return [];
  }
}

export function saveChat<T extends { promptText?: string; status?: string; original?: string }>(docId: string, messages: T[]): void {
  try {
    const trimmed = messages
      .filter((m) => m.status !== "streaming")
      .slice(-MAX_MESSAGES)
      .map((m) => {
        let out = m;
        if (out.promptText && out.promptText.length > MAX_PROMPT_CHARS) out = { ...out, promptText: out.promptText.slice(0, MAX_PROMPT_CHARS) };
        if (out.original && out.original.length > MAX_ORIGINAL_CHARS) out = { ...out, original: out.original.slice(0, MAX_ORIGINAL_CHARS) };
        return out;
      });
    localStorage.setItem(keyFor(docId), JSON.stringify(trimmed));
  } catch {
    // Quota exceeded: drop history for this document rather than failing.
    try {
      localStorage.removeItem(keyFor(docId));
    } catch {
      /* ignore */
    }
  }
}

export function clearChat(docId: string): void {
  try {
    localStorage.removeItem(keyFor(docId));
  } catch {
    /* ignore */
  }
}
