/** Chat history → Gemini turns. Pure and unit-tested. */

export interface HistoryTurn {
  role: "user" | "model";
  text: string;
}

export interface HistoryMessage {
  role: "user" | "assistant";
  text: string;
  promptText?: string;
  status: "streaming" | "done" | "error";
}

const CONTEXT_MARKER = /\n\n(--- (SELECTED TEXT|DOCUMENT|CONTEXT) START ---|\[(Selected text|Document))/;

/** Build alternating user/model history; older turns drop their bulky document context. */
export function buildHistory(messages: HistoryMessage[], next: string, maxTurns = 16): HistoryTurn[] {
  const turns: HistoryTurn[] = [];
  for (let i = 0; i < messages.length; i++) {
    const m = messages[i];
    const prev = messages[i - 1];
    if (m.role === "assistant" && m.status === "done" && prev?.role === "user") {
      turns.push({ role: "user", text: prev.promptText ?? prev.text }, { role: "model", text: m.text || "(no text)" });
    }
  }
  const recent = turns.slice(-maxTurns);
  const compacted = recent.map((t, i) => {
    if (t.role !== "user" || i >= recent.length - 4) return t;
    const idx = t.text.search(CONTEXT_MARKER);
    return idx > 0 ? { ...t, text: `${t.text.slice(0, idx)}\n\n(document context omitted)` } : t;
  });
  return [...compacted, { role: "user", text: next }];
}

