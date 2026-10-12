/** Saved prompts (per user, in this add-in's local storage). */
const KEY = "zeeai.word.prompts.v1";

export const DEFAULT_PROMPTS: string[] = [
  "Summarize this document in 5 bullet points and list any open questions.",
  "Check the document for inconsistent terms, dates and numbers, and list what to fix.",
  "Draft a professional introduction section for this document.",
  "Turn the selected notes into a clear, well-structured paragraph.",
  "Draft a polite follow-up email based on the selected text.",
];

export function loadPrompts(): string[] {
  try {
    const raw = localStorage.getItem(KEY);
    if (!raw) return [...DEFAULT_PROMPTS];
    const list = JSON.parse(raw) as unknown;
    return Array.isArray(list) ? list.filter((p): p is string => typeof p === "string" && p.trim() !== "") : [...DEFAULT_PROMPTS];
  } catch {
    return [...DEFAULT_PROMPTS];
  }
}

export function savePrompts(list: string[]): void {
  try {
    localStorage.setItem(KEY, JSON.stringify(list.slice(0, 50)));
  } catch {
    /* storage unavailable – keep for this session only */
  }
}
