/**
 * Formatting-preserving edits.
 *
 * Instead of replacing a whole selection (which loses bold words, links, styles and comments
 * inside it), Zee AI turns a rewrite into small "hunks" per paragraph:
 *   left context + deleted words + right context  →  inserted words
 * Each hunk is located in Word with a short search and only the changed words are replaced,
 * so everything else keeps its formatting – and Track Changes shows clean word-level revisions.
 *
 * This module is pure (no Office.js) and unit-tested; wordService.ts applies the plan.
 */
import { diffWords } from "./diff";

export interface Hunk {
  /** Unchanged text right before the change (search anchor, may be ""). */
  left: string;
  /** Original words to remove (may be "" for a pure insertion). */
  del: string;
  /** New words (may be "" for a pure deletion). */
  ins: string;
  /** Unchanged text right after the change (search anchor, may be ""). */
  right: string;
}

export interface ParagraphEdit {
  /** Position among the non-empty paragraphs of the selection. */
  index: number;
  before: string;
  after: string;
  /** null = this paragraph changed too much; replace its text as a whole (paragraph style is kept). */
  hunks: Hunk[] | null;
}

/** Word's search() accepts at most 255 characters. */
export const MAX_SEARCH = 255;
const CONTEXT_CHARS = 48;
const MAX_HUNKS_PER_PARAGRAPH = 40;

/** Last few words of `s` (at most CONTEXT_CHARS, cut at a word boundary). */
export function tailContext(s: string, max = CONTEXT_CHARS): string {
  if (s.length <= max) return s;
  const cut = s.slice(s.length - max);
  const sp = cut.search(/\s/);
  return sp >= 0 && sp < cut.length - 1 ? cut.slice(sp) : cut;
}

/** First few words of `s` (at most CONTEXT_CHARS, cut at a word boundary). */
export function headContext(s: string, max = CONTEXT_CHARS): string {
  if (s.length <= max) return s;
  const cut = s.slice(0, max);
  const sp = Math.max(cut.lastIndexOf(" "), cut.lastIndexOf("\t"));
  return sp > 0 ? cut.slice(0, sp) : cut;
}

/** Split one paragraph's change into searchable hunks, or null when a whole-paragraph replace is safer. */
export function hunksFor(before: string, after: string): Hunk[] | null {
  if (before === after) return [];
  if (/[\t\u0007]/.test(before)) return null; // tables / tabs: Word's text and search don't line up reliably
  const ops = diffWords(before, after);
  const hunks: Hunk[] = [];
  for (let i = 0; i < ops.length; i++) {
    if (ops[i].type === "eq") continue;
    let del = "";
    let ins = "";
    const start = i;
    while (i < ops.length && ops[i].type !== "eq") {
      if (ops[i].type === "del") del += ops[i].text;
      else ins += ops[i].text;
      i++;
    }
    const prevEq = start > 0 && ops[start - 1].type === "eq" ? ops[start - 1].text : "";
    const nextEq = i < ops.length && ops[i].type === "eq" ? ops[i].text : "";
    i--; // the for-loop increments
    const left = tailContext(prevEq);
    const right = headContext(nextEq);
    if (!left.trim() && !right.trim()) return null; // nothing stable to anchor on
    if (left.length + del.length + right.length > MAX_SEARCH) return null;
    hunks.push({ left, del, ins, right });
  }
  // Mostly rewritten paragraphs read better (and track more cleanly) as one replacement.
  if (hunks.length > MAX_HUNKS_PER_PARAGRAPH) return null;
  const changed = hunks.reduce((n, h) => n + h.del.length, 0);
  if (before.length > 0 && changed / before.length > 0.6) return null;
  return hunks;
}

/** Non-empty paragraphs of a (normalized) text, in order. */
export function splitParagraphs(text: string): string[] {
  return text
    .replace(/\r\n?/g, "\n")
    .split("\n")
    .map((p) => p.replace(/\s+$/, ""))
    .filter((p) => p.trim() !== "");
}

/**
 * Pair the selection's paragraphs with the model's paragraphs and plan the edits.
 * Returns null when the paragraph count changed (e.g. Shorten merged paragraphs) – the caller then
 * falls back to replacing the whole selection.
 */
export function planEdits(originalParagraphs: string[], newText: string): ParagraphEdit[] | null {
  const orig = originalParagraphs.filter((p) => p.trim() !== "");
  const next = splitParagraphs(newText).map((p) => p.trim());
  if (!orig.length || orig.length !== next.length) return null;
  const edits: ParagraphEdit[] = [];
  orig.forEach((before, index) => {
    const lead = before.match(/^\s*/)?.[0] ?? "";
    const core = before.trim();
    const after = next[index];
    if (core === after) return;
    const hunks = hunksFor(core, after);
    edits.push({ index, before, after: lead + after, hunks });
  });
  return edits;
}

/** Count non-overlapping occurrences of `needle` in `hay`. */
export function countOccurrences(hay: string, needle: string): number {
  if (!needle) return 0;
  let n = 0;
  let i = hay.indexOf(needle);
  while (i >= 0) {
    n++;
    i = hay.indexOf(needle, i + needle.length);
  }
  return n;
}

/** Word search treats "^" as a special-character prefix; escape it. */
export function escapeWordSearch(s: string): string {
  return s.replace(/\^/g, "^^");
}

/**
 * String version of what wordService does in Word: locate each hunk by its unique
 * left+del+right anchor and replace only `del`. Used by the tests to prove a plan is sound;
 * returns null where Word would fall back to a whole-paragraph replace.
 */
export function applyHunksToText(paragraph: string, hunks: Hunk[]): string | null {
  let text = paragraph;
  for (const h of [...hunks].reverse()) {
    const anchor = h.left + h.del + h.right;
    if (countOccurrences(text, anchor) !== 1) return null;
    const at = text.indexOf(anchor) + h.left.length;
    text = text.slice(0, at) + h.ins + text.slice(at + h.del.length);
  }
  return text;
}
