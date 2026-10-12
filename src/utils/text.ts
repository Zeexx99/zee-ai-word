/** Text helpers shared by the Word service and the prompts. Pure and unit-tested. */

/**
 * Word returns paragraph breaks as "\r", manual line breaks as "\v" (vertical tab), table cell
 * ends as "\u0007" and fields/objects as control characters. Turn that into clean "\n" text.
 */
export function normalizeWordText(s: string): string {
  return s
    .replace(/\r\n?/g, "\n")
    .replace(/\u000b/g, "\n")
    .replace(/\u0007/g, "\t")
    .replace(/[\u0001-\u0006\u0008\u000e-\u001f]/g, "")
    .replace(/[ \t]+\n/g, "\n")
    .replace(/\n{3,}/g, "\n\n");
}

const PREAMBLE = /^(sure[,!.]?\s*)?(here(?:'s| is| are)\b[^\n]*?:|certainly[^\n]*?:|below is[^\n]*?:)\s*\n+/i;

/**
 * Gemini sometimes wraps a rewrite in a code fence, adds "Here is the revised text:" or puts the
 * whole answer in quotes. Strip that so only the text that belongs in the document remains.
 */
export function cleanModelOutput(text: string, original = ""): string {
  let s = text.replace(/\r\n?/g, "\n").trim();
  s = s.replace(PREAMBLE, "");
  const fence = s.match(/^(```|~~~)[\w-]*\n([\s\S]*?)\n\1\s*$/);
  if (fence) s = fence[2];
  const quoted = s.match(/^["“](.*)["”]$/s);
  const origQuoted = /^\s*["“]/.test(original);
  if (quoted && !origQuoted && !quoted[1].includes('"')) s = quoted[1];
  return s.trim();
}

/** Keep the original's leading/trailing whitespace and final paragraph break when replacing. */
export function preserveEdges(original: string, replacement: string): string {
  const lead = original.match(/^\s*/)?.[0] ?? "";
  const trail = original.match(/\s*$/)?.[0] ?? "";
  return lead + replacement.trim() + trail;
}

export function countWords(s: string): number {
  return (s.match(/[\p{L}\p{N}]+(?:['’-][\p{L}\p{N}]+)*/gu) ?? []).length;
}

/** Truncate long document text for prompts, keeping the start and the end. */
export function truncateMiddle(s: string, maxChars: number): { text: string; truncated: boolean } {
  if (s.length <= maxChars) return { text: s, truncated: false };
  const head = Math.floor(maxChars * 0.7);
  const tail = maxChars - head;
  return {
    text: `${s.slice(0, head)}\n\n[… ${(s.length - maxChars).toLocaleString("en-US")} characters omitted …]\n\n${s.slice(s.length - tail)}`,
    truncated: true,
  };
}

export function shorten(s: string, n: number): string {
  const one = s.replace(/\s+/g, " ").trim();
  return one.length > n ? `${one.slice(0, n - 1)}…` : one;
}
