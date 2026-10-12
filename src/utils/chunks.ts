/** Split long documents into prompt-sized chunks on paragraph boundaries. Pure and unit-tested. */

export interface Chunk {
  /** 1-based number of the first paragraph in the chunk (as numbered in the document). */
  firstParagraph: number;
  text: string;
}

/**
 * Group paragraphs into chunks of at most `maxChars` characters. A single paragraph longer than
 * `maxChars` is split at sentence boundaries so no chunk ever exceeds the limit by much.
 * Paragraph numbers count every paragraph (including empty ones) so they match Word's order.
 */
export function chunkParagraphs(paragraphs: string[], maxChars: number): Chunk[] {
  const chunks: Chunk[] = [];
  let cur: string[] = [];
  let curLen = 0;
  let first = 0;
  const push = () => {
    if (cur.length) chunks.push({ firstParagraph: first, text: cur.join("\n") });
    cur = [];
    curLen = 0;
    first = 0;
  };
  paragraphs.forEach((raw, i) => {
    const p = raw.trim();
    if (!p) return;
    const pieces = p.length > maxChars ? splitLong(p, maxChars) : [p];
    for (const piece of pieces) {
      if (curLen && curLen + piece.length + 1 > maxChars) push();
      if (!cur.length) first = i + 1;
      cur.push(piece);
      curLen += piece.length + 1;
    }
  });
  push();
  return chunks;
}

function splitLong(p: string, maxChars: number): string[] {
  const sentences = p.match(/[^.!?。؟]+[.!?。؟]*\s*/g) ?? [p];
  const out: string[] = [];
  let buf = "";
  for (const s of sentences) {
    if (buf && buf.length + s.length > maxChars) {
      out.push(buf.trim());
      buf = "";
    }
    if (s.length > maxChars) {
      for (let i = 0; i < s.length; i += maxChars) out.push(s.slice(i, i + maxChars).trim());
      continue;
    }
    buf += s;
  }
  if (buf.trim()) out.push(buf.trim());
  return out;
}

/** Number paragraphs as [P1], [P2]… (skipping empty ones, keeping Word's numbering) for citations. */
export function numberParagraphs(paragraphs: string[]): string {
  return paragraphs
    .map((p, i) => (p.trim() ? `[P${i + 1}] ${p.trim()}` : ""))
    .filter(Boolean)
    .join("\n");
}

/** Find [P12] / [P3, P4] style citations in an answer and turn them into markdown links. */
export function linkCitations(markdown: string): string {
  return markdown.replace(/\[(P\d+(?:\s*[,–-]\s*P?\d+)*)\](?!\()/g, (whole, inner: string) => {
    const nums = inner.match(/\d+/g) ?? [];
    if (!nums.length) return whole;
    return nums.map((n) => `[¶${n}](#zee-para-${n})`).join(" ");
  });
}
