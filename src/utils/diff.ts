/**
 * Word-level diff (Myers O(ND) algorithm) used for the "review before applying" preview.
 * Pure and unit-tested; no DOM or Office dependencies.
 */

export type DiffOp = { type: "eq" | "ins" | "del"; text: string };

/** Split text into words, whitespace runs and single punctuation marks (keeps every character). */
export function tokenize(text: string): string[] {
  return text.match(/\s+|[\p{L}\p{N}\p{M}_'’-]+|[^\s\p{L}\p{N}\p{M}]/gu) ?? [];
}

/** Myers diff on token arrays. Returns null when the texts differ by more than `maxD` edits. */
function myers(a: string[], b: string[], maxD: number): DiffOp[] | null {
  const n = a.length;
  const m = b.length;
  const max = n + m;
  const offset = max + 1;
  const v = new Int32Array(2 * max + 3);
  const trace: Int32Array[] = [];
  const limit = Math.min(max, maxD);

  for (let d = 0; d <= limit; d++) {
    // Snapshot of the state after d-1 edits, for diagonals -d..d (index k + d).
    trace.push(v.slice(offset - d, offset + d + 1));
    for (let k = -d; k <= d; k += 2) {
      let x = k === -d || (k !== d && v[offset + k - 1] < v[offset + k + 1]) ? v[offset + k + 1] : v[offset + k - 1] + 1;
      let y = x - k;
      while (x < n && y < m && a[x] === b[y]) {
        x++;
        y++;
      }
      v[offset + k] = x;
      if (x >= n && y >= m) return backtrack(a, b, trace, d);
    }
  }
  return null;
}

function backtrack(a: string[], b: string[], trace: Int32Array[], D: number): DiffOp[] {
  const rev: DiffOp[] = [];
  let x = a.length;
  let y = b.length;
  for (let d = D; d > 0; d--) {
    const vp = trace[d];
    const get = (k: number) => vp[k + d];
    const k = x - y;
    const prevK = k === -d || (k !== d && get(k - 1) < get(k + 1)) ? k + 1 : k - 1;
    const prevX = get(prevK);
    const prevY = prevX - prevK;
    while (x > prevX && y > prevY) {
      rev.push({ type: "eq", text: a[x - 1] });
      x--;
      y--;
    }
    if (x === prevX) {
      rev.push({ type: "ins", text: b[y - 1] });
      y--;
    } else {
      rev.push({ type: "del", text: a[x - 1] });
      x--;
    }
  }
  while (x > 0 && y > 0) {
    rev.push({ type: "eq", text: a[x - 1] });
    x--;
    y--;
  }
  return rev.reverse();
}

/** Merge neighbours of the same type; put deletions before insertions inside a change block. */
export function mergeOps(ops: DiffOp[]): DiffOp[] {
  const out: DiffOp[] = [];
  let i = 0;
  while (i < ops.length) {
    if (ops[i].type === "eq") {
      let text = "";
      while (i < ops.length && ops[i].type === "eq") text += ops[i++].text;
      out.push({ type: "eq", text });
      continue;
    }
    let del = "";
    let ins = "";
    while (i < ops.length && ops[i].type !== "eq") {
      if (ops[i].type === "del") del += ops[i].text;
      else ins += ops[i].text;
      i++;
    }
    if (del) out.push({ type: "del", text: del });
    if (ins) out.push({ type: "ins", text: ins });
  }
  return out;
}

/**
 * Small whitespace-only "equal" runs between two changes are folded into the change, so the
 * preview reads "~~old phrase~~ new phrase" instead of a word-by-word zebra.
 */
function absorbWhitespace(ops: DiffOp[]): DiffOp[] {
  const out: DiffOp[] = [];
  for (let i = 0; i < ops.length; i++) {
    const op = ops[i];
    const prev = ops[i - 1];
    const next = ops[i + 1];
    if (op.type === "eq" && /^[ \t]+$/.test(op.text) && prev && next && prev.type !== "eq" && next.type !== "eq") {
      out.push({ type: "del", text: op.text }, { type: "ins", text: op.text });
    } else out.push(op);
  }
  return mergeOps(out);
}

export const MAX_DIFF_TOKENS = 12000;

/** Diff two texts at word level. Very large or completely different texts become one replace block. */
export function diffWords(before: string, after: string, maxD = 2500): DiffOp[] {
  if (before === after) return before ? [{ type: "eq", text: before }] : [];
  const a = tokenize(before);
  const b = tokenize(after);
  let ops: DiffOp[] | null = null;
  if (a.length + b.length <= MAX_DIFF_TOKENS) ops = myers(a, b, maxD);
  if (!ops) {
    ops = [];
    if (before) ops.push({ type: "del", text: before });
    if (after) ops.push({ type: "ins", text: after });
    return ops;
  }
  return absorbWhitespace(mergeOps(ops));
}

export interface DiffStats {
  added: number;
  removed: number;
  /** Share of the original text that is unchanged (0–1). */
  kept: number;
}

const countWords = (s: string) => (s.match(/[\p{L}\p{N}]+/gu) ?? []).length;

export function diffStats(ops: DiffOp[]): DiffStats {
  let added = 0;
  let removed = 0;
  let same = 0;
  for (const op of ops) {
    const w = countWords(op.text);
    if (op.type === "ins") added += w;
    else if (op.type === "del") removed += w;
    else same += w;
  }
  return { added, removed, kept: same + removed ? same / (same + removed) : 1 };
}

/** Apply the ops to rebuild either side (used by tests and as a sanity check). */
export function rebuild(ops: DiffOp[], side: "before" | "after"): string {
  return ops
    .filter((o) => o.type === "eq" || (side === "before" ? o.type === "del" : o.type === "ins"))
    .map((o) => o.text)
    .join("");
}
