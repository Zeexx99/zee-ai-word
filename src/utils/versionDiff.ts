/** Paragraph-level comparison of an older and the current version of a document. Pure – unit tested. */

export interface DiffBlock {
  kind: "changed" | "added" | "removed";
  oldText: string[];
  newText: string[];
  /** 1-based paragraph numbers in the CURRENT document (for removed text: the paragraph after the gap). */
  newParagraphs: number[];
}

const norm = (s: string) => s.replace(/\s+/g, " ").trim().toLowerCase();

/**
 * LCS over normalized paragraphs (empty paragraphs ignored). Adjacent removed + added paragraphs
 * become one "changed" block. Large documents fall back to a cheaper prefix/suffix comparison.
 */
export function diffParagraphs(oldParas: string[], newParas: string[], maxCells = 6_000_000): DiffBlock[] {
  const A = oldParas.map((t, i) => ({ t, i })).filter((x) => x.t.trim());
  const B = newParas.map((t, i) => ({ t, i })).filter((x) => x.t.trim());
  const a = A.map((x) => norm(x.t));
  const b = B.map((x) => norm(x.t));
  // Common prefix / suffix first (most edits are local).
  let pre = 0;
  while (pre < a.length && pre < b.length && a[pre] === b[pre]) pre++;
  let suf = 0;
  while (suf < a.length - pre && suf < b.length - pre && a[a.length - 1 - suf] === b[b.length - 1 - suf]) suf++;
  const a2 = a.slice(pre, a.length - suf);
  const b2 = b.slice(pre, b.length - suf);
  type Op = { op: "=" | "-" | "+"; ai?: number; bi?: number };
  const ops: Op[] = [];
  if ((a2.length + 1) * (b2.length + 1) <= maxCells) {
    const n = a2.length;
    const m = b2.length;
    const L: Uint32Array[] = Array.from({ length: n + 1 }, () => new Uint32Array(m + 1));
    for (let i = n - 1; i >= 0; i--) for (let j = m - 1; j >= 0; j--) L[i][j] = a2[i] === b2[j] ? L[i + 1][j + 1] + 1 : Math.max(L[i + 1][j], L[i][j + 1]);
    let i = 0;
    let j = 0;
    while (i < n && j < m) {
      if (a2[i] === b2[j]) {
        ops.push({ op: "=", ai: i++, bi: j++ });
      } else if (L[i + 1][j] >= L[i][j + 1]) ops.push({ op: "-", ai: i++ });
      else ops.push({ op: "+", bi: j++ });
    }
    while (i < n) ops.push({ op: "-", ai: i++ });
    while (j < m) ops.push({ op: "+", bi: j++ });
  } else {
    a2.forEach((_, i) => ops.push({ op: "-", ai: i }));
    b2.forEach((_, j) => ops.push({ op: "+", bi: j }));
  }

  const blocks: DiffBlock[] = [];
  let cur: DiffBlock | null = null;
  let lastNew = pre; // index into B of the next current paragraph
  const flush = () => {
    if (!cur) return;
    cur.kind = cur.oldText.length && cur.newText.length ? "changed" : cur.newText.length ? "added" : "removed";
    if (!cur.newParagraphs.length) {
      const anchor = B[Math.min(lastNew, B.length - 1)];
      if (anchor) cur.newParagraphs.push(anchor.i + 1);
    }
    blocks.push(cur);
    cur = null;
  };
  for (const o of ops) {
    if (o.op === "=") {
      flush();
      lastNew = pre + (o.bi as number) + 1;
      continue;
    }
    cur = cur ?? { kind: "changed", oldText: [], newText: [], newParagraphs: [] };
    if (o.op === "-") cur.oldText.push(A[pre + (o.ai as number)].t);
    else {
      const x = B[pre + (o.bi as number)];
      cur.newText.push(x.t);
      cur.newParagraphs.push(x.i + 1);
      lastNew = pre + (o.bi as number) + 1;
    }
  }
  flush();
  return blocks;
}

/** Compact text of the differences for the prompt (capped). */
export function blocksForPrompt(blocks: DiffBlock[], maxChars = 60_000): { text: string; truncated: boolean } {
  const parts: string[] = [];
  let size = 0;
  let truncated = false;
  for (const [k, b] of blocks.entries()) {
    const where = b.newParagraphs.length ? `[P${b.newParagraphs[0]}${b.newParagraphs.length > 1 ? `–P${b.newParagraphs[b.newParagraphs.length - 1]}` : ""}]` : "";
    const piece = `### Change ${k + 1} (${b.kind}) ${where}\n${b.oldText.length ? `OLD:\n${b.oldText.join("\n")}\n` : ""}${b.newText.length ? `NEW:\n${b.newText.join("\n")}\n` : ""}`;
    if (size + piece.length > maxChars && parts.length) {
      truncated = true;
      break;
    }
    parts.push(piece);
    size += piece.length;
  }
  return { text: parts.join("\n"), truncated };
}
