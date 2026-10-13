/**
 * "General" scope: turns the navigator's answer (which paragraphs a request is about) into a safe plan,
 * and builds the focused context sent with "answer" requests. Pure functions – unit tested.
 */

export interface LocateResult {
  action?: string;
  start?: number;
  end?: number;
  insertAfter?: number;
  paragraphs?: number[];
  reason?: string;
}

export type LocatePlan =
  | { action: "edit"; start: number; end: number; reason: string }
  | { action: "insert"; after: number; reason: string } // after = 0 → end of document
  | { action: "answer"; paragraphs: number[]; reason: string }
  | { action: "none"; reason: string };

const int = (v: unknown): number | null => (typeof v === "number" && Number.isFinite(v) ? Math.round(v) : typeof v === "string" && /^\s*\d+\s*$/.test(v) ? Number(v) : null);

/**
 * Validates the model's answer against the real document: paragraph numbers must exist, ranges are
 * ordered, empty paragraphs at the edges of an edit range are trimmed. `paras` = the document's
 * paragraphs (1-based numbering, as sent to the model).
 */
export function normalizeLocate(raw: LocateResult | null | undefined, paras: string[]): LocatePlan {
  const n = paras.length;
  const reason = typeof raw?.reason === "string" ? raw.reason.trim().slice(0, 300) : "";
  const action = typeof raw?.action === "string" ? raw.action.trim().toLowerCase() : "";
  const valid = (p: number | null): p is number => p !== null && p >= 1 && p <= n;
  const hasText = (p: number) => !!paras[p - 1]?.trim();

  if (action === "edit") {
    let s = int(raw?.start);
    let e = int(raw?.end);
    if (!valid(s) && valid(e)) s = e;
    if (valid(s) && !valid(e)) e = s;
    if (valid(s) && valid(e)) {
      if (s > e) [s, e] = [e, s];
      while (s < e && !hasText(s)) s++;
      while (e > s && !hasText(e)) e--;
      if (hasText(s)) return { action: "edit", start: s, end: e, reason };
    }
    // An "edit" without a usable range: fall back to answering from the listed paragraphs.
  }

  if (action === "insert") {
    const a = int(raw?.insertAfter);
    if (a === 0) return { action: "insert", after: 0, reason };
    if (valid(a)) return { action: "insert", after: a, reason };
    const e = int(raw?.end);
    if (valid(e)) return { action: "insert", after: e, reason };
    return { action: "insert", after: 0, reason };
  }

  const listed = Array.isArray(raw?.paragraphs) ? raw!.paragraphs.map(int).filter(valid) : [];
  const fromRange = [int(raw?.start), int(raw?.end)].filter(valid);
  const ps = Array.from(new Set([...listed, ...fromRange])).filter(hasText).slice(0, 40);
  if (action === "none" && !ps.length) return { action: "none", reason };
  if (ps.length) return { action: "answer", paragraphs: ps, reason };
  return { action: "none", reason };
}

/**
 * Focused context for "answer": the chosen paragraphs plus `around` neighbours each side, in document
 * order, keeping their real [P#] numbers; gaps are marked with "…". Capped at `maxChars`.
 */
export function relevantExcerpt(paras: string[], picks: number[], around = 1, maxChars = 30_000): { text: string; used: number[] } {
  const want = new Set<number>();
  for (const p of picks) for (let k = p - around; k <= p + around; k++) if (k >= 1 && k <= paras.length && paras[k - 1]?.trim()) want.add(k);
  const order = Array.from(want).sort((a, b) => a - b);
  // Keep the most relevant picks first when trimming to the budget.
  const priority = [...picks.filter((p) => want.has(p)), ...order.filter((p) => !picks.includes(p))];
  const keep = new Set<number>();
  let size = 0;
  for (const p of priority) {
    const len = paras[p - 1].length + 8;
    if (size + len > maxChars && keep.size) continue;
    keep.add(p);
    size += len;
  }
  const used = order.filter((p) => keep.has(p));
  const lines: string[] = [];
  used.forEach((p, i) => {
    if (i > 0 && p !== used[i - 1] + 1) lines.push("…");
    lines.push(`[P${p}] ${paras[p - 1]}`);
  });
  return { text: lines.join("\n"), used };
}

/** "¶3", "¶3–7" */
export function paraLabel(start: number, end: number = start): string {
  return start === end ? `¶${start}` : `¶${start}–${end}`;
}

/** "¶2, ¶5–7, ¶12" for a list of paragraph numbers. */
export function paraListLabel(ps: number[], max = 6): string {
  const sorted = Array.from(new Set(ps)).sort((a, b) => a - b);
  const runs: Array<[number, number]> = [];
  for (const p of sorted) {
    const last = runs[runs.length - 1];
    if (last && p === last[1] + 1) last[1] = p;
    else runs.push([p, p]);
  }
  const parts = runs.map(([a, b]) => paraLabel(a, b));
  return parts.length > max ? `${parts.slice(0, max).join(", ")} …` : parts.join(", ");
}
