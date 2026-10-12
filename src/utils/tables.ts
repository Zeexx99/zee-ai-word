/** Word table helpers (pure, unit-tested). */

export type Grid = string[][];

const clean = (v: unknown) => String(v ?? "").replace(/[\r\n\u000b\u0007]+/g, " ").trim();

/** A table as a GitHub-flavoured Markdown table (first row = header). */
export function tableToMarkdown(values: Grid): string {
  if (!values.length) return "";
  const cols = Math.max(...values.map((r) => r.length));
  const row = (r: string[]) => `| ${Array.from({ length: cols }, (_, i) => clean(r[i]).replace(/\|/g, "\\|")).join(" | ")} |`;
  return [row(values[0]), `|${" --- |".repeat(cols)}`, ...values.slice(1).map(row)].join("\n");
}

/** Force a model's answer into exactly rows × cols (missing cells keep the original value). */
export function fitGrid(original: Grid, proposed: unknown): Grid {
  const rows = Array.isArray(proposed) ? proposed : [];
  return original.map((r, i) =>
    r.map((cell, j) => {
      const p = Array.isArray(rows[i]) ? (rows[i] as unknown[])[j] : undefined;
      return p === undefined || p === null ? cell : clean(p);
    })
  );
}

export interface CellChange {
  row: number;
  col: number;
  before: string;
  after: string;
}

/** Cells whose text differs (whitespace-insensitive). */
export function diffGrid(before: Grid, after: Grid): CellChange[] {
  const out: CellChange[] = [];
  before.forEach((r, i) =>
    r.forEach((cell, j) => {
      const a = after[i]?.[j] ?? cell;
      if (clean(cell) !== clean(a)) out.push({ row: i, col: j, before: clean(cell), after: clean(a) });
    })
  );
  return out;
}

/** Numbers that appear in a cell (used to make sure a "clean-up" never changes values). */
export function numbersIn(s: string): string[] {
  return (s.replace(/,(?=\d{3}\b)/g, "").match(/-?\d+(?:\.\d+)?/g) ?? []).map((n) => String(Number(n)));
}

/** True when every number in `before` is still present in `after` (order-insensitive). */
export function keepsNumbers(before: string, after: string): boolean {
  const a = numbersIn(after);
  return numbersIn(before).every((n) => {
    const i = a.indexOf(n);
    if (i < 0) return false;
    a.splice(i, 1);
    return true;
  });
}
