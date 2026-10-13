/** Document outline for the section navigator. Pure – unit tested. */
import { countWords } from "./text";

export interface OutlineInput {
  text: string;
  /** Heading level 1–9, or null for body text. */
  level: number | null;
}

export interface Section {
  /** 1-based paragraph numbers (as numbered by getDocumentParagraphs). */
  start: number;
  end: number;
  level: number;
  title: string;
  /** Words in the whole section, sub-sections included. */
  words: number;
}

/** A heading's section runs until the next heading of the same or a higher level (or the end). */
export function buildSections(items: OutlineInput[]): Section[] {
  const out: Section[] = [];
  items.forEach((p, i) => {
    if (p.level === null || p.level < 1 || !p.text.trim()) return;
    let end = items.length;
    for (let k = i + 1; k < items.length; k++) {
      const l = items[k].level;
      if (l !== null && l >= 1 && l <= (p.level as number) && items[k].text.trim()) {
        end = k;
        break;
      }
    }
    // Trim trailing empty paragraphs.
    while (end - 1 > i && !items[end - 1].text.trim()) end--;
    const words = items.slice(i, end).reduce((n, x) => n + countWords(x.text), 0);
    out.push({ start: i + 1, end, level: p.level, title: p.text.trim().slice(0, 140), words });
  });
  return out;
}

/** Heading level from Word's style info: built-in "Heading1"…"Heading9" / "Title", localized style names, or the outline level. */
export function headingLevel(styleBuiltIn: string | undefined, styleName: string | undefined, outlineLevel?: number | string | null): number | null {
  const b = String(styleBuiltIn ?? "");
  const m = b.match(/^Heading(\d)$/i);
  if (m) return Number(m[1]);
  if (/^Title$/i.test(b)) return 1;
  const n = String(styleName ?? "").trim();
  const nm = n.match(/^(?:heading|عنوان)\s*(\d)$/i);
  if (nm) return Number(nm[1]);
  const ol = typeof outlineLevel === "number" ? outlineLevel : typeof outlineLevel === "string" ? Number(String(outlineLevel).replace(/\D/g, "")) : NaN;
  if (Number.isFinite(ol) && ol >= 1 && ol <= 9) return ol;
  return null;
}
