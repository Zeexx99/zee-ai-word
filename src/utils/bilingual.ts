/**
 * Arabic ⇄ English consistency helpers: language of a paragraph, numbers in both scripts, and local
 * pairing of English / Arabic counterparts (stacked paragraphs or side-by-side table cells). Pure – unit tested.
 */

export type Lang = "ar" | "en" | "mixed" | "none";

const AR = /[؀-ۿݐ-ݿࢠ-ࣿﭐ-﷿ﹰ-﻿]/g;
const LATIN = /[A-Za-z]/g;

export function langOf(text: string): Lang {
  const ar = (text.match(AR) ?? []).length;
  const en = (text.match(LATIN) ?? []).length;
  if (!ar && !en) return "none";
  if (ar >= en * 2) return "ar";
  if (en >= ar * 3) return "en";
  return "mixed";
}

/** Arabic-Indic / Persian digits → 0-9, Arabic decimal / thousands separators → "." / "". */
export function normalizeDigits(text: string): string {
  return text
    .replace(/[٠-٩]/g, (d) => String(d.charCodeAt(0) - 0x0660))
    .replace(/[۰-۹]/g, (d) => String(d.charCodeAt(0) - 0x06f0))
    .replace(/٫/g, ".")
    .replace(/٬/g, "");
}

/** Numbers in a text as canonical strings ("1,250.50" → "1250.5", "٣٥٪" → "35"). Years and decimals included. */
export function extractNumbers(text: string): string[] {
  const t = normalizeDigits(text);
  const out: string[] = [];
  const re = /\d[\d,]*(?:\.\d+)?/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(t))) {
    const raw = m[0].replace(/,(?=\d{3}\b)/g, "").replace(/,/g, "");
    const n = Number(raw);
    if (Number.isFinite(n)) out.push(String(n));
  }
  return out;
}

function multisetDiff(a: string[], b: string[]): string[] {
  const left = [...b];
  const out: string[] = [];
  for (const x of a) {
    const i = left.indexOf(x);
    if (i >= 0) left.splice(i, 1);
    else out.push(x);
  }
  return out;
}

export interface PairMismatch {
  /** 1-based paragraph numbers of the English and Arabic counterparts. */
  en: number;
  ar: number;
  onlyInEnglish: string[];
  onlyInArabic: string[];
}

/**
 * Pairs consecutive English/Arabic paragraphs (either order) and reports numbers that appear in only
 * one of them. Only pairs where at least one side has numbers are compared.
 */
export function numberMismatches(paras: string[]): { pairs: number; mismatches: PairMismatch[] } {
  const mismatches: PairMismatch[] = [];
  let pairs = 0;
  let i = 0;
  const next = (k: number) => {
    let j = k + 1;
    while (j < paras.length && !paras[j].trim()) j++;
    return j;
  };
  while (i < paras.length) {
    if (!paras[i].trim()) {
      i++;
      continue;
    }
    const j = next(i);
    const a = langOf(paras[i]);
    const b = j < paras.length ? langOf(paras[j]) : "none";
    if ((a === "en" && b === "ar") || (a === "ar" && b === "en")) {
      pairs++;
      const en = a === "en" ? i : j;
      const ar = a === "ar" ? i : j;
      const ne = extractNumbers(paras[en]);
      const na = extractNumbers(paras[ar]);
      const onlyInEnglish = multisetDiff(ne, na);
      const onlyInArabic = multisetDiff(na, ne);
      if (onlyInEnglish.length || onlyInArabic.length) mismatches.push({ en: en + 1, ar: ar + 1, onlyInEnglish, onlyInArabic });
      i = next(j);
    } else i = j;
  }
  return { pairs, mismatches };
}
