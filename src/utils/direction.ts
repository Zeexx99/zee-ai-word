/**
 * Text direction detection for text Zee AI writes into Word.
 *
 * Arabic (and other right-to-left scripts) → right-to-left, right-aligned.
 * English / Latin → left-to-right, left-aligned.
 * Headings are centred (handled by the callers).
 *
 * Mixed technical Arabic ("تم تركيب مضخة ESP بقدرة 150 HP") stays Arabic: a paragraph is RTL when
 * it starts with an RTL word or has at least as many RTL words as Latin ones. Pure and unit-tested.
 */

export type TextDirection = "rtl" | "ltr";

/** Hebrew, Arabic, Syriac, Thaana, NKo, Samaritan, Mandaic, Arabic Extended + presentation forms. */
const RTL_CHAR = /[֐-ࣿיִ-﷿ﹰ-ﻼ]/;
/** Latin, Greek and Cyrillic letters. */
const LTR_CHAR = /[A-Za-zÀ-ɏͰ-ϿЀ-ӿ]/;

/** Strips Markdown syntax that would skew the count (link URLs, inline code, list / heading markers). */
function stripMarkup(text: string): string {
  return text
    .replace(/`[^`]*`/g, " ")
    .replace(/\]\((https?:\/\/[^\s)]+)\)/g, "]")
    .replace(/https?:\/\/\S+/g, " ");
}

/** Direction of one paragraph, or null when it has no letters (numbers, symbols, empty). */
export function detectDirection(text: string): TextDirection | null {
  const words = stripMarkup(text).split(/\s+/);
  let rtl = 0;
  let ltr = 0;
  let first: TextDirection | null = null;
  for (const w of words) {
    // A word counts for the script of its first letter.
    for (const ch of w) {
      if (RTL_CHAR.test(ch)) {
        rtl++;
        first ??= "rtl";
        break;
      }
      if (LTR_CHAR.test(ch)) {
        ltr++;
        first ??= "ltr";
        break;
      }
    }
  }
  if (!rtl && !ltr) return null;
  if (!rtl) return "ltr";
  if (!ltr) return "rtl";
  return first === "rtl" || rtl >= ltr ? "rtl" : "ltr";
}

/** True when the text contains any right-to-left letter. */
export function hasRtl(text: string): boolean {
  return RTL_CHAR.test(text);
}

/** Word paragraph alignment for a direction. */
export function alignmentFor(dir: TextDirection): "Right" | "Left" {
  return dir === "rtl" ? "Right" : "Left";
}
