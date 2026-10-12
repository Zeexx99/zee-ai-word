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

/* ------------------------------------------------------------------ paragraph OOXML (bidi) */

/** pPr children that must come AFTER <w:bidi/> (ECMA-376 order), so bidi is inserted before the first of them. */
const AFTER_BIDI = [
  "adjustRightInd", "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc",
  "textDirection", "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr", "pPrChange",
];
/** pPr children that must come AFTER <w:jc/>. */
const AFTER_JC = ["textDirection", "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr", "pPrChange"];

/**
 * Content that a paragraph-level OOXML round trip could damage (comments, notes, list numbering,
 * tracked revisions, floating shapes, text boxes, tables). Those paragraphs only get an alignment.
 */
const RISKY = /<w:(commentRangeStart|commentReference|footnoteReference|endnoteReference|numPr|ins|del|moveFrom|moveTo|pPrChange|rPrChange|tbl|txbxContent|drawing|pict|object)\b/;

export interface ParagraphXmlInfo {
  /** The paragraph is right-to-left (<w:bidi/>). */
  bidi: boolean;
  /** Not safe to rewrite through OOXML. */
  risky: boolean;
}

interface ParaSpan {
  /** Index of "<w:p" of the first body paragraph. */
  start: number;
  /** Index just after the paragraph's opening tag. */
  openEnd: number;
  /** pPr bounds inside the paragraph (pPrStart = index of "<w:pPr", pPrEnd = index after "</w:pPr>"), or -1. */
  pPrStart: number;
  pPrEnd: number;
  /** Index of the paragraph's "</w:p>". */
  end: number;
}

function firstParagraph(pkg: string): ParaSpan | null {
  const body = pkg.indexOf("<w:body>");
  if (body < 0) return null;
  const re = /<w:p(?=[\s>/])/g;
  re.lastIndex = body;
  const m = re.exec(pkg);
  if (!m) return null;
  const start = m.index;
  const openEnd = pkg.indexOf(">", start) + 1;
  if (openEnd <= 0 || pkg[openEnd - 2] === "/") return null; // <w:p/> – nothing to change
  const end = pkg.indexOf("</w:p>", openEnd);
  if (end < 0) return null;
  let pPrStart = -1;
  let pPrEnd = -1;
  const after = pkg.slice(openEnd, openEnd + 7);
  if (after.startsWith("<w:pPr>") || after.startsWith("<w:pPr ")) {
    pPrStart = openEnd;
    const close = pkg.indexOf("</w:pPr>", pPrStart);
    if (close < 0 || close > end) return null;
    pPrEnd = close + "</w:pPr>".length;
  } else if (after.startsWith("<w:pPr/")) {
    pPrStart = openEnd;
    pPrEnd = pkg.indexOf(">", pPrStart) + 1;
  }
  return { start, openEnd, pPrStart, pPrEnd, end };
}

/** Bidi state of the first paragraph of a Word paragraph OOXML package, and whether it is safe to rewrite. */
export function inspectParagraphXml(pkg: string): ParagraphXmlInfo | null {
  const p = firstParagraph(pkg);
  if (!p) return null;
  const pPr = p.pPrStart >= 0 ? pkg.slice(p.pPrStart, p.pPrEnd) : "";
  const bidi = /<w:bidi(\s*\/>|\s+w:val="(1|true|on)"\s*\/>)/.test(pPr);
  // A block content control around the paragraph (e.g. Zee AI's undo marker) would be duplicated by a rewrite.
  const wrapped = /<w:sdt[\s>]/.test(pkg.slice(pkg.indexOf("<w:body>"), p.start));
  return { bidi, risky: wrapped || RISKY.test(pkg.slice(p.start, p.end)) };
}

/** Inserts `tag` into a pPr body before the first child listed in `before` (or at the end). */
function insertInPPr(pPrInner: string, tag: string, before: string[]): string {
  let at = pPrInner.length;
  for (const name of before) {
    const i = pPrInner.search(new RegExp(`<w:${name}(?=[\\s>/])`));
    if (i >= 0 && i < at) at = i;
  }
  return pPrInner.slice(0, at) + tag + pPrInner.slice(at);
}

/**
 * Sets the direction of the first paragraph in a Word paragraph OOXML package.
 * RTL: adds <w:bidi/>; LTR: removes it. A left/right/start/end <w:jc> is removed so the paragraph
 * falls back to its natural side (right for Arabic, left for English) – Word interprets left/right
 * differently in RTL paragraphs, so an explicit side is avoided. Centred / justified alignment is kept,
 * and `center` forces centring (headings). Returns null when nothing has to change.
 */
export function setParagraphDirectionXml(pkg: string, dir: TextDirection, center = false): string | null {
  const p = firstParagraph(pkg);
  if (!p) return null;
  const hasPPr = p.pPrStart >= 0;
  const selfClosing = hasPPr && pkg.slice(p.pPrStart, p.pPrEnd).endsWith("/>");
  const openTag = hasPPr ? pkg.slice(p.pPrStart, pkg.indexOf(">", p.pPrStart) + 1).replace(/\/>$/, ">") : "<w:pPr>";
  let inner = hasPPr && !selfClosing ? pkg.slice(pkg.indexOf(">", p.pPrStart) + 1, p.pPrEnd - "</w:pPr>".length) : "";
  const original = inner;

  const BIDI_RE = /<w:bidi(\s+w:val="[^"]*")?\s*\/>/g;
  const isBidi = /<w:bidi(\s*\/>|\s+w:val="(1|true|on)"\s*\/>)/.test(inner);
  if (dir === "rtl" && !isBidi) inner = insertInPPr(inner.replace(BIDI_RE, ""), "<w:bidi/>", AFTER_BIDI);
  if (dir === "ltr") inner = inner.replace(BIDI_RE, "");

  const JC_RE = /<w:jc\s+w:val="([^"]*)"\s*\/>/;
  const jc = inner.match(JC_RE)?.[1];
  if (center) {
    if (jc !== "center") inner = insertInPPr(inner.replace(JC_RE, ""), '<w:jc w:val="center"/>', AFTER_JC);
  } else if (jc && /^(left|right|start|end)$/.test(jc) && (dir === "rtl") !== isBidi) {
    // Only when the direction actually flips: the old side belonged to the old direction.
    inner = inner.replace(JC_RE, "");
  }

  if (inner === original) return null;
  const pPr = inner ? `${openTag}${inner}</w:pPr>` : "";
  return hasPPr ? pkg.slice(0, p.pPrStart) + pPr + pkg.slice(p.pPrEnd) : pkg.slice(0, p.openEnd) + pPr + pkg.slice(p.openEnd);
}
