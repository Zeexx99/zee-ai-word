/**
 * Minimal Markdown → HTML converter for inserting AI text into Word with insertHtml().
 * Covers what Gemini writes in documents: headings, paragraphs, **bold**, *italic*, `code`,
 * links, bullet / numbered lists (nested by indentation), block quotes, code fences, GFM tables
 * and horizontal rules. Everything is HTML-escaped; no raw HTML from the model is passed through.
 * With `autoDirection`, every block gets its direction from its own text: Arabic → dir="rtl" and
 * right-aligned, English → left-aligned, headings centred. Pure and unit-tested.
 */

import { detectDirection, type TextDirection } from "./direction";

export interface HtmlOptions {
  /** Body font of the insertion point, so inserted text matches the document (Word defaults to Times New Roman otherwise). */
  fontFamily?: string;
  /** Body font size in points. */
  fontSizePt?: number;
  /** Arabic paragraphs right-to-left + right-aligned, English left-aligned, headings centred. */
  autoDirection?: boolean;
}

/** Builds ` dir="…" style="…"` for one block. `align` "auto" = right for RTL text, left for LTR. */
function blockAttrs(styles: string[], text: string, auto: boolean, align: "auto" | "center" | "none" = "auto"): string {
  const css = [...styles];
  let dir: TextDirection | null = null;
  if (auto) {
    dir = detectDirection(text);
    if (dir === "rtl") css.push("direction:rtl");
    if (align === "center") css.push("text-align:center");
    else if (align === "auto" && dir) css.push(`text-align:${dir === "rtl" ? "right" : "left"}`);
  }
  return (dir === "rtl" ? ' dir="rtl"' : dir === "ltr" ? ' dir="ltr"' : "") + (css.length ? ` style="${css.join(";")}"` : "");
}

export function escapeHtml(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

/** Inline formatting on already-escaped text. */
export function inline(raw: string): string {
  const codes: string[] = [];
  // Protect inline code first so its contents are not formatted.
  let s = escapeHtml(raw).replace(/`([^`]+)`/g, (_, c: string) => {
    codes.push(`<code style="font-family:Consolas,'Courier New',monospace">${c}</code>`);
    return `\u0000${codes.length - 1}\u0000`;
  });
  s = s
    .replace(/\[([^\]]+)\]\((https?:\/\/[^\s)]+)\)/g, (_, t: string, u: string) => `<a href="${u}">${t}</a>`)
    .replace(/\*\*\*([^*]+)\*\*\*/g, "<b><i>$1</i></b>")
    .replace(/\*\*([^*]+)\*\*/g, "<b>$1</b>")
    .replace(/__([^_]+)__/g, "<b>$1</b>")
    .replace(/(^|[^*\w])\*([^*\s][^*]*?)\*(?!\w)/g, "$1<i>$2</i>")
    .replace(/(^|[^_\w])_([^_\s][^_]*?)_(?!\w)/g, "$1<i>$2</i>")
    .replace(/~~([^~]+)~~/g, "<s>$1</s>");
  return s.replace(/\u0000(\d+)\u0000/g, (_, i: string) => codes[Number(i)]);
}

const isTableSep = (line: string) => /^\s*\|?\s*:?-{2,}:?\s*(\|\s*:?-{2,}:?\s*)*\|?\s*$/.test(line);
const splitRow = (line: string) =>
  line
    .trim()
    .replace(/^\|/, "")
    .replace(/\|$/, "")
    .split(/(?<!\\)\|/)
    .map((c) => c.trim().replace(/\\\|/g, "|"));

interface ListItem {
  indent: number;
  ordered: boolean;
  text: string;
}

function renderList(items: ListItem[], pStyle: string, liAttrs?: (text: string) => string, listDir = ""): string {
  let html = "";
  const stack: Array<{ indent: number; tag: string }> = [];
  for (const it of items) {
    const tag = it.ordered ? "ol" : "ul";
    while (stack.length && it.indent < stack[stack.length - 1].indent) html += `</li></${stack.pop()!.tag}>`;
    const top = stack[stack.length - 1];
    if (!top || it.indent > top.indent) {
      html += `<${tag}${listDir}>`;
      stack.push({ indent: it.indent, tag });
    } else {
      html += "</li>";
      if (top.tag !== tag) {
        html += `</${top.tag}><${tag}${listDir}>`;
        top.tag = tag;
      }
    }
    html += `<li${liAttrs ? liAttrs(it.text) : pStyle}>${inline(it.text)}`;
  }
  while (stack.length) html += `</li></${stack.pop()!.tag}>`;
  return html;
}

export function markdownToHtml(md: string, opts: HtmlOptions = {}): string {
  const styleParts: string[] = [];
  if (opts.fontFamily) styleParts.push(`font-family:'${opts.fontFamily.replace(/'/g, "")}'`);
  if (opts.fontSizePt && Number.isFinite(opts.fontSizePt)) styleParts.push(`font-size:${opts.fontSizePt}pt`);
  const pStyle = styleParts.length ? ` style="${styleParts.join(";")}"` : "";
  const auto = !!opts.autoDirection;
  /** Attributes of a body paragraph (font + direction when autoDirection is on). */
  const pAttrs = (text: string) => (auto ? blockAttrs(styleParts, text, true) : pStyle);

  const lines = md.replace(/\r\n?/g, "\n").split("\n");
  const out: string[] = [];
  let para: string[] = [];
  const flushPara = () => {
    if (para.length) out.push(`<p${pAttrs(para.join(" "))}>${para.map(inline).join("<br>")}</p>`);
    para = [];
  };

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];

    // Code fence
    const fence = line.match(/^\s*(```|~~~)/);
    if (fence) {
      flushPara();
      const body: string[] = [];
      i++;
      while (i < lines.length && !lines[i].trim().startsWith(fence[1])) body.push(lines[i++]);
      out.push(`<pre style="font-family:Consolas,'Courier New',monospace;font-size:10pt">${escapeHtml(body.join("\n"))}</pre>`);
      continue;
    }

    if (!line.trim()) {
      flushPara();
      continue;
    }

    // Heading
    const h = line.match(/^\s{0,3}(#{1,6})\s+(.*?)\s*#*\s*$/);
    if (h) {
      flushPara();
      const level = h[1].length;
      out.push(`<h${level}${auto ? blockAttrs([], h[2], true, "center") : ""}>${inline(h[2])}</h${level}>`);
      continue;
    }

    // Horizontal rule
    if (/^\s{0,3}([-*_])(\s*\1){2,}\s*$/.test(line)) {
      flushPara();
      out.push("<hr>");
      continue;
    }

    // Table
    if (line.includes("|") && i + 1 < lines.length && isTableSep(lines[i + 1])) {
      flushPara();
      const header = splitRow(line);
      i += 2;
      const rows: string[][] = [];
      while (i < lines.length && lines[i].includes("|") && lines[i].trim()) rows.push(splitRow(lines[i++]));
      i--;
      const cellCss = ["border:1px solid #999", "padding:3pt 6pt"];
      const cell = (tag: string, c: string) =>
        auto
          ? `<${tag}${blockAttrs(cellCss, c, true)}><span${pStyle}>${inline(c)}</span></${tag}>`
          : `<${tag} style="${cellCss.join(";")}"><span${pStyle}>${inline(c)}</span></${tag}>`;
      const tableRtl = auto && detectDirection([header, ...rows].flat().join(" ")) === "rtl";
      out.push(
        `<table${tableRtl ? ' dir="rtl"' : ""} style="border-collapse:collapse${tableRtl ? ";direction:rtl" : ""}"><thead><tr>${header.map((c) => cell("th", c)).join("")}</tr></thead><tbody>` +
          rows.map((r) => `<tr>${header.map((_, ci) => cell("td", r[ci] ?? "")).join("")}</tr>`).join("") +
          "</tbody></table>"
      );
      continue;
    }

    // Lists
    if (/^\s*([-*+]|\d+[.)])\s+/.test(line)) {
      flushPara();
      const items: ListItem[] = [];
      while (i < lines.length) {
        const m = lines[i].match(/^(\s*)([-*+]|\d+[.)])\s+(.*)$/);
        if (m) {
          items.push({ indent: m[1].replace(/\t/g, "    ").length, ordered: /\d/.test(m[2]), text: m[3] });
          i++;
        } else if (lines[i].trim() && /^\s{2,}/.test(lines[i]) && items.length) {
          items[items.length - 1].text += ` ${lines[i].trim()}`; // continuation line
          i++;
        } else break;
      }
      i--;
      const listRtl = auto && detectDirection(items.map((it) => it.text).join(" ")) === "rtl";
      out.push(renderList(items, pStyle, auto ? pAttrs : undefined, listRtl ? ' dir="rtl"' : ""));
      continue;
    }

    // Block quote
    if (/^\s*>/.test(line)) {
      flushPara();
      const q: string[] = [];
      while (i < lines.length && /^\s*>/.test(lines[i])) q.push(lines[i++].replace(/^\s*>\s?/, ""));
      i--;
      const quoteRtl = auto && detectDirection(q.join(" ")) === "rtl";
      out.push(
        `<blockquote style="${quoteRtl ? "margin-right" : "margin-left"}:18pt;font-style:italic"><p${pAttrs(q.join(" "))}>${q.map(inline).join("<br>")}</p></blockquote>`
      );
      continue;
    }

    para.push(line.trim());
  }
  flushPara();
  return out.join("");
}

/**
 * Plain text → one HTML paragraph per line, each with its own direction and alignment
 * (used when inserting plain text with autoDirection on). Blank lines are dropped.
 */
export function plainToHtml(text: string, opts: HtmlOptions = {}): string {
  const styleParts: string[] = [];
  if (opts.fontFamily) styleParts.push(`font-family:'${opts.fontFamily.replace(/'/g, "")}'`);
  if (opts.fontSizePt && Number.isFinite(opts.fontSizePt)) styleParts.push(`font-size:${opts.fontSizePt}pt`);
  return text
    .replace(/\r\n?/g, "\n")
    .split("\n")
    .filter((l) => l.trim())
    .map((l) => `<p${blockAttrs(styleParts, l, opts.autoDirection !== false)}>${escapeHtml(l)}</p>`)
    .join("");
}

/** Markdown → readable plain text (for "plain text" insertion and the clipboard). */
export function markdownToPlain(md: string): string {
  return md
    .replace(/\r\n?/g, "\n")
    .replace(/^\s*(```|~~~).*$/gm, "")
    .replace(/^\s{0,3}#{1,6}\s+/gm, "")
    .replace(/\*\*([^*]+)\*\*/g, "$1")
    .replace(/__([^_]+)__/g, "$1")
    .replace(/(^|[^*\w])\*([^*\s][^*]*?)\*(?!\w)/g, "$1$2")
    .replace(/`([^`]+)`/g, "$1")
    .replace(/\[([^\]]+)\]\((https?:\/\/[^\s)]+)\)/g, "$1 ($2)")
    .replace(/^\s*[-*+]\s+/gm, "• ")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

/** True when the text uses Markdown syntax worth rendering (otherwise insert as plain text). */
export function looksLikeMarkdown(s: string): boolean {
  return /(^|\n)\s{0,3}(#{1,6}\s|[-*+]\s|\d+[.)]\s|>|```)|\*\*[^*]+\*\*|\|.*\|\s*\n\s*\|?\s*:?-{2,}/.test(s);
}
