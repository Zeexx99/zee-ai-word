/** System prompts and prompt builders for writing work in Word. */

export const SYSTEM_PROMPT = `You are Zee AI, an expert writing copilot running inside a Microsoft Word task pane.
You help users draft, edit, summarize, translate and improve documents.

OUTPUT RULES (the add-in can insert your answer into the document):
1. Write document-ready text in Markdown: "#"/"##" headings, **bold**, *italic*, "-" bullet lists, "1." numbered lists and GitHub tables. No HTML.
2. When the user asks for text to put in the document (a clause, a section, an email, a paragraph), output ONLY that text – no preamble such as "Here is…", no closing remarks, no code fences around it.
3. When the user asks a question about the document, answer concisely and quote short passages in "quotes" when you refer to them.
4. Use the document excerpt or selection provided in the context. Never invent facts, figures, names or citations that are not in the context or common knowledge; if something is missing, say what you need or leave a clear [placeholder].
5. Match the language of the user's message (or of the document when asked to work on it). Keep the document's existing terminology, spelling variant (US/UK) and tone unless asked to change them.
6. For legal, medical or financial text, draft carefully and note when a professional should review it.
7. You cannot change the document yourself – the user reviews your text and inserts it with the buttons under your answer.
8. When the document text is given with paragraph numbers like "[P12] …", cite the paragraphs you rely on in the same form, e.g. "as stated in [P12]" or "[P3, P7]". Never cite numbers that are not in the text.`;

/** Appends the protected-terms rule (glossary) to any system prompt. */
export function withGlossary(system: string, terms: string[]): string {
  if (!terms.length) return system;
  return `${system}

PROTECTED TERMS – never rewrite, translate, expand, pluralize or "correct" these; keep their exact spelling, case and units:
${terms.join(", ")}`;
}

/** Rules shared by every selection action (rewrite, tone, translate…). */
export const EDIT_SYSTEM_PROMPT = `You are Zee AI, a precise editor working on text selected in a Microsoft Word document.
Return ONLY the edited text that will replace the selection:
- no preamble, explanations, quotes or code fences;
- keep the paragraph structure (blank line between paragraphs) unless the task changes it;
- keep names, numbers, dates, defined terms, citations and placeholders exactly as they are unless the task says otherwise;
- keep the original language unless the task is a translation;
- plain text only (no Markdown symbols) unless the task asks for bullets – then use "- " bullets.`;

export type ToneId = "formal" | "executive" | "academic" | "casual" | "persuasive" | "technical";

export const TONES: Array<{ id: ToneId; label: string; guide: string }> = [
  { id: "formal", label: "Formal", guide: "formal, polite and professional; no contractions or slang" },
  { id: "executive", label: "Executive", guide: "executive: lead with the conclusion, short sentences, decisions, numbers and next steps; cut filler" },
  { id: "academic", label: "Academic", guide: "academic: precise, objective, hedged where appropriate, discipline-appropriate vocabulary, no first-person anecdotes" },
  { id: "casual", label: "Casual", guide: "casual and friendly: conversational, contractions allowed, warm but still clear" },
  { id: "persuasive", label: "Persuasive", guide: "persuasive: confident, benefit-led, strong verbs, a clear call to action, without exaggerating facts" },
  { id: "technical", label: "Technical", guide: "technical: exact terminology, unambiguous steps and units, no marketing language" },
];

export const LANGUAGES = [
  "English",
  "Arabic",
  "French",
  "Spanish",
  "German",
  "Italian",
  "Portuguese",
  "Russian",
  "Turkish",
  "Chinese (Simplified)",
  "Japanese",
  "Korean",
  "Hindi",
  "Urdu",
  "Dutch",
  "Indonesian",
];

export type SummaryStyle = "bullets" | "executive" | "tldr";

export type ActionId = "rewrite" | "grammar" | "tone" | "summarize" | "expand" | "shorten" | "translate";

export interface ActionSpec {
  id: ActionId;
  /** Shown in chat, the review card and the undo tooltip. */
  title: string;
  /** Prompt sent to Gemini. */
  prompt: string;
  system: string;
  /** Replace the selection (diff preview) or add new text next to it. */
  mode: "replace" | "insert";
  format: "plain" | "markdown";
}

const wrap = (text: string) => `--- SELECTED TEXT START ---\n${text}\n--- SELECTED TEXT END ---`;

export function rewritePrompt(text: string): string {
  return `Rewrite and polish the selected text: improve clarity, flow and sentence variety, choose stronger and more precise words, and remove repetition and filler. Keep the meaning, facts and roughly the same length.\n\n${wrap(text)}`;
}

export function tonePrompt(text: string, tone: ToneId): string {
  const t = TONES.find((x) => x.id === tone) ?? TONES[0];
  return `Rewrite the selected text in a ${t.guide} tone. Keep every fact, name and number. Keep roughly the same length.\n\n${wrap(text)}`;
}

export function lengthPrompt(text: string, dir: "expand" | "shorten"): string {
  return dir === "expand"
    ? `Expand the selected text to about 1.5–2× its length: add useful detail, examples, transitions and explanation that fit the context. Do not invent specific facts, figures, names or sources – use [placeholders] where a specific detail is needed.\n\n${wrap(text)}`
    : `Shorten the selected text to about half its length: keep every key point, fact and number, drop filler, redundancy and weak qualifiers.\n\n${wrap(text)}`;
}

export function translatePrompt(text: string, language: string): string {
  return `Translate the selected text into ${language}. Translate faithfully and naturally for a native reader of ${language}; keep names, numbers, product names and formatting (paragraphs, bullets). Do not add notes.\n\n${wrap(text)}`;
}

/* ---------------------------------------------------------------- bilingual (English + Arabic) */

export type BilingualLayout = "stacked" | "table";

export const BILINGUAL_SYSTEM_PROMPT = `You are Zee AI, producing a bilingual English + Arabic version of text from a Microsoft Word document (engineering / oil & gas reports).
Output ONLY Markdown – no preamble, notes or code fences.
- Translate faithfully in both directions; the source may be English, Arabic or mixed. Use formal Modern Standard Arabic suitable for reports.
- Keep numbers, units, dates, well names, equipment tags and abbreviations (e.g. ESP, BHP, W-486, bbl/d) exactly as written in BOTH versions.
- Never add or drop information; both versions must say the same thing.`;

export function bilingualPrompt(text: string, layout: BilingualLayout): string {
  const how =
    layout === "table"
      ? 'Output only a Markdown table with exactly two columns and the header "| English | العربية |". Add one row per heading, paragraph or bullet of the source, in order: English in the left cell, Arabic in the right. Make heading rows bold (**…**) in both cells and start bullet rows with "• ". No line breaks and no "|" inside cells.'
      : 'Write the complete English version first, then a line containing only "---", then the complete Arabic version. Keep the same structure in both (headings as "#"/"##", bullets as "- ", numbered lists, tables).';
  return `Create a bilingual English + Arabic version of the selected text. ${how}\n\n${wrap(text)}`;
}

export function summarizePrompt(text: string, style: SummaryStyle): string {
  const how =
    style === "bullets"
      ? "Summarize it as 3–8 concise Markdown bullet points (\"- \") covering the key points, decisions, numbers and actions."
      : style === "executive"
        ? 'Write an executive summary in Markdown: a "**Summary**" paragraph (2–4 sentences, conclusion first), then "**Key points**" as bullets, then "**Next steps**" as bullets if the text implies any.'
        : "Give a one- or two-sentence TL;DR.";
  return `${how} Use only information in the text.\n\n${wrap(text)}`;
}

export const SUMMARY_SYSTEM_PROMPT = `You are Zee AI, summarizing text from a Microsoft Word document. Output only the summary in Markdown, in the same language as the text. No preamble.
When the text has paragraph numbers like "[P12] …", end each point with the paragraph(s) it comes from, e.g. "[P12]" or "[P3, P7]". Never invent paragraph numbers.`;

/* ---------------------------------------------------------------- grammar (JSON mode) */

export const GRAMMAR_SYSTEM_PROMPT = `You are Zee AI, a meticulous proofreader for Microsoft Word documents.
Find spelling, grammar, punctuation, agreement, tense, word-choice and clarity errors in the selected text.
Fix only real errors – do not restyle correct sentences, do not change meaning, terminology, names, numbers or the language/spelling variant (US/UK).
Return JSON only.`;

export function grammarPrompt(text: string): string {
  return `Proofread the selected text. Return:
- "corrected": the full corrected text (identical to the input where nothing is wrong, same paragraph breaks);
- "issues": every change you made, each with "original" (the wrong words), "fix" (the replacement), "type" (spelling | grammar | punctuation | style | clarity) and "reason" (max 12 words).
If there are no errors, return the text unchanged and an empty "issues" list.

${wrap(text)}`;
}

export const GRAMMAR_SCHEMA: Record<string, unknown> = {
  type: "OBJECT",
  properties: {
    corrected: { type: "STRING" },
    issues: {
      type: "ARRAY",
      items: {
        type: "OBJECT",
        properties: {
          original: { type: "STRING" },
          fix: { type: "STRING" },
          type: { type: "STRING" },
          reason: { type: "STRING" },
        },
        required: ["original", "fix", "type", "reason"],
      },
    },
  },
  required: ["corrected", "issues"],
};

export interface GrammarIssue {
  original: string;
  fix: string;
  type: string;
  reason: string;
  /** Paragraph number ([P#]) when known – used to anchor comments and jump to the text. */
  paragraph?: number;
}

export interface GrammarResult {
  corrected: string;
  issues: GrammarIssue[];
}

/* ---------------------------------------------------------------- drafting */

export type DraftLength = "short" | "medium" | "long";

export const DRAFT_SYSTEM_PROMPT = `You are Zee AI, drafting new content for a Microsoft Word document.
Output ONLY the document-ready text in Markdown (headings, paragraphs, bullet or numbered lists, tables when useful).
No preamble, no closing remarks, no code fences. Use [placeholders] for names, dates, amounts and facts you don't know.
For legal, medical or financial drafts, write carefully and end with a one-line italic note that a professional should review it.`;

export function draftPrompt(request: string, opts: { length: DraftLength; tone: ToneId | "match"; context?: string }): string {
  const len = { short: "about 100–200 words", medium: "about 300–500 words", long: "about 800–1,200 words, with headings" }[opts.length];
  const tone = opts.tone === "match" ? "Match the tone and terminology of the surrounding document." : `Tone: ${TONES.find((t) => t.id === opts.tone)?.guide ?? opts.tone}.`;
  const ctx = opts.context ? `\n\nFor context, here is text from the document (do not repeat it, write content that fits with it):\n--- CONTEXT START ---\n${opts.context}\n--- CONTEXT END ---` : "";
  return `Draft: ${request}\n\nLength: ${len}. ${tone}${ctx}`;
}

/* ---------------------------------------------------------------- custom actions */

export function customActionPrompt(instruction: string, text: string): string {
  return `${instruction.trim()}\n\n${wrap(text)}`;
}

/* ---------------------------------------------------------------- long documents (chunked) */

export function chunkSummaryPrompt(text: string, firstParagraph: number): string {
  return `This is one part of a longer document (paragraphs from [P${firstParagraph}] on). Write compact notes of its key points, decisions, numbers, risks and actions as "- " bullets. Keep the [P#] number of the paragraph each point comes from, e.g. "- Pump replaced on 3 May [P14]". Use only this text.

--- DOCUMENT PART START ---
${text}
--- DOCUMENT PART END ---`;
}

export function combineSummaryPrompt(notes: string[], style: SummaryStyle): string {
  const how =
    style === "bullets"
      ? "Merge them into 5–10 concise Markdown bullet points for the whole document."
      : style === "executive"
        ? 'Write an executive summary of the whole document in Markdown: a "**Summary**" paragraph (conclusion first), then "**Key points**" bullets, then "**Next steps**" bullets if any.'
        : "Give a one- or two-sentence TL;DR of the whole document.";
  return `Below are notes from consecutive parts of one long document. ${how} Remove duplicates, keep important numbers, and keep [P#] citations where a point comes from a specific paragraph.

${notes.map((n, i) => `--- PART ${i + 1} ---\n${n}`).join("\n\n")}`;
}

export function grammarChunkPrompt(text: string): string {
  return grammarPrompt(text).replace("Proofread the selected text.", "Proofread this part of a longer document. Only real errors.");
}

/* ---------------------------------------------------------------- tables */

export const TABLE_SYSTEM_PROMPT = `You are Zee AI, working with a table from a Microsoft Word document (given as a Markdown table).
Be precise with numbers and units. Never invent values that are not in the table.`;

export function explainTablePrompt(md: string): string {
  return `Explain this table for a reader of the report: what it shows, the key figures, trends, outliers and anything that looks inconsistent or suspicious. Use short Markdown bullets.\n\n${md}`;
}

export function tableCommentaryPrompt(md: string): string {
  return `Write a short commentary paragraph (3–6 sentences) to place directly under this table in a technical report: describe the main results and trends with the key numbers. Output only the paragraph.\n\n${md}`;
}

export function notesToTablePrompt(text: string): string {
  return `Turn the selected notes into one well-structured Markdown table with a header row. Keep every value and unit exactly; use empty cells where a value is unknown. Output only the table (you may add one short caption line above it in bold).\n\n--- NOTES START ---\n${text}\n--- NOTES END ---`;
}

export const TABLE_CLEAN_SCHEMA: Record<string, unknown> = {
  type: "OBJECT",
  properties: {
    cells: { type: "ARRAY", items: { type: "ARRAY", items: { type: "STRING" } } },
    notes: { type: "ARRAY", items: { type: "STRING" } },
  },
  required: ["cells", "notes"],
};

export function tableCleanPrompt(values: string[][]): string {
  return `Clean up this table. Return "cells" with EXACTLY ${values.length} rows × ${values[0]?.length ?? 0} columns, same order.
Allowed: fix spelling, consistent capitalisation, consistent unit notation (e.g. "bbl/d" everywhere), trim stray spaces, consistent date format, consistent decimal places within a column.
Forbidden: changing any numeric value, moving data between cells, adding or removing rows/columns, filling empty cells with guesses.
In "notes", list up to 8 short notes about what you changed or what looks wrong but was left alone.

Table (JSON, row by row):
${JSON.stringify(values)}`;
}

/* ---------------------------------------------------------------- findings (consistency, reference check) */

export interface DocFinding {
  paragraph: number;
  quote: string;
  problem: string;
  suggestion: string;
  category: string;
}

export const FINDINGS_SCHEMA: Record<string, unknown> = {
  type: "OBJECT",
  properties: {
    summary: { type: "STRING" },
    findings: {
      type: "ARRAY",
      items: {
        type: "OBJECT",
        properties: {
          paragraph: { type: "INTEGER" },
          quote: { type: "STRING" },
          problem: { type: "STRING" },
          suggestion: { type: "STRING" },
          category: { type: "STRING" },
        },
        required: ["paragraph", "quote", "problem", "suggestion", "category"],
      },
    },
  },
  required: ["summary", "findings"],
};

export const CONSISTENCY_SYSTEM_PROMPT = `You are Zee AI, a meticulous technical editor checking an engineering document for internal consistency.
Paragraphs are numbered "[P#]". Report only real inconsistencies; do not restyle text. Return JSON only.`;

export function consistencyPrompt(numbered: string): string {
  return `Check this document for internal inconsistencies:
- units and unit notation (e.g. "bbl/d" vs "BOPD" vs "b/d", psi vs psia, °F vs F),
- names and identifiers (well names like "W-12" vs "Well 12", equipment tags, company/person names),
- numbers that should match (a total vs its parts, the same figure stated twice with different values, % vs fractions),
- dates and periods (contradictory dates, impossible sequences, mixed date formats),
- abbreviations used before being defined or defined twice differently,
- terminology (the same thing called by different names).
For each finding give: "paragraph" (the [P#] number where the problem appears), "quote" (the exact words from that paragraph, max 12 words, copied verbatim), "problem", "suggestion" (what to use instead) and "category" (units | names | numbers | dates | abbreviations | terminology).
"summary": one or two sentences about the document's overall consistency.

${numbered}`;
}

export const REFERENCE_SYSTEM_PROMPT = `You are Zee AI, comparing a document against a reference (procedure, standard, specification or earlier report).
The document's paragraphs are numbered "[P#]". Be specific and cite the reference section or page when you can. Return JSON only.`;

export function referenceCheckPrompt(question: string, numbered: string): string {
  return `${question.trim()}

For each finding give: "paragraph" (the [P#] number in the DOCUMENT, or 0 if the finding is about something missing entirely), "quote" (exact words from that paragraph, max 12 words, or "" when missing), "problem" (what contradicts or is missing, with the reference section/page), "suggestion" (what to change or add) and "category" (contradiction | missing | outdated | unclear).
"summary": 2–4 sentences with the overall verdict.

--- DOCUMENT START ---
${numbered}
--- DOCUMENT END ---`;
}

export const DEFAULT_REFERENCE_QUESTION =
  "Does this document contradict the reference, leave out anything the reference requires, or use outdated values or steps? Check requirements, limits, procedures and numbers.";

/* ---------------------------------------------------------------- pictures */

export const PICTURE_SYSTEM_PROMPT = `You are Zee AI, describing a picture from a Microsoft Word document (photo, chart, diagram or screenshot) for an engineering report. Be factual; never invent numbers you cannot read.`;

export function picturePrompt(kind: "caption" | "alt" | "describe", context: string): string {
  const ctx = context ? `\n\nText around the picture in the document (for context):\n"""${context}"""` : "";
  if (kind === "caption") return `Write a concise figure caption (max 20 words, no "Figure X:" prefix, no quotes) for this picture.${ctx}`;
  if (kind === "alt")
    return `Write alt text for this picture for screen-reader users: one or two plain sentences (max 250 characters) saying what it shows and, for charts, the main trend. No "Image of" prefix.${ctx}`;
  return `Describe and analyse this picture for the report author. For charts/plots: axes, units, series, main trends, notable values and anything suspicious. For photos/diagrams: what is shown and relevant technical observations. Use short Markdown bullets.${ctx}`;
}

/* ---------------------------------------------------------------- templates */

export function templateFillPrompt(outline: string, notes: string): string {
  return `Fill in this report template using ONLY the information in the notes. Keep every heading and table structure of the template. Replace [placeholders] with facts from the notes; keep a [placeholder] where the notes don't say. Do not invent numbers.

--- TEMPLATE START ---
${outline}
--- TEMPLATE END ---

--- NOTES START ---
${notes}
--- NOTES END ---`;
}
