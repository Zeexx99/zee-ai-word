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
