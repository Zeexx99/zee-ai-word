/**
 * Office.js helpers for Word: reading the selection / document, inserting AI text, formatting-
 * preserving edits, Track Changes, Word comments, paragraph navigation and the multi-step undo stack.
 *
 * Undo design: every insertion or edit is wrapped in a hidden content control tagged "zeeai-<id>".
 * Undo finds that control and either deletes it (pure insertions) or puts the original text back
 * (edits / replacements). Comments are undone by id. The stack lives in memory only; when the pane
 * reloads, leftover hidden controls are unwrapped (their text stays) so documents never accumulate
 * Zee AI markers – except around edits made with Track Changes ("zeeai-t…"): those stay until the
 * revisions are accepted or rejected, so "Accept / Reject Zee AI's changes" can find them later.
 */
import { markdownToHtml, markdownToPlain, plainToHtml } from "../utils/mdToHtml";
import { alignmentFor, detectDirection, hasRtl, inspectParagraphXml, setParagraphDirectionXml } from "../utils/direction";
import { loadSettings } from "./settingsService";
import { normalizeWordText } from "../utils/text";
import { countOccurrences, escapeWordSearch, planEdits, MAX_SEARCH } from "../utils/edits";
import { headingLevel } from "../utils/sections";

export type InsertLocation = "replace" | "after" | "end";
export type InsertFormat = "markdown" | "plain";

export const LOCATION_LABELS: Record<InsertLocation, string> = {
  replace: "Replace selection",
  after: "Insert after cursor",
  end: "Append to end of document",
};

const TAG_PREFIX = "zeeai-";
/** Markers around tracked edits keep this prefix (TAG_PREFIX + "t…") until accepted / rejected. */
const TRACKED_PREFIX = TAG_PREFIX + "t";
const MAX_UNDO = 25;
const MAX_OOXML_CHARS = 5_000_000;

export function isWordAvailable(): boolean {
  try {
    return typeof Office !== "undefined" && Office.context?.host === Office.HostType.Word && typeof Word !== "undefined";
  } catch {
    return false;
  }
}

/** WordApi requirement-set check (1.3: ranges & intersections, 1.4: comments & Track Changes, 1.6: accept / reject revisions). */
export function supportsWordApi(version: "1.3" | "1.4" | "1.6"): boolean {
  try {
    return isWordAvailable() && Office.context.requirements.isSetSupported("WordApi", version);
  } catch {
    return false;
  }
}

function requireWord() {
  if (!isWordAvailable()) throw new Error("Open Zee AI from Word to read or change the document.");
}

const sameText = (a: string, b: string) => a.replace(/\s+/g, " ").trim() === b.replace(/\s+/g, " ").trim();

/* ------------------------------------------------------------------ reading */

export interface SelectionInfo {
  text: string;
  isEmpty: boolean;
  fontName: string | null;
  fontSize: number | null;
}

export async function getSelectionInfo(): Promise<SelectionInfo> {
  requireWord();
  return Word.run(async (ctx) => {
    const sel = ctx.document.getSelection();
    sel.load("text,isEmpty");
    sel.font.load("name,size");
    await ctx.sync();
    const text = normalizeWordText(sel.text ?? "");
    return {
      text,
      isEmpty: sel.isEmpty || !text.trim(),
      // Mixed fonts come back as null/"" – the inserted text then uses the paragraph default.
      fontName: sel.font.name || null,
      fontSize: typeof sel.font.size === "number" && sel.font.size > 0 ? sel.font.size : null,
    };
  });
}

export async function getSelectionText(): Promise<string> {
  return (await getSelectionInfo()).text;
}

export async function getDocumentText(): Promise<string> {
  requireWord();
  return Word.run(async (ctx) => {
    const body = ctx.document.body;
    body.load("text");
    await ctx.sync();
    return normalizeWordText(body.text ?? "");
  });
}

/** Every paragraph of the document body, in order (empty ones included so numbering matches Word). */
export async function getDocumentParagraphs(): Promise<string[]> {
  requireWord();
  return Word.run(async (ctx) => {
    const paras = ctx.document.body.paragraphs;
    paras.load("items/text");
    await ctx.sync();
    return paras.items.map((p) => normalizeWordText(p.text ?? "").replace(/\n+/g, " ").trim());
  });
}

/**
 * Text + heading level of every paragraph, numbered exactly like getDocumentParagraphs (for the
 * section navigator).
 */
export async function getOutline(): Promise<Array<{ text: string; level: number | null }>> {
  requireWord();
  const rich = supportsWordApi("1.3");
  return Word.run(async (ctx) => {
    const paras = ctx.document.body.paragraphs;
    paras.load(rich ? "items/text,items/style,items/styleBuiltIn,items/outlineLevel" : "items/text,items/style,items/outlineLevel");
    await ctx.sync();
    return paras.items.map((p) => {
      const text = normalizeWordText(p.text ?? "").replace(/\n+/g, " ").trim();
      let outline: number | null = null;
      try {
        outline = typeof p.outlineLevel === "number" ? p.outlineLevel : null;
      } catch {
        outline = null;
      }
      return { text, level: headingLevel(rich ? String(p.styleBuiltIn ?? "") : "", p.style ?? "", outline) };
    });
  });
}

/** Inserts a picture (PNG/JPEG base64) as its own paragraph after the cursor's paragraph. Undoable. */
export async function insertPictureAfterCursor(base64: string, label: string, altText = ""): Promise<void> {
  requireWord();
  const id = uid();
  let undoable = false;
  await Word.run(async (ctx) => {
    const sel = ctx.document.getSelection();
    const p = sel.paragraphs.getLast().insertParagraph("", Word.InsertLocation.after);
    if (supportsWordApi("1.3")) p.styleBuiltIn = Word.BuiltInStyleName.normal;
    const pic = p.insertInlinePictureFromBase64(base64, Word.InsertLocation.start);
    pic.load("width,height");
    await ctx.sync();
    if (pic.width > 460) {
      const k = 460 / pic.width;
      pic.width = 460;
      pic.height = Math.round(pic.height * k);
    }
    if (altText) pic.altTextDescription = altText;
    p.alignment = Word.Alignment.centered;
    await ctx.sync();
    undoable = wrap(p, id);
    await ctx.sync().catch(() => {
      undoable = false;
    });
    p.getRange(Word.RangeLocation.end).select();
    await ctx.sync().catch(() => undefined);
  });
  if (undoable) pushUndo({ kind: "insert", id, label, originalOoxml: null, originalText: "", track: false });
}

/** Scrolls to and selects paragraph `n` (1-based, as numbered by getDocumentParagraphs). */
export async function selectParagraph(n: number): Promise<void> {
  requireWord();
  await Word.run(async (ctx) => {
    const paras = ctx.document.body.paragraphs;
    paras.load("items/text");
    await ctx.sync();
    const p = paras.items[n - 1];
    if (!p) throw new Error(`Paragraph ${n} no longer exists – the document changed since this answer.`);
    p.select(Word.SelectionMode.select);
    await ctx.sync();
  });
}

/**
 * Selects paragraphs `start`..`end` (1-based, as numbered by getDocumentParagraphs) – from the start of
 * the first to the end of the last one's text (its paragraph mark stays outside, so a replacement keeps
 * the paragraph's style). Returns the selected text.
 */
export async function selectParagraphRange(start: number, end: number = start): Promise<string> {
  requireWord();
  return Word.run(async (ctx) => {
    const paras = ctx.document.body.paragraphs;
    paras.load("items");
    await ctx.sync();
    const a = paras.items[start - 1];
    const b = paras.items[end - 1];
    if (!a || !b) throw new Error("Those paragraphs no longer exist – the document changed. Send the request again.");
    const range = start === end ? a.getRange(Word.RangeLocation.content) : a.getRange(Word.RangeLocation.start).expandTo(b.getRange(Word.RangeLocation.content));
    range.select(Word.SelectionMode.select);
    range.load("text");
    await ctx.sync();
    return normalizeWordText(range.text ?? "");
  });
}

/**
 * Puts the cursor in a new, empty paragraph right after paragraph `n` (so inserted text becomes its own
 * paragraph(s)). If paragraph `n` no longer has `expectText` (the document changed), the paragraph with
 * that text is used instead.
 */
export async function placeCursorAfterParagraph(n: number, expectText?: string): Promise<void> {
  requireWord();
  await Word.run(async (ctx) => {
    const paras = ctx.document.body.paragraphs;
    paras.load("items/text");
    await ctx.sync();
    const norm = (t: string) => normalizeWordText(t).replace(/\s+/g, " ").trim();
    let p = paras.items[n - 1];
    if (expectText && (!p || norm(p.text ?? "") !== norm(expectText))) {
      p = paras.items.find((x) => norm(x.text ?? "") === norm(expectText)) ?? p;
    }
    if (!p) throw new Error("That paragraph no longer exists – the document changed. Put the cursor where the text should go and use “Insert after cursor”.");
    const np = p.insertParagraph("", Word.InsertLocation.after);
    if (supportsWordApi("1.3")) np.styleBuiltIn = Word.BuiltInStyleName.normal;
    np.getRange(Word.RangeLocation.start).select();
    await ctx.sync();
  });
}

/** Font used at the end of the document (for "Append to end"). */
async function getEndFont(ctx: Word.RequestContext): Promise<{ name: string | null; size: number | null }> {
  const last = ctx.document.body.paragraphs.getLast();
  last.font.load("name,size");
  await ctx.sync();
  return { name: last.font.name || null, size: typeof last.font.size === "number" && last.font.size > 0 ? last.font.size : null };
}

/* ------------------------------------------------------------------ Track Changes */

/**
 * Runs `fn` with Word's Track Changes switched on (if requested and supported), then restores the
 * user's previous mode. Returns whether the change was actually tracked.
 */
async function withTracking<T>(ctx: Word.RequestContext, track: boolean, fn: () => Promise<T>): Promise<{ value: T; tracked: boolean }> {
  if (!track || !supportsWordApi("1.4")) return { value: await fn(), tracked: false };
  const doc = ctx.document;
  doc.load("changeTrackingMode");
  await ctx.sync();
  const previous = doc.changeTrackingMode;
  if (previous === Word.ChangeTrackingMode.off) {
    doc.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
    await ctx.sync();
  }
  try {
    return { value: await fn(), tracked: true };
  } finally {
    if (previous === Word.ChangeTrackingMode.off) {
      doc.changeTrackingMode = previous;
      await ctx.sync().catch(() => undefined);
    }
  }
}

/* ------------------------------------------------------------------ undo stack */

type UndoEntry =
  | {
      kind: "insert";
      id: string;
      label: string;
      /** Replaced content (OOXML) – null when the step only inserted new text. */
      originalOoxml: string | null;
      /** Plain text of the replaced content. */
      originalText: string;
      track: boolean;
    }
  | {
      kind: "edit";
      id: string;
      label: string;
      originalOoxml: string;
      originalText: string;
      track: boolean;
    }
  | { kind: "comments"; id: string; label: string; commentIds: string[] }
  | { kind: "table"; id: string; label: string; before: string[][] }
  | { kind: "alttext"; id: string; label: string; before: string };

const undoStack: UndoEntry[] = [];
type UndoListener = (label: string | null, depth: number) => void;
const listeners = new Set<UndoListener>();

function notify() {
  const top = undoStack[undoStack.length - 1];
  listeners.forEach((l) => l(top?.label ?? null, undoStack.length));
}

function pushUndo(entry: UndoEntry) {
  undoStack.push(entry);
  if (undoStack.length > MAX_UNDO) {
    const dropped = undoStack.splice(0, undoStack.length - MAX_UNDO);
    unwrap(untrackedTags(dropped)).catch(() => undefined);
  }
  notify();
}

export function onUndoChange(l: UndoListener): () => void {
  listeners.add(l);
  l(undoStack[undoStack.length - 1]?.label ?? null, undoStack.length);
  return () => listeners.delete(l);
}

/** Remove Zee AI's hidden wrappers but keep their text. "untracked" = all except markers around tracked edits. */
async function unwrap(tags: string[] | "all" | "untracked"): Promise<void> {
  if (!isWordAvailable() || (Array.isArray(tags) && !tags.length)) return;
  await Word.run(async (ctx) => {
    const ccs = ctx.document.contentControls;
    ccs.load("items/tag");
    await ctx.sync();
    for (const cc of ccs.items) {
      if (!cc.tag?.startsWith(TAG_PREFIX)) continue;
      if (tags === "all" || (tags === "untracked" ? !cc.tag.startsWith(TRACKED_PREFIX) : tags.includes(cc.tag))) cc.delete(true);
    }
    await ctx.sync();
  });
}

/** Tags of undo entries whose markers may be removed (tracked edits keep theirs for accept / reject). */
const untrackedTags = (entries: UndoEntry[]) => entries.filter((e) => e.kind !== "comments" && !(TAG_PREFIX + e.id).startsWith(TRACKED_PREFIX)).map((e) => TAG_PREFIX + e.id);

/**
 * Called on start-up: markers from an earlier session can't be undone any more. Markers around tracked
 * edits stay while they still contain revisions (so they can be accepted / rejected from Zee AI).
 */
export async function cleanupOrphanMarkers(): Promise<void> {
  try {
    await unwrap("untracked");
    if (!supportsWordApi("1.6")) return;
    await Word.run(async (ctx) => {
      const ccs = ctx.document.contentControls;
      ccs.load("items/tag");
      await ctx.sync();
      const tracked = ccs.items.filter((cc) => cc.tag?.startsWith(TRACKED_PREFIX));
      const changes = tracked.map((cc) => {
        const t = cc.getRange(Word.RangeLocation.whole).getTrackedChanges();
        t.load("items");
        return t;
      });
      await ctx.sync();
      tracked.forEach((cc, i) => {
        if (!changes[i].items.length) cc.delete(true); // accepted / rejected in Word's Review tab
      });
      await ctx.sync();
    });
  } catch {
    /* read-only document, protected content… – harmless */
  }
}

/** Forget the undo history (keeps all inserted text and comments). */
export async function clearUndoHistory(): Promise<void> {
  const tags = untrackedTags(undoStack);
  undoStack.length = 0;
  notify();
  await unwrap(tags).catch(() => undefined);
}

/* ------------------------------------------------------------------ Zee AI's tracked changes */

/** Revisions inside Zee AI's markers (edits made with Track Changes on). 0 when not supported. */
export async function countZeeRevisions(): Promise<number> {
  if (!supportsWordApi("1.6")) return 0;
  return Word.run(async (ctx) => {
    const ccs = ctx.document.contentControls;
    ccs.load("items/tag");
    await ctx.sync();
    const lists = ccs.items
      .filter((cc) => cc.tag?.startsWith(TRACKED_PREFIX))
      .map((cc) => {
        const t = cc.getRange(Word.RangeLocation.whole).getTrackedChanges();
        t.load("items");
        return t;
      });
    if (!lists.length) return 0;
    await ctx.sync();
    return lists.reduce((n, t) => n + t.items.length, 0);
  });
}

/**
 * Accepts or rejects only the revisions Zee AI made (inside its markers) – the user's own tracked
 * changes elsewhere are not touched. The markers are removed afterwards (text stays).
 */
export async function resolveZeeRevisions(accept: boolean): Promise<{ changes: number; edits: number }> {
  requireWord();
  if (!supportsWordApi("1.6")) throw new Error("Accepting or rejecting revisions from an add-in needs Word for Microsoft 365 (WordApi 1.6). Use Word's Review tab instead.");
  const done: string[] = [];
  let changes = 0;
  await Word.run(async (ctx) => {
    const ccs = ctx.document.contentControls;
    ccs.load("items/tag");
    await ctx.sync();
    const mine = ccs.items.filter((cc) => cc.tag?.startsWith(TRACKED_PREFIX));
    const lists = mine.map((cc) => {
      const t = cc.getRange(Word.RangeLocation.whole).getTrackedChanges();
      t.load("items");
      return t;
    });
    await ctx.sync();
    lists.forEach((t) => {
      changes += t.items.length;
      if (!t.items.length) return;
      if (accept) t.acceptAll();
      else t.rejectAll();
    });
    await ctx.sync();
    mine.forEach((cc) => {
      done.push(cc.tag);
      cc.delete(true);
    });
    await ctx.sync();
  });
  // Those steps can't be undone with ↶ any more (their markers are gone).
  const before = undoStack.length;
  for (let i = undoStack.length - 1; i >= 0; i--) if (done.includes(TAG_PREFIX + undoStack[i].id)) undoStack.splice(i, 1);
  if (undoStack.length !== before) notify();
  return { changes, edits: done.length };
}

export async function undoLast(): Promise<string> {
  requireWord();
  const entry = undoStack[undoStack.length - 1];
  if (!entry) throw new Error("Nothing to undo.");

  if (entry.kind === "comments") {
    await Word.run(async (ctx) => {
      const comments = ctx.document.body.getComments();
      comments.load("items/id");
      await ctx.sync();
      comments.items.filter((c) => entry.commentIds.includes(c.id)).forEach((c) => c.delete());
      await ctx.sync();
    });
    undoStack.pop();
    notify();
    return entry.label;
  }

  if (entry.kind === "table" || entry.kind === "alttext") {
    await Word.run(async (ctx) => {
      const found = ctx.document.contentControls.getByTag(TAG_PREFIX + entry.id);
      found.load("items");
      await ctx.sync();
      const cc = found.items[0];
      if (!cc) throw new Error(`Can't undo "${entry.label}": the table or picture was deleted or moved.`);
      if (entry.kind === "table") {
        const table = cc.tables.getFirst();
        table.load("rowCount,values");
        await ctx.sync();
        entry.before.forEach((row, r) =>
          row.forEach((v, c) => {
            if ((table.values[r]?.[c] ?? "") !== v) table.getCell(r, c).value = v;
          })
        );
      } else {
        const pic = cc.inlinePictures.getFirst();
        pic.altTextDescription = entry.before;
      }
      cc.delete(true);
      await ctx.sync();
    });
    undoStack.pop();
    notify();
    return entry.label;
  }

  const tag = TAG_PREFIX + entry.id;
  await Word.run(async (ctx) => {
    const found = ctx.document.contentControls.getByTag(tag);
    found.load("items");
    await ctx.sync();
    const cc = found.items[0];
    if (!cc) throw new Error(`Can't undo "${entry.label}": that text was deleted or moved. Use Word's own Undo (Ctrl+Z) instead.`);
    await withTracking(ctx, entry.track, async () => {
      if (entry.kind === "insert" && entry.originalOoxml === null) {
        cc.delete(false); // remove the inserted text together with its wrapper
        await ctx.sync();
        return;
      }
      // Edits: undo word by word as well, so formatting stays intact.
      if (entry.kind === "edit" && supportsWordApi("1.3")) {
        const ok = await applyMinimalEdits(ctx, cc.getRange(Word.RangeLocation.content), entry.originalText);
        if (ok) {
          cc.delete(true);
          await ctx.sync();
          return;
        }
      }
      if (entry.originalOoxml && /\n/.test(entry.originalText.trim())) {
        // Whole paragraphs: OOXML brings back the exact formatting, lists and tables.
        cc.insertOoxml(entry.originalOoxml, Word.InsertLocation.replace);
      } else {
        // Part of one paragraph: OOXML would add an extra paragraph break, so restore the text.
        cc.insertText(entry.originalText, Word.InsertLocation.replace);
      }
      cc.delete(true);
      await ctx.sync();
    });
  });
  undoStack.pop();
  notify();
  return entry.label;
}

/* ------------------------------------------------------------------ formatting-preserving edits */

/**
 * Turns the text inside `container` into `newText` by changing only the words that differ
 * (see utils/edits.ts). Paragraphs that changed too much are replaced as a whole, keeping their
 * paragraph style. Returns false (nothing changed) when the paragraph structure differs – the
 * caller then replaces the whole range.
 */
async function applyMinimalEdits(ctx: Word.RequestContext, container: Word.Range, newText: string): Promise<boolean> {
  const paras = container.paragraphs;
  paras.load("items");
  await ctx.sync();
  const segs = paras.items.map((p) => p.getRange(Word.RangeLocation.content).intersectWithOrNullObject(container));
  segs.forEach((s) => s.load("text,isNullObject"));
  await ctx.sync();

  const live: Array<{ range: Word.Range; text: string }> = [];
  for (const s of segs) {
    if (s.isNullObject) continue;
    const raw = s.text ?? "";
    if (/[\u000b\u0007]/.test(raw)) return false; // manual line breaks / table cells: use the safe path
    const text = raw.replace(/[\r\n]+$/, "");
    if (text.trim()) live.push({ range: s, text });
  }
  const plan = planEdits(
    live.map((l) => l.text),
    newText
  );
  if (!plan) return false;

  for (const edit of plan) {
    const seg = live[edit.index];
    let ok = edit.hunks !== null;
    if (edit.hunks) {
      for (const h of [...edit.hunks].reverse()) {
        if (!(await applyHunk(ctx, seg.range, h.left, h.del, h.ins, h.right))) {
          ok = false;
          break;
        }
      }
    }
    if (!ok) {
      // Replace the paragraph's text only (its paragraph mark – and so its style/list – stays).
      seg.range.insertText(edit.after.trim(), Word.InsertLocation.replace);
      await ctx.sync();
    }
  }
  return true;
}

async function applyHunk(ctx: Word.RequestContext, scope: Word.Range, left: string, del: string, ins: string, right: string): Promise<boolean> {
  const anchor = left + del + right;
  if (!anchor.trim() || anchor.length > MAX_SEARCH) return false;
  const found = scope.search(escapeWordSearch(anchor), { matchCase: true });
  found.load("items");
  await ctx.sync();
  if (found.items.length !== 1) return false;
  const area = found.items[0];

  if (del) {
    const inner = area.search(escapeWordSearch(del), { matchCase: true });
    inner.load("items");
    await ctx.sync();
    const target = inner.items[countOccurrences(left, del)];
    if (!target) return false;
    if (ins) target.insertText(ins, Word.InsertLocation.replace);
    else target.delete();
  } else if (left) {
    const inner = area.search(escapeWordSearch(left), { matchCase: true });
    inner.load("items");
    await ctx.sync();
    if (!inner.items[0]) return false;
    inner.items[0].insertText(ins, Word.InsertLocation.end);
  } else {
    const inner = area.search(escapeWordSearch(right), { matchCase: true });
    inner.load("items");
    await ctx.sync();
    if (!inner.items[0]) return false;
    inner.items[0].insertText(ins, Word.InsertLocation.start);
  }
  await ctx.sync();
  return true;
}

/* ------------------------------------------------------------------ inserting */

export interface InsertRequest {
  content: string;
  format: InsertFormat;
  location: InsertLocation;
  /** Shown in the undo tooltip, e.g. "Rewrite & polish". */
  label: string;
  /** Replace only the words that changed (keeps formatting). Plain-text replacements only. */
  preserveFormatting?: boolean;
  /** Record the change as Word revisions (Track Changes). */
  trackChanges?: boolean;
  /** Arabic right-to-left + right-aligned, English left-aligned, headings centred. Default: the setting. */
  autoDirection?: boolean;
  /** Font for Arabic characters ("" = keep). Default: the setting. */
  arabicFont?: string;
}

export interface InsertResult {
  location: InsertLocation;
  undoable: boolean;
  /** Formatting-preserving word-level edit was used. */
  minimal: boolean;
  /** Change was recorded with Track Changes. */
  tracked: boolean;
}

const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
/** Marker id: tracked edits get a "t" prefix so their markers survive until accepted / rejected. */
const markId = (track: boolean) => (track && supportsWordApi("1.4") ? "t" : "") + uid();

function wrap(target: Word.Range | Word.Paragraph, id: string): boolean {
  try {
    const cc = target.insertContentControl();
    cc.tag = TAG_PREFIX + id;
    cc.title = "Zee AI";
    cc.appearance = Word.ContentControlAppearance.hidden;
    return true;
  } catch {
    return false;
  }
}

/* ------------------------------------------------------------------ text direction */

const HEADING_BUILTIN = /^(heading\d|title|subtitle)$/i;
/** Style names as Word shows them (English and Arabic UI). */
const HEADING_NAME = /^(heading|title|subtitle|عنوان)/i;
/** Paragraphs checked per insert/edit (alignment). */
const MAX_DIRECTION_PARAS = 400;
/** Paragraphs whose right-to-left setting is fixed through OOXML (one OOXML read each). */
const MAX_BIDI_PARAS = 80;

export interface DirectionOptions {
  /** The paragraphs were created by this insert (centre headings; any alignment except Justified may change). */
  whole: boolean;
  /** Font for Arabic characters ("" = keep the document's). */
  arabicFont?: string;
  /** Fix direction/alignment (false = only apply the Arabic font). Default true. */
  direction?: boolean;
}

function supportsDesktopApi(version: string): boolean {
  try {
    return isWordAvailable() && Office.context.requirements.isSetSupported("WordApiDesktop", version);
  } catch {
    return false;
  }
}

/** True when the document records revisions – an OOXML rewrite would then show up as tracked changes. */
async function isTracking(ctx: Word.RequestContext): Promise<boolean> {
  if (!supportsWordApi("1.4")) return false;
  const doc = ctx.document;
  doc.load("changeTrackingMode");
  await ctx.sync();
  return doc.changeTrackingMode !== Word.ChangeTrackingMode.off;
}

/**
 * Gives text Zee AI wrote or edited the right direction:
 * - Arabic paragraphs become right-to-left (paragraph bidi, set through the paragraph's OOXML) and
 *   sit on the right; English paragraphs become left-to-right on the left;
 * - headings are centred (new content only);
 * - Arabic characters get the chosen Arabic font.
 * Paragraphs whose OOXML can't be rewritten safely (comments, footnotes, lists, tables, tracked
 * changes…) only get the alignment. Never throws – direction must not block an insert.
 */
export async function fixDirection(ctx: Word.RequestContext, target: Word.Range | Word.Paragraph, opts: DirectionOptions): Promise<void> {
  try {
    const v13 = supportsWordApi("1.3");
    const range = target instanceof Word.Paragraph ? (v13 ? target.getRange() : null) : target;

    // Arabic font: only complex-script characters are affected, so the whole range can be set.
    if (opts.arabicFont && supportsDesktopApi("1.3")) {
      const r = range ?? (target as Word.Paragraph);
      (r as Word.Range).font.nameBidirectional = opts.arabicFont;
      await ctx.sync().catch(() => undefined);
    }
    if (opts.direction === false) return;

    let paras: Word.Paragraph[];
    if (range) {
      const col = range.paragraphs;
      col.load(
        v13
          ? "items/text,items/alignment,items/style,items/styleBuiltIn,items/tableNestingLevel,items/isListItem"
          : "items/text,items/alignment,items/style"
      );
      await ctx.sync();
      paras = col.items.slice(0, MAX_DIRECTION_PARAS);
    } else {
      const p = target as Word.Paragraph;
      p.load("text,alignment,style");
      await ctx.sync();
      paras = [p];
    }

    const plan = paras.map((p) => {
      const builtIn = v13 ? String(p.styleBuiltIn ?? "") : "";
      const heading = HEADING_BUILTIN.test(builtIn) || HEADING_NAME.test(p.style ?? "");
      const inTableOrList = v13 && ((p.tableNestingLevel ?? 0) > 0 || !!p.isListItem);
      return { p, heading, dir: detectDirection(p.text ?? ""), inTableOrList, bidi: null as boolean | null, done: false };
    });

    /* ---- 1. paragraph direction (bidi) through OOXML */
    const tracking = await isTracking(ctx).catch(() => true);
    const candidates = tracking ? [] : plan.filter((x) => x.dir && !x.inTableOrList).slice(0, MAX_BIDI_PARAS);
    if (candidates.length) {
      const xml = candidates.map((x) => x.p.getOoxml());
      await ctx.sync();
      const changes: Array<{ x: (typeof plan)[number]; pkg: string; next: Word.Paragraph }> = [];
      candidates.forEach((x, i) => {
        const info = inspectParagraphXml(xml[i].value);
        if (!info) return;
        x.bidi = info.bidi;
        if (info.risky) return;
        const pkg = setParagraphDirectionXml(xml[i].value, x.dir!, opts.whole && x.heading);
        if (pkg === null) {
          x.done = true; // already right
          return;
        }
        const next = x.p.getNextOrNullObject();
        next.load("text");
        changes.push({ x, pkg, next });
      });
      if (changes.length) {
        await ctx.sync();
        const probes = changes.map((c) => {
          const r = c.x.p.insertOoxml(c.pkg, Word.InsertLocation.replace);
          const after = r.paragraphs.getLast().getNextOrNullObject();
          after.load("text");
          return { c, after };
        });
        await ctx.sync();
        // Replacing a paragraph through OOXML can leave an extra empty paragraph behind it.
        for (const { c, after } of probes) {
          c.x.done = true;
          const wasEmpty = !c.next.isNullObject && !(c.next.text ?? "").trim();
          if (!after.isNullObject && !(after.text ?? "").trim() && !wasEmpty) after.delete();
        }
        await ctx.sync();
      }
    }

    /* ---- 2. alignment for everything not handled above */
    for (const x of plan) {
      if (x.done) continue;
      const current = String(x.p.alignment);
      if (x.heading) {
        if (opts.whole && current !== "Centered") x.p.alignment = "Centered";
        continue;
      }
      if (!x.dir) continue;
      // A bidi paragraph's left/right is mirrored by Word – leave the side of RTL paragraphs alone.
      if (x.bidi) continue;
      const want = alignmentFor(x.dir);
      if (current === want || current === "Justified") continue;
      if (!opts.whole && current !== "Left" && current !== "Right") continue;
      x.p.alignment = want;
    }
    await ctx.sync();
  } catch {
    /* direction is best effort */
  }
}

export async function insertContent(req: InsertRequest): Promise<InsertResult> {
  requireWord();
  const id = markId(!!req.trackChanges);
  const content = req.content.replace(/\r\n?/g, "\n");
  if (!content.trim()) throw new Error("There is no text to insert.");

  // Filled inside Word.run (an object, so TypeScript doesn't narrow it to its initial value).
  const state = { ooxml: null as string | null, text: "", undoable: true, minimal: false, tracked: false };
  const track = !!req.trackChanges;
  const saved = loadSettings();
  const auto = req.autoDirection ?? saved.autoDirection;
  const arabicFont = (req.arabicFont ?? saved.arabicFont).trim();

  await Word.run(async (ctx) => {
    const sel = ctx.document.getSelection();
    sel.load("text,isEmpty");
    sel.font.load("name,size");
    await ctx.sync();

    const replacing = req.location === "replace" && !sel.isEmpty;
    if (replacing) {
      const ooxml = sel.getOoxml();
      await ctx.sync();
      state.text = normalizeWordText(sel.text ?? "");
      state.ooxml = ooxml.value.length <= MAX_OOXML_CHARS ? ooxml.value : "";
    }

    /* ---- formatting-preserving edit: wrap the selection, then change only the differing words */
    if (replacing && req.format === "plain" && req.preserveFormatting !== false && supportsWordApi("1.3")) {
      let cc: Word.ContentControl | null = null;
      try {
        cc = sel.insertContentControl();
        cc.tag = TAG_PREFIX + id;
        cc.title = "Zee AI";
        cc.appearance = Word.ContentControlAppearance.hidden;
        await ctx.sync();
      } catch {
        cc = null; // can't wrap (e.g. selection crosses a table edge) – use the plain replace below
      }
      if (cc) {
        const wrapped = cc;
        const r = await withTracking(ctx, track, async () => {
          const ok = await applyMinimalEdits(ctx, wrapped.getRange(Word.RangeLocation.content), content);
          if (!ok) return false;
          // Verify the result; if Word's text doesn't match, fall back to a full replacement.
          wrapped.load("text");
          await ctx.sync();
          if (!sameText(normalizeWordText(wrapped.text ?? ""), content)) {
            wrapped.insertText(content, Word.InsertLocation.replace);
            await ctx.sync();
          }
          return true;
        });
        state.tracked = r.tracked;
        if (r.value) {
          state.minimal = true;
          // Edits (e.g. Translate English → Arabic) flip the paragraph direction too.
          if (auto || arabicFont) await fixDirection(ctx, wrapped.getRange(Word.RangeLocation.content), { whole: false, arabicFont, direction: auto });
          wrapped.getRange(Word.RangeLocation.end).select();
          await ctx.sync().catch(() => undefined);
          return;
        }
        // Structure changed (paragraph count): replace the wrapped content as a whole.
        await withTracking(ctx, track, async () => {
          wrapped.insertText(content, Word.InsertLocation.replace);
          await ctx.sync();
          if (auto || arabicFont) await fixDirection(ctx, wrapped.getRange(Word.RangeLocation.content), { whole: false, arabicFont, direction: auto });
        });
        wrapped.getRange(Word.RangeLocation.end).select();
        await ctx.sync().catch(() => undefined);
        return;
      }
    }

    let font = { name: sel.font.name || null, size: typeof sel.font.size === "number" && sel.font.size > 0 ? sel.font.size : null };
    if (req.location === "end") font = await getEndFont(ctx);

    // Plain text that forms its own paragraphs (multi-line, at the end, or typed into an empty
    // paragraph) is inserted as HTML so each paragraph gets its own direction. A short text dropped
    // into the middle of an existing paragraph stays inline.
    let ownParagraphs = req.location === "end" || /\n/.test(content.trim());
    if (auto && !ownParagraphs && sel.isEmpty && supportsWordApi("1.3")) {
      const para = sel.paragraphs.getFirst();
      para.load("text");
      await ctx.sync();
      ownParagraphs = !(para.text ?? "").trim();
    }
    const plainAsHtml = auto && req.format !== "markdown" && ownParagraphs && hasRtl(content);

    // A tracked replacement is done inside a marker that wraps the old text first, so the deleted and
    // the inserted revisions both end up inside it (needed for "Accept / Reject Zee AI's changes").
    let pre: Word.ContentControl | null = null;
    if (replacing && track && supportsWordApi("1.4")) {
      try {
        pre = sel.insertContentControl();
        pre.tag = TAG_PREFIX + id;
        pre.title = "Zee AI";
        pre.appearance = Word.ContentControlAppearance.hidden;
        await ctx.sync();
      } catch {
        pre = null; // e.g. the selection crosses a table edge – replace without the outer marker
      }
    }

    const r = await withTracking(ctx, track, async () => {
      let inserted: Word.Range | Word.Paragraph;
      if (req.format === "markdown" || plainAsHtml) {
        const htmlOpts = { fontFamily: font.name ?? undefined, fontSizePt: font.size ?? undefined, autoDirection: auto, arabicFont: arabicFont || undefined };
        const html = req.format === "markdown" ? markdownToHtml(content, htmlOpts) : plainToHtml(content, htmlOpts);
        if (pre) inserted = pre.insertHtml(html, Word.InsertLocation.replace);
        else if (req.location === "end") inserted = ctx.document.body.insertHtml(html, Word.InsertLocation.end);
        else if (req.location === "after") inserted = sel.insertHtml(html, sel.isEmpty ? Word.InsertLocation.replace : Word.InsertLocation.after);
        else inserted = sel.insertHtml(html, Word.InsertLocation.replace);
      } else if (pre) {
        inserted = pre.insertText(content, Word.InsertLocation.replace);
      } else if (req.location === "end") {
        // Start on a new paragraph at the end of the document.
        inserted = ctx.document.body.insertParagraph(content, Word.InsertLocation.end);
      } else if (req.location === "after" && !sel.isEmpty) {
        inserted = sel.insertText(/\n/.test(content) ? `\n${content}` : ` ${content}`, Word.InsertLocation.after);
      } else {
        inserted = sel.insertText(content, Word.InsertLocation.replace);
      }
      await ctx.sync();
      return inserted;
    });
    state.tracked = r.tracked;
    const inserted = r.value;

    // Wrap for undo. Some ranges (e.g. inside equations or locked content controls) can't be wrapped:
    // the text is still inserted, only Zee AI's undo is unavailable (Word's Ctrl+Z still works).
    state.undoable = pre ? true : wrap(inserted, id);
    try {
      await ctx.sync();
    } catch {
      state.undoable = false;
    }
    let scope: Word.Range | Word.Paragraph = inserted;
    if (auto || arabicFont) {
      // After wrapping, so the undo marker keeps covering paragraphs that get rewritten for direction.
      // fixDirection never throws, so direction problems can't break the insert.
      if (state.undoable) scope = ctx.document.contentControls.getByTag(TAG_PREFIX + id).getFirst().getRange(Word.RangeLocation.content);
      await fixDirection(ctx, scope, { whole: ownParagraphs || req.format === "markdown", arabicFont, direction: auto });
    }
    scope.select(Word.SelectionMode.end); // cursor after the new text, ready for the next insert
    await ctx.sync().catch(() => undefined);
  });

  if (state.undoable) {
    const replaced = req.location === "replace" && state.text !== "";
    if (state.minimal) {
      pushUndo({ kind: "edit", id, label: req.label, originalOoxml: state.ooxml ?? "", originalText: state.text, track });
    } else {
      pushUndo({ kind: "insert", id, label: req.label, originalOoxml: replaced ? state.ooxml ?? "" : null, originalText: state.text, track });
    }
  }
  return { location: req.location, undoable: state.undoable, minimal: state.minimal, tracked: state.tracked };
}

/* ------------------------------------------------------------------ comments */

export interface CommentRequest {
  /** Text to find in the document (the comment is anchored to it). */
  find: string;
  /** Comment text. */
  text: string;
  /** 1-based paragraph to search first ([P#]); without a match the comment goes on that paragraph. */
  paragraph?: number;
}

/**
 * Adds Word comments anchored to each `find` text – inside the selection when possible, otherwise
 * anywhere in the document. Repeated identical `find` texts are anchored to successive occurrences.
 */
export async function addComments(items: CommentRequest[], label: string): Promise<{ added: number; missing: number }> {
  requireWord();
  if (!supportsWordApi("1.4")) throw new Error("Comments need a newer Word (WordApi 1.4: Microsoft 365 / Word 2021 or later).");
  const ids: string[] = [];
  let missing = 0;
  await Word.run(async (ctx) => {
    const sel = ctx.document.getSelection();
    sel.load("isEmpty");
    await ctx.sync();
    const used = new Map<string, number>();
    const created: Word.Comment[] = [];
    const needParas = items.some((i) => i.paragraph && i.paragraph > 0);
    const paras = needParas ? ctx.document.body.paragraphs : null;
    if (paras) {
      paras.load("items");
      await ctx.sync();
    }
    for (const item of items) {
      const find = item.find.replace(/\s+/g, " ").trim().slice(0, MAX_SEARCH);
      const para = paras && item.paragraph ? paras.items[item.paragraph - 1] : undefined;
      if (para) {
        let anchor: Word.Range | null = null;
        if (find) {
          const inPara = para.search(escapeWordSearch(find), { matchCase: false });
          inPara.load("items");
          await ctx.sync();
          anchor = inPara.items[0] ?? null;
        }
        const c = (anchor ?? para.getRange(Word.RangeLocation.whole)).insertComment(item.text);
        c.load("id");
        created.push(c);
        continue;
      }
      if (!find) {
        missing++;
        continue;
      }
      let hits = sel.isEmpty ? null : sel.search(escapeWordSearch(find), { matchCase: true });
      if (hits) {
        hits.load("items");
        await ctx.sync();
      }
      if (!hits || !hits.items.length) {
        hits = ctx.document.body.search(escapeWordSearch(find), { matchCase: true });
        hits.load("items");
        await ctx.sync();
      }
      const k = used.get(find) ?? 0;
      const target = hits.items[k] ?? hits.items[0];
      if (!target) {
        missing++;
        continue;
      }
      used.set(find, k + 1);
      const c = target.insertComment(item.text);
      c.load("id");
      created.push(c);
    }
    await ctx.sync();
    created.forEach((c) => ids.push(c.id));
  });
  if (ids.length) pushUndo({ kind: "comments", id: uid(), label, commentIds: ids });
  return { added: ids.length, missing };
}

/* ------------------------------------------------------------------ misc */

/** Plain-text version of AI output for the clipboard. */
export function toClipboardText(content: string, format: InsertFormat): string {
  return format === "markdown" ? markdownToPlain(content) : content;
}

/** Fires (debounced) when the user changes the selection in Word. Returns an unsubscribe function. */
export function onSelectionChange(cb: () => void): () => void {
  if (!isWordAvailable()) return () => undefined;
  let timer = 0;
  const handler = () => {
    window.clearTimeout(timer);
    timer = window.setTimeout(cb, 200);
  };
  try {
    Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, handler);
  } catch {
    return () => undefined;
  }
  return () => {
    window.clearTimeout(timer);
    try {
      Office.context.document.removeHandlerAsync(Office.EventType.DocumentSelectionChanged, { handler });
    } catch {
      /* ignore */
    }
  };
}

/* ------------------------------------------------------------------ tables */

export interface TableInfo {
  values: string[][];
  rows: number;
  cols: number;
}

/** The table the cursor is in (or the first table in the selection), or null. */
export async function getSelectedTable(): Promise<TableInfo | null> {
  requireWord();
  if (!supportsWordApi("1.3")) throw new Error("Table actions need Word 2016 or later (WordApi 1.3).");
  return Word.run(async (ctx) => {
    const sel = ctx.document.getSelection();
    const parent = sel.parentTableOrNullObject;
    parent.load("isNullObject,values,rowCount");
    const inSel = sel.tables.getFirstOrNullObject();
    inSel.load("isNullObject,values,rowCount");
    await ctx.sync();
    const t = !parent.isNullObject ? parent : !inSel.isNullObject ? inSel : null;
    if (!t) return null;
    const values = (t.values ?? []).map((r) => r.map((v) => normalizeWordText(String(v ?? "")).replace(/\n+/g, " ").trim()));
    return { values, rows: values.length, cols: Math.max(0, ...values.map((r) => r.length)) };
  });
}

/** Writes only the changed cells of the selected table (cell formatting stays). Undoable. */
export async function writeTableCells(after: string[][], label: string, trackChanges: boolean): Promise<{ changed: number; tracked: boolean }> {
  requireWord();
  const id = markId(trackChanges);
  let changed = 0;
  let tracked = false;
  let before: string[][] = [];
  let undoable = false;
  await Word.run(async (ctx) => {
    const sel = ctx.document.getSelection();
    const parent = sel.parentTableOrNullObject;
    parent.load("isNullObject,values");
    const inSel = sel.tables.getFirstOrNullObject();
    inSel.load("isNullObject,values");
    await ctx.sync();
    const table = !parent.isNullObject ? parent : !inSel.isNullObject ? inSel : null;
    if (!table) throw new Error("Put the cursor in the table again, then click Apply.");
    before = table.values.map((r) => r.map((v) => String(v ?? "")));
    try {
      const cc = table.getRange(Word.RangeLocation.whole).insertContentControl();
      cc.tag = TAG_PREFIX + id;
      cc.title = "Zee AI";
      cc.appearance = Word.ContentControlAppearance.hidden;
      await ctx.sync();
      undoable = true;
    } catch {
      undoable = false;
    }
    const r = await withTracking(ctx, trackChanges, async () => {
      after.forEach((row, i) =>
        row.forEach((v, j) => {
          const old = before[i]?.[j];
          if (old === undefined) return;
          if (normalizeWordText(old).replace(/\s+/g, " ").trim() !== v.replace(/\s+/g, " ").trim()) {
            table.getCell(i, j).value = v;
            changed++;
          }
        })
      );
      await ctx.sync();
    });
    tracked = r.tracked;
  });
  if (undoable && changed) pushUndo({ kind: "table", id, label, before });
  else if (undoable) unwrap([TAG_PREFIX + id]).catch(() => undefined);
  return { changed, tracked };
}

/* ------------------------------------------------------------------ pictures */

export interface PictureInfo {
  base64: string;
  altText: string;
  /** Text of the paragraphs around the picture (context for captions). */
  context: string;
}

/** The first picture in the selection (or the picture the cursor is on), or null. */
export async function getSelectedPicture(): Promise<PictureInfo | null> {
  requireWord();
  return Word.run(async (ctx) => {
    const pics = ctx.document.getSelection().inlinePictures;
    pics.load("items/altTextDescription");
    await ctx.sync();
    const pic = pics.items[0];
    if (!pic) return null;
    const b64 = pic.getBase64ImageSrc();
    let context = "";
    if (supportsWordApi("1.3")) {
      const para = pic.paragraph;
      const prev = para.getPreviousOrNullObject();
      const next = para.getNextOrNullObject();
      prev.load("text,isNullObject");
      next.load("text,isNullObject");
      await ctx.sync();
      context = [prev.isNullObject ? "" : prev.text, next.isNullObject ? "" : next.text].map((t) => normalizeWordText(t ?? "").trim()).filter(Boolean).join("\n").slice(0, 1500);
    } else {
      await ctx.sync();
    }
    return { base64: b64.value, altText: pic.altTextDescription ?? "", context };
  });
}

/** Sets the selected picture's alt text (undoable). */
export async function setPictureAltText(text: string, label: string): Promise<void> {
  requireWord();
  const id = uid();
  let before = "";
  let undoable = false;
  await Word.run(async (ctx) => {
    const pics = ctx.document.getSelection().inlinePictures;
    pics.load("items/altTextDescription");
    await ctx.sync();
    const pic = pics.items[0];
    if (!pic) throw new Error("Select the picture again, then click Apply.");
    before = pic.altTextDescription ?? "";
    try {
      const cc = pic.insertContentControl();
      cc.tag = TAG_PREFIX + id;
      cc.title = "Zee AI";
      cc.appearance = Word.ContentControlAppearance.hidden;
      await ctx.sync();
      undoable = true;
    } catch {
      undoable = false;
    }
    pic.altTextDescription = text;
    await ctx.sync();
  });
  if (undoable) pushUndo({ kind: "alttext", id, label, before });
}

/** Inserts a caption paragraph (Word's Caption style) directly below the selected picture (undoable). */
export async function insertPictureCaption(text: string, label: string, trackChanges: boolean): Promise<void> {
  requireWord();
  const id = markId(trackChanges);
  let undoable = false;
  await Word.run(async (ctx) => {
    const pics = ctx.document.getSelection().inlinePictures;
    pics.load("items");
    await ctx.sync();
    const pic = pics.items[0];
    if (!pic) throw new Error("Select the picture again, then click Insert caption.");
    await withTracking(ctx, trackChanges, async () => {
      const p = pic.paragraph.insertParagraph(text, Word.InsertLocation.after);
      if (supportsWordApi("1.3")) p.styleBuiltIn = Word.BuiltInStyleName.caption;
      await ctx.sync();
      undoable = wrap(p, id);
      await ctx.sync().catch(() => {
        undoable = false;
      });
    });
  });
  if (undoable) pushUndo({ kind: "insert", id, label, originalOoxml: null, originalText: "", track: trackChanges });
}
