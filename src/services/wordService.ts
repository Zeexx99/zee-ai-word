/**
 * Office.js helpers for Word: reading the selection / document, inserting AI text, formatting-
 * preserving edits, Track Changes, Word comments, paragraph navigation and the multi-step undo stack.
 *
 * Undo design: every insertion or edit is wrapped in a hidden content control tagged "zeeai-<id>".
 * Undo finds that control and either deletes it (pure insertions) or puts the original text back
 * (edits / replacements). Comments are undone by id. The stack lives in memory only; when the pane
 * reloads, leftover hidden controls are unwrapped (their text stays) so documents never accumulate
 * Zee AI markers.
 */
import { markdownToHtml, markdownToPlain, plainToHtml } from "../utils/mdToHtml";
import { alignmentFor, detectDirection, hasRtl } from "../utils/direction";
import { loadSettings } from "./settingsService";
import { normalizeWordText } from "../utils/text";
import { countOccurrences, escapeWordSearch, planEdits, MAX_SEARCH } from "../utils/edits";

export type InsertLocation = "replace" | "after" | "end";
export type InsertFormat = "markdown" | "plain";

export const LOCATION_LABELS: Record<InsertLocation, string> = {
  replace: "Replace selection",
  after: "Insert after cursor",
  end: "Append to end of document",
};

const TAG_PREFIX = "zeeai-";
const MAX_UNDO = 25;
const MAX_OOXML_CHARS = 5_000_000;

export function isWordAvailable(): boolean {
  try {
    return typeof Office !== "undefined" && Office.context?.host === Office.HostType.Word && typeof Word !== "undefined";
  } catch {
    return false;
  }
}

/** WordApi requirement-set check (1.3: ranges & intersections, 1.4: comments & Track Changes). */
export function supportsWordApi(version: "1.3" | "1.4"): boolean {
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
    unwrap(dropped.filter((e) => e.kind !== "comments").map((e) => TAG_PREFIX + e.id)).catch(() => undefined);
  }
  notify();
}

export function onUndoChange(l: UndoListener): () => void {
  listeners.add(l);
  l(undoStack[undoStack.length - 1]?.label ?? null, undoStack.length);
  return () => listeners.delete(l);
}

/** Remove Zee AI's hidden wrappers but keep their text. */
async function unwrap(tags: string[] | "all"): Promise<void> {
  if (!isWordAvailable() || (tags !== "all" && !tags.length)) return;
  await Word.run(async (ctx) => {
    const ccs = ctx.document.contentControls;
    ccs.load("items/tag");
    await ctx.sync();
    for (const cc of ccs.items) {
      if (!cc.tag?.startsWith(TAG_PREFIX)) continue;
      if (tags === "all" || tags.includes(cc.tag)) cc.delete(true);
    }
    await ctx.sync();
  });
}

/** Called on start-up: markers from an earlier session can't be undone any more. */
export async function cleanupOrphanMarkers(): Promise<void> {
  try {
    await unwrap("all");
  } catch {
    /* read-only document, protected content… – harmless */
  }
}

/** Forget the undo history (keeps all inserted text and comments). */
export async function clearUndoHistory(): Promise<void> {
  const tags = undoStack.filter((e) => e.kind !== "comments").map((e) => TAG_PREFIX + e.id);
  undoStack.length = 0;
  notify();
  await unwrap(tags).catch(() => undefined);
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
const MAX_DIRECTION_PARAS = 400;

/**
 * Aligns the paragraphs of freshly inserted content: headings centred, Arabic paragraphs right,
 * English paragraphs left. `whole` = the paragraphs were created by this insert (any alignment may
 * be changed except Justified); otherwise the text went into an existing paragraph, so only a
 * left/right alignment is flipped – centred or justified paragraphs the user set up are kept.
 */
async function applyAutoDirection(ctx: Word.RequestContext, target: Word.Range | Word.Paragraph, whole: boolean): Promise<void> {
  const v13 = supportsWordApi("1.3");
  const range = target instanceof Word.Paragraph ? (v13 ? target.getRange() : null) : target;
  let paras: Word.Paragraph[];
  if (range) {
    const col = range.paragraphs;
    col.load(v13 ? "items/text,items/alignment,items/style,items/styleBuiltIn,items/tableNestingLevel" : "items/text,items/alignment,items/style");
    await ctx.sync();
    paras = col.items;
  } else {
    const p = target as Word.Paragraph;
    p.load("text,alignment,style");
    await ctx.sync();
    paras = [p];
  }
  for (const p of paras.slice(0, MAX_DIRECTION_PARAS)) {
    const builtIn = v13 ? String((p as Word.Paragraph).styleBuiltIn ?? "") : "";
    const heading = HEADING_BUILTIN.test(builtIn) || HEADING_NAME.test(p.style ?? "");
    const current = String(p.alignment);
    if (heading) {
      if (whole && current !== "Centered") p.alignment = "Centered";
      continue;
    }
    const dir = detectDirection(p.text ?? "");
    if (!dir) continue;
    const want = alignmentFor(dir);
    if (current === want) continue;
    if (current === "Justified") continue; // justified reads correctly in both directions
    if (!whole && current !== "Left" && current !== "Right") continue;
    p.alignment = want;
  }
  await ctx.sync();
}

export async function insertContent(req: InsertRequest): Promise<InsertResult> {
  requireWord();
  const id = uid();
  const content = req.content.replace(/\r\n?/g, "\n");
  if (!content.trim()) throw new Error("There is no text to insert.");

  // Filled inside Word.run (an object, so TypeScript doesn't narrow it to its initial value).
  const state = { ooxml: null as string | null, text: "", undoable: true, minimal: false, tracked: false };
  const track = !!req.trackChanges;

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
          wrapped.getRange(Word.RangeLocation.end).select();
          await ctx.sync().catch(() => undefined);
          return;
        }
        // Structure changed (paragraph count): replace the wrapped content as a whole.
        await withTracking(ctx, track, async () => {
          wrapped.insertText(content, Word.InsertLocation.replace);
          await ctx.sync();
        });
        wrapped.getRange(Word.RangeLocation.end).select();
        await ctx.sync().catch(() => undefined);
        return;
      }
    }

    let font = { name: sel.font.name || null, size: typeof sel.font.size === "number" && sel.font.size > 0 ? sel.font.size : null };
    if (req.location === "end") font = await getEndFont(ctx);

    const auto = req.autoDirection ?? loadSettings().autoDirection;
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

    const r = await withTracking(ctx, track, async () => {
      let inserted: Word.Range | Word.Paragraph;
      if (req.format === "markdown" || plainAsHtml) {
        const htmlOpts = { fontFamily: font.name ?? undefined, fontSizePt: font.size ?? undefined, autoDirection: auto };
        const html = req.format === "markdown" ? markdownToHtml(content, htmlOpts) : plainToHtml(content, htmlOpts);
        if (req.location === "end") inserted = ctx.document.body.insertHtml(html, Word.InsertLocation.end);
        else if (req.location === "after") inserted = sel.insertHtml(html, sel.isEmpty ? Word.InsertLocation.replace : Word.InsertLocation.after);
        else inserted = sel.insertHtml(html, Word.InsertLocation.replace);
      } else if (req.location === "end") {
        // Start on a new paragraph at the end of the document.
        inserted = ctx.document.body.insertParagraph(content, Word.InsertLocation.end);
      } else if (req.location === "after" && !sel.isEmpty) {
        inserted = sel.insertText(/\n/.test(content) ? `\n${content}` : ` ${content}`, Word.InsertLocation.after);
      } else {
        inserted = sel.insertText(content, Word.InsertLocation.replace);
      }
      await ctx.sync();
      if (auto) {
        // Never let alignment problems block the insert itself.
        await applyAutoDirection(ctx, inserted, ownParagraphs || req.format === "markdown").catch(() => undefined);
      }
      return inserted;
    });
    state.tracked = r.tracked;
    const inserted = r.value;

    // Wrap for undo. Some ranges (e.g. inside equations or locked content controls) can't be wrapped:
    // the text is still inserted, only Zee AI's undo is unavailable (Word's Ctrl+Z still works).
    state.undoable = wrap(inserted, id);
    try {
      await ctx.sync();
    } catch {
      state.undoable = false;
    }
    inserted.select(Word.SelectionMode.end); // cursor after the new text, ready for the next insert
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
  const id = uid();
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
  const id = uid();
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
