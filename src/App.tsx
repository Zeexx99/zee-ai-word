import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import {
  Badge,
  Button,
  FluentProvider,
  Menu,
  MenuItem,
  MenuList,
  MenuPopover,
  MenuTrigger,
  MessageBar,
  MessageBarActions,
  MessageBarBody,
  Toast,
  ToggleButton,
  ToastBody,
  ToastTitle,
  Toaster,
  Tooltip,
  useId,
  useToastController,
  type ToastIntent,
} from "@fluentui/react-components";
import {
  Add20Regular,
  ArrowDownload20Regular,
  ArrowUndo20Regular,
  CheckmarkCircle20Regular,
  Copy20Regular,
  DismissCircle20Regular,
  DocumentAdd20Regular,
  DocumentCheckmark20Regular,
  DocumentEdit20Filled,
  DocumentEdit20Regular,
  Settings20Regular,
  WeatherMoon20Regular,
  WeatherSunny20Regular,
} from "@fluentui/react-icons";

import { darkTheme, lightTheme, useAutoDark, useBodyThemeClass } from "./theme";
import { useChat, type ChatMessage } from "./hooks/useChat";
import { ChatFeed } from "./components/ChatFeed";
import { ChatInput } from "./components/ChatInput";
import { QuickActions, useKey } from "./components/QuickActions";
// Imported eagerly (never lazy-loaded): a separate chunk breaks when Office holds a cached page from an older deploy.
import { SettingsDrawer } from "./components/SettingsDrawer";
import {
  ConfirmDialog,
  DraftDialog,
  ReferenceDialog,
  TemplateDialog,
  type ConfirmState,
  type DraftChoice,
  type ReferenceChoice,
  type TemplateChoice,
} from "./components/Dialogs";
import { InsertActionsContext, copyText, type InsertActions } from "./components/insertActions";

import {
  allCustomActions,
  bulkSettings,
  isConfigured,
  loadSettings,
  modelLabel,
  protectedTermList,
  sanitizeSettings,
  saveSettings,
  type CustomAction,
  type DocContext,
  type ZeeSettings,
} from "./services/settingsService";
import { onCommand } from "./services/commands";
import { fetchTeamConfig } from "./services/teamService";
import type { ReportTemplate } from "./services/templates";
import { diffGrid, fitGrid, keepsNumbers, tableToMarkdown } from "./utils/tables";
import { detectImageMime } from "./utils/files";
import {
  LOCATION_LABELS,
  addComments,
  getSelectedPicture,
  getSelectedTable,
  insertPictureCaption,
  setPictureAltText,
  writeTableCells,
  cleanupOrphanMarkers,
  clearUndoHistory,
  getDocumentParagraphs,
  getSelectionInfo,
  insertContent,
  isWordAvailable,
  onSelectionChange,
  onUndoChange,
  countZeeRevisions,
  placeCursorAfterParagraph,
  resolveZeeRevisions,
  selectParagraph,
  selectParagraphRange,
  supportsWordApi,
  undoLast,
  type InsertFormat,
  type InsertLocation,
} from "./services/wordService";
import {
  CONSISTENCY_SYSTEM_PROMPT,
  DRAFT_SYSTEM_PROMPT,
  FINDINGS_SCHEMA,
  PICTURE_SYSTEM_PROMPT,
  REFERENCE_SYSTEM_PROMPT,
  TABLE_CLEAN_SCHEMA,
  TABLE_SYSTEM_PROMPT,
  consistencyPrompt,
  explainTablePrompt,
  notesToTablePrompt,
  picturePrompt,
  referenceCheckPrompt,
  tableCleanPrompt,
  tableCommentaryPrompt,
  templateFillPrompt,
  type DocFinding,
  GRAMMAR_SCHEMA,
  GRAMMAR_SYSTEM_PROMPT,
  chunkSummaryPrompt,
  combineSummaryPrompt,
  customActionPrompt,
  grammarChunkPrompt,
  type GrammarIssue,
  type GrammarResult,
  EDIT_SYSTEM_PROMPT,
  SUMMARY_SYSTEM_PROMPT,
  TONES,
  draftPrompt,
  lengthPrompt,
  rewritePrompt,
  summarizePrompt,
  tonePrompt,
  translatePrompt,
  bilingualPrompt,
  BILINGUAL_SYSTEM_PROMPT,
  LOCATE_SCHEMA,
  LOCATE_SYSTEM_PROMPT,
  generalEditPrompt,
  generalInsertPrompt,
  locatePrompt,
  withGlossary,
  type SummaryStyle,
  type ToneId,
} from "./services/prompts";
import { estimateTokens, generateFromParts, generateJson, generateText, GeminiError, type GeminiPart, type Usage } from "./services/geminiService";
import { getDocTitle } from "./services/chatStore";
import { countWords, preserveEdges, shorten, truncateMiddle } from "./utils/text";
import { looksLikeMarkdown, markdownToPlain } from "./utils/mdToHtml";
import { chunkParagraphs, numberParagraphs } from "./utils/chunks";
import { normalizeLocate, paraLabel, paraListLabel, relevantExcerpt, type LocatePlan, type LocateResult } from "./utils/locate";
import { UpdateBanners, useUpdates } from "./qol/UpdateBanners";
import { recordUsage, usageForDay } from "./qol/usageLog";
import { buildDiagnostics } from "./qol/diagnostics";
import { bumpUse, loadUseCounts } from "./qol/recent";

export interface UsageTotals {
  requests: number;
  promptTokens: number;
  outputTokens: number;
}

/** Selections larger than this are refused for edit actions (keeps answers complete and costs predictable). */
const MAX_EDIT_CHARS = 60_000;
/** Size of one part when a whole document is processed in chunks. */
const CHUNK_CHARS = 12_000;

const sameText = (a: string, b: string) => a.replace(/\s+/g, " ").trim() === b.replace(/\s+/g, " ").trim();

export default function App() {
  /* ------------------------------------------------------------ settings & theme */
  const [settings, setSettings] = useState<ZeeSettings>(() => loadSettings());
  const settingsRef = useRef(settings);
  settingsRef.current = settings;
  const getSettings = useCallback(() => settingsRef.current, []);
  const updateSettings = useCallback((patch: Partial<ZeeSettings>) => {
    setSettings((prev) => {
      const next = { ...prev, ...patch };
      saveSettings(next);
      return next;
    });
  }, []);
  const autoDark = useAutoDark();
  const dark = settings.theme === "auto" ? autoDark : settings.theme === "dark";
  useBodyThemeClass(dark);
  // Word may finish initialising Office.js after the pane has rendered (main.tsx renders anyway after 8 s,
  // and a cleared add-in cache makes the first start slow). Re-check instead of freezing in preview mode.
  const [inWord, setInWord] = useState(() => isWordAvailable());
  useEffect(() => {
    if (inWord || typeof Office === "undefined") return;
    let alive = true;
    const check = () => {
      if (alive && isWordAvailable()) setInWord(true);
    };
    Office.onReady().then(check, () => undefined);
    const timer = window.setInterval(check, 1000);
    const stop = window.setTimeout(() => window.clearInterval(timer), 60_000);
    return () => {
      alive = false;
      window.clearInterval(timer);
      window.clearTimeout(stop);
    };
  }, [inWord]);
  const docTitle = useMemo(() => getDocTitle(), [inWord]);

  // Refresh the shared team config (glossary + actions) from its URL / the team proxy once per pane load.
  useEffect(() => {
    const src = settingsRef.current.teamSource;
    if (!src || src === "file") return;
    fetchTeamConfig(src, settingsRef.current)
      .then((team) => updateSettings({ team, teamLoadedAt: Date.now() }))
      .catch(() => undefined); // offline: keep the last loaded copy
  }, [updateSettings]);

  /* ------------------------------------------------------------ toasts */
  const toasterId = useId("toaster");
  const { dispatchToast } = useToastController(toasterId);
  const toast = useCallback(
    (intent: ToastIntent, title: string, body?: string) =>
      dispatchToast(
        <Toast>
          <ToastTitle>{title}</ToastTitle>
          {body && <ToastBody>{body}</ToastBody>}
        </Toast>,
        { intent, timeout: intent === "error" ? 8000 : 4000 }
      ),
    [dispatchToast]
  );
  const toastError = useCallback((e: unknown) => toast("error", "Something went wrong", (e as Error)?.message ?? String(e)), [toast]);

  /* ------------------------------------------------------------ usage */
  const [usage, setUsage] = useState<UsageTotals>({ requests: 0, promptTokens: 0, outputTokens: 0 });
  const [today, setToday] = useState(() => usageForDay());
  const addUsage = useCallback((u: Usage) => {
    setUsage((p) => ({ requests: p.requests + 1, promptTokens: p.promptTokens + u.promptTokens, outputTokens: p.outputTokens + u.outputTokens }));
    setToday(recordUsage(u));
  }, []);
  const prices = useMemo(() => ({ inPerM: settings.priceInPerM, outPerM: settings.priceOutPerM }), [settings.priceInPerM, settings.priceOutPerM]);

  /* ------------------------------------------------------------ updates, usage counts */
  const updates = useUpdates();
  const [useCounts, setUseCounts] = useState(() => loadUseCounts());
  const used = useCallback((key: string) => setUseCounts(bumpUse(key)), []);

  /* ------------------------------------------------------------ selection tracking */
  const [selectionWords, setSelectionWords] = useState(0);
  /** × on the selection chip: leave the selection out of the next message (resets when the selection changes). */
  const [ignoreSelection, setIgnoreSelection] = useState(false);
  useEffect(() => {
    if (!inWord) return;
    const refresh = () =>
      getSelectionInfo()
        .then((s) => {
          setSelectionWords(s.isEmpty ? 0 : countWords(s.text));
          setIgnoreSelection(false);
        })
        .catch(() => setSelectionWords(0));
    refresh();
    return onSelectionChange(refresh);
  }, [inWord]);

  /* ------------------------------------------------------------ undo */
  const [undoLabel, setUndoLabel] = useState<string | null>(null);
  const [undoDepth, setUndoDepth] = useState(0);
  useEffect(
    () =>
      onUndoChange((label, depth) => {
        setUndoLabel(label);
        setUndoDepth(depth);
      }),
    []
  );
  /* ------------------------------------------------------------ Zee AI's tracked changes */
  const [zeeRevisions, setZeeRevisions] = useState(0);
  const canResolve = inWord && supportsWordApi("1.6");
  const refreshRevisions = useCallback(() => {
    if (!canResolve) return;
    countZeeRevisions()
      .then(setZeeRevisions)
      .catch(() => undefined);
  }, [canResolve]);
  // Markers left by an earlier session can't be undone any more – unwrap them (text stays; markers
  // around tracked edits stay until those revisions are accepted or rejected).
  useEffect(() => {
    if (inWord) cleanupOrphanMarkers().then(refreshRevisions);
  }, [inWord, refreshRevisions]);
  useEffect(() => {
    if (settings.trackChanges || zeeRevisions) refreshRevisions();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [undoDepth, undoLabel]);
  const resolveRevisions = async (accept: boolean) => {
    try {
      const r = await resolveZeeRevisions(accept);
      setZeeRevisions(0);
      toast(
        "success",
        accept ? "Zee AI's changes accepted" : "Zee AI's changes rejected",
        `${r.changes} revision${r.changes === 1 ? "" : "s"} in ${r.edits} edit${r.edits === 1 ? "" : "s"}. Your own tracked changes were not touched.`
      );
    } catch (e) {
      toastError(e);
      refreshRevisions();
    }
  };
  const undo = async () => {
    try {
      const label = await undoLast();
      toast("success", "Undone", `Reverted: ${label}`);
    } catch (e) {
      toastError(e);
    }
  };

  /* ------------------------------------------------------------ confirm (promise based) */
  const [confirmState, setConfirmState] = useState<ConfirmState | null>(null);
  const confirmResolver = useRef<((ok: boolean) => void) | null>(null);
  const askConfirm = useCallback(
    (state: ConfirmState) =>
      new Promise<boolean>((resolve) => {
        confirmResolver.current?.(false);
        confirmResolver.current = resolve;
        setConfirmState(state);
      }),
    []
  );
  const onConfirmResult = (ok: boolean) => {
    confirmResolver.current?.(ok);
    confirmResolver.current = null;
    setConfirmState(null);
  };

  /** Cost guard: ask before sending a large prompt. */
  const confirmLarge = useCallback(
    (tokens: number) =>
      askConfirm({
        title: "Large request",
        message: `This request is about ${tokens.toLocaleString()} tokens (document text + chat history), above your cost guard of ${settingsRef.current.costGuardTokens.toLocaleString()}. Send it anyway? To reduce it, select less text, switch the chat to "Selected text", or start a new chat.`,
        confirmLabel: "Send",
      }),
    [askConfirm]
  );

  const { messages, busy, send, runAction, runGrammar, runTask, retry, markApplied, markCommented, stop, clear } = useChat({ getSettings, onUsage: addUsage, confirmLarge });

  /* ------------------------------------------------------------ "General" scope: locating */
  const [locating, setLocating] = useState<string | null>(null);
  const locateAbort = useRef<AbortController | null>(null);
  const working = busy || !!locating;
  const stopAll = useCallback(() => {
    locateAbort.current?.abort();
    stop();
  }, [stop]);
  const abortable = useRef(false);
  abortable.current = busy;
  // Esc stops a running answer from anywhere in the pane.
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.key === "Escape" && (abortable.current || locateAbort.current)) stopAll();
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [stopAll]);

  const [settingsOpen, setSettingsOpen] = useState(false);
  const ensureConfigured = useCallback(() => {
    if (isConfigured(settingsRef.current)) return true;
    setSettingsOpen(true);
    toast("warning", "Connect Gemini first", "Add an API key, sign in with Google or enter your team proxy.");
    return false;
  }, [toast]);

  /* ------------------------------------------------------------ document context */
  const docContext = useCallback(
    async (mode: DocContext): Promise<{ suffix: string; label?: string; text: string; numbered?: string; paragraphs?: string[]; citations?: boolean }> => {
      if (!inWord || mode === "none") return { suffix: "", text: "" };
      try {
        if (mode === "selection") {
          const s = await getSelectionInfo();
          if (s.isEmpty) return { suffix: "", text: "" };
          const t = truncateMiddle(s.text, settingsRef.current.maxContextChars);
          return {
            suffix: `\n\n[Selected text${t.truncated ? " – trimmed" : ""}]\n"""\n${t.text}\n"""`,
            label: `Selection · ${countWords(s.text).toLocaleString()} words${t.truncated ? " (trimmed)" : ""}`,
            text: s.text,
          };
        }
        const paras = await getDocumentParagraphs();
        const body = paras.filter(Boolean).join("\n");
        if (!body.trim()) return { suffix: "", text: "" };
        // Paragraph numbers let answers cite [P12] – clicking a citation jumps to that paragraph.
        const t = truncateMiddle(numberParagraphs(paras), settingsRef.current.maxContextChars);
        return {
          suffix: `\n\n[Document "${docTitle}", paragraphs numbered [P#]${t.truncated ? " – middle trimmed" : ""}]\n"""\n${t.text}\n"""`,
          label: `Whole document · ${countWords(body).toLocaleString()} words${t.truncated ? " (trimmed)" : ""}`,
          text: body,
          numbered: t.text,
          paragraphs: paras,
          citations: true,
        };
      } catch (e) {
        toast("warning", "Document text not included", (e as Error).message);
        return { suffix: "", text: "" };
      }
    },
    [docTitle, inWord, toast]
  );

  /**
   * "General" scope: a cheap first call finds where in the document the request applies, then Zee AI
   * selects that part and edits it (diff + Replace), writes new text for that spot, or answers from
   * just those paragraphs.
   */
  const runGeneral = useCallback(
    async (text: string) => {
      if (!inWord) {
        await send({ display: text, prompt: text });
        return;
      }
      let paras: string[];
      try {
        paras = await getDocumentParagraphs();
      } catch (e) {
        toastError(e);
        return;
      }
      if (!paras.some((p) => p.trim())) {
        // Empty document: just write what was asked at the cursor.
        await runAction({
          kind: "draft",
          title: "Draft",
          display: text,
          prompt: draftPrompt(text, { length: "medium", tone: "match" }),
          system: DRAFT_SYSTEM_PROMPT,
          contextLabel: "General · empty document",
          format: "markdown",
          defaultLocation: "after",
        });
        return;
      }
      const s = settingsRef.current;
      const numbered = truncateMiddle(numberParagraphs(paras), Math.max(s.maxContextChars, 120_000));
      const tokens = estimateTokens(numbered.text + text + LOCATE_SYSTEM_PROMPT);
      if (s.costGuardTokens > 0 && tokens > s.costGuardTokens && !(await confirmLarge(tokens))) return;

      const abort = new AbortController();
      locateAbort.current = abort;
      setLocating("Finding the relevant part of the document…");
      let plan: LocatePlan;
      try {
        const { data, usage: u } = await generateJson<LocateResult>({
          settings: bulkSettings(s),
          prompt: locatePrompt(text, numbered.text),
          systemInstruction: withGlossary(LOCATE_SYSTEM_PROMPT, protectedTermList(s)),
          schema: LOCATE_SCHEMA,
          signal: abort.signal,
        });
        if (u) addUsage(u);
        plan = normalizeLocate(data, paras);
      } catch (e) {
        if (abort.signal.aborted || (e instanceof GeminiError && e.kind === "aborted")) return;
        toast("warning", "Couldn't find the spot automatically", `${(e as Error).message} Answering from the whole document instead.`);
        plan = { action: "none", reason: "" };
      } finally {
        setLocating(null);
        locateAbort.current = null;
      }
      const why = plan.reason ? ` – ${plan.reason}` : "";

      if (plan.action === "edit") {
        const words = countWords(paras.slice(plan.start - 1, plan.end).join(" "));
        if (words * 6 > MAX_EDIT_CHARS) {
          toast(
            "warning",
            "That part is too long to edit in one go",
            `${paraLabel(plan.start, plan.end)} has about ${words.toLocaleString()} words. Ask for one section at a time, or select less and use “Selected text”.`
          );
          return;
        }
        let selected: string;
        try {
          selected = await selectParagraphRange(plan.start, plan.end);
        } catch (e) {
          toastError(e);
          return;
        }
        if (!selected.trim()) {
          toast("info", "Nothing to edit there", "The located paragraphs are empty – try describing the part more precisely.");
          return;
        }
        await runAction({
          kind: "action",
          title: `Edit ${paraLabel(plan.start, plan.end)}`,
          display: text,
          prompt: generalEditPrompt(text, selected),
          system: EDIT_SYSTEM_PROMPT,
          original: selected,
          contextLabel: `Found ${paraLabel(plan.start, plan.end)} · ${countWords(selected).toLocaleString()} words${why}`,
          format: "plain",
          defaultLocation: "replace",
          extra: { paraRange: [plan.start, plan.end] },
        });
        return;
      }

      if (plan.action === "insert") {
        let anchor = plan.after;
        if (anchor > 0) {
          try {
            await selectParagraph(anchor); // show where the new text will go
          } catch {
            anchor = 0;
          }
        }
        const last = paras.reduce((k, p, i) => (p.trim() ? i + 1 : k), 1);
        const around = relevantExcerpt(paras, [anchor || last], 3, 12_000).text;
        await runAction({
          kind: "draft",
          title: anchor ? `New text after ${paraLabel(anchor)}` : "New text at the end",
          display: text,
          prompt: generalInsertPrompt(text, paras[(anchor || last) - 1] ?? "", around),
          system: DRAFT_SYSTEM_PROMPT,
          contextLabel: `${anchor ? `Goes after ${paraLabel(anchor)}` : "Goes at the end"}${why}`,
          format: "markdown",
          defaultLocation: anchor ? "after" : "end",
          extra: anchor ? { anchorParagraph: anchor, anchorText: paras[anchor - 1] } : undefined,
        });
        return;
      }

      if (plan.action === "answer") {
        const ex = relevantExcerpt(paras, plan.paragraphs, 1, Math.min(Math.max(s.maxContextChars, 8000), 40_000));
        await send({
          display: text,
          prompt: `${text}\n\n[Relevant parts of the document "${docTitle}" found by Zee AI – paragraphs numbered [P#]; "…" marks skipped text]\n"""\n${ex.text}\n"""`,
          contextLabel: `Found ${paraListLabel(ex.used)}${why}`,
          citations: true,
        });
        return;
      }

      // No specific part: answer with the whole document as context.
      const doc = await docContext("document");
      await send({ display: text, prompt: text + doc.suffix, contextLabel: doc.label ? `${doc.label} · no specific part found` : undefined, citations: doc.citations });
    },
    [addUsage, confirmLarge, docContext, docTitle, inWord, runAction, send, toast, toastError]
  );

  const sendChat = useCallback(
    async (text: string) => {
      if (!ensureConfigured()) return;
      const mode = settingsRef.current.docContext;
      if (mode === "general") {
        await runGeneral(text);
        return;
      }
      const skip = mode === "selection" && ignoreSelection;
      setIgnoreSelection(false);
      const ctx = await docContext(skip ? "none" : mode);
      await send({ display: text, prompt: text + ctx.suffix, contextLabel: ctx.label, citations: ctx.citations });
    },
    [docContext, ensureConfigured, ignoreSelection, runGeneral, send]
  );

  /* ------------------------------------------------------------ selection actions */
  /** Reads the selection for an edit action; returns null (with a hint) when nothing usable is selected. */
  const requireSelection = useCallback(async (): Promise<string | null> => {
    if (!inWord) {
      toast("info", "Open Zee AI in Word", "Selection actions work on text selected in a Word document.");
      return null;
    }
    try {
      const s = await getSelectionInfo();
      if (s.isEmpty) {
        toast("info", "Select some text first", "Highlight the words, sentences or paragraphs you want Zee AI to work on.");
        return null;
      }
      if (s.text.length > MAX_EDIT_CHARS) {
        toast("warning", "Selection too long", `Select up to about ${(MAX_EDIT_CHARS / 6).toLocaleString()} words at a time (a few pages).`);
        return null;
      }
      return s.text;
    } catch (e) {
      toastError(e);
      return null;
    }
  }, [inWord, toast, toastError]);

  /** Selected text without hints (empty string when nothing usable is selected). */
  const peekSelection = useCallback(async (): Promise<string> => {
    if (!inWord) return "";
    try {
      const s = await getSelectionInfo();
      return s.isEmpty ? "" : s.text;
    } catch {
      return "";
    }
  }, [inWord]);

  const selLabel = (text: string) => `Selection · ${countWords(text).toLocaleString()} words`;

  const editAction = useCallback(
    async (title: string, build: (text: string) => string) => {
      if (!ensureConfigured()) return;
      const text = await requireSelection();
      if (!text) return;
      await runAction({
        kind: "action",
        title,
        display: `${title}: “${shorten(text, 80)}”`,
        prompt: build(text),
        system: EDIT_SYSTEM_PROMPT,
        original: text,
        contextLabel: selLabel(text),
        format: "plain",
        defaultLocation: "replace",
      });
    },
    [ensureConfigured, requireSelection, runAction]
  );

  const onRewrite = () => editAction("Rewrite & polish", rewritePrompt);
  const onTone = (tone: ToneId) => (used(useKey.tone(tone)), editAction(`Tone: ${TONES.find((t) => t.id === tone)?.label ?? tone}`, (t) => tonePrompt(t, tone)));
  const onLength = (dir: "expand" | "shorten") => editAction(dir === "expand" ? "Expand" : "Shorten", (t) => lengthPrompt(t, dir));
  const onTranslate = (language: string) => {
    used(useKey.translate(language));
    if (language !== settingsRef.current.translateTo) updateSettings({ translateTo: language });
    return editAction(`Translate to ${language}`, (t) => translatePrompt(t, language));
  };

  /** English + Arabic version of the selection, inserted as Markdown (each paragraph gets its own direction). */
  const onBilingual = async (layout: "stacked" | "table") => {
    if (!ensureConfigured()) return;
    const text = await requireSelection();
    if (!text) return;
    const title = layout === "table" ? "Bilingual – side by side" : "Bilingual – English + Arabic";
    await runAction({
      kind: "action",
      title,
      display: `${title}: “${shorten(text, 80)}”`,
      prompt: bilingualPrompt(text, layout),
      system: BILINGUAL_SYSTEM_PROMPT,
      contextLabel: selLabel(text),
      format: "markdown",
      defaultLocation: "replace",
    });
  };

  const onGrammar = async () => {
    if (!ensureConfigured()) return;
    if (inWord && !(await peekSelection())) {
      await proofreadDocument();
      return;
    }
    const text = await requireSelection();
    if (!text) return;
    await runGrammar({ display: `Grammar & spelling: “${shorten(text, 80)}”`, original: text, contextLabel: selLabel(text) });
  };

  const onSummarize = async (style: SummaryStyle) => {
    if (!ensureConfigured()) return;
    if (!inWord) {
      toast("info", "Open Zee AI in Word", "Summarize works on the selection or the whole document.");
      return;
    }
    const name = style === "bullets" ? "Summary (bullets)" : style === "executive" ? "Executive summary" : "TL;DR";
    // Selection if there is one, otherwise the whole document.
    const sel = await docContext("selection");
    if (sel.text) {
      const t = truncateMiddle(sel.text, settingsRef.current.maxContextChars);
      await runAction({
        kind: "action",
        title: name,
        display: `${name} of the selection`,
        prompt: summarizePrompt(t.text, style),
        system: SUMMARY_SYSTEM_PROMPT,
        contextLabel: sel.label,
        format: "markdown",
        defaultLocation: "after",
      });
      return;
    }
    const doc = await docContext("document");
    if (!doc.text || !doc.paragraphs || !doc.numbered) {
      toast("info", "Nothing to summarize", "The document is empty.");
      return;
    }
    if (numberParagraphs(doc.paragraphs).length <= settingsRef.current.maxContextChars) {
      await runAction({
        kind: "action",
        title: name,
        display: `${name} of the document`,
        prompt: summarizePrompt(doc.numbered, style),
        system: SUMMARY_SYSTEM_PROMPT,
        contextLabel: doc.label,
        format: "markdown",
        defaultLocation: "after",
        citations: true,
      });
      return;
    }
    await summarizeLongDocument(doc.paragraphs, style, name, doc.label);
  };

  /* ------------------------------------------------------------ drafting */
  const [draftOpen, setDraftOpen] = useState(false);
  const onDraft = async (c: DraftChoice) => {
    setDraftOpen(false);
    if (!ensureConfigured()) return;
    let context: string | undefined;
    let label: string | undefined;
    if (c.useContext && inWord) {
      let ctx = await docContext("selection");
      if (!ctx.text) ctx = await docContext("document");
      if (ctx.text) {
        context = truncateMiddle(ctx.text, Math.min(settingsRef.current.maxContextChars, 20000)).text;
        label = ctx.label;
      }
    }
    await runAction({
      kind: "draft",
      title: "Draft",
      display: `Draft: ${c.request}`,
      prompt: draftPrompt(c.request, { length: c.length, tone: c.tone, context }),
      system: DRAFT_SYSTEM_PROMPT,
      contextLabel: label,
      format: "markdown",
      defaultLocation: "after",
    });
  };

  /* ------------------------------------------------------------ whole-document tasks (chunked) */
  const proofreadDocument = useCallback(async () => {
    let paras: string[];
    try {
      paras = await getDocumentParagraphs();
    } catch (e) {
      toastError(e);
      return;
    }
    const chunks = chunkParagraphs(paras, CHUNK_CHARS);
    if (!chunks.length) {
      toast("info", "Nothing to check", "The document is empty.");
      return;
    }
    const words = countWords(paras.join(" "));
    const ok = await askConfirm({
      title: "Proofread the whole document?",
      message: `Nothing is selected, so Zee AI will check all ${words.toLocaleString()} words in ${chunks.length} part${chunks.length === 1 ? "" : "s"}. Nothing is changed: you get a list of issues that you can add as Word comments.`,
      confirmLabel: "Proofread",
    });
    if (!ok) return;
    await runTask({
      display: "Grammar & spelling: whole document",
      prompt: "Proofread the whole document.",
      estimateText: paras.join("\n").repeat(2), // input + corrected text coming back
      contextLabel: `Whole document · ${words.toLocaleString()} words`,
      assistant: { kind: "action", title: "Grammar & spelling", scope: "document", format: "plain", issueStyle: "fix" },
      bulk: true,
      work: async ({ settings, signal, system, progress, addUsage }) => {
        const issues: GrammarIssue[] = [];
        for (let i = 0; i < chunks.length; i++) {
          progress(`part ${i + 1} of ${chunks.length}`);
          const { data, usage } = await generateJson<GrammarResult>({
            settings,
            prompt: grammarChunkPrompt(chunks[i].text),
            systemInstruction: system(GRAMMAR_SYSTEM_PROMPT),
            schema: GRAMMAR_SCHEMA,
            signal,
          });
          addUsage(usage);
          (Array.isArray(data?.issues) ? data.issues : []).forEach((x) => {
            if (x && typeof x.original === "string" && x.original.trim() && x.original !== x.fix) issues.push(x);
          });
        }
        return {
          text: issues.length
            ? `Found ${issues.length} possible issue${issues.length === 1 ? "" : "s"}. Nothing was changed – add them as Word comments to review each one in place.`
            : "",
          issues,
        };
      },
    });
  }, [askConfirm, runTask, toast, toastError]);

  const summarizeLongDocument = useCallback(
    async (paras: string[], style: SummaryStyle, name: string, label?: string) => {
      const chunks = chunkParagraphs(
        paras.map((p, i) => (p ? `[P${i + 1}] ${p}` : "")),
        CHUNK_CHARS
      );
      await runTask({
        display: `${name} of the document (${chunks.length} parts)`,
        prompt: `${name} of the whole document.`,
        estimateText: paras.join("\n"),
        contextLabel: label,
        assistant: { kind: "action", title: name, format: "markdown", defaultLocation: "after", citations: true },
        bulk: true,
        work: async ({ settings, signal, system, progress, addUsage }) => {
          const notes: string[] = [];
          for (let i = 0; i < chunks.length; i++) {
            progress(`reading part ${i + 1} of ${chunks.length}`);
            const r = await generateText({ settings, prompt: chunkSummaryPrompt(chunks[i].text, chunks[i].firstParagraph), systemInstruction: system(SUMMARY_SYSTEM_PROMPT), signal });
            addUsage(r.usage);
            notes.push(r.text);
          }
          progress("writing the summary");
          const final = await generateText({ settings, prompt: combineSummaryPrompt(notes, style), systemInstruction: system(SUMMARY_SYSTEM_PROMPT), signal });
          addUsage(final.usage);
          return { text: final.text || "_(empty response)_" };
        },
      });
    },
    [runTask]
  );

  /* ------------------------------------------------------------ findings helper (consistency, reference) */
  const findingsToIssues = (findings: unknown): GrammarIssue[] =>
    (Array.isArray(findings) ? (findings as DocFinding[]) : [])
      .filter((f) => f && typeof f.problem === "string" && f.problem.trim())
      .map((f) => ({
        original: typeof f.quote === "string" ? f.quote : "",
        fix: typeof f.suggestion === "string" ? f.suggestion : "",
        type: typeof f.category === "string" ? f.category : "finding",
        reason: f.problem,
        paragraph: Number.isFinite(f.paragraph) && f.paragraph > 0 ? Math.round(f.paragraph) : undefined,
      }));

  /* ------------------------------------------------------------ consistency check */
  const onConsistency = useCallback(async () => {
    if (!ensureConfigured()) return;
    if (!inWord) {
      toast("info", "Open Zee AI in Word", "The consistency check reads the whole document.");
      return;
    }
    let paras: string[];
    try {
      paras = await getDocumentParagraphs();
    } catch (e) {
      toastError(e);
      return;
    }
    const numbered = numberParagraphs(paras);
    if (!numbered.trim()) {
      toast("info", "Nothing to check", "The document is empty.");
      return;
    }
    const t = truncateMiddle(numbered, Math.max(settingsRef.current.maxContextChars, 120_000));
    const words = countWords(paras.join(" "));
    await runTask({
      display: "Consistency check: whole document",
      prompt: "Check the whole document for inconsistent units, names, numbers, dates and terms.",
      estimateText: t.text,
      contextLabel: `Whole document · ${words.toLocaleString()} words${t.truncated ? " (trimmed)" : ""}`,
      assistant: { kind: "action", title: "Consistency check", scope: "document", format: "plain", issueStyle: "finding" },
      bulk: false, // needs the whole document in one view – use the main model
      work: async ({ settings, signal, system, addUsage }) => {
        const { data, usage } = await generateJson<{ summary: string; findings: DocFinding[] }>({
          settings,
          prompt: consistencyPrompt(t.text),
          systemInstruction: system(CONSISTENCY_SYSTEM_PROMPT),
          schema: FINDINGS_SCHEMA,
          signal,
        });
        addUsage(usage);
        const issues = findingsToIssues(data?.findings);
        return { text: `${data?.summary ?? ""}${issues.length ? ` ${issues.length} finding${issues.length === 1 ? "" : "s"} – add them as comments to fix each in place.` : ""}`.trim(), issues };
      },
    });
  }, [ensureConfigured, inWord, runTask, toast, toastError]);

  /* ------------------------------------------------------------ check against a reference */
  const [referenceOpen, setReferenceOpen] = useState(false);
  const onReference = useCallback(
    async (c: ReferenceChoice) => {
      setReferenceOpen(false);
      if (!ensureConfigured()) return;
      let paras: string[];
      try {
        paras = await getDocumentParagraphs();
      } catch (e) {
        toastError(e);
        return;
      }
      const t = truncateMiddle(numberParagraphs(paras), Math.max(settingsRef.current.maxContextChars, 120_000));
      if (!t.text.trim()) {
        toast("info", "Nothing to check", "The document is empty.");
        return;
      }
      const refText = c.reference.kind === "text" ? truncateMiddle(c.reference.text, 400_000).text : "";
      const parts: GeminiPart[] =
        c.reference.kind === "inline"
          ? [{ text: `REFERENCE DOCUMENT: ${c.reference.name} (attached)` }, { inlineData: { mimeType: c.reference.mimeType, data: c.reference.data } }]
          : [{ text: `--- REFERENCE "${c.reference.name}" START ---\n${refText}\n--- REFERENCE END ---` }];
      parts.push({ text: referenceCheckPrompt(c.question, t.text) });
      await runTask({
        display: `Check against “${c.reference.name}”: ${shorten(c.question, 90)}`,
        prompt: `Check the document against the reference "${c.reference.name}": ${c.question}`,
        // PDFs are billed at roughly 260 tokens per page; estimate from the file size.
        estimateText: t.text + refText + (c.reference.kind === "inline" ? "x".repeat(Math.min(4_000_000, c.reference.data.length / 3)) : ""),
        contextLabel: `Document + ${c.reference.name}`,
        assistant: { kind: "action", title: "Reference check", scope: "document", format: "plain", issueStyle: "finding" },
        work: async ({ settings, signal, system, addUsage }) => {
          const r = await generateFromParts<{ summary: string; findings: DocFinding[] }>({
            settings,
            parts,
            systemInstruction: system(REFERENCE_SYSTEM_PROMPT),
            schema: FINDINGS_SCHEMA,
            signal,
          });
          addUsage(r.usage);
          const issues = findingsToIssues(r.data?.findings);
          return { text: r.data?.summary ?? "", issues };
        },
      });
    },
    [ensureConfigured, runTask, toast, toastError]
  );

  /* ------------------------------------------------------------ tables */
  const onTable = async (action: "explain" | "commentary" | "clean" | "fromNotes") => {
    if (!ensureConfigured()) return;
    if (action === "fromNotes") {
      const text = await requireSelection();
      if (!text) return;
      await runAction({
        kind: "action",
        title: "Notes → table",
        display: `Turn into a table: “${shorten(text, 80)}”`,
        prompt: notesToTablePrompt(text),
        system: TABLE_SYSTEM_PROMPT,
        contextLabel: selLabel(text),
        format: "markdown",
        defaultLocation: "after",
      });
      return;
    }
    let table: Awaited<ReturnType<typeof getSelectedTable>>;
    try {
      table = await getSelectedTable();
    } catch (e) {
      toastError(e);
      return;
    }
    if (!table || !table.rows) {
      toast("info", "Click inside a table first", "Put the cursor anywhere in the table you want Zee AI to work on.");
      return;
    }
    const md = tableToMarkdown(table.values);
    const label = `Table · ${table.rows} × ${table.cols}`;
    if (action === "explain") {
      await runAction({ kind: "action", title: "Explain table", display: "Explain this table", prompt: explainTablePrompt(md), system: TABLE_SYSTEM_PROMPT, contextLabel: label, format: "markdown", defaultLocation: "after" });
      return;
    }
    if (action === "commentary") {
      await runAction({
        kind: "action",
        title: "Table commentary",
        display: "Write a commentary for this table",
        prompt: tableCommentaryPrompt(md),
        system: TABLE_SYSTEM_PROMPT,
        contextLabel: label,
        format: "markdown",
        defaultLocation: "after",
      });
      return;
    }
    const before = table.values;
    await runTask({
      display: "Clean up this table",
      prompt: "Clean up the table (spelling, units, formatting) without changing values.",
      estimateText: JSON.stringify(before).repeat(2),
      contextLabel: label,
      assistant: { kind: "action", title: "Table clean-up", format: "plain" },
      work: async ({ settings, signal, system, addUsage }) => {
        const { data, usage } = await generateJson<{ cells: string[][]; notes: string[] }>({
          settings,
          prompt: tableCleanPrompt(before),
          systemInstruction: system(TABLE_SYSTEM_PROMPT),
          schema: TABLE_CLEAN_SCHEMA,
          signal,
        });
        addUsage(usage);
        let after = fitGrid(before, data?.cells);
        // Safety net: any cell whose numbers changed is reverted to the original.
        let reverted = 0;
        after = after.map((r, i) =>
          r.map((v, j) => {
            if (keepsNumbers(before[i][j], v)) return v;
            reverted++;
            return before[i][j];
          })
        );
        const notes = (Array.isArray(data?.notes) ? data.notes : []).filter((n) => typeof n === "string").slice(0, 8);
        if (reverted) notes.push(`${reverted} proposed change${reverted === 1 ? "" : "s"} would have altered a number and ${reverted === 1 ? "was" : "were"} discarded.`);
        const changes = diffGrid(before, after).length;
        return { text: changes ? `${changes} cell${changes === 1 ? "" : "s"} to tidy.` : "The table already looks clean.", tableEdit: { before, after, notes } };
      },
    });
  };

  const applyTable = useCallback(
    async (m: ChatMessage) => {
      if (!m.tableEdit) return;
      try {
        const r = await writeTableCells(m.tableEdit.after, "Table clean-up", settingsRef.current.trackChanges);
        markApplied(m.id, "replace");
        toast("success", `Updated ${r.changed} cell${r.changed === 1 ? "" : "s"}`, `${r.tracked ? "Recorded as Track Changes. " : ""}Undo with ↶.`);
      } catch (e) {
        toastError(e);
      }
    },
    [markApplied, toast, toastError]
  );

  /* ------------------------------------------------------------ pictures */
  const onPicture = async (kind: "caption" | "alt" | "describe") => {
    if (!ensureConfigured()) return;
    if (!inWord) {
      toast("info", "Open Zee AI in Word", "Select a picture in the document first.");
      return;
    }
    let pic: Awaited<ReturnType<typeof getSelectedPicture>>;
    try {
      pic = await getSelectedPicture();
    } catch (e) {
      toastError(e);
      return;
    }
    if (!pic) {
      toast("info", "Select a picture first", "Click a picture or chart in the document (in line with text), then try again.");
      return;
    }
    const title = kind === "caption" ? "Picture caption" : kind === "alt" ? "Alt text" : "Picture analysis";
    const image = pic;
    await runTask({
      display: kind === "caption" ? "Write a caption for the selected picture" : kind === "alt" ? "Generate alt text for the selected picture" : "Describe / analyse the selected picture",
      prompt: `${title} for the selected picture.`,
      estimateText: "x".repeat(1200 * 4) + image.context,
      contextLabel: image.altText ? `Picture · current alt text: “${shorten(image.altText, 40)}”` : "Picture",
      assistant: { kind: "action", title, format: kind === "describe" ? "markdown" : "plain", defaultLocation: "after", pictureResult: kind === "describe" ? undefined : { kind } },
      work: async ({ settings, signal, system, addUsage }) => {
        const r = await generateFromParts({
          settings,
          parts: [{ inlineData: { mimeType: detectImageMime(image.base64), data: image.base64 } }, { text: picturePrompt(kind, image.context) }],
          systemInstruction: system(PICTURE_SYSTEM_PROMPT),
          signal,
        });
        addUsage(r.usage);
        return { text: kind === "describe" ? r.text : r.text.replace(/^["“]|["”]$/g, "").trim() };
      },
    });
  };

  const applyPicture = useCallback(
    async (m: ChatMessage) => {
      if (!m.pictureResult) return;
      try {
        if (m.pictureResult.kind === "caption") {
          await insertPictureCaption(m.text.trim(), "Picture caption", settingsRef.current.trackChanges);
          toast("success", "Caption inserted", "Added below the picture in Word's Caption style. Undo with ↶.");
        } else {
          await setPictureAltText(m.text.trim(), "Alt text");
          toast("success", "Alt text set", "Screen readers will now read it. Undo with ↶.");
        }
        markApplied(m.id, "after");
      } catch (e) {
        toastError(e);
      }
    },
    [markApplied, toast, toastError]
  );

  /* ------------------------------------------------------------ report templates */
  const [templatePick, setTemplatePick] = useState<ReportTemplate | null>(null);
  const onTemplateRun = async (c: TemplateChoice) => {
    setTemplatePick(null);
    used(useKey.template(c.template.id));
    if (c.fillFromSelection) {
      if (!ensureConfigured()) return;
      const notes = await requireSelection();
      if (!notes) return;
      await runAction({
        kind: "draft",
        title: c.template.name,
        display: `${c.template.name} from the selected notes`,
        prompt: templateFillPrompt(c.template.outline, notes),
        system: DRAFT_SYSTEM_PROMPT,
        contextLabel: selLabel(notes),
        format: "markdown",
        defaultLocation: c.location,
      });
      return;
    }
    try {
      await insertContent({ content: c.template.outline, format: "markdown", location: c.location, label: `Template: ${c.template.name}`, trackChanges: settingsRef.current.trackChanges });
      toast("success", "Template inserted", "Replace the [placeholders]. References → Table of Contents works with these headings. Undo with ↶.");
    } catch (e) {
      toastError(e);
    }
  };

  /* ------------------------------------------------------------ my quick actions */
  const [settingsFocus, setSettingsFocus] = useState<"actions" | null>(null);
  const onCustom = async (a: CustomAction) => {
    if (!ensureConfigured()) return;
    used(useKey.custom(a.id));
    if (a.mode === "replace") {
      await editAction(a.name, (t) => customActionPrompt(a.instruction, t));
      return;
    }
    const text = await requireSelection();
    if (!text) return;
    await runAction({
      kind: "action",
      title: a.name,
      display: `${a.name}: “${shorten(text, 80)}”`,
      prompt: customActionPrompt(a.instruction, text),
      system: DRAFT_SYSTEM_PROMPT,
      contextLabel: selLabel(text),
      format: "markdown",
      defaultLocation: "after",
    });
  };

  /* ------------------------------------------------------------ inserting into Word */
  const insert = useCallback(
    async (m: ChatMessage, location: InsertLocation) => {
      try {
        let content = m.text;
        let format: InsertFormat = m.format === "plain" ? "plain" : looksLikeMarkdown(content) ? "markdown" : "plain";
        if (format === "markdown" && !settingsRef.current.richInsert) {
          content = markdownToPlain(content);
          format = "plain";
        }
        if (format === "plain" && m.format !== "plain") content = markdownToPlain(content);

        // "General" scope: put the selection / cursor back on the located part first.
        if (location === "replace" && m.paraRange && m.original !== undefined) {
          // If the paragraphs changed meanwhile, the check below shows the difference and asks first.
          const again = await selectParagraphRange(m.paraRange[0], m.paraRange[1]).catch(() => "");
          if (!again) {
            toast("info", "The document changed", "Those paragraphs no longer exist – select the text to replace and click Replace again, or Regenerate.");
            return;
          }
        }
        if (location === "after" && m.anchorParagraph) await placeCursorAfterParagraph(m.anchorParagraph, m.anchorText);

        if (location === "replace") {
          const sel = await getSelectionInfo();
          if (sel.isEmpty) {
            const ok = await askConfirm({
              title: "Nothing is selected",
              message: "There is no selected text to replace. Insert at the cursor instead?",
              confirmLabel: "Insert at cursor",
            });
            if (!ok) return;
          } else if (m.original !== undefined) {
            // An edit action: the diff was already shown in the chat. Only ask if the selection moved.
            if (!sameText(sel.text, m.original)) {
              const ok = await askConfirm({
                title: "The selection has changed",
                message: "The selected text is no longer the text this suggestion was made for. Replace the current selection anyway? (Tip: select the original text again first.)",
                confirmLabel: "Replace anyway",
                diff: { before: sel.text, after: format === "plain" ? content : markdownToPlain(content) },
              });
              if (!ok) return;
            }
            if (format === "plain") content = preserveEdges(sel.text, content);
          } else {
            // A chat answer replacing text: always preview the change first.
            const ok = await askConfirm({
              title: "Replace the selected text?",
              message: `${countWords(sel.text).toLocaleString()} selected words will be replaced. Review the change:`,
              confirmLabel: "Replace",
              diff: { before: sel.text, after: format === "plain" ? content : markdownToPlain(content) },
            });
            if (!ok) return;
          }
        }

        const s = settingsRef.current;
        const r = await insertContent({
          content,
          format,
          location,
          label: `${m.title ?? "Chat answer"} – ${LOCATION_LABELS[location].toLowerCase()}`,
          preserveFormatting: s.preserveFormatting && m.original !== undefined,
          trackChanges: s.trackChanges,
          autoDirection: s.autoDirection,
        });
        markApplied(m.id, location);
        const notes = [
          r.minimal ? "Only the changed words were replaced – formatting kept." : "",
          r.tracked ? "Recorded as Track Changes (Review tab to accept/reject)." : "",
          r.undoable ? "Undo with ↶." : "Use Word's Undo (Ctrl+Z) to revert.",
        ].filter(Boolean);
        toast("success", location === "replace" ? "Text replaced" : "Text inserted", notes.join(" "));
      } catch (e) {
        toastError(e);
      }
    },
    [askConfirm, markApplied, toast, toastError]
  );

  const copyDiagnostics = useCallback(
    async (m: ChatMessage) => {
      const s = settingsRef.current;
      const report = buildDiagnostics({
        error: m.text,
        kind: m.errorKind,
        model: m.model ?? s.model,
        authMode: s.authMode,
        extra: {
          "Request": m.title ?? "chat",
          "Document context": s.docContext,
          "Track Changes": String(s.trackChanges),
          "WordApi 1.3 / 1.4 / 1.6": inWord ? ["1.3", "1.4", "1.6"].map((v) => (supportsWordApi(v as "1.3" | "1.4" | "1.6") ? "yes" : "no")).join(" / ") : "not in Word",
        },
      });
      if (await copyText(report)) toast("success", "Diagnostics copied", "Paste them to Claude or IT. No keys or document text are included.");
      else toast("error", "Couldn't copy", report.slice(0, 300));
    },
    [inWord, toast]
  );

  /** Edit & resend: put a prompt back into the chat box. */
  const [inject, setInject] = useState<{ text: string; n: number } | null>(null);
  const editPrompt = useCallback((text: string) => setInject({ text, n: Date.now() }), []);

  const comment = useCallback(
    async (m: ChatMessage) => {
      const issues = m.issues ?? [];
      if (!issues.length) return;
      try {
        const r = await addComments(
          issues.map((i) => ({
            find: i.original,
            paragraph: i.paragraph,
            text:
              m.issueStyle === "finding"
                ? `Zee AI · ${i.type || "finding"}: ${i.reason}${i.fix ? ` Suggestion: ${i.fix}` : ""}`
                : `Zee AI · ${i.type || "suggestion"}: “${i.original}” → “${i.fix}”${i.reason ? ` – ${i.reason}` : ""}`,
          })),
          `${m.title ?? "Grammar"} – ${issues.length} comment${issues.length === 1 ? "" : "s"}`
        );
        markCommented(m.id);
        toast(
          r.added ? "success" : "warning",
          r.added ? `Added ${r.added} comment${r.added === 1 ? "" : "s"}` : "No comments added",
          `${r.missing ? `${r.missing} passage${r.missing === 1 ? "" : "s"} couldn't be found (the text may have changed). ` : ""}Review them in Word's Review tab; undo with ↶.`
        );
      } catch (e) {
        toastError(e);
      }
    },
    [markCommented, toast, toastError]
  );

  const goToParagraph = useCallback(
    (n: number) => {
      selectParagraph(n).catch((e) => toastError(e));
    },
    [toastError]
  );

  const insertActions = useMemo<InsertActions>(
    () => ({ insert, comment, goToParagraph, applyTable, applyPicture, canInsert: inWord, retry: (id) => void retry(id), editPrompt, copyDiagnostics, busy: working }),
    [insert, comment, goToParagraph, applyTable, applyPicture, inWord, retry, editPrompt, copyDiagnostics, working]
  );

  /* ------------------------------------------------------------ ribbon / right-click / keyboard commands */
  const commandRef = useRef<(cmd: string) => void>(() => undefined);
  commandRef.current = (cmd: string) => {
    if (cmd === "rewrite") void onRewrite();
    else if (cmd === "grammar") void onGrammar();
    else if (cmd === "summarize") void onSummarize("bullets");
    else if (cmd === "consistency") void onConsistency();
  };
  useEffect(() => onCommand((cmd) => commandRef.current(cmd)), []);

  /* ------------------------------------------------------------ export chat */
  const chatAsText = useCallback(
    () =>
      messages
        .filter((m) => m.status !== "streaming")
        .map((m) => {
          const who = m.role === "user" ? "You" : `Zee AI${m.title ? ` (${m.title})` : ""}`;
          const ctx = m.contextLabel ? ` [${m.contextLabel}]` : "";
          return `### ${who}${ctx}\n\n${m.text}`;
        })
        .join("\n\n---\n\n"),
    [messages]
  );
  const exportCopy = async () => {
    if (await copyText(chatAsText())) toast("success", "Chat copied", "Paste it anywhere.");
    else toast("error", "Couldn't copy", "The clipboard is blocked here – use “Append chat to document” instead.");
  };
  const exportToDoc = async () => {
    try {
      await insertContent({ content: `# Zee AI chat – ${new Date().toLocaleString()}\n\n${chatAsText()}`, format: "markdown", location: "end", label: "Chat export" });
      toast("success", "Chat appended", "Added at the end of the document (undo with ↶).");
    } catch (e) {
      toastError(e);
    }
  };

  /* ------------------------------------------------------------ render */
  return (
    // Layout classes live on an inner div, NOT on FluentProvider: Fluent copies the provider's className
    // onto its portal containers (toasts, dialogs), and a full-height background class there covered the
    // whole pane with an empty layer in the Excel add-in.
    <FluentProvider theme={dark ? darkTheme : lightTheme} style={{ height: "100%" }}>
      <div className={`zee-app ${dark ? "zee-dark" : "zee-light"}`}>
        <InsertActionsContext.Provider value={insertActions}>
          <header className="zee-header">
            <img src="assets/icon-32.png" alt="" width={22} height={22} />
            <span className="zee-title">Zee AI</span>
            <Tooltip content="Change model in Settings" relationship="description">
              <Badge
                appearance="tint"
                size="small"
                className="zee-model-badge"
                role="button"
                tabIndex={0}
                onClick={() => setSettingsOpen(true)}
                onKeyDown={(e) => (e.key === "Enter" || e.key === " ") && setSettingsOpen(true)}
              >
                {modelLabel(settings.model)}
              </Badge>
            </Tooltip>
            <span className="zee-spacer" />
            <Tooltip content={undoLabel ? `Undo: ${undoLabel}${undoDepth > 1 ? ` (${undoDepth} steps available)` : ""}` : "Nothing to undo"} relationship="label">
              <Button appearance="subtle" size="small" icon={<ArrowUndo20Regular />} disabled={!undoLabel} onClick={undo}>
                {undoDepth > 1 ? undoDepth : undefined}
              </Button>
            </Tooltip>
            <Tooltip
              content={
                !supportsWordApi("1.4") && inWord
                  ? "Track Changes needs Word for Microsoft 365 / 2021+"
                  : settings.trackChanges
                    ? "Track Changes ON – Zee AI's edits become revisions you accept or reject"
                    : "Track Changes OFF – click to record Zee AI's edits as revisions"
              }
              relationship="label"
            >
              <ToggleButton
                appearance="subtle"
                size="small"
                checked={settings.trackChanges}
                className={settings.trackChanges ? "zee-track-on" : undefined}
                icon={settings.trackChanges ? <DocumentEdit20Filled /> : <DocumentEdit20Regular />}
                onClick={() => {
                  const on = !settingsRef.current.trackChanges;
                  updateSettings({ trackChanges: on });
                  toast("info", on ? "Track Changes on" : "Track Changes off", on ? "Zee AI's edits will appear as revisions in Word." : "Edits are applied directly (still undoable with ↶).");
                }}
              />
            </Tooltip>
            {zeeRevisions > 0 && (
              <Menu>
                <MenuTrigger disableButtonEnhancement>
                  <Tooltip content={`${zeeRevisions} tracked change${zeeRevisions === 1 ? "" : "s"} by Zee AI – accept or reject them (your own revisions stay)`} relationship="label">
                    <Button appearance="subtle" size="small" className="zee-track-on" icon={<DocumentCheckmark20Regular />}>
                      {zeeRevisions}
                    </Button>
                  </Tooltip>
                </MenuTrigger>
                <MenuPopover>
                  <MenuList>
                    <MenuItem icon={<CheckmarkCircle20Regular />} onClick={() => resolveRevisions(true)}>
                      Accept all Zee AI changes ({zeeRevisions})
                    </MenuItem>
                    <MenuItem icon={<DismissCircle20Regular />} onClick={() => resolveRevisions(false)}>
                      Reject all Zee AI changes
                    </MenuItem>
                  </MenuList>
                </MenuPopover>
              </Menu>
            )}
            <Tooltip content={dark ? "Switch to light mode" : "Switch to dark mode"} relationship="label">
              <Button appearance="subtle" size="small" icon={dark ? <WeatherSunny20Regular /> : <WeatherMoon20Regular />} onClick={() => updateSettings({ theme: dark ? "light" : "dark" })} />
            </Tooltip>
            <Menu>
              <MenuTrigger disableButtonEnhancement>
                <Tooltip content="Export chat" relationship="label">
                  <Button appearance="subtle" size="small" icon={<ArrowDownload20Regular />} disabled={!messages.length} />
                </Tooltip>
              </MenuTrigger>
              <MenuPopover>
                <MenuList>
                  <MenuItem icon={<Copy20Regular />} onClick={exportCopy}>
                    Copy chat as text
                  </MenuItem>
                  <MenuItem icon={<DocumentAdd20Regular />} onClick={exportToDoc} disabled={!inWord}>
                    Append chat to the end of the document
                  </MenuItem>
                </MenuList>
              </MenuPopover>
            </Menu>
            <Tooltip content="New chat (clears this document's chat history)" relationship="label">
              <Button appearance="subtle" size="small" icon={<Add20Regular />} onClick={clear} disabled={!messages.length} />
            </Tooltip>
            <Tooltip content="Settings" relationship="label">
              <Button appearance="subtle" size="small" icon={<Settings20Regular />} onClick={() => setSettingsOpen(true)} />
            </Tooltip>
          </header>

          <UpdateBanners newerVersion={updates.newerVersion} notes={updates.notes} onDismissNotes={updates.dismissNotes} />
          {!inWord && (
            <MessageBar intent="warning" className="zee-banner">
              <MessageBarBody>Preview mode – open Zee AI from Word to read and edit documents. Chat still works.</MessageBarBody>
            </MessageBar>
          )}
          {!isConfigured(settings) && (
            <MessageBar intent="info" className="zee-banner">
              <MessageBarBody>Connect a Gemini API key, Google account or team proxy to start.</MessageBarBody>
              <MessageBarActions>
                <Button size="small" onClick={() => setSettingsOpen(true)}>
                  Set up
                </Button>
              </MessageBarActions>
            </MessageBar>
          )}

          <QuickActions
            disabled={working}
            useCounts={useCounts}
            preferredLanguage={settings.translateTo}
            onRewrite={onRewrite}
            onGrammar={onGrammar}
            onTone={onTone}
            onSummarize={onSummarize}
            onLength={onLength}
            onTranslate={onTranslate}
            onBilingual={onBilingual}
            onDraft={() => setDraftOpen(true)}
            customActions={allCustomActions(settings)}
            onTable={onTable}
            onConsistency={onConsistency}
            onReference={() => (ensureConfigured() ? setReferenceOpen(true) : undefined)}
            onPicture={onPicture}
            onTemplate={setTemplatePick}
            onCustom={onCustom}
            onManageActions={() => {
              setSettingsFocus("actions");
              setSettingsOpen(true);
            }}
          />

          <ChatFeed messages={messages} onSuggestion={sendChat} />

          <ChatInput
            busy={working}
            status={locating}
            docContext={settings.docContext}
            onDocContext={(v) => updateSettings({ docContext: v })}
            selectionLabel={inWord ? (selectionWords ? `${selectionWords.toLocaleString()} words selected` : null) : "preview"}
            ignoreSelection={ignoreSelection}
            onIgnoreSelection={setIgnoreSelection}
            today={today}
            prices={prices}
            inject={inject}
            onSend={sendChat}
            onStop={stopAll}
          />

          {settingsOpen && (
            <SettingsDrawer
              open={settingsOpen}
              settings={settings}
              usage={usage}
              undoDepth={undoDepth}
              onClearUndo={() => clearUndoHistory().then(() => toast("success", "Undo history cleared", "All inserted text stays in the document."))}
              focus={settingsFocus}
              onImport={(patch) => {
                setSettings((prev) => {
                  const next = sanitizeSettings({ ...prev, ...patch });
                  saveSettings(next);
                  return next;
                });
                toast("success", "Settings imported", "Check Connection – the API key is only included when it was exported with it.");
              }}
              onChange={updateSettings}
              onClose={() => {
                setSettingsOpen(false);
                setSettingsFocus(null);
              }}
            />
          )}
          <DraftDialog open={draftOpen} hasSelection={selectionWords > 0} onClose={() => setDraftOpen(false)} onRun={onDraft} />
          <ReferenceDialog open={referenceOpen} onClose={() => setReferenceOpen(false)} onRun={onReference} />
          <TemplateDialog template={templatePick} hasSelection={selectionWords > 0} onClose={() => setTemplatePick(null)} onRun={onTemplateRun} />
          <ConfirmDialog state={confirmState} onResult={onConfirmResult} />
          <Toaster toasterId={toasterId} position="top-end" />
        </InsertActionsContext.Provider>
      </div>
    </FluentProvider>
  );
}
