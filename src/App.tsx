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
  Copy20Regular,
  DocumentAdd20Regular,
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
import { QuickActions } from "./components/QuickActions";
// Imported eagerly (never lazy-loaded): a separate chunk breaks when Office holds a cached page from an older deploy.
import { SettingsDrawer } from "./components/SettingsDrawer";
import { ConfirmDialog, DraftDialog, type ConfirmState, type DraftChoice } from "./components/Dialogs";
import { InsertActionsContext, copyText, type InsertActions } from "./components/insertActions";

import { isConfigured, loadSettings, modelLabel, saveSettings, type CustomAction, type DocContext, type ZeeSettings } from "./services/settingsService";
import {
  LOCATION_LABELS,
  addComments,
  cleanupOrphanMarkers,
  clearUndoHistory,
  getDocumentParagraphs,
  getSelectionInfo,
  insertContent,
  isWordAvailable,
  onSelectionChange,
  onUndoChange,
  selectParagraph,
  supportsWordApi,
  undoLast,
  type InsertFormat,
  type InsertLocation,
} from "./services/wordService";
import {
  DRAFT_SYSTEM_PROMPT,
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
  type SummaryStyle,
  type ToneId,
} from "./services/prompts";
import { generateJson, generateText, type Usage } from "./services/geminiService";
import { getDocTitle } from "./services/chatStore";
import { countWords, preserveEdges, shorten, truncateMiddle } from "./utils/text";
import { looksLikeMarkdown, markdownToPlain } from "./utils/mdToHtml";
import { chunkParagraphs, numberParagraphs } from "./utils/chunks";

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
  const inWord = useMemo(() => isWordAvailable(), []);
  const docTitle = useMemo(() => getDocTitle(), []);

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
  const addUsage = useCallback(
    (u: Usage) => setUsage((p) => ({ requests: p.requests + 1, promptTokens: p.promptTokens + u.promptTokens, outputTokens: p.outputTokens + u.outputTokens })),
    []
  );

  /* ------------------------------------------------------------ selection tracking */
  const [selectionWords, setSelectionWords] = useState(0);
  useEffect(() => {
    if (!inWord) return;
    const refresh = () =>
      getSelectionInfo()
        .then((s) => setSelectionWords(s.isEmpty ? 0 : countWords(s.text)))
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
  // Markers left by an earlier session can't be undone any more – unwrap them (text stays).
  useEffect(() => {
    if (inWord) cleanupOrphanMarkers();
  }, [inWord]);
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

  const { messages, busy, send, runAction, runGrammar, runTask, markApplied, markCommented, stop, clear } = useChat({ getSettings, onUsage: addUsage, confirmLarge });

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

  const sendChat = useCallback(
    async (text: string) => {
      if (!ensureConfigured()) return;
      const ctx = await docContext(settingsRef.current.docContext);
      await send({ display: text, prompt: text + ctx.suffix, contextLabel: ctx.label, citations: ctx.citations });
    },
    [docContext, ensureConfigured, send]
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
  const onTone = (tone: ToneId) => editAction(`Tone: ${TONES.find((t) => t.id === tone)?.label ?? tone}`, (t) => tonePrompt(t, tone));
  const onLength = (dir: "expand" | "shorten") => editAction(dir === "expand" ? "Expand" : "Shorten", (t) => lengthPrompt(t, dir));
  const onTranslate = (language: string) => {
    if (language !== settingsRef.current.translateTo) updateSettings({ translateTo: language });
    return editAction(`Translate to ${language}`, (t) => translatePrompt(t, language));
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
      assistant: { kind: "action", title: "Grammar & spelling", scope: "document", format: "plain" },
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

  /* ------------------------------------------------------------ my quick actions */
  const [settingsFocus, setSettingsFocus] = useState<"actions" | null>(null);
  const onCustom = async (a: CustomAction) => {
    if (!ensureConfigured()) return;
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

  const comment = useCallback(
    async (m: ChatMessage) => {
      const issues = m.issues ?? [];
      if (!issues.length) return;
      try {
        const r = await addComments(
          issues.map((i) => ({
            find: i.original,
            text: `Zee AI · ${i.type || "suggestion"}: “${i.original}” → “${i.fix}”${i.reason ? ` – ${i.reason}` : ""}`,
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

  const insertActions = useMemo<InsertActions>(() => ({ insert, comment, goToParagraph, canInsert: inWord }), [insert, comment, goToParagraph, inWord]);

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
  const sessionTokens = usage.promptTokens + usage.outputTokens;
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
            disabled={busy}
            preferredLanguage={settings.translateTo}
            onRewrite={onRewrite}
            onGrammar={onGrammar}
            onTone={onTone}
            onSummarize={onSummarize}
            onLength={onLength}
            onTranslate={onTranslate}
            onDraft={() => setDraftOpen(true)}
            customActions={settings.customActions}
            onCustom={onCustom}
            onManageActions={() => {
              setSettingsFocus("actions");
              setSettingsOpen(true);
            }}
          />

          <ChatFeed messages={messages} onSuggestion={sendChat} />

          <ChatInput
            busy={busy}
            docContext={settings.docContext}
            onDocContext={(v) => updateSettings({ docContext: v })}
            selectionLabel={inWord ? (selectionWords ? `${selectionWords.toLocaleString()} words selected` : null) : "preview"}
            sessionTokens={sessionTokens}
            onSend={sendChat}
            onStop={stop}
          />

          {settingsOpen && (
            <SettingsDrawer
              open={settingsOpen}
              settings={settings}
              usage={usage}
              undoDepth={undoDepth}
              onClearUndo={() => clearUndoHistory().then(() => toast("success", "Undo history cleared", "All inserted text stays in the document."))}
              focus={settingsFocus}
              onChange={updateSettings}
              onClose={() => {
                setSettingsOpen(false);
                setSettingsFocus(null);
              }}
            />
          )}
          <DraftDialog open={draftOpen} hasSelection={selectionWords > 0} onClose={() => setDraftOpen(false)} onRun={onDraft} />
          <ConfirmDialog state={confirmState} onResult={onConfirmResult} />
          <Toaster toasterId={toasterId} position="top-end" />
        </InsertActionsContext.Provider>
      </div>
    </FluentProvider>
  );
}
