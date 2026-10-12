import { useCallback, useEffect, useRef, useState } from "react";
import { estimateTokens, GeminiError, generateJson, streamChat, type ChatTurn, type Usage } from "../services/geminiService";
import { buildHistory } from "../utils/history";
import { protectedTermList, type ZeeSettings } from "../services/settingsService";
import { GRAMMAR_SCHEMA, GRAMMAR_SYSTEM_PROMPT, SYSTEM_PROMPT, grammarPrompt, withGlossary, type GrammarIssue, type GrammarResult } from "../services/prompts";
import { clearChat, getDocId, loadChat, saveChat } from "../services/chatStore";
import type { InsertFormat, InsertLocation } from "../services/wordService";
import { cleanModelOutput } from "../utils/text";

export interface ChatMessage {
  id: string;
  role: "user" | "assistant";
  /** What the user sees (assistant: the generated text). */
  text: string;
  /** What was actually sent (display text + document context). */
  promptText?: string;
  /** e.g. "Selection · 120 words" – shown as a chip on user messages. */
  contextLabel?: string;
  status: "streaming" | "done" | "error";
  /** chat = free conversation, action = edit of the selection, draft = new content. */
  kind?: "chat" | "action" | "draft";
  /** Action title, e.g. "Rewrite & polish". */
  title?: string;
  /** Selected text the action worked on – enables the diff preview and "Replace selection". */
  original?: string;
  /** How the result is inserted into Word. */
  format?: InsertFormat;
  /** Insert button offered first. */
  defaultLocation?: InsertLocation;
  /** Grammar check findings. */
  issues?: GrammarIssue[];
  /** "document" = result covers the whole document (findings can become Word comments). */
  scope?: "selection" | "document";
  /** Answer cites [P#] paragraphs that can be clicked to jump there. */
  citations?: boolean;
  /** Set after the findings were added as Word comments. */
  commented?: boolean;
  /** Set after the user inserted this result. */
  applied?: InsertLocation;
  model?: string;
  usage?: Usage;
}

export interface SendInput {
  display: string;
  prompt: string;
  contextLabel?: string;
}

export interface ActionInput extends SendInput {
  kind: "action" | "draft";
  title: string;
  system: string;
  original?: string;
  format: InsertFormat;
  defaultLocation: InsertLocation;
  /** The prompt contains [P#]-numbered paragraphs; make citations clickable. */
  citations?: boolean;
}

/** A multi-step job (e.g. chunked whole-document work) that reports progress into one message. */
export interface TaskInput extends SendInput {
  assistant: Partial<ChatMessage>;
  /** Text used only for the cost-guard estimate (e.g. the whole document). */
  estimateText: string;
  work: (ctx: {
    settings: ZeeSettings;
    signal: AbortSignal;
    /** System prompt with the glossary applied. */
    system: (base: string) => string;
    progress: (text: string) => void;
    addUsage: (u: Usage | undefined) => void;
  }) => Promise<Partial<ChatMessage>>;
}

export interface ChatHookOptions {
  getSettings: () => ZeeSettings;
  onUsage: (u: Usage) => void;
  /** Ask the user before sending a large prompt. Resolve true to continue. */
  confirmLarge: (estimatedTokens: number) => Promise<boolean>;
}

const uid = () => Math.random().toString(36).slice(2, 10) + Date.now().toString(36);

export function useChat({ getSettings, onUsage, confirmLarge }: ChatHookOptions) {
  const docIdRef = useRef<string>("preview");
  const [messages, setMessages] = useState<ChatMessage[]>([]);
  const [busy, setBusy] = useState(false);
  const messagesRef = useRef<ChatMessage[]>([]);
  const abortRef = useRef<AbortController | null>(null);
  const saveTimer = useRef(0);

  // Persist history, debounced: streaming patches the message many times per second and
  // serialising the whole chat into localStorage on every frame makes long chats sluggish.
  const scheduleSave = useCallback((next: ChatMessage[], immediate = false) => {
    window.clearTimeout(saveTimer.current);
    if (immediate) saveChat(docIdRef.current, next);
    else saveTimer.current = window.setTimeout(() => saveChat(docIdRef.current, messagesRef.current), 400);
  }, []);
  useEffect(() => () => window.clearTimeout(saveTimer.current), []);

  // Load this document's history once.
  useEffect(() => {
    docIdRef.current = getDocId();
    const saved = loadChat<ChatMessage>(docIdRef.current)
      .filter((m) => m && typeof m.text === "string" && (m.role === "user" || m.role === "assistant"))
      .map((m) => (m.status === "streaming" ? { ...m, status: "done" as const } : m));
    messagesRef.current = saved;
    setMessages(saved);
  }, []);

  const update = useCallback(
    (fn: (m: ChatMessage[]) => ChatMessage[]) => {
      setMessages((prev) => {
        const next = fn(prev);
        messagesRef.current = next;
        scheduleSave(next, !next.some((m) => m.status === "streaming"));
        return next;
      });
    },
    [scheduleSave]
  );

  const patch = useCallback(
    (id: string, p: Partial<ChatMessage> | ((m: ChatMessage) => Partial<ChatMessage>)) => {
      update((list) => list.map((x) => (x.id === id ? { ...x, ...(typeof p === "function" ? p(x) : p) } : x)));
    },
    [update]
  );

  /** Shared start: cost guard, add messages, controller. Returns null if cancelled or busy. */
  const begin = useCallback(
    async (input: SendInput, system: string, assistant: Partial<ChatMessage>, withHistory: boolean) => {
      if (abortRef.current) return null;
      const settings = getSettings();
      const history: ChatTurn[] = withHistory ? buildHistory(messagesRef.current, input.prompt) : [{ role: "user", text: input.prompt }];
      const tokens = estimateTokens(history.map((h) => h.text).join("\n") + system);
      if (settings.costGuardTokens > 0 && tokens > settings.costGuardTokens) {
        if (!(await confirmLarge(tokens))) return null;
      }
      const userMsg: ChatMessage = { id: uid(), role: "user", text: input.display, promptText: input.prompt, contextLabel: input.contextLabel, status: "done" };
      const asstId = uid();
      update((m) => [...m, userMsg, { id: asstId, role: "assistant", text: "", status: "streaming", model: settings.model, kind: "chat", ...assistant }]);
      const controller = new AbortController();
      abortRef.current = controller;
      setBusy(true);
      return { settings, history, asstId, controller, system: withGlossary(system, protectedTermList(settings)) };
    },
    [confirmLarge, getSettings, update]
  );

  const finish = useCallback(() => {
    abortRef.current = null;
    setBusy(false);
  }, []);

  const fail = useCallback(
    (asstId: string, e: unknown) => {
      const err =
        e instanceof GeminiError ? e : (e as Error)?.name === "AbortError" ? new GeminiError("Stopped.", "aborted") : new GeminiError((e as Error)?.message ?? String(e));
      patch(asstId, (x) => {
        if (err.kind === "aborted") return { text: x.text ? `${x.text}\n\n_…stopped._` : "_Stopped._", status: "done", original: undefined, defaultLocation: x.text ? x.defaultLocation : undefined };
        return { text: err.message, status: "error" };
      });
    },
    [patch]
  );

  /** Streams an answer into a new assistant message. */
  const stream = useCallback(
    async (input: SendInput, system: string, assistant: Partial<ChatMessage>, withHistory: boolean, clean: boolean) => {
      const ctx = await begin(input, system, assistant, withHistory);
      if (!ctx) return;
      const { settings, history, asstId, controller, system: fullSystem } = ctx;
      let pending: string | null = null;
      let frame = 0;
      const flush = () => {
        frame = 0;
        if (pending === null) return;
        const text = pending;
        pending = null;
        patch(asstId, { text });
      };
      try {
        const result = await streamChat({
          settings,
          history,
          systemInstruction: fullSystem,
          signal: controller.signal,
          onText: (full) => {
            pending = full;
            if (!frame) frame = requestAnimationFrame(flush);
          },
        });
        if (frame) cancelAnimationFrame(frame);
        const text = clean ? cleanModelOutput(result.text, assistant.original ?? "") : result.text;
        patch(asstId, { text: text || "_(empty response)_", status: "done", usage: result.usage });
        if (result.usage) onUsage(result.usage);
      } catch (e) {
        if (frame) cancelAnimationFrame(frame);
        fail(asstId, e);
      } finally {
        finish();
      }
    },
    [begin, fail, finish, onUsage, patch]
  );

  /** Free chat (document-aware through the prompt the caller builds). */
  const send = useCallback(
    (input: SendInput & { citations?: boolean }) =>
      stream(input, SYSTEM_PROMPT, { kind: "chat", format: "markdown", defaultLocation: "after", citations: input.citations }, true, false),
    [stream]
  );

  /** Selection actions and drafting: no chat history (cheaper, more predictable). */
  const runAction = useCallback(
    (input: ActionInput) =>
      stream(
        input,
        input.system,
        { kind: input.kind, title: input.title, original: input.original, format: input.format, defaultLocation: input.defaultLocation, citations: input.citations },
        false,
        true
      ),
    [stream]
  );

  /** Grammar & spelling check (JSON mode: corrected text + list of issues). */
  const runGrammar = useCallback(
    async (input: { display: string; original: string; contextLabel?: string }) => {
      const prompt = grammarPrompt(input.original);
      const ctx = await begin({ display: input.display, prompt, contextLabel: input.contextLabel }, GRAMMAR_SYSTEM_PROMPT, {
        kind: "action",
        title: "Grammar & spelling",
        original: input.original,
        format: "plain",
        defaultLocation: "replace",
      }, false);
      if (!ctx) return;
      const { settings, asstId, controller, system } = ctx;
      try {
        const { data, usage } = await generateJson<GrammarResult>({
          settings,
          prompt,
          systemInstruction: system,
          schema: GRAMMAR_SCHEMA,
          signal: controller.signal,
        });
        const issues = Array.isArray(data?.issues) ? data.issues.filter((i) => i && typeof i.original === "string") : [];
        const corrected = typeof data?.corrected === "string" && data.corrected.trim() ? data.corrected : input.original;
        patch(asstId, { text: corrected, issues, status: "done", usage });
        if (usage) onUsage(usage);
      } catch (e) {
        fail(asstId, e);
      } finally {
        finish();
      }
    },
    [begin, fail, finish, onUsage, patch]
  );

  /** Multi-step job with progress (whole-document summaries and proofreading). */
  const runTask = useCallback(
    async (input: TaskInput) => {
      const ctx = await begin({ display: input.display, prompt: input.estimateText, contextLabel: input.contextLabel }, "", input.assistant, false);
      if (!ctx) return;
      const { settings, asstId, controller } = ctx;
      const total: Usage = { promptTokens: 0, outputTokens: 0, totalTokens: 0 };
      // The user turn keeps a short prompt (not the whole document) in the saved history.
      patch(asstId, { text: "" });
      update((list) => list.map((m, i) => (list[i + 1]?.id === asstId ? { ...m, promptText: input.prompt } : m)));
      try {
        const result = await input.work({
          settings,
          signal: controller.signal,
          system: (base) => withGlossary(base, protectedTermList(settings)),
          progress: (text) => patch(asstId, { title: `${input.assistant.title ?? "Working"} – ${text}` }),
          addUsage: (u) => {
            if (!u) return;
            total.promptTokens += u.promptTokens;
            total.outputTokens += u.outputTokens;
            total.totalTokens += u.totalTokens;
            onUsage(u);
          },
        });
        patch(asstId, { ...result, title: input.assistant.title, status: "done", usage: total });
      } catch (e) {
        patch(asstId, { title: input.assistant.title });
        fail(asstId, e);
      } finally {
        finish();
      }
    },
    [begin, fail, finish, onUsage, patch, update]
  );

  const markCommented = useCallback((id: string) => patch(id, { commented: true }), [patch]);

  const markApplied = useCallback((id: string, location: InsertLocation) => patch(id, { applied: location }), [patch]);
  const stop = useCallback(() => abortRef.current?.abort(), []);
  const clear = useCallback(() => {
    abortRef.current?.abort();
    clearChat(docIdRef.current);
    update(() => []);
  }, [update]);

  return { messages, busy, send, runAction, runGrammar, runTask, markApplied, markCommented, stop, clear };
}
