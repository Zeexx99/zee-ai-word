/**
 * SHARED MODULE – keep host-agnostic. Synced to the other Zee AI add-in with scripts/sync-shared.mjs.
 *
 * Gemini client (direct REST, no SDK, no backend) supporting three auth paths:
 *  - API key        -> Gemini Developer API (generativelanguage.googleapis.com, key in the x-goog-api-key header)
 *  - Google sign-in -> Vertex AI            (aiplatform.googleapis.com, user's GCP project quota, Bearer token)
 *  - Team proxy     -> your proxy URL       (Bearer team token; the proxy holds the key)
 * Streaming uses `streamGenerateContent?alt=sse` and is parsed incrementally.
 */
import { getValidGoogleToken } from "./auth";
import { SseParser } from "./sse";
import { hashString } from "./hash";

/**
 * Connection + generation settings the client needs. Both Zee AI add-ins' settings objects
 * satisfy this shape structurally, so this file is host-agnostic (no Word / Excel imports).
 */
export interface GeminiConnection {
  authMode: "apiKey" | "google" | "proxy";
  apiKey: string;
  model: string;
  gcpProjectId: string;
  gcpLocation: string;
  proxyUrl: string;
  proxyToken: string;
  temperature: number | null;
}
type ZeeSettings = GeminiConnection;

const SYSTEM_PROMPT_FALLBACK = "You are Zee AI, a helpful assistant inside Microsoft Office.";

export interface ChatTurn {
  role: "user" | "model";
  text: string;
}

export interface Usage {
  promptTokens: number;
  outputTokens: number;
  totalTokens: number;
}

export type GeminiErrorKind = "config" | "auth" | "quota" | "notFound" | "safety" | "network" | "server" | "aborted" | "unknown";

export class GeminiError extends Error {
  readonly kind: GeminiErrorKind;
  readonly status?: number;
  constructor(message: string, kind: GeminiErrorKind = "unknown", status?: number) {
    super(message);
    this.name = "GeminiError";
    this.kind = kind;
    this.status = status;
  }
}

interface Endpoint {
  url: (method: string) => string;
  headers: Record<string, string>;
  label: string;
}

const GEMINI_API = "https://generativelanguage.googleapis.com/v1beta";

function resolveEndpoint(s: ZeeSettings, model: string): Endpoint {
  const m = encodeURIComponent(model.trim());
  if (!m) throw new GeminiError("Choose a model in Settings.", "config");
  if (s.authMode === "proxy") {
    const base = s.proxyUrl.trim().replace(/\/+$/, "");
    if (!base) throw new GeminiError("Enter your team proxy URL in Settings.", "config");
    const headers: Record<string, string> = { "Content-Type": "application/json" };
    if (s.proxyToken.trim()) headers.Authorization = `Bearer ${s.proxyToken.trim()}`;
    return { url: (method) => `${base}/v1beta/models/${m}:${method}`, headers, label: "team proxy" };
  }
  if (s.authMode === "apiKey") {
    const key = s.apiKey.trim();
    if (!key) throw new GeminiError("Add your Gemini API key in Settings (gear icon).", "config");
    return {
      url: (method) => `${GEMINI_API}/models/${m}:${method}`,
      headers: { "Content-Type": "application/json", "x-goog-api-key": key },
      label: "Gemini API",
    };
  }
  const token = getValidGoogleToken();
  if (!token) throw new GeminiError("Your Google sign-in has expired. Open Settings and sign in again.", "auth");
  const project = s.gcpProjectId.trim();
  const location = (s.gcpLocation.trim() || "global").toLowerCase();
  if (!project) throw new GeminiError("Enter your Google Cloud project ID in Settings.", "config");
  const host = location === "global" ? "aiplatform.googleapis.com" : `${location}-aiplatform.googleapis.com`;
  return {
    url: (method) =>
      `https://${host}/v1/projects/${encodeURIComponent(project)}/locations/${location}/publishers/google/models/${m}:${method}`,
    headers: { "Content-Type": "application/json", Authorization: `Bearer ${token}`, "x-goog-user-project": project },
    label: `Vertex AI (${project})`,
  };
}

function buildBody(history: ChatTurn[], systemInstruction: string, s: ZeeSettings, extraConfig: Record<string, unknown> = {}) {
  return buildRawBody(
    history.map((t) => ({ role: t.role, parts: [{ text: t.text }] })),
    systemInstruction,
    s,
    extraConfig
  );
}

function buildRawBody(contents: GeminiContent[], systemInstruction: string, s: ZeeSettings, extraConfig: Record<string, unknown> = {}) {
  const generationConfig: Record<string, unknown> = { ...extraConfig };
  if (s.temperature !== null && s.temperature !== undefined) generationConfig.temperature = s.temperature;
  return {
    systemInstruction: { parts: [{ text: systemInstruction }] },
    contents,
    generationConfig,
  };
}

/** Raw Gemini content parts (text, or inline files/images as base64). */
export interface GeminiPart {
  text?: string;
  thought?: boolean;
  thoughtSignature?: string;
  inlineData?: { mimeType: string; data: string };
}
export interface GeminiContent {
  role: "user" | "model";
  parts: GeminiPart[];
}

/** Rough token estimate (~4 characters per token) used by the cost guard. */
export function estimateTokens(text: string): number {
  return Math.ceil(text.length / 4);
}

interface ApiErrorBody {
  error?: { code?: number; message?: string; status?: string; details?: Array<Record<string, unknown>> };
}

function retryDelayMs(body: ApiErrorBody | null): number | null {
  const info = body?.error?.details?.find((d) => String(d["@type"] ?? "").includes("RetryInfo"));
  const delay = info?.retryDelay;
  if (typeof delay === "string") {
    const sec = parseFloat(delay);
    if (Number.isFinite(sec)) return Math.min(sec * 1000, 30_000);
  }
  return null;
}

/** Map an HTTP error to a user-friendly GeminiError. */
export function toGeminiError(status: number, bodyText: string, model: string): GeminiError {
  let body: ApiErrorBody | null = null;
  try {
    body = JSON.parse(bodyText) as ApiErrorBody;
  } catch {
    /* not JSON */
  }
  const apiMsg = body?.error?.message ?? bodyText.slice(0, 300);
  if (status === 400 && /api key/i.test(apiMsg)) return new GeminiError("Your Gemini API key is invalid. Check it in Settings.", "auth", status);
  if (status === 401) return new GeminiError("Authentication failed – sign in again or re-check your key.", "auth", status);
  if (status === 403)
    return new GeminiError(
      `Permission denied: ${apiMsg} (Check that the API is enabled for your project and your account has the "Vertex AI User" role.)`,
      "auth",
      status
    );
  if (status === 404) return new GeminiError(`Model "${model}" isn't available for this account. Pick another model in Settings (or click Refresh).`, "notFound", status);
  if (status === 429)
    return new GeminiError(
      "Rate limit or quota exceeded. Wait a minute, switch to a Flash-Lite model, or check billing/quota in Google AI Studio / Cloud console.",
      "quota",
      status
    );
  if (status >= 500) return new GeminiError(`Gemini is temporarily unavailable (${status}). Please try again.`, "server", status);
  return new GeminiError(`Gemini error ${status}: ${apiMsg}`, "unknown", status);
}

function sleep(ms: number, signal?: AbortSignal): Promise<void> {
  return new Promise((resolve, reject) => {
    if (signal?.aborted) {
      reject(new GeminiError("Stopped.", "aborted"));
      return;
    }
    const onAbort = () => {
      clearTimeout(t);
      reject(new GeminiError("Stopped.", "aborted"));
    };
    const t = setTimeout(() => {
      signal?.removeEventListener("abort", onAbort);
      resolve();
    }, ms);
    signal?.addEventListener("abort", onAbort, { once: true });
  });
}

/** fetch with automatic retry on 429 / 5xx (before any streamed bytes are consumed). */
async function fetchWithRetry(url: string, init: RequestInit, model: string, retries = 2): Promise<Response> {
  for (let attempt = 0; ; attempt++) {
    let res: Response;
    try {
      res = await fetch(url, init);
    } catch (e) {
      if ((e as Error).name === "AbortError") throw new GeminiError("Stopped.", "aborted");
      throw new GeminiError("Network error – check your internet connection or firewall.", "network");
    }
    if (res.ok) return res;
    const text = await res.text();
    const retryable = res.status === 429 || res.status === 500 || res.status === 503;
    if (retryable && attempt < retries) {
      let body: ApiErrorBody | null = null;
      try {
        body = JSON.parse(text) as ApiErrorBody;
      } catch {
        /* ignore */
      }
      const wait = retryDelayMs(body) ?? 1000 * 2 ** attempt + Math.random() * 400;
      await sleep(wait, init.signal ?? undefined);
      continue;
    }
    throw toGeminiError(res.status, text, model);
  }
}

interface Candidate {
  content?: { role?: string; parts?: GeminiPart[] };
  finishReason?: string;
}
interface GenerateResponse {
  candidates?: Candidate[];
  promptFeedback?: { blockReason?: string };
  usageMetadata?: { promptTokenCount?: number; candidatesTokenCount?: number; totalTokenCount?: number; thoughtsTokenCount?: number };
  error?: { message?: string };
}

function textOf(r: GenerateResponse): string {
  return (r.candidates?.[0]?.content?.parts ?? [])
    .filter((p) => !p.thought && typeof p.text === "string")
    .map((p) => p.text)
    .join("");
}

function usageOf(r: GenerateResponse): Usage | undefined {
  const u = r.usageMetadata;
  if (!u) return undefined;
  const out = (u.candidatesTokenCount ?? 0) + (u.thoughtsTokenCount ?? 0);
  return { promptTokens: u.promptTokenCount ?? 0, outputTokens: out, totalTokens: u.totalTokenCount ?? (u.promptTokenCount ?? 0) + out };
}

export interface StreamResult {
  text: string;
  usage?: Usage;
  finishReason?: string;
}

export interface StreamOptions {
  settings: ZeeSettings;
  history: ChatTurn[];
  systemInstruction?: string;
  signal?: AbortSignal;
  onText: (fullText: string) => void;
}

/** Stream a chat completion. Resolves with the full text when the stream ends. */
export async function streamChat(opts: StreamOptions): Promise<StreamResult> {
  const { settings, signal } = opts;
  const model = settings.model;
  const ep = resolveEndpoint(settings, model);
  const res = await fetchWithRetry(
    `${ep.url("streamGenerateContent")}?alt=sse`,
    { method: "POST", headers: ep.headers, body: JSON.stringify(buildBody(opts.history, opts.systemInstruction ?? SYSTEM_PROMPT_FALLBACK, settings)), signal },
    model
  );
  if (!res.body) throw new GeminiError("Streaming is not supported in this environment.", "unknown");

  const reader = res.body.getReader();
  const decoder = new TextDecoder();
  const parser = new SseParser();
  let full = "";
  let usage: Usage | undefined;
  let finishReason: string | undefined;

  const handle = (data: string) => {
    let chunk: GenerateResponse;
    try {
      chunk = JSON.parse(data) as GenerateResponse;
    } catch {
      return;
    }
    if (chunk.error) throw new GeminiError(chunk.error.message ?? "Gemini stream error", "server");
    if (chunk.promptFeedback?.blockReason)
      throw new GeminiError(`The request was blocked by Gemini safety filters (${chunk.promptFeedback.blockReason}).`, "safety");
    const delta = textOf(chunk);
    if (delta) {
      full += delta;
      opts.onText(full);
    }
    usage = usageOf(chunk) ?? usage;
    finishReason = chunk.candidates?.[0]?.finishReason ?? finishReason;
  };

  try {
    for (;;) {
      const { done, value } = await reader.read();
      if (done) break;
      parser.push(decoder.decode(value, { stream: true })).forEach(handle);
    }
    parser.push(decoder.decode()).forEach(handle);
    parser.flush().forEach(handle);
  } catch (e) {
    if ((e as Error).name === "AbortError") throw new GeminiError("Stopped.", "aborted");
    throw e;
  }

  if (!full && finishReason === "SAFETY") throw new GeminiError("The response was blocked by Gemini safety filters.", "safety");
  if (finishReason === "MAX_TOKENS") {
    full += "\n\n_…response truncated (max output tokens reached)._";
    opts.onText(full);
  }
  return { text: full, usage, finishReason };
}

/* ---------------------------------------------------------------- response cache */

/**
 * Small in-memory cache for one-shot requests: re-running the same action on the same text
 * (e.g. clicking Summarize twice, or re-proofreading unchanged parts) costs nothing.
 * Session only – never persisted, never shared.
 */
const CACHE_MAX = 80;
const cache = new Map<string, { text: string; usage?: Usage }>();
function cacheKey(s: ZeeSettings, parts: string[]): string {
  return hashString([s.authMode, s.model, String(s.temperature), ...parts].join("\u0001")) + ":" + parts.reduce((n, p) => n + p.length, 0);
}
function cacheGet(key: string) {
  const hit = cache.get(key);
  if (hit) {
    cache.delete(key);
    cache.set(key, hit); // most recently used
  }
  return hit;
}
function cachePut(key: string, value: { text: string; usage?: Usage }) {
  cache.set(key, value);
  while (cache.size > CACHE_MAX) cache.delete(cache.keys().next().value as string);
}
export function clearResponseCache(): void {
  cache.clear();
}
/** Number of cached answers (shown in Settings). */
export function responseCacheSize(): number {
  return cache.size;
}

/** One-shot JSON generation (used for bulk row transforms). */
export async function generateJson<T>(opts: {
  settings: ZeeSettings;
  prompt: string;
  systemInstruction: string;
  schema: Record<string, unknown>;
  signal?: AbortSignal;
}): Promise<{ data: T; usage?: Usage }> {
  const { settings } = opts;
  const key = cacheKey(settings, ["json", opts.systemInstruction, JSON.stringify(opts.schema), opts.prompt]);
  const hit = cacheGet(key);
  if (hit) return { data: JSON.parse(hit.text) as T };
  const ep = resolveEndpoint(settings, settings.model);
  const body = buildBody([{ role: "user", text: opts.prompt }], opts.systemInstruction, settings, {
    responseMimeType: "application/json",
    responseSchema: opts.schema,
  });
  const res = await fetchWithRetry(
    ep.url("generateContent"),
    { method: "POST", headers: ep.headers, body: JSON.stringify(body), signal: opts.signal },
    settings.model,
    3
  );
  const json = (await res.json()) as GenerateResponse;
  const text = textOf(json);
  try {
    const data = JSON.parse(text) as T;
    cachePut(key, { text });
    return { data, usage: usageOf(json) };
  } catch {
    throw new GeminiError("Gemini returned malformed JSON. Try again or use a Pro model.", "unknown");
  }
}

/** One-shot plain text generation. */
export async function generateText(opts: { settings: ZeeSettings; prompt: string; systemInstruction: string; signal?: AbortSignal }): Promise<{ text: string; usage?: Usage }> {
  const { settings } = opts;
  const key = cacheKey(settings, ["text", opts.systemInstruction, opts.prompt]);
  const hit = cacheGet(key);
  if (hit) return { text: hit.text };
  const ep = resolveEndpoint(settings, settings.model);
  const res = await fetchWithRetry(
    ep.url("generateContent"),
    { method: "POST", headers: ep.headers, body: JSON.stringify(buildBody([{ role: "user", text: opts.prompt }], opts.systemInstruction, settings)), signal: opts.signal },
    settings.model,
    3
  );
  const json = (await res.json()) as GenerateResponse;
  if (json.promptFeedback?.blockReason) throw new GeminiError(`Blocked by safety filters (${json.promptFeedback.blockReason}).`, "safety");
  const text = textOf(json).trim();
  if (text) cachePut(key, { text });
  return { text, usage: usageOf(json) };
}

/**
 * One-shot request with mixed parts – text plus inline files or images (base64), e.g. a PDF
 * reference document or a picture from the document. Optional JSON schema for structured output.
 */
export async function generateFromParts<T = unknown>(opts: {
  settings: ZeeSettings;
  parts: GeminiPart[];
  systemInstruction: string;
  schema?: Record<string, unknown>;
  signal?: AbortSignal;
}): Promise<{ text: string; data?: T; usage?: Usage }> {
  const { settings } = opts;
  const ep = resolveEndpoint(settings, settings.model);
  const extra: Record<string, unknown> = opts.schema ? { responseMimeType: "application/json", responseSchema: opts.schema } : {};
  const body = buildRawBody([{ role: "user", parts: opts.parts }], opts.systemInstruction, settings, extra);
  const res = await fetchWithRetry(ep.url("generateContent"), { method: "POST", headers: ep.headers, body: JSON.stringify(body), signal: opts.signal }, settings.model, 2);
  const json = (await res.json()) as GenerateResponse;
  if (json.promptFeedback?.blockReason) throw new GeminiError(`Blocked by safety filters (${json.promptFeedback.blockReason}).`, "safety");
  const text = textOf(json).trim();
  if (!opts.schema) return { text, usage: usageOf(json) };
  try {
    return { text, data: JSON.parse(text) as T, usage: usageOf(json) };
  } catch {
    throw new GeminiError("Gemini returned malformed JSON. Try again or use a Pro model.", "unknown");
  }
}

interface ModelListResponse {
  models?: Array<{ name: string; supportedGenerationMethods?: string[] }>;
}

/** Keep text-generation Gemini models only, without the "models/" prefix. */
export function filterModelIds(json: ModelListResponse): string[] {
  return (json.models ?? [])
    .filter((m) => m.name.includes("gemini") && (m.supportedGenerationMethods ?? []).includes("generateContent"))
    .map((m) => m.name.replace(/^models\//, ""))
    .filter((id) => !/(tts|embedding|image|live|audio|transcribe)/i.test(id))
    .sort();
}

/** List generateContent-capable Gemini models (API key directly, or through the team proxy). */
export async function listModels(s: ZeeSettings): Promise<string[]> {
  let url: string;
  const headers: Record<string, string> = {};
  if (s.authMode === "proxy") {
    const base = s.proxyUrl.trim().replace(/\/+$/, "");
    if (!base) throw new GeminiError("Enter your team proxy URL first.", "config");
    url = `${base}/v1beta/models?pageSize=1000`;
    if (s.proxyToken.trim()) headers.Authorization = `Bearer ${s.proxyToken.trim()}`;
  } else {
    if (!s.apiKey.trim()) throw new GeminiError("Enter an API key first.", "config");
    url = `${GEMINI_API}/models?pageSize=1000`;
    headers["x-goog-api-key"] = s.apiKey.trim();
  }
  let res: Response;
  try {
    res = await fetch(url, { headers });
  } catch {
    throw new GeminiError("Network error – check your internet connection (or the proxy URL).", "network");
  }
  if (!res.ok) throw toGeminiError(res.status, await res.text(), "");
  return filterModelIds((await res.json()) as ModelListResponse);
}

export interface ConnectionResult {
  message: string;
  /** Models available to these credentials (empty when the backend can't list them). */
  models: string[];
}

/** Validates credentials and model access, and refreshes the model list where possible. */
export async function testConnection(s: ZeeSettings): Promise<ConnectionResult> {
  if (s.authMode === "apiKey") {
    const models = await listModels(s);
    const hasModel = models.includes(s.model);
    return {
      message: `Connected to Gemini API – ${models.length} models available.${hasModel ? "" : ` Note: "${s.model}" is not in your list – pick another model.`}`,
      models,
    };
  }
  // Proxy and Vertex AI: prove the selected model answers (cheap countTokens call).
  const ep = resolveEndpoint(s, s.model);
  const res = await fetchWithRetry(
    ep.url("countTokens"),
    { method: "POST", headers: ep.headers, body: JSON.stringify({ contents: [{ role: "user", parts: [{ text: "ping" }] }] }) },
    s.model,
    0
  );
  await res.json();
  let models: string[] = [];
  if (s.authMode === "proxy") models = await listModels(s).catch(() => []);
  return { message: `Connected to ${ep.label} – model "${s.model}" is reachable.${models.length ? ` ${models.length} models available.` : ""}`, models };
}
