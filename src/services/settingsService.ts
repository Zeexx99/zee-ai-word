/**
 * Settings persistence.
 *
 * Office.context.roamingSettings exists only in Outlook add-ins. In Word the alternatives are
 * Office.context.document.settings (saved INSIDE the .docx – the wrong place for an API key) or
 * browser storage. Zee AI for Word therefore keeps settings in localStorage of the add-in's own
 * origin: per user, per machine, and never written into the document or into source code.
 */

import type { TeamConfig } from "../utils/teamConfig";

export type AuthMode = "apiKey" | "google" | "proxy";

/** "auto" follows the Office theme (or the OS colour scheme outside Office). */
export type ThemeMode = "auto" | "light" | "dark";

/**
 * How much of the document the chat sends with each message.
 * "general" = Zee AI first finds the part of the document the request is about, then edits it,
 * adds new text there, or answers from those paragraphs.
 */
export type DocContext = "none" | "selection" | "document" | "general";

/** A user-defined quick action (shown under "My actions"). */
export interface CustomAction {
  id: string;
  name: string;
  /** Instruction sent to Gemini together with the selected text. */
  instruction: string;
  /** replace = edit the selection (diff preview); insert = write new text after it. */
  mode: "replace" | "insert";
}

export const DEFAULT_PROTECTED_TERMS =
  "ESP, VSD, TDH, PIP, BHP, FBHP, SBHP, PI, IPR, GOR, GLR, WC, BS&W, API gravity, bbl/d, BOPD, BWPD, STB/d, Mscf/d, MMscf/d, psi, psig, psia, ft, Hz, kVA, NPSH, Vogel, Arps, ESP run life, MTBF, P&ID, HSE";

export const DEFAULT_CUSTOM_ACTIONS: CustomAction[] = [
  {
    id: "well-test-summary",
    name: "Well test summary",
    instruction:
      "Turn the selected notes or numbers into a concise well test summary paragraph for a daily report: rates, water cut, pressures and any anomaly. Keep every number and unit exactly.",
    mode: "insert",
  },
  {
    id: "client-ready",
    name: "Client-ready tone",
    instruction: "Rewrite the selected text so it is clear, polite and suitable to send to a client. Remove internal jargon and blunt wording; keep all facts and numbers.",
    mode: "replace",
  },
];

export interface ModelInfo {
  id: string;
  label: string;
  hint: string;
}

/**
 * Default models (checked against Google's model list in Oct 2026).
 * Use "Test connection" / ⟳ in Settings to pull the live list for your key.
 */
export const DEFAULT_MODELS: ModelInfo[] = [
  { id: "gemini-3.8-flash", label: "Gemini 3.8 Flash", hint: "Recommended default – fast, smart drafting and editing" },
  { id: "gemini-3.1-pro-preview", label: "Gemini 3.1 Pro (preview)", hint: "Deep reasoning and long-form document drafting" },
  { id: "gemini-3.5-flash-lite", label: "Gemini 3.5 Flash-Lite", hint: "Fastest and cheapest – bulk summaries and quick checks" },
];

export interface ZeeSettings {
  authMode: AuthMode;
  apiKey: string;
  model: string;
  /** Extra model IDs fetched from the API with "Refresh" / "Test connection". */
  fetchedModels: string[];
  googleClientId: string;
  gcpProjectId: string;
  gcpLocation: string;
  /** Team proxy (keeps the Gemini key on a server). */
  proxyUrl: string;
  proxyToken: string;
  /** null = let the model use its default (recommended for Gemini 3). */
  temperature: number | null;
  /** What the chat reads from the document by default. */
  docContext: DocContext;
  /** Documents longer than this many characters are truncated before sending. */
  maxContextChars: number;
  /** Ask before sending prompts larger than this many (estimated) tokens. 0 = never ask. */
  costGuardTokens: number;
  /** Insert AI text with Markdown formatting (headings, bold, lists, tables) instead of plain text. */
  richInsert: boolean;
  /** Default language offered by Translate. */
  translateTo: string;
  /** Appearance of the task pane. */
  theme: ThemeMode;
  /** Record Zee AI's edits as Word revisions (Track Changes) so they can be accepted or rejected. */
  trackChanges: boolean;
  /** Edits change only the words that differ, keeping bold, links, styles and comments. */
  preserveFormatting: boolean;
  /** Inserted text: Arabic right-to-left + right-aligned, English left-aligned, headings centred. */
  autoDirection: boolean;
  /** Font for Arabic text Zee AI writes ("" = same as the document). */
  arabicFont: string;
  /** Terms Zee AI must never rewrite, translate or "correct" (comma or newline separated). */
  protectedTerms: string;
  /** User-defined quick actions. */
  customActions: CustomAction[];
  /** Whole-document passes (proofreading, long summaries, consistency) use the cheaper bulk model. */
  useBulkModel: boolean;
  bulkModel: string;
  /** Shared team glossary + actions (imported file, URL or team proxy). */
  team: TeamConfig | null;
  /** Where the team config comes from: "", "file", an https URL, or "proxy". */
  teamSource: string;
  teamLoadedAt: number;
  /** Optional prices (USD per 1M tokens) for the usage meter's cost estimate. 0 = unknown. */
  priceInPerM: number;
  priceOutPerM: number;
}

export const DEFAULT_SETTINGS: ZeeSettings = {
  authMode: "apiKey",
  apiKey: "",
  model: DEFAULT_MODELS[0].id,
  fetchedModels: [],
  googleClientId: "",
  gcpProjectId: "",
  gcpLocation: "global",
  proxyUrl: "",
  proxyToken: "",
  temperature: null,
  docContext: "selection",
  maxContextChars: 60000,
  costGuardTokens: 20000,
  richInsert: true,
  translateTo: "Arabic",
  theme: "auto",
  trackChanges: false,
  preserveFormatting: true,
  autoDirection: true,
  arabicFont: "",
  protectedTerms: DEFAULT_PROTECTED_TERMS,
  customActions: DEFAULT_CUSTOM_ACTIONS,
  useBulkModel: true,
  bulkModel: "gemini-3.5-flash-lite",
  team: null,
  teamSource: "",
  teamLoadedAt: 0,
  priceInPerM: 0,
  priceOutPerM: 0,
};

export const STORAGE_KEY = "zeeai.word.settings.v1";

export function sanitizeSettings(parsed: Partial<ZeeSettings> | null | undefined): ZeeSettings {
  const merged: ZeeSettings = { ...DEFAULT_SETTINGS, ...(parsed ?? {}) };
  if (!["auto", "light", "dark"].includes(merged.theme)) merged.theme = "auto";
  if (!["apiKey", "google", "proxy"].includes(merged.authMode)) merged.authMode = "apiKey";
  if (!["none", "selection", "document", "general"].includes(merged.docContext)) merged.docContext = "selection";
  if (!Array.isArray(merged.fetchedModels)) merged.fetchedModels = [];
  if (typeof merged.model !== "string" || !merged.model.trim()) merged.model = DEFAULT_SETTINGS.model;
  if (!Number.isFinite(merged.costGuardTokens) || merged.costGuardTokens < 0) merged.costGuardTokens = DEFAULT_SETTINGS.costGuardTokens;
  if (!Number.isFinite(merged.maxContextChars) || merged.maxContextChars < 1000) merged.maxContextChars = DEFAULT_SETTINGS.maxContextChars;
  if (merged.temperature !== null && !Number.isFinite(merged.temperature)) merged.temperature = null;
  if (typeof merged.protectedTerms !== "string") merged.protectedTerms = DEFAULT_PROTECTED_TERMS;
  merged.trackChanges = merged.trackChanges === true;
  merged.preserveFormatting = merged.preserveFormatting !== false;
  merged.autoDirection = merged.autoDirection !== false;
  merged.arabicFont = typeof merged.arabicFont === "string" ? merged.arabicFont.trim().slice(0, 64) : "";
  merged.customActions = Array.isArray(merged.customActions)
    ? merged.customActions
        .filter((a) => a && typeof a.name === "string" && typeof a.instruction === "string" && a.name.trim() && a.instruction.trim())
        .map((a) => ({ id: String(a.id || a.name), name: a.name.trim(), instruction: a.instruction.trim(), mode: (a.mode === "insert" ? "insert" : "replace") as CustomAction["mode"] }))
    : DEFAULT_CUSTOM_ACTIONS;
  merged.useBulkModel = merged.useBulkModel !== false;
  if (typeof merged.bulkModel !== "string" || !merged.bulkModel.trim()) merged.bulkModel = DEFAULT_SETTINGS.bulkModel;
  if (merged.team && (typeof merged.team !== "object" || merged.team.zeeTeamConfig !== 1)) merged.team = null;
  if (typeof merged.teamSource !== "string") merged.teamSource = "";
  if (!Number.isFinite(merged.priceInPerM) || merged.priceInPerM < 0) merged.priceInPerM = 0;
  if (!Number.isFinite(merged.priceOutPerM) || merged.priceOutPerM < 0) merged.priceOutPerM = 0;
  return merged;
}

export function loadSettings(): ZeeSettings {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    return sanitizeSettings(raw ? (JSON.parse(raw) as Partial<ZeeSettings>) : null);
  } catch {
    return { ...DEFAULT_SETTINGS };
  }
}

export function saveSettings(s: ZeeSettings): void {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(s));
  } catch {
    // Storage can be unavailable (private mode / policy). Settings then live for this session only.
  }
}

export function modelLabel(id: string): string {
  return DEFAULT_MODELS.find((m) => m.id === id)?.label ?? id;
}

export function modelHint(id: string): string | undefined {
  return DEFAULT_MODELS.find((m) => m.id === id)?.hint;
}

export function isConfigured(s: ZeeSettings): boolean {
  if (s.authMode === "apiKey") return s.apiKey.trim().length > 0;
  if (s.authMode === "proxy") return s.proxyUrl.trim().length > 0;
  return s.googleClientId.trim().length > 0 && s.gcpProjectId.trim().length > 0;
}

/** Protected terms (personal + team) as a clean list. */
export function protectedTermList(s: Pick<ZeeSettings, "protectedTerms"> & { team?: TeamConfig | null }): string[] {
  const all = `${s.protectedTerms}\n${s.team?.protectedTerms ?? ""}`;
  return Array.from(new Set(all.split(/[,;\n]+/).map((t) => t.trim()).filter(Boolean))).slice(0, 300);
}

/** Personal actions followed by the team's. */
export function allCustomActions(s: Pick<ZeeSettings, "customActions" | "team">): CustomAction[] {
  return [...s.customActions, ...(s.team?.customActions ?? [])];
}

/** Settings for whole-document passes: the cheaper bulk model when enabled. */
export function bulkSettings(s: ZeeSettings): ZeeSettings {
  return s.useBulkModel && s.bulkModel.trim() ? { ...s, model: s.bulkModel.trim() } : s;
}
