import { useEffect, useMemo, useRef, useState } from "react";
import {
  Button,
  Combobox,
  Divider,
  DrawerBody,
  DrawerHeader,
  DrawerHeaderTitle,
  Dropdown,
  Field,
  Input,
  Link,
  MessageBar,
  MessageBarBody,
  Option,
  OptionGroup,
  OverlayDrawer,
  Radio,
  RadioGroup,
  Spinner,
  Switch,
  Textarea,
} from "@fluentui/react-components";
import { ArrowSync16Regular, Dismiss24Regular, Eye20Regular, EyeOff20Regular, Person16Regular } from "@fluentui/react-icons";
import { DEFAULT_MODELS, DEFAULT_PROTECTED_TERMS, DEFAULT_SETTINGS, modelHint, modelLabel, switchAuthMode, type AuthMode, type DocContext, type ThemeMode, type ZeeSettings } from "../services/settingsService";
import { OTHER_PROVIDERS, providerName, providerOf } from "../shared/openaiCompat";
import { loadPrompts, savePrompts } from "../services/promptStore";
import { BackupSection, UsageSection } from "../qol/BackupSection";
import { getGoogleSession, signInWithGoogle, signOutGoogle, type GoogleSession } from "../services/authService";
import { listModels, testConnection } from "../services/geminiService";
import { LANGUAGES } from "../services/prompts";
import type { UsageTotals } from "../App";
import { openExternal } from "./insertActions";
import { CustomActionsEditor } from "./CustomActionsEditor";
import { supportsWordApi } from "../services/wordService";
import { clearResponseCache, responseCacheSize } from "../services/geminiService";
import { downloadText, exportTeamConfig, fetchTeamConfig } from "../services/teamService";
import { parseTeamConfig } from "../utils/teamConfig";
import { copyText } from "./insertActions";

interface Props {
  open: boolean;
  settings: ZeeSettings;
  usage: UsageTotals;
  undoDepth: number;
  onClearUndo: () => void;
  /** Scroll to this section when the drawer opens. */
  focus?: "actions" | null;
  onChange: (patch: Partial<ZeeSettings>) => void;
  /** Settings from an export file (Backup & move). */
  onImport: (patch: Partial<ZeeSettings>) => void;
  onClose: () => void;
}

type Status = { intent: "success" | "error" | "info"; text: string } | null;

const CONTEXTS: Array<{ value: DocContext; label: string }> = [
  { value: "selection", label: "Selected text (recommended)" },
  { value: "document", label: "Whole document" },
  { value: "general", label: "General – Zee AI finds the relevant part" },
  { value: "none", label: "Nothing – chat only" },
];

const TEMPS: Array<{ value: string; label: string }> = [
  { value: "auto", label: "Auto (model default – recommended)" },
  { value: "0", label: "0 – precise, repeatable" },
  { value: "0.4", label: "0.4 – balanced" },
  { value: "1", label: "1.0 – creative" },
];

function minutesLeft(s: GoogleSession | null): number {
  return s ? Math.max(0, Math.round((s.expiresAt - Date.now()) / 60000)) : 0;
}

export function SettingsDrawer({ open, settings, usage, undoDepth, onClearUndo, focus, onChange, onImport, onClose }: Props) {
  const actionsRef = useRef<HTMLElement>(null);
  useEffect(() => {
    if (open && focus === "actions") window.setTimeout(() => actionsRef.current?.scrollIntoView({ behavior: "smooth", block: "start" }), 250);
  }, [open, focus]);
  const canTrack = supportsWordApi("1.4");
  const [teamUrl, setTeamUrl] = useState(settings.teamSource === "file" ? "" : settings.teamSource);
  const [teamBusy, setTeamBusy] = useState(false);
  const [teamStatus, setTeamStatus] = useState<Status>(null);
  const [cacheSize, setCacheSize] = useState(0);
  useEffect(() => {
    if (open) setCacheSize(responseCacheSize());
  }, [open]);
  const [showKey, setShowKey] = useState(false);
  const [status, setStatus] = useState<Status>(null);
  const [testing, setTesting] = useState(false);
  const [refreshing, setRefreshing] = useState(false);
  const [signingIn, setSigningIn] = useState(false);
  const [session, setSession] = useState<GoogleSession | null>(() => getGoogleSession());
  const [modelText, setModelText] = useState(settings.model);

  useEffect(() => setModelText(settings.model), [settings.model]);
  useEffect(() => {
    if (open) {
      setSession(getGoogleSession());
      setStatus(null);
    }
  }, [open]);

  const other = settings.authMode === "openai";
  const otherName = providerName(settings.otherBaseUrl);
  const otherPreset = providerOf(settings.otherBaseUrl);
  const extraModels = useMemo(
    () => settings.fetchedModels.filter((id) => !DEFAULT_MODELS.some((m) => m.id === id)),
    [settings.fetchedModels]
  );
  const otherModels = useMemo(() => {
    const ids = [...settings.otherFetchedModels];
    if (otherPreset?.model && !ids.includes(otherPreset.model)) ids.unshift(otherPreset.model);
    return ids;
  }, [settings.otherFetchedModels, otherPreset]);
  const signedIn = !!session && minutesLeft(session) > 0;

  const runTest = async () => {
    setTesting(true);
    setStatus(null);
    try {
      const r = await testConnection(settings);
      if (r.models.length) onChange(other ? { otherFetchedModels: r.models } : { fetchedModels: r.models });
      setStatus({ intent: "success", text: r.message });
    } catch (e) {
      setStatus({ intent: "error", text: (e as Error).message });
    } finally {
      setTesting(false);
    }
  };

  const refreshModels = async () => {
    setRefreshing(true);
    try {
      const ids = await listModels(settings);
      onChange(other ? { otherFetchedModels: ids } : { fetchedModels: ids });
      setStatus({ intent: "info", text: `Loaded ${ids.length} models from ${other ? otherName : "your account"}.` });
    } catch (e) {
      setStatus({ intent: "error", text: (e as Error).message });
    } finally {
      setRefreshing(false);
    }
  };

  const signIn = async () => {
    setSigningIn(true);
    setStatus(null);
    try {
      const s = await signInWithGoogle(settings.googleClientId);
      setSession(s);
      setStatus({ intent: "success", text: `Signed in${s.email ? ` as ${s.email}` : ""}.` });
    } catch (e) {
      setStatus({ intent: "error", text: (e as Error).message });
    } finally {
      setSigningIn(false);
    }
  };

  const commitModel = (id: string) => {
    const v = id.trim();
    if (v && v !== settings.model) onChange({ model: v });
  };

  return (
    <OverlayDrawer open={open} position="end" size="full" onOpenChange={(_, d) => !d.open && onClose()}>
      <DrawerHeader>
        <DrawerHeaderTitle action={<Button appearance="subtle" aria-label="Close" icon={<Dismiss24Regular />} onClick={onClose} />}>
          Settings
        </DrawerHeaderTitle>
      </DrawerHeader>
      <DrawerBody className="zee-settings">
        <section>
          <h3>Connection</h3>
          <RadioGroup value={settings.authMode} onChange={(_, d) => onChange(switchAuthMode(settings, d.value as AuthMode))}>
            <Radio value="apiKey" label="Gemini API key (bring your own key)" />
            <Radio value="openai" label="Other API key (Groq, OpenRouter, OpenAI, Mistral…)" />
            <Radio value="google" label="Google sign-in (Vertex AI, your Cloud project quota)" />
            <Radio value="proxy" label="Team proxy (key kept on your server)" />
          </RadioGroup>

          {settings.authMode === "proxy" ? (
            <div className="zee-form">
              <Field label="Proxy URL" hint="Your deployed Zee AI proxy (see proxy/README.md – the Excel add-in's proxy works too), e.g. https://zee-proxy.yourteam.workers.dev">
                <Input value={settings.proxyUrl} onChange={(_, d) => onChange({ proxyUrl: d.value })} placeholder="https://…" />
              </Field>
              <Field label="Team access token" hint="The shared secret your admin configured on the proxy.">
                <Input type="password" value={settings.proxyToken} onChange={(_, d) => onChange({ proxyToken: d.value })} autoComplete="off" />
              </Field>
            </div>
          ) : other ? (
            <div className="zee-form">
              <Field label="Provider" hint="Any service with an OpenAI-compatible API works. Pick Custom to type its base URL.">
                <Dropdown
                  value={otherPreset?.label ?? "Custom"}
                  selectedOptions={[otherPreset?.id ?? "custom"]}
                  onOptionSelect={(_, d) => {
                    const p = OTHER_PROVIDERS.find((x) => x.id === d.optionValue);
                    if (!p) return;
                    onChange({ otherBaseUrl: p.baseUrl, otherFetchedModels: [], ...(p.model ? { model: p.model } : {}) });
                  }}
                >
                  {OTHER_PROVIDERS.map((p) => (
                    <Option key={p.id} value={p.id}>
                      {p.label}
                    </Option>
                  ))}
                  <Option value="custom">Custom</Option>
                </Dropdown>
              </Field>
              <Field label="Base URL" hint="Ends in /v1 for most providers. Requests go to {base URL}/chat/completions.">
                <Input value={settings.otherBaseUrl} onChange={(_, d) => onChange({ otherBaseUrl: d.value })} placeholder="https://api.groq.com/openai/v1" />
              </Field>
              <Field
                label={`${otherName} API key`}
                hint={
                  <>
                    {otherPreset ? (
                      <>
                        Create one at{" "}
                        <Link inline onClick={() => openExternal(otherPreset.keyUrl)}>
                          {otherPreset.keyUrl.replace(/^https:\/\//, "")}
                        </Link>
                        .{" "}
                      </>
                    ) : null}
                    Stored only on this computer (in this add-in's browser storage) – never in the document.
                  </>
                }
              >
                <Input
                  type={showKey ? "text" : "password"}
                  value={settings.otherApiKey}
                  onChange={(_, d) => onChange({ otherApiKey: d.value })}
                  placeholder={otherPreset?.id === "groq" ? "gsk_…" : "sk-…"}
                  autoComplete="off"
                  contentAfter={
                    <Button
                      size="small"
                      appearance="transparent"
                      aria-label={showKey ? "Hide key" : "Show key"}
                      icon={showKey ? <EyeOff20Regular /> : <Eye20Regular />}
                      onClick={() => setShowKey(!showKey)}
                    />
                  }
                />
              </Field>
              <p className="zee-muted">
                Chat, editing, proofreading and the quick actions work with this key. Reading PDFs is Gemini-only, and the cheaper bulk model is skipped (everything uses
                the model below).
              </p>
            </div>
          ) : settings.authMode === "apiKey" ? (
            <Field
              label="Gemini API key"
              hint={
                <>
                  Get a free key at{" "}
                  <Link inline onClick={() => openExternal("https://aistudio.google.com/apikey")}>
                    aistudio.google.com/apikey
                  </Link>
                  . Stored only on this computer (in this add-in's browser storage) – never in the document.
                </>
              }
            >
              <Input
                type={showKey ? "text" : "password"}
                value={settings.apiKey}
                onChange={(_, d) => onChange({ apiKey: d.value })}
                placeholder="AIza…"
                autoComplete="off"
                contentAfter={
                  <Button
                    size="small"
                    appearance="transparent"
                    aria-label={showKey ? "Hide key" : "Show key"}
                    icon={showKey ? <EyeOff20Regular /> : <Eye20Regular />}
                    onClick={() => setShowKey(!showKey)}
                  />
                }
              />
            </Field>
          ) : (
            <div className="zee-form">
              <Field label="OAuth Client ID (Web application)" hint="Google Cloud console → APIs & Services → Credentials">
                <Input value={settings.googleClientId} onChange={(_, d) => onChange({ googleClientId: d.value })} placeholder="1234-abc.apps.googleusercontent.com" />
              </Field>
              <Field label="Google Cloud project ID" hint="Vertex AI API must be enabled; usage is billed to this project.">
                <Input value={settings.gcpProjectId} onChange={(_, d) => onChange({ gcpProjectId: d.value })} placeholder="my-project-123" />
              </Field>
              <Field label="Vertex AI location" hint='"global" works for most Gemini models; or e.g. us-central1, europe-west4.'>
                <Input value={settings.gcpLocation} onChange={(_, d) => onChange({ gcpLocation: d.value })} placeholder="global" />
              </Field>
              <div className="zee-row">
                {signedIn ? (
                  <>
                    <span className="zee-chip">
                      <Person16Regular /> {session?.email ?? "Signed in"} · {minutesLeft(session)} min left
                    </span>
                    <Button
                      size="small"
                      onClick={() => {
                        signOutGoogle();
                        setSession(null);
                      }}
                    >
                      Sign out
                    </Button>
                  </>
                ) : (
                  <Button appearance="primary" onClick={signIn} disabled={signingIn || !settings.googleClientId.trim()} icon={signingIn ? <Spinner size="tiny" /> : undefined}>
                    Sign in with Google
                  </Button>
                )}
              </div>
            </div>
          )}

          <div className="zee-row">
            <Button onClick={runTest} disabled={testing} icon={testing ? <Spinner size="tiny" /> : undefined}>
              Test connection &amp; refresh models
            </Button>
          </div>
          {status && (
            <MessageBar intent={status.intent}>
              <MessageBarBody>{status.text}</MessageBarBody>
            </MessageBar>
          )}
        </section>

        <Divider />

        <section>
          <h3>Model</h3>
          <Field
            label={other ? `${otherName} model` : "Gemini model"}
            hint={other ? "Model ID exactly as your provider lists it. Click ⟳ to load the models your key can use." : (modelHint(settings.model) ?? "Custom model ID")}
          >
            <div className="zee-row">
              <Combobox
                className="zee-grow"
                freeform
                value={modelText}
                selectedOptions={[settings.model]}
                onChange={(e) => setModelText(e.target.value)}
                onBlur={() => commitModel(modelText)}
                onOptionSelect={(_, d) => {
                  if (d.optionValue) {
                    setModelText(d.optionValue);
                    commitModel(d.optionValue);
                  }
                }}
              >
                {other ? (
                  otherModels.length > 0 ? (
                    <OptionGroup label={`From ${otherName}`}>
                      {otherModels.map((id) => (
                        <Option key={id} value={id}>
                          {id}
                        </Option>
                      ))}
                    </OptionGroup>
                  ) : (
                    <Option value={settings.model || " "} disabled>
                      Click ⟳ to load models
                    </Option>
                  )
                ) : (
                  <>
                    <OptionGroup label="Recommended">
                      {DEFAULT_MODELS.map((m) => (
                        <Option key={m.id} value={m.id} text={m.id}>
                          {m.label}
                        </Option>
                      ))}
                    </OptionGroup>
                    {extraModels.length > 0 && (
                      <OptionGroup label="From your account">
                        {extraModels.map((id) => (
                          <Option key={id} value={id}>
                            {id}
                          </Option>
                        ))}
                      </OptionGroup>
                    )}
                  </>
                )}
              </Combobox>
              {settings.authMode !== "google" && (
                <Button
                  aria-label="Refresh model list"
                  title="Load the models available to your key or proxy"
                  icon={refreshing ? <Spinner size="tiny" /> : <ArrowSync16Regular />}
                  disabled={
                    refreshing ||
                    (settings.authMode === "apiKey" ? !settings.apiKey.trim() : other ? !settings.otherApiKey.trim() : !settings.proxyUrl.trim())
                  }
                  onClick={refreshModels}
                />
              )}
            </div>
          </Field>
          <Field label="Creativity (temperature)">
            <Dropdown
              value={TEMPS.find((t) => t.value === (settings.temperature === null ? "auto" : String(settings.temperature)))?.label ?? String(settings.temperature)}
              selectedOptions={[settings.temperature === null ? "auto" : String(settings.temperature)]}
              onOptionSelect={(_, d) => onChange({ temperature: d.optionValue === "auto" ? null : Number(d.optionValue) })}
            >
              {TEMPS.map((t) => (
                <Option key={t.value} value={t.value}>
                  {t.label}
                </Option>
              ))}
            </Dropdown>
          </Field>
        </section>

        <Divider />

        <section>
          <h3>Appearance</h3>
          <RadioGroup layout="horizontal" value={settings.theme} onChange={(_, d) => onChange({ theme: d.value as ThemeMode })} aria-label="Theme">
            <Radio value="auto" label="Match Office" />
            <Radio value="light" label="Light" />
            <Radio value="dark" label="Dark" />
          </RadioGroup>
          <p className="zee-muted">"Match Office" follows your Office theme (File → Account → Office Theme). Tip: the moon/sun button in the header switches quickly.</p>
        </section>

        <Divider />

        <section>
          <h3>Document context</h3>
          <Field label="Chat reads by default" hint="You can also switch this above the chat box.">
            <Dropdown
              value={CONTEXTS.find((c) => c.value === settings.docContext)?.label}
              selectedOptions={[settings.docContext]}
              onOptionSelect={(_, d) => d.optionValue && onChange({ docContext: d.optionValue as DocContext })}
            >
              {CONTEXTS.map((c) => (
                <Option key={c.value} value={c.value}>
                  {c.label}
                </Option>
              ))}
            </Dropdown>
          </Field>
          <Field label="Max document characters sent per prompt" hint="Longer documents are trimmed in the middle (start and end are kept). ~4 characters ≈ 1 token.">
            <Input
              type="number"
              min={1000}
              max={1000000}
              step={5000}
              value={String(settings.maxContextChars)}
              onChange={(_, d) => {
                const n = Number(d.value);
                if (Number.isFinite(n)) onChange({ maxContextChars: Math.min(1000000, Math.max(1000, Math.round(n))) });
              }}
            />
          </Field>
        </section>

        <Divider />

        <section>
          <h3>Editing</h3>
          <Switch
            checked={settings.trackChanges}
            onChange={(_, d) => onChange({ trackChanges: d.checked })}
            label="Apply edits as Track Changes (accept or reject them in Word's Review tab)"
          />
          {!canTrack && <p className="zee-muted">Track Changes needs Word for Microsoft 365 / Word 2021 or later (WordApi 1.4). On older Word, edits are applied directly.</p>}
          <Switch
            checked={settings.preserveFormatting}
            onChange={(_, d) => onChange({ preserveFormatting: d.checked })}
            label="Keep formatting: change only the words that differ (bold, links, styles stay)"
          />
          <Switch
            checked={settings.autoDirection}
            onChange={(_, d) => onChange({ autoDirection: d.checked })}
            label="Auto direction: Arabic right-to-left and right-aligned, English left-aligned, headings centred"
          />
          <Field label="Arabic font" hint="Used for Arabic text Zee AI writes. Pick from the list or type any installed font.">
            <ArabicFontPicker value={settings.arabicFont} onChange={(arabicFont) => onChange({ arabicFont })} />
          </Field>
        </section>

        <Divider />

        <section>
          <h3>Protected terms</h3>
          <Field
            label="Never rewrite, translate or “correct” these terms"
            hint="Comma or line separated. Used by every action, the chat, drafting and translation."
          >
            <Textarea rows={4} resize="vertical" value={settings.protectedTerms} onChange={(_, d) => onChange({ protectedTerms: d.value })} />
          </Field>
          <div className="zee-row">
            <Button size="small" disabled={settings.protectedTerms === DEFAULT_PROTECTED_TERMS} onClick={() => onChange({ protectedTerms: DEFAULT_PROTECTED_TERMS })}>
              Reset to petroleum defaults
            </Button>
          </div>
        </section>

        <Divider />

        <section ref={actionsRef}>
          <h3>My quick actions</h3>
          <p className="zee-muted">Your own one-click actions, listed under “Mine” in the toolbar.</p>
          <CustomActionsEditor actions={settings.customActions} onChange={(customActions) => onChange({ customActions })} />
        </section>

        <Divider />

        <section>
          <h3>Inserting text</h3>
          <Switch
            checked={settings.richInsert}
            onChange={(_, d) => onChange({ richInsert: d.checked })}
            label="Keep formatting from AI answers (headings, bold, lists, tables)"
          />
          <Field label="Default translation language">
            <Dropdown value={settings.translateTo} selectedOptions={[settings.translateTo]} onOptionSelect={(_, d) => d.optionValue && onChange({ translateTo: d.optionValue })}>
              {LANGUAGES.map((l) => (
                <Option key={l} value={l}>
                  {l}
                </Option>
              ))}
            </Dropdown>
          </Field>
          <p className="zee-muted">
            Nothing is changed without your click: edits show a diff first, and every insert can be undone step by step with ↶ ({undoDepth} step{undoDepth === 1 ? "" : "s"} available).
          </p>
          <div className="zee-row">
            <Button size="small" disabled={!undoDepth} onClick={onClearUndo}>
              Clear undo history (keep the text)
            </Button>
          </div>
        </section>

        <Divider />

        <section>
          <h3>Team setup</h3>
          <p className="zee-muted">
            Share one glossary and set of quick actions with your team. Team terms and actions are added to your own (marked “team” under Mine).
          </p>
          {settings.team ? (
            <MessageBar intent="success">
              <MessageBarBody>
                Team config{settings.team.name ? ` “${settings.team.name}”` : ""}: {settings.team.customActions.length} action{settings.team.customActions.length === 1 ? "" : "s"},{" "}
                {settings.team.protectedTerms.split(/[,;\n]+/).filter((t) => t.trim()).length} terms
                {settings.teamSource ? ` · from ${settings.teamSource === "file" ? "a file" : settings.teamSource === "proxy" ? "the team proxy" : "a URL"}` : ""}
                {settings.teamLoadedAt ? ` · ${new Date(settings.teamLoadedAt).toLocaleString()}` : ""}
              </MessageBarBody>
            </MessageBar>
          ) : null}
          <Field label="Load from an address" hint='An https link to a team file (e.g. a raw GitHub file), or type "proxy" to load it from your team proxy. Reloaded each time the pane opens.'>
            <div className="zee-row">
              <Input className="zee-grow" value={teamUrl} placeholder="https://… or proxy" onChange={(_, d) => setTeamUrl(d.value)} />
              <Button
                disabled={!teamUrl.trim() || teamBusy}
                icon={teamBusy ? <Spinner size="tiny" /> : undefined}
                onClick={async () => {
                  setTeamBusy(true);
                  try {
                    const team = await fetchTeamConfig(teamUrl, settings);
                    onChange({ team, teamSource: teamUrl.trim(), teamLoadedAt: Date.now() });
                    setTeamStatus({ intent: "success", text: `Team config loaded: ${team.customActions.length} actions.` });
                  } catch (e) {
                    setTeamStatus({ intent: "error", text: (e as Error).message });
                  } finally {
                    setTeamBusy(false);
                  }
                }}
              >
                Load
              </Button>
            </div>
          </Field>
          <div className="zee-row">
            <label className="zee-file-pick">
              <span className="zee-muted">Import a team file:</span>
              <input
                type="file"
                accept=".json,application/json"
                onChange={async (e) => {
                  const f = e.target.files?.[0];
                  if (!f) return;
                  try {
                    const team = parseTeamConfig(await f.text());
                    onChange({ team, teamSource: "file", teamLoadedAt: Date.now() });
                    setTeamStatus({ intent: "success", text: `Imported “${f.name}”: ${team.customActions.length} actions.` });
                  } catch (err) {
                    setTeamStatus({ intent: "error", text: (err as Error).message });
                  }
                }}
              />
            </label>
          </div>
          <div className="zee-row">
            <Button
              size="small"
              onClick={async () => {
                const json = exportTeamConfig(settings, "Zee AI team");
                const copied = await copyText(json);
                const saved = downloadText("zee-ai-team.json", json);
                setTeamStatus({
                  intent: "info",
                  text: `${saved ? "Saved zee-ai-team.json (check Downloads). " : ""}${copied ? "Also copied to the clipboard." : ""} Share it, or put it on GitHub / the team proxy (TEAM_CONFIG).`,
                });
              }}
            >
              Export my glossary &amp; actions
            </Button>
            {settings.team && (
              <Button size="small" onClick={() => onChange({ team: null, teamSource: "", teamLoadedAt: 0 })}>
                Remove team config
              </Button>
            )}
          </div>
          {teamStatus && (
            <MessageBar intent={teamStatus.intent}>
              <MessageBarBody>{teamStatus.text}</MessageBarBody>
            </MessageBar>
          )}
        </section>

        <Divider />

        <section>
          <h3>Cost guard</h3>
          <Field label="Ask before sending prompts larger than (tokens)" hint="Default 20,000. 0 = never ask. ~4 characters ≈ 1 token.">
            <Input
              type="number"
              min={0}
              step={1000}
              value={String(settings.costGuardTokens)}
              onChange={(_, d) => {
                const n = Number(d.value);
                if (Number.isFinite(n) && n >= 0) onChange({ costGuardTokens: Math.round(n) });
              }}
            />
          </Field>
          <Switch
            checked={settings.useBulkModel}
            onChange={(_, d) => onChange({ useBulkModel: d.checked })}
            label="Use a cheaper model for whole-document passes (proofreading, long summaries)"
          />
          <Field label="Bulk model" hint="Flash-Lite is the fastest and cheapest. The consistency and reference checks always use the main model.">
            <Input value={settings.bulkModel} disabled={!settings.useBulkModel} onChange={(_, d) => onChange({ bulkModel: d.value })} />
          </Field>
          <div className="zee-row">
            <span className="zee-muted">Repeated requests are answered from a session cache ({cacheSize} saved answers).</span>
            <Button
              size="small"
              disabled={!cacheSize}
              onClick={() => {
                clearResponseCache();
                setCacheSize(0);
              }}
            >
              Clear cache
            </Button>
          </div>
        </section>

        <Divider />

        <UsageSection session={usage} prices={{ inPerM: settings.priceInPerM, outPerM: settings.priceOutPerM }} onPrices={(p) => onChange({ priceInPerM: p.inPerM, priceOutPerM: p.outPerM })} />
        <p className="zee-muted">
          Current model: {modelLabel(settings.model)}. Quotas and billing are managed in Google AI Studio (API key) or the Google Cloud console (Vertex AI).
        </p>

        <Divider />

        <BackupSection settings={settings} defaults={DEFAULT_SETTINGS} onImport={onImport} loadPrompts={loadPrompts} savePrompts={savePrompts} />
      </DrawerBody>
    </OverlayDrawer>
  );
}

const ARABIC_FONTS = ["Simplified Arabic", "Arial", "Traditional Arabic", "Sakkal Majalla", "Tahoma", "Times New Roman", "Segoe UI", "Calibri", "Arabic Typesetting", "Dubai"];
const SAME_AS_DOC = "Same as document";

/** Free-text font combobox: "" = keep the document's font. */
function ArabicFontPicker({ value, onChange }: { value: string; onChange: (font: string) => void }) {
  const [text, setText] = useState(value || SAME_AS_DOC);
  useEffect(() => setText(value || SAME_AS_DOC), [value]);
  const commit = (v: string) => onChange(v.trim() === SAME_AS_DOC ? "" : v.trim());
  return (
    <Combobox
      freeform
      value={text}
      selectedOptions={[value || SAME_AS_DOC]}
      onChange={(e) => setText(e.target.value)}
      onBlur={() => commit(text)}
      onOptionSelect={(_, d) => {
        if (d.optionValue) {
          setText(d.optionValue);
          commit(d.optionValue);
        }
      }}
    >
      <Option value={SAME_AS_DOC}>{SAME_AS_DOC}</Option>
      {ARABIC_FONTS.map((f) => (
        <Option key={f} value={f} text={f}>
          <span style={{ fontFamily: `'${f}'` }}>{f} – نص عربي</span>
        </Option>
      ))}
    </Combobox>
  );
}
