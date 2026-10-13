import { useEffect, useState } from "react";
import {
  Button,
  Checkbox,
  Dialog,
  DialogActions,
  DialogBody,
  DialogContent,
  DialogSurface,
  DialogTitle,
  Dropdown,
  Field,
  Option,
  Radio,
  RadioGroup,
  Spinner,
  Textarea,
} from "@fluentui/react-components";
import { DiffView } from "./DiffView";
import { DEFAULT_REFERENCE_QUESTION, TONES, type DraftLength, type ToneId } from "../services/prompts";
import type { ReportTemplate } from "../services/templates";
import { readReferenceFile, type ReferencePayload } from "../utils/files";
import { loadDailyReportPackage, parseDailyReportPackage, type DailyReportPackage } from "../qol/sharedPackage";

/* ------------------------------------------------------------------ confirm (optionally with a diff) */

export interface ConfirmState {
  title: string;
  message: string;
  confirmLabel?: string;
  /** Show a before/after diff (used before replacing text). */
  diff?: { before: string; after: string };
}

export function ConfirmDialog({ state, onResult }: { state: ConfirmState | null; onResult: (ok: boolean) => void }) {
  return (
    <Dialog open={!!state} onOpenChange={(_, d) => !d.open && onResult(false)}>
      <DialogSurface className="zee-dialog">
        <DialogBody>
          <DialogTitle>{state?.title}</DialogTitle>
          <DialogContent className="zee-form">
            <p className="zee-dialog-text">{state?.message}</p>
            {state?.diff && <DiffView before={state.diff.before} after={state.diff.after} />}
          </DialogContent>
          <DialogActions>
            <Button appearance="primary" onClick={() => onResult(true)}>
              {state?.confirmLabel ?? "OK"}
            </Button>
            <Button onClick={() => onResult(false)}>Cancel</Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}

/* ------------------------------------------------------------------ draft */

export interface DraftChoice {
  request: string;
  length: DraftLength;
  tone: ToneId | "match";
  useContext: boolean;
}

const EXAMPLES = [
  "An NDA non-compete clause: 12 months, worldwide, standard carve-outs",
  "An introduction section for this report",
  "A polite email declining the meeting and proposing next week",
  "A project risk section with a table of risks, impact and mitigation",
];

const LENGTHS: Array<{ id: DraftLength; label: string }> = [
  { id: "short", label: "Short (100–200 words)" },
  { id: "medium", label: "Medium (300–500 words)" },
  { id: "long", label: "Long (800–1,200 words, with headings)" },
];

export function DraftDialog({ open, hasSelection, onClose, onRun }: { open: boolean; hasSelection: boolean; onClose: () => void; onRun: (c: DraftChoice) => void }) {
  const [request, setRequest] = useState("");
  const [length, setLength] = useState<DraftLength>("medium");
  const [tone, setTone] = useState<ToneId | "match">("match");
  const [useContext, setUseContext] = useState(true);
  useEffect(() => {
    if (open) setUseContext(true);
  }, [open]);

  const toneLabel = tone === "match" ? "Match the document" : TONES.find((t) => t.id === tone)?.label ?? tone;
  const submit = () => request.trim() && onRun({ request: request.trim(), length, tone, useContext });

  return (
    <Dialog open={open} onOpenChange={(_, d) => !d.open && onClose()}>
      <DialogSurface className="zee-dialog">
        <DialogBody>
          <DialogTitle>Draft new content</DialogTitle>
          <DialogContent className="zee-form">
            <Field label="What should Zee AI write?">
              <Textarea
                value={request}
                onChange={(_, d) => setRequest(d.value)}
                rows={4}
                resize="vertical"
                placeholder={EXAMPLES[0]}
                onKeyDown={(e) => {
                  if (e.key === "Enter" && (e.ctrlKey || e.metaKey)) submit();
                }}
              />
            </Field>
            <div className="zee-examples">
              {EXAMPLES.map((x) => (
                <Button key={x} size="small" appearance="transparent" onClick={() => setRequest(x)}>
                  {x}
                </Button>
              ))}
            </div>
            <Field label="Length">
              <Dropdown value={LENGTHS.find((l) => l.id === length)?.label} selectedOptions={[length]} onOptionSelect={(_, d) => setLength(d.optionValue as DraftLength)}>
                {LENGTHS.map((l) => (
                  <Option key={l.id} value={l.id}>
                    {l.label}
                  </Option>
                ))}
              </Dropdown>
            </Field>
            <Field label="Tone">
              <Dropdown value={toneLabel} selectedOptions={[tone]} onOptionSelect={(_, d) => setTone(d.optionValue as ToneId | "match")}>
                <Option value="match">Match the document</Option>
                {TONES.map((t) => (
                  <Option key={t.id} value={t.id}>
                    {t.label}
                  </Option>
                ))}
              </Dropdown>
            </Field>
            <Checkbox
              checked={useContext}
              onChange={(_, d) => setUseContext(!!d.checked)}
              label={hasSelection ? "Use the selected text as context" : "Use the start of the document as context"}
            />
            <p className="zee-muted">The draft appears in the chat first. You choose where it goes: after the cursor, replacing the selection, or at the end.</p>
          </DialogContent>
          <DialogActions>
            <Button appearance="primary" disabled={!request.trim()} onClick={submit}>
              Draft
            </Button>
            <Button onClick={onClose}>Cancel</Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}

/* ------------------------------------------------------------------ check against a reference */

export interface ReferenceChoice {
  reference: ReferencePayload;
  question: string;
}

export function ReferenceDialog({
  open,
  onClose,
  onRun,
  mode = "reference",
}: {
  open: boolean;
  onClose: () => void;
  onRun: (c: ReferenceChoice) => void;
  /** "compare" = the file is an older version of this document. */
  mode?: "reference" | "compare";
}) {
  const compare = mode === "compare";
  const [file, setFile] = useState<ReferencePayload | null>(null);
  const [error, setError] = useState("");
  const [reading, setReading] = useState(false);
  const [question, setQuestion] = useState(DEFAULT_REFERENCE_QUESTION);
  useEffect(() => {
    if (open) {
      setFile(null);
      setError("");
    }
  }, [open]);

  const pick = async (f: File | undefined) => {
    if (!f) return;
    setReading(true);
    setError("");
    try {
      setFile(await readReferenceFile(f));
    } catch (e) {
      setFile(null);
      setError((e as Error).message);
    } finally {
      setReading(false);
    }
  };

  return (
    <Dialog open={open} onOpenChange={(_, d) => !d.open && onClose()}>
      <DialogSurface className="zee-dialog">
        <DialogBody>
          <DialogTitle>{compare ? "Compare with an older version" : "Check against a reference"}</DialogTitle>
          <DialogContent className="zee-form">
            <Field
              label={compare ? "Older version of this document" : "Reference document"}
              hint={
                compare
                  ? "DOCX or TXT works best (only the changed paragraphs are sent). A PDF is sent whole. Max 15 MB."
                  : "A procedure, standard, specification or earlier report: PDF, DOCX, TXT/MD/CSV (max 15 MB). It is sent to Gemini for this check only."
              }
              validationMessage={error || undefined}
            >
              <div className="zee-file-pick">
                <input type="file" accept=".pdf,.docx,.txt,.md,.csv,.json,.xml,.html,image/*" onChange={(e) => pick(e.target.files?.[0])} />
                {reading && <Spinner size="tiny" label="Reading…" />}
                {file && <span className="zee-chip">📎 {file.name}</span>}
              </div>
            </Field>
            {!compare && (
              <Field label="What should be checked?">
                <Textarea value={question} rows={4} resize="vertical" onChange={(_, d) => setQuestion(d.value)} />
              </Field>
            )}
            <p className="zee-muted">
              {compare
                ? "You get the changes in numbers, conclusions and recommendations, each pointing to its paragraph – add them as Word comments if you like."
                : "Findings point to the paragraph in your document and can be added as Word comments."}
            </p>
          </DialogContent>
          <DialogActions>
            <Button appearance="primary" disabled={!file || (!compare && !question.trim())} onClick={() => file && onRun({ reference: file, question })}>
              {compare ? "Compare" : "Check"}
            </Button>
            <Button onClick={onClose}>Cancel</Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}

/* ------------------------------------------------------------------ report templates */

export interface TemplateChoice {
  template: ReportTemplate;
  fillFromSelection: boolean;
  location: "after" | "end";
}

export function TemplateDialog({
  template,
  hasSelection,
  onClose,
  onRun,
}: {
  template: ReportTemplate | null;
  hasSelection: boolean;
  onClose: () => void;
  onRun: (c: TemplateChoice) => void;
}) {
  const [fill, setFill] = useState(false);
  const [location, setLocation] = useState<"after" | "end">("after");
  useEffect(() => {
    if (template) setFill(hasSelection);
  }, [template, hasSelection]);

  return (
    <Dialog open={!!template} onOpenChange={(_, d) => !d.open && onClose()}>
      <DialogSurface className="zee-dialog">
        <DialogBody>
          <DialogTitle>{template?.name}</DialogTitle>
          <DialogContent className="zee-form">
            <p className="zee-dialog-text">{template?.description}</p>
            <pre className="zee-template-preview">{template?.outline}</pre>
            <Checkbox
              checked={fill}
              disabled={!hasSelection}
              onChange={(_, d) => setFill(!!d.checked)}
              label={hasSelection ? "Fill it in with AI from the selected notes" : "Select your notes first to have AI fill it in"}
            />
            <RadioGroup layout="horizontal" value={location} onChange={(_, d) => setLocation(d.value as "after" | "end")}>
              <Radio value="after" label="At the cursor" />
              <Radio value="end" label="At the end of the document" />
            </RadioGroup>
            <p className="zee-muted">
              {fill ? "The filled-in report appears in the chat first, so you can review it before inserting." : "The empty template is inserted with Word heading styles – References → Table of Contents works right away."}
            </p>
          </DialogContent>
          <DialogActions>
            <Button appearance="primary" onClick={() => template && onRun({ template, fillFromSelection: fill && hasSelection, location })}>
              {fill && hasSelection ? "Fill in with AI" : "Insert template"}
            </Button>
            <Button onClick={onClose}>Cancel</Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}

/* ------------------------------------------------------------------ daily report from the Excel add-in */

export interface DailyFromExcelChoice {
  pkg: DailyReportPackage;
  notes: string;
  location: "after" | "end";
  includeChart: boolean;
}

export function DailyFromExcelDialog({ open, onClose, onRun }: { open: boolean; onClose: () => void; onRun: (c: DailyFromExcelChoice) => void }) {
  const [pkg, setPkg] = useState<DailyReportPackage | null>(null);
  const [paste, setPaste] = useState("");
  const [error, setError] = useState("");
  const [notes, setNotes] = useState("");
  const [location, setLocation] = useState<"after" | "end">("after");
  const [chart, setChart] = useState(true);
  useEffect(() => {
    if (!open) return;
    setPkg(loadDailyReportPackage());
    setPaste("");
    setError("");
  }, [open]);
  const age = pkg ? Math.round((Date.now() - Date.parse(pkg.createdAt)) / 60000) : 0;
  return (
    <Dialog open={open} onOpenChange={(_, d) => !d.open && onClose()}>
      <DialogSurface className="zee-dialog">
        <DialogBody>
          <DialogTitle>Daily report from Excel</DialogTitle>
          <DialogContent className="zee-form">
            {pkg ? (
              <p className="zee-dialog-text">
                Found the report for <b>{pkg.dateLabel}</b> – {pkg.wells} wells, made in Excel {age < 1 ? "just now" : age < 120 ? `${age} min ago` : new Date(pkg.createdAt).toLocaleString()}
                {pkg.chartPng ? ", with a chart" : ""}.
              </p>
            ) : (
              <p className="zee-dialog-text">
                No report from Excel found here. In Excel run <b>Petroleum → Daily production report</b> – it also copies the report to the clipboard: paste it below.
              </p>
            )}
            <Field label={pkg ? "Use a different one (paste from Excel)" : "Paste the report from Excel"} validationMessage={error || undefined}>
              <Textarea
                rows={2}
                value={paste}
                placeholder="Ctrl+V"
                onChange={(_, d) => {
                  setPaste(d.value);
                  if (!d.value.trim()) return setError("");
                  try {
                    setPkg(parseDailyReportPackage(d.value));
                    setError("");
                  } catch (e) {
                    setError((e as Error).message);
                  }
                }}
              />
            </Field>
            <Field label="Extra notes for the report (optional)" hint="Events, shut-ins, workovers… Zee AI works them into the text.">
              <Textarea rows={3} value={notes} onChange={(_, d) => setNotes(d.value)} placeholder="e.g. W-7 shut in for ESP change since 06:00" />
            </Field>
            {pkg?.chartPng && <Checkbox checked={chart} onChange={(_, d) => setChart(!!d.checked)} label="Offer the chart for insertion (button under the report)" />}
            <RadioGroup layout="horizontal" value={location} onChange={(_, d) => setLocation(d.value as "after" | "end")}>
              <Radio value="after" label="At the cursor" />
              <Radio value="end" label="At the end" />
            </RadioGroup>
          </DialogContent>
          <DialogActions>
            <Button appearance="primary" disabled={!pkg} onClick={() => pkg && onRun({ pkg, notes, location, includeChart: chart && !!pkg.chartPng })}>
              Write report
            </Button>
            <Button onClick={onClose}>Cancel</Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}
