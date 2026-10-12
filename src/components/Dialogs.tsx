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

export function ReferenceDialog({ open, onClose, onRun }: { open: boolean; onClose: () => void; onRun: (c: ReferenceChoice) => void }) {
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
          <DialogTitle>Check against a reference</DialogTitle>
          <DialogContent className="zee-form">
            <Field
              label="Reference document"
              hint="A procedure, standard, specification or earlier report: PDF, DOCX, TXT/MD/CSV (max 15 MB). It is sent to Gemini for this check only."
              validationMessage={error || undefined}
            >
              <div className="zee-file-pick">
                <input type="file" accept=".pdf,.docx,.txt,.md,.csv,.json,.xml,.html,image/*" onChange={(e) => pick(e.target.files?.[0])} />
                {reading && <Spinner size="tiny" label="Reading…" />}
                {file && <span className="zee-chip">📎 {file.name}</span>}
              </div>
            </Field>
            <Field label="What should be checked?">
              <Textarea value={question} rows={4} resize="vertical" onChange={(_, d) => setQuestion(d.value)} />
            </Field>
            <p className="zee-muted">Findings point to the paragraph in your document and can be added as Word comments.</p>
          </DialogContent>
          <DialogActions>
            <Button appearance="primary" disabled={!file || !question.trim()} onClick={() => file && onRun({ reference: file, question })}>
              Check
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
