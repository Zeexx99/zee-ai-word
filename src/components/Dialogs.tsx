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
  Textarea,
} from "@fluentui/react-components";
import { DiffView } from "./DiffView";
import { TONES, type DraftLength, type ToneId } from "../services/prompts";

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
