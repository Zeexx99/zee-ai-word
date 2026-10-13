import { useEffect, useRef, useState, type KeyboardEvent } from "react";
import {
  Button,
  Dropdown,
  Menu,
  MenuDivider,
  MenuGroupHeader,
  MenuItem,
  MenuList,
  MenuPopover,
  MenuTrigger,
  Option,
  Spinner,
  Textarea,
  Tooltip,
} from "@fluentui/react-components";
import { Bookmark20Regular, Delete16Regular, Send20Filled, Star16Regular, Stop20Filled } from "@fluentui/react-icons";
import { loadPrompts, savePrompts } from "../services/promptStore";
import type { DocContext } from "../services/settingsService";
import { ContextChip, PromptChips, UsageMeter, usePromptBox } from "../qol/inputHelpers";
import type { DayUsage, Prices } from "../qol/usageLog";

const short = (p: string) => (p.length > 70 ? `${p.slice(0, 68)}…` : p);

const CONTEXT_LABEL: Record<DocContext, string> = {
  none: "No document text",
  selection: "Selected text",
  document: "Whole document",
  general: "General – find it for me",
};

const CONTEXT_HINT: Record<DocContext, string> = {
  none: "Chat only – nothing from the document is sent",
  selection: "The selected text is sent with each message",
  document: "The whole document is sent with each message",
  general: "Describe what you want – Zee AI finds the right part of the document, then edits it, adds text there, or answers from it",
};

const PLACEHOLDER: Record<DocContext, string> = {
  none: "Ask or draft anything… (Enter to send, ↑ for earlier prompts)",
  selection: "Ask about the selection or draft anything… (Enter to send, ↑ for earlier prompts)",
  document: "Ask about the document… (Enter to send, ↑ for earlier prompts)",
  general: "e.g. “Shorten the conclusion”, “Fix the units in the pump section”, “Add a risks paragraph after the summary”…",
};

interface Props {
  busy: boolean;
  /** Shown while Zee AI prepares a request (e.g. "Finding the relevant part…"). */
  status?: string | null;
  docContext: DocContext;
  onDocContext: (v: DocContext) => void;
  /** e.g. "42 words selected" (null = nothing selected, "preview" outside Word) */
  selectionLabel: string | null;
  /** Leave the selection out of the next message. */
  ignoreSelection: boolean;
  onIgnoreSelection: (v: boolean) => void;
  today: DayUsage;
  prices: Prices;
  /** Text to put into the box (Edit & resend); `n` changes on every request. */
  inject?: { text: string; n: number } | null;
  onSend: (text: string) => void;
  onStop: () => void;
}

export function ChatInput({ busy, status, docContext, onDocContext, selectionLabel, ignoreSelection, onIgnoreSelection, today, prices, inject, onSend, onStop }: Props) {
  const { text, setText, onRecallKey, onSent } = usePromptBox();
  const box = useRef<HTMLTextAreaElement>(null);
  const [prompts, setPrompts] = useState<string[]>(() => loadPrompts());
  const updatePrompts = (list: string[]) => {
    setPrompts(list);
    savePrompts(list);
  };

  useEffect(() => {
    if (!inject) return;
    setText(inject.text);
    window.setTimeout(() => box.current?.focus(), 0);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [inject?.n]);

  // Saved prompts can change from Settings (import) – re-read when the menu opens.
  const refreshPrompts = () => setPrompts(loadPrompts());

  const submit = () => {
    const t = text.trim();
    if (!t || busy) return;
    onSend(t);
    onSent(t);
  };

  const onKeyDown = (e: KeyboardEvent<HTMLTextAreaElement>) => {
    if (e.key === "Escape" && busy) {
      e.preventDefault();
      onStop();
      return;
    }
    if (onRecallKey(e)) return;
    if (e.key === "Enter" && !e.shiftKey && !e.nativeEvent.isComposing) {
      e.preventDefault();
      submit();
    }
  };

  const trimmed = text.trim();
  const isFavourite = !!trimmed && prompts.includes(trimmed);
  const showChip = docContext === "selection" && selectionLabel !== "preview";

  return (
    <div className="zee-input">
      <div className="zee-input-context">
        <Tooltip content={CONTEXT_HINT[docContext]} relationship="description">
          <Dropdown
            size="small"
            className="zee-context-dd"
            aria-label="Document context"
            value={CONTEXT_LABEL[docContext]}
            selectedOptions={[docContext]}
            onOptionSelect={(_, d) => d.optionValue && onDocContext(d.optionValue as DocContext)}
          >
            {(Object.keys(CONTEXT_LABEL) as DocContext[]).map((k) => (
              <Option key={k} value={k} text={CONTEXT_LABEL[k]}>
                {CONTEXT_LABEL[k]}
              </Option>
            ))}
          </Dropdown>
        </Tooltip>
        {showChip &&
          (selectionLabel ? (
            <ContextChip label={selectionLabel} ignored={ignoreSelection} onToggle={onIgnoreSelection} title="Sent with your next message" />
          ) : (
            <span className="zee-chip zee-chip-inline">nothing selected</span>
          ))}
        {docContext === "general" && <span className="zee-chip zee-chip-inline" title={CONTEXT_HINT.general}>Zee finds the spot</span>}
        <span className="zee-spacer" />
        <UsageMeter today={today} prices={prices} />
      </div>
      {status && (
        <div className="zee-input-status">
          <Spinner size="extra-tiny" /> <span>{status}</span>
        </div>
      )}
      {!trimmed && !busy && <PromptChips prompts={prompts} onPick={setText} />}
      <div className="zee-input-row">
        <Textarea
          className="zee-textarea"
          value={text}
          ref={box}
          onChange={(_, d) => setText(d.value)}
          onKeyDown={onKeyDown}
          placeholder={PLACEHOLDER[docContext]}
          resize="vertical"
          rows={2}
        />
        <Menu positioning="above-end" onOpenChange={(_, d) => d.open && refreshPrompts()}>
          <MenuTrigger disableButtonEnhancement>
            <Tooltip content="Favourite prompts" relationship="label">
              <Button appearance="subtle" icon={<Bookmark20Regular />} />
            </Tooltip>
          </MenuTrigger>
          <MenuPopover>
            <MenuList>
              <MenuGroupHeader>Favourite prompts – click to use</MenuGroupHeader>
              {prompts.length === 0 && <MenuItem disabled>No favourites yet</MenuItem>}
              {prompts.map((p, i) => (
                <MenuItem key={i} title={p} onClick={() => setText(p)}>
                  {short(p)}
                </MenuItem>
              ))}
              <MenuDivider />
              <MenuItem icon={<Star16Regular />} disabled={!trimmed || isFavourite} onClick={() => updatePrompts([trimmed, ...prompts])}>
                {isFavourite ? "Already a favourite" : "Add current text to favourites"}
              </MenuItem>
              {prompts.length > 0 && (
                <Menu>
                  <MenuTrigger disableButtonEnhancement>
                    <MenuItem icon={<Delete16Regular />}>Remove a favourite</MenuItem>
                  </MenuTrigger>
                  <MenuPopover>
                    <MenuList>
                      {prompts.map((p, i) => (
                        <MenuItem key={i} title={p} onClick={() => updatePrompts(prompts.filter((_, j) => j !== i))}>
                          {short(p)}
                        </MenuItem>
                      ))}
                    </MenuList>
                  </MenuPopover>
                </Menu>
              )}
            </MenuList>
          </MenuPopover>
        </Menu>
        {busy ? (
          <Tooltip content="Stop (Esc)" relationship="label">
            <Button appearance="secondary" icon={<Stop20Filled />} onClick={onStop} />
          </Tooltip>
        ) : (
          <Tooltip content="Send (Enter)" relationship="label">
            <Button appearance="primary" icon={<Send20Filled />} onClick={submit} disabled={!trimmed} />
          </Tooltip>
        )}
      </div>
    </div>
  );
}
