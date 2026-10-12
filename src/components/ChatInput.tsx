import { useState, type KeyboardEvent } from "react";
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
  Textarea,
  Tooltip,
} from "@fluentui/react-components";
import { Bookmark20Regular, BookmarkAdd20Regular, Delete16Regular, Send20Filled, Stop20Filled } from "@fluentui/react-icons";
import { loadPrompts, savePrompts } from "../services/promptStore";
import type { DocContext } from "../services/settingsService";

const short = (p: string) => (p.length > 70 ? `${p.slice(0, 68)}…` : p);

const CONTEXT_LABEL: Record<DocContext, string> = {
  none: "No document text",
  selection: "Selected text",
  document: "Whole document",
};

interface Props {
  busy: boolean;
  docContext: DocContext;
  onDocContext: (v: DocContext) => void;
  /** e.g. "42 words selected" */
  selectionLabel: string | null;
  sessionTokens: number;
  onSend: (text: string) => void;
  onStop: () => void;
}

export function ChatInput({ busy, docContext, onDocContext, selectionLabel, sessionTokens, onSend, onStop }: Props) {
  const [text, setText] = useState("");
  const [prompts, setPrompts] = useState<string[]>(() => loadPrompts());
  const updatePrompts = (list: string[]) => {
    setPrompts(list);
    savePrompts(list);
  };

  const submit = () => {
    const t = text.trim();
    if (!t || busy) return;
    onSend(t);
    setText("");
  };

  const onKeyDown = (e: KeyboardEvent<HTMLTextAreaElement>) => {
    if (e.key === "Enter" && !e.shiftKey && !e.nativeEvent.isComposing) {
      e.preventDefault();
      submit();
    }
  };

  return (
    <div className="zee-input">
      <div className="zee-input-context">
        <Tooltip content="What the chat reads from your document with each message" relationship="description">
          <Dropdown
            size="small"
            className="zee-context-dd"
            aria-label="Document context"
            value={CONTEXT_LABEL[docContext]}
            selectedOptions={[docContext]}
            onOptionSelect={(_, d) => d.optionValue && onDocContext(d.optionValue as DocContext)}
          >
            {(Object.keys(CONTEXT_LABEL) as DocContext[]).map((k) => (
              <Option key={k} value={k}>
                {CONTEXT_LABEL[k]}
              </Option>
            ))}
          </Dropdown>
        </Tooltip>
        {docContext === "selection" && <span className="zee-chip zee-chip-inline">{selectionLabel ?? "nothing selected"}</span>}
        <span className="zee-spacer" />
        {sessionTokens > 0 && (
          <span className="zee-muted" title="Tokens used this session (see Settings for details)">
            Σ {sessionTokens >= 1000 ? `${(sessionTokens / 1000).toFixed(1)}k` : sessionTokens} tokens
          </span>
        )}
      </div>
      <div className="zee-input-row">
        <Textarea
          className="zee-textarea"
          value={text}
          onChange={(_, d) => setText(d.value)}
          onKeyDown={onKeyDown}
          placeholder="Ask or draft anything… (Enter to send, Shift+Enter for a new line)"
          resize="vertical"
          rows={2}
        />
        <Menu positioning="above-end">
          <MenuTrigger disableButtonEnhancement>
            <Tooltip content="Saved prompts" relationship="label">
              <Button appearance="subtle" icon={<Bookmark20Regular />} />
            </Tooltip>
          </MenuTrigger>
          <MenuPopover>
            <MenuList>
              <MenuGroupHeader>Saved prompts – click to use</MenuGroupHeader>
              {prompts.length === 0 && <MenuItem disabled>No saved prompts yet</MenuItem>}
              {prompts.map((p, i) => (
                <MenuItem key={i} title={p} onClick={() => setText(p)}>
                  {short(p)}
                </MenuItem>
              ))}
              <MenuDivider />
              <MenuItem icon={<BookmarkAdd20Regular />} disabled={!text.trim() || prompts.includes(text.trim())} onClick={() => updatePrompts([text.trim(), ...prompts])}>
                Save current text as a prompt
              </MenuItem>
              {prompts.length > 0 && (
                <Menu>
                  <MenuTrigger disableButtonEnhancement>
                    <MenuItem icon={<Delete16Regular />}>Remove a prompt</MenuItem>
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
          <Tooltip content="Stop generating" relationship="label">
            <Button appearance="secondary" icon={<Stop20Filled />} onClick={onStop} />
          </Tooltip>
        ) : (
          <Tooltip content="Send" relationship="label">
            <Button appearance="primary" icon={<Send20Filled />} onClick={submit} disabled={!text.trim()} />
          </Tooltip>
        )}
      </div>
    </div>
  );
}
