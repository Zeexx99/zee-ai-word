import { useEffect, useRef } from "react";
import { Button } from "@fluentui/react-components";
import type { ChatMessage } from "../hooks/useChat";
import { MessageBubble } from "./MessageBubble";

const SUGGESTIONS = [
  "Summarize this document in 5 bullet points",
  "Draft an introduction section for this document",
  "Draft an NDA non-compete clause (12 months, worldwide)",
  "What is unclear or inconsistent in the selected text?",
];

export function ChatFeed({ messages, onSuggestion }: { messages: ChatMessage[]; onSuggestion: (t: string) => void }) {
  const ref = useRef<HTMLDivElement>(null);
  const stick = useRef(true);

  useEffect(() => {
    const el = ref.current;
    if (el && stick.current) el.scrollTop = el.scrollHeight;
  }, [messages]);

  return (
    <div
      className="zee-feed"
      ref={ref}
      onScroll={(e) => {
        const el = e.currentTarget;
        stick.current = el.scrollHeight - el.scrollTop - el.clientHeight < 60;
      }}
    >
      {messages.length === 0 ? (
        <div className="zee-empty">
          <img src="assets/icon-64.png" alt="" width={48} height={48} />
          <h2>What shall we write today?</h2>
          <p>Select text and pick an action above, or ask below. You review every change before it touches your document.</p>
          <div className="zee-suggestions">
            {SUGGESTIONS.map((s) => (
              <Button key={s} appearance="outline" size="small" onClick={() => onSuggestion(s)}>
                {s}
              </Button>
            ))}
          </div>
        </div>
      ) : (
        messages.map((m) => <MessageBubble key={m.id} message={m} />)
      )}
    </div>
  );
}
