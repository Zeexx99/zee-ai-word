import { memo, useMemo, useState, type ReactNode } from "react";
import ReactMarkdown, { type Components, type Options } from "react-markdown";
import remarkGfm from "remark-gfm";
import rehypeHighlight from "rehype-highlight";
import { common } from "lowlight";
import {
  Button,
  Menu,
  MenuItem,
  MenuList,
  MenuPopover,
  MenuTrigger,
  Spinner,
  SplitButton,
  Tooltip,
  type MenuButtonProps,
} from "@fluentui/react-components";
import {
  ArrowSwap16Regular,
  Checkmark16Regular,
  Comment16Regular,
  Copy16Regular,
  DocumentArrowDown16Regular,
  ArrowEnter16Regular,
} from "@fluentui/react-icons";
import type { ChatMessage } from "../hooks/useChat";
import type { GrammarIssue } from "../services/prompts";
import { linkCitations } from "../utils/chunks";
import { copyText, openExternal, useInsertActions } from "./insertActions";
import { modelLabel } from "../services/settingsService";
import { LOCATION_LABELS, toClipboardText, type InsertLocation } from "../services/wordService";
import { DiffView } from "./DiffView";
import { ErrorBubble, RegenerateButton, UserMessageTools } from "../qol/MessageTools";

const rehypePlugins: Options["rehypePlugins"] = [[rehypeHighlight, { languages: common, detect: false }]];
const remarkPlugins: Options["remarkPlugins"] = [remarkGfm];

function CopyButton({ text, label = "Copy" }: { text: string; label?: string }) {
  const [done, setDone] = useState(false);
  return (
    <Tooltip content={done ? "Copied" : label} relationship="label">
      <Button
        size="small"
        appearance="subtle"
        icon={done ? <Checkmark16Regular /> : <Copy16Regular />}
        onClick={async () => {
          if (await copyText(text)) {
            setDone(true);
            setTimeout(() => setDone(false), 1500);
          }
        }}
      />
    </Tooltip>
  );
}

const LOCATION_ICON: Record<InsertLocation, JSX.Element> = {
  replace: <ArrowSwap16Regular />,
  after: <ArrowEnter16Regular />,
  end: <DocumentArrowDown16Regular />,
};

const SHORT_LABEL: Record<InsertLocation, string> = {
  replace: "Replace selection",
  after: "Insert after cursor",
  end: "Append to end",
};

/** Insert bar under an assistant answer: main action + the other two locations + copy. */
function InsertBar({ message }: { message: ChatMessage }) {
  const actions = useInsertActions();
  const [working, setWorking] = useState(false);
  const primary: InsertLocation = message.defaultLocation ?? "after";
  const others = (["replace", "after", "end"] as InsertLocation[]).filter((l) => l !== primary);

  const run = async (loc: InsertLocation) => {
    setWorking(true);
    try {
      await actions.insert(message, loc);
    } finally {
      setWorking(false);
    }
  };

  return (
    <div className="zee-insert-bar">
      <Menu positioning="below-start">
        <MenuTrigger disableButtonEnhancement>
          {(triggerProps: MenuButtonProps) => (
            <SplitButton
              size="small"
              appearance="primary"
              icon={working ? <Spinner size="extra-tiny" /> : LOCATION_ICON[primary]}
              disabled={!actions.canInsert || working}
              menuButton={triggerProps}
              primaryActionButton={{ onClick: () => run(primary), title: LOCATION_LABELS[primary] }}
            >
              {SHORT_LABEL[primary]}
            </SplitButton>
          )}
        </MenuTrigger>
        <MenuPopover>
          <MenuList>
            {others.map((loc) => (
              <MenuItem key={loc} icon={LOCATION_ICON[loc]} onClick={() => run(loc)}>
                {LOCATION_LABELS[loc]}
              </MenuItem>
            ))}
          </MenuList>
        </MenuPopover>
      </Menu>
      <CopyButton text={toClipboardText(message.text, message.format ?? "markdown")} label="Copy text" />
      {message.applied && (
        <span className="zee-applied">
          <Checkmark16Regular /> {message.applied === "replace" ? "Replaced" : "Inserted"}
        </span>
      )}
      {!actions.canInsert && <span className="zee-muted">Open in Word to insert</span>}
    </div>
  );
}

function IssueList({ issues, empty, style = "fix" }: { issues: GrammarIssue[]; empty: string; style?: "fix" | "finding" }) {
  const actions = useInsertActions();
  if (!issues.length)
    return (
      <div className="zee-issues">
        <span className="zee-issues-none">
          <Checkmark16Regular /> {empty}
        </span>
      </div>
    );
  return (
    <div className="zee-issues">
      <strong>{issues.length === 1 ? "1 issue" : `${issues.length} issues`}</strong>
      <ul>
        {issues.slice(0, 200).map((i, k) => (
          <li key={k}>
            {i.paragraph ? (
              <button type="button" className="zee-para-chip" title="Show this paragraph" onClick={() => actions.goToParagraph(i.paragraph!)}>
                ¶{i.paragraph}
              </button>
            ) : null}
            {style === "finding" ? (
              <>
                <span className="zee-issue-type">{i.type}</span> {i.reason}
                {i.original && (
                  <>
                    {" "}
                    <del>{i.original}</del>
                  </>
                )}
                {i.fix && (
                  <>
                    {" "}
                    → <ins>{i.fix}</ins>
                  </>
                )}
              </>
            ) : (
              <>
                <del>{i.original}</del> → <ins>{i.fix}</ins> <span className="zee-issue-type">{i.type}</span>
                {i.reason && <span className="zee-muted"> – {i.reason}</span>}
              </>
            )}
          </li>
        ))}
      </ul>
      {issues.length > 200 && <p className="zee-muted">…and {issues.length - 200} more (all are added as comments).</p>}
    </div>
  );
}

/** "Add as Word comments" for grammar findings. */
function CommentButton({ message, primary }: { message: ChatMessage; primary?: boolean }) {
  const actions = useInsertActions();
  const [working, setWorking] = useState(false);
  return (
    <Tooltip content="Mark each issue in the document with a Word comment instead of changing the text" relationship="description">
      <Button
        size="small"
        appearance={primary ? "primary" : "secondary"}
        icon={working ? <Spinner size="extra-tiny" /> : <Comment16Regular />}
        disabled={!actions.canInsert || working || !!message.commented}
        onClick={async () => {
          setWorking(true);
          try {
            await actions.comment(message);
          } finally {
            setWorking(false);
          }
        }}
      >
        {message.commented ? "Comments added" : "Add as comments"}
      </Button>
    </Tooltip>
  );
}

/** Table clean-up proposal: changed cells highlighted (old value struck through), Apply writes only those cells. */
function TableProposal({ message }: { message: ChatMessage }) {
  const actions = useInsertActions();
  const [working, setWorking] = useState(false);
  const edit = message.tableEdit!;
  const changed = edit.after.reduce((n, r, i) => n + r.filter((v, j) => v.trim() !== (edit.before[i]?.[j] ?? "").trim()).length, 0);
  return (
    <>
      <strong>{changed === 1 ? "1 cell changes" : `${changed} cells change`}</strong>
      <table className="zee-grid-preview">
        <tbody>
          {edit.after.map((r, i) => (
            <tr key={i}>
              {r.map((v, j) => {
                const old = edit.before[i]?.[j] ?? "";
                const diff = v.trim() !== old.trim();
                return (
                  <td key={j} className={diff ? "zee-cell-changed" : undefined}>
                    {diff && <del>{old || "∅"}</del>}
                    {v}
                  </td>
                );
              })}
            </tr>
          ))}
        </tbody>
      </table>
      {edit.notes.length > 0 && (
        <ul className="zee-actions">
          {edit.notes.map((n, i) => (
            <li key={i}>{n}</li>
          ))}
        </ul>
      )}
      <div className="zee-insert-bar">
        <Button
          size="small"
          appearance="primary"
          disabled={!changed || !actions.canInsert || working || !!message.applied}
          icon={working ? <Spinner size="extra-tiny" /> : <Checkmark16Regular />}
          onClick={async () => {
            setWorking(true);
            try {
              await actions.applyTable(message);
            } finally {
              setWorking(false);
            }
          }}
        >
          {message.applied ? "Applied" : "Apply to table"}
        </Button>
        <span className="zee-muted">Only the changed cells are written; formatting stays.</span>
      </div>
    </>
  );
}

function PictureApplyBar({ message }: { message: ChatMessage }) {
  const actions = useInsertActions();
  const [working, setWorking] = useState(false);
  const kind = message.pictureResult!.kind;
  return (
    <div className="zee-insert-bar">
      <Button
        size="small"
        appearance="primary"
        disabled={!actions.canInsert || working || !!message.applied}
        icon={working ? <Spinner size="extra-tiny" /> : <Checkmark16Regular />}
        onClick={async () => {
          setWorking(true);
          try {
            await actions.applyPicture(message);
          } finally {
            setWorking(false);
          }
        }}
      >
        {message.applied ? "Applied" : kind === "caption" ? "Insert caption below picture" : "Set as alt text"}
      </Button>
      <CopyButton text={message.text} />
      <span className="zee-muted">Keep the picture selected.</span>
    </div>
  );
}

function CodeBlock({ children, text, lang }: { children: ReactNode; text: string; lang: string }) {
  return (
    <div className="zee-code">
      <div className="zee-code-bar">
        <span className="zee-code-lang">{lang || "text"}</span>
        <span className="zee-spacer" />
        <CopyButton text={text} />
      </div>
      <pre>{children}</pre>
    </div>
  );
}

interface HastNode {
  tagName?: string;
  value?: string;
  properties?: Record<string, unknown>;
  children?: HastNode[];
}
const hastText = (n?: HastNode): string => (!n ? "" : n.value ?? (n.children ?? []).map(hastText).join(""));

function formatTokens(n: number): string {
  return n >= 1000 ? `${(n / 1000).toFixed(1)}k` : String(n);
}

function MessageBubbleImpl({ message }: { message: ChatMessage }) {
  const actions = useInsertActions();
  const streaming = message.status === "streaming";
  const components = useMemo<Components>(
    () => ({
      pre: ({ node, children }) => {
        const code = (node as unknown as HastNode | undefined)?.children?.find((c) => c.tagName === "code");
        const classes = (code?.properties?.className as string[] | undefined) ?? [];
        const lang = (classes.find((c) => c.startsWith("language-")) ?? "").replace("language-", "");
        return (
          <CodeBlock text={hastText(code).replace(/\n$/, "")} lang={lang}>
            {children}
          </CodeBlock>
        );
      },
      table: ({ children }) => (
        <div className="zee-table">
          <div className="zee-table-scroll">
            <table>{children}</table>
          </div>
        </div>
      ),
      a: ({ href, children }) => (
        <a
          href={href}
          target="_blank"
          rel="noreferrer"
          className={href?.startsWith("#zee-para-") ? "zee-cite" : undefined}
          title={href?.startsWith("#zee-para-") ? "Show this paragraph in the document" : undefined}
          onClick={(e) => {
            const para = href?.match(/^#zee-para-(\d+)$/);
            if (para) {
              e.preventDefault();
              actions.goToParagraph(Number(para[1]));
              return;
            }
            // Task panes can't open new tabs reliably (desktop Word ignores target=_blank or navigates the pane away).
            if (href && /^https?:/i.test(href)) {
              e.preventDefault();
              openExternal(href);
            }
          }}
        >
          {children}
        </a>
      ),
    }),
    [actions]
  );
  const markdown = useMemo(() => (message.citations ? linkCitations(message.text) : message.text), [message.citations, message.text]);

  if (message.role === "user") {
    return (
      <div className="zee-msg zee-msg-user">
        <div className="zee-bubble">{message.text}</div>
        {message.contextLabel && <div className="zee-chip">📎 {message.contextLabel}</div>}
        <UserMessageTools text={message.text} />
      </div>
    );
  }

  if (message.status === "error") return <ErrorBubble message={message} />;

  const done = message.status === "done";
  const isEdit = done && message.original !== undefined && message.defaultLocation === "replace";
  const isDocFindings = done && message.scope === "document" && !!message.issues;
  const plain = message.format === "plain";
  const hasResult = done && !!message.text && !/^_(Stopped\.|\(empty response\))_$/.test(message.text);
  const noIssues = !!message.issues && message.issues.length === 0;

  return (
    <div className="zee-msg zee-msg-assistant">
      {message.title && <div className="zee-msg-title">{message.title}</div>}
      <div className={`zee-bubble ${plain ? "zee-plain" : "zee-markdown"}`}>
        {message.tableEdit && done ? (
          <TableProposal message={message} />
        ) : isDocFindings ? (
          <>
            <p className="zee-dialog-text">{message.text}</p>
            <IssueList
              issues={message.issues ?? []}
              style={message.issueStyle ?? "fix"}
              empty={message.issueStyle === "finding" ? "No problems found." : "No spelling or grammar issues found in the document."}
            />
          </>
        ) : isEdit ? (
          <>
            {message.issues && <IssueList issues={message.issues} empty="No spelling or grammar issues found." />}
            {!noIssues && <DiffView before={message.original ?? ""} after={message.text} preferNew={/^Translate/i.test(message.title ?? "")} />}
          </>
        ) : message.text ? (
          plain ? (
            message.text
          ) : (
            <ReactMarkdown remarkPlugins={remarkPlugins} rehypePlugins={rehypePlugins} components={components}>
              {markdown}
            </ReactMarkdown>
          )
        ) : streaming ? (
          <Spinner size="tiny" label={message.kind === "chat" ? "Thinking…" : "Writing…"} labelPosition="after" />
        ) : null}
        {streaming && message.text && <span className="zee-caret" />}
      </div>
      {isDocFindings && !noIssues && (
        <div className="zee-insert-bar">
          <CommentButton message={message} primary />
        </div>
      )}
      {message.pictureResult && hasResult && <PictureApplyBar message={message} />}
      {!isDocFindings && !message.tableEdit && !message.pictureResult && hasResult && !noIssues && <InsertBar message={message} />}
      {isEdit && message.issues && !noIssues && (
        <div className="zee-insert-bar">
          <CommentButton message={message} />
        </div>
      )}
      {!streaming && (
        <div className="zee-meta">
          {message.model && <span>{modelLabel(message.model)}</span>}
          {message.usage && (
            <span title="Prompt tokens · completion tokens">
              {formatTokens(message.usage.promptTokens)} in · {formatTokens(message.usage.outputTokens)} out
            </span>
          )}
          <span className="zee-spacer" />
          <RegenerateButton message={message} />
        </div>
      )}
    </div>
  );
}

export const MessageBubble = memo(MessageBubbleImpl);
