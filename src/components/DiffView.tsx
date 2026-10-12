import { memo, useMemo, useState } from "react";
import { Tab, TabList } from "@fluentui/react-components";
import { diffStats, diffWords } from "../utils/diff";

interface Props {
  before: string;
  after: string;
  /** Translations etc.: start on "New text" because a word diff is meaningless. */
  preferNew?: boolean;
}

/** Word-level diff preview: deletions struck through in red, insertions underlined in green. */
function DiffViewImpl({ before, after, preferNew }: Props) {
  const ops = useMemo(() => diffWords(before, after), [before, after]);
  const stats = useMemo(() => diffStats(ops), [ops]);
  const mostlyNew = stats.kept < 0.25;
  const [tab, setTab] = useState<"changes" | "new" | "original">(preferNew || mostlyNew ? "new" : "changes");
  const unchanged = before.trim() === after.trim();

  return (
    <div className="zee-diff">
      <div className="zee-diff-bar">
        <TabList size="small" selectedValue={tab} onTabSelect={(_, d) => setTab(d.value as typeof tab)}>
          <Tab value="changes">Changes</Tab>
          <Tab value="new">New text</Tab>
          <Tab value="original">Original</Tab>
        </TabList>
        <span className="zee-spacer" />
        <span className="zee-diff-stats" title="Words added / removed">
          {unchanged ? (
            "no changes"
          ) : (
            <>
              <span className="zee-diff-plus">+{stats.added}</span> <span className="zee-diff-minus">−{stats.removed}</span>
            </>
          )}
        </span>
      </div>
      <div className="zee-diff-body" aria-live="polite">
        {tab === "changes" &&
          ops.map((op, i) =>
            op.type === "eq" ? (
              <span key={i}>{op.text}</span>
            ) : op.type === "del" ? (
              <del key={i}>{op.text}</del>
            ) : (
              <ins key={i}>{op.text}</ins>
            )
          )}
        {tab === "new" && after}
        {tab === "original" && before}
      </div>
    </div>
  );
}

export const DiffView = memo(DiffViewImpl);
