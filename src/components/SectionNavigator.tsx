import { useMemo, useState } from "react";
import { Button, Input, Spinner, Tooltip } from "@fluentui/react-components";
import { ArrowSync16Regular, Dismiss16Regular, Search16Regular } from "@fluentui/react-icons";
import type { Section } from "../utils/sections";

interface Props {
  sections: Section[];
  loading: boolean;
  /** Section currently used as scope (start paragraph). */
  active: number | null;
  onRefresh: () => void;
  onPick: (s: Section) => void;
  onClose: () => void;
}

/** Headings of the document: click one to select that section – your next request works on it. */
export function SectionNavigator({ sections, loading, active, onRefresh, onPick, onClose }: Props) {
  const [q, setQ] = useState("");
  const shown = useMemo(() => {
    const t = q.trim().toLowerCase();
    return t ? sections.filter((s) => s.title.toLowerCase().includes(t)) : sections;
  }, [q, sections]);
  const minLevel = sections.reduce((m, s) => Math.min(m, s.level), 9);
  return (
    <div className="zee-sections" role="navigation" aria-label="Document sections">
      <div className="zee-sections-bar">
        <strong>Sections</strong>
        <span className="zee-muted">click one to work on it</span>
        <span className="zee-spacer" />
        <Tooltip content="Re-read the headings" relationship="label">
          <Button size="small" appearance="subtle" icon={loading ? <Spinner size="extra-tiny" /> : <ArrowSync16Regular />} onClick={onRefresh} disabled={loading} />
        </Tooltip>
        <Tooltip content="Close" relationship="label">
          <Button size="small" appearance="subtle" icon={<Dismiss16Regular />} onClick={onClose} />
        </Tooltip>
      </div>
      {sections.length > 8 && <Input size="small" contentBefore={<Search16Regular />} placeholder="Filter headings" value={q} onChange={(_, d) => setQ(d.value)} />}
      <div className="zee-sections-list">
        {!loading && !sections.length && (
          <p className="zee-muted">No headings found. Use Word's Heading 1 / Heading 2 styles (Home → Styles) and click ⟳ – Zee AI's report templates already use them.</p>
        )}
        {shown.map((s) => (
          <button
            key={s.start}
            type="button"
            className={`zee-section-item${active === s.start ? " zee-section-active" : ""}`}
            style={{ paddingLeft: 6 + (s.level - minLevel) * 14 }}
            title={`¶${s.start}–${s.end} · ${s.words.toLocaleString()} words`}
            onClick={() => onPick(s)}
          >
            <span className="zee-section-title">{s.title}</span>
            <span className="zee-section-words">{s.words >= 1000 ? `${(s.words / 1000).toFixed(1)}k` : s.words}</span>
          </button>
        ))}
      </div>
    </div>
  );
}
