import { Menu, MenuButton, MenuDivider, MenuGroupHeader, MenuItem, MenuList, MenuPopover, MenuTrigger, Toolbar, ToolbarButton, Tooltip } from "@fluentui/react-components";
import {
  Compose20Regular,
  DocumentBulletList20Regular,
  DocumentCheckmark20Regular,
  Flash20Regular,
  ImageSparkle20Regular,
  Table20Regular,
  Emoji20Regular,
  Sparkle20Regular,
  TextBulletListSquare20Regular,
  TextExpand20Regular,
  TextGrammarCheckmark20Regular,
  Translate20Regular,
} from "@fluentui/react-icons";
import { LANGUAGES, TONES, type SummaryStyle, type ToneId } from "../services/prompts";
import type { CustomAction } from "../services/settingsService";
import { TEMPLATES, type ReportTemplate } from "../services/templates";

export interface QuickActionHandlers {
  onRewrite: () => void;
  onGrammar: () => void;
  onTone: (tone: ToneId) => void;
  onSummarize: (style: SummaryStyle) => void;
  onLength: (dir: "expand" | "shorten") => void;
  onTranslate: (language: string) => void;
  onBilingual: (layout: "stacked" | "table") => void;
  onDraft: () => void;
  onCustom: (action: CustomAction) => void;
  onManageActions: () => void;
  onTable: (action: "explain" | "commentary" | "clean" | "fromNotes") => void;
  onConsistency: () => void;
  onReference: () => void;
  onPicture: (kind: "caption" | "alt" | "describe") => void;
  onTemplate: (t: ReportTemplate) => void;
}

export function QuickActions(p: QuickActionHandlers & { disabled?: boolean; preferredLanguage: string; customActions: CustomAction[] }) {
  const languages = [p.preferredLanguage, ...LANGUAGES.filter((l) => l !== p.preferredLanguage)];
  return (
    <Toolbar className="zee-quick" size="small" aria-label="Quick actions">
      <Tooltip content="Rewrite & polish the selection" relationship="label">
        <ToolbarButton appearance="subtle" icon={<Sparkle20Regular />} disabled={p.disabled} onClick={p.onRewrite}>
          <span className="zee-quick-label">Rewrite</span>
        </ToolbarButton>
      </Tooltip>

      <Tooltip content="Grammar & spelling – the selection, or the whole document (as comments) if nothing is selected" relationship="label">
        <ToolbarButton appearance="subtle" icon={<TextGrammarCheckmark20Regular />} disabled={p.disabled} onClick={p.onGrammar}>
          <span className="zee-quick-label">Grammar</span>
        </ToolbarButton>
      </Tooltip>

      <Menu>
        <MenuTrigger disableButtonEnhancement>
          <MenuButton size="small" appearance="subtle" icon={<Emoji20Regular />} disabled={p.disabled} title="Tone" aria-label="Tone">
            <span className="zee-quick-label">Tone</span>
          </MenuButton>
        </MenuTrigger>
        <MenuPopover>
          <MenuList>
            <MenuGroupHeader>Rewrite the selection as…</MenuGroupHeader>
            {TONES.map((t) => (
              <MenuItem key={t.id} onClick={() => p.onTone(t.id)}>
                {t.label}
              </MenuItem>
            ))}
          </MenuList>
        </MenuPopover>
      </Menu>

      <Menu>
        <MenuTrigger disableButtonEnhancement>
          <MenuButton size="small" appearance="subtle" icon={<TextBulletListSquare20Regular />} disabled={p.disabled} title="Summarize" aria-label="Summarize">
            <span className="zee-quick-label">Summarize</span>
          </MenuButton>
        </MenuTrigger>
        <MenuPopover>
          <MenuList>
            <MenuGroupHeader>Selection – or the whole document if nothing is selected</MenuGroupHeader>
            <MenuItem onClick={() => p.onSummarize("bullets")}>Bullet points</MenuItem>
            <MenuItem onClick={() => p.onSummarize("executive")}>Executive summary</MenuItem>
            <MenuItem onClick={() => p.onSummarize("tldr")}>TL;DR (one or two sentences)</MenuItem>
          </MenuList>
        </MenuPopover>
      </Menu>

      <Menu>
        <MenuTrigger disableButtonEnhancement>
          <MenuButton size="small" appearance="subtle" icon={<TextExpand20Regular />} disabled={p.disabled} title="Expand or shorten" aria-label="Expand or shorten">
            <span className="zee-quick-label">Length</span>
          </MenuButton>
        </MenuTrigger>
        <MenuPopover>
          <MenuList>
            <MenuItem onClick={() => p.onLength("expand")} secondaryContent="≈2×">
              Expand
            </MenuItem>
            <MenuItem onClick={() => p.onLength("shorten")} secondaryContent="≈½">
              Shorten
            </MenuItem>
          </MenuList>
        </MenuPopover>
      </Menu>

      <Menu>
        <MenuTrigger disableButtonEnhancement>
          <MenuButton size="small" appearance="subtle" icon={<Translate20Regular />} disabled={p.disabled} title="Translate" aria-label="Translate">
            <span className="zee-quick-label">Translate</span>
          </MenuButton>
        </MenuTrigger>
        <MenuPopover className="zee-lang-popover">
          <MenuList>
            <MenuGroupHeader>Translate the selection into…</MenuGroupHeader>
            <MenuItem onClick={() => p.onTranslate(languages[0])}>{languages[0]}</MenuItem>
            <MenuDivider />
            {languages.slice(1).map((l) => (
              <MenuItem key={l} onClick={() => p.onTranslate(l)}>
                {l}
              </MenuItem>
            ))}
            <MenuDivider />
            <MenuGroupHeader>Bilingual – English + Arabic</MenuGroupHeader>
            <MenuItem onClick={() => p.onBilingual("stacked")}>English, then Arabic</MenuItem>
            <MenuItem onClick={() => p.onBilingual("table")}>Side by side (table)</MenuItem>
          </MenuList>
        </MenuPopover>
      </Menu>

      <Menu>
        <MenuTrigger disableButtonEnhancement>
          <MenuButton size="small" appearance="subtle" icon={<Flash20Regular />} disabled={p.disabled} title="My actions" aria-label="My actions">
            <span className="zee-quick-label">Mine</span>
          </MenuButton>
        </MenuTrigger>
        <MenuPopover className="zee-lang-popover">
          <MenuList>
            <MenuGroupHeader>My quick actions</MenuGroupHeader>
            {p.customActions.length === 0 && <MenuItem disabled>No custom actions yet</MenuItem>}
            {p.customActions.map((a) => (
              <MenuItem
                key={a.id}
                title={a.instruction}
                onClick={() => p.onCustom(a)}
                secondaryContent={`${a.id.startsWith("team-") ? "team · " : ""}${a.mode === "replace" ? "edit" : "insert"}`}
              >
                {a.name}
              </MenuItem>
            ))}
            <MenuDivider />
            <MenuItem onClick={p.onManageActions}>Add or edit actions…</MenuItem>
          </MenuList>
        </MenuPopover>
      </Menu>

      <Menu>
        <MenuTrigger disableButtonEnhancement>
          <MenuButton size="small" appearance="subtle" icon={<Table20Regular />} disabled={p.disabled} title="Table" aria-label="Table">
            <span className="zee-quick-label">Table</span>
          </MenuButton>
        </MenuTrigger>
        <MenuPopover>
          <MenuList>
            <MenuGroupHeader>Cursor in a table</MenuGroupHeader>
            <MenuItem onClick={() => p.onTable("explain")}>Explain this table</MenuItem>
            <MenuItem onClick={() => p.onTable("commentary")}>Write commentary below it</MenuItem>
            <MenuItem onClick={() => p.onTable("clean")} secondaryContent="keeps values">
              Clean up table…
            </MenuItem>
            <MenuDivider />
            <MenuGroupHeader>Notes selected</MenuGroupHeader>
            <MenuItem onClick={() => p.onTable("fromNotes")}>Turn notes into a table</MenuItem>
          </MenuList>
        </MenuPopover>
      </Menu>

      <Menu>
        <MenuTrigger disableButtonEnhancement>
          <MenuButton size="small" appearance="subtle" icon={<DocumentCheckmark20Regular />} disabled={p.disabled} title="Review" aria-label="Review">
            <span className="zee-quick-label">Review</span>
          </MenuButton>
        </MenuTrigger>
        <MenuPopover>
          <MenuList>
            <MenuItem onClick={p.onConsistency} secondaryContent="units, names, numbers">
              Consistency check
            </MenuItem>
            <MenuItem onClick={p.onReference} secondaryContent="PDF / DOCX">
              Check against a reference…
            </MenuItem>
          </MenuList>
        </MenuPopover>
      </Menu>

      <Menu>
        <MenuTrigger disableButtonEnhancement>
          <MenuButton size="small" appearance="subtle" icon={<ImageSparkle20Regular />} disabled={p.disabled} title="Picture" aria-label="Picture">
            <span className="zee-quick-label">Picture</span>
          </MenuButton>
        </MenuTrigger>
        <MenuPopover>
          <MenuList>
            <MenuGroupHeader>Select a picture or chart first</MenuGroupHeader>
            <MenuItem onClick={() => p.onPicture("caption")}>Write a caption</MenuItem>
            <MenuItem onClick={() => p.onPicture("alt")}>Generate alt text</MenuItem>
            <MenuItem onClick={() => p.onPicture("describe")}>Describe / analyse chart</MenuItem>
          </MenuList>
        </MenuPopover>
      </Menu>

      <Menu>
        <MenuTrigger disableButtonEnhancement>
          <MenuButton size="small" appearance="subtle" icon={<DocumentBulletList20Regular />} disabled={p.disabled} title="Templates" aria-label="Templates">
            <span className="zee-quick-label">Templates</span>
          </MenuButton>
        </MenuTrigger>
        <MenuPopover className="zee-lang-popover">
          <MenuList>
            <MenuGroupHeader>Report templates</MenuGroupHeader>
            {TEMPLATES.map((t) => (
              <MenuItem key={t.id} title={t.description} onClick={() => p.onTemplate(t)}>
                {t.name}
              </MenuItem>
            ))}
          </MenuList>
        </MenuPopover>
      </Menu>

      <Tooltip content="Draft new content (clause, section, email…)" relationship="label">
        <ToolbarButton appearance="subtle" icon={<Compose20Regular />} disabled={p.disabled} onClick={p.onDraft}>
          <span className="zee-quick-label">Draft</span>
        </ToolbarButton>
      </Tooltip>
    </Toolbar>
  );
}
