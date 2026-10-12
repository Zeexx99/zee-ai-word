import { useState } from "react";
import { Button, Dropdown, Field, Input, Option, Textarea } from "@fluentui/react-components";
import { Add16Regular, Delete16Regular } from "@fluentui/react-icons";
import type { CustomAction } from "../services/settingsService";

const MODES: Record<CustomAction["mode"], string> = {
  replace: "Edit the selection (shows a diff)",
  insert: "Write new text after the selection",
};

/** Add / edit / remove the user's own quick actions ("My actions" in the toolbar). */
export function CustomActionsEditor({ actions, onChange }: { actions: CustomAction[]; onChange: (a: CustomAction[]) => void }) {
  const [draft, setDraft] = useState<CustomAction>({ id: "", name: "", instruction: "", mode: "replace" });

  const update = (i: number, patch: Partial<CustomAction>) => onChange(actions.map((a, j) => (j === i ? { ...a, ...patch } : a)));

  const add = () => {
    if (!draft.name.trim() || !draft.instruction.trim()) return;
    onChange([...actions, { ...draft, id: `${Date.now().toString(36)}`, name: draft.name.trim(), instruction: draft.instruction.trim() }]);
    setDraft({ id: "", name: "", instruction: "", mode: "replace" });
  };

  return (
    <div className="zee-form">
      {actions.map((a, i) => (
        <div key={a.id} className="zee-action-card">
          <div className="zee-row">
            <Input className="zee-grow" size="small" value={a.name} onChange={(_, d) => update(i, { name: d.value })} aria-label="Action name" />
            <Button size="small" appearance="subtle" icon={<Delete16Regular />} aria-label={`Delete ${a.name}`} onClick={() => onChange(actions.filter((_, j) => j !== i))} />
          </div>
          <Textarea size="small" value={a.instruction} rows={2} resize="vertical" onChange={(_, d) => update(i, { instruction: d.value })} aria-label="Instruction" />
          <Dropdown size="small" value={MODES[a.mode]} selectedOptions={[a.mode]} onOptionSelect={(_, d) => d.optionValue && update(i, { mode: d.optionValue as CustomAction["mode"] })}>
            {(Object.keys(MODES) as CustomAction["mode"][]).map((m) => (
              <Option key={m} value={m}>
                {MODES[m]}
              </Option>
            ))}
          </Dropdown>
        </div>
      ))}

      <div className="zee-action-card zee-action-new">
        <Field label="New action name">
          <Input size="small" value={draft.name} placeholder="e.g. Daily report paragraph" onChange={(_, d) => setDraft({ ...draft, name: d.value })} />
        </Field>
        <Field label="What should Zee AI do with the selected text?">
          <Textarea
            size="small"
            rows={3}
            resize="vertical"
            value={draft.instruction}
            placeholder="e.g. Turn these notes into a short, formal paragraph for the weekly ESP performance report."
            onChange={(_, d) => setDraft({ ...draft, instruction: d.value })}
          />
        </Field>
        <Dropdown size="small" value={MODES[draft.mode]} selectedOptions={[draft.mode]} onOptionSelect={(_, d) => d.optionValue && setDraft({ ...draft, mode: d.optionValue as CustomAction["mode"] })}>
          {(Object.keys(MODES) as CustomAction["mode"][]).map((m) => (
            <Option key={m} value={m}>
              {MODES[m]}
            </Option>
          ))}
        </Dropdown>
        <div className="zee-row">
          <Button size="small" appearance="primary" icon={<Add16Regular />} disabled={!draft.name.trim() || !draft.instruction.trim()} onClick={add}>
            Add action
          </Button>
        </div>
      </div>
    </div>
  );
}
