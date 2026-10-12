/** Team configuration file: shared glossary and quick actions (pure, unit-tested). */

export interface TeamAction {
  id: string;
  name: string;
  instruction: string;
  mode: "replace" | "insert";
}

export interface TeamConfig {
  zeeTeamConfig: 1;
  name?: string;
  protectedTerms: string;
  customActions: TeamAction[];
}

export function buildTeamConfig(name: string, protectedTerms: string, customActions: TeamAction[]): TeamConfig {
  return { zeeTeamConfig: 1, name: name || undefined, protectedTerms, customActions };
}

/** Validates any JSON the user imports or a URL returns. Throws a readable error. */
export function parseTeamConfig(raw: string | unknown): TeamConfig {
  let o: unknown = raw;
  if (typeof raw === "string") {
    try {
      o = JSON.parse(raw);
    } catch {
      throw new Error("This is not a valid Zee AI team file (not JSON).");
    }
  }
  if (!o || typeof o !== "object" || (o as { zeeTeamConfig?: unknown }).zeeTeamConfig !== 1) throw new Error("This is not a Zee AI team file (missing zeeTeamConfig: 1).");
  const c = o as Record<string, unknown>;
  const actions = Array.isArray(c.customActions) ? c.customActions : [];
  return {
    zeeTeamConfig: 1,
    name: typeof c.name === "string" ? c.name.slice(0, 80) : undefined,
    protectedTerms: typeof c.protectedTerms === "string" ? c.protectedTerms.slice(0, 20000) : "",
    customActions: actions
      .filter((a): a is Record<string, unknown> => !!a && typeof a === "object")
      .filter((a) => typeof a.name === "string" && typeof a.instruction === "string" && a.name.trim() !== "" && a.instruction.trim() !== "")
      .slice(0, 100)
      .map((a, i) => ({
        id: `team-${typeof a.id === "string" && a.id ? a.id : i}`,
        name: String(a.name).trim().slice(0, 60),
        instruction: String(a.instruction).trim().slice(0, 4000),
        mode: a.mode === "insert" ? "insert" : "replace",
      })),
  };
}
