/** Loading / exporting the shared team configuration (glossary + quick actions). */
import { buildTeamConfig, parseTeamConfig, type TeamConfig } from "../utils/teamConfig";
import type { ZeeSettings } from "./settingsService";

/** "proxy" = the team proxy's /zee-config endpoint; otherwise any https URL (e.g. a raw GitHub file). */
export async function fetchTeamConfig(source: string, s: ZeeSettings): Promise<TeamConfig> {
  let url = source.trim();
  const headers: Record<string, string> = {};
  if (url === "proxy") {
    const base = s.proxyUrl.trim().replace(/\/+$/, "");
    if (!base) throw new Error("Set up the team proxy in Connection first.");
    url = `${base}/zee-config`;
  }
  if (!/^https:\/\//i.test(url)) throw new Error("The team config address must start with https://");
  const proxyBase = s.proxyUrl.trim().replace(/\/+$/, "");
  if (proxyBase && url.startsWith(proxyBase) && s.proxyToken.trim()) headers.Authorization = `Bearer ${s.proxyToken.trim()}`;
  let res: Response;
  try {
    res = await fetch(url, { headers, cache: "no-store" });
  } catch {
    throw new Error("Couldn't reach the team config address (network or CORS). Raw GitHub links and the team proxy work.");
  }
  if (!res.ok) throw new Error(`The team config address answered ${res.status}.`);
  return parseTeamConfig(await res.text());
}

/** The user's own glossary and actions as a shareable team file. */
export function exportTeamConfig(s: ZeeSettings, name: string): string {
  return JSON.stringify(buildTeamConfig(name, s.protectedTerms, s.customActions), null, 2);
}

/** Offers the JSON as a download (falls back to the clipboard where downloads are blocked). */
export function downloadText(filename: string, text: string): boolean {
  try {
    const url = URL.createObjectURL(new Blob([text], { type: "application/json" }));
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 5000);
    return true;
  } catch {
    return false;
  }
}
