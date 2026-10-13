/**
 * Per-add-in details for the shared quality-of-life modules in this folder.
 * Everything else in src/qol is identical in the Excel and Word add-ins – only this file differs.
 */

export const APP = {
  id: "word" as const,
  name: "Zee AI for Word",
  host: "Word",
  /** Where the add-in keeps its settings (used to tell a first install from an update). */
  settingsKey: "zeeai.word.settings.v1",
  /** Prefix for every localStorage key the qol modules use. */
  storagePrefix: "zeeai.word",
};

export interface ReleaseNote {
  version: string;
  items: string[];
}

/** Shown once after an update ("What's new"). Newest first. */
export const RELEASE_NOTES: ReleaseNote[] = [
  {
    version: "1.3.0",
    items: [
      "New “General” scope (next to Selected text / Whole document): describe what you want and Zee finds the right part of the document, selects it and edits it, adds new text in the right place, or answers from the relevant paragraphs.",
      "Live selection chip above the chat box – click × to leave the selection out of your next message.",
      "Press ↑ / ↓ in the chat box to recall earlier prompts; favourite prompts appear as one-click chips.",
      "Every message: copy, edit & resend, regenerate. Esc stops a running answer.",
      "Errors now explain what to do, with Retry and “Copy diagnostics”.",
      "Usage meter: today’s tokens (and estimated cost if you enter your prices in Settings).",
      "Accept or reject only Zee AI’s tracked changes (your own revisions stay untouched).",
      "Your most-used quick actions and templates move to the top of their menus.",
      "Settings → Backup & updates: export / import settings to another PC, and “Reload latest version”.",
      "A banner tells you when a newer version is online, and this “What’s new” note appears after each update.",
    ],
  },
];
