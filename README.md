# Zee AI for Word – Gemini writing copilot for Microsoft Word

Zee AI for Word is a Word task-pane add-in for Windows, Mac and Word on the web. It is built with React 18, TypeScript, Vite, Fluent UI v9 and Office.js, and powered by Google Gemini. It is the Word sibling of the Zee AI Excel add-in and follows the same architecture, security rules and scripts.

## Features

| | |
|---|---|
| **Rewrite & polish** | Improves clarity, flow, sentence variety and word choice in the selection. |
| **Grammar & spelling** | Proofreads the selection in JSON mode: corrected text plus a list of every fix with a reason. Apply the fixes or **add them as Word comments**. With nothing selected it proofreads the **whole document** in parts and lists the issues for you to add as comments. |
| **Tone** | Formal, Executive, Academic, Casual, Persuasive or Technical. |
| **Summarize** | Bullet points, an executive summary or a TL;DR of the selection. With nothing selected it summarizes the whole document. Long documents are read in parts and then combined, with clickable **[¶12]** citations. |
| **Length** | Expand (about 2×) or shorten (about ½). |
| **Translate** | 16 languages. Your last choice is remembered. |
| **Draft** | Writes new content from a free-form request, such as "Draft an NDA non-compete clause" or "Generate an introduction section". You can set length and tone and use the selection or the start of the document as context. |
| **My actions** | Your own one-click actions ("Mine" in the toolbar), for example *Well test summary* or *Client-ready tone*. Each one either edits the selection or writes new text after it. Add and edit them in Settings. |
| **Tables** | With the cursor in a table you can *explain* it, *write a commentary* below it, or *clean it up*. Clean-up fixes spelling, units, capitalisation and date formats, changes only the affected cells and keeps their formatting. A safety check discards any proposed change that would alter a number. *Turn notes into a table* builds a table from selected text. |
| **Consistency check** | Review → Consistency check scans the whole document for mismatched units (bbl/d vs BOPD), names (W-12 vs Well 12), numbers, dates, abbreviations and terms. Each finding points to its paragraph and can become a Word comment. |
| **Check against a reference** | Review → Check against a reference… compares the document with a procedure, standard or earlier report (PDF, DOCX, TXT, MD or CSV). It lists contradictions, missing requirements and outdated values, which can become comments. |
| **Pictures** | Select a picture or chart to write a caption (inserted below it in Word's Caption style), generate alt text, or describe and analyse a chart. Gemini reads the image itself. |
| **Report templates** | Daily production report, ESP failure analysis (RCA), well test report, MOC, technical memo and meeting minutes. Each is inserted with Word heading styles, so References → Table of Contents works. With notes selected, AI fills the template in from them. |
| **Ribbon, right-click & shortcuts** | Home → Zee AI → **Quick actions**, the **Zee AI** entry in the right-click menu, and Ctrl+Alt+Shift+**R** (rewrite), **G** (grammar), **S** (summarize), **K** (consistency) and **Z** (open the pane). They open the pane with the usual preview, and nothing changes without a click. These need a Word version with shared-runtime support (Microsoft 365). |
| **Cheaper bulk work** | Whole-document proofreading and long summaries use Flash-Lite (this can be switched off in Settings → Cost). Repeated requests are answered from a session cache at no cost. |
| **Team setup** | Settings → Team setup exports your glossary and quick actions as a team file. Colleagues import that file or load it from a URL or the team proxy (`TEAM_CONFIG`). It refreshes every time the pane opens. |
| **Chat** | Streaming answers (`streamGenerateContent?alt=sse`). The chat can read nothing, the selection or the whole document. Long documents are trimmed in the middle. History is saved **per document**. |
| **Track Changes** | The header toggle records Zee AI's edits as Word revisions you accept or reject in the Review tab. Needs Word for Microsoft 365 / 2021+. Your own Track Changes setting is restored afterwards. |
| **Keeps formatting** | Edits change only the words that differ, so bold words, links, styles and comments in the rest of the text stay. If a paragraph changed too much, only that paragraph's text is replaced and its style is kept. |
| **Protected terms** | A glossary (petroleum defaults: ESP, TDH, PIP, GOR, bbl/d, psi…) that no action, translation or chat answer may rewrite, translate or "correct". |
| **Citations** | When the chat reads the whole document, answers cite paragraphs as [¶12]. Click one to jump to that paragraph. |
| **Diff preview** | Edits show **Changes / New text / Original** tabs before you apply them. Deletions are struck through in red and insertions are underlined in green. Text is never overwritten silently. |
| **Smart insertion** | **Replace selection** (`Word.InsertLocation.replace`), **Insert after cursor** (`after`) or **Append to end** (`end`). Markdown answers (headings, bold, lists, tables) are inserted as real Word formatting in the font at the cursor. |
| **Multi-step undo** | ↶ reverts Zee AI's insertions, edits and comments one step at a time, up to 25 steps. Edits are undone word by word, so formatting comes back too. Word's own Ctrl+Z works as well. |
| **Cost guard** | Asks before sending a prompt above the token limit (default 20,000). The pane shows prompt and completion tokens per answer, a Σ session total and a breakdown in Settings. |
| **Auth** | Gemini API key, Google sign-in (Vertex AI quota) or a **team proxy** that keeps the key on a server (`proxy/`). |

### Security rules

* No backend is needed. The pane calls `https://generativelanguage.googleapis.com/v1beta` directly and sends the key in the `x-goog-api-key` header.
* Settings (key, model, theme, temperature…) live in `localStorage` under the add-in's own origin, key **`zeeai.word.settings.v1`**.
* The API key is **never** written into the document or into source code or git.
* Chat history is also stored in `localStorage`, keyed by a hash of the document's path. A never-saved document gets a tiny random id in its document settings. That id is never a key or a setting.

---

## Option A – Hosted on GitHub Pages (recommended, the same as the Excel add-in)

1. Run **`setup.bat`** once. It installs packages, creates `package-lock.json` and checks that everything builds.
2. Double-click **`publish-github.bat`**. It will:
   - ask for your GitHub username and open GitHub, where you create an empty **public** repo named `zee-ai-word`
   - generate `manifest.prod.xml` for `https://<you>.github.io/zee-ai-word`
   - run the type check, tests and build, and upload only if they pass (a GitHub sign-in window may appear)
   - open **Settings → Pages**, where you set **Source = GitHub Actions**
3. Wait for the **Actions** tab to show a green check (about 2 minutes).
4. Double-click **`use-online.bat`**. It checks that the site is up, registers the hosted manifest with Word and clears Office's cache.
5. Close every Word window, reopen Word and click **Home → Zee AI**.

To ship changes later, run **`update-zee.bat`**. It builds, tests, pushes and clears the cache. To install on another PC with no Node.js needed, run **`install-on-another-pc.bat`**. If your GitHub username isn't `zeexx99`, edit `GHUSER` in that file first.

> On a free GitHub plan the repo must be public. The code contains no keys: keys are typed into Settings and stay on each PC.

## Option B – Local (`https://localhost:3200`)

```bat
setup.bat        :: first time: npm install + localhost certificate + build + open Word
serve-zee.bat    :: every time: start the local server (keep its window open); renews the certificate
start-zee.bat    :: rebuild after code changes, restart the server and open Word
diagnose.bat     :: writes diag-log.txt if the pane doesn't load
```

The add-in uses port **3200**. Port 3000 is often taken by other dev servers and 3100 belongs to the Zee AI Excel add-in, so both add-ins can run at the same time.

---

## Connect Gemini

Open **⚙ Settings** in the pane.

- **API key**: paste a key from <https://aistudio.google.com/apikey>, then click **Test connection & refresh models**. This validates the key and loads every model it can use.
- **Google sign-in (Vertex AI)**:
  1. Enable the Vertex AI API.
  2. Create a *Web application* OAuth client. Set the JavaScript origin to your add-in origin (`https://localhost:3200` or `https://<you>.github.io`) and the redirect URI to `<add-in URL>/auth/callback.html`.
  3. Give your account the *Vertex AI User* role, then sign in.
- **Team proxy**: deploy `proxy/worker.js` (see `proxy/README.md`), then paste its URL and the team token. The Excel add-in's proxy works for Word too.

Models: **Gemini 3.8 Flash** (default, for fast drafting and editing), **Gemini 3.1 Pro preview** (deep reasoning and long-form drafting) and **Gemini 3.5 Flash-Lite** (fastest and cheapest, for bulk summaries and quick checks). Use ⟳ to load all models your key can use.

## Project structure

```
manifest.xml                 Word manifest (Host "Document", https://localhost:3200)
scripts/make-manifest.mjs    Generates manifest.prod.xml for the hosted URL (separate add-in Id)
deploy/deploy.yml            GitHub Pages workflow (copied to .github/workflows by publish-github.bat)
public/                      icons, ribbon commands page, Google sign-in dialog pages
proxy/                       Optional Cloudflare Worker team proxy
taskpane.html                Pane entry + early error capture (shows start-up errors instead of a blank pane)
src/
  main.tsx                   Boot: Office.onReady + 8 s safety net, error boundary, stale-build reload
  App.tsx                    UI orchestration: quick actions, insert/replace flow, confirmations, undo
  theme.ts / styles.css      Fluent themes (light/dark/auto following Office) and layout
  hooks/useChat.ts           Streaming chat, selection actions, grammar JSON mode, cost guard, history
  components/
    QuickActions.tsx         Rewrite · Grammar · Tone · Summarize · Length · Translate · Draft
    MessageBubble.tsx        Markdown answers, diff preview, grammar findings, insert buttons
    DiffView.tsx             Word-level diff (Changes / New text / Original)
    CustomActionsEditor.tsx  Add/edit "My actions"
    ChatInput.tsx / ChatFeed.tsx / Dialogs.tsx / SettingsDrawer.tsx / insertActions.ts
  services/
    wordService.ts           Office.js: read/insert, word-level edits, Track Changes, comments, paragraph jump, undo
    geminiService.ts         REST client: SSE streaming, JSON mode, retries, key / Vertex / proxy
    settingsService.ts       Settings in localStorage (zeeai.word.settings.v1), models
    chatStore.ts             Per-document chat history
    authService.ts / prompts.ts / promptStore.ts
  utils/                     Pure, unit-tested helpers: diff, edit planner, chunks/citations, markdown→HTML, text, SSE, history
  __tests__/                 Vitest tests (run by every build)
```

## Shared code with the Excel add-in

`src/shared/` holds the host-agnostic modules: the Gemini client, Google sign-in, the SSE parser and the hash. Edit them there, then sync them to the other add-in:

```bat
node scripts\sync-shared.mjs D:\ADD_IN_CLAUDE --check   :: list differences
node scripts\sync-shared.mjs D:\ADD_IN_CLAUDE           :: copy src\shared into the Excel project
```

See `src/shared/README.md`. The Excel add-in still imports its own copies until it is switched over.

## Lessons carried over from the Excel add-in

| Problem seen in the Excel add-in | How this project prevents it |
|---|---|
| Port 3000 taken by another app | Port 3200 with `strictPort`. `start-zee.bat` only stops a server started from **this** folder and tells you if another app holds the port. |
| Blank or white pane | Layout classes are never put on `FluentProvider`. Errors are caught early in `taskpane.html`, there's an error boundary, and an 8 s fallback if `Office.onReady` stalls. |
| Old version kept loading | No lazy-loaded chunks. A stale build reloads itself once. `update-zee.bat` and `use-online.bat` clear the Office `Wef` cache. The build date shows in Settings → About. |
| `createRoot` called twice warning | The root is reused if `main.tsx` runs again. |
| "Keyborg instance disposed incorrectly" noise | No `React.StrictMode` double mount. The message is also filtered as benign. |
| `setup.bat` had LF line endings | Every `.bat` / `.ps1` uses CRLF, and `.gitattributes` keeps it that way. |
| "LF will be replaced by CRLF" git warnings | `.gitattributes` normalises line endings. |
| `office-addin-dev-settings` used but not installed | It's a direct devDependency, called with `npx --no` so nothing is downloaded on the fly. |
| GitHub build needs `package-lock.json` | `publish-github.bat` runs `npm install` first so the lock file always matches, and the workflow falls back to `npm install` if it's missing. |
| Slow 4-minute install, deprecated-package warnings | `office-addin-debugging`, which brought in the Teams CLI and `keytar`, was removed. The scripts only need `office-addin-dev-settings`. |
| Heavy deploys (≈14 MB of source maps) | Source maps are off for builds. Set `ZEE_SOURCEMAP=1` to get them back for debugging. |
| Upload of a broken build | `publish-github.bat` and `update-zee.bat` only upload after the type check, tests and build pass. |
| Chat sluggish while streaming | History is saved with a debounce, and streamed text is batched per animation frame. |
| Clipboard and links blocked in the pane | `execCommand` fallback for copy. Links open with `Office.context.ui.openBrowserWindow`. |

## Troubleshooting

| Symptom | Fix |
|---|---|
| "We can't load the add-in" (local) | Run `serve-zee.bat` and keep it open, then close and reopen Word. |
| "We can't load the add-in" (hosted) | Check that the Actions deploy succeeded and that `https://<you>.github.io/zee-ai-word/taskpane.html` opens in a browser. |
| Old version keeps loading | Close Word and run `update-zee.bat`, or delete `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`. |
| ↶ says the text was moved or deleted | Use Word's own Undo (Ctrl+Z). |
| Inserted text uses a different font | Zee AI copies the font at the cursor. If the selection mixes fonts, Word's paragraph default is used. |
| Port 3200 in use | `diagnose.bat` shows which program holds it. |
| `npm warn allow-scripts …` during install | Harmless. It comes from `office-addin-debugging`'s dependencies, and the build doesn't need those scripts. |
