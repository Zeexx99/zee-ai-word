# Shared Zee AI modules

These files are **host-agnostic** (no Word or Excel APIs) and are meant to be identical in every
Zee AI add-in:

| File | What it does |
|---|---|
| `gemini.ts` | Gemini REST client: key / Vertex AI / team proxy, SSE streaming, JSON mode, files & images, retries, response cache, model list, connection test |
| `auth.ts` | Google sign-in through the Office dialog (call `configureAuth("zeeai.word", "Word")` once per add-in) |
| `sse.ts` | Incremental Server-Sent-Events parser |
| `hash.ts` | Small stable string hash |

Edit them here, then copy them to the other add-in:

```bat
node scripts\sync-shared.mjs D:\ADD_IN_CLAUDE --check   :: what differs
node scripts\sync-shared.mjs D:\ADD_IN_CLAUDE           :: copy
```

The Excel add-in still has its own copies in `src/services` and `src/utils`; switching it to import
from `src/shared` is a separate, deliberate step (its agent mode needs a few extra exports).
