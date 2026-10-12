# Zee AI team proxy (optional)

Use this when several people use Zee AI for Word and you don't want a Gemini API key on every PC.
If you already deployed the proxy for the Zee AI Excel add-in, reuse it – this version also allows
listing models, so **Test connection** can refresh the model list through the proxy.

1. Install Wrangler: `npm i -g wrangler` and run `wrangler login`.
2. In this folder: `wrangler init zee-proxy --yes` (or create a Worker in the Cloudflare dashboard) and use `worker.js` as the code.
3. Add the secrets:
   ```
   wrangler secret put GEMINI_API_KEY
   wrangler secret put TEAM_TOKEN
   ```
   Optionally set the variable `ALLOWED_ORIGIN` to your add-in origin, e.g. `https://<user>.github.io`.
4. `wrangler deploy` → you get a URL like `https://zee-proxy.<account>.workers.dev`.
5. In Zee AI → Settings → **Team proxy**, paste that URL and the team token.

The proxy only forwards `generateContent`, `streamGenerateContent`, `countTokens` and the model list,
and it refuses requests without the token.

### Shared team setup (optional)

Export a team file in Zee AI → Settings → Team setup → **Export my glossary & actions**, then store it on the proxy:

```
wrangler secret put TEAM_CONFIG   # paste the whole JSON
```

Team members choose **Team setup → Load from an address → `proxy`**. The glossary and actions refresh every time the pane opens.
