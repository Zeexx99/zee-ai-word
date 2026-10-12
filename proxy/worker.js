/**
 * Zee AI team proxy – Cloudflare Worker (the free tier is plenty).
 * Keeps the Gemini API key on the server; add-in users only get a shared team token.
 * The same worker serves both the Excel and the Word add-in.
 *
 * Secrets (wrangler secret put …):
 *   GEMINI_API_KEY  – your Google AI Studio key
 *   TEAM_TOKEN      – any long random string you give to your team (Settings → Team proxy)
 * Optional variables:
 *   ALLOWED_ORIGIN  – e.g. https://<user>.github.io (defaults to *)
 *   TEAM_CONFIG     – a Zee AI team file (JSON from Settings → Team setup → Export). Served at /zee-config
 *                     so everyone gets the same glossary and quick actions (Settings → Team setup → "proxy").
 */
const UPSTREAM = "https://generativelanguage.googleapis.com";

export default {
  async fetch(request, env) {
    const origin = env.ALLOWED_ORIGIN || "*";
    const cors = {
      "Access-Control-Allow-Origin": origin,
      "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
      "Access-Control-Allow-Headers": "Authorization, Content-Type",
      "Access-Control-Max-Age": "86400",
    };
    if (request.method === "OPTIONS") return new Response(null, { status: 204, headers: cors });
    const json = (status, message) =>
      new Response(JSON.stringify({ error: { code: status, message } }), { status, headers: { ...cors, "Content-Type": "application/json" } });

    const url = new URL(request.url);
    // Only model calls (/v1beta/models/<model>:<method>) and the model list (/v1beta/models).
    const isModelCall = /^\/v1beta\/models\/[\w.\-]+:(generateContent|streamGenerateContent|countTokens)$/.test(url.pathname);
    const isList = request.method === "GET" && url.pathname === "/v1beta/models";
    const isConfig = request.method === "GET" && url.pathname === "/zee-config";
    if (!isModelCall && !isList && !isConfig) return json(404, "Not allowed");

    const auth = request.headers.get("Authorization") || "";
    if (!env.TEAM_TOKEN || auth !== `Bearer ${env.TEAM_TOKEN}`) return json(401, "Invalid team token");

    if (isConfig) {
      if (!env.TEAM_CONFIG) return json(404, "No TEAM_CONFIG set on the proxy");
      return new Response(env.TEAM_CONFIG, { status: 200, headers: { ...cors, "Content-Type": "application/json", "Cache-Control": "no-store" } });
    }

    const upstream = new URL(url.pathname + url.search, UPSTREAM);
    const res = await fetch(upstream, {
      method: request.method,
      headers: { "Content-Type": "application/json", "x-goog-api-key": env.GEMINI_API_KEY },
      body: request.method === "POST" ? request.body : undefined,
    });
    const headers = new Headers(res.headers);
    Object.entries(cors).forEach(([k, v]) => headers.set(k, v));
    return new Response(res.body, { status: res.status, headers });
  },
};
