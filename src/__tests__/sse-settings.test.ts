import { describe, expect, it } from "vitest";
import { SseParser } from "../utils/sse";
import { DEFAULT_SETTINGS, STORAGE_KEY, isConfigured, sanitizeSettings } from "../services/settingsService";
import { filterModelIds } from "../services/geminiService";

describe("SseParser", () => {
  it("parses data lines across chunk boundaries", () => {
    const p = new SseParser();
    expect(p.push('data: {"a":1}\r\n\r\ndata: {"b"')).toEqual(['{"a":1}']);
    expect(p.push(':2}\n\n')).toEqual(['{"b":2}']);
    expect(p.push("data: [DONE]\n")).toEqual([]);
    p.push('data: {"c":3}');
    expect(p.flush()).toEqual(['{"c":3}']);
  });
});

describe("settings", () => {
  it("uses the Word storage key and Gemini 3.8 Flash by default", () => {
    expect(STORAGE_KEY).toBe("zeeai.word.settings.v1");
    expect(DEFAULT_SETTINGS.model).toBe("gemini-3.8-flash");
    expect(DEFAULT_SETTINGS.costGuardTokens).toBe(20000);
  });
  it("repairs invalid stored values", () => {
    const s = sanitizeSettings({ theme: "neon" as never, authMode: "x" as never, costGuardTokens: -5, model: "" });
    expect(s.theme).toBe("auto");
    expect(s.authMode).toBe("apiKey");
    expect(s.costGuardTokens).toBe(20000);
    expect(s.model).toBe("gemini-3.8-flash");
  });
  it("knows when a connection is configured", () => {
    expect(isConfigured(DEFAULT_SETTINGS)).toBe(false);
    expect(isConfigured({ ...DEFAULT_SETTINGS, apiKey: "AIza" })).toBe(true);
    expect(isConfigured({ ...DEFAULT_SETTINGS, authMode: "proxy", proxyUrl: "https://p" })).toBe(true);
    expect(isConfigured({ ...DEFAULT_SETTINGS, authMode: "google", googleClientId: "id" })).toBe(false);
  });
});

describe("filterModelIds", () => {
  it("keeps text Gemini models only", () => {
    expect(
      filterModelIds({
        models: [
          { name: "models/gemini-3.8-flash", supportedGenerationMethods: ["generateContent"] },
          { name: "models/gemini-embedding-001", supportedGenerationMethods: ["embedContent"] },
          { name: "models/gemini-3.1-flash-image", supportedGenerationMethods: ["generateContent"] },
          { name: "models/gemini-3.5-flash-lite", supportedGenerationMethods: ["generateContent", "countTokens"] },
        ],
      })
    ).toEqual(["gemini-3.5-flash-lite", "gemini-3.8-flash"]);
  });
});
