import { describe, expect, it } from "vitest";
import { diffStats, diffWords, rebuild, tokenize } from "../utils/diff";

describe("tokenize", () => {
  it("keeps every character", () => {
    const s = "Hello,  world! It's 3.5% — fine.\nNew line";
    expect(tokenize(s).join("")).toBe(s);
  });
  it("handles Arabic and accented words", () => {
    const s = "مرحبا بالعالم café";
    expect(tokenize(s).join("")).toBe(s);
    expect(tokenize(s).filter((t) => t.trim()).length).toBe(3);
  });
});

describe("diffWords", () => {
  it("returns a single equal block for identical text", () => {
    expect(diffWords("same text", "same text")).toEqual([{ type: "eq", text: "same text" }]);
  });

  it("marks a replaced word", () => {
    const ops = diffWords("The quick brown fox", "The quick red fox");
    expect(rebuild(ops, "before")).toBe("The quick brown fox");
    expect(rebuild(ops, "after")).toBe("The quick red fox");
    expect(ops).toContainEqual({ type: "del", text: "brown" });
    expect(ops).toContainEqual({ type: "ins", text: "red" });
  });

  it("rebuilds both sides for multi-paragraph edits", () => {
    const a = "First paragraph has a eror.\n\nSecond one is fine.\nThird line here.";
    const b = "First paragraph has an error.\n\nSecond one is fine.\nThird line is here, too.";
    const ops = diffWords(a, b);
    expect(rebuild(ops, "before")).toBe(a);
    expect(rebuild(ops, "after")).toBe(b);
  });

  it("handles empty sides", () => {
    expect(diffWords("", "new")).toEqual([{ type: "ins", text: "new" }]);
    expect(diffWords("old", "")).toEqual([{ type: "del", text: "old" }]);
    expect(diffWords("", "")).toEqual([]);
  });

  it("falls back to one replace block when the edit budget is exceeded", () => {
    const ops = diffWords("a b c d e f", "u v w x y z", 2);
    expect(ops).toEqual([
      { type: "del", text: "a b c d e f" },
      { type: "ins", text: "u v w x y z" },
    ]);
  });

  it("is fast enough for long texts", () => {
    const words = Array.from({ length: 3000 }, (_, i) => `word${i % 97}`);
    const a = words.join(" ");
    const b = words.map((w, i) => (i % 50 === 0 ? "changed" : w)).join(" ");
    const t = Date.now();
    const ops = diffWords(a, b);
    expect(Date.now() - t).toBeLessThan(2000);
    expect(rebuild(ops, "after")).toBe(b);
  });
});

describe("diffStats", () => {
  it("counts added and removed words", () => {
    const s = diffStats(diffWords("one two three", "one 2 three four"));
    expect(s.added).toBe(2);
    expect(s.removed).toBe(1);
    expect(s.kept).toBeCloseTo(2 / 3);
  });
});
