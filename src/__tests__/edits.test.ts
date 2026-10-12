import { describe, expect, it } from "vitest";
import { applyHunksToText, countOccurrences, escapeWordSearch, headContext, hunksFor, planEdits, splitParagraphs, tailContext } from "../utils/edits";
import { chunkParagraphs, linkCitations, numberParagraphs } from "../utils/chunks";

describe("hunksFor", () => {
  it("replaces only the changed words", () => {
    const before = "The pump was ran at 50 Hz and the intake pressure were stable.";
    const after = "The pump was run at 50 Hz and the intake pressure was stable.";
    const h = hunksFor(before, after)!;
    expect(h.length).toBe(2);
    expect(h[0]).toMatchObject({ del: "ran", ins: "run" });
    expect(applyHunksToText(before, h)).toBe(after);
  });

  it("handles insertions and deletions", () => {
    const before = "Install the ESP in the well and record the data.";
    const after = "Install the new ESP in the well and record data.";
    const h = hunksFor(before, after)!;
    expect(applyHunksToText(before, h)).toBe(after);
  });

  it("falls back to whole-paragraph replacement for heavy rewrites", () => {
    expect(hunksFor("One two three four five.", "Completely different words here now.")).toBeNull();
  });

  it("returns no hunks for identical text and null for tab-separated text", () => {
    expect(hunksFor("same", "same")).toEqual([]);
    expect(hunksFor("a\tb c", "a\tb d")).toBeNull();
  });

  it("works on long paragraphs with repeated words", () => {
    const before = Array.from({ length: 30 }, (_, i) => `Well ${i} produced oil at a steady rate.`).join(" ");
    const after = before.replace("Well 17 produced oil at a steady rate.", "Well 17 produced oil at a declining rate.");
    const h = hunksFor(before, after)!;
    expect(h.length).toBe(1);
    expect(applyHunksToText(before, h)).toBe(after);
  });
});

describe("planEdits", () => {
  it("pairs paragraphs and skips unchanged ones", () => {
    const plan = planEdits(["First line is fine.", "", "Second has a eror here."], "First line is fine.\n\nSecond has an error here.")!;
    expect(plan.length).toBe(1);
    expect(plan[0].index).toBe(1);
    expect(applyHunksToText("Second has a eror here.", plan[0].hunks!)).toBe("Second has an error here.");
  });
  it("returns null when the paragraph count changes", () => {
    expect(planEdits(["a b c", "d e f"], "a b c d e f")).toBeNull();
  });
  it("keeps leading indentation", () => {
    const plan = planEdits(["   Indented paragrah text here."], "Indented paragraph text here.")!;
    expect(plan[0].after).toBe("   Indented paragraph text here.");
  });
});

describe("helpers", () => {
  it("cuts context at word boundaries", () => {
    expect(tailContext("x".repeat(10) + " alpha beta gamma", 12)).toBe(" beta gamma");
    expect(headContext("alpha beta gamma delta", 12)).toBe("alpha beta");
  });
  it("counts and escapes", () => {
    expect(countOccurrences("a a a", "a")).toBe(3);
    expect(escapeWordSearch("x^2")).toBe("x^^2");
    expect(splitParagraphs("a\r\n\r\nb  \n")).toEqual(["a", "b"]);
  });
});

describe("chunks and citations", () => {
  it("chunks on paragraph boundaries and keeps Word numbering", () => {
    const paras = ["aaaa", "", "bbbb", "cccc", "dddd"];
    const c = chunkParagraphs(paras, 10);
    expect(c).toEqual([
      { firstParagraph: 1, text: "aaaa\nbbbb" },
      { firstParagraph: 4, text: "cccc\ndddd" },
    ]);
  });
  it("splits a single huge paragraph", () => {
    const p = "Sentence one is here. ".repeat(50);
    const c = chunkParagraphs([p], 120);
    expect(c.every((x) => x.text.length <= 130)).toBe(true);
    expect(c.length).toBeGreaterThan(5);
  });
  it("numbers paragraphs and links citations", () => {
    expect(numberParagraphs(["Intro", "", "Body"])).toBe("[P1] Intro\n[P3] Body");
    expect(linkCitations("See [P3] and [P4, P7].")).toBe("See [¶3](#zee-para-3) and [¶4](#zee-para-4) [¶7](#zee-para-7).");
    expect(linkCitations("[P2](http://x)")).toBe("[P2](http://x)");
  });
});
