import { describe, expect, it } from "vitest";
import { cleanModelOutput, countWords, normalizeWordText, preserveEdges, truncateMiddle } from "../utils/text";
import { hashString } from "../utils/hash";

describe("normalizeWordText", () => {
  it("converts Word breaks and control characters", () => {
    expect(normalizeWordText("Para one\rPara two\u000bline\u0007cell\u0001")).toBe("Para one\nPara two\nline\tcell");
  });
});

describe("cleanModelOutput", () => {
  it("removes preambles, fences and wrapping quotes", () => {
    expect(cleanModelOutput("Here is the revised text:\n\nBetter sentence.")).toBe("Better sentence.");
    expect(cleanModelOutput("```text\nBetter sentence.\n```")).toBe("Better sentence.");
    expect(cleanModelOutput('"Better sentence."')).toBe("Better sentence.");
  });
  it("keeps quotes when the original was quoted", () => {
    expect(cleanModelOutput('"Quoted."', '"original"')).toBe('"Quoted."');
  });
  it("leaves normal text alone", () => {
    expect(cleanModelOutput("Here we go again: nothing to strip.")).toBe("Here we go again: nothing to strip.");
  });
});

describe("helpers", () => {
  it("preserves leading and trailing whitespace of the original", () => {
    expect(preserveEdges("  old text\n", "new text")).toBe("  new text\n");
  });
  it("counts words in several scripts", () => {
    expect(countWords("It's a well-known fact, 42 times. مرحبا بك")).toBe(8);
  });
  it("truncates in the middle", () => {
    const r = truncateMiddle("a".repeat(100) + "b".repeat(100), 100);
    expect(r.truncated).toBe(true);
    expect(r.text.startsWith("a".repeat(70))).toBe(true);
    expect(r.text.endsWith("b".repeat(30))).toBe(true);
    expect(truncateMiddle("short", 100)).toEqual({ text: "short", truncated: false });
  });
  it("hashes deterministically", () => {
    expect(hashString("C:\\Docs\\Report.docx")).toBe(hashString("C:\\Docs\\Report.docx"));
    expect(hashString("a")).not.toBe(hashString("b"));
    expect(hashString("")).toMatch(/^[0-9a-f]{8}$/);
  });
});
