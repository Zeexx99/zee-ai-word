import { describe, expect, it } from "vitest";
import { normalizeLocate, paraLabel, paraListLabel, relevantExcerpt } from "../utils/locate";

const paras = ["Title", "", "Intro text.", "Pump details.", "More pump.", "", "Conclusion text."];

describe("normalizeLocate", () => {
  it("keeps a valid edit range and trims empty edge paragraphs", () => {
    expect(normalizeLocate({ action: "edit", start: 2, end: 6, reason: "pump" }, paras)).toEqual({ action: "edit", start: 3, end: 5, reason: "pump" });
  });
  it("orders reversed ranges and fills a missing end", () => {
    expect(normalizeLocate({ action: "edit", start: 5, end: 4 }, paras)).toMatchObject({ action: "edit", start: 4, end: 5 });
    expect(normalizeLocate({ action: "EDIT", start: 7 }, paras)).toMatchObject({ action: "edit", start: 7, end: 7 });
  });
  it("falls back to answering when an edit range is unusable", () => {
    expect(normalizeLocate({ action: "edit", start: 99, end: 120, paragraphs: [4] }, paras)).toMatchObject({ action: "answer", paragraphs: [4] });
    expect(normalizeLocate({ action: "edit", start: 2, end: 2 }, paras)).toMatchObject({ action: "none" });
  });
  it("handles inserts (0 = end of document, invalid numbers too)", () => {
    expect(normalizeLocate({ action: "insert", insertAfter: 5 }, paras)).toMatchObject({ action: "insert", after: 5 });
    expect(normalizeLocate({ action: "insert", insertAfter: 0 }, paras)).toMatchObject({ action: "insert", after: 0 });
    expect(normalizeLocate({ action: "insert", insertAfter: 50 }, paras)).toMatchObject({ action: "insert", after: 0 });
    expect(normalizeLocate({ action: "insert", end: 3 }, paras)).toMatchObject({ action: "insert", after: 3 });
  });
  it("answers from existing, non-empty paragraphs only", () => {
    expect(normalizeLocate({ action: "answer", paragraphs: [7, 2, 4, 4, 30] }, paras)).toEqual({ action: "answer", paragraphs: [7, 4], reason: "" });
    expect(normalizeLocate({ action: "answer", paragraphs: [] }, paras)).toMatchObject({ action: "none" });
    expect(normalizeLocate(null, paras)).toMatchObject({ action: "none" });
    expect(normalizeLocate({ action: "none" }, paras)).toMatchObject({ action: "none" });
  });
  it("accepts numbers sent as strings", () => {
    expect(normalizeLocate({ action: "edit", start: "3" as unknown as number, end: "4" as unknown as number }, paras)).toMatchObject({ start: 3, end: 4 });
  });
});

describe("relevantExcerpt", () => {
  it("adds neighbours, keeps real numbers and marks gaps", () => {
    const r = relevantExcerpt(paras, [7, 3], 1);
    expect(r.used).toEqual([3, 4, 7]); // empty neighbours (¶2, ¶6) are skipped
    expect(r.text).toBe("[P3] Intro text.\n[P4] Pump details.\n…\n[P7] Conclusion text.");
  });
  it("respects the size budget, keeping the most relevant pick", () => {
    const big = ["a".repeat(50), "b".repeat(50), "c".repeat(50)];
    const r = relevantExcerpt(big, [3, 1], 0, 70);
    expect(r.used).toEqual([3]);
  });
});

describe("labels", () => {
  it("formats paragraph labels", () => {
    expect(paraLabel(3)).toBe("¶3");
    expect(paraLabel(3, 7)).toBe("¶3–7");
    expect(paraListLabel([9, 2, 3, 4, 12])).toBe("¶2–4, ¶9, ¶12");
    expect(paraListLabel([1, 3, 5, 7, 9, 11, 13], 3)).toBe("¶1, ¶3, ¶5 …");
  });
});
