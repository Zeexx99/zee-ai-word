import { describe, expect, it } from "vitest";
import { buildSections, headingLevel } from "../utils/sections";
import { extractNumbers, langOf, normalizeDigits, numberMismatches } from "../utils/bilingual";
import { blocksForPrompt, diffParagraphs } from "../utils/versionDiff";

describe("sections", () => {
  it("builds nested sections up to the next heading of the same level", () => {
    const items = [
      { text: "Report", level: 1 },
      { text: "intro words here", level: null },
      { text: "Pumps", level: 2 },
      { text: "pump text", level: null },
      { text: "Wells", level: 2 },
      { text: "well text", level: null },
      { text: "", level: null },
      { text: "Annex", level: 1 },
      { text: "x", level: null },
    ];
    const s = buildSections(items);
    expect(s.map((x) => [x.title, x.start, x.end, x.level])).toEqual([
      ["Report", 1, 6, 1],
      ["Pumps", 3, 4, 2],
      ["Wells", 5, 6, 2],
      ["Annex", 8, 9, 1],
    ]);
    expect(s[1].words).toBe(3);
  });
  it("reads heading levels from styles or outline level", () => {
    expect(headingLevel("Heading2", "Heading 2")).toBe(2);
    expect(headingLevel("Title", "")).toBe(1);
    expect(headingLevel("Other", "عنوان 3")).toBe(3);
    expect(headingLevel("Normal", "Normal", 4)).toBe(4);
    expect(headingLevel("Normal", "Normal", "BodyText")).toBeNull();
  });
});

describe("bilingual", () => {
  it("detects language and normalizes Arabic digits", () => {
    expect(langOf("The pump was replaced.")).toBe("en");
    expect(langOf("تم استبدال المضخة.")).toBe("ar");
    expect(langOf("123")).toBe("none");
    expect(normalizeDigits("٣٥٫٥")).toBe("35.5");
    expect(extractNumbers("Rate 1,250.50 bbl/d on 2026, WC ٣٥٪")).toEqual(["1250.5", "2026", "35"]);
  });
  it("pairs EN/AR paragraphs and flags numbers in only one language", () => {
    const paras = ["Oil rate was 1,250 bbl/d at 35% water cut.", "", "كان معدل الزيت ١٢٥٠ برميل/يوم عند نسبة مياه ٣٨٪.", "Pump depth 6,500 ft.", "عمق المضخة ٦٥٠٠ قدم."];
    const r = numberMismatches(paras);
    expect(r.pairs).toBe(2);
    expect(r.mismatches).toEqual([{ en: 1, ar: 3, onlyInEnglish: ["35"], onlyInArabic: ["38"] }]);
  });
});

describe("version diff", () => {
  const oldP = ["Title", "Oil rate 1,200 bbl/d.", "Same text.", "Recommend acid job.", "End."];
  const newP = ["Title", "Oil rate 1,310 bbl/d.", "Same text.", "", "New risk paragraph.", "End."];
  it("finds changed, added and removed paragraphs with current paragraph numbers", () => {
    const b = diffParagraphs(oldP, newP);
    expect(b).toHaveLength(2);
    expect(b[0]).toMatchObject({ kind: "changed", oldText: ["Oil rate 1,200 bbl/d."], newText: ["Oil rate 1,310 bbl/d."], newParagraphs: [2] });
    expect(b[1]).toMatchObject({ kind: "changed", oldText: ["Recommend acid job."], newText: ["New risk paragraph."], newParagraphs: [5] });
  });
  it("reports pure removals next to the following paragraph", () => {
    const b = diffParagraphs(["A", "B", "C"], ["A", "C"]);
    expect(b).toEqual([{ kind: "removed", oldText: ["B"], newText: [], newParagraphs: [2] }]);
    expect(diffParagraphs(["A"], ["A"])).toEqual([]);
  });
  it("formats blocks for the prompt within a budget", () => {
    const r = blocksForPrompt(diffParagraphs(oldP, newP), 80);
    expect(r.truncated).toBe(true);
    expect(r.text).toContain("Change 1 (changed) [P2]");
  });
});
