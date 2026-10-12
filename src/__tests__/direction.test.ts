import { describe, expect, it } from "vitest";
import { alignmentFor, detectDirection, hasRtl } from "../utils/direction";
import { markdownToHtml, plainToHtml } from "../utils/mdToHtml";
import { sanitizeSettings } from "../services/settingsService";

describe("detectDirection", () => {
  it("detects Arabic and English paragraphs", () => {
    expect(detectDirection("تم الانتهاء من اختبار البئر بنجاح")).toBe("rtl");
    expect(detectDirection("The well test was completed successfully")).toBe("ltr");
  });
  it("keeps technical Arabic with English terms right-to-left", () => {
    expect(detectDirection("تم تركيب مضخة ESP بقدرة 150 HP في البئر W-486")).toBe("rtl");
    expect(detectDirection("ESP تم تركيب مضخة جديدة")).toBe("rtl");
  });
  it("keeps English with one Arabic word left-to-right", () => {
    expect(detectDirection("The client name is شركة in the report header today")).toBe("ltr");
  });
  it("returns null without letters", () => {
    expect(detectDirection("")).toBeNull();
    expect(detectDirection("1,250 – 35% (2026)")).toBeNull();
  });
  it("ignores link URLs and inline code", () => {
    expect(detectDirection("[الرابط](https://example.com/very/long/english/path) هنا")).toBe("rtl");
  });
  it("helpers", () => {
    expect(hasRtl("abc")).toBe(false);
    expect(hasRtl("abc د")).toBe(true);
    expect(alignmentFor("rtl")).toBe("Right");
    expect(alignmentFor("ltr")).toBe("Left");
  });
});

describe("markdownToHtml with autoDirection", () => {
  it("is unchanged when autoDirection is off", () => {
    expect(markdownToHtml("# Title\n\nنص عربي")).toBe("<h1>Title</h1><p>نص عربي</p>");
  });
  it("centres headings and sets paragraph direction", () => {
    const html = markdownToHtml("# تقرير يومي\n\nنص عربي هنا\n\nEnglish text here", { autoDirection: true });
    expect(html).toContain('<h1 dir="rtl" style="direction:rtl;text-align:center">تقرير يومي</h1>');
    expect(html).toContain('<p dir="rtl" style="direction:rtl;text-align:right">نص عربي هنا</p>');
    expect(html).toContain('<p dir="ltr" style="text-align:left">English text here</p>');
  });
  it("keeps the font with direction", () => {
    const html = markdownToHtml("نص", { autoDirection: true, fontFamily: "Arial", fontSizePt: 12 });
    expect(html).toBe(`<p dir="rtl" style="font-family:'Arial';font-size:12pt;direction:rtl;text-align:right">نص</p>`);
  });
  it("sets list and table direction", () => {
    const list = markdownToHtml("- البند الأول\n- البند الثاني", { autoDirection: true });
    expect(list.startsWith('<ul dir="rtl">')).toBe(true);
    expect(list).toContain('<li dir="rtl" style="direction:rtl;text-align:right">');
    const table = markdownToHtml("| البئر | الإنتاج |\n|---|---|\n| W-1 | 500 |", { autoDirection: true });
    expect(table).toContain('<table dir="rtl" style="border-collapse:collapse;direction:rtl">');
  });
  it("puts Arabic block quotes on the right", () => {
    expect(markdownToHtml("> ملاحظة مهمة", { autoDirection: true })).toContain("margin-right:18pt");
  });
});

describe("plainToHtml", () => {
  it("gives each line its own direction", () => {
    expect(plainToHtml("سطر عربي\n\nEnglish line")).toBe(
      '<p dir="rtl" style="direction:rtl;text-align:right">سطر عربي</p><p dir="ltr" style="text-align:left">English line</p>'
    );
  });
  it("escapes HTML", () => {
    expect(plainToHtml("a < b")).toBe('<p dir="ltr" style="text-align:left">a &lt; b</p>');
  });
});

describe("autoDirection setting", () => {
  it("defaults to on and respects off", () => {
    expect(sanitizeSettings(null).autoDirection).toBe(true);
    expect(sanitizeSettings({ autoDirection: false }).autoDirection).toBe(false);
  });
});
