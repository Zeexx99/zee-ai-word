import { describe, expect, it } from "vitest";
import { alignmentFor, detectDirection, hasRtl, inspectParagraphXml, setParagraphDirectionXml } from "../utils/direction";
import { markdownToHtml, plainToHtml } from "../utils/mdToHtml";
import { sanitizeSettings } from "../services/settingsService";
import { bilingualPrompt } from "../services/prompts";

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

const pkg = (p: string) =>
  `<pkg:package><pkg:part><pkg:xmlData><w:document><w:body>${p}<w:sectPr/></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>`;
const para = (x: string | null) => (x === null ? null : x.replace(/.*<w:body>|<w:sectPr\/>.*/g, ""));

describe("paragraph OOXML direction", () => {
  it("adds bidi to an Arabic paragraph without pPr", () => {
    expect(para(setParagraphDirectionXml(pkg("<w:p><w:r><w:t>نص</w:t></w:r></w:p>"), "rtl"))).toBe("<w:p><w:pPr><w:bidi/></w:pPr><w:r><w:t>نص</w:t></w:r></w:p>");
  });
  it("puts bidi in schema order and drops the old left alignment", () => {
    const x = pkg('<w:p><w:pPr><w:pStyle w:val="Normal"/><w:spacing w:after="0"/><w:jc w:val="left"/><w:rPr><w:b/></w:rPr></w:pPr><w:r><w:t>نص</w:t></w:r></w:p>');
    expect(para(setParagraphDirectionXml(x, "rtl"))).toBe('<w:p><w:pPr><w:pStyle w:val="Normal"/><w:bidi/><w:spacing w:after="0"/><w:rPr><w:b/></w:rPr></w:pPr><w:r><w:t>نص</w:t></w:r></w:p>');
  });
  it("keeps justified and centres headings", () => {
    const j = pkg('<w:p><w:pPr><w:jc w:val="both"/></w:pPr><w:r><w:t>نص</w:t></w:r></w:p>');
    expect(para(setParagraphDirectionXml(j, "rtl"))).toBe('<w:p><w:pPr><w:bidi/><w:jc w:val="both"/></w:pPr><w:r><w:t>نص</w:t></w:r></w:p>');
    const h = pkg('<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>عنوان</w:t></w:r></w:p>');
    expect(para(setParagraphDirectionXml(h, "rtl", true))).toBe('<w:p><w:pPr><w:pStyle w:val="Heading1"/><w:bidi/><w:jc w:val="center"/></w:pPr><w:r><w:t>عنوان</w:t></w:r></w:p>');
  });
  it("turns an English paragraph back to left-to-right", () => {
    const x = pkg('<w:p><w:pPr><w:bidi/><w:jc w:val="right"/></w:pPr><w:r><w:t>Text</w:t></w:r></w:p>');
    expect(para(setParagraphDirectionXml(x, "ltr"))).toBe("<w:p><w:r><w:t>Text</w:t></w:r></w:p>");
  });
  it("returns null when nothing changes", () => {
    expect(setParagraphDirectionXml(pkg("<w:p><w:pPr><w:bidi/></w:pPr><w:r><w:t>نص</w:t></w:r></w:p>"), "rtl")).toBeNull();
    expect(setParagraphDirectionXml(pkg("<w:p><w:r><w:t>Text</w:t></w:r></w:p>"), "ltr")).toBeNull();
  });
  it("flags paragraphs that are unsafe to rewrite", () => {
    expect(inspectParagraphXml(pkg("<w:p><w:pPr><w:bidi/></w:pPr></w:p>"))).toEqual({ bidi: true, risky: false });
    expect(inspectParagraphXml(pkg('<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/></w:numPr></w:pPr><w:r><w:t>x</w:t></w:r></w:p>'))?.risky).toBe(true);
    expect(inspectParagraphXml(pkg('<w:p><w:commentRangeStart w:id="1"/><w:r><w:t>x</w:t></w:r></w:p>'))?.risky).toBe(true);
    expect(inspectParagraphXml(pkg("<w:sdt><w:sdtContent><w:p><w:r><w:t>x</w:t></w:r></w:p></w:sdtContent></w:sdt>"))?.risky).toBe(true);
  });
});

describe("Arabic font and bilingual tables", () => {
  it("uses the Arabic font for RTL blocks only", () => {
    const html = markdownToHtml("نص عربي\n\nEnglish", { autoDirection: true, fontFamily: "Aptos", arabicFont: "Simplified Arabic" });
    expect(html).toContain(`<p dir="rtl" style="font-family:'Simplified Arabic';direction:rtl;text-align:right">`);
    expect(html).toContain(`<p dir="ltr" style="font-family:'Aptos';text-align:left">English</p>`);
  });
  it("keeps an English | Arabic table left-to-right with per-cell direction", () => {
    const html = markdownToHtml("| English | العربية |\n|---|---|\n| Pump installed | تم تركيب المضخة |", { autoDirection: true });
    expect(html).toContain('<table style="border-collapse:collapse">');
    expect(html).toContain('dir="rtl" style="border:1px solid #999;padding:3pt 6pt;direction:rtl;text-align:right"');
  });
  it("builds both bilingual layouts", () => {
    expect(bilingualPrompt("Hello", "table")).toContain("| English | العربية |");
    expect(bilingualPrompt("Hello", "stacked")).toContain('"---"');
  });
});
