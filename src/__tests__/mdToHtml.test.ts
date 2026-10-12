import { describe, expect, it } from "vitest";
import { escapeHtml, inline, looksLikeMarkdown, markdownToHtml, markdownToPlain } from "../utils/mdToHtml";

describe("inline", () => {
  it("formats bold, italic, code and links", () => {
    expect(inline("**b** and *i* and `x<y` [site](https://a.com)")).toBe(
      `<b>b</b> and <i>i</i> and <code style="font-family:Consolas,'Courier New',monospace">x&lt;y</code> <a href="https://a.com">site</a>`
    );
  });
  it("does not treat snake_case or math as italic", () => {
    expect(inline("use file_name_here and 2*3*4")).toBe("use file_name_here and 2*3*4");
  });
});

describe("markdownToHtml", () => {
  it("escapes HTML from the model", () => {
    expect(markdownToHtml("<script>alert(1)</script>")).toBe("<p>&lt;script&gt;alert(1)&lt;/script&gt;</p>");
    expect(escapeHtml(`"&'`)).toBe("&quot;&amp;'");
  });

  it("renders headings and paragraphs with the body font", () => {
    const html = markdownToHtml("# Title\n\nFirst line\nsecond line\n\nNext", { fontFamily: "Calibri", fontSizePt: 11 });
    expect(html).toBe(
      `<h1>Title</h1><p style="font-family:'Calibri';font-size:11pt">First line<br>second line</p><p style="font-family:'Calibri';font-size:11pt">Next</p>`
    );
  });

  it("renders nested and mixed lists", () => {
    const html = markdownToHtml("- a\n- b\n  - b1\n  - b2\n- c\n\n1. one\n2. two");
    expect(html).toBe("<ul><li>a</li><li>b<ul><li>b1</li><li>b2</li></ul></li><li>c</li></ul><ol><li>one</li><li>two</li></ol>");
  });

  it("renders GFM tables", () => {
    const html = markdownToHtml("| A | B |\n|---|---:|\n| 1 | 2 |\n| 3 |");
    expect(html).toContain("<table");
    expect(html).toContain(">A</span></th>");
    expect((html.match(/<tr>/g) ?? []).length).toBe(3);
    expect(html).toContain('<td style="border:1px solid #999;padding:3pt 6pt"><span></span></td>'); // missing cell padded
  });

  it("renders code fences verbatim", () => {
    expect(markdownToHtml("```\n<b>x</b>\n```")).toContain("&lt;b&gt;x&lt;/b&gt;</pre>");
  });

  it("renders quotes and rules", () => {
    expect(markdownToHtml("> quoted\n\n---")).toBe(`<blockquote style="margin-left:18pt;font-style:italic"><p>quoted</p></blockquote><hr>`);
  });
});

describe("markdownToPlain", () => {
  it("strips markup", () => {
    expect(markdownToPlain("## Head\n\n**Bold** and *it*\n- item")).toBe("Head\n\nBold and it\n• item");
  });
});

describe("looksLikeMarkdown", () => {
  it("detects markdown", () => {
    expect(looksLikeMarkdown("# Title")).toBe(true);
    expect(looksLikeMarkdown("text\n- item")).toBe(true);
    expect(looksLikeMarkdown("Some **bold** words")).toBe(true);
    expect(looksLikeMarkdown("Just a plain sentence - with a dash.")).toBe(false);
  });
});
