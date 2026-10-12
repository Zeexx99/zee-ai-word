import { describe, expect, it } from "vitest";
import { deflateRawSync } from "node:zlib";
import { Buffer } from "node:buffer";
import { diffGrid, fitGrid, keepsNumbers, tableToMarkdown } from "../utils/tables";
import { detectImageMime, docxToText, readZipEntry, wordXmlToText } from "../utils/files";
import { buildTeamConfig, parseTeamConfig } from "../utils/teamConfig";
import { TEMPLATES } from "../services/templates";
import { markdownToHtml } from "../utils/mdToHtml";

/** Builds a minimal ZIP (one deflated entry) like a .docx. */
function makeZip(name: string, content: string): Uint8Array {
  const data = deflateRawSync(Buffer.from(content, "utf8"));
  const nameBuf = Buffer.from(name, "utf8");
  const local = Buffer.alloc(30);
  local.writeUInt32LE(0x04034b50, 0);
  local.writeUInt16LE(8, 8);
  local.writeUInt32LE(data.length, 18);
  local.writeUInt32LE(content.length, 22);
  local.writeUInt16LE(nameBuf.length, 26);
  const central = Buffer.alloc(46);
  central.writeUInt32LE(0x02014b50, 0);
  central.writeUInt16LE(8, 10);
  central.writeUInt32LE(data.length, 20);
  central.writeUInt32LE(content.length, 24);
  central.writeUInt16LE(nameBuf.length, 28);
  central.writeUInt32LE(0, 42);
  const cdOffset = local.length + nameBuf.length + data.length;
  const eocd = Buffer.alloc(22);
  eocd.writeUInt32LE(0x06054b50, 0);
  eocd.writeUInt16LE(1, 8);
  eocd.writeUInt16LE(1, 10);
  eocd.writeUInt32LE(central.length + nameBuf.length, 12);
  eocd.writeUInt32LE(cdOffset, 16);
  return new Uint8Array(Buffer.concat([local, nameBuf, data, central, nameBuf, eocd]));
}

describe("tables", () => {
  const t = [
    ["Well", "Oil (bbl/d)"],
    ["W-1", "1,250"],
    ["w-2", "980 BOPD"],
  ];
  it("renders markdown", () => {
    expect(tableToMarkdown(t)).toBe("| Well | Oil (bbl/d) |\n| --- | --- |\n| W-1 | 1,250 |\n| w-2 | 980 BOPD |");
  });
  it("fits model output to the original shape", () => {
    expect(fitGrid(t, [["Well", "Oil (bbl/d)"], ["W-1"]])).toEqual([["Well", "Oil (bbl/d)"], ["W-1", "1,250"], ["w-2", "980 BOPD"]]);
  });
  it("diffs cells and protects numbers", () => {
    const after = [["Well", "Oil (bbl/d)"], ["W-1", "1,250"], ["W-2", "980 bbl/d"]];
    expect(diffGrid(t, after).map((c) => [c.row, c.col])).toEqual([[2, 0], [2, 1]]);
    expect(keepsNumbers("980 BOPD", "980 bbl/d")).toBe(true);
    expect(keepsNumbers("1,250", "1250")).toBe(true);
    expect(keepsNumbers("980", "98")).toBe(false);
  });
});

describe("files", () => {
  it("detects image types", () => {
    expect(detectImageMime("iVBORw0KGgo")).toBe("image/png");
    expect(detectImageMime("/9j/4AAQ")).toBe("image/jpeg");
  });
  it("turns WordprocessingML into text", () => {
    expect(wordXmlToText('<w:p><w:r><w:t>Hello</w:t></w:r><w:tab/><w:r><w:t>A&amp;B</w:t></w:r></w:p><w:p><w:r><w:t>Next</w:t></w:r></w:p>')).toBe("Hello\tA&B\nNext");
  });
  it("reads a deflated zip entry and a docx", async () => {
    const xml = '<w:document><w:body><w:p><w:r><w:t>Pump intake 450 psi</w:t></w:r></w:p></w:body></w:document>';
    const zip = makeZip("word/document.xml", xml);
    const entry = await readZipEntry(zip, "word/document.xml");
    expect(new TextDecoder().decode(entry!)).toBe(xml);
    expect(await docxToText(zip)).toBe("Pump intake 450 psi");
    expect(await readZipEntry(zip, "missing.xml")).toBeNull();
  });
});

describe("team config", () => {
  it("round-trips and validates", () => {
    const json = JSON.stringify(buildTeamConfig("Ops", "ESP, TDH", [{ id: "a", name: "Daily", instruction: "Do it", mode: "insert" }]));
    const c = parseTeamConfig(json);
    expect(c.name).toBe("Ops");
    expect(c.customActions[0]).toEqual({ id: "team-a", name: "Daily", instruction: "Do it", mode: "insert" });
  });
  it("rejects foreign files", () => {
    expect(() => parseTeamConfig("{}")).toThrow();
    expect(() => parseTeamConfig("not json")).toThrow();
  });
});

describe("templates", () => {
  it("all templates convert to HTML with headings and tables", () => {
    for (const t of TEMPLATES) {
      const html = markdownToHtml(t.outline);
      expect(html.startsWith("<h1>")).toBe(true);
      expect(html.includes("<h2>")).toBe(true);
    }
  });
});
