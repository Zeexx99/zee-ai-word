/**
 * Reference files and images for Gemini (pure, unit-tested where possible).
 * PDFs and images are sent as-is (Gemini reads them natively); .docx is unzipped in the
 * browser and reduced to text; text formats are read as text.
 */

export const MAX_REFERENCE_BYTES = 15 * 1024 * 1024;

export type ReferencePayload = { kind: "inline"; mimeType: string; data: string; name: string } | { kind: "text"; text: string; name: string };

export function detectImageMime(base64: string): string {
  const b = base64.replace(/^data:[^,]+,/, "").slice(0, 16);
  if (b.startsWith("iVBOR")) return "image/png";
  if (b.startsWith("/9j/")) return "image/jpeg";
  if (b.startsWith("R0lGOD")) return "image/gif";
  if (b.startsWith("UklGR")) return "image/webp";
  if (b.startsWith("Qk")) return "image/bmp";
  return "image/png";
}

export function bytesToBase64(bytes: Uint8Array): string {
  let s = "";
  const CH = 0x8000;
  for (let i = 0; i < bytes.length; i += CH) s += String.fromCharCode(...bytes.subarray(i, i + CH));
  return btoa(s);
}

/** Minimal ZIP reader: returns the (inflated) bytes of one entry, or null. Uses DecompressionStream. */
export async function readZipEntry(zip: Uint8Array, name: string): Promise<Uint8Array | null> {
  const dv = new DataView(zip.buffer, zip.byteOffset, zip.byteLength);
  // End of central directory record (search backwards).
  let eocd = -1;
  for (let i = zip.length - 22; i >= Math.max(0, zip.length - 65557); i--) {
    if (dv.getUint32(i, true) === 0x06054b50) {
      eocd = i;
      break;
    }
  }
  if (eocd < 0) return null;
  const count = dv.getUint16(eocd + 10, true);
  let p = dv.getUint32(eocd + 16, true);
  const dec = new TextDecoder();
  for (let n = 0; n < count && p + 46 <= zip.length; n++) {
    if (dv.getUint32(p, true) !== 0x02014b50) return null;
    const method = dv.getUint16(p + 10, true);
    const compSize = dv.getUint32(p + 20, true);
    const nameLen = dv.getUint16(p + 28, true);
    const extraLen = dv.getUint16(p + 30, true);
    const commentLen = dv.getUint16(p + 32, true);
    const localOff = dv.getUint32(p + 42, true);
    const entryName = dec.decode(zip.subarray(p + 46, p + 46 + nameLen));
    if (entryName === name) {
      const lNameLen = dv.getUint16(localOff + 26, true);
      const lExtraLen = dv.getUint16(localOff + 28, true);
      const start = localOff + 30 + lNameLen + lExtraLen;
      const data = zip.subarray(start, start + compSize);
      if (method === 0) return data;
      if (method !== 8) return null;
      const stream = new Blob([new Uint8Array(data)]).stream().pipeThrough(new DecompressionStream("deflate-raw"));
      return new Uint8Array(await new Response(stream).arrayBuffer());
    }
    p += 46 + nameLen + extraLen + commentLen;
  }
  return null;
}

/** Plain text of a WordprocessingML document.xml (paragraphs → lines, tabs/breaks kept). */
export function wordXmlToText(xml: string): string {
  return xml
    .replace(/<w:tab\/>/g, "\t")
    .replace(/<w:br[^>]*\/>/g, "\n")
    .replace(/<\/w:p>/g, "\n")
    .replace(/<\/w:tc>/g, "\t")
    .replace(/<[^>]+>/g, "")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&")
    .replace(/[ \t]+\n/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

export async function docxToText(bytes: Uint8Array): Promise<string> {
  const xml = await readZipEntry(bytes, "word/document.xml");
  if (!xml) throw new Error("This .docx file could not be read (no word/document.xml).");
  return wordXmlToText(new TextDecoder().decode(xml));
}

const TEXT_EXT = /\.(txt|md|markdown|csv|tsv|json|xml|html?|log)$/i;

/** Turns a user-picked file into something Gemini can read. */
export async function readReferenceFile(file: File): Promise<ReferencePayload> {
  if (file.size > MAX_REFERENCE_BYTES) throw new Error(`"${file.name}" is larger than 15 MB. Use a smaller file or a PDF of the relevant section.`);
  const name = file.name;
  if (/\.pdf$/i.test(name) || file.type === "application/pdf") {
    return { kind: "inline", mimeType: "application/pdf", data: bytesToBase64(new Uint8Array(await file.arrayBuffer())), name };
  }
  if (/\.docx$/i.test(name)) return { kind: "text", text: await docxToText(new Uint8Array(await file.arrayBuffer())), name };
  if (/^image\//.test(file.type)) return { kind: "inline", mimeType: file.type, data: bytesToBase64(new Uint8Array(await file.arrayBuffer())), name };
  if (TEXT_EXT.test(name) || file.type.startsWith("text/")) return { kind: "text", text: await file.text(), name };
  if (/\.doc$/i.test(name)) throw new Error("Old .doc files can't be read here. Save it as .docx or PDF first.");
  throw new Error(`"${name}" isn't a supported type. Use PDF, DOCX, TXT/MD/CSV or an image.`);
}
