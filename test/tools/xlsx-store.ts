// Test tooling: writes one worksheet of formulas, numbers, text and error values
// as a real .xlsx (five XML parts in a stored zip) another spreadsheet engine can
// open. Owns the zip bytes and the sheet XML; a cell carries its formula and no
// cached value, so the engine has to calculate it.

import fs from "node:fs";

/** One cell of the sheet: an A1 address and exactly one kind of content. */
export interface XlsxCell {
  address: string;
  formula?: string;
  number?: number;
  text?: string;
  error?: string;
}

const CRC_TABLE = Array.from({ length: 256 }, (_unused, n) => {
  let c = n;
  for (let bit = 0; bit < 8; bit += 1)
    c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
  return c >>> 0;
});

function crc32(data: Buffer): number {
  let crc = 0xffffffff;
  for (const byte of data)
    crc = (CRC_TABLE[(crc ^ byte) & 0xff] ?? 0) ^ (crc >>> 8);
  return (crc ^ 0xffffffff) >>> 0;
}

// 1980-01-01 00:00, the earliest time a zip entry can carry.
const DOS_DATE = 0x21;

function u16(value: number): Buffer {
  const out = Buffer.alloc(2);
  out.writeUInt16LE(value);
  return out;
}

function u32(value: number): Buffer {
  const out = Buffer.alloc(4);
  out.writeUInt32LE(value);
  return out;
}

// The local header and data of one stored entry, and its central directory record.
function zipEntry(name: string, body: string, offset: number) {
  const data = Buffer.from(body, "utf8");
  const nameBytes = Buffer.from(name, "utf8");
  const crc = crc32(data);
  const common = [
    u16(20), // version needed
    u16(0), // flags
    u16(0), // method: stored
    u16(0), // time
    u16(DOS_DATE),
    u32(crc),
    u32(data.length),
    u32(data.length),
    u16(nameBytes.length),
    u16(0), // extra length
  ];
  const local = Buffer.concat([u32(0x04034b50), ...common, nameBytes, data]);
  const central = Buffer.concat([
    u32(0x02014b50),
    u16(20), // version made by
    ...common,
    u16(0), // comment length
    u16(0), // disk number
    u16(0), // internal attributes
    u32(0), // external attributes
    u32(offset),
    nameBytes,
  ]);
  return { local, central };
}

function zip(parts: [string, string][]): Buffer {
  const locals: Buffer[] = [];
  const centrals: Buffer[] = [];
  let offset = 0;
  for (const [name, body] of parts) {
    const entry = zipEntry(name, body, offset);
    locals.push(entry.local);
    centrals.push(entry.central);
    offset += entry.local.length;
  }
  const directory = Buffer.concat(centrals);
  const end = Buffer.concat([
    u32(0x06054b50),
    u16(0),
    u16(0),
    u16(parts.length),
    u16(parts.length),
    u32(directory.length),
    u32(offset),
    u16(0),
  ]);
  return Buffer.concat([...locals, directory, end]);
}

const escapeXml = (text: string): string =>
  text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

function cellXml({ address, formula, number, text, error }: XlsxCell): string {
  if (formula !== undefined)
    return `<c r="${address}"><f>${escapeXml(formula)}</f></c>`;
  if (number !== undefined)
    return `<c r="${address}"><v>${String(number)}</v></c>`;
  if (error !== undefined) return `<c r="${address}" t="e"><v>${error}</v></c>`;
  const body = escapeXml(text ?? "");
  return `<c r="${address}" t="inlineStr"><is><t xml:space="preserve">${body}</t></is></c>`;
}

function sheetXml(cells: XlsxCell[]): string {
  const rows = new Map<number, string[]>();
  for (const cell of cells) {
    const row = Number(/\d+$/.exec(cell.address)?.[0]);
    rows.set(row, [...(rows.get(row) ?? []), cellXml(cell)]);
  }
  const body = [...rows.entries()]
    .sort(([a], [b]) => a - b)
    .map(([row, xml]) => `<row r="${String(row)}">${xml.join("")}</row>`);
  return `<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${body.join("")}</sheetData></worksheet>`;
}

const NS_RELS = "http://schemas.openxmlformats.org/package/2006/relationships";
const NS_DOC = "http://schemas.openxmlformats.org/officeDocument/2006";
const XML_HEAD = `<?xml version="1.0" encoding="UTF-8"?>`;

/** Writes `cells` as the only sheet of a new workbook at `file`. */
export function writeXlsx(file: string, cells: XlsxCell[]): void {
  const parts: [string, string][] = [
    [
      "[Content_Types].xml",
      `${XML_HEAD}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`,
    ],
    [
      "_rels/.rels",
      `${XML_HEAD}<Relationships xmlns="${NS_RELS}"><Relationship Id="rId1" Type="${NS_DOC}/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    ],
    [
      "xl/workbook.xml",
      `${XML_HEAD}<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${NS_DOC}/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    ],
    [
      "xl/_rels/workbook.xml.rels",
      `${XML_HEAD}<Relationships xmlns="${NS_RELS}"><Relationship Id="rId1" Type="${NS_DOC}/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
    ],
    ["xl/worksheets/sheet1.xml", sheetXml(cells)],
  ];
  fs.writeFileSync(file, zip(parts));
}
