// The deck's copy of a link as a pure codec: what a record keeps of each kind
// of payload, the XML it travels as and the caps that cut it to identity only.
// What is refused on the way back in is deck-record.refuse.test.ts.

import { describe, expect, it } from "vitest";
import {
  buildDeckRecord,
  decodeDeckRecord,
  DECK_NAMESPACE,
  DECK_RECORD_MAX_BYTES,
  DECK_STORE_MAX_BYTES,
  encodeDeckRecord,
  packDeckRecord,
  payloadFromRecord,
  type DeckRecord,
} from "./deck-record";
import {
  CHART,
  HASH,
  ID,
  LINK,
  picture,
  PUSHED,
  record,
  SAVED,
  SRC,
  table,
  TEXT,
} from "./deck-record.support";

describe("what a record keeps", () => {
  it("names the link, the revision and the three times", () => {
    expect(record(table(2, 2))).toMatchObject({
      v: 1,
      id: ID,
      kind: "table",
      label: "Model!B4:F12 table",
      project: "Amasty",
      src: SRC,
      token: LINK.token,
      rev: 7,
      sentAt: LINK.sentAt,
      pushedAt: PUSHED,
      savedAt: SAVED,
      dataKept: true,
    });
  });

  it("keeps a table's cells, widths, header flag and hash, and says nothing twice", () => {
    const source = table(2, 3);
    const kept = record(source).payload;
    expect(kept).toEqual({
      v: 1,
      kind: "table",
      rows: 2,
      cols: 3,
      cells: source.cells,
      widths: source.widths,
      hash: HASH,
      h: true,
    });
    expect(kept).not.toHaveProperty("src");
    expect(kept).not.toHaveProperty("pushedAt");
  });

  it("keeps a text link's text", () => {
    expect(record(TEXT, { ...LINK, kind: "text" }).payload).toEqual({
      v: 1,
      kind: "text",
      text: "EUR 3.2m",
      hash: HASH,
    });
  });

  it("keeps a chart's data and size but never the picture bytes", () => {
    const kept = record(picture({ chart: CHART }), { ...LINK, kind: "chart" });
    expect(kept.dataKept).toBe(true);
    expect(kept.payload).toEqual({
      v: 1,
      kind: "picture",
      width: 800,
      height: 400,
      hash: HASH,
      chart: CHART,
    });
    expect(JSON.stringify(kept)).not.toContain("iVBOR");
  });

  it.each([
    ["a range picture", picture(), "range" as const],
    [
      "a chart Excel could not describe",
      picture({ chartIssue: "chart type unsupported" }),
      "chart" as const,
    ],
  ])("keeps identity only for %s", (_name, payload, kind) => {
    const kept = record(payload, { ...LINK, kind });
    expect(kept.dataKept).toBe(false);
    expect(kept.payload).toBeNull();
    expect(kept).toMatchObject({ id: ID, rev: 7, src: SRC, pushedAt: PUSHED });
  });

  it("leaves a link with no project without one", () => {
    const bare = { ...LINK, project: undefined };
    expect(record(table(1, 1), bare)).not.toHaveProperty("project");
  });

  it("hands a payload back as it was, less the picture bytes", () => {
    const source = table(3, 2);
    expect(payloadFromRecord(record(source))).toEqual(source);
    expect(payloadFromRecord(record(TEXT))).toEqual(TEXT);
    const chart = picture({ chart: CHART });
    expect(payloadFromRecord(record(chart))).toEqual({ ...chart, png: "" });
    expect(payloadFromRecord(record(picture()))).toBeNull();
  });
});

describe("the XML a record travels as", () => {
  it("is one link element in the deck namespace, its body base64url", () => {
    const xml = encodeDeckRecord(record(table(2, 2)));
    const match =
      /^<link xmlns="urn:plsfix:deck-link:1" v="1" id="([0-9a-f]{32})">([A-Za-z0-9_-]+)<\/link>$/.exec(
        xml,
      );
    expect(match?.[1]).toBe(ID);
    expect(DECK_NAMESPACE).toBe("urn:plsfix:deck-link:1");
  });

  it.each([
    ["a table", record(table(3, 4))],
    ["a text", record(TEXT, { ...LINK, kind: "text" })],
    ["a chart", record(picture({ chart: CHART }), { ...LINK, kind: "chart" })],
    ["a range picture", record(picture(), { ...LINK, kind: "range" })],
  ])("round trips %s", (_name, source) => {
    expect(decodeDeckRecord(encodeDeckRecord(source))).toEqual({
      ok: true,
      record: source,
    });
  });

  it("carries cell text that would break the XML, byte for byte", () => {
    const nasty = ["]]>", "<b>", "a & b", `"quoted"`, "'single'", "]]><x/>"];
    const source = record(table(2, 3, (r, c) => nasty[(r * 3 + c) % 6]!));
    const xml = encodeDeckRecord(source);
    const inner = xml.slice(xml.indexOf(">") + 1, xml.lastIndexOf("</link>"));
    expect(inner).not.toMatch(/[<>&\]]/);
    expect(decodeDeckRecord(xml)).toEqual({ ok: true, record: source });
  });

  it("carries non-ASCII text: Latvian letters, the euro sign, an emoji", () => {
    const words = ["Ieņēmumi", "EUR 1 094 417 €", "Peļņa 📈"];
    const source = record(table(1, 3, (_r, c) => words[c]!));
    expect(decodeDeckRecord(encodeDeckRecord(source))).toEqual({
      ok: true,
      record: source,
    });
  });

  it("carries a 60 x 20 table whole, and well inside the record cap", () => {
    const big = table(60, 20, (r, c) => `${String((r * 20 + c) * 1234.5)}`);
    const source = record(big);
    const xml = encodeDeckRecord(source);
    expect(source.dataKept).toBe(true);
    expect(xml.length).toBeLessThan(DECK_RECORD_MAX_BYTES);
    const decoded = decodeDeckRecord(xml);
    expect(decoded).toEqual({ ok: true, record: source });
    const back = (decoded as { record: DeckRecord }).record;
    expect(payloadFromRecord(back)).toEqual(big);
  });

  it("refuses to write a record whose id is not a link id", () => {
    expect(() => encodeDeckRecord({ ...record(TEXT), id: "nope" })).toThrow(
      /encodeDeckRecord/,
    );
  });

  it("builds the same record from the same inputs, whenever it is asked", () => {
    const again = buildDeckRecord(LINK, table(2, 2), 7, SAVED);
    expect(again).toEqual(record(table(2, 2)));
  });
});

describe("the caps", () => {
  it("are 256 KB a record and 4 MB a deck", () => {
    expect(DECK_RECORD_MAX_BYTES).toBe(256 * 1024);
    expect(DECK_STORE_MAX_BYTES).toBe(4 * 1024 * 1024);
  });

  it("cut a record over the cap to identity only, which still says who it is", () => {
    const huge = record(table(60, 20, () => "x".repeat(400)));
    const packed = packDeckRecord(huge);
    expect(packed.xml.length).toBeLessThanOrEqual(DECK_RECORD_MAX_BYTES);
    expect(packed.record).toMatchObject({
      id: ID,
      kind: "table",
      label: "Model!B4:F12 table",
      src: SRC,
      token: LINK.token,
      rev: 7,
      dataKept: false,
      payload: null,
    });
    expect(encodeDeckRecord(huge)).toBe(packed.xml);
    expect(decodeDeckRecord(packed.xml)).toEqual({
      ok: true,
      record: packed.record,
    });
  });

  it("leave a record under the cap whole", () => {
    const source = record(table(60, 20));
    const packed = packDeckRecord(source);
    expect(packed.record).toBe(source);
    expect(packed.record.dataKept).toBe(true);
  });
});
