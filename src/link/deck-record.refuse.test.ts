// What the deck record codec refuses on the way back in: a part that is not
// XML, not ours, not whole, not this version - each with its reason - and the
// shapes of a host's own re-serialising it still reads.

import { describe, expect, it } from "vitest";
import { toBase64Url } from "./crypto";
import {
  decodeDeckRecord,
  DECK_NAMESPACE,
  DECK_RECORD_MAX_BYTES,
  encodeDeckRecord,
  NEWER_DECK_RECORD,
} from "./deck-record";
import {
  ATTRS,
  CHART,
  HASH,
  ID,
  LINK,
  part,
  picture,
  record,
  refusal,
  table,
  TEXT,
  wire,
} from "./deck-record.support";

const NS = `xmlns="${DECK_NAMESPACE}"`;

describe("a part that is not ours or not whole", () => {
  const good = encodeDeckRecord(record(TEXT, { ...LINK, kind: "text" }));

  it.each([
    ["an empty string", ""],
    ["plain text", "not xml at all"],
    ["an unclosed element", good.replace("</link>", "")],
    ["two elements", `${good}${good}`],
    ["a self-closing element", `<link ${ATTRS}/>`],
    ["a different closing tag", good.replace("</link>", "</other>")],
  ])("refuses %s as not XML it can read", (_name, xml) => {
    expect(refusal(xml).reason).toBe("notXml");
  });

  it("refuses a part longer than any record may be", () => {
    const xml = `<link ${ATTRS}>${"A".repeat(2 * DECK_RECORD_MAX_BYTES)}</link>`;
    expect(refusal(xml).reason).toBe("tooBig");
  });

  it("refuses an element in another namespace", () => {
    const other = good.replace(DECK_NAMESPACE, "urn:someone:else:1");
    expect(refusal(other).reason).toBe("wrongNamespace");
    expect(refusal(good.replace(` ${NS}`, "")).reason).toBe("wrongNamespace");
  });

  it("ignores a foreign element that sits in the namespace", () => {
    const foreign = `<note ${ATTRS}>QUJD</note>`;
    expect(refusal(foreign).reason).toBe("notALink");
  });

  it("tells a part from a newer pls,fix to update the add-in", () => {
    const newer = part(wire(record(TEXT)), `${NS} v="2" id="${ID}"`);
    const said = refusal(newer);
    expect(said.reason).toBe("newerVersion");
    expect(said.message).toBe(NEWER_DECK_RECORD);
    expect(said.message).toMatch(/newer pls,fix\. Update the add-in\.$/);
  });

  it("does the same when only the JSON inside names the newer version", () => {
    expect(refusal(part({ ...wire(record(TEXT)), v: 2 })).reason).toBe(
      "newerVersion",
    );
  });

  it("reads a newer part's version before it judges its shape", () => {
    const newer = part({ v: 9, whatever: true }, `${NS} v="3" id="${ID}"`);
    expect(refusal(newer).reason).toBe("newerVersion");
  });

  it.each([
    ["no version", `${NS} id="${ID}"`],
    ["a version that is not a number", `${NS} v="x" id="${ID}"`],
    ["version 0", `${NS} v="0" id="${ID}"`],
    ["version 1.5", `${NS} v="1.5" id="${ID}"`],
  ])("refuses %s", (_name, attrs) => {
    expect(refusal(part(wire(record(TEXT)), attrs)).reason).toBe("badRecord");
  });

  it.each([
    ["no id", `${NS} v="1"`],
    ["an id in capitals", `${NS} v="1" id="${ID.toUpperCase()}"`],
    ["a short id", `${NS} v="1" id="abc"`],
    ["an id that is not hex", `${NS} v="1" id="${"g".repeat(32)}"`],
  ])("refuses %s", (_name, attrs) => {
    expect(refusal(part(wire(record(TEXT)), attrs)).reason).toBe("badId");
  });

  it("refuses an element whose id and body name different links", () => {
    const said = refusal(part({ ...wire(record(TEXT)), id: "f".repeat(32) }));
    expect(said.reason).toBe("badId");
  });

  const bodyOf = (text: string | Uint8Array): string =>
    `<link ${ATTRS}>${toBase64Url(
      typeof text === "string" ? new TextEncoder().encode(text) : text,
    )}</link>`;

  it.each([
    [
      "a body outside the base64url alphabet",
      `<link ${ATTRS}>not base64!</link>`,
    ],
    ["a body that is no JSON", bodyOf("{oops")],
    ["a body that is no text", bodyOf(new Uint8Array([0xff, 0xfe, 0xfd]))],
    ["an empty body", `<link ${ATTRS}></link>`],
    ["JSON that is not an object", part([1, 2, 3])],
    ["JSON null", part(null)],
  ])("refuses %s", (_name, xml) => {
    expect(refusal(xml).reason).toBe("badBody");
  });

  const base = wire(record(table(2, 2)));
  const without = (key: string): Record<string, unknown> =>
    Object.fromEntries(Object.entries(base).filter(([name]) => name !== key));
  const brokenChart = {
    ...wire(record(picture({ chart: CHART }))),
    payload: {
      v: 1,
      kind: "picture",
      width: 800,
      height: 400,
      hash: HASH,
      chart: { v: 1, kind: "column" },
    },
  };

  it.each([
    ["no label", without("label")],
    ["a label that is not text", { ...base, label: 5 }],
    ["no token", without("token")],
    ["an unknown kind", { ...base, kind: "video" }],
    ["a source that is not one", { ...base, src: { sheet: "x" } }],
    ["a revision that is not a number", { ...base, rev: "7" }],
    ["a negative revision", { ...base, rev: -1 }],
    ["a fractional revision", { ...base, rev: 1.5 }],
    ["a revision that is not finite", { ...base, rev: null }],
    ["no time of its own", without("savedAt")],
    ["no time it was sent", without("sentAt")],
    ["no time it was pushed", without("pushedAt")],
    ["a project that is not text", { ...base, project: 4 }],
    ["data kept with no payload", { ...base, payload: null }],
    ["a payload with no data kept", { ...base, dataKept: false }],
    ["a data flag that is not a boolean", { ...base, dataKept: "yes" }],
    ["a payload of an unknown kind", { ...base, payload: { kind: "video" } }],
    [
      "a table whose cells are not rows x cols",
      { ...base, payload: { ...(base.payload as object), rows: 3 } },
    ],
    ["a chart that is not drawable", brokenChart],
  ])("refuses a body with %s", (_name, json) => {
    expect(refusal(part(json)).reason).toBe("badRecord");
  });

  it("drops keys it does not know instead of trusting them", () => {
    const source = record(TEXT, { ...LINK, kind: "text" });
    const json = wire(source);
    const stray = { ...(json.payload as object), stray: 1 };
    const decoded = decodeDeckRecord(
      part({ ...json, extra: "ignored", payload: stray }),
    );
    expect(decoded).toEqual({ ok: true, record: source });
  });

  it("cleans a project name the way every other reader does", () => {
    const named = part({ ...wire(record(TEXT)), project: "  Amasty\u0007  " });
    const decoded = decodeDeckRecord(named);
    expect(decoded.ok && decoded.record.project).toBe("Amasty");
  });
});

describe("a part as a host may hand it back", () => {
  const source = record(table(2, 2));
  const xml = encodeDeckRecord(source);
  const reordered = xml.replace(
    /<link (.*?) (v="1") (id="[0-9a-f]+")>/,
    "<link $3 $2 $1>",
  );

  it.each([
    [
      "an XML declaration in front",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${xml}`,
    ],
    ["a byte order mark in front", `﻿${xml}`],
    ["white space around it", `\n  ${xml}\n`],
    ["single quotes", xml.replace(/"/g, "'")],
    ["its attributes in another order", reordered],
    [
      "white space inside the tag",
      xml.replace('" v=', '"   v=').replace('">', '" >'),
    ],
  ])("reads it with %s", (_name, text) => {
    expect(decodeDeckRecord(text)).toEqual({ ok: true, record: source });
  });
});
