// The deck's design record: the brand list a deck keeps and the theme colors it
// had before the brand was written into them, one custom XML part in its own
// namespace. The codec is held to the deck-link record's rules: strict decode
// with a reason, a newer version says "update pls,fix", and nothing is half used.
import { describe, expect, it } from "vitest";
import { toBase64Url } from "../link/crypto";
import {
  buildDeckDesign,
  decodeDeckDesign,
  DESIGN_NAMESPACE,
  DESIGN_RECORD_MAX_BYTES,
  encodeDeckDesign,
  NEWER_DECK_DESIGN,
  type DeckDesign,
  type SavedTheme,
} from "./deck-design-record";

const GOLD = "#B27E54";
const MARBLE = "#282623";
const GREEN = "#273631";

const OFFICE: SavedTheme = {
  master: "master-1",
  accents: ["#4472C4", "#ED7D31", "#A5A5A5", "#FFC000", "#5B9BD5", "#70AD47"],
};

function design(): DeckDesign {
  return buildDeckDesign([GOLD, MARBLE, GREEN], [OFFICE])!;
}

// A part as the host would hold it, built by hand so a test can bend one field.
function part(
  json: unknown,
  { v = "1", id = "design", ns = DESIGN_NAMESPACE } = {},
): string {
  const body = toBase64Url(new TextEncoder().encode(JSON.stringify(json)));
  return `<${id} xmlns="${ns}" v="${v}">${body}</${id}>`;
}

const GOOD = { v: 1, brand: [GOLD, MARBLE], themes: [] as unknown[] };

function refusal(xml: string): string {
  const read = decodeDeckDesign(xml);
  if (read.ok) throw new Error("expected a refusal");
  return read.reason;
}

describe("buildDeckDesign", () => {
  it("keeps upper-case colors, each once, in the order given", () => {
    const built = buildDeckDesign(["#b27e54", GOLD, MARBLE, "273631"]);
    expect(built).toEqual({
      v: 1,
      brand: [GOLD, MARBLE, GREEN],
      themes: [],
    });
  });

  it("refuses fewer than two distinct colors", () => {
    expect(buildDeckDesign([GOLD])).toBeNull();
    expect(buildDeckDesign([GOLD, "#b27e54"])).toBeNull();
    expect(buildDeckDesign(["nope", "worse"])).toBeNull();
  });

  it("holds at most ten colors", () => {
    const twelve = Array.from(
      { length: 12 },
      (_, at) => `#${String(at + 10)}A0C0`,
    );
    expect(buildDeckDesign(twelve)!.brand).toHaveLength(10);
  });

  it("takes a saved theme only whole: six colors under one master", () => {
    const built = buildDeckDesign([GOLD, MARBLE], [OFFICE])!;
    expect(built.themes).toEqual([OFFICE]);
    expect(
      buildDeckDesign([GOLD, MARBLE], [{ master: "m", accents: ["#111111"] }]),
    ).toBeNull();
    expect(
      buildDeckDesign(
        [GOLD, MARBLE],
        [{ master: "", accents: OFFICE.accents }],
      ),
    ).toBeNull();
  });
});

describe("a record's round trip", () => {
  it("is one element in the design namespace", () => {
    const xml = encodeDeckDesign(design());
    expect(xml).toMatch(
      /^<design xmlns="urn:plsfix:deck-design:1" v="1">[A-Za-z0-9_-]+<\/design>$/,
    );
  });

  it("never uses the link store's namespace", () => {
    expect(DESIGN_NAMESPACE).toBe("urn:plsfix:deck-design:1");
    expect(DESIGN_NAMESPACE).not.toBe("urn:plsfix:deck-link:1");
  });

  it("reads back what was written, the saved theme included", () => {
    const read = decodeDeckDesign(encodeDeckDesign(design()));
    expect(read).toEqual({ ok: true, design: design() });
  });

  it("reads back a design with no saved theme", () => {
    const bare = buildDeckDesign([GOLD, MARBLE])!;
    expect(decodeDeckDesign(encodeDeckDesign(bare))).toEqual({
      ok: true,
      design: bare,
    });
  });

  it("reads a part with a declaration, a mark and white space around it", () => {
    const xml = encodeDeckDesign(design());
    const wrapped = `﻿<?xml version="1.0" encoding="UTF-8"?>\n  ${xml}\n`;
    expect(decodeDeckDesign(wrapped)).toEqual({ ok: true, design: design() });
  });

  it("reads single-quoted attributes in any order", () => {
    const body = toBase64Url(new TextEncoder().encode(JSON.stringify(GOOD)));
    const xml = `<design v='1' xmlns='${DESIGN_NAMESPACE}'>${body}</design>`;
    expect(decodeDeckDesign(xml).ok).toBe(true);
  });
});

describe("decode trusts nothing", () => {
  it("refuses text that is not one XML element", () => {
    expect(refusal("")).toBe("notXml");
    expect(refusal("plain text")).toBe("notXml");
    expect(refusal("<design></design><design></design>")).toBe("notXml");
    expect(refusal(42 as unknown as string)).toBe("notXml");
  });

  it("refuses another namespace, and another element in ours", () => {
    expect(refusal(part(GOOD, { ns: "urn:plsfix:deck-link:1" }))).toBe(
      "wrongNamespace",
    );
    expect(refusal(part(GOOD, { id: "link" }))).toBe("notDesign");
  });

  it("says update pls,fix for a newer version, outside or inside", () => {
    const outside = decodeDeckDesign(part(GOOD, { v: "2" }));
    expect(outside).toEqual({
      ok: false,
      reason: "newerVersion",
      message: NEWER_DECK_DESIGN,
    });
    expect(refusal(part({ ...GOOD, v: 2 }))).toBe("newerVersion");
    expect(NEWER_DECK_DESIGN).toMatch(/update/i);
  });

  it("refuses a version it cannot read", () => {
    expect(refusal(part(GOOD, { v: "x" }))).toBe("badRecord");
    expect(refusal(part(GOOD, { v: "0" }))).toBe("badRecord");
    expect(refusal(part({ ...GOOD, v: "1" }))).toBe("badRecord");
  });

  it("refuses a body that is not base64url JSON", () => {
    const xml = `<design xmlns="${DESIGN_NAMESPACE}" v="1">!!!</design>`;
    expect(refusal(xml)).toBe("badBody");
    const empty = `<design xmlns="${DESIGN_NAMESPACE}" v="1"></design>`;
    expect(refusal(empty)).toBe("badBody");
    expect(refusal(part([1, 2]))).toBe("badBody");
  });

  it("refuses a brand list that is not two to ten distinct colors", () => {
    expect(refusal(part({ ...GOOD, brand: [GOLD] }))).toBe("badRecord");
    expect(refusal(part({ ...GOOD, brand: [GOLD, GOLD] }))).toBe("badRecord");
    expect(refusal(part({ ...GOOD, brand: [GOLD, "red"] }))).toBe("badRecord");
    expect(refusal(part({ ...GOOD, brand: "gold" }))).toBe("badRecord");
    const eleven = Array.from(
      { length: 11 },
      (_, at) => `#${String(at + 10)}A0C0`,
    );
    expect(refusal(part({ ...GOOD, brand: eleven }))).toBe("badRecord");
  });

  it("refuses a saved theme that is not six colors under a named master", () => {
    const bad = (themes: unknown): string => refusal(part({ ...GOOD, themes }));
    expect(bad("themes")).toBe("badRecord");
    expect(bad([{ master: "m", accents: ["#111111"] }])).toBe("badRecord");
    expect(
      bad([{ master: "m", accents: [...OFFICE.accents.slice(1), "x"] }]),
    ).toBe("badRecord");
    expect(bad([{ accents: OFFICE.accents }])).toBe("badRecord");
    expect(bad([OFFICE, OFFICE])).toBe("badRecord");
  });

  it("reads a record written with no themes key as none saved", () => {
    const read = decodeDeckDesign(part({ v: 1, brand: [GOLD, MARBLE] }));
    expect(read).toEqual({
      ok: true,
      design: { v: 1, brand: [GOLD, MARBLE], themes: [] },
    });
  });

  it("refuses a part larger than any design record", () => {
    const big = `<design xmlns="${DESIGN_NAMESPACE}" v="1">${"A".repeat(DESIGN_RECORD_MAX_BYTES + 2000)}</design>`;
    expect(refusal(big)).toBe("tooBig");
  });

  it("ignores a key it does not know, and does not write it back", () => {
    const read = decodeDeckDesign(part({ ...GOOD, later: { x: 1 } }));
    expect(read.ok).toBe(true);
    if (read.ok)
      expect(Object.keys(read.design).sort()).toEqual(["brand", "themes", "v"]);
  });
});
