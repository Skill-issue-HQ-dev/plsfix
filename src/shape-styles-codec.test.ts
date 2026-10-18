// The codec of the stored shape styles list. A newer version reads as "update
// pls,fix", a corrupt store reads as empty with a reason and never throws, a bad
// entry is left out and counted, and what is written holds the known fields
// only. No Office.js.

import fc from "fast-check";
import { describe, expect, it } from "vitest";
import {
  decodeShapeStyles,
  encodeShapeStyles,
  NEWER_SHAPE_STYLES,
  readLook,
  SHAPE_STYLES_KEY,
  SHAPE_STYLES_VERSION,
} from "./shape-styles-codec";
import { starterShapeStyles, type ShapeStyle } from "./shape-styles";
import { DEFAULT_SETTINGS } from "./settings";

const USER: ShapeStyle = {
  name: "Margin box",
  kind: "user",
  fill: "#EBEDDE",
  fillTransparency: 0.25,
  line: {
    visible: true,
    color: "#525F0E",
    weight: 1.5,
    dashStyle: "Dash",
    transparency: 0,
    style: "Single",
  },
  text: {
    name: "Aptos",
    size: 14,
    color: "#1F1D1B",
    bold: true,
    italic: false,
  },
};

function readOk(raw: string | null | undefined) {
  const read = decodeShapeStyles(raw);
  if (!read.ok)
    throw new Error(`expected a readable list, got: ${read.message}`);
  return read;
}

describe("the stored key", () => {
  it("is plsfix.shapeStyles.v1 and the codec writes version 1", () => {
    expect(SHAPE_STYLES_KEY).toBe("plsfix.shapeStyles.v1");
    expect(SHAPE_STYLES_VERSION).toBe(1);
    const stored = JSON.parse(encodeShapeStyles([])) as { v: number };
    expect(stored.v).toBe(1);
  });
});

describe("encodeShapeStyles and decodeShapeStyles", () => {
  it("round-trips the starters and a user style in order", () => {
    const list = [...starterShapeStyles(DEFAULT_SETTINGS), USER];
    const read = readOk(encodeShapeStyles(list));
    expect(read.styles).toEqual(list);
    expect(read.reason).toBeUndefined();
  });

  it("keeps the newest user style last", () => {
    const list: ShapeStyle[] = [
      { name: "First", kind: "user", fill: "#111111" },
      { name: "Second", kind: "user", fill: "#222222" },
    ];
    expect(readOk(encodeShapeStyles(list)).styles.map((s) => s.name)).toEqual([
      "First",
      "Second",
    ]);
  });

  it("reads nothing stored, and a blank value, as an empty list with no reason", () => {
    for (const raw of [null, undefined, "", "  \n"]) {
      const read = readOk(raw);
      expect(read.styles).toEqual([]);
      expect(read.reason).toBeUndefined();
    }
  });

  it("reads a newer version as update pls,fix, and writes nothing of it", () => {
    const raw = JSON.stringify({ v: SHAPE_STYLES_VERSION + 1, styles: [USER] });
    const read = decodeShapeStyles(raw);
    expect(read.ok).toBe(false);
    if (!read.ok) {
      expect(read.message).toBe(NEWER_SHAPE_STYLES);
      expect(read.message).toContain("Update");
    }
  });

  it("reads a corrupt store as empty, with a reason, and never throws", () => {
    for (const raw of [
      "{not json",
      "42",
      "null",
      "[]",
      '"text"',
      "{}",
      '{"styles": []}',
      '{"v": "one", "styles": []}',
      '{"v": 1, "styles": "nope"}',
    ]) {
      const read = readOk(raw);
      expect(read.styles, raw).toEqual([]);
      expect(read.reason, raw).toMatch(/could not be read/);
    }
  });

  it("leaves out an entry it cannot trust and says how many", () => {
    const raw = JSON.stringify({
      v: 1,
      styles: [
        { name: "Good", kind: "user", fill: "#112233" },
        { name: "", kind: "user" },
        { name: "No kind" },
        { name: "Odd kind", kind: "other" },
        "text",
        null,
        { name: "Bad fill", kind: "user", fill: 7 },
        { name: "Bad line", kind: "user", line: { visible: "yes" } },
      ],
    });
    const read = readOk(raw);
    expect(read.styles.map((style) => style.name)).toEqual(["Good"]);
    expect(read.reason).toBe("7 styles could not be read and were left out.");
  });

  it("says one style was left out, not one styles were", () => {
    const raw = JSON.stringify({
      v: 1,
      styles: [{ name: "Good", kind: "user" }, { name: "" }],
    });
    expect(readOk(raw).reason).toBe(
      "1 style could not be read and was left out.",
    );
  });

  it("keeps the first of two names that differ only in case", () => {
    const raw = JSON.stringify({
      v: 1,
      styles: [
        { name: "Callout", kind: "starter", fill: "#111111" },
        { name: "CALLOUT", kind: "user", fill: "#222222" },
      ],
    });
    const read = readOk(raw);
    expect(read.styles).toEqual([
      { name: "Callout", kind: "starter", fill: "#111111" },
    ]);
  });

  it("normalizes colors to #RRGGBB", () => {
    const raw = JSON.stringify({
      v: 1,
      styles: [
        {
          name: "Loose",
          kind: "user",
          fill: "#abc",
          line: { visible: true, color: "b27e54", weight: 2 },
          text: { color: " #112233 ", bold: true },
        },
      ],
    });
    const [style] = readOk(raw).styles;
    expect(style!.fill).toBe("#AABBCC");
    expect(style!.line).toEqual({ visible: true, color: "#B27E54", weight: 2 });
    expect(style!.text).toEqual({ color: "#112233", bold: true });
  });

  it("leaves out an entry with a color, number or flag a host would refuse", () => {
    const bad: Record<string, unknown>[] = [
      { fill: "not a color" },
      { fill: 7 },
      { fillTransparency: 1.5 },
      { line: { visible: true, weight: 0 } },
      { line: { visible: true, transparency: -1 } },
      { line: { visible: true, color: "red" } },
      { line: { visible: "yes" } },
      { line: "thin" },
      { text: { size: -3 } },
      { text: { color: "not a color" } },
      { text: { bold: "yes" } },
      { text: { name: "" } },
    ];
    for (const fields of bad) {
      const raw = JSON.stringify({
        v: 1,
        styles: [{ name: "Bad", kind: "user", ...fields }],
      });
      const read = readOk(raw);
      expect(read.styles, JSON.stringify(fields)).toEqual([]);
      expect(read.reason, JSON.stringify(fields)).toMatch(
        /^1 style could not be read/,
      );
    }
  });

  it("keeps null as no fill and an absent fill as leave it alone", () => {
    const raw = JSON.stringify({
      v: 1,
      styles: [
        { name: "None", kind: "user", fill: null },
        { name: "Alone", kind: "user" },
      ],
    });
    const [none, alone] = readOk(raw).styles;
    expect(none!.fill).toBeNull();
    expect(alone).not.toHaveProperty("fill");
  });

  it("holds an invisible line to its one fact", () => {
    const raw = JSON.stringify({
      v: 1,
      styles: [
        {
          name: "Hidden",
          kind: "user",
          line: { visible: false, color: "#112233", weight: 4 },
        },
      ],
    });
    expect(readOk(raw).styles[0]!.line).toEqual({ visible: false });
  });

  it("writes the known fields only", () => {
    const stored = JSON.parse(
      encodeShapeStyles([
        {
          ...USER,
          secret: "x",
          line: { ...USER.line!, extra: 1 },
        } as ShapeStyle,
      ]),
    ) as { styles: Record<string, unknown>[] };
    expect(Object.keys(stored.styles[0]!).sort()).toEqual(
      ["fill", "fillTransparency", "kind", "line", "name", "text"].sort(),
    );
    expect(Object.keys(stored.styles[0]!.line as object).sort()).toEqual(
      [
        "color",
        "dashStyle",
        "style",
        "transparency",
        "visible",
        "weight",
      ].sort(),
    );
  });

  it("never throws on any JSON value", () => {
    fc.assert(
      fc.property(fc.jsonValue(), (value) => {
        const read = decodeShapeStyles(JSON.stringify(value));
        if (read.ok) {
          for (const style of read.styles) {
            expect(typeof style.name).toBe("string");
            expect(["starter", "user"]).toContain(style.kind);
          }
        } else {
          expect(read.message).toBe(NEWER_SHAPE_STYLES);
        }
      }),
      { seed: 20261007 },
    );
  });

  it("never throws on any text", () => {
    fc.assert(
      fc.property(fc.string(), (text) => {
        expect(() => decodeShapeStyles(text)).not.toThrow();
      }),
      { seed: 20261007 },
    );
  });
});

describe("readLook", () => {
  it("keeps what a host answered that a style may hold", () => {
    expect(
      readLook({
        fill: "#b27e54",
        fillTransparency: 0.25,
        line: {
          visible: true,
          color: "#282623",
          weight: 1.5,
          dashStyle: "Dash",
          transparency: 0,
          style: "Single",
        },
        text: {
          name: "Aptos",
          size: 14,
          color: "#1f1d1b",
          bold: true,
          italic: false,
        },
      }),
    ).toEqual({
      fill: "#B27E54",
      fillTransparency: 0.25,
      line: {
        visible: true,
        color: "#282623",
        weight: 1.5,
        dashStyle: "Dash",
        transparency: 0,
        style: "Single",
      },
      text: {
        name: "Aptos",
        size: 14,
        color: "#1F1D1B",
        bold: true,
        italic: false,
      },
    });
  });

  it("reads null as no fill, and leaves out a fill that is no color", () => {
    expect(readLook({ fill: null })).toEqual({ fill: null });
    for (const fill of ["", "gold", 7, undefined]) {
      expect(readLook({ fill, fillTransparency: 0.5 }), String(fill)).toEqual(
        {},
      );
    }
  });

  it("leaves out every field the host answered null, empty or beyond what a style holds", () => {
    expect(
      readLook({
        fill: "#112233",
        fillTransparency: 2,
        line: {
          visible: true,
          color: null,
          weight: 0,
          dashStyle: "",
          transparency: -1,
          style: null,
        },
        text: { name: null, size: 0, color: "", bold: null, italic: undefined },
      }),
    ).toEqual({ fill: "#112233", line: { visible: true } });
  });

  it("holds an invisible line to its one fact, and leaves out a line whose visibility is unknown", () => {
    expect(
      readLook({ line: { visible: false, color: "#112233", weight: 3 } }),
    ).toEqual({ line: { visible: false } });
    expect(readLook({ line: { visible: null, color: "#112233" } })).toEqual({});
  });

  it("leaves a text look out when none of it was readable", () => {
    expect(readLook({ text: { name: null, size: null, bold: null } })).toEqual(
      {},
    );
  });
});
