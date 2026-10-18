// The custom styles' pure core, part one: a Paintbrush capture as a style record
// and the naming rules. No Office.js: src/excel/custom-styles.ts writes Excel.
// Starters are in custom-styles.starters.test.ts, the registry in
// custom-styles.registry.test.ts.

import { describe, expect, it } from "vitest";
import { checkStyleName, fromPaintSlot, STYLE_NAME_MAX } from "./custom-styles";
import type { PaintSlot } from "./paintbrush";

const edge = (style: string, color = "#000000") => ({
  style,
  weight: "Thin",
  color,
});

function slot(overrides: Partial<PaintSlot> = {}): PaintSlot {
  return {
    numberFormat: "0.0%",
    font: {
      name: "Aptos",
      size: 10,
      bold: true,
      italic: false,
      color: "#1F1D1B",
    },
    fill: "#EEF2D2",
    horizontalAlignment: "Right",
    borders: {
      top: edge("None"),
      bottom: edge("Continuous", "#6D7E12"),
      left: edge("None"),
      right: edge("None"),
    },
    ...overrides,
  };
}

describe("fromPaintSlot", () => {
  it("makes a user style that includes everything but protection", () => {
    const style = fromPaintSlot("Margin %", slot());
    expect(style.name).toBe("Margin %");
    expect(style.kind).toBe("user");
    expect(style.include).toEqual({
      number: true,
      font: true,
      patterns: true,
      border: true,
      alignment: true,
    });
    expect(Object.keys(style.include)).not.toContain("protection");
  });

  it("carries the number format, the font, the alignment and all four edges", () => {
    const style = fromPaintSlot("Margin %", slot());
    expect(style.numberFormat).toBe("0.0%");
    expect(style.font).toEqual({
      name: "Aptos",
      size: 10,
      bold: true,
      italic: false,
      color: "#1F1D1B",
    });
    expect(style.horizontalAlignment).toBe("Right");
    expect(style.borders?.bottom).toEqual(edge("Continuous", "#6D7E12"));
    expect(Object.keys(style.borders ?? {}).sort()).toEqual([
      "bottom",
      "left",
      "right",
      "top",
    ]);
  });

  it("keeps a fill and reads a missing fill as no fill", () => {
    expect(fromPaintSlot("A", slot()).fill).toBe("#EEF2D2");
    expect(fromPaintSlot("A", slot({ fill: null })).fill).toBeNull();
  });

  it("does not share the slot's objects: editing the style leaves the slot alone", () => {
    const source = slot();
    const style = fromPaintSlot("A", source);
    style.font!.bold = false;
    style.borders!.top!.color = "#FF0000";
    expect(source.font.bold).toBe(true);
    expect(source.borders.top.color).toBe("#000000");
  });
});

describe("checkStyleName", () => {
  const table = ["Normal", "Good", "Header", "Margin %"];

  it("trims and accepts a free name", () => {
    expect(checkStyleName("  Spread  ", table)).toEqual({
      ok: true,
      name: "Spread",
    });
  });

  it("refuses an empty name, in one plain sentence", () => {
    const result = checkStyleName("   ", table);
    expect(result).toEqual({
      ok: false,
      message: "Type a name for the style first.",
    });
  });

  it("allows 40 characters and refuses 41", () => {
    expect(STYLE_NAME_MAX).toBe(40);
    expect(checkStyleName("x".repeat(40), table).ok).toBe(true);
    expect(checkStyleName("x".repeat(41), table)).toEqual({
      ok: false,
      message: "A style name can be up to 40 characters.",
    });
  });

  it("counts the trimmed name, not the typed one", () => {
    expect(checkStyleName(` ${"x".repeat(40)} `, table).ok).toBe(true);
  });

  it("refuses a taken name whatever its case, built-in styles included", () => {
    expect(checkStyleName("Header", table)).toEqual({
      ok: false,
      message: "A style named Header already exists in this workbook.",
    });
    expect(checkStyleName("  gOOd ", table)).toEqual({
      ok: false,
      message: "A style named gOOd already exists in this workbook.",
    });
    expect(checkStyleName("normal", table).ok).toBe(false);
  });

  // The starters are named with this prefix and the Starter styles button writes
  // them: a user style of that name would be overwritten and filed as a starter.
  const RESERVED =
    'Style names starting with "pls,fix " are for the starter styles. Pick another name.';

  it("refuses a name in the starters' own prefix, whatever its case", () => {
    for (const name of ["pls,fix Margin", " PLS,FIX margin ", "pls,fix  x"]) {
      expect(checkStyleName(name, table), name).toEqual({
        ok: false,
        message: RESERVED,
      });
    }
  });

  it("says it is reserved, not taken, for a starter's own name", () => {
    expect(
      checkStyleName("pls,fix Title", [...table, "pls,fix Title"]),
    ).toEqual({ ok: false, message: RESERVED });
  });

  it("lets a name that only looks like the prefix through", () => {
    for (const name of [
      "pls,fix",
      "pls,fixed",
      "my pls,fix style",
      "pls fix X",
    ]) {
      expect(checkStyleName(name, table).ok, name).toBe(true);
    }
  });
});
