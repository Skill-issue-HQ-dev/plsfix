// The shape styles' pure core, part one: the six starters drawn from the brand
// palette and the naming rules. No Office.js. The list's rules are in
// shape-styles.list.test.ts, the stored codec in shape-styles-codec.test.ts.

import { describe, expect, it } from "vitest";
import { STYLE_NAME_MAX } from "./custom-styles";
import {
  checkShapeStyleName,
  STARTER_SHAPE_NAMES,
  starterShapeStyles,
  type ShapeStyle,
} from "./shape-styles";
import {
  contrastText,
  DEFAULT_SETTINGS,
  tint,
  type BrandSettings,
} from "./settings";

const GOLD = "#B27E54";
const MARBLE = "#282623";
const GREEN = "#273631";
const SLATE = "#41564F";

function settingsOf(brand: string[]): BrandSettings {
  return {
    ...DEFAULT_SETTINGS,
    primary: brand[0]!,
    accent: brand[1]!,
    brand,
  };
}

const find = (styles: ShapeStyle[], name: string): ShapeStyle =>
  styles.find((style) => style.name === name)!;

describe("starterShapeStyles", () => {
  const starters = starterShapeStyles(settingsOf([GOLD, MARBLE, GREEN, SLATE]));

  it("makes exactly these six, in this order, all starters", () => {
    expect(starters.map((style) => style.name)).toEqual([
      "Brand fill",
      "Accent fill",
      "Soft fill",
      "Brand outline",
      "Callout",
      "Highlight",
    ]);
    expect([...STARTER_SHAPE_NAMES]).toEqual(starters.map((s) => s.name));
    expect(new Set(starters.map((style) => style.kind))).toEqual(
      new Set(["starter"]),
    );
  });

  it("Brand fill: the first brand color, no outline, readable text", () => {
    expect(find(starters, "Brand fill")).toEqual({
      name: "Brand fill",
      kind: "starter",
      fill: GOLD,
      line: { visible: false },
      text: { color: contrastText(GOLD) },
    });
  });

  it("Accent fill: the second brand color, no outline, readable text", () => {
    expect(find(starters, "Accent fill")).toEqual({
      name: "Accent fill",
      kind: "starter",
      fill: MARBLE,
      line: { visible: false },
      text: { color: contrastText(MARBLE) },
    });
  });

  it("Soft fill: a tint of the first color at 0.85, no outline, text in that color", () => {
    expect(find(starters, "Soft fill")).toEqual({
      name: "Soft fill",
      kind: "starter",
      fill: tint(GOLD, 0.85),
      line: { visible: false },
      text: { color: GOLD },
    });
  });

  it("Brand outline: no fill, a 1.5 pt solid outline in the first color, text in it", () => {
    expect(find(starters, "Brand outline")).toEqual({
      name: "Brand outline",
      kind: "starter",
      fill: null,
      line: { visible: true, color: GOLD, weight: 1.5, dashStyle: "Solid" },
      text: { color: GOLD },
    });
  });

  it("Callout: a tint of the second color, a 1 pt solid outline in it, the formula color for text", () => {
    const settings = settingsOf([GOLD, MARBLE, GREEN, SLATE]);
    expect(find(starters, "Callout")).toEqual({
      name: "Callout",
      kind: "starter",
      fill: tint(MARBLE, 0.85),
      line: { visible: true, color: MARBLE, weight: 1, dashStyle: "Solid" },
      text: { color: settings.formula },
    });
  });

  it("Highlight: the third color, no outline, readable text", () => {
    expect(find(starters, "Highlight")).toEqual({
      name: "Highlight",
      kind: "starter",
      fill: GREEN,
      line: { visible: false },
      text: { color: contrastText(GREEN) },
    });
  });

  it("Highlight falls back to the second color when the palette has two", () => {
    const two = starterShapeStyles(settingsOf([GOLD, MARBLE]));
    expect(find(two, "Highlight").fill).toBe(MARBLE);
    expect(find(two, "Highlight").text).toEqual({
      color: contrastText(MARBLE),
    });
  });

  it("sets no font name and no size on any starter: the deck's theme fonts stay", () => {
    for (const style of starters) {
      expect(style.text, style.name).not.toHaveProperty("name");
      expect(style.text, style.name).not.toHaveProperty("size");
    }
  });

  it("follows a palette change", () => {
    const other = starterShapeStyles(settingsOf([SLATE, GREEN]));
    expect(find(other, "Brand fill").fill).toBe(SLATE);
    expect(find(other, "Accent fill").fill).toBe(GREEN);
  });

  it("draws from the shipped palette when that is the palette", () => {
    const shipped = starterShapeStyles(DEFAULT_SETTINGS);
    expect(find(shipped, "Brand fill").fill).toBe(DEFAULT_SETTINGS.primary);
    expect(find(shipped, "Accent fill").fill).toBe(DEFAULT_SETTINGS.accent);
  });

  it("makes every color a #RRGGBB the host accepts", () => {
    for (const style of starters) {
      for (const color of [style.fill, style.line?.color, style.text?.color]) {
        if (typeof color === "string") expect(color).toMatch(/^#[0-9A-F]{6}$/);
      }
    }
  });
});

describe("checkShapeStyleName", () => {
  it("trims the name", () => {
    expect(checkShapeStyleName("  Callout  ", [])).toEqual({
      ok: true,
      name: "Callout",
    });
  });

  it("refuses an empty or blank name in one sentence", () => {
    for (const raw of ["", "   "]) {
      const check = checkShapeStyleName(raw, []);
      expect(check.ok).toBe(false);
      if (!check.ok)
        expect(check.message).toBe("Type a name for the style first.");
    }
  });

  it("takes STYLE_NAME_MAX characters and refuses one more", () => {
    expect(checkShapeStyleName("a".repeat(STYLE_NAME_MAX), []).ok).toBe(true);
    const check = checkShapeStyleName("a".repeat(STYLE_NAME_MAX + 1), []);
    expect(check.ok).toBe(false);
    if (!check.ok) expect(check.message).toContain(String(STYLE_NAME_MAX));
  });

  it("measures the trimmed name, not what was typed", () => {
    const padded = ` ${"a".repeat(STYLE_NAME_MAX)} `;
    expect(checkShapeStyleName(padded, []).ok).toBe(true);
  });

  it("refuses a name that is in the list, whatever its case", () => {
    const check = checkShapeStyleName("callout", ["Brand fill", "Callout"]);
    expect(check.ok).toBe(false);
    if (!check.ok)
      expect(check.message).toBe("A shape style named callout already exists.");
  });

  it("says each refusal as one sentence", () => {
    const refusals = [
      checkShapeStyleName("", []),
      checkShapeStyleName("x".repeat(STYLE_NAME_MAX + 1), []),
      checkShapeStyleName("Callout", ["callout"]),
    ];
    for (const check of refusals) {
      expect(check.ok).toBe(false);
      if (!check.ok) {
        expect(check.message).toMatch(/^[^.!?]+\.$/);
      }
    }
  });
});
