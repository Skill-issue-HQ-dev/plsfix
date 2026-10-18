// The brand list as an Excel copy carries it: each color once, the first kept, in
// order; nothing at all when fewer than two different colors remain. Split out of
// palette.test.ts, which stands over the line cap.
import { describe, expect, it } from "vitest";
import { brandColors, distinctBrandColors, withBrandColor } from "./palette";
import { DEFAULT_SETTINGS, type BrandSettings } from "./settings";

const GOLD = "#B27E54";
const MARBLE = "#282623";
const GREEN = "#273631";
const SLATE = "#41564F";
const NEON = "#DEF970";

const fiveColors: BrandSettings = {
  ...DEFAULT_SETTINGS,
  primary: GOLD,
  accent: MARBLE,
  brand: [GOLD, MARBLE, GREEN, SLATE, NEON],
};

describe("distinctBrandColors", () => {
  it("is the brand list when every color differs", () => {
    expect(distinctBrandColors(fiveColors)).toEqual([
      GOLD,
      MARBLE,
      GREEN,
      SLATE,
      NEON,
    ]);
  });

  it("keeps the first of a color set twice: primary and accent are set apart", () => {
    const twice = withBrandColor(
      withBrandColor(fiveColors, 0, MARBLE)!,
      5,
      MARBLE,
    )!;
    expect(brandColors(twice)).toEqual([
      MARBLE,
      MARBLE,
      GREEN,
      SLATE,
      NEON,
      MARBLE,
    ]);
    expect(distinctBrandColors(twice)).toEqual([MARBLE, GREEN, SLATE, NEON]);
  });

  it("folds a color that differs only in case, and answers upper case", () => {
    const folded = { ...fiveColors, accent: GOLD.toLowerCase() };
    expect(distinctBrandColors(folded)).toEqual([GOLD, GREEN, SLATE, NEON]);
  });

  it("is null when fewer than two different colors are left: nothing to carry", () => {
    const one = { ...DEFAULT_SETTINGS, primary: GOLD, accent: GOLD, brand: [] };
    expect(distinctBrandColors(one)).toBeNull();
  });
});
