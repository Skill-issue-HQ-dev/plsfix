// The brand colour rules every chart follows, on the sheet and on the slide:
// the series palette order, the waterfall's totals/rises/falls, a pie's cycle.

import { describe, expect, it } from "vitest";
import {
  pieColors,
  seriesColor,
  seriesPalette,
  waterfallColors,
} from "./chart-colors";
import { shade, tint } from "./settings";

const brand = {
  primary: "#14213D",
  accent: "#2EC4B6",
  external: "#B27E54",
};

describe("seriesPalette", () => {
  it("leads with the accent, then the primary, then their tints", () => {
    expect(seriesPalette(brand)).toEqual([
      brand.accent,
      brand.primary,
      tint(brand.primary, 0.55),
      tint(brand.accent, 0.45),
      tint(brand.primary, 0.78),
      tint(brand.accent, 0.7),
    ]);
  });
});

describe("waterfallColors", () => {
  it("paints the totals primary, a fall external and a rise accent", () => {
    expect(waterfallColors([100, 20, -30, 90], brand)).toEqual([
      brand.primary,
      brand.accent,
      brand.external,
      brand.primary,
    ]);
  });

  it("treats two points as two totals", () => {
    expect(waterfallColors([100, 90], brand)).toEqual([
      brand.primary,
      brand.primary,
    ]);
  });
});

describe("pieColors", () => {
  it("cycles the palette past six slices", () => {
    const colors = pieColors(8, brand);
    expect(colors).toHaveLength(8);
    expect(colors[0]).toBe(brand.accent);
    expect(colors[6]).toBe(colors[0]);
    expect(colors[7]).toBe(colors[1]);
  });
});

// Twelve series are drawn: the six palette colours first, then the same six
// shaded, so no two series of one chart wear the same colour.
describe("seriesColor", () => {
  it("is the palette's own colour for the first six series", () => {
    const palette = seriesPalette(brand);
    for (let index = 0; index < 6; index += 1) {
      expect(seriesColor(index, brand)).toBe(palette[index]);
    }
  });

  it("shades the palette for the next six", () => {
    const palette = seriesPalette(brand);
    for (let index = 6; index < 12; index += 1) {
      expect(seriesColor(index, brand)).toBe(shade(palette[index - 6]!, 0.3));
    }
  });

  it("gives twelve series twelve different colours", () => {
    const colors = Array.from({ length: 12 }, (_one, index) =>
      seriesColor(index, brand),
    );
    expect(new Set(colors).size).toBe(12);
  });

  it("keeps shading a third time round, never past black", () => {
    const third = seriesColor(12, brand);
    expect(third).toMatch(/^#[0-9A-F]{6}$/);
    expect(third).not.toBe(seriesColor(6, brand));
  });
});
