// The receipt a fill or font cycle toasts: the same step order as the cycle
// itself, so a palette with more brand colors names its extra steps Brand 3,
// Brand 4 and so on, and a skipped repeat never shifts the labels.
import { describe, expect, it } from "vitest";
import { fillLadder, fontLadder } from "./brand-ladders";
import {
  fillCycleLabels,
  FILL_CYCLE_LABELS,
  fontCycleLabels,
  FONT_CYCLE_LABELS,
} from "./cycle-labels";
import { buildFillCycle, buildFontCycle } from "./cycles";
import { DEFAULT_SETTINGS, type BrandSettings } from "./settings";

function palette(...brand: string[]): BrandSettings {
  return { ...DEFAULT_SETTINGS, primary: brand[0]!, accent: brand[1]!, brand };
}

const five = palette("#B27E54", "#282623", "#273631", "#41564F", "#DEF970");

describe("the labels of a two-color palette", () => {
  it("are the ones the cycles have always toasted", () => {
    expect(FILL_CYCLE_LABELS).toEqual([
      "Header tint",
      "Result tint",
      "Accent",
      "Primary",
      "none",
    ]);
    expect(fillCycleLabels(DEFAULT_SETTINGS)).toEqual(FILL_CYCLE_LABELS);
    expect(FONT_CYCLE_LABELS).toEqual([
      "Formulas",
      "Inputs",
      "Cross-sheet links",
      "Accent",
      "Primary",
    ]);
    expect(fontCycleLabels(DEFAULT_SETTINGS)).toEqual(FONT_CYCLE_LABELS);
  });
});

describe("the labels of a palette with more brand colors", () => {
  it("name the extra fill steps Brand 3, Brand 4, Brand 5 and end on none", () => {
    expect(fillCycleLabels(five)).toEqual([
      "Header tint",
      "Result tint",
      "Accent",
      "Primary",
      "Brand 3",
      "Brand 4",
      "Brand 5",
      "none",
    ]);
  });

  it("name the extra font steps the same way", () => {
    expect(fontCycleLabels(five)).toEqual([
      "Formulas",
      "Inputs",
      "Cross-sheet links",
      "Accent",
      "Primary",
      "Brand 3",
      "Brand 4",
      "Brand 5",
    ]);
  });

  it("stay one for one with the cycle, repeats skipped on both sides", () => {
    const repeating = palette(
      "#B27E54",
      "#282623",
      DEFAULT_SETTINGS.input,
      "#DEF970",
      "#b27e54",
    );
    expect(fillCycleLabels(repeating)).toHaveLength(
      buildFillCycle(repeating).length,
    );
    expect(fontCycleLabels(repeating)).toHaveLength(
      buildFontCycle(repeating).length,
    );
    // The blue is already Inputs on the font ladder, so only Brand 4 survives
    // of the extras there; on the fill ladder the blue is a new color.
    expect(fontCycleLabels(repeating).slice(5)).toEqual(["Brand 4"]);
    expect(fillCycleLabels(repeating).slice(4)).toEqual([
      "Brand 3",
      "Brand 4",
      "none",
    ]);
    expect(
      fillLadder(repeating)
        .map((step) => step.slot)
        .slice(4),
    ).toEqual([2, 3]);
    expect(fontLadder(repeating)).toHaveLength(6);
  });
});
