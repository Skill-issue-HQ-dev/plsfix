// heldElsewhere: whether another entry of the brand list (Primary and Accent included) already
// has a color, the one rule the palette editor's rows and the Brand tab's Primary and Accent
// pickers refuse a repeat by. Pure, so it is held here without a DOM.

import { describe, expect, it } from "vitest";
import { heldElsewhere } from "./palette";
import { DEFAULT_SETTINGS, type BrandSettings } from "./settings";

const GOLD = "#B27E54";
const MARBLE = "#282623";
const GREEN = "#273631";
const SLATE = "#41564F";

function palette(...brand: string[]): BrandSettings {
  return { ...DEFAULT_SETTINGS, primary: brand[0]!, accent: brand[1]!, brand };
}

const four = palette(GOLD, MARBLE, GREEN, SLATE);

describe("heldElsewhere", () => {
  it("is true for a color another entry holds, whichever entry asks", () => {
    expect(heldElsewhere(four, 0, MARBLE)).toBe(true);
    expect(heldElsewhere(four, 1, GOLD)).toBe(true);
    expect(heldElsewhere(four, 1, SLATE)).toBe(true);
    expect(heldElsewhere(four, 3, GOLD)).toBe(true);
  });

  it("counts Primary and Accent among the entries", () => {
    expect(heldElsewhere(four, 2, GOLD)).toBe(true);
    expect(heldElsewhere(four, 2, MARBLE)).toBe(true);
  });

  it("is false for the entry's own color", () => {
    for (const [index, color] of four.brand.entries()) {
      expect(heldElsewhere(four, index, color), String(index)).toBe(false);
    }
  });

  it("is false for a color the list does not hold", () => {
    expect(heldElsewhere(four, 0, "#123456")).toBe(false);
    expect(heldElsewhere(four, 1, "#123456")).toBe(false);
  });

  it("judges a color to append (an index off the list) against the whole list", () => {
    expect(heldElsewhere(four, 4, GOLD)).toBe(true);
    expect(heldElsewhere(four, 4, "#123456")).toBe(false);
  });
});
