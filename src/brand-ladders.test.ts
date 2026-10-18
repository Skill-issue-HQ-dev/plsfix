// The fill and font ladders the Ctrl+Shift+F / G cycles step through: today's
// steps exactly for a two-color palette, the extra brand colors appended, and a
// color already on the ladder skipped so the cycle can never stall on it.
import { describe, expect, it } from "vitest";
import { fillLadder, fontLadder } from "./brand-ladders";
import {
  buildFillCycle,
  buildFontCycle,
  CLEAR_FILL,
  nextInCycle,
} from "./cycles";
import { DEFAULT_SETTINGS, deriveTheme, type BrandSettings } from "./settings";

const GOLD = "#B27E54";
const MARBLE = "#282623";
const GREEN = "#273631";
const SLATE = "#41564F";
const NEON = "#DEF970";

function palette(...brand: string[]): BrandSettings {
  return { ...DEFAULT_SETTINGS, primary: brand[0]!, accent: brand[1]!, brand };
}

const theme = (settings: BrandSettings) => deriveTheme(settings);

describe("the two-color ladders (today's cycles)", () => {
  it("fill: header tint, result tint, accent, primary", () => {
    const t = theme(DEFAULT_SETTINGS);
    expect(fillLadder(DEFAULT_SETTINGS).map((step) => step.color)).toEqual([
      t.headerFill,
      t.resultFill,
      DEFAULT_SETTINGS.accent,
      DEFAULT_SETTINGS.primary,
    ]);
  });

  it("font: formulas, inputs, links, accent, primary", () => {
    expect(fontLadder(DEFAULT_SETTINGS).map((step) => step.color)).toEqual([
      DEFAULT_SETTINGS.formula,
      DEFAULT_SETTINGS.input,
      DEFAULT_SETTINGS.link,
      DEFAULT_SETTINGS.accent,
      DEFAULT_SETTINGS.primary,
    ]);
  });

  it("buildFillCycle ends on the clear step, buildFontCycle does not", () => {
    expect(buildFillCycle(DEFAULT_SETTINGS).at(-1)).toBe(CLEAR_FILL);
    expect(buildFontCycle(DEFAULT_SETTINGS)).toEqual(
      fontLadder(DEFAULT_SETTINGS).map((step) => step.color),
    );
  });
});

describe("a palette with more brand colors", () => {
  const five = palette(GOLD, MARBLE, GREEN, SLATE, NEON);

  it("appends brand colors 3 and up after primary, in list order", () => {
    const t = theme(five);
    expect(buildFillCycle(five)).toEqual([
      t.headerFill,
      t.resultFill,
      MARBLE,
      GOLD,
      GREEN,
      SLATE,
      NEON,
      CLEAR_FILL,
    ]);
    expect(buildFontCycle(five)).toEqual([
      five.formula,
      five.input,
      five.link,
      MARBLE,
      GOLD,
      GREEN,
      SLATE,
      NEON,
    ]);
  });

  it("names each appended step by its place in the brand list", () => {
    expect(
      fillLadder(five)
        .slice(4)
        .map((step) => step.slot),
    ).toEqual([2, 3, 4]);
    expect(
      fontLadder(five)
        .slice(5)
        .map((step) => step.slot),
    ).toEqual([2, 3, 4]);
  });

  it("takes the head from primary and accent, whatever the stored list says", () => {
    const stale = { ...five, primary: SLATE, brand: [GOLD, MARBLE, GREEN] };
    expect(buildFontCycle(stale).slice(3)).toEqual([MARBLE, SLATE, GREEN]);
  });

  it("holds ten brand colors", () => {
    const ten = palette(
      ...Array.from(
        { length: 10 },
        (_, at) => `#${String(at + 1).padStart(2, "0")}A0C0`,
      ),
    );
    expect(buildFillCycle(ten)).toHaveLength(4 + 8 + 1);
    expect(buildFontCycle(ten)).toHaveLength(5 + 8);
  });
});

describe("a color the ladder already holds", () => {
  it("is skipped when a brand color repeats an earlier step", () => {
    const t = theme(DEFAULT_SETTINGS);
    const repeating = palette(
      DEFAULT_SETTINGS.primary,
      DEFAULT_SETTINGS.accent,
      t.headerFill,
      NEON,
      DEFAULT_SETTINGS.accent.toLowerCase(),
    );
    expect(buildFillCycle(repeating)).toEqual([
      t.headerFill,
      t.resultFill,
      DEFAULT_SETTINGS.accent,
      DEFAULT_SETTINGS.primary,
      NEON,
      CLEAR_FILL,
    ]);
  });

  it("skips a brand color equal to the input blue on the font ladder", () => {
    const blue = palette(GOLD, MARBLE, DEFAULT_SETTINGS.input, GREEN);
    expect(buildFontCycle(blue)).toEqual([
      blue.formula,
      blue.input,
      blue.link,
      MARBLE,
      GOLD,
      GREEN,
    ]);
  });

  it("keeps one step for a primary equal to the accent, named Accent", () => {
    const same = palette(GOLD, GOLD, MARBLE);
    expect(fontLadder(same).map((step) => step.slot)).toEqual([
      "formula",
      "input",
      "link",
      "accent",
      2,
    ]);
    expect(buildFillCycle(same).filter((color) => color === GOLD)).toHaveLength(
      1,
    );
  });

  it("never stalls: stepping from every step lands on a different one", () => {
    const messy = palette(
      GOLD,
      GOLD,
      MARBLE,
      MARBLE.toLowerCase(),
      DEFAULT_SETTINGS.link,
    );
    for (const cycle of [buildFillCycle(messy), buildFontCycle(messy)]) {
      for (const step of cycle) {
        expect(nextInCycle(step, cycle), step).not.toBe(step);
      }
      expect(new Set(cycle.map((step) => step.toUpperCase())).size).toBe(
        cycle.length,
      );
    }
  });

  it("walks a whole cycle and comes back to where it started", () => {
    const five = palette(GOLD, MARBLE, GREEN, SLATE, NEON);
    const cycle = buildFillCycle(five);
    const seen = [cycle[0]!];
    while (seen.length < cycle.length) {
      seen.push(nextInCycle(seen.at(-1)!, cycle));
    }
    expect(seen).toEqual(cycle);
    expect(nextInCycle(seen.at(-1)!, cycle)).toBe(cycle[0]);
  });
});
