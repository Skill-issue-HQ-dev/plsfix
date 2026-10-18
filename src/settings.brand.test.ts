// The ordered brand list on BrandSettings: the shipped default, what
// serializeSettings writes and every way parsePalette reads it back, including
// the older pane that edits only primary / accent and drops `brand` on write.
import { describe, expect, it } from "vitest";
import {
  DEFAULT_SETTINGS,
  parsePalette,
  readStoredSettings,
  serializeSettings,
} from "./settings";

const GOLD = "#B27E54";
const MARBLE = "#282623";
const GREEN = "#273631";
const SLATE = "#41564F";

function palette(source: Record<string, unknown>) {
  const parsed = parsePalette(JSON.stringify(source));
  if (!parsed) throw new Error("expected a palette");
  return parsed;
}

describe("the shipped brand list", () => {
  it("is the primary and the accent, in that order", () => {
    expect(DEFAULT_SETTINGS.brand).toEqual([
      DEFAULT_SETTINGS.primary,
      DEFAULT_SETTINGS.accent,
    ]);
  });

  it("cannot be changed in place by a stray push or splice", () => {
    expect(() => DEFAULT_SETTINGS.brand.push("#000000")).toThrow();
    expect(() => DEFAULT_SETTINGS.brand.splice(0, 1)).toThrow();
    expect(DEFAULT_SETTINGS.brand).toHaveLength(2);
  });
});

describe("serializeSettings", () => {
  it("writes the brand list beside every existing key, in the same JSON", () => {
    const json = JSON.parse(serializeSettings(DEFAULT_SETTINGS)) as Record<
      string,
      unknown
    >;
    expect(json.brand).toEqual([
      DEFAULT_SETTINGS.primary,
      DEFAULT_SETTINGS.accent,
    ]);
    expect(Object.keys(json)).toEqual(Object.keys(DEFAULT_SETTINGS));
    expect(Object.keys(json).slice(0, 3)).toEqual([
      "primary",
      "accent",
      "brand",
    ]);
  });

  it("round-trips a ten-color palette through parsePalette", () => {
    const ten = Array.from(
      { length: 10 },
      (_, at) => `#${String(at + 1).padStart(2, "0")}A0C0`,
    );
    const settings = {
      ...DEFAULT_SETTINGS,
      primary: ten[0]!,
      accent: ten[1]!,
      brand: ten,
    };
    expect(parsePalette(serializeSettings(settings))).toEqual(settings);
  });
});

describe("parsePalette: the brand list", () => {
  it("reads a valid list, upper-cased, and takes primary and accent from its first two", () => {
    const parsed = palette({ brand: ["#b27e54", "282623", "#abc"] });
    expect(parsed.brand).toEqual([GOLD, MARBLE, "#AABBCC"]);
    expect(parsed.primary).toBe(GOLD);
    expect(parsed.accent).toBe(MARBLE);
  });

  it("de-duplicates case-insensitively and keeps the first of each", () => {
    expect(
      palette({ brand: [GOLD, "#b27e54", MARBLE, GOLD, GREEN] }).brand,
    ).toEqual([GOLD, MARBLE, GREEN]);
  });

  it("caps the list at ten colors", () => {
    const twelve = Array.from(
      { length: 12 },
      (_, at) => `#${String(at + 1).padStart(2, "0")}00FF`,
    );
    const parsed = palette({ brand: twelve });
    expect(parsed.brand).toHaveLength(10);
    expect(parsed.brand).toEqual(twelve.slice(0, 10));
  });

  it("skips entries that are not colors and keeps the rest", () => {
    const parsed = palette({ brand: [GOLD, 42, "zzz", null, MARBLE] });
    expect(parsed.brand).toEqual([GOLD, MARBLE]);
  });

  it("treats a list with fewer than two valid colors as absent", () => {
    const parsed = palette({
      primary: GREEN,
      accent: SLATE,
      brand: [GOLD, "nope"],
    });
    expect(parsed.brand).toEqual([GREEN, SLATE]);
    expect(palette({ brand: [GOLD] }).brand).toEqual([
      DEFAULT_SETTINGS.primary,
      DEFAULT_SETTINGS.accent,
    ]);
    expect(palette({ brand: [GOLD, GOLD.toLowerCase()] }).brand).toEqual([
      DEFAULT_SETTINGS.primary,
      DEFAULT_SETTINGS.accent,
    ]);
  });

  it("treats a brand that is not a list as absent", () => {
    expect(
      palette({ primary: GOLD, accent: MARBLE, brand: "nope" }).brand,
    ).toEqual([GOLD, MARBLE]);
    expect(palette({ primary: GOLD, brand: { 0: GREEN } }).brand).toEqual([
      GOLD,
      DEFAULT_SETTINGS.accent,
    ]);
  });

  it("builds the list from primary and accent when the JSON has none", () => {
    expect(palette({ primary: GOLD, accent: MARBLE }).brand).toEqual([
      GOLD,
      MARBLE,
    ]);
    expect(palette({ primary: GOLD }).brand).toEqual([
      GOLD,
      DEFAULT_SETTINGS.accent,
    ]);
    expect(palette({ font: "Arial" }).brand).toEqual([
      DEFAULT_SETTINGS.primary,
      DEFAULT_SETTINGS.accent,
    ]);
  });

  it("lets a primary in the JSON win position 0 (an older pane drops brand on write)", () => {
    const parsed = palette({ primary: GREEN, brand: [GOLD, MARBLE, SLATE] });
    expect(parsed.brand).toEqual([GREEN, MARBLE, SLATE]);
    expect(parsed.primary).toBe(GREEN);
    expect(parsed.accent).toBe(MARBLE);
  });

  it("lets primary and accent in the JSON win positions 0 and 1", () => {
    const parsed = palette({
      primary: GREEN,
      accent: SLATE,
      brand: [GOLD, MARBLE, "#DEF970"],
    });
    expect(parsed.brand).toEqual([GREEN, SLATE, "#DEF970"]);
  });

  it("drops a later color the overriding head already holds", () => {
    const parsed = palette({ primary: SLATE, brand: [GOLD, MARBLE, SLATE] });
    expect(parsed.brand).toEqual([SLATE, MARBLE]);
    expect(parsed.accent).toBe(MARBLE);
  });

  it("keeps primary and accent apart from the list when both name the same color", () => {
    const parsed = palette({
      primary: GOLD,
      accent: GOLD,
      brand: [GREEN, SLATE, MARBLE],
    });
    expect(parsed.brand).toEqual([GOLD, GOLD, MARBLE]);
  });

  // A palette saved by v3.1.0 to v3.5.0 with Primary set apart to the Accent's color: it still
  // loads, repeat and all, and keeps every extra color it had.
  it("keeps the extras of a palette stored with primary equal to accent", () => {
    const parsed = palette({
      primary: GOLD,
      accent: GOLD,
      brand: [GOLD, GOLD, MARBLE, GREEN],
    });
    expect(parsed.primary).toBe(GOLD);
    expect(parsed.accent).toBe(GOLD);
    expect(parsed.brand).toEqual([GOLD, GOLD, MARBLE, GREEN]);
  });

  it("keeps the extras past a list that repeats its head, whatever the keys name", () => {
    const parsed = palette({
      primary: SLATE,
      accent: GREEN,
      brand: [GOLD, GOLD, MARBLE],
    });
    expect(parsed.brand).toEqual([SLATE, GREEN, MARBLE]);
  });

  it("loads a primary key equal to the list's second color, with the rest of the list", () => {
    const parsed = palette({ primary: MARBLE, brand: [GOLD, MARBLE, GREEN] });
    expect(parsed.primary).toBe(MARBLE);
    expect(parsed.accent).toBe(MARBLE);
    expect(parsed.brand).toEqual([MARBLE, MARBLE, GREEN]);
  });

  it("holds a list that repeats its head and runs past ten to ten colors", () => {
    const twelve = Array.from(
      { length: 12 },
      (_, at) => `#${String(at + 1).padStart(2, "0")}00FF`,
    );
    const parsed = palette({
      primary: GOLD,
      accent: GOLD,
      brand: [GOLD, GOLD, ...twelve],
    });
    expect(parsed.brand).toEqual([GOLD, GOLD, ...twelve.slice(0, 8)]);
  });

  it("recognises a palette that carries only a brand list", () => {
    expect(
      parsePalette(JSON.stringify({ brand: [GOLD, MARBLE] })),
    ).not.toBeNull();
  });

  it("still refuses JSON naming none of the palette's keys", () => {
    expect(parsePalette('{"name":"not a palette"}')).toBeNull();
    expect(parsePalette("[1,2]")).toBeNull();
  });

  it("still refuses a primary that is not a color", () => {
    expect(
      parsePalette(JSON.stringify({ primary: "nope", brand: [GOLD, MARBLE] })),
    ).toBeNull();
  });

  it("never hands out the shipped default's own list", () => {
    const parsed = palette({ font: "Arial" });
    expect(parsed.brand).not.toBe(DEFAULT_SETTINGS.brand);
    expect(Object.isFrozen(parsed.brand)).toBe(false);
  });
});

describe("readStoredSettings", () => {
  it("lifts a palette stored before the brand list existed", () => {
    const legacy = JSON.stringify({
      primary: GOLD,
      accent: MARBLE,
      input: "#0057B8",
    });
    const stored = readStoredSettings(legacy);
    expect(stored.brand).toEqual([GOLD, MARBLE]);
  });

  it("falls back to the shipped list for nothing or garbage", () => {
    expect(readStoredSettings(null).brand).toEqual(DEFAULT_SETTINGS.brand);
    expect(readStoredSettings("{not json").brand).toEqual(
      DEFAULT_SETTINGS.brand,
    );
  });
});
