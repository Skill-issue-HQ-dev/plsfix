// The brand list's pure editing core: reading colors out of any pasted text,
// replacing / changing / removing entries without breaking the primary and
// accent head, and the five shades of a base color in Office's own order.
import { describe, expect, it } from "vitest";
import {
  brandColors,
  MAX_BRAND_COLORS,
  parseColorList,
  shadeLadder,
  withBrandColor,
  withBrandList,
  withoutBrandColor,
} from "./palette";
import {
  DEFAULT_SETTINGS,
  serializeSettings,
  type BrandSettings,
} from "./settings";

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

function channels(hex: string): number[] {
  return [1, 3, 5].map((at) => parseInt(hex.slice(at, at + 2), 16));
}

// Office rounds its own HSL maths a step differently at the odd channel, so a
// shade matches Office's published value to within one step per channel.
function expectWithinOne(actual: string, office: string): void {
  const got = channels(actual);
  const want = channels(office);
  got.forEach((channel, at) => {
    expect(
      Math.abs(channel - want[at]!),
      `${actual} against Office's ${office}`,
    ).toBeLessThanOrEqual(1);
  });
}

describe("MAX_BRAND_COLORS", () => {
  it("is ten", () => {
    expect(MAX_BRAND_COLORS).toBe(10);
  });
});

describe("brandColors", () => {
  it("is the brand list as the settings hold it", () => {
    expect(brandColors(fiveColors)).toEqual([GOLD, MARBLE, GREEN, SLATE, NEON]);
  });

  it("puts primary and accent first even when the stored head is stale", () => {
    const stale = { ...DEFAULT_SETTINGS, primary: GOLD, accent: MARBLE };
    expect(brandColors(stale)).toEqual([GOLD, MARBLE]);
  });
});

describe("parseColorList", () => {
  it("reads a realistic brand sheet in the order the colors appear", () => {
    const sheet =
      "Brand Gold #B27E54 Pantone 729 / Marble #282623 / Dark Green 273631 / rgb(65, 86, 79) / Neon #DEF970";
    expect(parseColorList(sheet)).toEqual({
      colors: [GOLD, MARBLE, GREEN, SLATE, NEON],
      dropped: 0,
    });
  });

  it("reads every spelling in one pass, keeping the order of appearance", () => {
    const { colors } = parseColorList(
      "rgba(65,86,79,0.5), #abc; 00ff00 '#B27E54' (282623) #def970",
    );
    expect(colors).toEqual([SLATE, "#AABBCC", "#00FF00", GOLD, MARBLE, NEON]);
  });

  it("reads rgb with spaces or commas and any case", () => {
    expect(parseColorList("RGB(0, 0, 0)").colors).toEqual(["#000000"]);
    expect(parseColorList("rgb(255 128 0)").colors).toEqual(["#FF8000"]);
    expect(parseColorList("rgba(  1 , 2 , 3 , .5 )").colors).toEqual([
      "#010203",
    ]);
  });

  it("skips an rgb with a channel past 255 or a non-integer channel", () => {
    expect(parseColorList("rgb(256, 0, 0)").colors).toEqual([]);
    expect(parseColorList("rgb(300, 1, 2) #B27E54").colors).toEqual([GOLD]);
    expect(parseColorList("rgb(1.5, 2, 3)").colors).toEqual([]);
    expect(parseColorList("rgb(10%, 20%, 30%)").colors).toEqual([]);
  });

  it("expands #RGB and upper-cases everything", () => {
    expect(parseColorList("#abc #b27e54").colors).toEqual(["#AABBCC", GOLD]);
  });

  it("never reads a hex inside a longer letter or digit run", () => {
    for (const text of [
      "x273631",
      "273631x",
      "1273631",
      "2736312",
      "abc273631def",
      "color:273631",
      "273631/",
      "#B27E54A",
      "#B27E54FF",
    ]) {
      expect(parseColorList(text).colors, text).toEqual([]);
    }
  });

  it("reads a bare six-digit hex only between its delimiters", () => {
    for (const text of [
      "273631",
      " 273631 ",
      "a,273631;b",
      "(273631)",
      '"273631"',
      "'273631'",
      "gold\t273631\n",
    ]) {
      expect(parseColorList(text).colors, text).toEqual([GREEN]);
    }
  });

  it("does not read the three-digit codes of a Pantone number as colors", () => {
    expect(parseColorList("Pantone 729 and 7621 C").colors).toEqual([]);
  });

  it("keeps the first of each color, whatever its spelling", () => {
    expect(
      parseColorList("#B27E54 b27e54 rgb(178, 126, 84) #b27e54").colors,
    ).toEqual([GOLD]);
  });

  it("caps the list at ten and counts the valid colors that did not fit", () => {
    const twelve = Array.from(
      { length: 12 },
      (_, at) => `#${String(at + 1).padStart(2, "0")}00FF`,
    );
    const parsed = parseColorList(twelve.join(" "));
    expect(parsed.colors).toEqual(twelve.slice(0, 10));
    expect(parsed.dropped).toBe(2);
  });

  it("does not count repeats as dropped", () => {
    const text = `${"#112233 ".repeat(15)}#445566`;
    expect(parseColorList(text)).toEqual({
      colors: ["#112233", "#445566"],
      dropped: 0,
    });
  });

  it("answers nothing for text with no color in it", () => {
    expect(parseColorList("")).toEqual({ colors: [], dropped: 0 });
    expect(parseColorList("   \n ")).toEqual({ colors: [], dropped: 0 });
    expect(parseColorList("our brand is warm")).toEqual({
      colors: [],
      dropped: 0,
    });
  });

  it("takes the brand list of palette JSON, as Copy palette JSON writes it", () => {
    expect(parseColorList(serializeSettings(fiveColors))).toEqual({
      colors: [GOLD, MARBLE, GREEN, SLATE, NEON],
      dropped: 0,
    });
  });

  it("takes a palette's primary and accent from the JSON over its list", () => {
    const json = JSON.stringify({
      primary: SLATE,
      brand: [GOLD, MARBLE, GREEN],
    });
    expect(parseColorList(json).colors).toEqual([SLATE, MARBLE, GREEN]);
  });

  it("counts what a palette JSON's list held past ten", () => {
    const twelve = Array.from(
      { length: 12 },
      (_, at) => `#${String(at + 1).padStart(2, "0")}00FF`,
    );
    const parsed = parseColorList(JSON.stringify({ brand: twelve }));
    expect(parsed.colors).toHaveLength(10);
    expect(parsed.dropped).toBe(2);
  });

  it("scans JSON that is not a palette like any other text", () => {
    expect(
      parseColorList('{"name":"theme","accent":"#B27E54x"}').colors,
    ).toEqual([]);
    expect(
      parseColorList('{"swatch":"#B27E54","also":"#282623"}').colors,
    ).toEqual([GOLD, MARBLE]);
    expect(parseColorList('"#B27E54"').colors).toEqual([GOLD]);
    expect(parseColorList("[1, 2, 3]").colors).toEqual([]);
  });
});

describe("withBrandList", () => {
  it("replaces the list and moves primary and accent with it", () => {
    const next = withBrandList(DEFAULT_SETTINGS, [GOLD, MARBLE, GREEN]);
    expect(next).not.toBeNull();
    expect(next!.brand).toEqual([GOLD, MARBLE, GREEN]);
    expect(next!.primary).toBe(GOLD);
    expect(next!.accent).toBe(MARBLE);
  });

  it("leaves every other setting and the original object alone", () => {
    const before = { ...fiveColors, font: "Arial" };
    const next = withBrandList(before, [GREEN, SLATE])!;
    expect(next.font).toBe("Arial");
    expect(next.input).toBe(before.input);
    expect(before.brand).toEqual([GOLD, MARBLE, GREEN, SLATE, NEON]);
    expect(before.primary).toBe(GOLD);
  });

  it("refuses fewer than two colors", () => {
    expect(withBrandList(DEFAULT_SETTINGS, [])).toBeNull();
    expect(withBrandList(DEFAULT_SETTINGS, [GOLD])).toBeNull();
    expect(
      withBrandList(DEFAULT_SETTINGS, [GOLD, GOLD.toLowerCase()]),
    ).toBeNull();
  });

  it("normalizes, drops repeats and entries that are not colors, and caps at ten", () => {
    const many = Array.from(
      { length: 13 },
      (_, at) => `#${String(at + 1).padStart(2, "0")}00FF`,
    );
    const next = withBrandList(DEFAULT_SETTINGS, [
      "b27e54",
      "not a color",
      "#B27E54",
      ...many,
    ])!;
    expect(next.brand).toHaveLength(MAX_BRAND_COLORS);
    expect(next.brand[0]).toBe(GOLD);
    expect(next.brand[1]).toBe("#0100FF");
  });
});

describe("withBrandColor", () => {
  it("sets primary through position 0 and accent through position 1", () => {
    const first = withBrandColor(fiveColors, 0, "#112233")!;
    expect(first.primary).toBe("#112233");
    expect(first.brand).toEqual(["#112233", MARBLE, GREEN, SLATE, NEON]);
    const second = withBrandColor(fiveColors, 1, "#445566")!;
    expect(second.accent).toBe("#445566");
    expect(second.brand).toEqual([GOLD, "#445566", GREEN, SLATE, NEON]);
  });

  it("changes an extra color without touching the head", () => {
    const next = withBrandColor(fiveColors, 3, "#abcdef")!;
    expect(next.brand).toEqual([GOLD, MARBLE, GREEN, "#ABCDEF", NEON]);
    expect(next.primary).toBe(GOLD);
    expect(next.accent).toBe(MARBLE);
  });

  it("appends at the end of the list and no further", () => {
    const next = withBrandColor(DEFAULT_SETTINGS, 2, "#abc")!;
    expect(next.brand).toEqual([
      DEFAULT_SETTINGS.primary,
      DEFAULT_SETTINGS.accent,
      "#AABBCC",
    ]);
    expect(withBrandColor(DEFAULT_SETTINGS, 3, "#abc")).toBeNull();
  });

  it("refuses an append at ten colors", () => {
    const ten = withBrandList(
      DEFAULT_SETTINGS,
      Array.from(
        { length: 10 },
        (_, at) => `#${String(at + 1).padStart(2, "0")}00FF`,
      ),
    )!;
    expect(ten.brand).toHaveLength(10);
    expect(withBrandColor(ten, 10, "#ABCDEF")).toBeNull();
    expect(withBrandColor(ten, 9, "#ABCDEF")!.brand[9]).toBe("#ABCDEF");
  });

  it("refuses a value that is not a color, and a negative or fractional index", () => {
    expect(withBrandColor(fiveColors, 2, "nope")).toBeNull();
    expect(withBrandColor(fiveColors, -1, "#112233")).toBeNull();
    expect(withBrandColor(fiveColors, 1.5, "#112233")).toBeNull();
  });

  it("does not change the settings it was given", () => {
    const frozen = Object.freeze({
      ...fiveColors,
      brand: Object.freeze([...fiveColors.brand]) as string[],
    });
    expect(() => withBrandColor(frozen, 2, "#ABCDEF")).not.toThrow();
    expect(frozen.brand[2]).toBe(GREEN);
  });
});

describe("withoutBrandColor", () => {
  it("removes an extra color and closes the gap", () => {
    const next = withoutBrandColor(fiveColors, 3)!;
    expect(next.brand).toEqual([GOLD, MARBLE, GREEN, NEON]);
  });

  it("never removes primary or accent", () => {
    expect(withoutBrandColor(fiveColors, 0)).toBeNull();
    expect(withoutBrandColor(fiveColors, 1)).toBeNull();
  });

  it("refuses an index past the end", () => {
    expect(withoutBrandColor(fiveColors, 5)).toBeNull();
    expect(withoutBrandColor(DEFAULT_SETTINGS, 2)).toBeNull();
  });
});

describe("shadeLadder", () => {
  it("gives a mid color Office's five shades: three lighter, two darker", () => {
    const ladder = shadeLadder("#4472C4");
    expect(ladder.map((shade) => shade.label)).toEqual([
      "Lighter 80%",
      "Lighter 60%",
      "Lighter 40%",
      "Darker 25%",
      "Darker 50%",
    ]);
  });

  // Excel's own palette for the Office theme's six accents, read off its color
  // picker: the maths here is Office's, so every shade sits within one step.
  it.each([
    ["#4472C4", ["#D9E2F3", "#B4C6E7", "#8FAADC", "#2F5597", "#1F3864"]],
    ["#ED7D31", ["#FBE5D6", "#F8CBAD", "#F4B183", "#C55A11", "#833C0B"]],
    ["#A5A5A5", ["#EDEDED", "#DBDBDB", "#C9C9C9", "#7C7C7C", "#525252"]],
    ["#FFC000", ["#FFF2CC", "#FFE699", "#FFD966", "#BF9000", "#7F6000"]],
    ["#5B9BD5", ["#DEEBF7", "#BDD7EE", "#9DC3E6", "#2E75B6", "#1F4E79"]],
    ["#70AD47", ["#E2F0D9", "#C5E0B4", "#A9D18E", "#548235", "#375623"]],
  ])("matches Office's shades of %s", (base, office) => {
    const ladder = shadeLadder(base);
    expect(ladder).toHaveLength(5);
    ladder.forEach((shade, at) => expectWithinOne(shade.hex, office[at]!));
  });

  it("returns upper-case #RRGGBB", () => {
    for (const shade of shadeLadder("#b27e54")) {
      expect(shade.hex).toMatch(/^#[0-9A-F]{6}$/);
    }
  });

  it("gives a dark base only lighter shades, from 90 % down to 10 %", () => {
    const ladder = shadeLadder("#111111");
    expect(ladder.map((shade) => shade.label)).toEqual([
      "Lighter 90%",
      "Lighter 75%",
      "Lighter 50%",
      "Lighter 25%",
      "Lighter 10%",
    ]);
    // Black lightened by x % is the grey at x % of white.
    const black = shadeLadder("#000000");
    ["#E6E6E6", "#BFBFBF", "#808080", "#404040", "#1A1A1A"].forEach(
      (grey, at) => expectWithinOne(black[at]!.hex, grey),
    );
  });

  it("gives a light base only darker shades, from 10 % to 90 %", () => {
    const ladder = shadeLadder("#F2F2F2");
    expect(ladder.map((shade) => shade.label)).toEqual([
      "Darker 10%",
      "Darker 25%",
      "Darker 50%",
      "Darker 75%",
      "Darker 90%",
    ]);
    const white = shadeLadder("#FFFFFF");
    ["#E6E6E6", "#BFBFBF", "#808080", "#404040", "#1A1A1A"].forEach(
      (grey, at) => expectWithinOne(white[at]!.hex, grey),
    );
  });

  it("changes regime strictly past 0.2 and 0.8 of lightness", () => {
    const labels = (hex: string): string => shadeLadder(hex)[0]!.label;
    expect(labels("#323232")).toBe("Lighter 90%"); // L = 50 / 255, under 0.2
    expect(labels("#333333")).toBe("Lighter 80%"); // L = 0.2 exactly
    expect(labels("#CCCCCC")).toBe("Lighter 80%"); // L = 0.8 exactly
    expect(labels("#CDCDCD")).toBe("Darker 10%"); // just over 0.8
  });

  it("keeps each shade's hue family: lighter is lighter, darker is darker", () => {
    const mean = (hex: string): number =>
      channels(hex).reduce((total, part) => total + part, 0) / 3;
    const ladder = shadeLadder(GOLD);
    const base = mean(GOLD);
    expect(mean(ladder[0]!.hex)).toBeGreaterThan(mean(ladder[1]!.hex));
    expect(mean(ladder[1]!.hex)).toBeGreaterThan(mean(ladder[2]!.hex));
    expect(mean(ladder[2]!.hex)).toBeGreaterThan(base);
    expect(mean(ladder[3]!.hex)).toBeLessThan(base);
    expect(mean(ladder[4]!.hex)).toBeLessThan(mean(ladder[3]!.hex));
  });

  it("answers nothing for a value that is not a color", () => {
    expect(shadeLadder("nope")).toEqual([]);
  });
});
