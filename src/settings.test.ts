import { describe, expect, it } from "vitest";
import {
  contrastText,
  currencyNumberFormat,
  DEFAULT_SETTINGS,
  deriveTheme,
  extractPaletteFromPixels,
  normalizeHex,
  parsePalette,
  readStoredSettings,
  serializeSettings,
  shade,
  tint,
} from "./settings";

// WCAG 2.x contrast ratio between two #RRGGBB colours, for the readability checks.
function luminance(hex: string): number {
  const linear = [1, 3, 5].map((at) => {
    const channel = parseInt(hex.slice(at, at + 2), 16) / 255;
    return channel <= 0.03928
      ? channel / 12.92
      : ((channel + 0.055) / 1.055) ** 2.4;
  });
  return 0.2126 * linear[0]! + 0.7152 * linear[1]! + 0.0722 * linear[2]!;
}
function contrastRatio(a: string, b: string): number {
  const [light, dark] = [luminance(a), luminance(b)].sort((x, y) => y - x);
  return (light! + 0.05) / (dark! + 0.05);
}

describe("hex handling", () => {
  it("normalizes hex colors to #RRGGBB uppercase", () => {
    expect(normalizeHex("#b27e54")).toBe("#B27E54");
    expect(normalizeHex("b27e54")).toBe("#B27E54");
    expect(normalizeHex("#abc")).toBe("#AABBCC");
  });

  it("rejects invalid hex colors", () => {
    expect(normalizeHex("")).toBeNull();
    expect(normalizeHex("#12")).toBeNull();
    expect(normalizeHex("xyzxyz")).toBeNull();
  });
});

describe("color math", () => {
  it("tints toward white and shades toward black", () => {
    expect(tint("#000000", 0.5)).toBe("#808080");
    expect(shade("#FFFFFF", 0.5)).toBe("#808080");
  });

  it("picks readable text color for a background", () => {
    expect(contrastText("#282623")).toBe("#FFFFFF");
    expect(contrastText("#B27E54")).toBe("#FFFFFF");
    expect(contrastText("#E8EEF3")).toBe("#1A1918");
  });

  // The highlighter palette (D4): ink and olive are the dark fills, the lime and
  // the light grey are the light ones, so they take the dark text.
  it("picks white on the shipped ink and olive, dark text on lime and light grey", () => {
    expect(contrastText("#111111")).toBe("#FFFFFF");
    expect(contrastText("#6D7E12")).toBe("#FFFFFF");
    expect(contrastText("#B9D62E")).toBe("#1A1918");
    expect(contrastText("#D9F24A")).toBe("#1A1918");
    expect(contrastText("#D3D3CC")).toBe("#1A1918");
  });

  it("reads at 4.5:1 or better on every default fill", () => {
    const fills = [
      DEFAULT_SETTINGS.primary,
      DEFAULT_SETTINGS.accent,
      "#B9D62E",
      "#D3D3CC",
    ];
    for (const fill of fills) {
      expect(
        contrastRatio(fill, contrastText(fill)),
        fill,
      ).toBeGreaterThanOrEqual(4.5);
    }
  });
});

describe("theme derivation", () => {
  it("derives all workbook colors from the palette", () => {
    const theme = deriveTheme(DEFAULT_SETTINGS);
    expect(theme.titleFill).toBe(DEFAULT_SETTINGS.primary);
    expect(theme.titleText).toBe(contrastText(DEFAULT_SETTINGS.primary));
    expect(theme.titleText).toBe("#FFFFFF");
    expect(theme.headerFill).toBe(tint(DEFAULT_SETTINGS.primary, 0.92));
    expect(theme.headerBorder).toBe(tint(DEFAULT_SETTINGS.primary, 0.55));
    expect(theme.resultFill).toBe(tint(DEFAULT_SETTINGS.accent, 0.86));
    expect(theme.resultBorder).toBe(shade(DEFAULT_SETTINGS.accent, 0.25));
    expect(theme.inputFont).toBe(DEFAULT_SETTINGS.input);
    expect(theme.formulaFont).toBe(DEFAULT_SETTINGS.formula);
    expect(theme.linkFont).toBe(DEFAULT_SETTINGS.link);
  });

  it("gives the autocolor link classes their own fonts", () => {
    const theme = deriveTheme(DEFAULT_SETTINGS);
    expect(theme.externalFont).toBe(DEFAULT_SETTINGS.external);
    expect(theme.partialFont).toBe(DEFAULT_SETTINGS.partial);
  });
});

describe("shipped palette", () => {
  it("ships ink and olive, the highlighter palette's own two colors", () => {
    expect(DEFAULT_SETTINGS.primary).toBe("#111111");
    expect(DEFAULT_SETTINGS.accent).toBe("#6D7E12");
  });

  // Blue typed numbers, black formulas, green links, red external, purple
  // partial: the modelling convention, not part of the brand.
  it("keeps the modelling colors", () => {
    expect(DEFAULT_SETTINGS.input).toBe("#0057B8");
    expect(DEFAULT_SETTINGS.formula).toBe("#1F1D1B");
    expect(DEFAULT_SETTINGS.link).toBe("#17823B");
    expect(DEFAULT_SETTINGS.external).toBe("#C00000");
    expect(DEFAULT_SETTINGS.partial).toBe("#7A3E9D");
  });
});

describe("autocolor defaults", () => {
  it("ships a red external color, a violet partial color and no live coloring", () => {
    expect(DEFAULT_SETTINGS.external).toBe("#C00000");
    expect(DEFAULT_SETTINGS.partial).toBe("#7A3E9D");
    expect(DEFAULT_SETTINGS.autocolorOnEdit).toBe(false);
  });
});

describe("number formats", () => {
  it("builds a currency format around the symbol", () => {
    expect(currencyNumberFormat("€", "en")).toBe("€ #,##0;[Red](€ #,##0);-");
    expect(currencyNumberFormat("$", "en")).toBe("$ #,##0;[Red]($ #,##0);-");
  });

  it("drops the symbol when none is configured", () => {
    expect(currencyNumberFormat("", "en")).toBe("#,##0;[Red](#,##0);-");
  });
});

describe("palette import and export", () => {
  it("round-trips settings through JSON", () => {
    const json = serializeSettings(DEFAULT_SETTINGS);
    expect(parsePalette(json)).toEqual(DEFAULT_SETTINGS);
  });

  it("merges partial palettes over defaults and normalizes hex", () => {
    const parsed = parsePalette('{"primary":"b27e54","unknown":1}');
    expect(parsed).toEqual({
      ...DEFAULT_SETTINGS,
      primary: "#B27E54",
      brand: ["#B27E54", DEFAULT_SETTINGS.accent],
    });
  });

  it("rejects malformed JSON and invalid colors", () => {
    expect(parsePalette("not json")).toBeNull();
    expect(parsePalette('{"primary":"nope"}')).toBeNull();
    expect(parsePalette('{"currency":"€€€€"}')).toBeNull();
    expect(parsePalette('{"external":"nope"}')).toBeNull();
    expect(parsePalette('{"partial":"#12"}')).toBeNull();
  });

  it("round-trips the autocolor-on-edit flag and rejects non-booleans", () => {
    const live = { ...DEFAULT_SETTINGS, autocolorOnEdit: true };
    expect(parsePalette(serializeSettings(live))).toEqual(live);
    expect(parsePalette('{"autocolorOnEdit":"yes"}')).toBeNull();
  });

  it("keeps palettes stored before the external and partial colors existed", () => {
    const stored = JSON.stringify({
      primary: "#14213D",
      accent: "#2EC4B6",
      input: "#0057B8",
      formula: "#1F1D1B",
      link: "#17823B",
      font: "Aptos",
      currency: "€",
    });
    // The saved navy and mint win; only the keys the file never had come from
    // the defaults. Nothing rewrites a saved palette to the shipped ink and olive.
    expect(parsePalette(stored)).toEqual({
      ...DEFAULT_SETTINGS,
      primary: "#14213D",
      accent: "#2EC4B6",
      brand: ["#14213D", "#2EC4B6"],
    });
  });
});

describe("stored settings", () => {
  it("falls back to defaults for missing or corrupt storage", () => {
    expect(readStoredSettings(null)).toEqual(DEFAULT_SETTINGS);
    expect(readStoredSettings("{broken")).toEqual(DEFAULT_SETTINGS);
    const custom = {
      ...DEFAULT_SETTINGS,
      accent: "#112233",
      brand: [DEFAULT_SETTINGS.primary, "#112233"],
    };
    expect(readStoredSettings(serializeSettings(custom))).toEqual(custom);
  });

  it("does not migrate a palette saved with the old navy and mint", () => {
    const saved = {
      ...DEFAULT_SETTINGS,
      primary: "#14213D",
      accent: "#2EC4B6",
      brand: ["#14213D", "#2EC4B6"],
    };
    const read = readStoredSettings(serializeSettings(saved));
    expect(read).toEqual(saved);
    expect(read.primary).not.toBe(DEFAULT_SETTINGS.primary);
    expect(read.accent).not.toBe(DEFAULT_SETTINGS.accent);
  });
});

describe("logo color extraction", () => {
  const pixel = (r: number, g: number, b: number, a = 255) => [r, g, b, a];

  it("returns dominant colors, skipping white background and transparency", () => {
    const data = new Uint8ClampedArray(
      [
        ...Array.from({ length: 10 }, () => pixel(178, 126, 84)),
        ...Array.from({ length: 5 }, () => pixel(40, 38, 35)),
        ...Array.from({ length: 3 }, () => pixel(255, 255, 255)),
        ...Array.from({ length: 2 }, () => pixel(255, 0, 0, 40)),
      ].flat(),
    );

    expect(extractPaletteFromPixels(data, 4)).toEqual(["#B27E54", "#282623"]);
  });

  it("caps the number of returned colors", () => {
    const data = new Uint8ClampedArray(
      [
        ...Array.from({ length: 4 }, () => pixel(178, 126, 84)),
        ...Array.from({ length: 3 }, () => pixel(40, 38, 35)),
        ...Array.from({ length: 2 }, () => pixel(0, 87, 184)),
      ].flat(),
    );

    expect(extractPaletteFromPixels(data, 2)).toEqual(["#B27E54", "#282623"]);
  });
});

describe("language", () => {
  it("defaults to Latvian, keeps a known language and refuses an unknown one", () => {
    expect(DEFAULT_SETTINGS.language).toBe("lv");
    expect(parsePalette(JSON.stringify({ language: "ru" }))?.language).toBe(
      "ru",
    );
    expect(parsePalette(JSON.stringify({ language: "de" }))).toBeNull();
    // A palette that names something else keeps the default language; one
    // naming nothing at all is not a palette (see settings.audit.test.ts).
    expect(parsePalette('{"primary":"#14213D"}')?.language).toBe("lv");
  });
});
