// The Ctrl+Shift+F / G cycles over the real adapter and the strict fake host
// with a palette of five brand colors: the extra colors are steps, the receipt
// names them "Brand 3", "Brand 4", "Brand 5", and a repeated color never stalls
// the cycle. The two-color walks stay in host.integration.test.ts untouched.
import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  enableStrictLoadSemantics,
  type FakeHelpers,
  installFakeHost,
  uninstallFakeHost,
} from "./fakehost";
import type * as ExcelModule from "../src/excel";
import type * as SettingsModule from "../src/settings";

enableStrictLoadSemantics();

let helpers: FakeHelpers;
let smt: typeof ExcelModule;
let settings: typeof SettingsModule;

const GOLD = "#B27E54";
const MARBLE = "#282623";
const GREEN = "#273631";
const SLATE = "#41564F";
const NEON = "#DEF970";

async function boot(brand: string[]): Promise<void> {
  vi.resetModules();
  uninstallFakeHost();
  const host = installFakeHost({ sheets: ["Model"] });
  helpers = host.helpers;
  smt = await import("../src/excel");
  settings = await import("../src/settings");
  settings.setActiveSettings({
    ...settings.DEFAULT_SETTINGS,
    primary: brand[0]!,
    accent: brand[1]!,
    brand,
  });
  helpers.seed("Model!A1", [[1, 2]]);
  helpers.select("Model!A1:B1");
}

describe("the fill cycle with five brand colors", () => {
  beforeEach(async () => {
    await boot([GOLD, MARBLE, GREEN, SLATE, NEON]);
  });

  it("walks the tints, accent, primary, then every extra color, then clears", async () => {
    const receipts: string[] = [];
    const colors: (string | null)[] = [];
    for (let press = 0; press < 8; press += 1) {
      receipts.push(await smt.applyFillCycle());
      const fill = helpers.fill("Model!B1");
      colors.push(fill.pattern === "None" ? null : fill.color);
    }
    expect(receipts).toEqual([
      "Fill: Header tint",
      "Fill: Result tint",
      "Fill: Accent",
      "Fill: Primary",
      "Fill: Brand 3",
      "Fill: Brand 4",
      "Fill: Brand 5",
      "Fill: none",
    ]);
    expect(colors.slice(2)).toEqual([MARBLE, GOLD, GREEN, SLATE, NEON, null]);
    expect(await smt.applyFillCycle()).toBe("Fill: Header tint");
  });

  it("steps on from a brand color the modeller painted by hand", async () => {
    helpers.setFill("Model!A1", { color: SLATE, pattern: "Solid" });
    expect(await smt.applyFillCycle()).toBe("Fill: Brand 5");
  });
});

describe("the font color cycle with five brand colors", () => {
  beforeEach(async () => {
    await boot([GOLD, MARBLE, GREEN, SLATE, NEON]);
  });

  it("walks the model colors, accent, primary, then every extra color", async () => {
    const receipts: string[] = [];
    for (let press = 0; press < 8; press += 1) {
      receipts.push(await smt.applyFontColorCycle());
    }
    expect(receipts).toEqual([
      "Font colour: Formulas",
      "Font colour: Inputs",
      "Font colour: Cross-sheet links",
      "Font colour: Accent",
      "Font colour: Primary",
      "Font colour: Brand 3",
      "Font colour: Brand 4",
      "Font colour: Brand 5",
    ]);
    expect(helpers.font("Model!B1").color).toBe(NEON);
    expect(await smt.applyFontColorCycle()).toBe("Font colour: Formulas");
  });
});

describe("a palette whose colors repeat", () => {
  it("never stalls the fill cycle on a primary equal to the accent", async () => {
    await boot([GOLD, GOLD, GREEN]);
    const receipts: string[] = [];
    for (let press = 0; press < 6; press += 1) {
      receipts.push(await smt.applyFillCycle());
    }
    expect(receipts).toEqual([
      "Fill: Header tint",
      "Fill: Result tint",
      "Fill: Accent",
      "Fill: Brand 3",
      "Fill: none",
      "Fill: Header tint",
    ]);
  });

  // A palette stored with Primary equal to Accent (v3.1.0 to v3.5.0): the font lap has one step for
  // the shared color, named Accent, and lands on each color once before it wraps.
  it("steps the font cycle through a primary equal to the accent once, then wraps", async () => {
    await boot([GOLD, GOLD, GREEN]);
    const receipts: string[] = [];
    const colors: string[] = [];
    for (let press = 0; press < 6; press += 1) {
      receipts.push(await smt.applyFontColorCycle());
      colors.push(helpers.font("Model!B1").color.toUpperCase());
    }
    expect(receipts).toEqual([
      "Font colour: Formulas",
      "Font colour: Inputs",
      "Font colour: Cross-sheet links",
      "Font colour: Accent",
      "Font colour: Brand 3",
      "Font colour: Formulas",
    ]);
    expect(colors.slice(3, 5)).toEqual([GOLD, GREEN]);
    expect(new Set(colors.slice(0, 5)).size).toBe(5);
    expect(colors[5]).toBe(colors[0]);
  });

  it("names the surviving step when a brand color repeats the input blue", async () => {
    await boot([GOLD, MARBLE, "#0057B8", NEON]);
    const receipts: string[] = [];
    for (let press = 0; press < 7; press += 1) {
      receipts.push(await smt.applyFontColorCycle());
    }
    expect(receipts.slice(5)).toEqual([
      "Font colour: Brand 4",
      "Font colour: Formulas",
    ]);
  });
});
