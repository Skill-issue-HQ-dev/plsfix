// @vitest-environment jsdom
// A palette stored with Primary equal to Accent (v3.1.0 to v3.5.0 could save one) draws each color once
// on the Tools tab's strips, while the Brand tab still shows both slots as stored so the user can fix
// one. The cycles' own walk is test/brand-cycles.integration.test.ts. Shipped taskpane.html, Excel mocked.
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

vi.mock("../excel", () => ({
  copySourceLabel: vi.fn(() => null),
  inspectSelection: vi.fn(async () => ({
    address: "A1",
    cells: 1,
    formulas: 0,
    errors: 0,
    blanks: 0,
  })),
  lastUndoSkipped: vi.fn(() => false),
  undoTarget: vi.fn(() => null),
  readWorkbookBrand: vi.fn(async () => null),
  writeWorkbookBrand: vi.fn(async () => undefined),
  setAutocolorOnEdit: vi.fn(async () => undefined),
  paintSwatch: vi.fn(async () => "Fill"),
  readSelectionFills: vi.fn(async () => []),
}));
vi.mock("./custom-styles-panel", () => ({
  refreshCustomStyles: vi.fn(async () => undefined),
}));

import { readWorkbookBrand, writeWorkbookBrand } from "../excel";

const GOLD = "#B27E54";
const GREEN = "#273631";
const FRESH = "#112233";

let store: Map<string, string>;
let writes: number;

// The pane booted over a machine copy that holds this brand list, Primary and Accent as the list's first two.
async function load(brand: string[]) {
  vi.resetModules();
  document.body.innerHTML = readFileSync(
    join(process.cwd(), "taskpane.html"),
    "utf8",
  );
  store.set(
    "plsfix.brand.v1",
    JSON.stringify({ primary: brand[0], accent: brand[1], brand }),
  );
  const shared = await import("./shared");
  const tab = await import("./brand-tab");
  const settings = await import("../settings");
  tab.loadSettings();
  tab.wireBrand();
  tab.renderBrand();
  shared.setExcelReady(true);
  writes = 0;
  return { settings };
}

const hexes = (mode: string): string[] =>
  [
    ...document.querySelectorAll<HTMLButtonElement>(
      `#brand-swatches [data-swatch-strip="${mode}"] [data-swatch-hex]`,
    ),
  ].map((button) => button.dataset.swatchHex ?? "");
const shades = (): number =>
  document.querySelectorAll("#brand-swatches [data-swatch-shades] button")
    .length;
const hexField = (slot: string): HTMLInputElement =>
  document.querySelector<HTMLInputElement>(`[data-slot-hex="${slot}"]`)!;
const editorHexes = (): string[] =>
  [
    ...document.querySelectorAll<HTMLInputElement>(
      '#brand-colors-editor [data-brand-index] input[type="text"]',
    ),
  ].map((field) => field.value);

beforeEach(() => {
  vi.resetAllMocks();
  vi.mocked(readWorkbookBrand).mockResolvedValue(null);
  vi.mocked(writeWorkbookBrand).mockResolvedValue(undefined);
  store = new Map();
  writes = 0;
  vi.stubGlobal("localStorage", {
    getItem: (name: string) => store.get(name) ?? null,
    setItem: (name: string, value: string) => {
      writes += 1;
      store.set(name, value);
    },
  });
});

afterEach(() => {
  vi.unstubAllGlobals();
});

describe("the Tools tab's strips over a palette stored with Primary equal to Accent", () => {
  it("draw each color once, on the Fill and the Font strip alike", async () => {
    await load([GOLD, GOLD, GREEN]);

    expect(hexes("fill")).toEqual([GOLD, GREEN]);
    expect(hexes("font")).toEqual([GOLD, GREEN]);
  });

  it("draw one swatch for a palette that is one color twice, with its Shades row", async () => {
    await load([GOLD, GOLD]);

    expect(hexes("fill")).toEqual([GOLD]);
    expect(hexes("font")).toEqual([GOLD]);
    expect(shades()).toBe(5);
  });

  it("size their row by the colors drawn, not by the list", async () => {
    await load([GOLD, GOLD, GREEN]);

    expect(
      document
        .querySelector<HTMLElement>(
          '#brand-swatches [data-swatch-strip="fill"]',
        )
        ?.style.getPropertyValue("--swatch-count"),
    ).toBe("2");
  });
});

describe("the Brand tab over the same palette", () => {
  it("shows both slots as stored, and the extras in the editor", async () => {
    await load([GOLD, GOLD, GREEN]);

    expect(hexField("primary").value).toBe(GOLD);
    expect(hexField("accent").value).toBe(GOLD);
    expect(editorHexes()).toEqual([GREEN]);
  });

  it("lets the user fix one slot: the new color is taken, saved once, and the strips follow", async () => {
    const { settings } = await load([GOLD, GOLD, GREEN]);

    hexField("accent").value = FRESH;
    hexField("accent").dispatchEvent(new Event("change"));

    expect(settings.getActiveSettings().brand).toEqual([GOLD, FRESH, GREEN]);
    expect(writes).toBe(1);
    expect(hexes("fill")).toEqual([GOLD, FRESH, GREEN]);
  });
});
