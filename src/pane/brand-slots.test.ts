// @vitest-environment jsdom
// The Brand tab's Primary and Accent never repeat a brand color: a hex or picker pick that equals
// another brand color is refused on commit in the palette editor's status line and the slot keeps
// its color (a drag may show the repeat while it lasts); modelling colors are never refused.
// Through the real pane markup with the Excel adapter mocked.
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
}));
vi.mock("./custom-styles-panel", () => ({
  refreshCustomStyles: vi.fn(async () => undefined),
}));

import { writeWorkbookBrand } from "../excel";

const GOLD = "#B27E54";
const MARBLE = "#282623";
const GREEN = "#273631";
const SLATE = "#41564F";
const FRESH = "#112233";

let store: Map<string, string>;
let writes: number;

// The pane with a four-color palette: Primary gold, Accent marble, then green and slate.
async function load() {
  vi.resetModules();
  document.body.innerHTML = readFileSync(
    join(process.cwd(), "taskpane.html"),
    "utf8",
  );
  const shared = await import("./shared");
  const tab = await import("./brand-tab");
  const settings = await import("../settings");
  tab.wireBrand();
  settings.setActiveSettings({
    ...settings.getActiveSettings(),
    primary: GOLD,
    accent: MARBLE,
    brand: [GOLD, MARBLE, GREEN, SLATE],
  });
  tab.renderBrand();
  shared.setExcelReady(true);
  return { settings };
}

const picker = (slot: string): HTMLInputElement =>
  document.querySelector<HTMLInputElement>(`[data-slot-color="${slot}"]`)!;
const hexField = (slot: string): HTMLInputElement =>
  document.querySelector<HTMLInputElement>(`[data-slot-hex="${slot}"]`)!;
const status = (): string =>
  document.querySelector("#brand-colors-editor [data-palette-status]")
    ?.textContent ?? "";
const fillStrip = (): (string | undefined)[] =>
  [
    ...document.querySelectorAll<HTMLButtonElement>(
      '#brand-swatches [data-swatch-strip="fill"] [data-swatch-hex]',
    ),
  ].map((button) => button.dataset.swatchHex);

function typeInto(slot: string, value: string): void {
  hexField(slot).value = value;
  hexField(slot).dispatchEvent(new Event("change"));
}
function drag(slot: string, color: string): void {
  picker(slot).value = color;
  picker(slot).dispatchEvent(new Event("input"));
}
function letGo(slot: string): void {
  picker(slot).dispatchEvent(new Event("change"));
}

beforeEach(() => {
  vi.resetAllMocks();
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

describe("Primary or Accent typed in its hex field", () => {
  it("is refused on commit when another brand color is that color: the status line says so, the field keeps its color", async () => {
    const { settings } = await load();

    typeInto("primary", SLATE);

    expect(status()).toBe(`${SLATE} is already in the list.`);
    expect(hexField("primary").value).toBe(GOLD);
    expect(picker("primary").value).toBe(GOLD.toLowerCase());
    expect(settings.getActiveSettings().primary).toBe(GOLD);
  });

  it("is refused for the Accent too, and for the color of the other head", async () => {
    const { settings } = await load();

    typeInto("accent", GOLD);
    expect(status()).toBe(`${GOLD} is already in the list.`);
    expect(hexField("accent").value).toBe(MARBLE);

    typeInto("primary", MARBLE);
    expect(status()).toBe(`${MARBLE} is already in the list.`);
    expect(hexField("primary").value).toBe(GOLD);
    expect(settings.getActiveSettings().brand).toEqual([
      GOLD,
      MARBLE,
      GREEN,
      SLATE,
    ]);
  });

  it("names the color the way the list holds it, whatever was typed", async () => {
    await load();

    typeInto("accent", "273631");

    expect(status()).toBe(`${GREEN} is already in the list.`);
    expect(hexField("accent").value).toBe(MARBLE);
  });

  it("saves nothing: not on the machine, not in the workbook", async () => {
    await load();

    typeInto("primary", GREEN);

    expect(writes).toBe(0);
    expect(writeWorkbookBrand).not.toHaveBeenCalled();
  });

  it("draws the strips from the list as it was, never from the refused color", async () => {
    await load();

    typeInto("primary", SLATE);

    expect(fillStrip()).toEqual([GOLD, MARBLE, GREEN, SLATE]);
  });

  it("is not refused for its own color: nothing repeats, and the usual save runs", async () => {
    await load();

    typeInto("primary", GOLD);

    expect(status()).toBe("");
    expect(writes).toBe(1);
  });

  it("still commits a color the list does not hold, after a refusal", async () => {
    const { settings } = await load();
    typeInto("primary", SLATE);

    typeInto("primary", FRESH);

    expect(settings.getActiveSettings().brand).toEqual([
      FRESH,
      MARBLE,
      GREEN,
      SLATE,
    ]);
    expect(writes).toBe(1);
    expect(JSON.parse(store.get("plsfix.brand.v1")!).primary).toBe(FRESH);
  });

  it("is not a brand color's business for the modelling colors: Inputs may be gold", async () => {
    const { settings } = await load();

    typeInto("input", GOLD);

    expect(status()).toBe("");
    expect(settings.getActiveSettings().input).toBe(GOLD);
    expect(writes).toBe(1);
  });
});

describe("Primary or Accent picked with the color picker", () => {
  it("may show a repeat while it is dragged, and is refused when it is let go", async () => {
    const { settings } = await load();

    drag("primary", MARBLE.toLowerCase());
    // The preview is in memory only: it may show the repeat.
    expect(settings.getActiveSettings().primary).toBe(MARBLE);
    expect(writes).toBe(0);

    letGo("primary");

    expect(status()).toBe(`${MARBLE} is already in the list.`);
    expect(settings.getActiveSettings().primary).toBe(GOLD);
    expect(hexField("primary").value).toBe(GOLD);
    expect(picker("primary").value).toBe(GOLD.toLowerCase());
    expect(fillStrip()).toEqual([GOLD, MARBLE, GREEN, SLATE]);
    expect(writes).toBe(0);
    expect(writeWorkbookBrand).not.toHaveBeenCalled();
  });

  it("takes back the color from before the drag, not one it passed on the way", async () => {
    const { settings } = await load();

    drag("accent", FRESH);
    drag("accent", "#445566");
    drag("accent", GREEN.toLowerCase());
    letGo("accent");

    expect(status()).toBe(`${GREEN} is already in the list.`);
    expect(settings.getActiveSettings().accent).toBe(MARBLE);
    expect(settings.getActiveSettings().brand).toEqual([
      GOLD,
      MARBLE,
      GREEN,
      SLATE,
    ]);
    expect(hexField("accent").value).toBe(MARBLE);
  });

  it("leaves the next edit to build on the list as it was", async () => {
    const { settings } = await load();
    drag("primary", SLATE.toLowerCase());
    letGo("primary");

    drag("accent", FRESH);
    letGo("accent");

    expect(settings.getActiveSettings().brand).toEqual([
      GOLD,
      FRESH,
      GREEN,
      SLATE,
    ]);
    expect(writes).toBe(1);
  });

  it("saves once, as before, when the drag ends on a color the list does not hold", async () => {
    const { settings } = await load();

    drag("primary", FRESH);
    drag("primary", "#445566");
    letGo("primary");

    expect(settings.getActiveSettings().primary).toBe("#445566");
    expect(status()).toBe("");
    expect(writes).toBe(1);
    expect(writeWorkbookBrand).toHaveBeenCalledTimes(1);
  });

  it("does not refuse a modelling color picked as a brand color", async () => {
    const { settings } = await load();

    drag("link", GREEN.toLowerCase());
    letGo("link");

    expect(settings.getActiveSettings().link).toBe(GREEN);
    expect(status()).toBe("");
    expect(writes).toBe(1);
  });
});
